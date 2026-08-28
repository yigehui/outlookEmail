# -*- coding: utf-8 -*-
"""
outlook_web/segments/12_cloudflare_mail.py —— 通过 cloudflare_temp_email 项目的 API 收发邮件。

用途:微软「绑定辅助邮箱」流程里,给微软账号绑定一个 cf 临时邮箱作为辅助邮箱,
然后从这个 cf 邮箱里收取微软发来的验证码。

本文件从 reg-factory/common/cloudflare_mail.py 原样移植,仅把配置读取改为
「环境变量优先、DB settings 回退」(get_setting 来自 segment 01_bootstrap)。

配置来源(环境变量优先 > DB settings > 默认空):
  环境变量 CF_MAIL_BASE / DB cf_mail_base      Worker base URL,如 https://mail.example.com
  环境变量 CF_MAIL_ADMIN / DB cf_mail_admin    admin 密码(x-admin-auth header)
  环境变量 CF_MAIL_DOMAIN / DB cf_mail_domain  可用域名(已绑 Email Routing catch-all)
  环境变量 CF_MAIL_SITE_PASS / DB cf_mail_site_pass  站点私有密码(x-custom-auth,可选)
  环境变量 CF_MAIL_PROXY / DB cf_mail_proxy    cf 收信代理(可选,None=直连)

API 参考(官方文档 temp-mail-docs.awsl.uk):
  建固定邮箱(管理员 API,可指定固定 name):
    POST {base}/admin/new_address
    headers: x-admin-auth: <admin密码>
    body: {"enablePrefix": false, "name": "<固定名>", "domain": "<域名>"}
    -> {"jwt": "...", "address": "...", "address_id": 123}
  收解析好的邮件(地址 JWT):
    GET {base}/api/parsed_mails?limit=&offset=
    headers: Authorization: Bearer <地址JWT>
    -> {results: [{id, subject, text, html, sender, ...}], count}
  旧部署可能没有 parsed_mails,回退 raw:
    GET {base}/api/mails?limit=&offset=
    -> 返回 raw RFC822,需客户端解析

辅助邮箱命名约定:ms-<微软号前缀>@<域名>,与微软号一一对应、长期保留。
微软号前缀 = 邮箱 @ 前部分(如 joseph_lee239@outlook.com -> joseph_lee239)。
"""

import os
import re
import secrets
import string
import time
import urllib.parse

import requests

# cf 域名(nuo.dpdns.org 在 Cloudflare 后)在本地/机房 DNS 常解析不稳,且 Worker 需外网可达。
# 故 cf 收信走代理(与 MS token 直连策略相反——MS 端点直连、cf worker 走代理)。
# 代理来源优先级:显式 set_proxy > CF_MAIL_PROXY 环境变量 > DB cf_mail_proxy > 不走代理(直连)。
_PROXY = None  # 模块级缓存


def set_proxy(proxy_url):
    """显式设置 cf 收信代理(socks5h://... 或 http://...)。None=直连。"""
    global _PROXY
    _PROXY = proxy_url or None


def _resolve_proxy():
    # 优先级:显式 set_proxy(_PROXY) > CF_MAIL_PROXY 环境变量 > DB cf_mail_proxy > 直连。
    # get_setting 触及 DB,可能在没有 app context 的调用方里抛错,故包 try/except。
    env = os.environ.get("CF_MAIL_PROXY")
    if env:
        return _PROXY or env
    try:
        return _PROXY or (get_setting('cf_mail_proxy', '') or None)
    except Exception:
        return _PROXY or None


def _cf_session():
    s = requests.Session()
    s.trust_env = False
    px = _resolve_proxy()
    if px:
        s.proxies = {"http": px, "https": px}
    else:
        s.proxies = {"http": None, "https": None}  # 显式直连
    s.headers.update({"Content-Type": "application/json"})
    return s


def _resolve_bind_channel():
    """优先用设置页指定的 CF 渠道（bind_cf_channel 渠道名），否则默认渠道，都没有返回 None。
    include_secret=True 才返回解密后的 admin_password。需在 app context 内调用。"""
    try:
        name = (get_setting('bind_cf_channel', '') or '').strip()
        if name:
            ch = get_cloudflare_channel_by_name(name, include_secret=True)
            if ch:
                return ch
        ch = get_default_cloudflare_channel(include_secret=True)
        if ch:
            return ch
    except Exception:
        return None
    return None


def _cfg():
    # 渠道优先：设置页 bind_cf_channel 指定渠道 > 默认渠道；都没有才回退环境变量/DB cf_mail_*。
    # 渠道表无 site_pass/proxy 字段，这两项仍只走环境变量（可选）。
    # F3.6 之后可在设置页选 CF 渠道而无需环境变量。
    channel = _resolve_bind_channel()
    if channel:
        worker = (channel.get('worker_domain', '') or '').strip().rstrip('/')
        base = f"https://{worker}" if worker else ''
        admin = channel.get('admin_password', '') or ''
        domains = channel.get('email_domains')
        if isinstance(domains, list):
            domain = domains[0].strip() if domains else ''
        else:
            parts = str(domains or '').split(',')
            domain = parts[0].strip() if parts else ''
        site_pass = os.environ.get("CF_MAIL_SITE_PASS") or get_setting('cf_mail_site_pass', '')
        return base, admin, domain, site_pass
    # 回退：环境变量优先 > DB settings(get_setting) > 默认空。
    base = (os.environ.get("CF_MAIL_BASE") or get_setting('cf_mail_base', '') or "").strip().rstrip("/")
    admin = os.environ.get("CF_MAIL_ADMIN") or get_setting('cf_mail_admin', '')
    domain = (os.environ.get("CF_MAIL_DOMAIN") or get_setting('cf_mail_domain', '') or "").strip()
    site_pass = os.environ.get("CF_MAIL_SITE_PASS") or get_setting('cf_mail_site_pass', '')
    return base, admin, domain, site_pass


def _auth_headers(admin=None, jwt=None, site_pass=None):
    h = {}
    if admin:
        h["x-admin-auth"] = admin
    if jwt:
        h["Authorization"] = f"Bearer {jwt}"
    if site_pass:
        h["x-custom-auth"] = site_pass
    return h


def ms_prefix(ms_email):
    """joseph_lee239@outlook.com -> joseph_lee239"""
    return (ms_email or "").split("@")[0].strip().lower()


def build_address_name(ms_email):
    """微软号 -> cf 辅助邮箱 name(固定,长期保留)。
    joseph_lee239@outlook.com -> ms-joseph_lee239
    注意:cf 会把 name 里的非字母数字字符(如下划线)规范化进地址,
    但建邮箱时 name 保持原样用于「复用找回」(list 里 name 字段比对)。"""
    p = ms_prefix(ms_email)
    return f"ms-{p}" if p else ""


def create_or_get_address(ms_email, verbose=True):
    """给微软号建(或复用)一个固定 cf 辅助邮箱,返回 {address, jwt, use_admin, address_id, name, password}。
    - 首次创建:返回地址 JWT,收码用地址 JWT(/api/parsed_mails)
      ENABLE_ADDRESS_PASSWORD 开启时,worker 自动生成 8 位密码,响应带 password(明文仅此一次)。
    - 已存在(cf 地址 JWT 无法重新获取):返回 use_admin=True,收码改走 admin API(/admin/mails?address=)
      admin API 用 x-admin-auth,不依赖地址 JWT。
      此时若地址密码功能开启,会 admin 重设一个新随机密码并返回明文(旧密码作废)。
    - 密码功能未开启(ENABLE_ADDRESS_PASSWORD=false):password=None。"""
    base, admin, domain, site_pass = _cfg()
    if not (base and admin and domain):
        raise RuntimeError("CF_MAIL_BASE/CF_MAIL_ADMIN/CF_MAIL_DOMAIN 未配置")
    name = build_address_name(ms_email)
    if not name:
        raise RuntimeError(f"无法从 {ms_email} 推导辅助邮箱 name")
    # 推断最终邮箱地址(cf 会把 name 规范化:去掉非字母数字字符)
    # ms-joseph_lee239 -> msjosephlee239@domain
    norm_name = re.sub(r'[^A-Za-z0-9]', '', name)
    expected_address = f"{norm_name}@{domain}"
    sess = _cf_session()
    url = f"{base}/admin/new_address"
    body = {"enablePrefix": False, "name": name, "domain": domain}
    headers = _auth_headers(admin=admin, site_pass=site_pass)
    try:
        resp = sess.post(url, json=body, headers=headers, timeout=30)
    except Exception as e:
        raise RuntimeError(f"建 cf 邮箱请求失败: {e}")
    if resp.status_code == 200:
        d = resp.json()
        if verbose:
            pwd_note = " 密码已生成" if d.get("password") else ""
            print(f"  [cf] 新建辅助邮箱 {d.get('address')} (name={name}){pwd_note}")
        return {"jwt": d.get("jwt"), "address": d.get("address"),
                "use_admin": False, "address_id": d.get("address_id"), "name": name,
                "password": d.get("password")}
    # 已存在 -> 用 admin 收码(地址 JWT 拿不回来)
    if "already exists" in (resp.text or "").lower() or resp.status_code == 400:
        if verbose:
            print(f"  [cf] 辅助邮箱已存在 {expected_address},改用 admin API 收码")
        # 已存在地址:worker 不返回密码(明文不可逆)。admin 重设一个新随机密码,
        # 明文返回给调用方记录。重设失败不影响收码,静默降级 password=None。
        password = _reset_address_password(expected_address, verbose=verbose)
        return {"jwt": None, "address": expected_address, "use_admin": True,
                "address_id": None, "name": name, "password": password}
    raise RuntimeError(f"无法为 {ms_email} 建立辅助邮箱: {resp.status_code} {resp.text[:200]}")


def _find_address_id(address):
    """从 admin/address 列表里找地址的 id。找不到返回 None。"""
    base, admin, domain, site_pass = _cfg()
    if not base:
        return None
    sess = _cf_session()
    headers = _auth_headers(admin=admin, site_pass=site_pass)
    offset = 0
    target = (address or "").strip().lower()
    while offset < 20000:  # 上限保护
        try:
            resp = sess.get(f"{base}/admin/address", headers=headers,
                            params={"limit": 100, "offset": offset}, timeout=30)
        except Exception:
            return None
        if resp.status_code != 200:
            return None
        items = resp.json().get("results", [])
        if not items:
            return None
        for it in items:
            if str(it.get("name", "")).strip().lower() == target:
                return it.get("id")
        offset += len(items)
    return None


def _reset_address_password(address, verbose=True):
    """admin 重设地址密码为新的随机密码,返回明文密码(失败返回 None)。
    POST {base}/admin/address/:id/reset_password  body {"password": "<SHA-256 hex>"}
    仅在 worker 开了 ENABLE_ADDRESS_PASSWORD 时有效(403=没开,静默降级)。"""
    import hashlib
    base, admin, domain, site_pass = _cfg()
    if not (base and admin):
        return None
    address_id = _find_address_id(address)
    if not address_id:
        if verbose:
            print(f"  [cf] 找不到 {address} 的 address_id,跳过设密码")
        return None
    alphabet = string.ascii_letters + string.digits
    plain = ''.join(secrets.choice(alphabet) for _ in range(12))
    hashed = hashlib.sha256(plain.encode("utf-8")).hexdigest()
    sess = _cf_session()
    headers = _auth_headers(admin=admin, site_pass=site_pass)
    try:
        resp = sess.post(f"{base}/admin/address/{address_id}/reset_password",
                         json={"password": hashed}, headers=headers, timeout=30)
    except Exception as e:
        if verbose:
            print(f"  [cf] 重设密码请求失败: {e}")
        return None
    if resp.status_code == 200:
        if verbose:
            print(f"  [cf] 已重设 {address} 的邮箱密码")
        return plain
    # 403 = ENABLE_ADDRESS_PASSWORD 未开启,静默降级
    if verbose and resp.status_code != 403:
        print(f"  [cf] 重设密码失败 {resp.status_code}: {resp.text[:120]}")
    return None


def list_addresses(ms_email=None):
    """列出地址(admin API)。旧部署返回的项里不一定有 jwt,需逐个调(简化:不拿 jwt)。"""
    base, admin, domain, site_pass = _cfg()
    if not base:
        return []
    sess = _cf_session()
    url = f"{base}/admin/address"
    headers = _auth_headers(admin=admin, site_pass=site_pass)
    try:
        resp = sess.get(url, headers=headers, timeout=30)
        if resp.status_code != 200:
            return []
        d = resp.json()
        items = d.get("results") or d.get("address") or d if isinstance(d, list) else (d.get("results") or [])
        return items
    except Exception as e:
        print(f"  [cf] list_addresses error: {e}")
        return []


def fetch_parsed_mails(jwt, limit=20, offset=0, site_pass=None):
    """取服务端解析好的邮件列表。返回 [{subject,text,html,sender,...}]。"""
    base, _, _, sp = _cfg()
    sess = _cf_session()
    url = f"{base}/api/parsed_mails"
    headers = _auth_headers(jwt=jwt, site_pass=site_pass or sp)
    resp = sess.get(url, headers=headers, params={"limit": limit, "offset": offset}, timeout=20)
    if resp.status_code == 200:
        return resp.json().get("results", [])
    if resp.status_code == 404:
        # 旧部署没 parsed_mails,回退 raw(返回原始,调用方需自行解析,这里先给空结构)
        return _fetch_raw_as_parsed(jwt, site_pass or sp)
    raise RuntimeError(f"fetch_parsed_mails {resp.status_code}: {resp.text[:200]}")


def _fetch_raw_as_parsed(jwt, site_pass, limit=20):
    """旧部署回退:取 raw RFC822,简单提取 subject/text。"""
    base, _, _, sp = _cfg()
    sess = _cf_session()
    url = f"{base}/api/mails"
    headers = _auth_headers(jwt=jwt, site_pass=site_pass or sp)
    resp = sess.get(url, headers=headers, params={"limit": limit, "offset": 0}, timeout=20)
    if resp.status_code != 200:
        return []
    out = []
    for m in resp.json().get("results", []):
        raw = m.get("raw") or ""
        subj = re.search(r"Subject:\s*(.+)", raw)
        out.append({"id": m.get("id"), "subject": subj.group(1) if subj else "",
                    "text": raw, "html": "", "sender": m.get("source", "")})
    return out


def extract_code(text):
    """从邮件主题/正文提取微软 6-7 位验证码。微软安全信息验证码通常 6-7 位数字。"""
    if not text:
        return None
    # 去掉 HTML 标签
    t = re.sub(r"<style.*?</style>", " ", text, flags=re.S | re.I)
    t = re.sub(r"<script.*?</script>", " ", t, flags=re.S | re.I)
    t = re.sub(r"<[^>]+>", " ", t)
    # 优先匹配 "code is 123456" / "验证码：123456" / "code: 123456"
    for pat in (r"(?:code is|your code|verification code|验证码|security code)[^\d]{0,15}(\d{6,7})",
                r"\b(\d{6,7})\b"):
        m = re.search(pat, t, re.I)
        if m:
            return m.group(1)
    return None


def fetch_admin_mails(address, limit=20, offset=0):
    """admin API 取指定地址的邮件(无需地址 JWT)。返回 raw 邮件列表(含 raw 字段)。
    GET {base}/admin/mails?address=&limit=&offset=  headers: x-admin-auth"""
    base, admin, _, site_pass = _cfg()
    sess = _cf_session()
    url = f"{base}/admin/mails"
    headers = _auth_headers(admin=admin, site_pass=site_pass)
    resp = sess.get(url, headers=headers, params={"limit": limit, "offset": offset, "address": address}, timeout=20)
    if resp.status_code != 200:
        raise RuntimeError(f"fetch_admin_mails {resp.status_code}: {resp.text[:200]}")
    return resp.json().get("results", [])


def parse_admin_mail(m):
    """把 admin/mails 返回的单条(raw)转成 {subject,text,sender,id}。raw 是 RFC822。
    简单解析:不引依赖,正则提 Subject/From 和正文。"""
    raw = m.get("raw") or ""
    subject = ""
    frm = ""
    sm = re.search(r'^Subject:\s*(.+(?:\r?\n[ \t].+)*)', raw, re.M)
    if sm:
        subject = re.sub(r'\r?\n[ \t]', '', sm.group(1)).strip()
    fm = re.search(r'^From:\s*(.+(?:\r?\n[ \t].+)*)', raw, re.M)
    if fm:
        frm = re.sub(r'\r?\n[ \t]', '', fm.group(1)).strip()
    # 正文:取 Content-Transfer-Encoding 后的纯文本段(简化:取 text/plain 部分)
    body = raw
    pm = re.search(r'Content-Type:\s*text/plain[\s\S]*?\r?\n\r?\n([\s\S]*?)(?:\r?\n--|\Z)', raw, re.I)
    if pm:
        body = pm.group(1)
    return {"id": m.get("id"), "subject": subject, "text": body, "sender": frm,
            "from": frm, "raw": raw}


def wait_for_code(addr_or_jwt, sender_contains=("account-security-noreply", "microsoftaccount",
                                                "account@accountprotection", "microsoft.com"),
                  subject_contains=("code", "verify", "verification", "security", "验证"),
                  max_wait=150, poll=5, received_after_id=None, use_admin=False, address=None, verbose=True):
    """轮询 cf 邮箱取微软验证码。返回 code 字符串或 None。
    addr_or_jwt: 地址 JWT(use_admin=False) 或忽略(use_admin=True 时用 address)。
    received_after_id: 只取 id > 该值的邮件(避免取到旧的)。"""
    start = time.time()
    last_id = received_after_id or 0
    while time.time() - start < max_wait:
        try:
            if use_admin:
                raws = fetch_admin_mails(address, limit=20)
                mails = [parse_admin_mail(m) for m in raws]
            else:
                mails = fetch_parsed_mails(addr_or_jwt, limit=20)
        except Exception as e:
            if verbose:
                print(f"  [cf] 取信异常: {e}")
            time.sleep(poll)
            continue
        for m in mails:
            mid = m.get("id", 0)
            if mid and mid <= last_id:
                continue
            subj = (m.get("subject") or "").lower()
            frm = (m.get("sender") or m.get("from") or "").lower()
            text = m.get("text") or m.get("html") or ""
            hit_sender = any(s in frm for s in sender_contains) if sender_contains else False
            hit_subject = any(s in subj for s in subject_contains) if subject_contains else False
            if not (hit_sender or hit_subject):
                continue
            code = extract_code(text) or extract_code(m.get("subject"))
            if code:
                if verbose:
                    print(f"  [cf] 取到验证码 {code} (from={frm[:40]} subject={m.get('subject','')[:40]})")
                return code
        elapsed = int(time.time() - start)
        if verbose:
            print(f"  [cf] 等待验证码... ({elapsed}s/{max_wait}s)")
        time.sleep(poll)
    if verbose:
        print("  [cf] 取码超时")
    return None


# ============================== 自检 ==============================
def _selftest():
    """配置好凭据后跑:python -m common.cloudflare_mail"""
    base, admin, domain, sp = _cfg()
    print(f"base={base} domain={domain} admin={'有' if admin else '无'} site_pass={'有' if sp else '无'}")
    if not (base and admin and domain):
        print("缺配置,无法自检")
        return
    # 自检:建一个测试邮箱 + 校验 JWT
    test = create_or_get_address("selftest_probe@outlook.com")
    jwt = test["jwt"]
    print(f"邮箱就绪: {test['address']}")
    # 校验 JWT
    sess = _cf_session()
    r = sess.get(f"{base}/api/settings", headers=_auth_headers(jwt=jwt, site_pass=sp), timeout=20)
    print(f"JWT 校验(/api/settings): {r.status_code} {r.text[:120]}")


    # 自检需在应用上下文内手动调用，不在 segment 加载阶段自动触发
    # （segment 经 exec 加载，__name__ 始终为 __main__，__main__ 保护会误触发）

