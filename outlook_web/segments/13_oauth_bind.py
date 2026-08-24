# -*- coding: utf-8 -*-
"""OAuth 绑定辅助邮箱:在已登录 session 里把 proofs/Add 绑成 CF 临时邮箱。
移植自 reg-factory/extract_graph_tokens.py(bind_proof_in_session + form 解析器)。
依赖 segment 12 (cloudflare_mail) 提供的收码函数(裸名:wait_for_code/fetch_admin_mails/fetch_parsed_mails/parse_admin_mail/extract_code)。"""
import re
import urllib.parse


def _bind_log(tag, msg, level="INFO"):
    """bind_proof_in_session 的最小日志(移植自 reg-factory _graph_log,降级为 stderr print)。"""
    import sys
    print(f"[{level}] {tag} {msg}", file=sys.stderr, flush=True)


def _save_bind_debug_html(email, idx, reason, html):
    """占位:不落盘 debug HTML(outlookEmail 未配 debug 目录),返回空串保持日志格式。"""
    return ""


def _parse_proof_add_form(text, url):
    """从 proofs/Add 真表单页解析出 form action + hidden 字段(含 canary)。
    返回 (form_action, form_data) 或 (None, None)。"""
    form_match = re.search(r'<form[^>]*action="([^"]+)"[^>]*>(.*?)</form>', text, re.DOTALL | re.IGNORECASE)
    if not form_match:
        return None, None
    form_action = form_match.group(1).replace("&amp;", "&")
    form_body = form_match.group(2)
    hidden = re.findall(r'<input[^>]*name="([^"]*)"[^>]*value="([^"]*)"', form_body)
    form_data = {n: v for n, v in hidden}
    if not form_action.startswith("http"):
        base = urllib.parse.urlparse(url)
        form_action = f"{base.scheme}://{base.netloc}{form_action}"
    return form_action, form_data


def _parse_proof_verify_form(text, url):
    """从 proofs/Verify 页解析 frmVerifyProof 的 action(含 epid)+ hidden(canary/action=VerifyProof)。
    返回 (form_action, form_data) 或 (None, None)。epid 在 action URL 里,一起带回去。"""
    # 只取 frmVerifyProof 这个 form(页面还有个 frmSubmitSLT 干扰,且 slt 为空不提交)
    form_match = re.search(
        r'<form[^>]*(?:id|name)="frmVerifyProof"[^>]*action="([^"]+)"[^>]*>(.*?)</form>',
        text, re.DOTALL | re.IGNORECASE)
    if not form_match:
        # 退而求其次:含 Verify 的 form
        form_match = re.search(r'<form[^>]*action="([^"]*proofs/Verify[^"]*)"[^>]*>(.*?)</form>',
                               text, re.DOTALL | re.IGNORECASE)
        if not form_match:
            return None, None
    form_action = form_match.group(1).replace("&amp;", "&")
    form_body = form_match.group(2)
    hidden = re.findall(r'<input[^>]*name="([^"]*)"[^>]*value="([^"]*)"', form_body)
    form_data = {n: v for n, v in hidden}
    if not form_action.startswith("http"):
        base = urllib.parse.urlparse(url)
        form_action = f"{base.scheme}://{base.netloc}{form_action}"
    return form_action, form_data


def bind_proof_in_session(session, html, url, cf_address, cf_jwt=None, use_admin=False, idx=0,
                          cm_module=None, max_wait=180, poll=3):
    """在一个已登录的 requests.Session 里,把 proofs/Add 真绑成 cf 辅助邮箱。
    步骤:AddProof(填 EmailAddress)→ 微软发码到 cf → 收码 → VerifyProof(填 iOttText)。
    成功后返回下一个响应 resp(通常落在 Consent 或已登录页),调用方继续跟 oauth。
    失败返回 None。

    cm_module: 保留以兼容旧签名;segment 12 的收码函数已是 web_outlook_app 裸名,故忽略此参数。
    """
    tag = f"[#{idx}]"
    # segment 12 已加载,wait_for_code/fetch_admin_mails/fetch_parsed_mails 均为裸名,无需 import。

    # 1) proofs/Add:解析 form + 填 EmailAddress
    form_action, form_data = _parse_proof_add_form(html, url)
    if not form_action:
        debug_path = _save_bind_debug_html(cf_address or "bind", idx, "proofs_add_no_form", html)
        _bind_log(tag, f"bind: proofs/Add 无 form debug={debug_path}", "WARN")
        return None
    if not cf_address:
        _bind_log(tag, "bind: 缺 cf 辅助邮箱地址", "ERR")
        return None
    # AddProof 提交字段:canary(已有)+ action=AddProof + EmailAddress + iProofOptions=Email
    form_data["action"] = "AddProof"
    form_data["EmailAddress"] = cf_address
    form_data.setdefault("iProofOptions", "Email")
    # 提交前记 cf 收件箱基线 id(微软发码很快,基线必须在提交前取)
    base_last_id = 0
    try:
        if use_admin:
            raws = fetch_admin_mails(cf_address, limit=20)
            base_last_id = max([m.get("id", 0) for m in raws] or [0])
        else:
            mails = fetch_parsed_mails(cf_jwt, limit=20)
            base_last_id = max([m.get("id", 0) for m in mails] or [0])
    except Exception as e:
        _bind_log(tag, f"bind: 取 cf 基线失败(继续): {e}", "WARN")
    _bind_log(tag, f"bind: 提交 AddProof email={cf_address} base_id={base_last_id}")

    resp = session.post(form_action, data=form_data, timeout=30, allow_redirects=True)
    vurl = getattr(resp, "url", "") or ""
    vtext = resp.text or ""
    _bind_log(tag, f"bind: AddProof -> {resp.status_code} url={vurl[:80]}")

    # 2) 跟重定向到 proofs/Verify(可能要追一跳)
    for _ in range(5):
        if "proofs/verify" in (vurl or "").lower():
            break
        # 有时落在中间页(Consent 不应在这阶段;proofs/Add 重复说明没绑成功)
        if "consent" in (vurl or "").lower() or "localhost" in (vurl or ""):
            break
        fm = re.search(r'<form[^>]*action="([^"]+)"', vtext, re.IGNORECASE)
        if fm and ("DoSubmit" in vtext or "fmHF" in vtext):
            fa = fm.group(1).replace("&amp;", "&")
            if not fa.startswith("http"):
                base = urllib.parse.urlparse(vurl)
                fa = f"{base.scheme}://{base.netloc}{fa}"
            hid = re.findall(r'<input[^>]*name="([^"]*)"[^>]*value="([^"]*)"', vtext)
            resp = session.post(fa, data={n: v for n, v in hid}, timeout=30, allow_redirects=True)
            vurl = getattr(resp, "url", "") or ""; vtext = resp.text or ""
            continue
        break

    if "proofs/verify" not in (vurl or "").lower():
        debug_path = _save_bind_debug_html(cf_address or "bind", idx, "proofs_verify_not_reached", vtext)
        _bind_log(tag, f"bind: 未到 proofs/Verify (url={vurl[:80]}) debug={debug_path}", "WARN")
        return None

    # 3) 解析 Verify form(canary/epid 在 action URL)+ 收码
    vaction, vdata = _parse_proof_verify_form(vtext, vurl)
    if not vaction:
        debug_path = _save_bind_debug_html(cf_address or "bind", idx, "proofs_verify_no_form", vtext)
        _bind_log(tag, f"bind: proofs/Verify 无 form debug={debug_path}", "WARN")
        return None
    vdata["action"] = "VerifyProof"
    code = wait_for_code(cf_jwt, received_after_id=base_last_id, max_wait=max_wait,
                         poll=poll, use_admin=use_admin, address=cf_address)
    if not code:
        # 兜底:基线后没新码,可能是重试同一 proof(微软限频不发新码),
        # 但收件箱里基线那封码对当前 pending proof 仍有效 —— 取最新一封匹配码试。
        code = _fetch_latest_code_fallback(cf_address, cf_jwt, use_admin)
        if not code:
            _bind_log(tag, "bind: cf 取码超时(且无兜底可用码)", "WARN")
            return None
        _bind_log(tag, f"bind: 用兜底最新码 {code}(重试场景基线码仍有效)")
    else:
        _bind_log(tag, f"bind: 取到验证码 {code}")
    vdata["iOttText"] = code
    # allow_redirects=False:绑完后重定向链会一路跟到 http://localhost/?code=...
    # 若 allow_redirects=True,requests 会真去连本地 80 端口 → ConnectionError。
    # 让主循环手动 follow Location(它有 localhost code 拦截),避免本地连不上炸掉。
    resp = session.post(vaction, data=vdata, timeout=30, allow_redirects=False)
    _bind_log(tag, f"bind: VerifyProof -> {resp.status_code} url={getattr(resp,'url','')[:80]}")
    return resp


def _fetch_latest_code_fallback(address, jwt, use_admin):
    """wait_for_code 超时兜底:直接取收件箱里最新一封匹配的微软安全码(忽略基线)。
    用于重试同一 pending proof 时微软限频不发新码、但旧码仍有效的场景。"""
    try:
        if use_admin:
            raws = fetch_admin_mails(address, limit=10)
            mails = [parse_admin_mail(m) for m in raws]
        else:
            mails = fetch_parsed_mails(jwt, limit=10)
        for m in mails:
            code = extract_code(m.get("text") or m.get("html") or "") or extract_code(m.get("subject"))
            if code:
                return code
    except Exception:
        pass
    return None
