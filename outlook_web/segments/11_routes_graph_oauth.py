from __future__ import annotations

import json
import os
import queue
import re
import threading
import urllib.parse
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import TYPE_CHECKING, Any, Callable, Dict, Optional

from flask import stream_with_context

if TYPE_CHECKING:
    from web_outlook_app import *  # noqa: F403


# ==================== Graph OAuth 自动提取 ====================

# client_id 和 redirect_uri 复用 01_bootstrap.py 中的 OAUTH_CLIENT_ID / OAUTH_REDIRECT_URI，
# 不再单独定义 GRAPH_EXTRACT_CLIENT_ID / GRAPH_EXTRACT_REDIRECT_URI 环境变量。
# scope 和 authority 为 Graph 自动提取专用，值与常规 OAuth 不同，保持独立。
GRAPH_EXTRACT_SCOPE = os.getenv(
    "GRAPH_EXTRACT_SCOPE",
    "offline_access https://outlook.office.com/IMAP.AccessAsUser.All",
)
# GraphAPI：与 OAUTH_GRAPH_SCOPES 对齐，含读信 / 标已读写权限 / User.Read
GRAPH_EXTRACT_GRAPH_SCOPE = os.getenv(
    "GRAPH_EXTRACT_GRAPH_SCOPE",
    " ".join(["offline_access", *OAUTH_GRAPH_SCOPES]),
)
GRAPH_EXTRACT_AUTHORITY = os.getenv("GRAPH_EXTRACT_AUTHORITY", "consumers")
GRAPH_EXTRACT_SCOPE_BY_MODE = {
    "imap": GRAPH_EXTRACT_SCOPE,
    "graph": GRAPH_EXTRACT_GRAPH_SCOPE,
}

# Backward-compatible aliases for existing docs/tests.
GRAPH_CLIENT_ID = OAUTH_CLIENT_ID
GRAPH_REDIRECT_URI = OAUTH_REDIRECT_URI
GRAPH_SCOPE = GRAPH_EXTRACT_SCOPE

# 授权主循环步数上限：登录/绑定后的重定向链会夹多张 Consent/DoSubmit 跳板页，
# 每跳算一步。记在模块级便于按实测调整与测试覆盖。
MAX_OAUTH_STEPS = int(os.getenv("GRAPH_OAUTH_MAX_STEPS", "40"))

GRAPH_OAUTH_TASKS: Dict[str, Dict[str, Any]] = {}
GRAPH_OAUTH_DONE = object()


def normalize_graph_oauth_mode(mode: Any) -> str:
    normalized = str(mode or "graph").strip().lower()
    return normalized if normalized in GRAPH_EXTRACT_SCOPE_BY_MODE else "graph"


def graph_oauth_mode_label(mode: str) -> str:
    return "GraphAPI" if mode == "graph" else "IMAP授权"


def graph_oauth_sse(payload: Dict[str, Any]) -> str:
    return f"data: {json.dumps(payload, ensure_ascii=False)}\n\n"


def graph_oauth_safe_details(details: Any) -> str:
    return sanitize_error_details(str(details or ""))[:500]


def graph_oauth_log(log: Optional[Callable[[str], None]], message: str) -> None:
    if log:
        log(graph_oauth_safe_details(message))


def make_graph_oauth_response(success: bool, error: str = "", details: str = "",
                              **extra: Any) -> Dict[str, Any]:
    payload = {"success": bool(success)}
    if error:
        payload["error"] = graph_oauth_safe_details(error)
    if details:
        payload["details"] = graph_oauth_safe_details(details)
    payload.update(extra)
    return payload


def build_graph_authorize_url(client_id: str, redirect_uri: str, scope: str,
                              authority: str) -> str:
    return (
        f"https://login.microsoftonline.com/{authority}/oauth2/v2.0/authorize"
        f"?client_id={urllib.parse.quote(client_id, safe='')}"
        f"&response_type=code"
        f"&redirect_uri={urllib.parse.quote(redirect_uri, safe='')}"
        f"&scope={urllib.parse.quote(scope)}"
        f"&response_mode=query"
    )


def is_oauth_code_redirect(loc: str) -> bool:
    """Location 是否已回到 redirect_uri 并带上授权码/错误。

    判定必须解析 query 参数名,不能做子串判断:
    - `"localhost" in loc` 会把微软中间跳转误判成终点 —— authorize URL 里
      redirect_uri 是 URL 编码的(http%3a%2f%2flocalhost%3a8080),字面就含 localhost;
    - `"code=" in loc` 同样误判 —— `response_type=code&...` 字面含 "code="。
    两者叠加 → 拿空响应体死循环到「授权流程卡住」。2026-09 绑辅助邮箱后重新授权时
    踩到的真坑；按参数判定后,中间跳转会被继续跟下去,而不是当成终点。
    """
    text = str(loc or "").strip()
    if not text:
        return False
    if not re.match(r'^(https?://|/|\?)', text):
        return False
    parts = urllib.parse.urlsplit(text)
    params = urllib.parse.parse_qs(parts.query)
    if "code" not in params and "error" not in params:
        return False
    host = parts.netloc.split("@")[-1].split(":")[0].lower()
    # 相对 Location(/?code=…)没有 host,视为回到 redirect_uri
    return host in ("localhost", "127.0.0.1") if host else True


def make_light_response(url: str, text: str = "", status_code: int = 200):
    return type("GraphOauthResponse", (), {
        "url": url,
        "text": text,
        "status_code": status_code,
        "headers": {},
    })()


def extract_hidden_inputs(html: str) -> Dict[str, str]:
    return {
        name: value
        for name, value in re.findall(
            r'<input[^>]*name="([^"]*)"[^>]*value="([^"]*)"',
            html or "",
            re.IGNORECASE,
        )
    }


def _unescape_form_fields(html: str) -> Dict[str, str]:
    """取表单 hidden 字段并做 HTML 反转义。

    DoSubmit(fmHF) 表单的 hidden value 是 HTML 转义的（&quot; &amp; &#39; 等），
    浏览器提交前会解码；纯 HTTP 必须先 unescape 再发，否则服务端收到坏 JSON
    （scenarios 字段被破坏）→ 302 oauth server_error。2026-09 发现的主坑。
    同名 input 保留首个（与 extract_hidden_inputs 的 dict 覆盖语义不同，对齐上游）。
    """
    import html as _html
    out: Dict[str, str] = {}
    for name, value in re.findall(
        r'<input[^>]*name="([^"]*)"[^>]*value="([^"]*)"',
        str(html or ""),
    ):
        if name in out:
            continue
        out[name] = _html.unescape(value)
    return out


def absolute_form_action(action: str, current_url: str) -> str:
    action = (action or "").replace("&amp;", "&")
    if action.startswith("http"):
        return action
    base = urllib.parse.urlparse(current_url)
    if action.startswith("/"):
        return f"{base.scheme}://{base.netloc}{action}"
    path = urllib.parse.urljoin(f"{base.scheme}://{base.netloc}{base.path}", action)
    return path


def parse_server_data(html: str) -> Dict[str, Any]:
    """解析页面里的 `ServerData = {...};`（非贪婪 + 括号平衡兜底），失败返回 {}。"""
    raw = str(html or "")
    match = re.search(r'ServerData\s*=\s*(\{.*?\});', raw, re.DOTALL)
    if not match:
        return {}
    for candidate in (match.group(1), _parse_cfg_json_balanced(match.group(1))):
        try:
            parsed = json.loads(candidate)
        except Exception:
            continue
        if isinstance(parsed, dict):
            return parsed
    return {}


def handle_consent_update(session, text: str, current_url: str, log=None):
    """Consent/Update 同意页：用 canary/scope 接受。返回响应；无 ServerData 返回 None。

    两种页面形态：
    - 老形态：写 `ucaction` + `sClientId`/`sRawInputScopes`/`sCanary`；
    - 2026-09 绑辅助邮箱(id=293577)之后的新形态：ServerData 里是
      `arrConsentInfoServerData`（无 sCanary/sRawInputScopes），必须补上
      `sRawInputGrantedScopes`（取首个 scope 的 id)并把 `sCanary` 置空，
      否则页面下发的空 scope 会把已授权 scope 覆盖掉 → AAD 回 900144 bad request。
    """
    sd = parse_server_data(text)
    if not sd:
        return None
    graph_oauth_log(log, "接受 Outlook 授权同意页面")
    consent_form = {
        "ucaction": "Yes",
        "client_id": sd.get("sClientId", ""),
        "scope": sd.get("sRawInputScopes", ""),
        "cscope": sd.get("sRawInputGrantedScopes", ""),
        "canary": sd.get("sCanary", ""),
    }
    consent_info = sd.get("arrConsentInfoServerData")
    if sd.get("sCanary") in (None, "") and consent_info:
        raw_scopes = ""
        if isinstance(consent_info, list) and consent_info:
            client = consent_info[0] if isinstance(consent_info[0], dict) else {}
            scopes = client.get("arrScopes") or client.get("arrRawScopes") or []
            if isinstance(scopes, list):
                raw_scopes = " ".join(
                    str(s.get("id") or s.get("scope") or "").strip()
                    for s in scopes if isinstance(s, dict)
                ).strip()
        consent_form["cscope"] = raw_scopes
    resp = session.post(current_url, data=consent_form, timeout=30, allow_redirects=False)
    return resp


def is_dosubmit_bouncer(html: str) -> bool:
    """是否是微软的 DoSubmit(fmHF) 自动提交跳板页(登录后 / 绑定后的「Continue」页)。"""
    text = str(html or "")
    return ("DoSubmit" in text or ("fmHF" in text and "onload" in text)) and "action" in text


def submit_dosubmit_bouncer(session, html: str, url: str, log=None):
    """提交 DoSubmit(fmHF) 自动提交跳板页,返回响应;找不到 action 返回 None。

    hidden value 必须先 HTML 反转义(_unescape_form_fields),否则服务端收到坏 JSON
    → 302 oauth server_error(2026-09 发现的主坑)。
    """
    match = re.search(r"""action\s*=\s*["']([^"']+)["']""", str(html or ""))
    if not match:
        return None
    graph_oauth_log(log, "处理 Microsoft 中间自动提交页面")
    return session.post(
        absolute_form_action(match.group(1), url),
        data=_unescape_form_fields(html),
        timeout=30,
        allow_redirects=False,
    )


def _parse_cfg_json_balanced(raw: str) -> str:
    """从 $Config= 或 ServerData= 后的文本提取平衡 JSON（处理转义不破坏结构）。"""
    depth = 0
    instr = False
    esc = False
    for i, ch in enumerate(raw):
        if esc:
            esc = False
            continue
        if ch == '\\':
            esc = True
            continue
        if ch == '"':
            instr = not instr
            continue
        if instr:
            continue
        if ch == '{':
            depth += 1
        elif ch == '}':
            depth -= 1
            if depth == 0:
                return raw[:i + 1]
    return raw


def _skip_createfido(session, text: str, url: str, idx: int = 0,
                     log: Optional[Callable[[str], None]] = None):
    """CreateFido 页（强制 passkey 注册）纯协议跳过。

    JS 逻辑（ConvergedCreateFido_Core）：Skip 按钮 → form POST 到 $Config.urlPost，
    字段 canary($Config.sCanary) + error_code="Cancel" + i19。
    成功 → 302 回 oauth20_authorize；失败（未过 fido/create 状态）→ errcode=1078。
    返回响应；无 $Config / 无 urlPost 时返回 None（调用方继续走通用表单分支）。
    """
    m_cfg = re.search(r'\$Config\s*=\s*(\{.*?\});', str(text or ""), re.DOTALL)
    if not m_cfg:
        return None
    try:
        cfg = json.loads(_parse_cfg_json_balanced(m_cfg.group(1)))
    except Exception:
        return None
    url_post = str(cfg.get("urlPost", "") or "").replace("\\u0026", "&").replace("&amp;", "&")
    if not url_post:
        return None
    resp = session.post(
        url_post,
        data={"canary": cfg.get("sCanary", ""), "error_code": "Cancel", "i19": "3"},
        timeout=30,
        allow_redirects=False,
        headers={
            "Referer": url,
            "Origin": "https://login.microsoft.com",
            "Sec-Fetch-Dest": "document",
            "Sec-Fetch-Mode": "navigate",
            "Sec-Fetch-Site": "cross-site",
            "Sec-Fetch-User": "?F0D1",
            "Content-Type": "application/x-www-form-urlencoded",
        },
    )
    graph_oauth_log(
        log,
        f"[#{idx}] CreateFido skip -> {resp.status_code} "
        f"{(resp.headers.get('Location') or '')[:90]}",
    )
    return resp


def _resolve_bind_secondary(email: str, recovery_email: str, recovery_password: str,
                            log: Optional[Callable[[str], None]] = None) -> Optional[Dict[str, Any]]:
    """解析本号要绑的 CF 辅助邮箱，返回 {cf_address, cf_jwt, use_admin, cf_password}。

    优先用导入时带的辅助邮箱（recovery_email，收码走 CF admin API），
    否则 create_or_get_address 自动分配 ms-<前缀>@<CF域名>。
    失败返回 None（调用方回退 Skip 或不绑，不阻断授权）。
    """
    recovery_email = str(recovery_email or '').strip()
    if recovery_email:
        graph_oauth_log(log, f"使用导入的辅助邮箱绑定: {recovery_email}")
        return {
            "cf_address": recovery_email,
            "cf_jwt": None,
            "use_admin": True,
            "cf_password": str(recovery_password or ''),
        }
    try:
        cf_info = create_or_get_address(email)
    except Exception as exc:
        graph_oauth_log(log, f"CF 辅助邮箱分配失败，回退 Skip: {exc}")
        return None
    cf_address = cf_info.get("address")
    if not cf_address:
        return None
    return {
        "cf_address": cf_address,
        "cf_jwt": cf_info.get("jwt"),
        "use_admin": cf_info.get("use_admin", False),
        "cf_password": cf_info.get("password") or "",
    }


def _bind_credentialaction_in_session(session, html: str, url: str, email: str, *,
                                      bind_secondary: Any = None, idx: int = 0,
                                      log: Optional[Callable[[str], None]] = None,
                                      max_wait: int = 150, poll: int = 4):
    """credentialaction 中断页（mode=mpb）纯协议绑定辅助邮箱。

    2026-09 起未绑辅助邮箱的号登录后落到此页（老 proofs/Add 链路已废）。
    页面 ServerData：acmaInitialResponse.continuationToken + apiCanary。
    POST api/v1.0/auth/methods/email → CF 收码 → POST .../activate（200=绑定成功）。
    成功返回 True；失败返回 None。
    """
    tag = f"[#{idx}]"
    m_sd = re.search(r'ServerData\s*=\s*(\{.*?\});', str(html or ""), re.DOTALL)
    if not m_sd:
        graph_oauth_log(log, f"{tag} bind_ca: 无 ServerData，无法绑定")
        return None
    try:
        sd = json.loads(m_sd.group(1))
    except Exception:
        graph_oauth_log(log, f"{tag} bind_ca: ServerData 解析失败")
        return None
    canary = sd.get("apiCanary", "")
    cont_token = str((sd.get("acmaInitialResponse") or {}).get("continuationToken", "") or "")
    if not canary or not cont_token:
        graph_oauth_log(log, f"{tag} bind_ca: canary/continuationToken 缺失，无法绑定")
        return None

    # uaid：页面 hidden input（DoSubmit 表单里），作 correlationId 用
    uaid = _unescape_form_fields(html).get("uaid", "")

    # CF 辅助邮箱：调用方给了就直接用，否则现场分配
    bs = bind_secondary
    if bs and not isinstance(bs, dict):
        bs = None
    cf_address = str((bs or {}).get("cf_address") or "")
    cf_jwt = (bs or {}).get("cf_jwt")
    use_admin = bool((bs or {}).get("use_admin", False))
    if not cf_address:
        info = _resolve_bind_secondary(email, '', '', log)
        if not info:
            graph_oauth_log(log, f"{tag} bind_ca: 无可用 CF 辅助邮箱")
            return None
        cf_address = info["cf_address"]
        cf_jwt = info["cf_jwt"]
        use_admin = info["use_admin"]
    graph_oauth_log(log, f"{tag} bind_ca: 绑定辅助邮箱 {cf_address}")

    h_api = {
        "Accept": "application/json",
        "canary": canary,
        "correlationId": uaid,
        "client-request-id": uaid,
        "Content-Type": "application/json",
        "Referer": "https://account.live.com/interrupt/credentialaction",
        "Origin": "https://account.live.com",
        "Sec-Fetch-Dest": "empty",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Site": "same-origin",
    }
    resp1 = session.post(
        "https://account.live.com/api/v1.0/auth/methods/email",
        json={"email": cf_address, "continuationToken": cont_token},
        headers=h_api,
        timeout=30,
    )
    try:
        jd1 = resp1.json() if resp1.status_code in (200, 201) else {}
    except Exception:
        jd1 = {}
    graph_oauth_log(log, f"{tag} bind_ca: POST email -> {resp1.status_code} state={jd1.get('state', '')}")
    if jd1.get("error") or resp1.status_code not in (200, 201):
        graph_oauth_log(log, f"{tag} bind_ca: email API 失败 {str(jd1)[:150]}")
        return None
    if jd1.get("apiCanary"):
        h_api["canary"] = jd1["apiCanary"]

    # CF 收码：先取基线 id，只取基线之后的新邮件
    base_id = 0
    try:
        if use_admin:
            raws = fetch_admin_mails(cf_address, limit=20)
            base_id = max([m.get("id", 0) for m in raws] or [0])
        else:
            mails = fetch_parsed_mails(cf_jwt, limit=20)
            base_id = max([m.get("id", 0) for m in mails] or [0])
    except Exception as exc:
        graph_oauth_log(log, f"{tag} bind_ca: 取 CF 基线失败（继续）: {exc}")

    code = wait_for_code(cf_jwt, received_after_id=base_id, max_wait=max_wait, poll=poll,
                         use_admin=use_admin, address=cf_address)
    if not code:
        code = _fetch_latest_code_fallback(cf_address, cf_jwt, use_admin)
    if not code:
        graph_oauth_log(log, f"{tag} bind_ca: CF 取码超时")
        return None
    graph_oauth_log(log, f"{tag} bind_ca: 取到验证码 {code}")

    resp2 = session.post(
        "https://account.live.com/api/v1.0/auth/methods/email/activate",
        json={
            "activationDetails": {"displayName": cf_address, "id": cf_address, "otp": code},
            "otp": code,
            "continuationToken": cont_token,
        },
        headers=h_api,
        timeout=30,
    )
    ok = resp2.status_code in (200, 201)
    graph_oauth_log(log, f"{tag} bind_ca: activate -> {resp2.status_code} {'OK' if ok else str(getattr(resp2, 'text', '') or '')[:120]}")
    return True if ok else None


def extract_graph_refresh_token(
    email: str,
    password: str,
    *,
    client_id: str = OAUTH_CLIENT_ID,
    redirect_uri: str = OAUTH_REDIRECT_URI,
    scope: str = GRAPH_EXTRACT_SCOPE,
    authority: str = GRAPH_EXTRACT_AUTHORITY,
    log: Optional[Callable[[str], None]] = None,
    session_factory: Optional[Callable[[], Any]] = None,
    proxy_url: str = None,
    bind_secondary: Any = None,
    recovery_email: str = '',
    recovery_password: str = '',
) -> Dict[str, Any]:
    """使用纯 HTTP OAuth2 授权码流程提取 Outlook refresh_token。

    bind_secondary 真值时,流程到达需要绑辅助邮箱的页面会真绑（而非 Skip）：
    - proofs/Add（老链路）：bind_proof_in_session 填表 → CF 收码 → VerifyProof。
    - interrupt/credentialaction（2026-09 新链路，未绑号登录后落到此页）：
      _bind_credentialaction_in_session 走 auth/methods/email + activate 纯协议绑定。
    - recovery_email 非空（导入带了辅助邮箱）：直接用该地址绑定,跳过 CF 建邮箱,
      收码走 CF admin API（需 CF 渠道）；recovery_password 作为辅助邮箱密码回传。
    - 否则：create_or_get_address 建 CF 临时邮箱（ms-<前缀>@<CF域名>）绑定。
    成功后 recovery_email/recovery_email_password 随成功 dict 返回,由调用方透传给 upsert。

    流程还会纯协议处理 passkey 强制页（fido/create 跳板 + CreateFido Skip）与
    App/Confirm 页（GET successUrl）。
    """
    _pending_recovery_email = ""
    _pending_recovery_password = ""
    try:
        session = session_factory() if session_factory else requests.Session()
        resolved_proxy = str(proxy_url or '').strip()
        if resolved_proxy:
            proxies = build_proxies(resolved_proxy)
            if proxies:
                session.proxies.update(proxies)
            # 已配置应用代理时避免与环境代理叠加
            session.trust_env = False
        else:
            session.trust_env = True
        session.headers.update({
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/130.0.0.0 Safari/537.36"
            )
        })

        graph_oauth_log(log, f"获取 Microsoft 授权页面: {email}")
        log_outbound_proxy_usage(f'Outlook自动授权 {email}', resolved_proxy or '')
        if resolved_proxy:
            graph_oauth_log(log, f"OAuth 全程固定代理: {format_proxy_for_log(resolved_proxy)}")
        auth_url = build_graph_authorize_url(client_id, redirect_uri, scope, authority)
        resp = session.get(auth_url, timeout=30, allow_redirects=True)
        text = resp.text or ""

        flow_token = ""
        sft_tag = re.search(r'sFTTag.*?value=\\?"([^"\\]+)', text, re.DOTALL)
        if sft_tag:
            flow_token = sft_tag.group(1)
        if not flow_token:
            ppft = re.search(r'name="PPFT"[^>]*value="([^"]+)"', text, re.IGNORECASE)
            if ppft:
                flow_token = ppft.group(1)
        if not flow_token:
            return make_graph_oauth_response(False, "无法提取 Flow Token", "未在授权页面找到 PPFT 字段")

        post_url = ""
        urlpost_match = re.search(r'"urlPost"\s*:\s*"([^"]+)"', text)
        if urlpost_match:
            post_url = urlpost_match.group(1).replace("\\u0026", "&")
        if not post_url:
            post_url = "https://login.live.com/ppsecure/post.srf"

        ctx = ""
        sctx_match = re.search(r'"sCtx"\s*:\s*"([^"]+)"', text)
        if sctx_match:
            ctx = sctx_match.group(1)

        graph_oauth_log(log, "提交 Microsoft 登录凭据")
        resp2 = session.post(
            post_url,
            data={
                "login": email,
                "loginfmt": email,
                "passwd": password,
                "PPFT": flow_token,
                "ctx": ctx,
                "type": "11",
                "LoginOptions": "3",
                "i13": "0",
                "CookieDisclosure": "0",
                "IsFidoSupported": "0",
                "isSignupPost": "0",
                "i19": "16393",
            },
            timeout=30,
            allow_redirects=False,
        )

        # 检测登录失败的情况
        post_html = resp2.text or ""
        post_url_check = getattr(resp2, "url", "") or post_url

        if resp2.status_code == 200 and "ppsecure/post.srf" in post_url_check:
            # 情况1：检查JavaScript错误变量和HTML错误元素
            error_markers = [
                (r'sErrTxt["\s:=]+["\']([^"\']+)', "JavaScript错误信息"),
                (r'<div[^>]*id=["\']error["\'][^>]*>([^<]+)', "错误提示框"),
                (r'data-bind=["\']text:\s*unsafe_(\w+)["\']', "验证失败"),
                (r'<div[^>]*class=["\'][^"\']*error[^"\']*["\'][^>]*>([^<]+)', "错误样式"),
            ]

            for pattern, error_type in error_markers:
                match = re.search(pattern, post_html, re.IGNORECASE | re.DOTALL)
                if match:
                    error_detail = match.group(1).strip() if match.lastindex and len(match.groups()) > 0 else error_type
                    # 清理HTML标签
                    error_detail = re.sub(r'<[^>]+>', '', error_detail).strip()
                    return make_graph_oauth_response(
                        False,
                        "Microsoft 登录失败",
                        f"{error_type}: {graph_oauth_safe_details(error_detail)}"
                    )

            # 情况2：没有重定向且停留在post.srf，检查是否返回了登录表单
            if not resp2.headers.get("Location"):
                # 如果页面包含密码输入框，说明登录失败返回了登录页面
                if re.search(r'name=["\']passwd["\']', post_html, re.IGNORECASE):
                    # 尝试提取更具体的错误信息
                    specific_errors = [
                        (r'incorrect|invalid|wrong', "密码不正确或账号不存在"),
                        (r'verify|verification|confirm', "需要额外验证"),
                        (r'suspicious|unusual', "检测到异常活动"),
                        (r'disabled|locked|blocked', "账号被锁定或禁用"),
                    ]

                    error_hint = "密码不正确、账号不存在或需要额外验证"
                    for pattern, hint in specific_errors:
                        if re.search(pattern, post_html, re.IGNORECASE):
                            error_hint = hint
                            break

                    return make_graph_oauth_response(
                        False,
                        "登录凭据验证失败",
                        f"提交凭据后返回了登录表单，通常表示{error_hint}。请手动登录 https://outlook.live.com 确认账号状态。"
                    )

        # 登录后的重定向链可能夹着若干 DoSubmit 自动提交跳板页(Continue)，
        # 每跳一次算一步，故按「已提交次数」计上限而不是只试 5 次。
        for _ in range(MAX_OAUTH_STEPS):
            html = resp2.text or ""
            if is_dosubmit_bouncer(html):
                bounced = submit_dosubmit_bouncer(session, html, getattr(resp2, "url", "") or post_url, log)
                if bounced is not None:
                    resp2 = bounced
                    continue
            break

        auth_code = None
        for _ in range(MAX_OAUTH_STEPS):
            while resp2.status_code in (301, 302, 303, 307):
                loc = resp2.headers.get("Location", "")
                # 只有真正回到 redirect_uri 且带 code=/error= 才算终点；否则继续跟。
                # 不能用 "localhost" in loc 裸判：authorize URL 里 redirect_uri 是编码过的
                # (http%3a%2f%2flocalhost%3a8080)，会把微软中间跳转误当终点、拿空响应死循环。
                if is_oauth_code_redirect(loc):
                    resp2 = make_light_response(loc)
                    break
                if not loc:
                    break
                resp2 = session.get(loc, timeout=30, allow_redirects=False)

            current_url = getattr(resp2, "url", "") or ""
            text = resp2.text if getattr(resp2, "text", "") else ""

            if is_oauth_code_redirect(current_url):
                params = urllib.parse.parse_qs(urllib.parse.urlparse(current_url).query)
                if params.get("error"):
                    err = params.get("error_description", params.get("error", ["?"]))[0]
                    return make_graph_oauth_response(False, "OAuth 错误", err)
                auth_code = params.get("code", [None])[0]
                if auth_code:
                    graph_oauth_log(log, "已捕获授权码")
                    break

            # DoSubmit(fmHF) 自动提交跳板（登录后 / 绑定后的 "Continue" 页）。
            # 必须先于 Consent 判定：绑定成功重新 POST authorize 后会先落到
            # Consent/Update 的跳板壳（HTML 转义 hidden + DoSubmit）上，把这层壳提交掉
            # 才能看到真正的落点。判定对齐 reg-factory（DoSubmit 或 fmHF+onload）。
            if is_dosubmit_bouncer(text):
                bounced = submit_dosubmit_bouncer(session, text, current_url, log)
                if bounced is not None:
                    resp2 = bounced
                    continue

            if "Consent/Update" in current_url or "Consent/update" in current_url:
                consent_resp = handle_consent_update(session, text, current_url, log)
                if consent_resp is None:
                    return make_graph_oauth_response(False, "同意页面处理失败", "无法解析 ServerData")
                resp2 = consent_resp
                continue

            if "interrupt/credentialaction" in current_url:
                # 2026-09 强制中断页：未绑辅助邮箱的号登录后落到此页（老 proofs/Add 链路已废）。
                # 纯协议绑定：apiCanary+continuationToken → POST auth/methods/email
                # → CF 收码 → activate → 重新 GET authorize 跟链。
                if not bind_secondary:
                    # 用户显式关掉了「绑定辅助邮箱」：不偷偷建 CF 邮箱,明确报出必须先绑。
                    return make_graph_oauth_response(
                        False, "需要绑定辅助邮箱",
                        "账号落到微软强制中断页(credentialaction)，未绑辅助邮箱必须绑定后才能授权；"
                        "请开启「绑定辅助邮箱」后重试",
                    )
                resolved = _resolve_bind_secondary(email, recovery_email, recovery_password, log)
                ca_bind_ok = _bind_credentialaction_in_session(
                    session, text, current_url, email,
                    bind_secondary=resolved, log=log,
                )
                if ca_bind_ok is None:
                    return make_graph_oauth_response(
                        False, "辅助邮箱绑定失败", "中断页(credentialaction)绑定辅助邮箱未完成"
                    )
                if resolved and resolved.get("cf_address"):
                    _pending_recovery_email = resolved["cf_address"]
                    _pending_recovery_password = resolved.get("cf_password") or ""
                # 绑定成功：重新 GET authorize（沿用现有登录态 → 落到 Consent 跳板页）。
                # 必须用 GET：POST authorize 不带 body 会被 AAD 拒
                # （AADSTS900144: request body must contain 'client_id'）。
                # 这一跳大概率直接 302 回 login.live.com/oauth20_authorize，
                # 循环顶部跟过去后是 DoSubmit(fmHF) 跳板页，由循环顶部的
                # is_dosubmit_bouncer 处理提交掉，再继续跟到 Consent/授权码。
                resp2 = session.get(auth_url, timeout=30, allow_redirects=False)
                continue

            # passkey 强制中断：fido/create 是自动提交跳板，POST 后继续跟链
            if "fido/create" in text and "onload" in text:
                fido_action = re.search(r"action='([^']*)'", text)
                if fido_action:
                    fido_url = fido_action.group(1).replace("&amp;", "&")
                    resp2 = session.post(
                        absolute_form_action(fido_url, current_url),
                        data=_unescape_form_fields(text),
                        timeout=30,
                        allow_redirects=False,
                    )
                    graph_oauth_log(log, f"fido/create POST -> {(getattr(resp2, 'url', '') or '')[:100]}")
                    continue

            # CreateFido 页：Skip = POST urlPost + canary + error_code=Cancel
            if "CreateFido" in text or ("$Config" in text and "sFidoChallenge" in text):
                skip_resp = _skip_createfido(session, text, current_url, 0, log)
                if skip_resp is not None:
                    resp2 = skip_resp
                    continue

            # App/Confirm 页：点 Continue → GET successUrl（带 res=success）
            if "App/Confirm" in current_url:
                success_match = re.search(r'"successUrl"\s*:\s*"([^"]+)"', text)
                if success_match:
                    success_url = success_match.group(1).replace("\\u0026", "&").replace("\\u002f", "/")
                    success_url = success_url.replace("&amp;", "&")
                    graph_oauth_log(log, f"App/Confirm -> GET successUrl ...{success_url[-60:]}")
                    resp2 = session.get(success_url, timeout=30, allow_redirects=False)
                    continue
                return make_graph_oauth_response(False, "App/Confirm 处理失败", "页面未找到 successUrl")

            if "proofs/Add" in current_url or "proofs/add" in current_url:
                if bind_secondary:
                    resolved = _resolve_bind_secondary(email, recovery_email, recovery_password, log)
                    if resolved and resolved.get("cf_address"):
                        bound_resp = bind_proof_in_session(
                            session, text, current_url,
                            cf_address=resolved["cf_address"], cf_jwt=resolved.get("cf_jwt"),
                            use_admin=resolved.get("use_admin", False), idx=0,
                        )
                        if bound_resp is not None:
                            resp2 = bound_resp
                            _pending_recovery_email = resolved["cf_address"]
                            _pending_recovery_password = resolved.get("cf_password") or ""
                            continue
                        graph_oauth_log(log, "bind 失败，回退 Skip proofs/Add")
                # 回退 / 未启用绑定：原 Skip 逻辑
                form_match = re.search(
                    r'<form[^>]*action="([^"]+)"[^>]*>(.*?)</form>',
                    text,
                    re.DOTALL | re.IGNORECASE,
                )
                if not form_match:
                    return make_graph_oauth_response(False, "安全信息页面处理失败", "无法找到表单")
                graph_oauth_log(log, "跳过 Microsoft 安全信息添加页面")
                form_data = _unescape_form_fields(form_match.group(2))
                form_data["action"] = "Skip"
                resp2 = session.post(
                    absolute_form_action(form_match.group(1), current_url),
                    data=form_data,
                    timeout=30,
                    allow_redirects=False,
                )
                continue

            form_match = re.search(
                r'<form[^>]*action="([^"]+)"[^>]*>(.*?)</form>',
                text,
                re.DOTALL | re.IGNORECASE,
            )
            if form_match:
                form_action = absolute_form_action(form_match.group(1), current_url)
                form_data = _unescape_form_fields(form_match.group(2))
                if "consent" in form_action.lower() or "consent" in current_url.lower():
                    graph_oauth_log(log, "提交通用同意表单")
                    form_data["ucaccept"] = "Yes"

                resp2 = session.post(form_action, data=form_data, timeout=30, allow_redirects=False)
                while resp2.status_code in (301, 302, 303, 307):
                    loc = resp2.headers.get("Location", "")
                    if is_oauth_code_redirect(loc):
                        resp2 = make_light_response(loc)
                        break
                    if not loc:
                        break
                    resp2 = session.get(loc, timeout=30, allow_redirects=False)
                continue

            return make_graph_oauth_response(
                False,
                "授权流程卡住",
                f"在 {current_url[:100]} 无法继续 (status={resp2.status_code})",
            )

        if not auth_code:
            return make_graph_oauth_response(False, "未能获取授权码", "完成所有步骤但未捕获到授权码")

        graph_oauth_log(log, "使用授权码换取 Outlook token")
        token_resp = session.post(
            f"https://login.microsoftonline.com/{authority}/oauth2/v2.0/token",
            data={
                "client_id": client_id,
                "grant_type": "authorization_code",
                "code": auth_code,
                "redirect_uri": redirect_uri,
                "scope": scope,
            },
            timeout=30,
        )
        token_data = token_resp.json()

        if "access_token" not in token_data:
            err = token_data.get("error_description", token_data.get("error", "?"))
            return make_graph_oauth_response(False, "Token 换取失败", err)

        refresh_token = str(token_data.get("refresh_token") or "").strip()
        if not refresh_token:
            return make_graph_oauth_response(False, "未获取到 refresh_token", "响应中包含 access_token 但没有 refresh_token")

        graph_oauth_log(log, "已获取 Outlook refresh_token")
        return {
            "success": True,
            "refresh_token": refresh_token,
            "client_id": client_id,
            "recovery_email": _pending_recovery_email,
            "recovery_email_password": _pending_recovery_password,
        }
    except Exception as exc:
        return make_graph_oauth_response(False, f"异常: {type(exc).__name__}", str(exc))


def get_upload_account_for_graph_auth(account_id: int):
    db = get_db()
    return db.execute(
        '''
        SELECT id, email, password, is_authorized, remark, group_id, proxy_url, tag_ids,
               recovery_email, recovery_email_password
        FROM outlook_upload_accounts
        WHERE id = ?
        ''',
        (account_id,),
    ).fetchone()


def upsert_graph_authorized_account(email: str, password: str, client_id: str,
                                    refresh_token: str, *,
                                    group_id: Any = None,
                                    proxy_url: str = '',
                                    tag_ids: Any = None,
                                    remark: str = '',
                                    authorization_type: Optional[str] = None,
                                    recovery_email: str = '',
                                    recovery_email_password: str = '') -> Dict[str, Any]:
    db = get_db()
    existing = db.execute(
        'SELECT id, authorization_type FROM accounts WHERE LOWER(email) = ? LIMIT 1',
        (normalize_email_address(email),),
    ).fetchone()
    encrypted_password = encrypt_data(password) if password else password
    encrypted_refresh_token = encrypt_data(refresh_token) if refresh_token else refresh_token
    encrypted_recovery_password = encrypt_data(recovery_email_password) if recovery_email_password else recovery_email_password
    if authorization_type is None:
        normalized_authorization_type = normalize_outlook_authorization_type(
            existing['authorization_type'] if existing else ''
        )
    else:
        normalized_authorization_type = normalize_outlook_authorization_type(
            authorization_type,
            strict=True,
        )

    if existing:
        account_id = int(existing['id'])
        # 已有正式账号：仅覆盖授权相关字段，保留分组/标签/代理等业务字段
        db.execute(
            '''
            UPDATE accounts
            SET password = ?,
                client_id = ?,
                refresh_token = ?,
                account_type = 'outlook',
                provider = 'outlook',
                authorization_type = ?,
                recovery_email = ?,
                recovery_email_password = ?,
                refresh_token_updated_at = CURRENT_TIMESTAMP,
                last_refresh_status = 'never',
                last_refresh_error = NULL,
                updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
            ''',
            (encrypted_password, client_id, encrypted_refresh_token, normalized_authorization_type,
             recovery_email, encrypted_recovery_password, account_id),
        )
        return {"account_id": account_id, "created": False}

    resolved_group_id = resolve_upload_group_id(group_id)
    normalized_proxy = str(proxy_url or '').strip()
    cursor = db.execute(ACCOUNT_INSERT_SQL, build_account_insert_values(
        normalize_email_address(email),
        password,
        client_id,
        refresh_token,
        resolved_group_id,
        remark or '',
        'outlook',
        'outlook',
        IMAP_SERVER_NEW,
        IMAP_PORT,
        '',
        False,
        None,
        'active',
        normalized_proxy,
        '',
        '',
        recovery_email,
        recovery_email_password,
    ))
    account_id = int(cursor.lastrowid)
    apply_account_tag_ids(account_id, tag_ids, db)
    db.execute(
        '''
        UPDATE accounts
        SET authorization_type = ?,
            refresh_token_updated_at = CURRENT_TIMESTAMP,
            last_refresh_status = 'never',
            last_refresh_error = NULL,
            updated_at = CURRENT_TIMESTAMP
        WHERE id = ?
        ''',
        (normalized_authorization_type, account_id),
    )
    return {"account_id": account_id, "created": True}


def mark_upload_account_authorized(account_id: int) -> None:
    get_db().execute(
        '''
        UPDATE outlook_upload_accounts
        SET is_authorized = 1,
            updated_at = CURRENT_TIMESTAMP
        WHERE id = ?
        ''',
        (account_id,),
    )


def save_graph_authorization_result(upload_row: Any, client_id: str,
                                    refresh_token: str,
                                    authorization_type: Optional[str] = None,
                                    recovery_email: str = '',
                                    recovery_email_password: str = '') -> Dict[str, Any]:
    email = str(upload_row['email'] or '').strip()
    password = get_upload_account_plain_password(upload_row)
    row_data = dict(upload_row) if hasattr(upload_row, 'keys') else {}
    # recovery 优先用流程返回值（走绑定页时=实际绑上去的邮箱，如建的 CF 邮箱）；
    # 流程没返回（没进 proofs/Add，如账号已绑过）则回退到导入时入库的 recovery，
    # 保证导入带的辅助邮箱不会在转正式表时丢失。
    final_recovery_email = (recovery_email or '').strip() or str(row_data.get('recovery_email') or '').strip()
    final_recovery_password = recovery_email_password
    if not final_recovery_password:
        enc_pw = str(row_data.get('recovery_email_password') or '')
        if enc_pw:
            try:
                final_recovery_password = decrypt_data(enc_pw) or ''
            except Exception:
                final_recovery_password = ''
    save_result = upsert_graph_authorized_account(
        email,
        password,
        client_id,
        refresh_token,
        group_id=row_data.get('group_id'),
        proxy_url=row_data.get('proxy_url') or '',
        tag_ids=decode_upload_tag_ids(row_data.get('tag_ids')),
        remark=str(row_data.get('remark') or ''),
        authorization_type=authorization_type,
        recovery_email=final_recovery_email,
        recovery_email_password=final_recovery_password,
    )
    mark_upload_account_authorized(int(upload_row['id']))
    get_db().commit()
    return save_result


def run_graph_oauth_task(account_id: int, output_queue: "queue.Queue[Dict[str, Any] | object]",
                         mode: str = "graph", bind_secondary: Any = None) -> None:
    def emit(payload: Dict[str, Any]) -> None:
        output_queue.put(payload)

    def log(message: str) -> None:
        emit({"type": "log", "message": graph_oauth_safe_details(message)})

    with app.app_context():
        try:
            mode = normalize_graph_oauth_mode(mode)
            upload_row = get_upload_account_for_graph_auth(account_id)
            if not upload_row:
                emit({"type": "error", "success": False, "mode": mode, "message": "上传账号不存在"})
                emit({"type": "complete", "success": False})
                return

            email = str(upload_row['email'] or '').strip()
            password = get_upload_account_plain_password(upload_row)
            if not email or not password:
                emit({"type": "error", "success": False, "mode": mode, "message": "邮箱或密码为空"})
                emit({"type": "complete", "success": False})
                return

            # 导入时若带了辅助邮箱，授权到 proofs/Add 直接用它绑定，跳过 CF 建邮箱。
            recovery_email = str(upload_row['recovery_email'] or '').strip()
            recovery_password = ''
            recovery_password_enc = str(upload_row['recovery_email_password'] or '')
            if recovery_email and recovery_password_enc:
                try:
                    recovery_password = decrypt_data(recovery_password_enc) or ''
                except Exception:
                    recovery_password = ''

            mode_label = graph_oauth_mode_label(mode)
            scope = GRAPH_EXTRACT_SCOPE_BY_MODE[mode]
            proxy_config = get_upload_account_resolved_proxy_config(upload_row)
            # OAuth 多跳必须固定同一主代理；不做中途 failover
            auth_proxy_url = proxy_config.get('proxy_url', '') or ''
            emit({
                "type": "start",
                "email": email,
                "mode": mode,
                "message": f"开始 {mode_label} OAuth 授权",
            })
            log(f"授权模式: {mode_label}")
            log(f"授权 Scope: {scope}")
            if auth_proxy_url:
                log("使用上传账号/分组代理进行自动授权")
            result = extract_graph_refresh_token(
                email,
                password,
                scope=scope,
                log=log,
                proxy_url=auth_proxy_url,
                bind_secondary=bind_secondary,
                recovery_email=recovery_email,
                recovery_password=recovery_password,
            )
            if not result.get("success"):
                emit({
                    "type": "error",
                    "success": False,
                    "mode": mode,
                    "message": graph_oauth_safe_details(result.get("error") or "授权失败"),
                    "details": graph_oauth_safe_details(result.get("details") or ""),
                })
                emit({"type": "complete", "success": False})
                return

            client_id = str(result.get("client_id") or "").strip()
            refresh_token = str(result.get("refresh_token") or "").strip()
            log(f"验证 {mode_label} refresh_token")
            refresh_result = test_refresh_token(
                client_id,
                refresh_token,
                proxy_url=auth_proxy_url,
                authorization_type=mode,
            )
            try:
                ok, error_msg, rotated_refresh_token, actual_channel = refresh_result
            except (TypeError, ValueError):
                ok, error_msg, rotated_refresh_token = refresh_result
                actual_channel = mode
            actual_channel = normalize_outlook_authorization_type(actual_channel)
            if not ok:
                emit({
                    "type": "error",
                    "success": False,
                    "mode": mode,
                    "message": f"{mode_label} refresh_token 验证失败",
                    "details": graph_oauth_safe_details(error_msg),
                })
                emit({"type": "complete", "success": False})
                return

            token_to_save = rotated_refresh_token or refresh_token
            recovery_email = str(result.get("recovery_email") or "")
            recovery_email_password = str(result.get("recovery_email_password") or "")
            save_result = save_graph_authorization_result(
                upload_row,
                client_id,
                token_to_save,
                authorization_type=actual_channel or mode,
                recovery_email=recovery_email,
                recovery_email_password=recovery_email_password,
            )
            emit({
                "type": "success",
                "success": True,
                "mode": mode,
                "authorization_type": actual_channel or mode,
                "email": email,
                "account_id": save_result["account_id"],
                "created": save_result["created"],
                "client_id": client_id,
                "message": "授权成功，已保存到正式账号",
            })
            emit({"type": "complete", "success": True})
        except Exception as exc:
            try:
                get_db().rollback()
            except Exception:
                pass
            emit({
                "type": "error",
                "success": False,
                "mode": normalize_graph_oauth_mode(mode),
                "message": "授权任务异常",
                "details": graph_oauth_safe_details(str(exc)),
            })
            emit({"type": "complete", "success": False})
        finally:
            output_queue.put(GRAPH_OAUTH_DONE)


def run_batch_oauth_task(account_ids, output_queue, *, mode="graph", bind_secondary=False, max_workers=5):
    """并行复用单账号 run_graph_oauth_task 处理一批上传账号。

    每个账号起一个子 queue 收集单账号 task 的 SSE 载荷,捕获最后一个 success/error
    payload 作为该账号的结果;逐 future 回调进度,结束后发 complete 汇总到
    output_queue,并返回汇总 dict。单账号失败不阻断整批。
    """
    total = len(account_ids)
    output_queue.put({
        "type": "start",
        "total": total,
        "mode": normalize_graph_oauth_mode(mode),
        "bind_secondary": bool(bind_secondary),
    })

    def _do_one(account_id):
        sub_q: "queue.Queue[Dict[str, Any] | object]" = queue.Queue()
        last_payload: Dict[str, Any] = {"type": "error", "success": False, "account_id": account_id}
        try:
            run_graph_oauth_task(account_id, sub_q, mode=mode, bind_secondary=bind_secondary)
            while True:
                payload = sub_q.get()
                if payload is GRAPH_OAUTH_DONE:
                    break
                if isinstance(payload, dict):
                    last_payload = payload
        except Exception as exc:
            last_payload = {
                "type": "error",
                "success": False,
                "account_id": account_id,
                "message": graph_oauth_safe_details(str(exc)),
            }
        return account_id, last_payload

    if not account_ids:
        summary = {"total": 0, "success_count": 0, "failed": []}
        output_queue.put({"type": "complete", **summary})
        return summary

    workers = min(max(1, max_workers), len(account_ids))
    workers = min(workers, 20)
    completed_index = 0
    success_count = 0
    failed: list = []

    with ThreadPoolExecutor(max_workers=workers, thread_name_prefix='oauth-batch') as executor:
        future_map = {executor.submit(_do_one, aid): aid for aid in account_ids}
        for future in as_completed(future_map):
            account_id = future_map[future]
            completed_index += 1
            try:
                aid, result = future.result()
            except Exception as exc:
                aid = account_id
                result = {
                    "type": "error",
                    "success": False,
                    "account_id": account_id,
                    "message": graph_oauth_safe_details(str(exc)),
                }
            is_success = bool(result.get("success")) if isinstance(result, dict) else False
            if is_success:
                success_count += 1
            else:
                failed.append({
                    "account_id": aid,
                    "error": graph_oauth_safe_details(
                        str(result.get("message") or result.get("error") or "")
                        if isinstance(result, dict) else "未知错误"
                    ),
                })
            output_queue.put({
                "type": "progress",
                "index": completed_index,
                "total": total,
                "account_id": aid,
                "success": is_success,
            })

    summary = {"total": total, "success_count": success_count, "failed": failed}
    output_queue.put({"type": "complete", **summary})
    return summary


@app.route('/api/oauth/graph-extract-token', methods=['POST'])
@login_required
def api_graph_extract_token():
    data = request.get_json(silent=True) or {}
    raw_account_id = data.get('account_id')
    try:
        account_id = int(raw_account_id)
    except (TypeError, ValueError):
        return jsonify({'success': False, 'error': 'account_id 不能为空'}), 400

    row = get_upload_account_for_graph_auth(account_id)
    if not row:
        return jsonify({'success': False, 'error': '上传账号不存在'}), 404
    if not str(row['email'] or '').strip() or not str(row['password'] or ''):
        return jsonify({'success': False, 'error': '邮箱或密码为空'}), 400

    mode = normalize_graph_oauth_mode(data.get('mode'))
    bind_secondary = bool(data.get('bind_secondary', True))
    task_id = uuid.uuid4().hex
    GRAPH_OAUTH_TASKS[task_id] = {'account_id': account_id, 'mode': mode, 'bind_secondary': bind_secondary}
    return jsonify({
        'success': True,
        'task_id': task_id,
        'mode': mode,
        'stream_url': f'/api/oauth/graph-extract-token/{task_id}/stream',
    })


@app.route('/api/oauth/graph-extract-token/<task_id>/stream', methods=['GET'])
@login_required
def api_graph_extract_token_stream(task_id: str):
    task = GRAPH_OAUTH_TASKS.pop(task_id, None)
    if not task:
        return Response(
            graph_oauth_sse({'type': 'error', 'success': False, 'message': '授权任务不存在或已过期'})
            + graph_oauth_sse({'type': 'complete', 'success': False}),
            mimetype='text/event-stream',
        )

    def generate():
        output_queue: "queue.Queue[Dict[str, Any] | object]" = queue.Queue()
        worker = threading.Thread(
            target=run_graph_oauth_task,
            args=(int(task['account_id']), output_queue, normalize_graph_oauth_mode(task.get('mode'))),
            kwargs={'bind_secondary': task.get('bind_secondary')},
            name=f"graph-oauth-{task_id[:8]}",
            daemon=True,
        )
        worker.start()

        while True:
            payload = output_queue.get()
            if payload is GRAPH_OAUTH_DONE:
                break
            yield graph_oauth_sse(payload)
        worker.join(timeout=1)

    return Response(stream_with_context(generate()), mimetype='text/event-stream')


@app.route('/api/oauth/graph-extract-batch', methods=['POST'])
@login_required
def api_graph_extract_batch():
    data = request.get_json(silent=True) or {}
    account_ids = data.get('account_ids') or []
    if not isinstance(account_ids, list) or not account_ids:
        return jsonify({'success': False, 'error': 'account_ids 不能为空'}), 400
    try:
        account_ids = [int(aid) for aid in account_ids]
    except (TypeError, ValueError):
        return jsonify({'success': False, 'error': 'account_ids 包含非法值'}), 400
    mode = normalize_graph_oauth_mode(data.get('mode'))
    # 默认 True(符合设计:批量授权自动分配 CF 辅助邮箱);显式传 False 可关闭
    bind_secondary = bool(data.get('bind_secondary', True))
    try:
        max_workers = min(20, max(1, int(data.get('max_workers', 5))))
    except (TypeError, ValueError):
        max_workers = 5
    task_id = uuid.uuid4().hex
    GRAPH_OAUTH_TASKS[task_id] = {
        'account_ids': account_ids,
        'mode': mode,
        'bind_secondary': bind_secondary,
        'max_workers': max_workers,
    }
    return jsonify({
        'success': True,
        'task_id': task_id,
        'stream_url': f'/api/oauth/graph-extract-batch/{task_id}/stream',
    })


@app.route('/api/oauth/graph-extract-batch/<task_id>/stream')
@login_required
def api_graph_extract_batch_stream(task_id: str):
    task = GRAPH_OAUTH_TASKS.pop(task_id, None)
    if not task:
        return Response(
            graph_oauth_sse({'type': 'error', 'success': False, 'message': '任务不存在或已完成'})
            + graph_oauth_sse({'type': 'complete', 'success': False}),
            mimetype='text/event-stream',
        )

    def generate():
        out_q: "queue.Queue[Dict[str, Any] | object]" = queue.Queue()
        worker = threading.Thread(
            target=run_batch_oauth_task,
            args=(task['account_ids'], out_q),
            kwargs={
                'mode': task['mode'],
                'bind_secondary': task['bind_secondary'],
                'max_workers': task['max_workers'],
            },
            name=f"oauth-batch-{task_id[:8]}",
            daemon=True,
        )
        worker.start()

        while True:
            try:
                payload = out_q.get(timeout=120)
            except queue.Empty:
                yield graph_oauth_sse({'type': 'ping'})
                continue
            if payload is GRAPH_OAUTH_DONE:
                break
            yield graph_oauth_sse(payload)
            if isinstance(payload, dict) and payload.get('type') == 'complete':
                break
        worker.join(timeout=1)

    return Response(stream_with_context(generate()), mimetype='text/event-stream')
