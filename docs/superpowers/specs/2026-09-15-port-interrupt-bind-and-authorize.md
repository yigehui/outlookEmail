# 对齐 reg-factory 授权改动：DoSubmit unescape + 中断页纯协议绑定（passkey/AppConfirm）

Date: 2026-09-15
Status: Implemented — 真号已验证（未绑号 fresh bind / 已绑号回归均通过）
Scope: `outlookEmail` repo (`D:\officeProject\yigehui\outlookEmail`)
来源改动: `reg-factory` 提交 `a650a8f`、`0e3b2b3`（另有 `1f960e9` 为 GitHub 注册链路，不适用）

## 背景

2026-08 起微软对新号收紧 Graph 授权：没有辅助邮箱的账号走完登录后不再落到老的
`proofs/Add`，而是 302 到 `account.live.com/interrupt/credentialaction`（id=293577, mode=mpb）
强制中断页；另有 passkey 强制注册页与 `App/Confirm` 页。outlookEmail 的
`extract_graph_refresh_token`（`11_routes_graph_oauth.py:124`）是 8 月从 reg-factory
移植的，因此**同样落在这些新页面上会「授权流程卡住」失败**。

reg-factory 这两天把这段跑通了，根因和三个新落点都解掉了：

1. **真正的主坑：`DoSubmit`(`fmHF`) 表单的 hidden value 是 HTML 转义的**
   （`&quot;` `&amp;` `&#39;`）。浏览器提交前会解码，纯 HTTP 必须先 `html.unescape`
   再 POST，否则服务端收到坏 JSON（`scenarios` 字段被破坏）→ 302 `oauth server_error`。
   此前的「PX / TLS / cookie」结论是误判。
2. **`interrupt/credentialaction` 纯协议绑定**：页面 `ServerData` 里有
   `apiCanary` + `acmaInitialResponse.continuationToken` →
   `POST api/v1.0/auth/methods/email` → CF 收码 → `POST .../email/activate`（200=绑定成功）
   → 重新 GET `authorize` 跟链。不再需要浏览器兜底。
3. **passkey 跳板 + `CreateFido` 页 Skip**：`fido/create` 是自动提交跳板（POST 后继续），
   落到 `CreateFido` 页时点 Skip = `POST urlPost` + `canary` + `error_code=Cancel` + `i19=3`
   （直接发 `postBackUrl` 会 errcode=1078）。
4. **`App/Confirm` 页**：GET `successUrl`（带 `res=success`）即继续，无需 POST。

outlookEmail 当前这三个落点全部未处理（`grep interrupt|credentialaction|CreateFido|App/Confirm`
只命中注释），且 DoSubmit 提交仍在用未 unescape 的 `extract_hidden_inputs`。因此**本项目现在
对新号授权是失败状态**，需要同步。

## 关键架构事实（已核对）

- 授权主循环：`extract_graph_refresh_token`（`11:124`），本地 helper
  `extract_hidden_inputs`（`11:102`，未 unescape）、`absolute_form_action`（`11:113`）、
  `make_light_response`（`11:93`）；日志 `graph_oauth_log`（`11:65`，经
  `sanitize_error_details` 截断 500）。
- 单位约定与 reg-factory 不同：本项目 `idx` 恒为 `0`，日志前缀 `[#0]`（现网如此），本次不改。
- **落点判断方式不同（重要的坑）**：本项目 `FakeSession.get(url, **kwargs)` 会**丢掉**
  `status_code`/`Location`，且 `url` 跟随重定向后的值；因此现有代码一律用
  `getattr(resp2, "url", "")` 做 URL 判断、用 `resp2.text` 做页面判断。
  → 新增落点判断必须沿用「`current_url` 判 URL + `text` 判页面」的写法，不能读
  `resp2.status_code` / `resp2.headers`（会破坏已有单测与录制回放）。
- 绑定已就绪：`bind_proof_in_session`（`13:59`，已有单测）、`_parse_proof_add_form` /
  `_parse_proof_verify_form`（`13:20/:36`）、`_fetch_latest_code_fallback`（`13:154`）、
  CF 收码裸名（`wait_for_code`/`fetch_admin_mails`/`fetch_parsed_mails`/`parse_admin_mail`/
  `extract_code`，`12_cloudflare_mail.py`）。
- `recovery_email`/`recovery_email_password` 从流程结果经
  `save_graph_authorization_result` → `upsert_graph_authorized_account` 落主表，已通。

## 改造内容

### Part A — 移植 reg-factory `0e3b2b3` 的 `extract_graph_tokens.py` 改动

**A1. `_unescape_form_fields`（`11_routes_graph_oauth.py` 新增，顶层 def）**

对齐 reg-factory 实现：正则取 `input[name][value]`（同名保留首个），对 value 做
`html.unescape`。**替换 3 处 `extract_hidden_inputs(...)` 调用点**，与 reg-factory 逐处对齐：

| 位置 | 现状 | 改为 |
|---|---|---|
| `11:283`（登录后 DoSubmit 跳板） | `extract_hidden_inputs(html)` | `_unescape_form_fields(html)` |
| `11:377`（proofs/Add Skip 分支） | `extract_hidden_inputs(form_match.group(2))` | `_unescape_form_fields(form_match.group(2))` |
| `11:394`（通用 form 提交） | `extract_hidden_inputs(form_match.group(2))` | `_unescape_form_fields(form_match.group(2))` |

`extract_hidden_inputs` 保留（`13_oauth_bind.py` 与其它调用点/测试仍在用），不删。
补充说明：`13_oauth_bind.py` 的 `bind_proof_in_session` 内部两处 hidden 解析与 reg-factory
**逐字相同（同样不 unescape）**，本次一并对齐为不 unescape，保持与上游同构。

**A2. `CreateFido` 页 Skip — 顶层 def `_skip_createfido(session, text, url, idx=0)`**

移植 reg-factory（`extract_graph_tokens.py` 内 `_skip_createfido`）：从 `$Config`（用平衡
括号提取器保证 JSON 完整）取 `urlPost`（还原 `&`/`&amp;`）与 `sCanary`，POST
`{"canary", "error_code": "Cancel", "i19": "3"}`，`allow_redirects=False`，带 Referer/Origin/
Sec-Fetch-* 头；成功 302 回 `oauth20_authorize`；返回响应或 `None`（无 `$Config`/无 `urlPost`）。
日志经 `graph_oauth_log(log, ...)` 走（本函数需要 `log` 参数以复用项目日志收口）。

**A3. `_bind_credentialaction_in_session(...)` — 顶层 def**

移植 reg-factory 同名函数，接口本地化：

```
_bind_credentialaction_in_session(session, html, url, email, *, bind_secondary=None,
                                  idx=0, log=None, max_wait=150, poll=4)
```

- 从 `ServerData` 取 `apiCanary` + `acmaInitialResponse.continuationToken`；
  缺任一 → 返回 `None`（日志标明原因；不落盘 debug HTML，对齐项目未配 debug 目录的现状）。
- `uaid` 取 `_unescape_form_fields(html)["uaid"]`（作为 `correlationId`/`client-request-id`）。
- CF 辅助邮箱解析顺序（**与现有 `11:337-355` 的语义保持一致**）：
  1. `recovery_email`（导入时带了辅助邮箱，用调用方给的 `bind_secondary` 里的地址）
  2. 否则 `create_or_get_address(email)`（自动分配 `ms-<前缀>@<CF域名>`）
  - `bind_secondary` 复用 `extract_graph_refresh_token` 现有的两种形态：`True`（bool）或
    带 `cf_address` 的 dict。为此新增内部归一化函数
    `_resolve_bind_secondary(email, recovery_email, recovery_password, log)`，
    返回 `{"cf_address", "cf_jwt", "use_admin", "cf_password"} | None`，
    **`11:333-366`（proofs/Add 分支）与 credentialaction 分支共用**，消除重复且行为一致。
- `POST https://account.live.com/api/v1.0/auth/methods/email`（json `{email, continuationToken}`），
  成功后用返回的 `apiCanary` 覆盖请求头 canary。
- CF 收码：先取基线 id（`use_admin` → `fetch_admin_mails`，否则 `fetch_parsed_mails`），
  再 `wait_for_code(...)`，超时走 `_fetch_latest_code_fallback`。
- `POST .../auth/methods/email/activate`（json `{activationDetails:{displayName,id,otp}, otp, continuationToken}`），
  200/201 → 返回 `True`；否则 `None`。
- 成功时把实际绑上的地址经 `_pending_recovery_email` / `_pending_recovery_password` 回传
  （沿用 `extract_graph_refresh_token` 现有机制，从而落主表 `recovery_email`）。

**A4. 主循环新增 4 个落点分支（`11:289` 的 15 步循环内，顺序对齐 reg-factory）**

在 `Consent/Update` 之后、`proofs/Add` 之前依次插入：

1. `if "interrupt/credentialaction" in current_url:` → 调 A3；失败 → 返回
   「辅助邮箱绑定失败」(details 说明中断页绑定失败)；成功 → `session.get(auth_url)` 重新跟链 `continue`。
2. `if "fido/create" in text and "onload" in text:` → 从 `text` 取 `action='...'`（相对路径补
   `current_url` 的 scheme+netloc），`_unescape_form_fields(text)` 为 data，POST，`continue`。
3. `if "CreateFido" in text or ("$Config" in text and "sFidoChallenge" in text):` → 调 A2，
   成功 `continue`；`None` → 继续落到后面通用 form 分支（与 reg-factory 一致，不直接失败）。
4. `if "App/Confirm" in current_url:` → 取 `"successUrl":"..."`（还原 `&`/`/`/`&amp;`），
   `session.get(suc, allow_redirects=False)`，`continue`；无 `successUrl` → 返回失败
   （details 标明 `App/Confirm 无 successUrl`）。

不改动：`localhost` code/error 拦截、`Consent/Update`、通用 form 提交、5 步 DoSubmit 预循环、
token 换取、返回值结构（`success/refresh_token/client_id/recovery_email/recovery_email_password`）。

### Part B — 顺带对齐（来自 `a650a8f`，本项目已在代码里、不在流程里）

`a650a8f` 的两点本项目**已经具备**，本次无需改代码，仅在验证清单里复核：
- 注册后授权前必须先有 CF 辅助邮箱（本项目由 `create_or_get_address` 在
  `bind_secondary` 为真时自动分配，等价）。
- `extract_graph_token_http` 透传 `bind_secondary`（本项目对应
  `extract_graph_refresh_token(bind_secondary=...)` 已有，路由 `11:846/:902` 默认 `True`）。

**明确不移植**（与授权无关 / 不适用）：
- `1f960e9`（GitHub `ruoyi` 注册页 mail-tab 登录处理）——本项目无注册链路。
- `auth_nograph_to_all.py` 的浏览器兜底（`bind_secondary_browser.py`、ruyi Firefox）——
  A3/A4 已把纯协议打通，浏览器兜底属注册机专用重依赖，本项目不引入。
- 代理策略调整（`--proxy` 不再默认读 env、CF 收码只认 `CF_MAIL_PROXY`）——本项目 CF 收码
  代理来源已是「`set_proxy`/env/DB `cf_mail_proxy`」优先级链（`12:57`），语义等价，不动。

## 涉及文件

| 文件 | 改动 |
|---|---|
| `outlook_web/segments/11_routes_graph_oauth.py` | 新增 `_unescape_form_fields` / `_parse_cfg_json_balanced` / `_skip_createfido` / `_bind_credentialaction_in_session` / `_resolve_bind_secondary`；替换 3 处 hidden 解析；主循环新增 4 个落点分支；proofs/Add 分支改用 `_resolve_bind_secondary` |
| `tests/test_project_runtime.py` | 新增 `InterruptPageAuthTests`（见下） |
| `tests/test_graph_oauth_extraction.py` | 新增 DoSubmit unescape 回放用例（断言提交的 data 已解码） |
| `docs/superpowers/specs/2026-09-15-port-interrupt-bind-and-authorize.md` | 本文件 |

不改：`12_cloudflare_mail.py`、`13_oauth_bind.py`（仅 A1 复核结论，无代码改动）、
数据库 schema、前端模板/JS、设置项（无新增配置）。

## 测试计划

全部沿用现有 `FakeSession` 录制回放风格（`FakeSession.get` 丢弃 status/headers，`url` 已跟随：
**URL 类落点用 `current_url` 触发，页面类落点用 `text` 触发**）。

新增 `InterruptPageAuthTests`：

1. **credentialaction 绑定成功** — 登录 POST 后 `get` 返回
   `url=.../interrupt/credentialaction` + `text` 含 `ServerData={"apiCanary":"C","acmaInitialResponse":{"continuationToken":"T"}}`
   与 `uaid` hidden；patch `create_or_get_address` / `wait_for_code` / `fetch_admin_mails`；
   断言 `POST api/v1.0/auth/methods/email`、`POST .../email/activate` 被调用，且最终
   `session.get` 重新取 `authorize` 后拿到 `code=`，函数返回 `success=True`、
   `recovery_email == create_or_get_address` 返回的地址。
2. **credentialaction 绑定失败** — `ServerData` 缺 canary/continuationToken → 返回
   `success=False`，details 含「中断页」/「绑定」字样，且不再继续跟链。
3. **CreateFido skip** — `text` 含 `$Config = {...urlPost,sCanary...}` 且 `sFidoChallenge` →
   断言 POST `urlPost`、data 为 `{canary, error_code:"Cancel", i19:"3"}`、`allow_redirects=False`，
   随后继续到 `code=`。
4. **App/Confirm** — `current_url` 含 `App/Confirm`，`text` 含 `"successUrl":"...\\u0026res=success"`
   → 断言 GET 的 URL 已还原为 `&res=success` 且无转义残留；无 `successUrl` 时返回失败。
5. **fido/create 跳板** — `text` 含 `fido/create` + `onload` + `action='/x'` → 断言 POST 到
   `current_url` 的 scheme+netloc + `/x`。
6. **DoSubmit unescape 回归**（`tests/test_graph_oauth_extraction.py`）— DoSubmit 页 hidden value
   含 `&quot;`/`&amp;`，断言提交的 form data 已解码（如 value 为 `a"b&c`），且
   `extract_hidden_inputs` 原行为不变。
7. **已绑号回归** — 无中断页、直接 302 `localhost?code=` 的老链路行为不变（现有用例覆盖，
   确保改动未破坏）。

## 验证

0. 当前状态：`python -m pytest tests/ -q` → **644 passed, 15 subtests passed**
   （含 `InterruptPageAuthTests` 16 例 + 原有 `test_graph_oauth_extraction.py` 回归）。
1. `python -m pytest tests/ -q` 全绿。
2. 真号（未绑辅助邮箱，fresh bind）已验证通过：`susan_jones363291@outlook.com`、
   `elizabeth_davis210@hotmail.com`、`sarah_anderson944@hotmail.com` 三个新号
   走 `extract_graph_refresh_token(bind_secondary=True)` 全链路
   登录 → `interrupt/credentialaction` → `POST email -> 200 interactionRequired`
   → CF 收码 → `activate -> 200 OK` → 重新 GET authorize → Consent → `success=True`。
   其中前两个号当场新建了 CF 辅助邮箱（`ms<prefix>@nuo.dpdns.org`）。
3. 真号（已绑辅助邮箱）回归验证通过：`zbiyvea533249@outlook.com`、
   `robert_white186@hotmail.com` 等直接拿 code，不进中断页。
4. passkey/App-Confirm 分支在真号上未遇到（本批号未触发），仅由合成 fixture 用例覆盖。

## 真号验证中发现并修掉的真坑（P0，未在 reg-factory 记录）

**绑定成功后重新授权会「假性拿到授权码」→ 死循环到「授权流程卡住」。**

绑定完成后再 GET authorize，微软先 302 到：

```
https://login.live.com/oauth20_authorize.srf?client_id=...&scope=...
  &redirect_uri=http%3a%2f%2flocalhost%3a8080&response_type=code&...&msproxy=1
```

旧代码用两个子串判定「已回到 redirect_uri」：`"localhost" in loc` 与 `"code=" in loc`。
而这条 Location **两个都命中** —— `redirect_uri` 是 URL 编码的（字面含 `localhost`），
`response_type=code` 字面含 `code=`。于是把中间跳转当成终点，
`make_light_response(loc)` 造了个空响应体，主循环反复消费同一 hop 直到 `MAX_OAUTH_STEPS`，
最终报「授权流程卡住」。

修法：新增 `is_oauth_code_redirect(loc)`，**解析 query 参数名**（`urlsplit` + `parse_qs`）
判 `code`/`error` 是否存在，host 限 `localhost`/`127.0.0.1`（或相对 Location），
替换主循环、预循环与兜底表单三处子串判定。reg-factory 的 `_is_code_redirect` 用
`"code=" in loc` 同样会让其它带 `response_type=code` 的中间页踩到，本项目按参数判定更稳。

## 风险与对策

- **微软页面结构再变**：与 reg-factory 同样依赖 HTML 抓取，风险既存；新分支按 URL/文本双条件
  触发，命中不了就退化为现有「授权流程卡住」失败，不会误伤老链路。
- **CF 服务不可达**：`create_or_get_address` 失败 → 与现状一致地回退（返回绑定失败并记
  `bind_fail` 语义的错误信息），不阻断其它账号的批量授权。
- **回归面**：Part A2/A3 是新增顶层函数 + 新增分支，A1 仅改 3 处解析口径；已有 26 个
  `test_graph_oauth_extraction.py` 用例 + `test_project_runtime.py` 绑定用例作为回归护栏。