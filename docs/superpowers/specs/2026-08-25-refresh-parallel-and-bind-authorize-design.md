# Token 刷新并行化 + 外部 txt 上传接口 + 用户名密码导入/绑定辅助邮箱/批量授权

Date: 2026-08-25
Status: Approved (verbal) — pending written spec review
Scope: `outlookEmail` repo (`D:\officeProject\yigehui\outlookEmail`)

## 背景

三个独立需求，但对账号都围绕同一批 Outlook 账号的生命周期：

1. **Token 刷新多线程** — 当前 `run_full_refresh` 是单线程串行 `for` 循环 + 每账号 `time.sleep`（默认 5s，可配 `refresh_delay_seconds`），外加进程级 `token_refresh_run_lock`。1W 账号按 1s/个需 ~3 小时。转发任务已有并行模式可镜像。
2. **reg-factory 推送 txt 上传接口** — reg-factory 想把本地 `email_all.txt`（`----` 分隔）直接 API 推送到 outlookEmail 主 `accounts` 表方便管理。现有 `POST /api/accounts` 接受该格式但只支持登录会话（`@login_required`），不支持 API Key。唯一的 API-key 端点 `/api/external/outlook/upload` 只收 `{email,password}` 且写入暂存表 `outlook_upload_accounts`，不写主表。
3. **用户名密码导入 + 绑定辅助邮箱 + 批量授权** — 关键发现：outlookEmail **已有**与 reg-factory 完全相同的用户名+密码 HTTP 登录流程（`11_routes_graph_oauth.py:190-208` POST `login`+`passwd` 到 `login.live.com/ppsecure/post.srf`，与 `reg-factory/extract_graph_tokens.py:432-452` 一致）。缺失的是：(a) 绑定辅助邮箱（proofs/Add 流程在现有循环里被 `skip`），(b) 批量/并行授权，(c) CF 临时邮箱收码依赖。因此 Feature 3 是**扩展**而非移植登录逻辑。

### 关键架构事实（已验证）

- **刷新**：`resolve_account_record` 是解密收口；`run_full_refresh`/`stream_full_refresh_events`/`stream_failed_refresh_events`/`stream_selected_refresh_events` 四个串行循环，共享 `token_refresh_run_lock`；`wait_refresh_delay` 分片 sleep。转发已有并行样板 `run_forwarding_accounts_parallel`（`08_forwarding_scheduler_errors.py:840`）+ `forward_parallel_workers`（clamp 1-10）/`forward_execution_mode`。
- **Token HTTP 调用**：`request_graph_token_response`（`03_mail_helpers.py:416`）POST `TOKEN_URL_GRAPH`，超时 `HTTP_REQUEST_TIMEOUT`(30s)，**无 429 退避**（429 当硬失败）。IMAP 对称 `request_imap_token_response:986`。
- **上传**：`POST /api/accounts`（`04_routes_groups_accounts.py:1368`）→ `parse_account_import`（`02_groups_accounts.py:3804`，按 `provider` 分派 `parse_outlook_account_string`）→ `add_accounts_bulk`（`:1490`）写主表。`@login_required` 装饰器 `03_mail_helpers.py:2869`；`@api_key_required` `:2889`（接受 `X-API-Key`/`?api_key=`/`?apikey=`，`secrets.compare_digest` 比对 `external_api_key` 设置）。
- **授权**：`POST /api/oauth/graph-extract-token`（`11_routes_graph_oauth.py:651`）单账号，建 task_id，SSE 流。worker `run_graph_oauth_task` → `extract_graph_refresh_token`（HTTP 抓取）→ `test_refresh_token` → `save_graph_authorization_result` → `upsert_graph_authorized_account`（`:414`，已存在则 UPDATE 认证字段、新建则 INSERT 主表 `:462`）+ `mark_upload_account_authorized`（暂存 `is_authorized=1`）。暂存表 `outlook_upload_accounts`（`01_bootstrap.py:1766`）列：`id,email,password,is_authorized,status,remark,source,group_id,proxy_url,tag_ids,created_at,updated_at`，无 `client_id/refresh_token/recovery_email`。
- **主表 recovery 列**：`accounts.recovery_email`/`recovery_email_password`（`01_bootstrap.py:1363`，迁移 `:1829`）已由前序工作支持，`build_account_insert_values`/`update_account` 已接受这两参。但 `upsert_graph_authorized_account` **未传**，固定写 `''`。
- **绑定源**：`reg-factory/extract_graph_tokens.py:251-349` `bind_proof_in_session` + `_parse_proof_add_form`/`_parse_proof_verify_form` 正则；`reg-factory/common/cloudflare_mail.py`（379 行，自包含，仅依赖 `requests` + `CF_MAIL_*`）。`create_or_get_address(email)` 返回 `ms-<prefix>@<cf-domain>` 地址 + 可选密码；`wait_for_code()` 轮询收码。

### 用户已确认决策

- 辅助邮箱依赖：**复用 reg-factory 的 Cloudflare 临时邮箱服务**（完整迁移 bind+authorize）。
- 刷新并发：**1-20，默认 5**。
- 上传接口：**新增 API-Key 端点**。
- 辅助邮箱来源：**outlookEmail 自动从 CF 分配**（`create_or_get_address(email)`），无需用户在 txt 里提供。

---

## 改造范围

### Feature 1 — Token 刷新并行化

**新增设置（`settings` 表）：**
- `refresh_parallel_workers` — clamp 1-20，默认 5。
- `refresh_execution_mode` — `serial` | `parallel`，默认 `parallel`（保留 `serial` 作网络不稳回退）。
- `refresh_delay_seconds`（现有，DB 种子值 5 **不变**）— 语义按模式切换：**并行模式下运行时忽略**（并发度本身即节流，对齐转发并行模式置 0 的做法，但不改 DB 值以免影响串行用户）；**串行模式**保持现有"每账号间隔"含义。设置项 UI 注明此行为。

**共享并行辅助** — `refresh_accounts_parallel(accounts, refresh_fn, max_workers, progress_callback, stop_check)`（`05_routes_refresh_mail.py`）：
- `ThreadPoolExecutor(max_workers, thread_name_prefix='refresh-account')`。
- 提交所有账号，逐 future 收集；每完成一个累加 success/fail 并调 `progress_callback({type:'progress', index, total, email, success, error})`——**SSE 事件形态不变**，前端无需改。
- 提交间检查 `is_token_refresh_stop_requested()` 早退。
- `token_refresh_run_lock` 保留——并行/串行一次只能一个刷新进程级运行，不会与定时任务重叠。

**429 退避（新增，`request_graph_token_response`/`request_imap_token_response`，`03_mail_helpers.py`）：** 当前 429 当硬失败。新增：遇 429 读 `Retry-After`（缺省指数 2→4→8s，单账号最多 3 次重试），仍 429 则记 `rate_limited` 失败。限制爆炸半径——一批 429 不会永久杀死账号，退避自然降有效并发。

**重构点：** `run_full_refresh`、`stream_full_refresh_events`、`stream_failed_refresh_events`、`stream_selected_refresh_events` 各加 `if execution_mode == 'parallel': refresh_accounts_parallel(...) else: <现有串行循环>`。非流式 `refresh-selected`/`refresh-failed`（本无 sleep）也一并并行。定时路径 `scheduled_refresh_task → run_full_refresh` 自动继承。

**测试（`tests/test_project_runtime.py`）：**
- (a) 并行模式处理全部账号且 success/fail 计数正确。
- (b) 429 触发退避后成功。
- (c) stop-request 中途取消。
- (d) `token_refresh_run_lock` 拒绝第二次并发运行。
- (e) 串行模式行为不变（回归）。

### Feature 2 — API-Key 上传端点

**新路由** `POST /api/external/accounts/import`（`07_routes_oauth_settings_external.py`，与现有 external 路由并列）：
- 装饰器 `@csrf_exempt` + `@api_key_required`（与 `/api/external/outlook/upload` 同机制）。
- 输入 JSON `{account_string, group_id?, tag_ids?, account_format?, provider?}`，`account_string` 为完整多行 `----` 文本。
- **复用** `parse_account_import` → `add_accounts_bulk`（与 `POST /api/accounts` 同管道）——2/4/6 段自动识别 + `recovery_email` 存储自动继承。
- 写 **主 `accounts` 表**。
- 返回 `{success, added_count, skipped_count, invalid_count, tagged_count}`（与 `POST /api/accounts` 同结构）。
- 未配置 API Key → 403。

**接口说明（提供给 reg-factory）：**
```
POST /api/external/accounts/import
Header: X-API-Key: <external_api_key>
Content-Type: application/json
Body: {
  "account_string": "主邮箱----主密码----client_id----refresh_token----辅助邮箱----辅助邮箱密码\n...",
  "group_id": 1,                  // 可选，默认 1
  "account_format": "client_id_refresh_token",  // 可选，默认 client_id_refresh_token
  "provider": "outlook",          // 可选，默认 outlook
  "tag_ids": "1,2"                // 可选，逗号分隔
}
成功: 200 {"success": true, "added_count": N, "skipped_count": N, "invalid_count": N, "tagged_count": N}
未配置Key: 403 {"success": false, "error": "未配置对外 API Key..."}
Key无效: 401 {"success": false, "error": "API Key 无效"}
```

**测试：** (a) API-key 成功导入多段并写主表；(b) 无 key/错 key 拒绝；(c) 与 `POST /api/accounts` 行为一致（复用解析器）。

### Feature 3 — 用户名密码导入 + 绑定辅助邮箱 + 批量授权

#### 3a. CF 邮件 + 绑定辅助邮箱模块（移植）

两个新文件，直接移植 reg-factory：
- `outlook_web/segments/12_cloudflare_mail.py` ← `reg-factory/common/cloudflare_mail.py`（379 行，自包含，仅依赖 `requests`）。导出 `create_or_get_address(email)`、`wait_for_code(...)`、`set_proxy(...)`。
- `outlook_web/segments/13_oauth_bind.py` ← `reg-factory/extract_graph_tokens.py:251-349` 的 `bind_proof_in_session` + `_parse_proof_add_form`/`_parse_proof_verify_form` 正则辅助。
- **接入** 现有 `extract_graph_refresh_token`（`11_routes_graph_oauth.py`）：当传入 `bind_secondary` 且重定向循环到 `proofs/Add` 时，调 `bind_proof_in_session` 而非 skip。辅助地址由 `cloudflare_mail.create_or_get_address(email)` 自动分配（按用户确认）。
- **持久化** `recovery_email`/`recovery_email_password` 到主 `accounts` 行：经 `upsert_graph_authorized_account`（加 2 参）+ `build_account_insert_values`（已支持）。
- **自动检测已绑定** — 若账号已有辅助邮箱（无 `proofs/Add` 重定向），流程自然跳过绑定（与 reg-factory `auth_nograph_to_all` 同逻辑）。

#### 3b. 批量授权（并行）

**新路由** `POST /api/oauth/graph-extract-batch`：
- 输入 `{account_ids: [int], mode?: 'graph'|'imap', bind_secondary?: bool, max_workers?: int}`，返回 `{task_id, stream_url}`。
- SSE 流逐账号中继进度（复用现有 `GRAPH_OAUTH_TASKS` 模式但 fan-out）。
- worker 用 `ThreadPoolExecutor(max_workers=min(20, requested))` 跑现有单账号 OAuth 流（含可选 bind）。复用 Feature 1 的并行范式。默认 `max_workers=5`，`bind_secondary=true`（CF 已配）。
- 成功 → 提升主表 + 暂存 `is_authorized=1`；失败 → 分类（`abuse`/`noexist`/`pwd_incorrect`/`rate_limited`/`oauth_error`），暂存行 `is_authorized=0` 保留待重试。

#### 3c. UI

暂存页新增「批量授权」按钮 → 调批量路由 + SSE 进度面板（复用现有 refresh SSE UI 组件）。提供「绑定辅助邮箱并授权」开关（默认开）。

#### 3d. 设置

经 `/api/settings` 配置（env 回退）：`cf_mail_base`/`cf_mail_admin`/`cf_mail_domain`/`cf_mail_site_pass`/`cf_mail_proxy`；`oauth_bind_parallel_workers`（1-20，默认5）；`oauth_bind_execution_mode`（serial|parallel，默认 parallel）。

#### 3e. 测试

- (a) 绑定流程针对录制的 HTML 固件（proofs/Add、proofs/Verify），CI 不打真实 MS。
- (b) 批量路由入队 + 流。
- (c) 已绑定账号跳过 proofs。
- (d) `upsert_graph_authorized_account` 写入 recovery 字段。
- (e) 失败分类正确。

---

## 涉及文件

| Feature | 文件 |
|---|---|
| 1 | `outlook_web/segments/05_routes_refresh_mail.py`、`outlook_web/segments/03_mail_helpers.py`、`outlook_web/segments/07_routes_oauth_settings_external.py`（设置项）、`01_bootstrap.py`（设置项种子）、`tests/test_project_runtime.py` |
| 2 | `outlook_web/segments/07_routes_oauth_settings_external.py`、`tests/test_project_runtime.py` |
| 3 | `outlook_web/segments/12_cloudflare_mail.py`（新）、`outlook_web/segments/13_oauth_bind.py`（新）、`outlook_web/segments/11_routes_graph_oauth.py`、`outlook_web/segments/02_groups_accounts.py`（`upsert_graph_authorized_account` 透传）、`outlook_web/segments/07_routes_oauth_settings_external.py`（批量路由+设置）、`01_bootstrap.py`（设置项种子）、`templates/partials/index/...`、`static/js/index/...`、`tests/test_project_runtime.py` |

## 验证

1. `python -m pytest tests/ -q` 全绿。
2. 本地：导入 1W 账号 → 并行刷新（默认5并发）观察完成时间与 429 退避 → reg-factory 用 `X-API-Key` 推 `email_all.txt` 验证主表写入 → 批量授权未绑定账号观察自动绑定 CF 辅助邮箱 → 已绑定账号跳过绑定。
3. SQLite 查 `accounts` 确认 `recovery_email` 明文、`recovery_email_password` 为 `enc:` 密文；`outlook_upload_accounts.is_authorized=1`。

## 关键复用点

- 加解密：`encrypt_data`/`decrypt_data`（`01_bootstrap.py:1100/:1117`）；集中解密 `resolve_account_record`（`02_groups_accounts.py:981`）。
- 转发并行样板：`run_forwarding_accounts_parallel`（`08:840`）、`forward_parallel_workers`/`forward_execution_mode` 设置。
- 上传解析：`parse_account_import`（`02:3804`）→ `add_accounts_bulk`（`02:1490`）。
- API Key 装饰器：`api_key_required`（`03:2889`）。
- 授权提升：`upsert_graph_authorized_account`（`11:414`）+ `build_account_insert_values`（已支持 recovery 两参）。
- 绑定源：`reg-factory/extract_graph_tokens.py:251-349`、`reg-factory/common/cloudflare_mail.py`。

## 风险与对策

- **微软改登录页结构**：现有 `11_routes_graph_oauth.py` 与 reg-factory 同样依赖 HTML 抓取，风险既存；绑定模块作为独立段，失败不影响刷新/上传。
- **429 封 IP**：默认 5 并发 + 单账号退避 + 串行回退；`refresh_execution_mode=serial` 一键降级。
- **CF 服务不可达**：绑定功能软依赖——未配 `CF_MAIL_*` 时批量授权自动回退 `bind_secondary=false`（仅授权不绑定），不报错。
- **API Key 暴露面**：新端点仅写主表，复用现有 key 与限流；不新增鉴权面。
