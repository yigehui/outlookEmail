# 刷新并行 + txt上传接口 + 绑定辅助邮箱/批量授权 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 让 1W Outlook 账号 token 刷新从串行 ~3h 降到并行分钟级；给 reg-factory 一个 API-Key 端点推送 `----` 格式 txt 到主表；扩展现有用户名+密码 OAuth 流程以支持绑定 CF 辅助邮箱 + 批量并行授权。

**Architecture:** 三阶段垂直切片，各阶段独立可测可发布。F1 镜像现有转发并行样板（`run_forwarding_accounts_parallel`）+ 新增 429 退避；F2 新增 `@api_key_required` 端点复用 `parse_account_import`→`add_accounts_bulk`；F3 把 reg-factory 的 `bind_proof_in_session`+`cloudflare_mail` 移植为两个新 segment，接入现有 `extract_graph_refresh_token` 的 proofs/Add 分支（当前 `action="Skip"` 改为可选 `AddProof`+收码+`VerifyProof`），批量授权复用并行范式。

**Tech Stack:** Flask, raw sqlite3, APScheduler, requests, ThreadPoolExecutor。测试 `unittest.TestCase` 跑 `python -m pytest tests/`。

## 关键架构事实（执行前必读）

**Segment 加载机制（`web_outlook_app.py:15-41`）：** `outlook_web/segments/*.py` 不是独立子模块——它们被 `exec(code, globals())` 加载进 `web_outlook_app` 模块的**同一个全局命名空间**。因此：

1. **新增 segment 文件必须加入 `SEGMENT_FILES` 元组**（`web_outlook_app.py:15-27`），否则不会被加载。`12_cloudflare_mail.py` 与 `13_oauth_bind.py` 都要在此追加。
2. **跨 segment 引用是裸名**，不是 import。例如 `11_routes_graph_oauth.py:426` 直接用 `encrypt_data(...)`（它由 `01_bootstrap.py` 定义并存在于共享 globals），没有 `from ... import`。同理 `extract_graph_refresh_token` 调 `bind_proof_in_session`、`create_or_get_address` 都是裸名——只要定义它们的 segment 已在 `SEGMENT_FILES` 中且排在调用方之前（`12`/`13` 在 `11` 之后，但因 `exec` 顺序加载、函数体在调用时才执行，**运行时**引用成立；**模块加载期**不能在 `12`/`13` 顶层代码里反向调用 `11` 的函数）。
3. **测试一律经 `web_outlook_app` 访问**：`web_outlook_app = importlib.import_module('web_outlook_app')`（已在 `tests/test_project_runtime.py:22` 顶部 import），然后 `web_outlook_app.request_graph_token_response(...)`、`web_outlook_app.add_accounts_bulk(...)`、`web_outlook_app.upsert_graph_authorized_account(...)`、`web_outlook_app.run_batch_oauth_task(...)`。**禁止** `from outlook_web.segments.seg_NN_xxx import` 或 `from outlook_web.segments import mail_helpers`——这些路径不存在。**禁止** `init_app(reinit_db=True)`、`tempfile.mkdtemp()+chdir` 等模式——该模块级单例已加载，重载会冲突。
4. **测试 setUp 用既有工作样板**（照抄 `RecoveryEmailTests`，tests/test_project_runtime.py:2102）：
   ```python
   class XxxTests(unittest.TestCase):
       def setUp(self):
           self.app = web_outlook_app.app
           self.app.config['TESTING'] = True
           self.app.config['WTF_CSRF_ENABLED'] = False
           self.client = self.app.test_client()
           with self.app.app_context():
               web_outlook_app.init_db()
               web_outlook_app.set_setting(
                   web_outlook_app.LOGIN_SESSION_VERSION_SETTING_KEY,
                   web_outlook_app.DEFAULT_LOGIN_SESSION_VERSION,
               )
               db = web_outlook_app.get_db()
               db.execute('DELETE FROM accounts')
               db.execute('DELETE FROM outlook_upload_accounts')  # 暂存表
               db.execute("DELETE FROM groups WHERE name NOT IN ('默认分组', '临时邮箱')")
               web_outlook_app.set_setting('login_password', web_outlook_app.hash_password('export-pass'))
               db.commit()
           with self.client.session_transaction() as sess:
               sess['logged_in'] = True
               sess['login_session_version'] = web_outlook_app.DEFAULT_LOGIN_SESSION_VERSION

       def _verify(self):  # 需要调导出时用
           resp = self.client.post('/api/export/verify', json={'password': 'export-pass'})
           self.assertEqual(resp.status_code, 200)
           return resp.get_json()['verify_token']
   ```
   - 查 DB：`with self.app.app_context(): row = web_outlook_app.get_db().execute("SELECT ...").fetchone()`，**不要** `sqlite3.connect`/`_open_db()`。
   - 需 API Key 的测试：`with self.app.app_context(): web_outlook_app.set_setting('external_api_key', 'sk-test-123')`。
5. mock 目标用 `web_outlook_app.<name>`：`@patch("web_outlook_app.post_with_proxy_fallback")`、`@patch("web_outlook_app.refresh_accounts_parallel")`、`@patch("web_outlook_app.run_graph_oauth_task")`、`@patch("web_outlook_app.create_or_get_address")`、`@patch("web_outlook_app.bind_proof_in_session")`、`@patch.object(web_outlook_app, 'encrypt_data', ...)`、`@patch.object(web_outlook_app, 'test_refresh_token', ...)`。被 patch 的名字必须存在于 `web_outlook_app` globals（被某 segment 顶层 `def` 定义），patch 后所有 segment 内对该裸名的引用都受影响。
6. **设置读写**：`web_outlook_app.set_setting(key, value)` / `web_outlook_app.get_setting_value(key)`（已存在；实现时 grep `def set_setting` 确认）。

**本计划中所有 `from outlook_web.segments... import`、`mh.xxx`/`refresh_mail.xxx`/`graph_oauth.xxx`/`groups_accounts.xxx`、`init_app(reinit_db=True)`、`tempfile+chdir`、`_open_db()` 写法均应按上述规则改写为 `web_outlook_app.<name>` + 既有 setUp 样板。** 下文任务代码块保留原写法作为逻辑参考，实现时按本节规则转译。

设计稿：`docs/superpowers/specs/2026-08-25-refresh-parallel-and-bind-authorize-design.md`

---

## File Structure

| 文件 | 责任 | 动作 |
|---|---|---|
| `outlook_web/segments/05_routes_refresh_mail.py` | 刷新入口 + 新增 `refresh_accounts_parallel` 辅助 | 修改 |
| `outlook_web/segments/03_mail_helpers.py` | token HTTP 调用 + 429 退避 | 修改 |
| `outlook_web/segments/07_routes_oauth_settings_external.py` | 设置读写 + external 路由 + 批量授权路由 | 修改 |
| `outlook_web/segments/01_bootstrap.py` | 设置项种子 | 修改 |
| `outlook_web/segments/12_cloudflare_mail.py` | CF 临时邮箱收发（移植自 reg-factory） | 新建（需加入 SEGMENT_FILES） |
| `outlook_web/segments/13_oauth_bind.py` | `bind_proof_in_session` + 正则辅助（移植自 reg-factory） | 新建（需加入 SEGMENT_FILES） |
| `outlook_web/segments/11_routes_graph_oauth.py` | 接入 bind + recovery 透传 + 批量授权 worker | 修改 |
| `outlook_web/segments/02_groups_accounts.py` | `upsert_graph_authorized_account` 透传 recovery | 修改 |
| `templates/partials/index/dialogs-primary.html` | 批量授权 UI | 修改 |
| `static/js/index/07-settings.js` | 批量授权前端 | 修改 |
| `tests/test_project_runtime.py` | 全部新测试 | 修改 |

---

# Phase F1 — Token 刷新并行化

## Task F1.0: 创建工作分支

**Files:** (none)

- [ ] **Step 1: 建分支**

```bash
git checkout -b feat/refresh-parallel-and-bind-authorize
```

- [ ] **Step 2: 确认基线测试绿**

Run: `python -m pytest tests/ -q 2>&1 | tail -5`
Expected: `557 passed`（或当前数，记录下来作回归基线）

---

## Task F1.1: 429 退避重试（token HTTP 调用层）

**Files:**
- Modify: `outlook_web/segments/03_mail_helpers.py` (`request_graph_token_response` 约 :416，`request_imap_token_response` 约 :986)
- Test: `tests/test_project_runtime.py`

- [ ] **Step 1: 写失败测试**

在 `tests/test_project_runtime.py` 末尾、最后一个测试类之前，新增 `GraphTokenRetryTests`：

```python
class GraphTokenRetryTests(unittest.TestCase):
    def _mock_429_then_200(self):
        """返回一个 response 序列：429(Retry-After=1) -> 200。"""
        import types

        def make(status, payload=None, retry_after=None):
            r = types.SimpleNamespace()
            r.status_code = status
            r.headers = {"Retry-After": retry_after} if retry_after else {}
            r.json = lambda: payload or {}
            r.text = "" if status != 200 else "{}"
            return r

        seq = [make(429, {"error": "temporarily_unavailable"}, retry_after="0"), make(200, {"access_token": "at", "refresh_token": "rt"})]
        return seq

    @patch("outlook_web.segments.03_mail_helpers.post_with_proxy_fallback")
    def test_graph_token_retries_on_429_then_succeeds(self, mock_post):
        from outlook_web.segments import mail_helpers as mh
        seq = self._mock_429_then_200()
        mock_post.side_effect = seq
        resp = mh.request_graph_token_response("cid", "rt")
        self.assertEqual(resp.status_code, 200)
        self.assertEqual(mock_post.call_count, 2)

    @patch("outlook_web.segments.03_mail_helpers.post_with_proxy_fallback")
    def test_graph_token_gives_up_after_max_429(self, mock_post):
        from outlook_web.segments import mail_helpers as mh
        import types
        def make429():
            r = types.SimpleNamespace()
            r.status_code = 429
            r.headers = {"Retry-After": "0"}
            r.json = lambda: {"error": "temporarily_unavailable"}
            r.text = ""
            return r
        # 超过最大重试次数(3+1=4 次 429)，最后返回 429
        mock_post.side_effect = [make429() for _ in range(4)]
        resp = mh.request_graph_token_response("cid", "rt")
        self.assertEqual(resp.status_code, 429)
```

- [ ] **Step 2: 验证失败**

Run: `python -m pytest tests/test_project_runtime.py::GraphTokenRetryTests -v`
Expected: FAIL — 当前无 429 重试逻辑，第一次 429 即返回，`call_count == 1`。

- [ ] **Step 3: 实现 429 退避**

在 `outlook_web/segments/03_mail_helpers.py` 顶部常量区加：

```python
GRAPH_TOKEN_MAX_429_RETRIES = 3
```

改造 `request_graph_token_response`（约 :416），在 `for index, (_label, scope) in enumerate(candidates):` 循环内的 `response = post_with_proxy_fallback(...)` 之后、`last_response = response` 之后插入 429 重试块。完整替换函数体为：

```python
def request_graph_token_response(client_id: str, refresh_token: str, proxy_url: str = None,
                                 fallback_proxy_urls: Optional[List[str]] = None,
                                 include_original_scope_fallback: bool = False):
    """请求 Graph token，优先使用授权时的显式委托 scope，避免 .default 依赖应用预配置权限。
    遇 429 限流时读 Retry-After 退避重试（单 scope 最多 GRAPH_TOKEN_MAX_429_RETRIES 次）。"""
    last_response = None
    candidates = get_graph_token_scope_candidates(include_original_scope_fallback)
    for index, (_label, scope) in enumerate(candidates):
        data = {
            "client_id": client_id,
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
        }
        if scope:
            data["scope"] = scope

        response = _post_token_with_429_backoff(
            TOKEN_URL_GRAPH,
            data,
            proxy_url=proxy_url,
            fallback_proxy_urls=fallback_proxy_urls,
        )
        last_response = response
        if response.status_code == 200:
            return response
        if index == len(candidates) - 1 or not is_graph_token_scope_retryable_response(response):
            return response

    return last_response
```

新增私有退避函数（紧接 `request_graph_token_response` 之前）：

```python
def _post_token_with_429_backoff(url, data, *, proxy_url=None, fallback_proxy_urls=None):
    """POST token 端点，遇 429 读 Retry-After 退避重试（指数兜底 2->4->8s），最多 GRAPH_TOKEN_MAX_429_RETRIES 次。"""
    import time as _time
    response = post_with_proxy_fallback(
        url,
        data=data,
        timeout=HTTP_REQUEST_TIMEOUT,
        proxy_url=proxy_url,
        fallback_proxy_urls=fallback_proxy_urls,
    )
    for attempt in range(GRAPH_TOKEN_MAX_429_RETRIES):
        if response.status_code != 429:
            return response
        retry_after_raw = (response.headers or {}).get("Retry-After", "")
        try:
            delay = float(retry_after_raw) if retry_after_raw else (2 ** attempt)
        except (TypeError, ValueError):
            delay = 2 ** attempt
        _time.sleep(min(delay, 8.0))
        response = post_with_proxy_fallback(
            url,
            data=data,
            timeout=HTTP_REQUEST_TIMEOUT,
            proxy_url=proxy_url,
            fallback_proxy_urls=fallback_proxy_urls,
        )
    return response
```

对称改造 `request_imap_token_response`（约 :986）——把其中对 `post_with_proxy_fallback(TOKEN_URL_IMAP, ...)` 的直接调用替换为 `_post_token_with_429_backoff(TOKEN_URL_IMAP, ...)`。

- [ ] **Step 4: 验证通过**

Run: `python -m pytest tests/test_project_runtime.py::GraphTokenRetryTests -v`
Expected: PASS

- [ ] **Step 5: 提交**

```bash
git add outlook_web/segments/03_mail_helpers.py tests/test_project_runtime.py
git commit -m "feat(refresh): token 刷新 429 退避重试

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

## Task F1.2: 并行刷新设置项 + 种子

**Files:**
- Modify: `outlook_web/segments/07_routes_oauth_settings_external.py`（设置读写，约 :754 附近 `refresh_delay_seconds` 处理）
- Modify: `outlook_web/segments/01_bootstrap.py`（设置项种子，约 :2111）
- Test: `tests/test_project_runtime.py`

- [ ] **Step 1: 写失败测试**

在 `tests/test_project_runtime.py` 新增 `RefreshParallelSettingsTests`：

```python
class RefreshParallelSettingsTests(unittest.TestCase):
    def setUp(self):
        import tempfile, os
        from outlook_web import web_outlook_app
        self._tmp_dir = tempfile.mkdtemp()
        self._orig_cwd = os.getcwd()
        os.chdir(self._tmp_dir)
        os.makedirs("data", exist_ok=True)
        web_outlook_app.init_app(reinit_db=True)
        from outlook_web.segments import bootstrap as bs
        bs.set_login_password("p")
        self.client = web_outlook_app.app.test_client()

    def tearDown(self):
        import os, shutil
        os.chdir(self._orig_cwd)
        shutil.rmtree(self._tmp_dir, ignore_errors=True)

    def _login(self):
        self.client.post("/api/login", data={"password": "p"})

    def test_default_parallel_workers_is_5(self):
        from outlook_web.segments import mail_helpers as mh
        self._login()
        with mh.get_db_conn() if hasattr(mh, "get_db_conn") else _open_db() as conn:
            row = conn.execute("SELECT value FROM settings WHERE key='refresh_parallel_workers'").fetchone()
            self.assertEqual(int(row["value"]), 5)

    def test_parallel_workers_clamped_to_1_20(self):
        from outlook_web.segments import mail_helpers as mh
        self._login()
        self.client.post("/api/settings", data={"refresh_parallel_workers": "99"})
        self.client.post("/api/settings", data={"refresh_parallel_workers": "0"})
        with _open_db() as conn:
            vals = {r["key"]: r["value"] for r in conn.execute("SELECT key,value FROM settings WHERE key='refresh_parallel_workers'")}
        # 写入原值（钳制在读时做），读取应得 clamp 值
        from outlook_web.segments import mail_helpers as mh2
        self.assertLessEqual(int(mh2.normalize_refresh_parallel_workers("99")), 20)
        self.assertEqual(mh2.normalize_refresh_parallel_workers("0"), 1)
```

注：`_open_db` 若测试文件已有则复用；否则在文件顶部 helper 区加：
```python
import sqlite3 as _sqlite3
def _open_db():
    conn = _sqlite3.connect("data/outlook_accounts.db")
    conn.row_factory = _sqlite3.Row
    return conn
```

- [ ] **Step 2: 验证失败**

Run: `python -m pytest tests/test_project_runtime.py::RefreshParallelSettingsTests -v`
Expected: FAIL — `refresh_parallel_workers` 设置项不存在，`normalize_refresh_parallel_workers` 未定义。

- [ ] **Step 3: 实现设置项读写 + 归一化**

在 `outlook_web/segments/05_routes_refresh_mail.py` 的 `get_refresh_delay_seconds`（约 :1008）附近新增：

```python
def normalize_refresh_parallel_workers(value) -> int:
    """钳制刷新并发度到 1-20，默认 5。"""
    try:
        v = int(value)
    except (TypeError, ValueError):
        return 5
    return max(1, min(20, v))


def get_refresh_parallel_workers(db_conn) -> int:
    row = db_conn.execute(
        "SELECT value FROM settings WHERE key = 'refresh_parallel_workers'"
    ).fetchone()
    if not row or row['value'] is None:
        return 5
    return normalize_refresh_parallel_workers(row['value'])


REFRESH_EXECUTION_MODES = {'serial', 'parallel'}


def get_refresh_execution_mode(db_conn) -> str:
    row = db_conn.execute(
        "SELECT value FROM settings WHERE key = 'refresh_execution_mode'"
    ).fetchone()
    mode = str(row['value']) if row and row['value'] else 'parallel'
    return mode if mode in REFRESH_EXECUTION_MODES else 'parallel'
```

在 `outlook_web/segments/01_bootstrap.py` 设置种子区（`refresh_delay_seconds` 行 `:2111` 附近）追加：

```python
        ('refresh_parallel_workers', '5'),
        ('refresh_execution_mode', 'parallel'),
```

在 `outlook_web/segments/07_routes_oauth_settings_external.py` 设置写入路由（处理 `refresh_delay_seconds` 的同一函数，约 :754）追加对两个新 key 的接受：

```python
    if 'refresh_parallel_workers' in (request.form or {}) or request.json and 'refresh_parallel_workers' in (request.json or {}):
        raw = request.form.get('refresh_parallel_workers') if request.form else request.json.get('refresh_parallel_workers')
        clamped = max(1, min(20, int(raw))) if str(raw).strip().lstrip('-').isdigit() else 5
        set_setting('refresh_parallel_workers', str(clamped))
    if 'refresh_execution_mode' in (request.form or {}) or request.json and 'refresh_execution_mode' in (request.json or {}):
        mode = (request.form.get('refresh_execution_mode') if request.form else request.json.get('refresh_execution_mode')) or 'parallel'
        if mode not in ('serial', 'parallel'):
            mode = 'parallel'
        set_setting('refresh_execution_mode', mode)
```

注：精确变量名以文件里 `set_setting` 实际名为准——若该文件用 `update_setting` 则替换。实现时读 :754 上下文确认。

- [ ] **Step 4: 验证通过**

Run: `python -m pytest tests/test_project_runtime.py::RefreshParallelSettingsTests -v`
Expected: PASS

- [ ] **Step 5: 提交**

```bash
git add outlook_web/segments/05_routes_refresh_mail.py outlook_web/segments/01_bootstrap.py outlook_web/segments/07_routes_oauth_settings_external.py tests/test_project_runtime.py
git commit -m "feat(refresh): 新增 refresh_parallel_workers/refresh_execution_mode 设置项

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

## Task F1.3: 共享并行刷新辅助

**Files:**
- Modify: `outlook_web/segments/05_routes_refresh_mail.py`
- Test: `tests/test_project_runtime.py`

- [ ] **Step 1: 写失败测试**

新增 `RefreshParallelDispatchTests`：

```python
class RefreshParallelDispatchTests(unittest.TestCase):
    def setUp(self):
        import tempfile, os
        from outlook_web import web_outlook_app
        self._tmp_dir = tempfile.mkdtemp()
        self._orig_cwd = os.getcwd()
        os.chdir(self._tmp_dir)
        os.makedirs("data", exist_ok=True)
        web_outlook_app.init_app(reinit_db=True)
        from outlook_web.segments import bootstrap as bs
        bs.set_login_password("p")

    def tearDown(self):
        import os, shutil
        os.chdir(self._orig_cwd)
        shutil.rmtree(self._tmp_dir, ignore_errors=True)

    def test_parallel_runs_all_accounts_and_counts_correctly(self):
        from outlook_web.segments import refresh_mail as rm
        calls = []
        def fake_refresh(account, log_type, db_conn=None):
            calls.append(account['email'])
            return {'success': True, 'email': account['email']}
        accounts = [{'id': i, 'email': f'u{i}@x.com'} for i in range(7)]
        results = rm.refresh_accounts_parallel(
            accounts,
            refresh_fn=fake_refresh,
            max_workers=5,
            progress_callback=None,
            stop_check=lambda: False,
            db_conn=None,
            log_refresh_type='manual',
        )
        self.assertEqual(len(results), 7)
        self.assertEqual(set(calls), {a['email'] for a in accounts})
        self.assertTrue(all(r['success'] for r in results))

    def test_parallel_stop_request_cancels_remaining(self):
        from outlook_web.segments import refresh_mail as rm
        submitted = []
        def fake_refresh(account, log_type, db_conn=None):
            submitted.append(account['email'])
            return {'success': True, 'email': account['email']}
        # stop_check 在第 3 个提交后返回 True
        counter = {'n': 0}
        def stop_check():
            counter['n'] += 1
            return counter['n'] > 3
        accounts = [{'id': i, 'email': f'u{i}@x.com'} for i in range(20)]
        results = rm.refresh_accounts_parallel(
            accounts, refresh_fn=fake_refresh, max_workers=2,
            progress_callback=None, stop_check=stop_check,
            db_conn=None, log_refresh_type='manual',
        )
        # 至少应少于全部 20 个
        self.assertLess(len(submitted), 20)
```

- [ ] **Step 2: 验证失败**

Run: `python -m pytest tests/test_project_runtime.py::RefreshParallelDispatchTests -v`
Expected: FAIL — `refresh_accounts_parallel` 未定义。

- [ ] **Step 3: 实现并行辅助**

在 `outlook_web/segments/05_routes_refresh_mail.py`，`wait_refresh_delay`（:449）之前或 `get_refresh_delay_seconds` 附近新增：

```python
def refresh_accounts_parallel(accounts, *, refresh_fn, max_workers, progress_callback,
                              stop_check, db_conn, log_refresh_type):
    """并行执行单账号刷新，逐 future 回调进度，保持 SSE 事件形态与串行一致。
    stop_check 返回 True 时停止提交剩余账号（已提交的 future 仍等其完成）。"""
    from concurrent.futures import ThreadPoolExecutor, as_completed

    if not accounts:
        return []
    effective_workers = min(max(1, max_workers), len(accounts))
    results = []
    total = len(accounts)
    completed_index = 0

    with ThreadPoolExecutor(max_workers=effective_workers, thread_name_prefix='refresh-account') as executor:
        future_map = {}
        for account in accounts:
            if stop_check():
                break
            future = executor.submit(refresh_fn, account, log_refresh_type, db_conn=db_conn)
            future_map[future] = account

        for future in as_completed(future_map):
            account = future_map[future]
            completed_index += 1
            try:
                result = future.result()
            except Exception as exc:
                result = {'success': False, 'email': account.get('email', ''), 'error': str(exc)}
            results.append(result)
            if progress_callback:
                progress_callback({
                    'type': 'progress',
                    'index': completed_index,
                    'total': total,
                    'email': account.get('email', ''),
                    'success': bool(result.get('success')),
                    'error': result.get('error', ''),
                })
            if stop_check():
                # 取消尚未开始的；正在跑的让其结束
                for f in future_map:
                    f.cancel()
                break
    return results
```

注：`refresh_fn` 即 `refresh_outlook_account_token`。提交后 worker 自己 `conn.commit()`（与串行 `run_full_refresh` 一致，每账号后 commit）。

- [ ] **Step 4: 验证通过**

Run: `python -m pytest tests/test_project_runtime.py::RefreshParallelDispatchTests -v`
Expected: PASS

- [ ] **Step 5: 提交**

```bash
git add outlook_web/segments/05_routes_refresh_mail.py tests/test_project_runtime.py
git commit -m "feat(refresh): 共享并行刷新辅助 refresh_accounts_parallel

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

## Task F1.4: 四个刷新循环接入并行分支

**Files:**
- Modify: `outlook_web/segments/05_routes_refresh_mail.py` (`run_full_refresh` :1018, `stream_full_refresh_events` :1190, `stream_failed_refresh_events` :1323, `stream_selected_refresh_events` :1497)

- [ ] **Step 1: 写失败测试**

新增 `RefreshModeDispatchTests`，断言 `run_full_refresh` 在 `refresh_execution_mode=parallel` 时调用 `refresh_accounts_parallel`（mock 掉并行辅助与单账号刷新，验证不会走串行 sleep）：

```python
class RefreshModeDispatchTests(unittest.TestCase):
    def setUp(self):
        import tempfile, os
        from outlook_web import web_outlook_app
        self._tmp_dir = tempfile.mkdtemp()
        self._orig_cwd = os.getcwd()
        os.chdir(self._tmp_dir)
        os.makedirs("data", exist_ok=True)
        web_outlook_app.init_app(reinit_db=True)
        from outlook_web.segments import bootstrap as bs
        bs.set_login_password("p")
        self._patch_parallel = patch("outlook_web.segments.refresh_mail.refresh_accounts_parallel")

    def tearDown(self):
        import os, shutil
        os.chdir(self._orig_cwd)
        shutil.rmtree(self._tmp_dir, ignore_errors=True)

    @patch("outlook_web.segments.refresh_mail.refresh_accounts_parallel")
    def test_run_full_refresh_uses_parallel_when_configured(self, mock_parallel):
        from outlook_web.segments import refresh_mail as rm
        # 准备 2 个账号
        rm.add_account("a@x.com", "pw", "cid", "rt", group_id=1)
        rm.add_account("b@x.com", "pw", "cid", "rt", group_id=1)
        # 设置并行模式
        with _open_db() as conn:
            conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('refresh_execution_mode','parallel')")
            conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('refresh_parallel_workers','3')")
            conn.commit()
        mock_parallel.return_value = [
            {'success': True, 'email': 'a@x.com'},
            {'success': True, 'email': 'b@x.com'},
        ]
        result = rm.run_full_refresh('manual_all', 'manual', progress_callback=None)
        self.assertTrue(mock_parallel.called)
        self.assertEqual(result['success_count'], 2)

    @patch("outlook_web.segments.refresh_mail.refresh_accounts_parallel")
    def test_run_full_refresh_serial_mode_not_called(self, mock_parallel):
        from outlook_web.segments import refresh_mail as rm
        rm.add_account("c@x.com", "pw", "cid", "rt", group_id=1)
        with _open_db() as conn:
            conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('refresh_execution_mode','serial')")
            conn.commit()
        # mock 掉单账号刷新避免真实 HTTP
        with patch("outlook_web.segments.refresh_mail.refresh_outlook_account_token", return_value={'success': True}):
            rm.run_full_refresh('manual_all', 'manual', progress_callback=None)
        self.assertFalse(mock_parallel.called)
```

- [ ] **Step 2: 验证失败**

Run: `python -m pytest tests/test_project_runtime.py::RefreshModeDispatchTests -v`
Expected: FAIL — `run_full_refresh` 当前无并行分支，`mock_parallel` 不会被调用。

- [ ] **Step 3: 改造 run_full_refresh**

在 `run_full_refresh`（:1018）加载 `accounts` 后、原 `for index, account in enumerate(accounts, 1):` 串行循环之前，插入并行分支：

```python
    execution_mode = get_refresh_execution_mode(conn)
    if execution_mode == 'parallel':
        parallel_workers = get_refresh_parallel_workers(conn)
        def _parallel_progress(payload):
            if progress_callback:
                progress_callback(payload)
        parallel_results = refresh_accounts_parallel(
            accounts,
            refresh_fn=refresh_outlook_account_token,
            max_workers=parallel_workers,
            progress_callback=_parallel_progress,
            stop_check=is_token_refresh_stop_requested,
            db_conn=conn,
            log_refresh_type=log_refresh_type,
        )
        for r in parallel_results:
            if r.get('success'):
                success_count += 1
            else:
                failed_count += 1
                failed_list.append({'email': r.get('email', ''), 'error': r.get('error', '')})
        conn.commit()
        return finalize_full_refresh(
            conn, snapshot_trigger_type, total, success_count, failed_count, failed_list, 0
        )
```

注：`finalize_full_refresh` 为现有收尾函数——实现时读 `run_full_refresh` 末尾确认其确切名与参数；若该函数无独立收尾而是内联，则把并行分支末尾的内联收尾代码复制自串行分支末尾（`return {success:..., success_count, failed_count, failed_list, ...}` 结构）。

对 `stream_full_refresh_events`(:1190)、`stream_failed_refresh_events`(:1323)、`stream_selected_refresh_events`(:1497) 做对称改造：在加载 accounts 后、串行循环前加 `if get_refresh_execution_mode(conn) == 'parallel':` 分支，分支内调 `refresh_accounts_parallel`，`progress_callback` 用各自的 SSE `event_queue.put({'type':'progress', ...})`。SSE 形态保持，前端无需改。stream 版仍要把 start/delay/done/complete 事件按现有顺序发出。

- [ ] **Step 4: 验证通过**

Run: `python -m pytest tests/test_project_runtime.py::RefreshModeDispatchTests -v`
Expected: PASS

- [ ] **Step 5: 回归全量**

Run: `python -m pytest tests/ -q 2>&1 | tail -10`
Expected: 全绿（基线数 + 新增）

- [ ] **Step 6: 提交**

```bash
git add outlook_web/segments/05_routes_refresh_mail.py tests/test_project_runtime.py
git commit -m "feat(refresh): 四个刷新循环接入并行分支，默认 parallel

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

# Phase F2 — API-Key txt 上传端点

## Task F2.1: 新增 external 导入端点

**Files:**
- Modify: `outlook_web/segments/07_routes_oauth_settings_external.py`（现有 `/api/external/outlook/upload` :1387 附近）
- Test: `tests/test_project_runtime.py`

- [ ] **Step 1: 写失败测试**

新增 `ExternalImportApiTests`：

```python
class ExternalImportApiTests(unittest.TestCase):
    def setUp(self):
        import tempfile, os
        from outlook_web import web_outlook_app
        self._tmp_dir = tempfile.mkdtemp()
        self._orig_cwd = os.getcwd()
        os.chdir(self._tmp_dir)
        os.makedirs("data", exist_ok=True)
        web_outlook_app.init_app(reinit_db=True)
        from outlook_web.segments import bootstrap as bs
        bs.set_login_password("p")
        self.client = web_outlook_app.app.test_client()
        # 设置 API Key
        with _open_db() as conn:
            conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('external_api_key','sk-test-123')")
            conn.commit()

    def tearDown(self):
        import os, shutil
        os.chdir(self._orig_cwd)
        shutil.rmtree(self._tmp_dir, ignore_errors=True)

    def test_import_with_valid_api_key_writes_main_table(self):
        line = "imp1@x.com----pw1----clientid1----reftoken1----aux1@cf.com----auxpw1"
        resp = self.client.post(
            "/api/external/accounts/import",
            headers={"X-API-Key": "sk-test-123", "Content-Type": "application/json"},
            json={"account_string": line, "group_id": 1},
        )
        self.assertEqual(resp.status_code, 200)
        data = resp.get_json()
        self.assertTrue(data["success"])
        self.assertEqual(data["added_count"], 1)
        with _open_db() as conn:
            row = conn.execute("SELECT email, recovery_email FROM accounts WHERE email='imp1@x.com'").fetchone()
        self.assertIsNotNone(row)
        self.assertEqual(row["recovery_email"], "aux1@cf.com")

    def test_import_without_api_key_rejected(self):
        resp = self.client.post(
            "/api/external/accounts/import",
            json={"account_string": "x@x.com----pw----cid----rt"},
        )
        self.assertEqual(resp.status_code, 401)

    def test_import_invalid_key_rejected(self):
        resp = self.client.post(
            "/api/external/accounts/import",
            headers={"X-API-Key": "wrong"},
            json={"account_string": "x@x.com----pw----cid----rt"},
        )
        self.assertEqual(resp.status_code, 401)

    def test_import_4_segment_backward_compatible(self):
        resp = self.client.post(
            "/api/external/accounts/import",
            headers={"X-API-Key": "sk-test-123"},
            json={"account_string": "imp2@x.com----pw2----cid2----rt2"},
        )
        self.assertEqual(resp.status_code, 200)
        with _open_db() as conn:
            row = conn.execute("SELECT recovery_email FROM accounts WHERE email='imp2@x.com'").fetchone()
        self.assertIsNotNone(row)
        self.assertEqual(row["recovery_email"], "")
```

- [ ] **Step 2: 验证失败**

Run: `python -m pytest tests/test_project_runtime.py::ExternalImportApiTests -v`
Expected: FAIL — 404（路由不存在）。

- [ ] **Step 3: 实现端点**

在 `outlook_web/segments/07_routes_oauth_settings_external.py`，紧邻 `/api/external/outlook/upload`（:1387）之后新增：

```python
@app.route('/api/external/accounts/import', methods=['POST'])
@csrf_exempt
@api_key_required
def api_external_import_accounts():
    """对外 API-Key 端点：接受 ---- 分隔的账号文本，写入主 accounts 表。
    复用与 POST /api/accounts 相同的 parse_account_import -> add_accounts_bulk 管道。"""
    data = request.get_json(silent=True) or {}
    account_string = (data.get('account_string') or '').strip()
    if not account_string:
        return jsonify({'success': False, 'error': 'account_string 不能为空'}), 400

    group_id = data.get('group_id', 1)
    account_format = data.get('account_format', 'client_id_refresh_token')
    provider = data.get('provider', 'outlook')
    tag_ids_raw = data.get('tag_ids', '')
    tag_ids = [int(x) for x in str(tag_ids_raw).split(',') if str(x).strip().lstrip('-').isdigit()] if tag_ids_raw else []
    proxy_url = data.get('proxy_url', '')
    fallback_proxy_url_1 = data.get('fallback_proxy_url_1', '')
    fallback_proxy_url_2 = data.get('fallback_proxy_url_2', '')
    imap_host = data.get('imap_host', '')
    imap_port = data.get('imap_port')

    lines = [ln for ln in account_string.split('\n') if ln.strip()]
    parsed_rows = []
    skipped = 0
    invalid = 0
    for line in lines:
        line = line.strip()
        try:
            parsed = parse_account_import(line, provider, account_format, imap_host, imap_port)
            if parsed:
                parsed_rows.append(parsed)
            else:
                invalid += 1
        except Exception:
            invalid += 1

    result = add_accounts_bulk(
        parsed_rows,
        group_id=group_id,
        tag_ids=tag_ids,
        proxy_url=proxy_url,
        fallback_proxy_url_1=fallback_proxy_url_1,
        fallback_proxy_url_2=fallback_proxy_url_2,
    )
    return jsonify({
        'success': True,
        'added_count': result.get('added_count', 0),
        'skipped_count': result.get('skipped_count', 0) + skipped,
        'invalid_count': result.get('invalid_count', 0) + invalid,
        'tagged_count': result.get('tagged_count', 0),
    })
```

注：`parse_account_import`、`add_accounts_bulk`、`parse_imap_account_string` 等需在该文件顶部已 import（确认 `from .segments import` 或 `from outlook_web.segments import` 风格——照搬该文件现有 import 写法）。若 `add_accounts_bulk` 签名不接受这些 kwargs，去掉不支持的 kwargs（实现时读 `02_groups_accounts.py:1490` 的 `add_accounts_bulk` 签名确认）。

- [ ] **Step 4: 验证通过**

Run: `python -m pytest tests/test_project_runtime.py::ExternalImportApiTests -v`
Expected: PASS

- [ ] **Step 5: 提交**

```bash
git add outlook_web/segments/07_routes_oauth_settings_external.py tests/test_project_runtime.py
git commit -m "feat(import): 新增 /api/external/accounts/import API-Key 端点

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

# Phase F3 — 绑定辅助邮箱 + 批量授权

## Task F3.1: 移植 cloudflare_mail 模块

**Files:**
- Create: `outlook_web/segments/12_cloudflare_mail.py`
- Source: `D:\officeProject\yigehui\reg-factory\common\cloudflare_mail.py`

- [ ] **Step 1: 复制源文件**

```bash
cp "D:/officeProject/yigehui/reg-factory/common/cloudflare_mail.py" "D:/officeProject/yigehui/outlookEmail/outlook_web/segments/12_cloudflare_mail.py"
```

- [ ] **Step 2: 改 import 路径适配 Flask 应用上下文**

打开 `outlook_web/segments/12_cloudflare_mail.py`，顶部加模块说明并确认 `import requests`、`import os`、`import re` 等 stdlib 不变。把 reg-factory 里 `from common import ...` 风格的内联引用（若 cloudflare_mail.py 内部无跨文件 import 则无需改——确认它自包含）。

`_resolve_proxy`（:53）改为先读 env 再回退 DB settings（裸名 `get_setting_value`，由 `01_bootstrap.py` 定义于共享 globals）：

```python
def _resolve_proxy():
    env = os.environ.get("CF_MAIL_PROXY")
    if env:
        return env
    try:
        return get_setting_value('cf_mail_proxy') or None
    except Exception:
        return None
```

对称：`create_or_get_address` 读 `CF_MAIL_BASE/ADMIN/DOMAIN/SITE_PASS` 处加 `get_setting_value(...)` 回退（env 优先）。

- [ ] **Step 3: 注册到 SEGMENT_FILES**

在 `web_outlook_app.py:15-27` 的 `SEGMENT_FILES` 元组追加两项：

```python
    "10_routes_email_shares.py",
    "11_routes_graph_oauth.py",
    "12_cloudflare_mail.py",
    "13_oauth_bind.py",
)
```

（`13_oauth_bind.py` 文件在 F3.2 创建，但元组一次性加好。）

- [ ] **Step 4: 写测试**

新增 `CloudflareMailTests`（纯单元，mock requests，不打真实网络）。patch 目标是 `web_outlook_app.requests`（因 `12_cloudflare_mail.py` 顶层 `import requests` 后 `requests` 进入 `web_outlook_app` globals）：

```python
class CloudflareMailTests(unittest.TestCase):
    @patch("web_outlook_app.requests.post")
    def test_create_or_get_address_returns_address(self, mock_post):
        import types
        resp = types.SimpleNamespace()
        resp.status_code = 200
        resp.json = lambda: {"jwt": "j", "address": "ms-foo@bar.com", "address_id": 1}
        resp.text = "{}"
        mock_post.return_value = resp
        # 注入 base/admin/domain（避免依赖 env）——经 web_outlook_app globals
        import web_outlook_app
        web_outlook_app._CF_BASE = "https://mail.example.com"
        web_outlook_app._CF_ADMIN = "adminpw"
        web_outlook_app._CF_DOMAIN = "example.com"
        result = web_outlook_app.create_or_get_address("foo@outlook.com")
        self.assertEqual(result["address"], "ms-foo@bar.com")
```

注：若源模块用 `os.environ.get` 读 `CF_MAIL_BASE` 而非模块级常量，则测试用 `@patch.dict("web_outlook_app.os.environ", {...})`。实现时以源码实际为准，统一经 `web_outlook_app.<name>` 访问。

- [ ] **Step 5: 验证通过**

Run: `python -m pytest tests/test_project_runtime.py::CloudflareMailTests -v`
Expected: PASS

- [ ] **Step 6: 提交**

```bash
git add web_outlook_app.py outlook_web/segments/12_cloudflare_mail.py tests/test_project_runtime.py
git commit -m "feat(bind): 移植 cloudflare_mail 模块为 segment 12

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

## Task F3.2: 移植 bind_proof_in_session + 正则辅助

**Files:**
- Create: `outlook_web/segments/13_oauth_bind.py`
- Source: `D:\officeProject\yigehui\reg-factory\extract_graph_tokens.py:212-248` (`_parse_proof_add_form`, `_parse_proof_verify_form`) 与 `:251-349` (`bind_proof_in_session`) 及 `:352+` (`_fetch_latest_code_fallback`)

- [ ] **Step 1: 提取源函数到新文件**

新建 `outlook_web/segments/13_oauth_bind.py`，从 reg-factory `extract_graph_tokens.py` 复制以下函数（保持 verbatim）：
- `_parse_proof_add_form(html, url)` (:212-225)
- `_parse_proof_verify_form(html, url)` (:228-248)
- `bind_proof_in_session(session, html, url, cf_address, cf_jwt=None, use_admin=False, idx=0, cm_module=None, max_wait=180, poll=3)` (:251-349)
- `_fetch_latest_code_fallback(cm, address, jwt, use_admin)` (:352-369 附近，读完整复制)

文件头：

```python
# -*- coding: utf-8 -*-
"""OAuth 绑定辅助邮箱：在已登录 session 里把 proofs/Add 绑成 CF 临时邮箱。
移植自 reg-factory/extract_graph_tokens.py。依赖 segment 12 (cloudflare_mail) 收码。"""
import re
import urllib.parse
```

- [ ] **Step 2: 适配 cm_module 默认导入**

把 `bind_proof_in_session` 里：
```python
        from common import cloudflare_mail as cm_module
```
改为裸名引用（`cloudflare_mail` 的函数已由 `12_cloudflare_mail.py` 定义在共享 globals，但 `bind_proof_in_session` 需要 *模块对象* 来调 `cm.wait_for_code` 等）——把参数从 `cm_module` 改为直接用 `12_cloudflare_mail.py` 里被顶层 `def` 暴露的函数名。最简做法：在 `bind_proof_in_session` 顶部用裸名引用 CF 函数：

```python
def bind_proof_in_session(session, html, url, cf_address, cf_jwt=None, use_admin=False, idx=0,
                          cm_module=None, max_wait=180, poll=3):
    """..."""
    tag = f"[#{idx}]"
    # cm_module 兼容旧签名；新实现直接用共享 globals 里的 cf 函数（裸名）
    cf_wait_for_code = wait_for_code          # 由 12_cloudflare_mail.py 定义
    cf_fetch_admin_mails = fetch_admin_mails
    cf_fetch_parsed_mails = fetch_parsed_mails
    ...
```

把函数体内所有 `cm.wait_for_code(...)`→`cf_wait_for_code(...)`、`cm.fetch_admin_mails(...)`→`cf_fetch_admin_mails(...)`、`cm.fetch_parsed_mails(...)`→`cf_fetch_parsed_mails(...)`、`cm.parse_admin_mail(...)`→`parse_admin_mail(...)`（裸名，需 `12_cloudflare_mail.py` 顶层定义这些函数——移植时确认它们都是顶层 `def`）。

`_graph_log` 用最小实现（同 segment 内顶层 `def`，裸名可用）：
```python
def _graph_log(tag, msg, level="INFO"):
    import sys
    print(f"{level} {tag} {msg}", file=sys.stderr)
```
（若 `11_routes_graph_oauth.py` 已有 `graph_oauth_log` 可直接复用——它是共享 globals 裸名。）

- [ ] **Step 3: 写测试（HTML 固件）**

新增 `BindProofTests`，用录制的 HTML 片段验证 `_parse_proof_add_form` / `_parse_proof_verify_form`，经 `web_outlook_app` 访问：

```python
class BindProofTests(unittest.TestCase):
    def test_parse_proof_add_form_extracts_action_and_hidden_inputs(self):
        import web_outlook_app
        html = '''
        <form action="/proofs/Add?canary=ABC" method="post">
          <input type="hidden" name="canary" value="ABC"/>
          <input type="hidden" name="hid" value="X"/>
        </form>'''
        action, data = web_outlook_app._parse_proof_add_form(html, "https://account.live.com/proofs/Add")
        self.assertIsNotNone(action)
        self.assertEqual(data.get("canary"), "ABC")

    def test_parse_proof_verify_form_extracts_epid_action(self):
        import web_outlook_app
        html = '''
        <form action="/proofs/Verify?epid=ZZ" method="post">
          <input type="hidden" name="canary" value="C"/>
        </form>'''
        action, data = web_outlook_app._parse_proof_verify_form(html, "https://account.live.com/proofs/Verify")
        self.assertIn("epid", action)
```

- [ ] **Step 4: 验证通过**

Run: `python -m pytest tests/test_project_runtime.py::BindProofTests -v`
Expected: PASS

- [ ] **Step 5: 提交**

```bash
git add outlook_web/segments/13_oauth_bind.py tests/test_project_runtime.py
git commit -m "feat(bind): 移植 bind_proof_in_session 为 segment 13

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

## Task F3.3: 接入现有 OAuth 流程的 proofs/Add 分支

**Files:**
- Modify: `outlook_web/segments/11_routes_graph_oauth.py` (`extract_graph_refresh_token` 签名 + proofs/Add 分支 :319-336)
- Modify: `outlook_web/segments/02_groups_accounts.py` (`upsert_graph_authorized_account` :414 透传 recovery)

- [ ] **Step 1: 写失败测试**

新增 `OauthBindIntegrationTests`，mock `bind_proof_in_session` 与 `cloudflare_mail.create_or_get_address`，断言 `extract_graph_refresh_token` 在 `bind_secondary=True` 时调用 bind 并把 recovery 写入：

```python
class OauthBindIntegrationTests(unittest.TestCase):
    def setUp(self):
        import tempfile, os
        from outlook_web import web_outlook_app
        self._tmp_dir = tempfile.mkdtemp()
        self._orig_cwd = os.getcwd()
        os.chdir(self._tmp_dir)
        os.makedirs("data", exist_ok=True)
        web_outlook_app.init_app(reinit_db=True)
        from outlook_web.segments import bootstrap as bs
        bs.set_login_password("p")

    def tearDown(self):
        import os, shutil
        os.chdir(self._orig_cwd)
        shutil.rmtree(self._tmp_dir, ignore_errors=True)

    @patch("outlook_web.segments.seg_13_oauth_bind.bind_proof_in_session")
    @patch("outlook_web.segments.seg_12_cloudflare_mail.create_or_get_address")
    def test_bind_path_called_when_bind_secondary_true(self, mock_create, mock_bind):
        import types
        from outlook_web.segments import graph_oauth as go
        mock_create.return_value = {"address": "ms-foo@cf.com", "jwt": "j", "password": "auxpw"}
        # 构造一个落在 proofs/Add 的 session 响应链，bind 返回一个带 code 的 localhost 响应
        bound_resp = types.SimpleNamespace()
        bound_resp.status_code = 302
        bound_resp.headers = {"Location": "http://localhost/?code=THECODE"}
        bound_resp.url = "http://localhost/?code=THECODE"
        bound_resp.text = ""
        mock_bind.return_value = bound_resp
        # mock 整个 session 交互太重——此处只验证 create_or_get_address 与 bind 被调用。
        # 真实端到端在手动验证步骤。这里跳过全 session mock，断言函数参数透传即可。
        # （若实现把 bind 调用与 create_or_get_address 调用放在 proofs/Add 分支，单测改为直接调分支函数。）
```

注：OAuth 全 session 难以单测（需 mock 整条 HTTP 链）。本任务的真正保护是：**(a)** bind 仅在 `bind_secondary=True` 且到达 proofs/Add 时调用；**(b)** bind 失败不致命——回退到原 `action="Skip"`。用更聚焦的测试替代上面的占位：

```python
    def test_upsert_graph_authorized_account_persists_recovery_fields(self):
        from outlook_web.segments import graph_oauth as go
        from outlook_web.segments import groups_accounts as ga
        # 模拟绑定成功后 upsert 传入 recovery
        result = ga.upsert_graph_authorized_account(
            "bindacc@x.com", "pw", "cid", "rt",
            recovery_email="ms-bindacc@cf.com",
            recovery_email_password="auxpw",
            authorization_type="graph",
        )
        with _open_db() as conn:
            row = conn.execute("SELECT recovery_email, recovery_email_password FROM accounts WHERE email='bindacc@x.com'").fetchone()
        self.assertEqual(row["recovery_email"], "ms-bindacc@cf.com")
        # recovery_email_password 应加密存储
        self.assertTrue(row["recovery_email_password"].startswith("enc:") or row["recovery_email_password"] != "auxpw")
```

- [ ] **Step 2: 验证失败**

Run: `python -m pytest tests/test_project_runtime.py::OauthBindIntegrationTests::test_upsert_graph_authorized_account_persists_recovery_fields -v`
Expected: FAIL — `upsert_graph_authorized_account` 不接受 `recovery_email`/`recovery_email_password` kwargs。

- [ ] **Step 3: upsert_graph_authorized_account 透传 recovery**

在 `outlook_web/segments/02_groups_accounts.py` 的 `upsert_graph_authorized_account`（:414）签名末尾加两参，并在两条 UPDATE/INSERT 透传。

签名改为（加在 `authorization_type` 后）：
```python
def upsert_graph_authorized_account(email: str, password: str, client_id: str,
                                    refresh_token: str, *,
                                    group_id: Any = None,
                                    proxy_url: str = '',
                                    tag_ids: Any = None,
                                    remark: str = '',
                                    authorization_type: Optional[str] = None,
                                    recovery_email: str = '',
                                    recovery_email_password: str = '') -> Dict[str, Any]:
```

existing 分支（:441 UPDATE）SET 子句加两列、参数元组加两值：
```python
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
            (encrypted_password, client_id, encrypted_refresh_token,
             normalized_authorization_type,
             recovery_email,
             encrypt_data(recovery_email_password) if recovery_email_password else recovery_email_password,
             account_id),
        )
```

新建分支（:462 `build_account_insert_values(...)`）末尾传两参：
```python
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
```
（`build_account_insert_values` 末两参即前序工作已加的 `recovery_email`/`recovery_email_password`。）

- [ ] **Step 4: 验证通过**

Run: `python -m pytest tests/test_project_runtime.py::OauthBindIntegrationTests -v`
Expected: PASS

- [ ] **Step 5: 接入 proofs/Add 分支调 bind**

在 `outlook_web/segments/11_routes_graph_oauth.py`：

1. 顶部 import：**无需 import**（共享 globals 裸名）。`create_or_get_address` 与 `bind_proof_in_session` 由 `12_cloudflare_mail.py`/`13_oauth_bind.py` 顶层定义并已在 `SEGMENT_FILES` 中。在 `extract_graph_refresh_token` 函数体内直接用裸名 `create_or_get_address(...)`、`bind_proof_in_session(...)` 即可。注意加载顺序���`12`/`13` 在 `11` 之后加载，但函数体在运行时才执行，运行时引用成立。

2. `extract_graph_refresh_token` 签名加参：
```python
def extract_graph_refresh_token(email, password, client_id, redirect_uri, scope,
    authority=GRAPH_EXTRACT_AUTHORITY, log=None, session_factory=None,
    proxy_url=None, bind_secondary=False):
```

3. proofs/Add 分支（:319-336）改为：
```python
            if "proofs/Add" in current_url or "proofs/add" in current_url:
                if bind_secondary:
                    # 绑定 CF 辅助邮箱：create_or_get_address -> bind_proof_in_session
                    try:
                        cf_info = cf_create_or_get_address(email)
                        cf_address = cf_info.get("address")
                        cf_jwt = cf_info.get("jwt")
                        cf_pw = cf_info.get("password") or ""
                        use_admin = cf_info.get("use_admin", False)
                    except Exception as exc:
                        graph_oauth_log(log, f"CF 辅助邮箱分配失败，回退 Skip: {exc}")
                        cf_address = None
                    if cf_address:
                        bound_resp = bind_proof_in_session(
                            session, text, current_url,
                            cf_address=cf_address, cf_jwt=cf_jwt,
                            use_admin=use_admin, idx=0,
                        )
                        if bound_resp is not None:
                            resp2 = bound_resp
                            _pending_recovery_email = cf_address
                            _pending_recovery_password = cf_pw
                            continue
                        graph_oauth_log(log, "bind 失败，回退 Skip proofs/Add")
                # 回退/未启用绑定：原 Skip 逻辑
                form_match = re.search(
                    r'<form[^>]*action="([^"]+)"[^>]*>(.*?)</form>',
                    text, re.DOTALL | re.IGNORECASE,
                )
                if not form_match:
                    return make_graph_oauth_response(False, "安全信息页面处理失败", "无法找到表单")
                graph_oauth_log(log, "跳过 Microsoft 安全信息添加页面")
                form_data = extract_hidden_inputs(form_match.group(2))
                form_data["action"] = "Skip"
                resp2 = session.post(
                    absolute_form_action(form_match.group(1), current_url),
                    data=form_data, timeout=30, allow_redirects=False,
                )
                continue
```

4. 在函数开头初始化 `_pending_recovery_email = ""` / `_pending_recovery_password = ""`，在最终 return 成功字典处把它们带出：
```python
        return {
            "success": True,
            "refresh_token": refresh_token,
            "client_id": client_id,
            "recovery_email": _pending_recovery_email,
            "recovery_email_password": _pending_recovery_password,
        }
```

5. `save_graph_authorization_result`（:510）把这两个值透传给 `upsert_graph_authorized_account`——读 `run_graph_oauth_task` 调 `extract_graph_refresh_token` 的返回，取 `recovery_email`/`recovery_email_password` 传入。`save_graph_authorization_result` 签名加两参并透传到 `upsert_graph_authorized_account(...recovery_email=..., recovery_email_password=...)`。

- [ ] **Step 6: 手动验证（端到端，不打单测）**

导入一个真实未绑定辅助邮箱的 outlook 账号到暂存表，调 `POST /api/oauth/graph-extract-token` 带 `bind_secondary=true`（注：当前单账号路由暂未透传 bind_secondary，下个 Task 加；此处先手动在代码里临时设 `bind_secondary=True` 跑一次），观察 CF 收码 + 主表 `recovery_email` 写入。恢复临时改动。

- [ ] **Step 7: 回归**

Run: `python -m pytest tests/ -q 2>&1 | tail -10`
Expected: 全绿

- [ ] **Step 8: 提交**

```bash
git add outlook_web/segments/11_routes_graph_oauth.py outlook_web/segments/02_groups_accounts.py tests/test_project_runtime.py
git commit -m "feat(bind): OAuth 流程接入辅助邮箱绑定，upsert 透传 recovery 字段

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

## Task F3.4: 批量授权端点（并行）

**Files:**
- Modify: `outlook_web/segments/11_routes_graph_oauth.py`（新增 `api_graph_extract_batch` + worker）
- Test: `tests/test_project_runtime.py`

- [ ] **Step 1: 写失败测试**

新增 `BatchAuthorizeApiTests`：

```python
class BatchAuthorizeApiTests(unittest.TestCase):
    def setUp(self):
        import tempfile, os
        from outlook_web import web_outlook_app
        self._tmp_dir = tempfile.mkdtemp()
        self._orig_cwd = os.getcwd()
        os.chdir(self._tmp_dir)
        os.makedirs("data", exist_ok=True)
        web_outlook_app.init_app(reinit_db=True)
        from outlook_web.segments import bootstrap as bs
        bs.set_login_password("p")
        self.client = web_outlook_app.app.test_client()
        self.client.post("/api/login", data={"password": "p"})
        # 暂存表插入 3 个待授权账号
        from outlook_web.segments import groups_accounts as ga
        ga.add_upload_account("batch1@x.com", "pw", group_id=1)
        ga.add_upload_account("batch2@x.com", "pw", group_id=1)
        ga.add_upload_account("batch3@x.com", "pw", group_id=1)

    def tearDown(self):
        import os, shutil
        os.chdir(self._orig_cwd)
        shutil.rmtree(self._tmp_dir, ignore_errors=True)

    def test_batch_endpoint_returns_task_id_and_stream_url(self):
        with _open_db() as conn:
            ids = [r["id"] for r in conn.execute("SELECT id FROM outlook_upload_accounts ORDER BY id").fetchall()]
        resp = self.client.post(
            "/api/oauth/graph-extract-batch",
            json={"account_ids": ids, "bind_secondary": False, "max_workers": 2},
        )
        self.assertEqual(resp.status_code, 200)
        data = resp.get_json()
        self.assertTrue(data["success"])
        self.assertIn("task_id", data)
        self.assertTrue(data["stream_url"].startswith("/api/oauth/graph-extract-batch/"))

    @patch("outlook_web.segments.graph_oauth.run_graph_oauth_task")
    def test_batch_worker_processes_all_accounts(self, mock_single):
        mock_single.return_value = {"success": True}
        from outlook_web.segments import graph_oauth as go
        with _open_db() as conn:
            ids = [r["id"] for r in conn.execute("SELECT id FROM outlook_upload_accounts ORDER BY id").fetchall()]
        import queue
        q = queue.Queue()
        results = go.run_batch_oauth_task(ids, q, mode="graph", bind_secondary=False, max_workers=2)
        self.assertEqual(mock_single.call_count, 3)
```

- [ ] **Step 2: 验证失败**

Run: `python -m pytest tests/test_project_runtime.py::BatchAuthorizeApiTests -v`
Expected: FAIL — 路由与 `run_batch_oauth_task` 不存在。

- [ ] **Step 3: 实现批量 worker + 路由**

在 `outlook_web/segments/11_routes_graph_oauth.py` 新增 worker（复用 F1 并行范式）：

```python
def run_batch_oauth_task(account_ids, output_queue, *, mode="graph", bind_secondary=False, max_workers=5):
    """并行跑单账号 OAuth 授权。复用现有 run_graph_oauth_task 单账号逻辑。"""
    from concurrent.futures import ThreadPoolExecutor, as_completed
    total = len(account_ids)
    output_queue.put({"type": "start", "total": total, "mode": mode, "bind_secondary": bind_secondary})
    completed = 0
    success_count = 0
    failed_list = []

    def _do_one(account_id):
        # 复用单账号 task，但用独立 queue 收尾再汇总
        import queue as _q
        sub_q = _q.Queue()
        run_graph_oauth_task(account_id, sub_q, mode, bind_secondary=bind_secondary)
        # 取最后一条结果
        last = None
        while not sub_q.empty():
            last = sub_q.get_nowait()
        return account_id, last

    workers = min(max(1, max_workers), total) if total else 1
    with ThreadPoolExecutor(max_workers=workers, thread_name_prefix="oauth-batch") as executor:
        future_map = {executor.submit(_do_one, aid): aid for aid in account_ids}
        for future in as_completed(future_map):
            account_id, last = future.result()
            completed += 1
            ok = bool(last and last.get("success"))
            if ok:
                success_count += 1
            else:
                failed_list.append({"account_id": account_id, "error": (last or {}).get("error", "")})
            output_queue.put({
                "type": "progress",
                "index": completed,
                "total": total,
                "account_id": account_id,
                "success": ok,
            })
    summary = {"type": "complete", "total": total, "success_count": success_count, "failed": failed_list}
    output_queue.put(summary)
    return summary
```

注：`run_graph_oauth_task` 当前签名是 `(account_id, output_queue, mode)`——需加 `bind_secondary=False` 参数并透传给 `extract_graph_refresh_token`。改 `run_graph_oauth_task` 签名 + 调用处。

新增路由（紧邻单账号 `/api/oauth/graph-extract-token` :651）：

```python
@app.route('/api/oauth/graph-extract-batch', methods=['POST'])
@login_required
def api_graph_extract_batch():
    data = request.get_json(silent=True) or {}
    account_ids = data.get('account_ids') or []
    if not isinstance(account_ids, list) or not account_ids:
        return jsonify({'success': False, 'error': 'account_ids 不能为空'}), 400
    mode = data.get('mode', 'graph')
    bind_secondary = bool(data.get('bind_secondary', True))
    max_workers = min(20, max(1, int(data.get('max_workers', 5))))
    task_id = uuid.uuid4().hex
    GRAPH_OAUTH_TASKS[task_id] = {'account_ids': account_ids, 'mode': mode,
                                  'bind_secondary': bind_secondary, 'max_workers': max_workers}
    return jsonify({'success': True, 'task_id': task_id,
                    'stream_url': f'/api/oauth/graph-extract-batch/{task_id}/stream'})


@app.route('/api/oauth/graph-extract-batch/<task_id>/stream')
@login_required
def api_graph_extract_batch_stream(task_id):
    task = GRAPH_OAUTH_TASKS.pop(task_id, None)
    if not task:
        return jsonify({'success': False, 'error': '任务不存在或已完成'}), 404

    def generate():
        import queue
        out_q = queue.Queue()
        worker = threading.Thread(
            target=run_batch_oauth_task,
            args=(task['account_ids'], out_q),
            kwargs={'mode': task['mode'], 'bind_secondary': task['bind_secondary'],
                    'max_workers': task['max_workers']},
            daemon=True,
        )
        worker.start()
        while True:
            try:
                payload = out_q.get(timeout=120)
            except queue.Empty:
                yield f"event: ping\ndata: {{}}\n\n"
                continue
            yield f"data: {json.dumps(payload)}\n\n"
            if payload.get('type') == 'complete':
                break

    return Response(generate(), mimetype='text/event-stream')
```

- [ ] **Step 4: 验证通过**

Run: `python -m pytest tests/test_project_runtime.py::BatchAuthorizeApiTests -v`
Expected: PASS

- [ ] **Step 5: 回归**

Run: `python -m pytest tests/ -q 2>&1 | tail -10`
Expected: 全绿

- [ ] **Step 6: 提交**

```bash
git add outlook_web/segments/11_routes_graph_oauth.py tests/test_project_runtime.py
git commit -m "feat(authorize): 批量并行授权端点 /api/oauth/graph-extract-batch

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

## Task F3.5: 单账号路由透传 bind_secondary + 前端 UI

**Files:**
- Modify: `outlook_web/segments/11_routes_graph_oauth.py`（`api_graph_extract_token` :651 透传 `bind_secondary`）
- Modify: `templates/partials/index/dialogs-primary.html`
- Modify: `static/js/index/07-settings.js`

- [ ] **Step 1: 单账号路由透传**

`api_graph_extract_token`（:651）读 `bind_secondary = bool(data.get('bind_secondary', True))`，存入 `GRAPH_OAUTH_TASKS[task_id]`，stream worker 透传给 `run_graph_oauth_task(..., bind_secondary=...)`。

- [ ] **Step 2: 前端 UI（暂存页批量授权按钮）**

在 `templates/partials/index/dialogs-primary.html` 暂存账号区块（搜 `outlook_upload` 或现有「批量」按钮附近）加：

```html
<button type="button" class="btn btn-primary" id="batchAuthorizeBtn" onclick="batchAuthorizeUploadAccounts()">
  批量授权
</button>
<label class="ml-3">
  <input type="checkbox" id="batchBindSecondary" checked> 绑定辅助邮箱并授权
</label>
```

- [ ] **Step 3: 前端 JS**

在 `static/js/index/07-settings.js` 加：

```javascript
async function batchAuthorizeUploadAccounts() {
  const ids = getSelectedUploadAccountIds(); // 复用现有选中逻辑，若无则取全部
  if (!ids || !ids.length) { alert('请先选择账号'); return; }
  const bindSecondary = document.getElementById('batchBindSecondary').checked;
  const res = await fetch('/api/oauth/graph-extract-batch', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({account_ids: ids, bind_secondary: bindSecondary, max_workers: 5})
  });
  const data = await res.json();
  if (data.success) {
    streamBatchAuthorizeProgress(data.stream_url);
  } else {
    alert(data.error || '批量授权启动失败');
  }
}

function streamBatchAuthorizeProgress(streamUrl) {
  const evt = new EventSource(streamUrl);
  evt.onmessage = (e) => {
    const p = JSON.parse(e.data);
    if (p.type === 'progress') {
      // 复用现有进度 UI 更新；无则 console
      console.log(`授权 ${p.index}/${p.total} ${p.success ? '✓' : '✗'}`);
    } else if (p.type === 'complete') {
      evt.close();
      alert(`完成：成功 ${p.success_count}/${p.total}，失败 ${p.failed.length}`);
      reloadUploadAccountsTable();
    }
  };
  evt.onerror = () => evt.close();
}
```

- [ ] **Step 4: 手动验证**

启动应用 → 暂存页选 2 个未授权账号 → 勾选「绑定辅助邮箱并授权」→ 点批量授权 → 观察 SSE 进度 → 成功后主表有 `recovery_email`、暂存行 `is_authorized=1`。

- [ ] **Step 5: 提交**

```bash
git add outlook_web/segments/11_routes_graph_oauth.py templates/partials/index/dialogs-primary.html static/js/index/07-settings.js
git commit -m "feat(authorize): 单账号路由透传 bind_secondary + 批量授权前端

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

## Task F3.6: CF 设置项 + 文档

**Files:**
- Modify: `outlook_web/segments/01_bootstrap.py`（CF 设置种子）
- Modify: `outlook_web/segments/07_routes_oauth_settings_external.py`（设置写入接受 CF keys）
- Modify: `docs/api.md`

- [ ] **Step 1: CF 设置种子**

在 `01_bootstrap.py` 设置种子区追加：
```python
        ('cf_mail_base', ''),
        ('cf_mail_admin', ''),
        ('cf_mail_domain', ''),
        ('cf_mail_site_pass', ''),
        ('cf_mail_proxy', ''),
```

- [ ] **Step 2: 设置写入接受**

在 `07_routes_oauth_settings_external.py` 设置写入路由（与 `refresh_parallel_workers` 同处）追加 5 个 CF key 的接受与 `set_setting`。

- [ ] **Step 3: 更新 docs/api.md**

在 `docs/api.md` 末尾追加两节：

```markdown
## POST /api/external/accounts/import （API-Key）

推送 ---- 分隔的账号文本到主 accounts 表。

Header: `X-API-Key: <external_api_key>`
Body:
```json
{"account_string": "主邮箱----主密码----client_id----refresh_token----辅助邮箱----辅助邮箱密码\n...",
 "group_id": 1, "account_format": "client_id_refresh_token", "provider": "outlook", "tag_ids": "1,2"}
```
返回: `{"success": true, "added_count": N, "skipped_count": N, "invalid_count": N, "tagged_count": N}`
未配置 Key: 403；Key 无效: 401。

## POST /api/oauth/graph-extract-batch （需登录）

批量并行授权暂存账号。

Body: `{"account_ids": [1,2], "mode": "graph", "bind_secondary": true, "max_workers": 5}`
返回: `{"success": true, "task_id": "...", "stream_url": "/api/oauth/graph-extract-batch/<id>/stream"}`

进度经 SSE 流推送 progress/complete 事件。
```

- [ ] **Step 4: 提交**

```bash
git add outlook_web/segments/01_bootstrap.py outlook_web/segments/07_routes_oauth_settings_external.py docs/api.md
git commit -m "feat(settings): CF 邮箱设置项 + API 文档

Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>"
```

---

## Task F3.7: 全量回归 + release 提交

**Files:** (none)

- [ ] **Step 1: 全量测试**

Run: `python -m pytest tests/ -q 2>&1 | tail -15`
Expected: 全绿（基线 557 + 全部新增）

- [ ] **Step 2: 合并到 main（或开 PR）**

```bash
git checkout main
git merge --no-ff feat/refresh-parallel-and-bind-authorize -m "merge: 刷新并行化 + txt上传接口 + 绑定辅助邮箱/批量授权"
```

或推送开 PR：
```bash
git push -u origin feat/refresh-parallel-and-bind-authorize
```

---

## Self-Review 自查（已在编写时完成）

1. **Spec 覆盖**：F1（并发1-20默认5 ✓ F1.2、429退避 ✓ F1.1、镜像转发并行 ✓ F1.3、四循环接入 ✓ F1.4、串行回退 ✓ F1.2 mode）；F2（API-Key 端点 ✓ F2.1、复用 parse_account_import ✓、写主表 ✓、6/4段兼容 ✓）；F3（移植 cloudflare_mail ✓ F3.1、移植 bind ✓ F3.2、接入 proofs/Add ✓ F3.3、recovery 透传 ✓ F3.3、批量并行 ✓ F3.4、CF 自动分配 ✓ F3.3、已绑定跳过 ✓ F3.3 bind None 回退 Skip、UI ✓ F3.5、设置 ✓ F3.6、CF 不可用回退 ✓ F3.3 except 回退 Skip）。
2. **占位符扫描**：F3.3 Step1 注明了「占位」并给出聚焦替代测试；F3.3 Step5 命名约定注明「实现时确认」。其余步骤均有完整代码。
3. **类型一致**：`refresh_accounts_parallel` 参数在 F1.3 定义与 F1.4 调用一致；`upsert_graph_authorized_account` 的 recovery 两参在 F3.3 定义与调用一致；`run_batch_oauth_task` 在 F3.4 定义与测试一致。
