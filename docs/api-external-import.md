# 外部账号导入 API

通过 HTTP 请求批量导入邮箱账号到主 `accounts` 表，**支持指定分组**。无需浏览器登录态，使用 API Key 鉴权。

对应代码：`outlook_web/segments/07_routes_oauth_settings_external.py` 的 `api_external_import_accounts`（commit `51f9ef1` 新增）。复用前端 `POST /api/accounts` 的解析与批量写入管道（`parse_account_import` → `add_accounts_bulk`），所以导入能力与界面导入完全一致，包括「辅助邮箱 + 辅助邮箱密码」6 段格式。

---

## 1. 鉴权

所有外部导入请求必须携带 API Key，三选一：

| 传递方式 | 示例 |
|---|---|
| 请求头（推荐） | `X-API-Key: <你的 key>` |
| 查询参数 | `?api_key=<你的 key>` |
| 查询参数（别名） | `?apikey=<你的 key>` |

- API Key 在「系统设置 → 对外 API Key」里配置，存库为 `external_api_key`（`get_external_api_key()` 读取）。
- 服务端用 `secrets.compare_digest` 常量时间比较，未配置 key 返回 `403`，key 不匹配返回 `401`，缺少 key 返回 `401`。
- 本接口 `@csrf_exempt`，不校验 CSRF Token。

---

## 2. 请求

```
POST /api/external/accounts/import
Content-Type: application/json
X-API-Key: <你的 key>
```

### 请求体（JSON）

| 字段 | 类型 | 必填 | 默认 | 说明 |
|---|---|---|---|---|
| `account_string` | string | **是** | — | 多行账号文本，每行一个账号，字段以 `----` 分隔。空行自动跳过。 |
| `group_id` | int | 否 | `1` | **分组 ID**。导入的所有账号归入此分组。默认 `1` 为系统默认分组。 |
| `account_format` | string | 否 | `client_id_refresh_token` | 仅 Outlook 生效，决定第 3/4 段顺序：`client_id_refresh_token`（第 3 段 client_id、第 4 段 refresh_token）或 `refresh_token_client_id`（反之）。 |
| `provider` | string | 否 | `outlook` | 邮箱服务商，决定解析格式与 IMAP 默认主机。见下方取值表。 |
| `imap_host` | string | 否 | `""` | `provider=custom` 时必填；其它 provider 留空则用内置默认主机。 |
| `imap_port` | int | 否 | `993` | IMAP 端口，`custom` 时可覆盖。非法值返回 `400`。 |
| `tag_ids` | int[] | 否 | `[]` | 标签 ID 数组，导入成功的账号会打上这些标签（标签需已存在，不存在的 ID 被忽略）。 |
| `proxy_url` | string | 否 | `""` | 主代理 URL，可含字面量 `{mail}`（运行时替换为邮箱 local-part）。 |
| `fallback_proxy_url_1` | string | 否 | `""` | 备用代理 1。 |
| `fallback_proxy_url_2` | string | 否 | `""` | 备用代理 2。 |
| `forward_enabled` | bool | 否 | `false` | 是否开启邮件转发。 |

> `group_id` 是整数主键。系统默认分组通常为 `1`。若要导入到自定义分组，需先知道该分组的 ID——见本文档第 5 节「如何拿到 group_id」。

### `provider` 取值

| provider | 说明 | 账号行格式 |
|---|---|---|
| `outlook` | Outlook/Hotmail/Live（走 OAuth，非 IMAP） | 4~6 段 |
| `gmail` | Gmail | 2 段（邮箱+密码），IMAP 主机内置 |
| `qq` | QQ 邮箱 | 2 段 |
| `163` | 163 邮箱 | 2 段 |
| `126` | 126 邮箱 | 2 段 |
| `yahoo` | Yahoo | 2 段 |
| `aliyun` | 阿里邮箱 | 2 段 |
| `2925` | 2925 邮箱 | 2 段 |
| `custom` | 自定义 IMAP | 2 或 4 段（4 段时第 3 段 host、第 4 段 port） |
| `auto` | 由邮箱域名自动推断 | — |

非 `outlook` 的 provider 均按 IMAP 账号解析（`account_type=imap`），密码写入 `imap_password`，`client_id`/`refresh_token` 留空。

### 账号行格式

字段之间用 `----`（四个减号）分隔，行内空白会被 trim。

**Outlook（provider=outlook）**

```
邮箱----主密码----client_id----refresh_token----辅助邮箱----辅助邮箱密码
```

- 4 段：基础账号（无辅助邮箱）
- 6 段：含辅助邮箱地址（第 5 段）和辅助邮箱密码（第 6 段，**入库加密存储**）
- 第 3/4 段顺序由 `account_format` 决定

**IMAP 邮箱（provider=gmail/qq/163/126/yahoo/aliyun/2925）**

```
邮箱----密码
```

主机/端口按 provider 内置默认值，无需在行内提供。

**自定义 IMAP（provider=custom）**

- 2 段：`邮箱----密码`，必须配合请求体的 `imap_host`/`imap_port`
- 4 段：`邮箱----密码----imap_host----imap_port`，行内覆盖请求体

### 解析失败处理

每行独立解析。解析失败的行计入 `invalid_count`，不会中断整体导入；解析成功的行进入批量写入。**邮箱已存在的行会被跳过**（计入 `skipped_count`），事务单次提交。

---

## 3. 响应

### 成功（HTTP 200）

```json
{
  "success": true,
  "added_count": 3,
  "skipped_count": 1,
  "invalid_count": 2,
  "tagged_count": 3
}
```

| 字段 | 说明 |
|---|---|
| `added_count` | 实际新增入库的账号数（重复邮箱不计） |
| `skipped_count` | 解析成功但邮箱已存在、未写入的行数 |
| `invalid_count` | 解析失败的行数 |
| `tagged_count` | 打了标签的新增账号数（仅当传了 `tag_ids` 且非空才非零） |

### 失败

| HTTP | 场景 | 响应体 |
|---|---|---|
| 400 | `account_string` 为空 | `{"success": false, "error": "请输入账号信息"}` |
| 400 | `imap_port` 非整数 | `{"success": false, "error": "IMAP 端口无效"}` |
| 401 | 缺少 API Key | `{"success": false, "error": "缺少 API Key，请通过 Header X-API-Key 或查询参数 api_key 提供"}` |
| 401 | API Key 不匹配 | `{"success": false, "error": "API Key 无效"}` |
| 403 | 服务端未配置对外 API Key | `{"success": false, "error": "未配置对外 API Key，请在系统设置中配置"}` |

---

## 4. 示例

### 4.1 导入 Outlook 账号到指定分组（含辅助邮箱）

```bash
curl -X POST http://127.0.0.1:5000/api/external/accounts/import \
  -H "Content-Type: application/json" \
  -H "X-API-Key: 38cc7780e3adbc1cf921deba9b891b22bc15e834dd1a7b4f9ba3f7eb625ecc14" \
  -d '{
    "group_id": 5,
    "account_string": "user1@outlook.com----Pass123----9e5f94bc-...----eyJ0eXAiOi...----aux1@cf.com----AuxPass456\nuser2@hotmail.com----Pass234----9e5f94bc-...----eyJ0eXAiOi..."
  }'
```

返回：

```json
{"success": true, "added_count": 2, "skipped_count": 0, "invalid_count": 0, "tagged_count": 0}
```

### 4.2 导入 QQ 邮箱到默认分组（group_id=1）

```bash
curl -X POST http://127.0.0.1:5000/api/external/accounts/import \
  -H "Content-Type: application/json" \
  -H "X-API-Key: <key>" \
  -d '{
    "provider": "qq",
    "account_string": "123456@qq.com----授权码A\n789012@qq.com----授权码B"
  }'
```

### 4.3 导入自定义 IMAP 账号并打标签

```bash
curl -X POST http://127.0.0.1:5000/api/external/accounts/import \
  -H "Content-Type: application/json" \
  -H "X-API-Key: <key>" \
  -d '{
    "provider": "custom",
    "group_id": 3,
    "imap_host": "imap.example.com",
    "imap_port": 993,
    "tag_ids": [7, 12],
    "proxy_url": "socks5://user:pass@proxy:1080",
    "account_string": "a@example.com----pwdA----imap.example.com----993\nb@example.com----pwdB"
  }'
```

### 4.4 Python 调用

```python
import requests

resp = requests.post(
    "http://127.0.0.1:5000/api/external/accounts/import",
    headers={
        "Content-Type": "application/json",
        "X-API-Key": "38cc7780e3adbc1cf921deba9b891b22bc15e834dd1a7b4f9ba3f7eb625ecc14",
    },
    json={
        "group_id": 5,
        "account_string": "user@outlook.com----Pass123----client-id----refresh-token----aux@cf.com----auxpass",
    },
    timeout=30,
)
print(resp.status_code, resp.json())
```

---

## 5. 如何拿到 group_id

分组 ID 是整数主键。`group_id` 不传时默认 `1`（系统默认分组）。若要导入到自建分组：

- **从界面查**：登录 Web 界面 → 分组管理，分组编辑 URL 里的数字即 ID。
- **从数据库查**：`SELECT id, name FROM groups WHERE name != '临时邮箱' ORDER BY id;`
- **创建新分组**：目前分组创建/查询接口（`GET/POST /api/groups`）是 `@login_required`，**不走 API Key**，无法用同一把外部 key 调用。需要先用浏览器登录态创建分组拿到 ID，再把这个 ID 用于外部导入。

  `POST /api/groups`（需登录态 Cookie）请求体示例：
  ```json
  {"name": "新分组名", "description": "", "color": "#1a1a1a", "parent_id": null}
  ```
  返回 `{"success": true, "group_id": 7}`。

> 如果你的场景需要「纯 API Key 创建/查询分组」（不依赖浏览器登录态），目前接口不支持，需要新增对外端点。可以提需求。

---

## 6. 约束与注意

- **单 worker**：服务需单 worker 运行（官方 Docker 已固定 Gunicorn 单 worker + 多线程）。外部导入用进程内 SQLite 单事务，多 worker 下重复邮箱的跳过判断可能竞态。
- **明文凭据**：`account_string` 里的主密码、refresh_token、辅助邮箱密码都是明文传输，请走 HTTPS 或内网。主密码与 refresh_token、辅助邮箱密码均在库内加密存储（`encrypt_data`）；辅助邮箱地址明文存。
- **重复邮箱**：以 `email` 唯一，已存在则跳过（不更新），计入 `skipped_count`。
- **批量大小**：无硬性行数上限，但单次请求体过大会受 Web 服务器 body 限制约束，建议单次几百行内，分批调用。
- **超时**：导入是同步处理，行数很多时响应可能较慢，调用方建议设 `timeout >= 30s`。
