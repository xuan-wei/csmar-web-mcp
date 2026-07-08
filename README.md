# CSMAR Web-API MCP

An MCP server that lets Claude (or any MCP client) **search, browse, preview, and download CSMAR (国泰安) data** — **no account, no password, no browser, no official SDK**.

Authentication is done purely through your **institutional IP** (e.g. a subscribing university's campus network): the CSMAR website (`data.csmar.com`) grants an access token to recognized IPs, and this server replicates the website's `IP-token + JSON` HTTP API with `requests` (plus one WebSocket for download-ready notifications).

> ⚠️ **You must be on a network IP that already has CSMAR access** (a campus / institutional network that subscribes to CSMAR). Auth is IP-based — there is no login form. On any other network `csmar_status` will report `IP授权: false` and nothing will work. This tool only replicates the public web endpoints you are already authorized to use; it does **not** bypass access control or share accounts. Follow CSMAR's Terms of Service.

---

## English

### Requirements
- **A network IP with CSMAR access** (subscribing institution / campus network).
- Python 3.10+ with:
  ```bash
  pip install mcp requests websocket-client pandas openpyxl
  ```

### Install (project-scoped)
1. Copy the example config and edit paths for your machine:
   ```bash
   cp .mcp.json.example .mcp.json
   # edit "command" -> your python, and the server.py absolute path
   ```
2. **Restart your MCP client** (e.g. Claude Code) to load the `csmar` server.

Environment variables (in `.mcp.json` `env`):
- `CSMAR_LANG`: `1` = English UI (English field names; search with English terms), `0` = Chinese UI. Default `1`.
- `CSMAR_OUT_DIR`: default directory for bulk downloads. Default `~/Downloads`.

### Tools (11)

**Discovery**

| Tool | Purpose | Key args |
|---|---|---|
| `csmar_status` | Check login / institution / IP-access status | — |
| `csmar_list_databases` | List the database catalog (with subscription flag) | `only_accessible`, `keyword` |
| `csmar_list_tables` | List tables under a database | `database_id` |
| `csmar_search` | Keyword search over tables/fields → get `table_id` | `keyword`, `page_size` |
| `csmar_list_fields` | List a table's fields, time/code field, samples | `table` (id or name) |

**Data** (`table` accepts a `table_id` or a table name)

| Tool | Purpose | Key args |
|---|---|---|
| `csmar_query_count` | Count matching records | `table`, `start_date`, `end_date`, `codes?` |
| `csmar_preview` | Preview rows as JSON (≤200 rows) | `table`, `start_date`, `end_date`, `fields?`, `codes?` |
| `csmar_query` | General query → JSON; `limit>200` auto-uses background packaging + download | `table`, `start_date`, `end_date`, `fields?`, `codes?`, `limit` |
| `csmar_download` | Download full data to local xlsx/csv | `table`, `start_date`, `end_date`, `fields?`, `codes?`, `out_dir?`, `as_csv?` |

**Convenience**

| Tool | Purpose | Key args |
|---|---|---|
| `get_stock_data` | Stock quotes by code (daily/weekly/monthly/annual, `TRD_Dalyr` etc.) | `stock_code`, `start_date`, `end_date`, `frequency`, `limit` |
| `get_company_info` | Company profile by code (`TRD_Co`) | `stock_code` |

Dates are `YYYY-MM-DD`; `fields` defaults to all; `codes` defaults to all; `frequency` ∈ daily/weekly/monthly/annual.

All four data tools also accept **`conditions`** — a list of field filters, e.g.
`[{"field":"Clsprc","op":">","value":"100"}]`. Operators: `>` `>=` `<` `<=` `=` `!=` `like` `not like` `is null` `not is null`; multiple conditions are AND-joined by default (`"relation":"or"` per item to OR).

### Example prompts
- "Check the CSMAR connection" → `csmar_status`
- "Search for stock trading tables" → `csmar_search("stock")`
- "What fields does the daily stock table have?" → `csmar_list_fields("Daily Stock Price & Returns")`
- "Preview PE/PB for the first week of July 2026" → `csmar_preview(3673,"2026-07-01","2026-07-07",["TradingDate","Symbol","PE","PB"])`
- "Download this range as Excel" → `csmar_download(3673,"2026-07-01","2026-07-07")`
- "Daily quotes for 000001 in Dec 2024" → `get_stock_data("000001","2024-12-01","2024-12-31")`

### Limits (CSMAR side)
- Preview is capped at **200 rows** (use `csmar_query`/`csmar_download` for more).
- A single query returns at most **200,000 records**; the time span per query is capped by frequency (e.g. daily ≤ 2 years). Narrow the range if it errors.
- Identical queries may be rate-limited briefly; transient slowness/timeouts are retried automatically.
- Only single-table queries (Data Query) are covered; cross-table query is not implemented.

**Deliberate non-goals**
- *Resolving a table by its physical name (e.g. `TRD_Dalyr`) cold*: CSMAR's search index does not contain physical names and there is no name→id endpoint, so pass a `table_id` (from `csmar_search`/`csmar_list_tables`) or a table display name. A physical name works only after that table has been accessed once in the session.
- *A `get_financial_data` convenience tool*: financial figures span many tables with consolidation / parent-only / report-type nuances that a naive `indicators=[…]` wrapper would silently get wrong. Use `csmar_search` → `csmar_query` (with `fields`/`conditions`) instead, which is explicit and correct.

### How it works
- Auth: `POST /api/csmar-main/automaticLogin` (self-generated `signCode` UUID + IP) → token.
- Search: `POST /api/csmar-main/highlight/searchData`.
- Catalog: `GET /api/csmar-main/single/getSeriesTree/-1`; tables: `GET .../getSingleTableLeftTree/{dbId}`.
- Fields: `GET /api/csmar-main/single/getSampleData/{tableId}`.
- Preview: `POST .../csmar-single/single/cacheCondition` → `GET .../preview/{id}` (selected fields need `isChecked=1`).
- Download: `saveOutline` → `singleData/pack` (async) → WebSocket `wss://data.csmar.com/ws`
  send `{outlineId,status:"start",token,signCode}`, receive `{filePath}` → fetch zip from
  `https://file.csmar.com/{filePath}` (contains the `.xlsx`). CSMAR xlsx has 3 header rows
  (name / description / unit).

---

## 中文

### 前提
- **本机处于已订阅 CSMAR 的机构 IP（校园网）环境**。鉴权靠 IP，无登录表单；换到没授权的网络，`csmar_status` 会显示 `IP授权: false`，所有功能不可用。
- Python 3.10+：
  ```bash
  pip install mcp requests websocket-client pandas openpyxl
  ```

### 安装（项目级）
1. 复制示例配置并按本机改路径：
   ```bash
   cp .mcp.json.example .mcp.json
   # 改 "command" 为你的 python，args 改成 server.py 的绝对路径
   ```
2. **重启 MCP 客户端**（如 Claude Code）即加载 `csmar` 服务。

环境变量（`.mcp.json` 的 `env`）：
- `CSMAR_LANG`：`1`=英文界面（字段名英文、搜索用英文词），`0`=中文界面。默认 `1`。
- `CSMAR_OUT_DIR`：批量下载默认目录。默认 `~/Downloads`。

### 工具（11 个）

**探索类**

| 工具 | 说明 | 关键参数 |
|---|---|---|
| `csmar_status` | 检查登录/机构/IP 授权状态 | 无 |
| `csmar_list_databases` | 列出数据库目录（含是否已订阅） | `only_accessible`, `keyword` |
| `csmar_list_tables` | 列出某数据库下的表 | `database_id` |
| `csmar_search` | 关键词搜索库/表/字段，拿到 `table_id` | `keyword`, `page_size` |
| `csmar_list_fields` | 列出某表的字段、时间/代码字段、样本 | `table`（table_id 或表名） |

**取数类**（`table` 可传 `table_id` 或表名）

| 工具 | 说明 | 关键参数 |
|---|---|---|
| `csmar_query_count` | 统计满足条件的记录总数 | `table`, `start_date`, `end_date`, `codes?` |
| `csmar_preview` | 取数预览，直接返回 JSON（≤200行） | `table`, `start_date`, `end_date`, `fields?`, `codes?` |
| `csmar_query` | 通用取数返回 JSON；`limit>200` 自动走后台打包下载再读取 | `table`, `start_date`, `end_date`, `fields?`, `codes?`, `limit` |
| `csmar_download` | 批量下载完整数据，存本地 xlsx/csv | `table`, `start_date`, `end_date`, `fields?`, `codes?`, `out_dir?`, `as_csv?` |

**便捷封装**

| 工具 | 说明 | 关键参数 |
|---|---|---|
| `get_stock_data` | 按股票代码取日/周/月/年行情（TRD_Dalyr 等） | `stock_code`, `start_date`, `end_date`, `frequency`, `limit` |
| `get_company_info` | 按股票代码取公司基本信息（TRD_Co） | `stock_code` |

日期格式 `YYYY-MM-DD`；`fields` 缺省=全部；`codes` 缺省=全部；`frequency` ∈ daily/weekly/monthly/annual。

四个取数工具都支持 **`conditions`**（字段筛选），例：`[{"field":"Clsprc","op":">","value":"100"}]`。
运算符：`>` `>=` `<` `<=` `=` `!=` `like` `not like` `is null` `not is null`；多条件默认 AND（单条加 `"relation":"or"` 可 OR）。

### 典型用法（对 Claude 说）
- “确认 CSMAR 能连吗” → `csmar_status`
- “搜股票交易相关的表” → `csmar_search("stock")`
- “日行情表有哪些字段” → `csmar_list_fields("Daily Stock Price & Returns")`
- “预览 2026-07 第一周的 PE、PB” → `csmar_preview(3673,"2026-07-01","2026-07-07",["TradingDate","Symbol","PE","PB"])`
- “这段完整下载成 Excel” → `csmar_download(3673,"2026-07-01","2026-07-07")`
- “000001 在 2024-12 的日行情” → `get_stock_data("000001","2024-12-01","2024-12-31")`

### 限制（CSMAR 侧）
- 预览每次最多 **200 行**（更多用 `csmar_query`/`csmar_download`）。
- 单次查询最多 **20 万条**；时间跨度上限因频率而异（如日频 ≤ 2 年）。超限就缩小范围。
- 相同查询短时间内可能被限流；偶发慢/超时已内置重试。
- 仅覆盖单表查询（Data Query）；跨表查询未实现。

**有意不做的两点**
- *冷启动用物理表名（如 `TRD_Dalyr`）解析表*：CSMAR 搜索索引不含物理名，也没有「按名查 id」的端点，所以请传 `table_id`（来自 `csmar_search`/`csmar_list_tables`）或表显示名；物理名仅在该表本会话内被访问过后才可用。
- *`get_financial_data` 便捷工具*：财务数据跨多张表，且有合并/母公司、报表类型等口径差异，`indicators=[…]` 式的封装很容易给出「看似对实则错」的数字。改用 `csmar_search` → `csmar_query`（配 `fields`/`conditions`）更明确也更可靠。

### 实现说明
参见上方 English → How it works。核心：IP 换 token → 搜索/浏览/预览（HTTP JSON）→ 下载走 `saveOutline`+`pack`（异步）+ WebSocket 拿文件路径 → 从 `file.csmar.com` 取 zip（内含 xlsx，前 3 行为 字段名/描述/单位）。

---

## Disclaimer
This project replicates the public web endpoints that your own institutional IP is already authorized to access. It does not bypass access control, crack authentication, or share accounts. Use it in accordance with CSMAR's Terms of Service and your institution's license.
