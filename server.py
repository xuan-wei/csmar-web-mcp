#!/usr/bin/env python3
"""
CSMAR Web-API MCP —— 无需账号密码，走校园网 IP 自动登录，纯 HTTP 复刻 data.csmar.com 网页接口。

前提：本机处于已订阅 CSMAR 的机构 IP（校园网）环境。服务器通过 IP 自动换取 token，
不绕过权限、不共享账号，仅复刻用户自己有权访问的网页查询接口。

工具：
  探索类
    csmar_status         检查登录/机构/IP 授权状态
    csmar_list_databases 列出数据库目录（按系列，含是否已订阅）
    csmar_list_tables    列出某数据库下的表
    csmar_search         关键词搜索表/字段
    csmar_list_fields    列出某表的字段、时间/代码字段、样本
  取数类（table 参数可传 table_id 数字，或表名（英文/中文，与 csmar_search 结果一致）
    csmar_query_count    统计满足条件的记录数
    csmar_preview        取数预览，直接返回 JSON（≤200 行）
    csmar_query          通用取数，返回 JSON 行（>200 行自动走后台打包下载再读取）
    csmar_download       批量下载完整数据，存本地 xlsx/csv
  便捷封装
    get_stock_data       按股票代码取日/周/月/年行情
    get_company_info     按股票代码取公司基本信息
"""
import os
import io
import copy
import json
import uuid
import time
import zipfile
import warnings
import requests
from websocket import create_connection, WebSocketTimeoutException
from mcp.server.fastmcp import FastMCP

warnings.filterwarnings("ignore", message="Workbook contains no default style")

# ==================== 配置 ====================
BASE = "https://data.csmar.com"
FILE_HOST = "https://file.csmar.com/"
WS_URL = "wss://data.csmar.com/ws"
LANG = os.environ.get("CSMAR_LANG", "1")          # 1=英文界面, 0=中文界面
DEFAULT_OUT = os.environ.get("CSMAR_OUT_DIR", os.path.expanduser("~/Downloads"))
UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/150.0.0.0 Safari/537.36")
FILEOUT_XLSX = "2"                                # Excel2007（*.xlsx）格式码
PREVIEW_CAP = 200                                 # 网页 preview 单次上限

# 便捷封装用的标准表（Stock Trading / Company 库，已实测）
STOCK_TABLES = {"daily": 3596, "weekly": 3597, "monthly": 3598, "annual": 3599}
COMPANY_TABLE = 3593                              # TRD_Co 公司基本信息

# 条件运算符 -> CSMAR character id（从网页前端逆向）
OP_MAP = {
    ">": "0", ">=": "1", "<": "2", "<=": "3", "=": "4", "==": "4",
    "!=": "5", "<>": "5", "like": "6", "not like": "7",
    "is null": "8", "isnull": "8", "not is null": "9", "notnull": "9", "is not null": "9",
}


def _clean(x):
    return (x or "").replace("<span class='font-red'>", "").replace("</span>", "")


class Csmar:
    """CSMAR 网页接口会话：持有 signCode + token，自动登录、失效重登。"""

    def __init__(self):
        self.s = requests.Session()
        self.sign = str(uuid.uuid4())
        self.s.cookies.set("signCode", self.sign, domain="data.csmar.com")
        self.token = None
        self.info = {}
        self._meta_cache = {}     # table_id -> getSampleData data
        self._db_cache = {}       # table_id -> {databaseId, databaseName, seriesName}
        self._name2id = {}        # 物理表名/表名(lower) -> table_id
        self._catalog = None      # getSeriesTree 缓存

    # ---------- 底层 ----------
    def _headers(self):
        h = {
            "signcode": self.sign, "lang": LANG, "belong": "0",
            "content-type": "application/json;charse=UTF-8",
            "accept": "application/json, text/plain, */*",
            "origin": BASE, "referer": BASE + "/csmar.html", "user-agent": UA,
        }
        if self.token:
            h["token"] = self.token
        return h

    def login(self, force=False):
        if self.token and not force:
            return self.info
        r = self.s.post(
            BASE + "/api/csmar-main/automaticLogin", headers=self._headers(),
            data=json.dumps({"resolvingPower": "1920X1080", "browser": "Chrome 150.0",
                             "website": "data.csmar.com", "clientType": "0"}),
            timeout=30)
        d = r.json()
        if d.get("code") != 0 or not d.get("data", {}).get("token"):
            raise RuntimeError(f"IP 自动登录失败: {d.get('msg')}（请确认处于校园网/机构 IP 环境）")
        self.info = d["data"]
        self.token = self.info["token"]
        return self.info

    @staticmethod
    def _is_auth_err(d):
        msg = (d.get("msg") or "")
        return d.get("code") in (401, -401, -100, -101) or "登录" in msg or "token" in msg.lower()

    def _req(self, method, path, host=BASE, body=None, timeout=60, retries=1, _reauth=True):
        """统一请求：校验 code==0；超时/瞬时错误自动重试；鉴权失效重登一次。"""
        self.login()
        url = host + path
        last = None
        for attempt in range(retries + 1):
            try:
                if method == "GET":
                    r = self.s.get(url, headers=self._headers(), timeout=timeout)
                else:
                    r = self.s.post(url, headers=self._headers(), data=json.dumps(body), timeout=timeout)
                d = r.json()
            except (requests.Timeout, requests.ConnectionError, ValueError) as e:
                last = f"网络/响应异常: {e}"
                time.sleep(2 * (attempt + 1))
                continue
            if d.get("code") != 0:
                if _reauth and self._is_auth_err(d):
                    self.login(force=True)
                    return self._req(method, path, host, body, timeout, retries, _reauth=False)
                raise RuntimeError(d.get("msg") or f"接口返回 code={d.get('code')}")
            return d.get("data")
        raise RuntimeError(last or "请求失败")

    # ---------- 目录/表解析 ----------
    def catalog(self):
        """getSeriesTree/-1：数据库目录（系列 + 数据库清单，含 useable）。带缓存。"""
        if self._catalog is None:
            self._catalog = self._req("GET", "/api/csmar-main/single/getSeriesTree/-1", timeout=30)
        return self._catalog

    def table_meta(self, table_id):
        """getSampleData：表名、物理名、字段元数据、样本。带缓存。"""
        if table_id not in self._meta_cache:
            self._meta_cache[table_id] = self._req(
                "GET", f"/api/csmar-main/single/getSampleData/{table_id}")
            # 顺手登记物理名/表名 -> id
            m = self._meta_cache[table_id]
            self._name2id[m["tableNamePhy"].lower()] = table_id
            self._name2id[m["tableName"].lower()] = table_id
        return self._meta_cache[table_id]

    def resolve_table(self, table):
        """table 可为 table_id(int/数字串) 或 物理表名/表名(str)，统一解析为 table_id。"""
        if isinstance(table, int):
            return table
        t = str(table).strip()
        if t.isdigit():
            return int(t)
        key = t.lower()
        if key in self._name2id:
            return self._name2id[key]
        # 搜索反查
        data = self._req("POST", "/api/csmar-main/highlight/searchData",
                         body={"searchType": "0", "searchKey": t, "pageNo": 1, "pageSize": 50})
        recs = (((data or {}).get("highlightEsFieldVo") or {}).get("esFieldPage") or {}).get("records") or []
        for r in recs:
            tid = r.get("tableId")
            if tid:
                self._db_cache.setdefault(tid, {
                    "databaseId": r.get("databaseId"),
                    "databaseName": _clean(r.get("databaseName")),
                    "seriesName": _clean(r.get("seriesName"))})
        # 优先按物理名精确匹配（需取 meta），否则取第一个表名匹配
        for r in recs:
            tid = r.get("tableId")
            if tid and _clean(r.get("tableName")).lower() == key:
                return tid
        for r in recs:
            if r.get("tableId"):
                try:
                    if self.table_meta(r["tableId"])["tableNamePhy"].lower() == key:
                        return r["tableId"]
                except Exception:                # noqa: BLE001
                    pass
        raise RuntimeError(f"无法解析表 '{table}'，请用 csmar_search 确认 table_id 或物理表名")

    def resolve_db(self, table_id):
        """按表名反查所属数据库信息（预览需要 databaseName）。带缓存。"""
        if table_id in self._db_cache and self._db_cache[table_id].get("databaseName"):
            return self._db_cache[table_id]
        meta = self.table_meta(table_id)
        data = self._req("POST", "/api/csmar-main/highlight/searchData",
                         body={"searchType": "0", "searchKey": meta["tableName"],
                               "pageNo": 1, "pageSize": 50})
        recs = (((data or {}).get("highlightEsFieldVo") or {}).get("esFieldPage") or {}).get("records") or []
        match = next((r for r in recs if r.get("tableId") == table_id), recs[0] if recs else {})
        info = {"databaseId": match.get("databaseId"),
                "databaseName": _clean(match.get("databaseName")),
                "seriesName": _clean(match.get("seriesName"))}
        self._db_cache[table_id] = info
        return info

    @staticmethod
    def _time_code_fields(fields):
        code_f = next((f["field"] for f in fields if f.get("fieldSign") == "2"), "")
        time_f = next((f["field"] for f in fields if f.get("fieldSign") == "1"), "")
        return time_f, code_f

    def _select_fields(self, fields, names):
        """挑出要查询的字段元数据（缺省=全部），强制 isChecked=1；返回副本不改缓存。"""
        if not names:
            sel = [copy.deepcopy(f) for f in fields]
        else:
            want = {n.lower() for n in names}
            sel = [copy.deepcopy(f) for f in fields
                   if f["field"].lower() in want or (f.get("fieldTitle") or "").lower() in want]
            if not sel:
                raise RuntimeError(f"未匹配到字段 {names}；可用字段见 csmar_list_fields")
        for f in sel:
            f["isChecked"] = "1"
        return sel

    def _build_conditions(self, meta, conditions):
        """把 [{field, op, value, relation?}] 转成 CSMAR 的 filterCondition DTO 列表。"""
        if not conditions:
            return []
        by_field = {f["field"].lower(): f for f in meta["fieldInfoVos"]}
        by_title = {(f.get("fieldTitle") or "").lower(): f for f in meta["fieldInfoVos"]}
        out = []
        for cnd in conditions:
            field = cnd.get("field") or cnd.get("column")
            op = str(cnd.get("op") or cnd.get("operator") or "=").strip().lower()
            ch = OP_MAP.get(op)
            if ch is None:
                raise RuntimeError(f"不支持的运算符 '{op}'，可用: > >= < <= = != like 'not like' 'is null' 'not is null'")
            fld = by_field.get(str(field).lower()) or by_title.get(str(field).lower())
            if not fld:
                raise RuntimeError(f"条件字段 '{field}' 不存在，见 csmar_list_fields")
            out.append({"field": fld["field"], "character": ch,
                        "condition": str(cnd.get("value", "")),
                        "name": fld.get("fieldTitle") or fld["field"],
                        "conditionRelation": str(cnd.get("relation") or "and").lower()})
        return out

    # ---------- 取数核心 ----------
    def cache_condition(self, table_id, start, end, fields, codes, conditions=None):
        """构造并提交 cacheCondition，返回 (cache_id, meta, selected_fields)。"""
        meta = self.table_meta(table_id)
        db = self.resolve_db(table_id)
        if not db.get("databaseName"):
            raise RuntimeError("无法解析该表所属数据库，请先用 csmar_search 定位该表")
        time_f, code_f = self._time_code_fields(meta["fieldInfoVos"])
        sel = self._select_fields(meta["fieldInfoVos"], fields)
        body = {
            "codeStr": ",".join(codes) if codes else "", "codeId": "",
            "fileOutType": FILEOUT_XLSX, "tableName": meta["tableName"],
            "tableNamePhy": meta["tableNamePhy"], "codeSetField": code_f, "timeSetField": time_f,
            "id": table_id, "startDate": start, "endDate": end,
            "fieldInfoDtos": sel, "filterConditionDtos": self._build_conditions(meta, conditions),
            "databaseName": db["databaseName"], "databaseId": db["databaseId"],
            "seriesName": db["seriesName"], "frequency": "0", "showCount": "1",
        }
        cache_id = self._req("POST", "/api/csmar-single/single/cacheCondition",
                             body=body, timeout=90, retries=2)
        return cache_id, meta, sel

    def preview_rows(self, table_id, start, end, fields=None, codes=None, conditions=None):
        """cacheCondition + preview，返回 (dataCount, rows(≤200), selected_fields, meta)。"""
        cache_id, meta, sel = self.cache_condition(table_id, start, end, fields, codes, conditions)
        pv = self._req("GET", f"/api/csmar-single/single/preview/{cache_id}", timeout=90, retries=2)
        rows = pv.get("previewDatas") or []
        for r in rows:
            r.pop("updateid", None)
        return pv.get("dataCount"), rows, sel, meta

    def download_file(self, table_id, start, end, fields=None, codes=None, conditions=None):
        """saveOutline → pack → WS → 下 zip 解出 xlsx，返回 (xlsx_bytes, dataCount, meta)。"""
        meta = self.table_meta(table_id)
        time_f, code_f = self._time_code_fields(meta["fieldInfoVos"])
        sel = self._select_fields(meta["fieldInfoVos"], fields)
        field_ids = ",".join(str(f["id"]) for f in sel)
        oid = self._req("POST", "/api/csmar-main/single/saveOutline", body={
            "codeId": "", "codeStr": ",".join(codes) if codes else "",
            "fieldStr": field_ids, "fileOutType": FILEOUT_XLSX,
            "tableName": meta["tableName"], "tableNamePhy": meta["tableNamePhy"],
            "planName": "", "mail": "", "codeSetField": code_f, "timeSetField": time_f,
            "conditionDtos": self._build_conditions(meta, conditions),
            "startTime": start, "endTime": end,
            "estimateTime": 1, "tableId": table_id, "downloadType": "0", "frequency": "0"})
        pack_id = self._req("POST", "/api/csmar-main/singleData/pack", body={
            "codeSetField": code_f, "mail": "", "outlineId": oid,
            "timeSetField": time_f, "planName": ""})
        res = self._retrieve(pack_id or oid)
        content = requests.get(FILE_HOST + res["filePath"], timeout=180).content
        zf = zipfile.ZipFile(io.BytesIO(content))
        xlsx_name = next((n for n in zf.namelist() if n.lower().endswith(".xlsx")), None)
        if not xlsx_name:
            raise RuntimeError(f"压缩包内未找到 xlsx：{zf.namelist()}")
        return zf.read(xlsx_name), res.get("dataCount"), meta

    def _retrieve(self, ws_id, timeout=300):
        deadline = time.time() + timeout
        last_err = None
        while time.time() < deadline:
            ws = None
            try:
                ws = create_connection(WS_URL, timeout=20,
                                       header=[f"User-Agent: {UA}", f"Origin: {BASE}"])
                ws.send(json.dumps({"outlineId": ws_id, "status": "start",
                                    "token": self.token, "signCode": self.sign}))
                while True:
                    m = json.loads(ws.recv())
                    if m.get("code") not in (0, None):
                        raise RuntimeError(m.get("msg") or "WS 返回错误")
                    data = m.get("data") or {}
                    if data.get("filePath"):
                        return data
            except WebSocketTimeoutException:
                last_err = "WS 读取超时，重试中"
            except Exception as e:                # noqa: BLE001
                last_err = str(e)
            finally:
                if ws:
                    try:
                        ws.close()
                    except Exception:            # noqa: BLE001
                        pass
            time.sleep(3)
        raise RuntimeError(f"下载打包超时：{last_err}")


csmar = Csmar()
mcp = FastMCP("csmar")


def _text_cols(meta):
    """CSMAR 文本类字段（代码/名称等），读 xlsx 时须按字符串处理以保留前导零。"""
    return [f["field"] for f in meta["fieldInfoVos"]
            if "char" in (f.get("fieldType") or "").lower()]


def _xlsx_to_rows(xlsx_bytes, limit, str_cols=None):
    """读 xlsx 首个 sheet 为 JSON 行（≤limit）。CSMAR 前 3 行为 字段名/描述/单位，跳过第 2、3 行；
    文本类列按字符串读，避免股票代码丢前导零。"""
    import pandas as pd
    dtype = {c: str for c in (str_cols or [])}
    df = pd.read_excel(io.BytesIO(xlsx_bytes), skiprows=[1, 2], dtype=dtype)
    if limit:
        df = df.head(limit)
    return json.loads(df.to_json(orient="records", date_format="iso"))


# ==================== 探索类工具 ====================
@mcp.tool()
def csmar_status() -> dict:
    """检查 CSMAR 登录与授权状态。确认当前是否处于已订阅机构的校园网 IP 环境。"""
    info = csmar.login(force=True)
    return {
        "机构": info.get("cousManageName") or info.get("school"),
        "学校": info.get("school"),
        "IP": info.get("clientIp"),
        "IP授权": info.get("isIPRange"),
        "有效期至": info.get("expireTime"),
        "界面语言": "英文" if LANG == "1" else "中文",
    }


@mcp.tool()
def csmar_list_databases(only_accessible: bool = True, keyword: str = "") -> dict:
    """
    列出 CSMAR 数据库目录（Data Query 页的库清单）。
    only_accessible=True 只列本机构已订阅（useable）的库；keyword 可按库名过滤。
    返回的 database_id 可传给 csmar_list_tables。
    """
    cat = csmar.catalog()
    series = {s["id"]: s["name"] for s in cat.get("seriesTrees", [])}
    dbs = cat.get("typeDatabaseVo", {}).get("databaseList", [])
    out = []
    for d in dbs:
        if only_accessible and not d.get("useable"):
            continue
        if keyword and keyword.lower() not in (d.get("databaseName") or "").lower():
            continue
        out.append({"database_id": d["id"], "database": d["databaseName"],
                    "series": series.get(str(d.get("seriesId")), ""),
                    "accessible": d.get("useable")})
    out.sort(key=lambda x: (x["series"], x["database"]))
    return {"count": len(out),
            "note": "仅列 Data Query 目录中的库；如未见到目标库，用 csmar_search 直接按关键词找表",
            "databases": out}


@mcp.tool()
def csmar_list_tables(database_id: int) -> dict:
    """列出某数据库下的全部表（含 table_id）。database_id 来自 csmar_list_databases。"""
    d = csmar._req("GET", f"/api/csmar-main/single/getSingleTableLeftTree/{database_id}", timeout=30)
    tables = []

    def walk(nodes):
        for n in nodes or []:
            tid = n.get("tableId")
            if not tid and isinstance(n.get("id"), str) and n["id"].startswith("table-"):
                tid = int(n["id"].split("-")[1])
            if tid:
                tables.append({"table_id": tid, "table": n.get("name")})
            walk(n.get("children"))
            walk(n.get("tableTrees"))
    walk(d.get("tableTrees"))
    return {"database": d.get("databaseName"), "series": d.get("seriesName"),
            "count": len(tables), "tables": tables}


@mcp.tool()
def csmar_search(keyword: str, page_size: int = 20) -> dict:
    """关键词搜索 CSMAR（英文界面用英文词）。返回匹配的表（含 table_id）和字段。"""
    data = csmar._req("POST", "/api/csmar-main/highlight/searchData",
                      body={"searchType": "0", "searchKey": keyword, "pageNo": 1, "pageSize": page_size})
    page = (((data or {}).get("highlightEsFieldVo") or {}).get("esFieldPage") or {})
    recs = page.get("records") or []
    tables, seen = [], set()
    for r in recs:
        tid = r.get("tableId")
        if tid and tid not in seen:
            seen.add(tid)
            csmar._db_cache.setdefault(tid, {
                "databaseId": r.get("databaseId"),
                "databaseName": _clean(r.get("databaseName")),
                "seriesName": _clean(r.get("seriesName"))})
            tables.append({"table_id": tid, "table": _clean(r.get("tableName")),
                           "database": _clean(r.get("databaseName")),
                           "series": _clean(r.get("seriesName")), "useable": r.get("useable")})
    fields = [{"table_id": r.get("tableId"), "table": _clean(r.get("tableName")),
               "field": _clean(r.get("field")), "title": _clean(r.get("fieldTitle")),
               "explain": _clean(r.get("fieldExplain"))} for r in recs]
    return {"total_fields_matched": page.get("total"), "tables": tables, "fields": fields[:page_size]}


@mcp.tool()
def csmar_list_fields(table: str) -> dict:
    """列出某表的全部字段、时间/代码字段与样本。table 可传 table_id 或表名。"""
    tid = csmar.resolve_table(table)
    meta = csmar.table_meta(tid)
    time_f, code_f = csmar._time_code_fields(meta["fieldInfoVos"])
    db = csmar.resolve_db(tid)
    return {
        "table_id": tid, "table": meta["tableName"], "table_phy": meta["tableNamePhy"],
        "database": db.get("databaseName"), "time_field": time_f, "code_field": code_f,
        "start_time": meta.get("startTime"), "end_time": meta.get("endTime"),
        "fields": [{"field": f["field"], "title": f.get("fieldTitle"),
                    "type": f.get("fieldType"), "explain": f.get("fieldExplain")}
                   for f in meta["fieldInfoVos"]],
        "sample": (meta.get("sampleList") or [])[:3],
    }


# ==================== 取数类工具 ====================
@mcp.tool()
def csmar_query_count(table: str, start_date: str, end_date: str,
                      codes: list[str] | None = None,
                      conditions: list[dict] | None = None) -> dict:
    """统计满足条件的记录总数（不取数据）。table 可传 table_id 或表名。
    conditions 例：[{"field":"Clsprc","op":">","value":"100"}]，op 支持 > >= < <= = != like。"""
    tid = csmar.resolve_table(table)
    cache_id, meta, _ = csmar.cache_condition(tid, start_date, end_date, None, codes, conditions)
    pv = csmar._req("GET", f"/api/csmar-single/single/preview/{cache_id}", timeout=90, retries=2)
    return {"table": meta["tableName"], "total": pv.get("dataCount")}


@mcp.tool()
def csmar_preview(table: str, start_date: str, end_date: str,
                  fields: list[str] | None = None, codes: list[str] | None = None,
                  conditions: list[dict] | None = None) -> dict:
    """
    取数预览：直接返回 JSON 行（每次≤200 行）。table 可传 table_id 或表名。
    日期 YYYY-MM-DD；fields 缺省=全部；codes 缺省=全部代码。
    conditions 例：[{"field":"Clsprc","op":">","value":"100","relation":"and"}]。
    """
    tid = csmar.resolve_table(table)
    total, rows, sel, meta = csmar.preview_rows(tid, start_date, end_date, fields, codes, conditions)
    return {"table": meta["tableName"], "total": total, "returned": len(rows),
            "note": "预览上限 200 行；要更多用 csmar_query 或 csmar_download",
            "columns": [{"field": f["field"], "title": f.get("fieldTitle")} for f in sel],
            "rows": rows}


@mcp.tool()
def csmar_query(table: str, start_date: str, end_date: str,
                fields: list[str] | None = None, codes: list[str] | None = None,
                conditions: list[dict] | None = None, limit: int = 200) -> dict:
    """
    通用取数，返回 JSON 行。limit≤200 走快速预览；limit>200 自动后台打包下载再读取（较慢）。
    table 可传 table_id 或表名。conditions 例：[{"field":"Clsprc","op":">=","value":"100"}]。
    注意 CSMAR 单次≤20万条、时间跨度上限因表频率而异。
    """
    tid = csmar.resolve_table(table)
    if limit <= PREVIEW_CAP:
        total, rows, sel, meta = csmar.preview_rows(tid, start_date, end_date, fields, codes, conditions)
        rows = rows[:limit]
        return {"table": meta["tableName"], "total": total, "returned": len(rows),
                "via": "preview",
                "columns": [{"field": f["field"], "title": f.get("fieldTitle")} for f in sel],
                "rows": rows}
    xlsx_bytes, total, meta = csmar.download_file(tid, start_date, end_date, fields, codes, conditions)
    rows = _xlsx_to_rows(xlsx_bytes, limit, _text_cols(meta))
    return {"table": meta["tableName"], "total": total, "returned": len(rows),
            "via": "download", "rows": rows}


@mcp.tool()
def csmar_download(table: str, start_date: str, end_date: str,
                   fields: list[str] | None = None, codes: list[str] | None = None,
                   conditions: list[dict] | None = None,
                   out_dir: str | None = None, as_csv: bool = False) -> dict:
    """
    批量下载完整数据，打包后存本地。默认 xlsx；as_csv=True 另存 CSV。返回本地路径与记录数。
    table 可传 table_id 或表名。conditions 例：[{"field":"Clsprc","op":">","value":"100"}]。
    注意 CSMAR 单次≤20万条、时间跨度上限因表频率而异。
    """
    tid = csmar.resolve_table(table)
    out_dir = out_dir or DEFAULT_OUT
    os.makedirs(out_dir, exist_ok=True)
    xlsx_bytes, total, meta = csmar.download_file(tid, start_date, end_date, fields, codes, conditions)
    base = f"{meta['tableNamePhy']}_{start_date}_{end_date}".replace(":", "")
    xlsx_path = os.path.join(out_dir, base + ".xlsx")
    with open(xlsx_path, "wb") as fh:
        fh.write(xlsx_bytes)
    result = {"table": meta["tableName"], "records": total, "xlsx": xlsx_path}
    if as_csv:
        try:
            import pandas as pd
            csv_path = os.path.join(out_dir, base + ".csv")
            pd.read_excel(xlsx_path, skiprows=[1, 2], dtype={c: str for c in _text_cols(meta)}
                          ).to_csv(csv_path, index=False, encoding="utf-8-sig")
            result["csv"] = csv_path
        except Exception as e:                    # noqa: BLE001
            result["csv_error"] = f"转 CSV 失败: {e}"
    return result


# ==================== 便捷封装 ====================
@mcp.tool()
def get_stock_data(stock_code: str, start_date: str, end_date: str,
                   frequency: str = "daily", limit: int = 200) -> dict:
    """
    取某只股票的行情（Stock Trading 库）。frequency: daily/weekly/monthly/annual。
    返回开高低收、成交量额、收益率等。日频单次时间跨度≤2年。
    """
    tid = STOCK_TABLES.get(frequency)
    if not tid:
        raise RuntimeError(f"frequency 需为 {list(STOCK_TABLES)}")
    return csmar_query(str(tid), start_date, end_date, codes=[stock_code], limit=limit)


@mcp.tool()
def get_company_info(stock_code: str) -> dict:
    """取某只股票的公司基本信息（名称、行业、省份、上市日期等，TRD_Co 表）。"""
    total, rows, sel, meta = csmar.preview_rows(
        COMPANY_TABLE, "1990-01-01", "2099-12-31", None, [stock_code])
    return {"table": meta["tableName"], "stock_code": stock_code,
            "count": len(rows), "info": rows}


if __name__ == "__main__":
    mcp.run()
