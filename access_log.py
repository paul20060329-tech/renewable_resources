# -*- coding: utf-8 -*-
"""访问审计日志模块（独立于业务库）。

职责：
1. 在 Flask 中间件中捕获每次请求，解析 ``X-Forwarded-For`` 得到真实访客 IP；
2. 记录访问时间、访客 IP、登录用户/角色、请求方法、请求地址、状态码、User-Agent；
3. 写入**独立数据库**：
   - 若配置了环境变量 ``ACCESS_LOG_DATABASE_URL``（或 ``DATABASE_URL``），
     使用 PostgreSQL（Neon）持久化，Render 重启不丢数据；
   - 否则退化为项目根目录下的本地 SQLite（``access_log.db``），
     方便本地开发，不影响业务库 ``greenloop.db``。

设计要点：
- 写入采用后台线程 + 内存队列缓冲，避免每请求同步落库拖慢响应；
- 连接失败自动重试，Neon 免费实例冷启动时不会丢日志；
- 队列满时丢弃最旧记录，保证主站不被拖垮；
- 该模块只被 ``app.py`` 引用，不反向依赖业务代码。
"""

from __future__ import annotations

import atexit
import ipaddress
import os
import queue
import sqlite3
import threading
import time
from datetime import datetime, timedelta, timezone

# --------------------------------------------------------------------------- #
# 配置
# --------------------------------------------------------------------------- #

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 独立数据库连接串：优先专用变量，其次复用通用变量；都没有则用本地 SQLite。
ACCESS_LOG_DATABASE_URL = (
    os.environ.get("ACCESS_LOG_DATABASE_URL")
    or os.environ.get("DATABASE_URL")
    or ""
).strip()

SQLITE_PATH = os.environ.get("ACCESS_LOG_SQLITE_PATH", os.path.join(BASE_DIR, "access_log.db"))

# 队列上限与批量写入参数
QUEUE_MAXSIZE = 20000
BATCH_SIZE = 50
FLUSH_INTERVAL = 2.0          # 秒：即使不满一批也定期落库
CONNECT_RETRY = 3             # 单次批量写入的连接重试次数
CONNECT_RETRY_WAIT = 1.5      # 秒

WRITE_QUEUE: "queue.Queue[dict]" = queue.Queue(maxsize=QUEUE_MAXSIZE)
_STOP_EVENT = threading.Event()
_WORKER_LOCK = threading.Lock()
_WORKER: threading.Thread | None = None
_PG_POOL_LOCK = threading.Lock()
_PG_CONN = None

# 仅记录这些前缀之外的请求（静态资源不记，避免噪声）
SKIP_PREFIXES = (
    "/static/",
    "/system_admin/static/",
    "/favicon.ico",
)
# 后台自己的访问也记录，但可以按需排除轮询类接口
SKIP_EXACT = set()

_PG_AVAILABLE = None  # 延迟探测结果


# --------------------------------------------------------------------------- #
# 真实 IP 解析
# --------------------------------------------------------------------------- #

def client_ip(req) -> str:
    """从请求中解析真实访客 IP。

    Render / Nginx 等反向代理会把真实 IP 放在 ``X-Forwarded-For`` 的最左侧。
    这里按以下优先级取值：
    1. ``X-Forwarded-For`` 第一个非私有、非空地址；
    2. ``X-Real-IP``；
    3. ``CF-Connecting-IP``（若走 Cloudflare）；
    4. ``request.remote_addr``。
    """
    candidates: list[str] = []

    xff = req.headers.get("X-Forwarded-For", "")
    if xff:
        candidates.extend(part.strip() for part in xff.split(","))

    for header in ("X-Real-IP", "CF-Connecting-IP", "True-Client-IP"):
        value = req.headers.get(header)
        if value:
            candidates.append(value.strip())

    remote = req.remote_addr
    if remote:
        candidates.append(remote.strip())

    public_ip = ""
    fallback = ""
    for raw in candidates:
        if not raw:
            continue
        cleaned = raw.strip("[]")
        # 去掉端口（形如 1.2.3.4:5678）
        if cleaned.count(":") == 1 and "." in cleaned:
            cleaned = cleaned.split(":", 1)[0]
        if not fallback:
            fallback = cleaned
        try:
            addr = ipaddress.ip_address(cleaned)
        except ValueError:
            continue
        if addr.is_global:
            public_ip = cleaned
            break

    return public_ip or fallback or "unknown"


def _is_local(ip: str) -> bool:
    try:
        addr = ipaddress.ip_address(ip)
    except ValueError:
        return False
    return addr.is_private or addr.is_loopback


# --------------------------------------------------------------------------- #
# 数据库访问层
# --------------------------------------------------------------------------- #

def _use_postgres() -> bool:
    """判断是否使用 PostgreSQL。

    只有当连接串可用且 psycopg 安装成功时才走 PG，否则安全退化为 SQLite。
    """
    global _PG_AVAILABLE
    if not ACCESS_LOG_DATABASE_URL:
        return False
    if _PG_AVAILABLE is not None:
        return _PG_AVAILABLE
    try:
        import psycopg  # noqa: F401
        _PG_AVAILABLE = True
    except Exception:
        _PG_AVAILABLE = False
    return _PG_AVAILABLE


def _normalize_pg_url(url: str) -> str:
    """把 SQLAlchemy 风格的 ``postgres://`` 转成 psycopg 可识别的 ``postgresql://``。"""
    if url.startswith("postgres://"):
        return "postgresql://" + url[len("postgres://"):]
    return url


def backend_name() -> str:
    if _use_postgres():
        return "PostgreSQL (Neon)"
    return "SQLite (本地文件)"


CREATE_TABLE_SQL_PG = """
CREATE TABLE IF NOT EXISTS access_logs (
    id           BIGSERIAL PRIMARY KEY,
    visited_at   TIMESTAMPTZ NOT NULL,
    ip           TEXT        NOT NULL,
    username     TEXT,
    role         TEXT,
    method       TEXT,
    path         TEXT,
    full_url     TEXT,
    status_code  INTEGER,
    user_agent   TEXT,
    referer      TEXT,
    is_login     BOOLEAN     DEFAULT FALSE,
    note         TEXT
);
"""

CREATE_TABLE_SQL_SQLITE = """
CREATE TABLE IF NOT EXISTS access_logs (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    visited_at   TEXT    NOT NULL,
    ip           TEXT    NOT NULL,
    username     TEXT,
    role         TEXT,
    method       TEXT,
    path         TEXT,
    full_url     TEXT,
    status_code  INTEGER,
    user_agent   TEXT,
    referer      TEXT,
    is_login     INTEGER DEFAULT 0,
    note         TEXT
);
"""

INDEX_SQL_PG = [
    "CREATE INDEX IF NOT EXISTS idx_access_logs_visited_at ON access_logs (visited_at DESC);",
    "CREATE INDEX IF NOT EXISTS idx_access_logs_ip ON access_logs (ip);",
    "CREATE INDEX IF NOT EXISTS idx_access_logs_username ON access_logs (username);",
]
INDEX_SQL_SQLITE = [
    "CREATE INDEX IF NOT EXISTS idx_access_logs_visited_at ON access_logs (visited_at DESC);",
    "CREATE INDEX IF NOT EXISTS idx_access_logs_ip ON access_logs (ip);",
    "CREATE INDEX IF NOT EXISTS idx_access_logs_username ON access_logs (username);",
]

INSERT_SQL_PG = """
INSERT INTO access_logs
    (visited_at, ip, username, role, method, path, full_url,
     status_code, user_agent, referer, is_login, note)
VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
"""

INSERT_SQL_SQLITE = """
INSERT INTO access_logs
    (visited_at, ip, username, role, method, path, full_url,
     status_code, user_agent, referer, is_login, note)
VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
"""


def _pg_connect():
    """获取（并复用）PostgreSQL 连接，带重试。"""
    global _PG_CONN
    import psycopg

    with _PG_POOL_LOCK:
        if _PG_CONN is not None:
            try:
                if not _PG_CONN.closed:
                    return _PG_CONN
            except Exception:
                _PG_CONN = None
        last_error = None
        for attempt in range(CONNECT_RETRY):
            try:
                conn = psycopg.connect(
                    _normalize_pg_url(ACCESS_LOG_DATABASE_URL),
                    autocommit=True,
                    connect_timeout=10,
                    prepare_threshold=None,  # 兼容 Neon pooler 的连接复用
                )
                _PG_CONN = conn
                return conn
            except Exception as exc:  # noqa: BLE001
                last_error = exc
                time.sleep(CONNECT_RETRY_WAIT * (attempt + 1))
        raise RuntimeError(f"无法连接访问日志数据库: {last_error}")


def _reset_pg_connection() -> None:
    global _PG_CONN
    with _PG_POOL_LOCK:
        if _PG_CONN is not None:
            try:
                _PG_CONN.close()
            except Exception:
                pass
        _PG_CONN = None


def init_access_db() -> str:
    """建表 + 建索引。返回实际使用的后端名称。"""
    if _use_postgres():
        conn = _pg_connect()
        with conn.cursor() as cur:
            cur.execute(CREATE_TABLE_SQL_PG)
            for stmt in INDEX_SQL_PG:
                cur.execute(stmt)
        return backend_name()

    con = sqlite3.connect(SQLITE_PATH)
    try:
        con.execute(CREATE_TABLE_SQL_SQLITE)
        for stmt in INDEX_SQL_SQLITE:
            con.execute(stmt)
        con.commit()
    finally:
        con.close()
    return backend_name()


def _write_batch(rows: list[dict]) -> None:
    if not rows:
        return
    if _use_postgres():
        conn = _pg_connect()
        try:
            with conn.cursor() as cur:
                cur.executemany(
                    INSERT_SQL_PG,
                    [
                        (
                            r["visited_at"], r["ip"], r["username"], r["role"],
                            r["method"], r["path"], r["full_url"], r["status_code"],
                            r["user_agent"], r["referer"], r["is_login"], r["note"],
                        )
                        for r in rows
                    ],
                )
            return
        except Exception:
            _reset_pg_connection()
            raise

    con = sqlite3.connect(SQLITE_PATH, timeout=10)
    try:
        con.executemany(
            INSERT_SQL_SQLITE,
            [
                (
                    r["visited_at"], r["ip"], r["username"], r["role"],
                    r["method"], r["path"], r["full_url"], r["status_code"],
                    r["user_agent"], r["referer"], 1 if r["is_login"] else 0, r["note"],
                )
                for r in rows
            ],
        )
        con.commit()
    finally:
        con.close()


def _worker() -> None:
    """后台写入线程：批量 + 定期落库，失败重试。"""
    buffer: list[dict] = []
    last_flush = time.time()
    while not _STOP_EVENT.is_set():
        timeout = max(0.1, FLUSH_INTERVAL - (time.time() - last_flush))
        try:
            item = WRITE_QUEUE.get(timeout=timeout)
            buffer.append(item)
        except queue.Empty:
            pass

        should_flush = len(buffer) >= BATCH_SIZE or (
            buffer and (time.time() - last_flush) >= FLUSH_INTERVAL
        )
        if not should_flush:
            continue

        pending, buffer = buffer, []
        last_flush = time.time()
        for attempt in range(CONNECT_RETRY):
            try:
                _write_batch(pending)
                break
            except Exception:
                if attempt == CONNECT_RETRY - 1:
                    # 最终失败：不阻塞主站，仅丢弃该批（避免内存无限增长）
                    break
                time.sleep(CONNECT_RETRY_WAIT * (attempt + 1))

    # 退出前尽力冲刷
    if buffer:
        try:
            _write_batch(buffer)
        except Exception:
            pass


def start_worker() -> None:
    global _WORKER
    with _WORKER_LOCK:
        if _WORKER is not None and _WORKER.is_alive():
            return
        _STOP_EVENT.clear()
        _WORKER = threading.Thread(target=_worker, name="access-log-writer", daemon=True)
        _WORKER.start()


def stop_worker() -> None:
    _STOP_EVENT.set()
    if _WORKER is not None:
        _WORKER.join(timeout=5)


atexit.register(stop_worker)


# --------------------------------------------------------------------------- #
# 查询接口（供后台审计台使用）
# --------------------------------------------------------------------------- #

def _row_to_dict(row, pg: bool) -> dict:
    if pg:
        keys = ["id", "visited_at", "ip", "username", "role", "method", "path",
                "full_url", "status_code", "user_agent", "referer", "is_login", "note"]
        data = dict(zip(keys, row))
        ts = data.get("visited_at")
        if isinstance(ts, datetime):
            data["visited_at"] = ts.astimezone(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
        data["location"] = describe_ip(data.get("ip", ""))
        return data
    data = {
        "id": row[0], "visited_at": row[1], "ip": row[2], "username": row[3],
        "role": row[4], "method": row[5], "path": row[6], "full_url": row[7],
        "status_code": row[8], "user_agent": row[9], "referer": row[10],
        "is_login": bool(row[11]), "note": row[12],
    }
    data["location"] = describe_ip(data.get("ip", ""))
    return data


def _build_filters(ip: str = "", username: str = "", keyword: str = "",
                   only_login: bool = False, date_from: str = "", date_to: str = "") -> tuple[str, list]:
    """构造 WHERE 子句（占位符按 PG 使用 %s / SQLite 使用 ?，由调用方替换）。"""
    clauses, params = [], []
    if ip:
        clauses.append("ip = %s")
        params.append(ip)
    if username:
        clauses.append("username = %s")
        params.append(username)
    if keyword:
        clauses.append("(path LIKE %s OR full_url LIKE %s OR user_agent LIKE %s)")
        like = f"%{keyword}%"
        params.extend([like, like, like])
    if only_login:
        clauses.append("is_login = TRUE")
    if date_from:
        clauses.append("visited_at >= %s")
        params.append(date_from)
    if date_to:
        clauses.append("visited_at <= %s")
        params.append(date_to)
    where = (" WHERE " + " AND ".join(clauses)) if clauses else ""
    return where, params


def query_logs(page: int = 1, page_size: int = 50, **filters) -> dict:
    """分页查询访问日志。返回 {total, page, page_size, rows}。"""
    page = max(1, int(page or 1))
    page_size = max(1, min(200, int(page_size or 50)))
    offset = (page - 1) * page_size
    where, params = _build_filters(**filters)

    if _use_postgres():
        conn = _pg_connect()
        with conn.cursor() as cur:
            count_sql = f"SELECT COUNT(*) FROM access_logs{where}"
            cur.execute(count_sql, params)
            total = cur.fetchone()[0]
            sql = (
                "SELECT id, visited_at, ip, username, role, method, path, full_url, "
                "status_code, user_agent, referer, is_login, note "
                f"FROM access_logs{where} ORDER BY visited_at DESC, id DESC LIMIT %s OFFSET %s"
            )
            cur.execute(sql, params + [page_size, offset])
            rows = [_row_to_dict(r, True) for r in cur.fetchall()]
        return {"total": total, "page": page, "page_size": page_size, "rows": rows}

    con = sqlite3.connect(SQLITE_PATH, timeout=10)
    try:
        sqlite_where = where.replace("%s", "?")
        sqlite_params = list(params)
        if filters.get("only_login"):
            # SQLite 用 1 表示 TRUE
            sqlite_where = sqlite_where.replace("is_login = TRUE", "is_login = 1")
        cur = con.execute(f"SELECT COUNT(*) FROM access_logs{sqlite_where}", sqlite_params)
        total = cur.fetchone()[0]
        sql = (
            "SELECT id, visited_at, ip, username, role, method, path, full_url, "
            "status_code, user_agent, referer, is_login, note "
            f"FROM access_logs{sqlite_where} ORDER BY visited_at DESC, id DESC LIMIT ? OFFSET ?"
        )
        cur = con.execute(sql, sqlite_params + [page_size, offset])
        rows = [_row_to_dict(r, False) for r in cur.fetchall()]
    finally:
        con.close()
    return {"total": total, "page": page, "page_size": page_size, "rows": rows}


def query_stats(days: int = 7) -> dict:
    """汇总统计：总量、独立 IP、独立用户、今日访问、最近 N 天趋势、Top IP / Top 用户。"""
    days = max(1, min(90, int(days or 7)))

    if _use_postgres():
        conn = _pg_connect()
        with conn.cursor() as cur:
            cur.execute("SELECT COUNT(*) FROM access_logs")
            total = cur.fetchone()[0]
            cur.execute("SELECT COUNT(DISTINCT ip) FROM access_logs")
            unique_ip = cur.fetchone()[0]
            cur.execute("SELECT COUNT(DISTINCT username) FROM access_logs WHERE username IS NOT NULL AND username <> ''")
            unique_user = cur.fetchone()[0]
            cur.execute("SELECT COUNT(*) FROM access_logs WHERE visited_at::date = CURRENT_DATE")
            today = cur.fetchone()[0]
            cur.execute(
                "SELECT to_char(visited_at::date, 'YYYY-MM-DD') AS d, COUNT(*) "
                "FROM access_logs WHERE visited_at >= CURRENT_DATE - INTERVAL '%s days' "
                "GROUP BY d ORDER BY d" % (days - 1)
            )
            trend = [{"date": r[0], "count": r[1]} for r in cur.fetchall()]
            cur.execute(
                "SELECT ip, COUNT(*) c FROM access_logs GROUP BY ip ORDER BY c DESC LIMIT 10"
            )
            top_ips = [{"ip": r[0], "count": r[1], "location": describe_ip(r[0])} for r in cur.fetchall()]
            cur.execute(
                "SELECT COALESCE(NULLIF(username,''),'(未登录)') u, COUNT(*) c FROM access_logs "
                "GROUP BY u ORDER BY c DESC LIMIT 10"
            )
            top_users = [{"username": r[0], "count": r[1]} for r in cur.fetchall()]
            logins = _recent_logins_pg(cur, 10)
        return {
            "total": total, "unique_ip": unique_ip, "unique_user": unique_user,
            "today": today, "trend": trend, "top_ips": top_ips,
            "top_users": top_users, "recent_logins": logins,
        }

    con = sqlite3.connect(SQLITE_PATH, timeout=10)
    try:
        total = con.execute("SELECT COUNT(*) FROM access_logs").fetchone()[0]
        unique_ip = con.execute("SELECT COUNT(DISTINCT ip) FROM access_logs").fetchone()[0]
        unique_user = con.execute(
            "SELECT COUNT(DISTINCT username) FROM access_logs WHERE username IS NOT NULL AND username <> ''"
        ).fetchone()[0]
        today = con.execute(
            "SELECT COUNT(*) FROM access_logs WHERE date(visited_at) = date('now')"
        ).fetchone()[0]
        trend = [
            {"date": r[0], "count": r[1]}
            for r in con.execute(
                "SELECT date(visited_at) d, COUNT(*) FROM access_logs "
                "WHERE date(visited_at) >= date('now', ?) GROUP BY d ORDER BY d",
                (f"-{days - 1} days",),
            ).fetchall()
        ]
        top_ips = [
            {"ip": r[0], "count": r[1], "location": describe_ip(r[0])}
            for r in con.execute(
                "SELECT ip, COUNT(*) c FROM access_logs GROUP BY ip ORDER BY c DESC LIMIT 10"
            ).fetchall()
        ]
        top_users = [
            {"username": r[0], "count": r[1]}
            for r in con.execute(
                "SELECT COALESCE(NULLIF(username,''),'(未登录)') u, COUNT(*) c "
                "FROM access_logs GROUP BY u ORDER BY c DESC LIMIT 10"
            ).fetchall()
        ]
        logins = [
            {"visited_at": r[0], "ip": r[1], "username": r[2], "role": r[3],
             "location": describe_ip(r[1])}
            for r in con.execute(
                "SELECT visited_at, ip, username, role FROM access_logs "
                "WHERE is_login = 1 ORDER BY visited_at DESC LIMIT 10"
            ).fetchall()
        ]
    finally:
        con.close()
    return {
        "total": total, "unique_ip": unique_ip, "unique_user": unique_user,
        "today": today, "trend": trend, "top_ips": top_ips,
        "top_users": top_users, "recent_logins": logins,
    }


def _recent_logins_pg(cur, limit: int = 10) -> list:
    cur.execute(
        "SELECT visited_at, ip, username, role FROM access_logs "
        "WHERE is_login = TRUE ORDER BY visited_at DESC LIMIT %s",
        (limit,),
    )
    out = []
    for r in cur.fetchall():
        ts = r[0]
        if isinstance(ts, datetime):
            ts = ts.astimezone(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
        out.append({"visited_at": ts, "ip": r[1], "username": r[2],
                    "role": r[3], "location": describe_ip(r[1])})
    return out


def distinct_values(column: str, limit: int = 200) -> list[str]:
    """取某列的去重值，用于后台筛选下拉。"""
    if column not in ("ip", "username"):
        return []
    if _use_postgres():
        conn = _pg_connect()
        with conn.cursor() as cur:
            cur.execute(
                f"SELECT {column}, COUNT(*) c FROM access_logs "
                f"WHERE {column} IS NOT NULL AND {column} <> '' "
                f"GROUP BY {column} ORDER BY c DESC LIMIT %s",
                (limit,),
            )
            return [r[0] for r in cur.fetchall()]
    con = sqlite3.connect(SQLITE_PATH, timeout=10)
    try:
        return [
            r[0]
            for r in con.execute(
                f"SELECT {column} FROM access_logs "
                f"WHERE {column} IS NOT NULL AND {column} <> '' "
                f"GROUP BY {column} ORDER BY COUNT(*) DESC LIMIT ?",
                (limit,),
            ).fetchall()
        ]
    finally:
        con.close()


def clear_logs(older_than_days: int | None = None) -> int:
    """清理日志；给定天数则只删更早的记录。返回删除条数。"""
    if _use_postgres():
        conn = _pg_connect()
        with conn.cursor() as cur:
            if older_than_days:
                cur.execute(
                    "DELETE FROM access_logs WHERE visited_at < NOW() - INTERVAL '%s days'"
                    % int(older_than_days)
                )
            else:
                cur.execute("DELETE FROM access_logs")
            return cur.rowcount or 0
    con = sqlite3.connect(SQLITE_PATH, timeout=10)
    try:
        if older_than_days:
            cutoff = (datetime.now() - timedelta(days=int(older_than_days))).strftime("%Y-%m-%d %H:%M:%S")
            cur = con.execute("DELETE FROM access_logs WHERE visited_at < ?", (cutoff,))
        else:
            cur = con.execute("DELETE FROM access_logs")
        con.commit()
        return cur.rowcount or 0
    finally:
        con.close()


# --------------------------------------------------------------------------- #
# IP 归属地（离线近似，不外呼，避免请求变慢）
# --------------------------------------------------------------------------- #

def describe_ip(ip: str) -> str:
    """给出 IP 的粗粒度归类。

    说明：完整的 IP 地理位置需要付费 API 或离线库（如 MaxMind）。
    这里只做**不产生网络请求**的本地判断，避免拖慢主站：
    - Render 内网 / 私有地址 → 代理或内网
    - 回环地址 → 本机
    - 其余 → 公网 IP（归属地需接入外部库后补充）
    后台界面会原样展示 IP，你可按 IP 自行核查归属地。
    """
    if not ip or ip == "unknown":
        return "未知"
    try:
        addr = ipaddress.ip_address(ip)
    except ValueError:
        return "未知"
    if addr.is_loopback:
        return "本机"
    if addr.is_private:
        return "内网/代理"
    if addr.is_global:
        return "公网"
    return "保留地址"


# --------------------------------------------------------------------------- #
# Flask 集成
# --------------------------------------------------------------------------- #

def init_app(app, skip_prefixes: tuple = SKIP_PREFIXES) -> str:
    """把审计中间件挂载到 Flask 应用，返回后端描述。"""
    start_worker()

    @app.after_request
    def _record_access(response):  # noqa: ANN001
        path = request_path()
        if any(path.startswith(p) for p in skip_prefixes) or path in SKIP_EXACT:
            return response
        try:
            record_request(response)
        except Exception:
            # 审计失败绝不影响主站
            pass
        return response

    return backend_name()


def request_path() -> str:
    from flask import request as _req
    return _req.path or "/"


def record_request(response) -> None:
    """把当前请求封装成一条日志记录，投递到写入队列。"""
    from flask import request as _req, session as _session

    user = None
    role = None
    try:
        session_user = _session.get("user")
        if isinstance(session_user, dict):
            user = session_user.get("name")
            role = session_user.get("role")
        elif session_user:
            user = str(session_user)
    except Exception:
        pass

    ip = client_ip(_req)
    path = _req.path or "/"
    is_login = bool(user) and path.startswith("/api/")

    record = {
        "visited_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S"),
        "ip": ip,
        "username": user,
        "role": role,
        "method": _req.method,
        "path": path,
        "full_url": _req.full_path.rstrip("?")[:1000],
        "status_code": getattr(response, "status_code", None),
        "user_agent": (_req.headers.get("User-Agent") or "")[:500],
        "referer": (_req.headers.get("Referer") or "")[:500],
        "is_login": is_login,
        "note": "",
    }
    try:
        WRITE_QUEUE.put_nowait(record)
    except queue.Full:
        # 队列满：丢弃最旧的一条，保证最新记录进得来
        try:
            WRITE_QUEUE.get_nowait()
            WRITE_QUEUE.put_nowait(record)
        except Exception:
            pass


def mark_event(username: str, role: str, event: str, ip: str = "", note: str = "",
               path: str = "") -> None:
    """手工记录一个语义事件（如"登录成功"、"登录失败"），补充到审计流。"""
    record = {
        "visited_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S"),
        "ip": ip or "unknown",
        "username": username,
        "role": role,
        "method": "EVENT",
        "path": path or f"/event/{event}",
        "full_url": path or f"/event/{event}",
        "status_code": 200,
        "user_agent": "",
        "referer": "",
        "is_login": event.startswith("login"),
        "note": f"[{event}] {note}".strip(),
    }
    try:
        WRITE_QUEUE.put_nowait(record)
    except queue.Full:
        pass
