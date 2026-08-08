"""Microsoft Graph API client for downloading CC6 report attachments."""

from __future__ import annotations

import logging
import re
import time
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Callable, Iterable

import requests

logger = logging.getLogger(__name__)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SCOPES = ["Mail.Read"]

# 可重试的网络异常类型
RETRYABLE_EXCEPTIONS = (
    requests.exceptions.ConnectionError,
    requests.exceptions.Timeout,
    requests.exceptions.ChunkedEncodingError,
)
# Graph API 限流状态码
RETRYABLE_STATUSES = (429, 500, 502, 503, 504)

# 重试配置
MAX_RETRIES = 4          # 最多重试 4 次（总共 5 次请求）
BASE_BACKOFF_SECS = 2    # 基础退避秒数
MAX_BACKOFF_SECS = 60    # 单次退避上限




@dataclass
class DownloadedAttachment:
    path: Path
    original_filename: str
    subject: str
    is_update: bool
    message_id: str
    received_time: datetime | None


def subject_matches(subject: str | None, keywords: Iterable[str]) -> bool:
    text = (subject or "").lower()
    return all(kw.lower() in text for kw in keywords)


def is_update_subject(subject: str | None, update_keyword: str = "update") -> bool:
    return update_keyword.lower() in (subject or "").lower()


def safe_filename(name: str) -> str:
    name = name.replace("\xa0", " ").strip()
    name = re.sub(r'[<>:"/\\|?*]', "_", name)
    return name or "attachment.xlsx"


def unique_path(directory: Path, filename: str) -> Path:
    directory.mkdir(parents=True, exist_ok=True)
    candidate = directory / filename
    if not candidate.exists():
        return candidate
    stem = candidate.stem
    suffix = candidate.suffix
    for i in range(1, 1000):
        alt = directory / f"{stem}__{i}{suffix}"
        if not alt.exists():
            return alt
    raise RuntimeError(f"Too many name collisions for {filename}")


# ---------------------------------------------------------------------------
# Proxy & HTTP session
# ---------------------------------------------------------------------------

def _get_http_session(config: dict[str, Any]) -> requests.Session:
    """Create a requests.Session with proxy settings from config (if any)."""
    proxy_cfg = config.get("proxy", {})
    http_proxy = proxy_cfg.get("http", "") or None
    https_proxy = proxy_cfg.get("https", "") or None

    session = requests.Session()
    if http_proxy or https_proxy:
        session.proxies = {
            "http": http_proxy or https_proxy,
            "https": https_proxy or http_proxy,
        }
        logger.info("Using proxy: http=%s https=%s", http_proxy, https_proxy)
    return session


# ---------------------------------------------------------------------------
# Authentication (interactive browser + token cache)
# ---------------------------------------------------------------------------

def _resolve_token_cache(config: dict[str, Any]) -> Path:
    """Compute token-cache path relative to the project root."""
    graph_cfg = config.get("graph", {})
    raw = graph_cfg.get("token_cache_file")
    if raw:
        p = Path(raw)
        if p.is_absolute():
            return p
    # fallback: logs/ dir next to log_file
    log_file = config.get("log_file", "./logs/run.log")
    log_dir = Path(log_file).parent if Path(log_file).is_absolute() else Path(log_file)
    if not log_dir.is_absolute():
        log_dir = Path.cwd() / log_dir
    return (log_dir / "graph_token_cache.json").resolve()


def _acquire_token(config: dict[str, Any]) -> str:
    """Authenticate via authorization-code flow with localhost redirect.
    Prints the login URL and tries to open the browser; caches the token."""
    import msal, threading, urllib.parse, webbrowser
    from http.server import HTTPServer, BaseHTTPRequestHandler

    graph_cfg = config.get("graph", {})
    client_id = graph_cfg.get("client_id", "")
    if not client_id:
        raise RuntimeError(
            "请先在 config.json -> graph -> client_id 中填入你的 Azure AD 应用 ID。\n"
            "获取方式：https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/CreateApplicationBlade"
        )
    tenant = graph_cfg.get("tenant", "common")
    cache_path = _resolve_token_cache(config)

    cache = msal.SerializableTokenCache()
    if cache_path.exists():
        with open(cache_path, "r", encoding="utf-8") as f:
            cache.deserialize(f.read())

    app = msal.PublicClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant}",
        token_cache=cache,
    )

    # 1) Try silent (cached) authentication (with retry for transient network issues)
    accounts = app.get_accounts()
    if accounts:
        for attempt in range(MAX_RETRIES + 1):
            try:
                result = app.acquire_token_silent(SCOPES, account=accounts[0])
                if result and "access_token" in result:
                    logger.info("Graph API: using cached token (%s)",
                                 accounts[0].get("username", "unknown"))
                    _save_cache(cache_path, cache)
                    return result["access_token"]
                # token acquisition returned without token but no exception -> expired cache
                break
            except Exception as e:
                if attempt < MAX_RETRIES:
                    wait = min(BASE_BACKOFF_SECS * (2 ** attempt), MAX_BACKOFF_SECS)
                    logger.warning(
                        "[auth] silent token 获取失败 (attempt %d/%d): %s，%ds 后重试...",
                        attempt + 1, MAX_RETRIES + 1, e, wait,
                    )
                    time.sleep(wait)
                else:
                    logger.warning(
                        "[auth] silent token 获取失败，已重试 %d 次: %s，将重新登录",
                        MAX_RETRIES, e,
                    )

    # 2) Authorization-code flow with PKCE – own localhost server
    # Find an available port
    server: HTTPServer | None = None
    port = 8400
    for p in range(8400, 8420):
        try:
            server = HTTPServer(("127.0.0.1", p), type("_H", (BaseHTTPRequestHandler,), {}))
            port = p
            break
        except OSError:
            continue
    if server is None:
        raise RuntimeError("无法绑定本地端口 8400-8419，请关闭占用端口的程序后重试")
    server.server_close()

    redirect_uri = f"http://localhost:{port}"
    flow = app.initiate_auth_code_flow(
        scopes=SCOPES,
        redirect_uri=redirect_uri,
    )

    print()
    print("=" * 58)
    print("  Microsoft Graph API 登录")
    print("=" * 58)
    print()
    print(f"  请用浏览器打开以下地址并完成登录：")
    print()
    print(f"  {flow['auth_uri']}")
    print()
    print("=" * 58)
    print()
    logger.info("Auth code flow initiated, redirect_uri=%s", redirect_uri)

    # Try to open the browser automatically
    webbrowser.open(flow["auth_uri"])

    # Start local server to catch the redirect
    auth_response: dict[str, str] = {}
    response_event = threading.Event()

    class CallbackHandler(BaseHTTPRequestHandler):
        def do_GET(self):
            parsed = urllib.parse.urlparse(self.path)
            qs = urllib.parse.parse_qs(parsed.query)
            # flatten single-value params
            for k, v in qs.items():
                auth_response[k] = v[0] if len(v) == 1 else v
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            if "code" in auth_response:
                self.wfile.write(
                    b"<html><body style='font-family:sans-serif;text-align:center;"
                    b"padding-top:60px'><h2>&#x2705; &#x767B;&#x5F55;&#x6210;&#x529F;</h2>"
                    b"<p>&#x53EF;&#x4EE5;&#x5173;&#x95ED;&#x6B64;&#x9875;&#x9762;&#x3002;</p>"
                    b"</body></html>"
                )
            else:
                err = auth_response.get("error_description", auth_response.get("error", "unknown"))
                self.wfile.write(
                    f"<html><body style='font-family:sans-serif;text-align:center;"
                    f"padding-top:60px'><h2 style='color:red'>&#x26A0; "
                    f"&#x767B;&#x5F55;&#x5931;&#x8D25;</h2><p>{err}</p>"
                    f"</body></html>".encode()
                )
            threading.Thread(target=self.server.shutdown, daemon=True).start()

        def log_message(self, fmt, *args):
            pass  # suppress access logs

    server = HTTPServer(("127.0.0.1", port), CallbackHandler)
    server.timeout = 180  # 3 minutes for user to login
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()

    if "code" not in auth_response:
        err = auth_response.get("error_description", auth_response.get("error", "login timeout"))
        raise RuntimeError(
            f"Login failed: {err}\n"
            "请重新运行并完成浏览器中的登录"
        )

    # Token exchange (with retry for network issues)
    last_err: Exception | None = None
    for attempt in range(MAX_RETRIES + 1):
        try:
            result = app.acquire_token_by_auth_code_flow(
                flow, auth_response, scopes=SCOPES
            )
            if "access_token" in result:
                break
            err = result.get("error_description", result.get("error", "unknown"))
            raise RuntimeError(f"Token exchange failed: {err}")
        except RuntimeError:
            raise  # 非网络错误，直接抛出
        except Exception as e:
            last_err = e
            if attempt < MAX_RETRIES:
                wait = min(BASE_BACKOFF_SECS * (2 ** attempt), MAX_BACKOFF_SECS)
                logger.warning(
                    "[auth] token 交换失败 (attempt %d/%d): %s，%ds 后重试...",
                    attempt + 1, MAX_RETRIES + 1, e, wait,
                )
                time.sleep(wait)
            else:
                raise RuntimeError(
                    f"Token 交换失败，已重试 {MAX_RETRIES} 次。"
                    f"最后错误: {last_err}"
                ) from last_err

    _save_cache(cache_path, cache)
    user = (result.get("id_token_claims", {})
            .get("preferred_username", "unknown"))
    logger.info("Graph API: authenticated as %s", user)
    print(f"  \u2713 登录成功 ({user})\n")
    print("  开始下载邮件附件...\n")
    return result["access_token"]


def _save_cache(path: Path, cache) -> None:
    if cache.has_state_changed:
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            f.write(cache.serialize())


# ---------------------------------------------------------------------------
# Exponential backoff retry helper
# ---------------------------------------------------------------------------

def _retry_with_backoff(
    operation: Callable[[], requests.Response],
    description: str = "API call",
) -> requests.Response:
    """Execute an HTTP request with exponential backoff on transient failures."""
    last_exc: Exception | None = None

    for attempt in range(MAX_RETRIES + 1):
        try:
            resp = operation()
            # 509: Bandwidth limit exceeded (Graph specific)
            # 429: Too many requests (rate limit)
            if resp.status_code in RETRYABLE_STATUSES:
                if attempt < MAX_RETRIES:
                    wait = min(
                        BASE_BACKOFF_SECS * (2 ** attempt),
                        MAX_BACKOFF_SECS,
                    )
                    # Graph API 可能在 Retry-After 头中指定等待时间
                    retry_after = resp.headers.get("Retry-After")
                    if retry_after and retry_after.isdigit():
                        wait = max(wait, int(retry_after))
                    logger.warning(
                        "[%s] HTTP %s (attempt %d/%d), retrying in %ds...",
                        description, resp.status_code, attempt + 1, MAX_RETRIES + 1, wait,
                    )
                    time.sleep(wait)
                    continue
                else:
                    # 最后一次尝试仍然失败，raise
                    resp.raise_for_status()
            # 401: 不重试，直接抛出清晰的错误
            if resp.status_code == 401:
                raise RuntimeError(
                    "Graph API token expired – delete logs/graph_token_cache.json "
                    "and re-run to re-authenticate."
                )
            resp.raise_for_status()
            return resp

        except RETRYABLE_EXCEPTIONS as e:
            last_exc = e
            if attempt < MAX_RETRIES:
                wait = min(BASE_BACKOFF_SECS * (2 ** attempt), MAX_BACKOFF_SECS)
                logger.warning(
                    "[%s] 网络错误 (attempt %d/%d): %s，%ds 后重试...",
                    description, attempt + 1, MAX_RETRIES + 1, e, wait,
                )
                time.sleep(wait)
            # else: 最后一次尝试失败，继续到下面的 raise

    # 所有重试都失败
    raise RuntimeError(
        f"[{description}] 请求失败，已重试 {MAX_RETRIES} 次。"
        f"最后错误: {last_exc}"
    ) from last_exc


# ---------------------------------------------------------------------------
# Graph API HTTP helpers
# ---------------------------------------------------------------------------

def _api_get(token: str, endpoint: str,
             params: dict | None = None) -> dict:
    url = f"{GRAPH_BASE}{endpoint}"
    headers = {"Authorization": f"Bearer {token}"}

    def _do() -> requests.Response:
        return requests.get(url, headers=headers, params=params, timeout=45)

    description = endpoint.split("?")[0].rstrip("/").split("/")[-1] or "GET"
    resp = _retry_with_backoff(_do, f"GET {description}")
    return resp.json()


def _api_get_bytes(token: str, endpoint: str) -> bytes:
    url = f"{GRAPH_BASE}{endpoint}"
    headers = {"Authorization": f"Bearer {token}"}

    def _do() -> requests.Response:
        return requests.get(url, headers=headers, timeout=60)

    # 提取描述信息（如 attachment id）
    parts = endpoint.rstrip("/").split("/")
    desc = "attachment" if "$value" in parts else (parts[-1] if parts else "download")
    resp = _retry_with_backoff(_do, f"GET {desc}")
    return resp.content


# ---------------------------------------------------------------------------
# Folder resolution
# ---------------------------------------------------------------------------

def _resolve_folder_id(token: str, folder_name: str) -> str | None:
    """Find a mail folder by name under Inbox or root. Returns folder id or None."""
    if not folder_name or folder_name.lower() == "inbox":
        return None  # caller uses /inbox directly

    # Search under Inbox first
    data = _api_get(token, "/me/mailFolders/inbox/childFolders")
    for f in data.get("value", []):
        if f.get("displayName", "").lower() == folder_name.lower():
            return f["id"]

    # Search root mail folders
    data = _api_get(token, "/me/mailFolders")
    for f in data.get("value", []):
        if f.get("displayName", "").lower() == folder_name.lower():
            return f["id"]

    logger.warning("Folder '%s' not found; falling back to Inbox", folder_name)
    return None


# ---------------------------------------------------------------------------
# Message listing & attachment download
# ---------------------------------------------------------------------------

def _mails_endpoint(folder_id: str | None) -> str:
    if folder_id:
        return f"/me/mailFolders/{folder_id}/messages"
    return "/me/mailFolders/inbox/messages"


def _message_key(message: dict) -> str:
    """Stable id for deduplication."""
    return message.get("id", "") or message.get("internetMessageId", "")


def _list_messages(
    token: str,
    folder_name: str,
    lookback_days: int,
    subject_keywords: list[str],
) -> list[dict]:
    folder_id = _resolve_folder_id(token, folder_name)
    since = (datetime.now(timezone.utc) - timedelta(days=max(1, int(lookback_days))))
    since_str = since.strftime("%Y-%m-%dT%H:%M:%SZ")

    params: dict[str, Any] = {
        "$filter": f"receivedDateTime ge {since_str}",
        "$orderby": "receivedDateTime desc",
        "$top": 200,
        "$select": "id,subject,receivedDateTime,hasAttachments,internetMessageId",
    }

    all_msgs: list[dict] = []
    endpoint = _mails_endpoint(folder_id)
    first_page = True

    while endpoint:
        p = params if first_page else None
        data = _api_get(token, endpoint, p)
        all_msgs.extend(data.get("value", []))
        next_link = data.get("@odata.nextLink")
        if next_link:
            # nextLink is a full URL; strip the Graph base
            endpoint = next_link[len(GRAPH_BASE):] if next_link.startswith(GRAPH_BASE) else next_link
        else:
            endpoint = None
        first_page = False

    matched = [m for m in all_msgs
               if subject_matches(m.get("subject"), subject_keywords)]
    logger.info("Matched %d mails (folder=%s, keywords=%s, total=%d)",
                len(matched), folder_name, subject_keywords, len(all_msgs))
    return matched


def _list_archive_messages(
    token: str,
    lookback_days: int,
    subject_keywords: list[str],
) -> list[dict]:
    """Search matching messages in Archive folder and Online Archive mailbox.

    Handles two archiving scenarios:
    1. Archive folder in main mailbox   – well-known name ``archive``
    2. Online Archive (In-Place Archive) – well-known name ``archivemsgfolderroot``
    """
    since = (datetime.now(timezone.utc) - timedelta(days=max(1, int(lookback_days))))
    since_str = since.strftime("%Y-%m-%dT%H:%M:%SZ")

    params: dict[str, Any] = {
        "$filter": f"receivedDateTime ge {since_str}",
        "$orderby": "receivedDateTime desc",
        "$top": 200,
        "$select": "id,subject,receivedDateTime,hasAttachments,internetMessageId",
    }

    # Collect archive folder IDs to search
    archive_folders: list[tuple[str, str]] = []  # [(folder_id, display_name), ...]

    # 1) Archive folder in main mailbox
    try:
        archive_id = _resolve_folder_id(token, "Archive")
        if archive_id:
            archive_folders.append((archive_id, "Archive"))
        else:
            # well-known archive folder may not be listed under childFolders;
            # try the well-known endpoint directly
            data = _api_get(token, "/me/mailFolders/archive")
            fid = data.get("id")
            if fid:
                archive_folders.append((fid, "Archive"))
    except Exception:
        logger.debug("Archive folder not found in main mailbox; skipping")

    # 2) Online Archive (In-Place Archive) – recurse into archive mailbox
    try:
        root_data = _api_get(token, "/me/mailFolders/archivemsgfolderroot")
        archive_root_id = root_data.get("id")
        if archive_root_id:
            # list child folders recursively
            _collect_child_folders(token, archive_root_id, "ArchiveMailbox", archive_folders)
    except Exception:
        logger.debug("Online Archive mailbox not found; skipping")

    if not archive_folders:
        logger.info("No archive folders found (neither Archive nor Online Archive)")
        return []

    # Search each archive folder
    all_matched: list[dict] = []
    for fid, display_name in archive_folders:
        try:
            endpoint = f"/me/mailFolders/{fid}/messages"
            folder_msgs: list[dict] = []
            first_page = True
            while endpoint:
                p = params if first_page else None
                data = _api_get(token, endpoint, p)
                folder_msgs.extend(data.get("value", []))
                next_link = data.get("@odata.nextLink")
                endpoint = next_link[len(GRAPH_BASE):] if next_link and next_link.startswith(GRAPH_BASE) else (next_link or None)
                first_page = False
            matched = [m for m in folder_msgs
                       if subject_matches(m.get("subject"), subject_keywords)]
            all_matched.extend(matched)
            logger.info("Archive folder '%s': matched %d / %d mails",
                        display_name, len(matched), len(folder_msgs))
        except Exception as exc:
            logger.debug("Error searching archive folder '%s': %s", display_name, exc)

    return all_matched


def _collect_child_folders(
    token: str,
    parent_id: str,
    prefix: str,
    result: list[tuple[str, str]],
    depth: int = 0,
) -> None:
    """Recursively collect child mail folder ids under *parent_id*."""
    if depth > 5:  # safety limit
        return
    try:
        data = _api_get(token, f"/me/mailFolders/{parent_id}/childFolders")
        for f in data.get("value", []):
            fid = f.get("id")
            name = f.get("displayName", "unknown")
            if fid:
                result.append((fid, f"{prefix}/{name}"))
                _collect_child_folders(token, fid, f"{prefix}/{name}", result, depth + 1)
    except Exception:
        logger.debug("Cannot list child folders under %s/%s", prefix, parent_id)


def _download_attachments_from_message(
    token: str,
    message: dict,
    download_dir: Path,
    *,
    update_keyword: str = "update",
    attachment_extensions: list[str] | None = None,
) -> list[DownloadedAttachment]:
    ext_set = {
        e.lower() if e.startswith(".") else f".{e.lower()}"
        for e in (attachment_extensions or [".xlsx", ".xls"])
    }
    subject = message.get("subject") or ""
    is_update = is_update_subject(subject, update_keyword)
    msg_id = _message_key(message)
    received_str = message.get("receivedDateTime")
    received = (
        datetime.fromisoformat(received_str.replace("Z", "+00:00"))
        if received_str else None
    )

    if not message.get("hasAttachments"):
        return []

    msg_graph_id = message["id"]
    att_data = _api_get(
        token, f"/me/messages/{msg_graph_id}/attachments"
    )
    results: list[DownloadedAttachment] = []

    for att in att_data.get("value", []):
        if att.get("@odata.type") != "#microsoft.graph.fileAttachment":
            continue
        filename = safe_filename(str(att.get("name", "")))
        if Path(filename).suffix.lower() not in ext_set:
            continue

        dest = unique_path(download_dir, filename)
        content = _api_get_bytes(
            token,
            f"/me/messages/{msg_graph_id}/attachments/{att['id']}/$value",
        )
        dest.write_bytes(content)

        result = DownloadedAttachment(
            path=dest,
            original_filename=filename,
            subject=subject,
            is_update=is_update,
            message_id=msg_id,
            received_time=received,
        )
        results.append(result)
        logger.info("Saved attachment: %s (update=%s)", dest.name, is_update)

    return results


# ---------------------------------------------------------------------------
# Public entry-point (same signature as before)
# ---------------------------------------------------------------------------

def fetch_report_attachments(
    config: dict[str, Any],
    processed_message_ids: set[str] | None = None,
) -> list[DownloadedAttachment]:
    outlook_cfg = config.get("outlook", {})
    download_dir = Path(config.get("download_dir", "./downloads"))
    if not download_dir.is_absolute():
        download_dir = (Path.cwd() / download_dir).resolve()

    processed_message_ids = processed_message_ids or set()
    token = _acquire_token(config)

    lookback_days = int(outlook_cfg.get("lookback_days", 30))
    subject_keywords = list(outlook_cfg.get("subject_keywords", ["CC6", "report"]))

    # Search main folder (Inbox / custom folder)
    messages = _list_messages(
        token,
        folder_name=outlook_cfg.get("folder", "Inbox"),
        lookback_days=lookback_days,
        subject_keywords=subject_keywords,
    )

    # Search Archive + Online Archive, deduplicate by message key
    archive_messages = _list_archive_messages(
        token,
        lookback_days=lookback_days,
        subject_keywords=subject_keywords,
    )
    inbox_count = len(messages)
    seen_keys: set[str] = {_message_key(m) for m in messages}
    for am in archive_messages:
        key = _message_key(am)
        if key not in seen_keys:
            seen_keys.add(key)
            messages.append(am)

    logger.info(
        "Total matched: %d (inbox=%d, archive=%d)",
        len(messages), inbox_count, len(archive_messages),
    )

    downloaded: list[DownloadedAttachment] = []
    for msg in messages:
        msg_id = _message_key(msg)
        if msg_id in processed_message_ids:
            subject = msg.get("subject", "(unknown)")[:60]
            logger.info("Skip already processed: %s", subject)
            continue

        try:
            downloaded.extend(
                _download_attachments_from_message(
                    token,
                    msg,
                    download_dir,
                    update_keyword=outlook_cfg.get("update_keyword", "update"),
                    attachment_extensions=list(
                        outlook_cfg.get("attachment_extensions", [".xlsx", ".xls"])
                    ),
                )
            )
        except Exception as exc:
            subject = msg.get("subject", "(unknown)")[:80]
            logger.error(
                "下载附件失败 [%s]: %s（跳过，继续处理下一封）",
                subject, exc,
            )

    return downloaded
