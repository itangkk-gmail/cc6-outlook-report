"""CC6 Outlook daily report downloader and master-table merger."""

from __future__ import annotations

import argparse
import json
import logging
import re
import sys
import traceback
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

from excel_merge import merge_into_master, parse_date_from_filename
from outlook_client import fetch_report_attachments, is_update_subject

ROOT = Path(__file__).resolve().parent
DEFAULT_CONFIG = ROOT / "config.json"




def setup_logging(
    log_file: str | Path | None,
    max_size_mb: int = 0,
    retention_days: int = 0,
) -> None:
    """Configure console + append log, with optional log rotation and cleanup.
    
    Args:
        log_file: Path to the cumulative run.log file.
        max_size_mb: Rotate run.log if it exceeds this many MB (0 = disabled).
        retention_days: Delete archived run.*.log files older than N days (0 = disabled).
    """
    handlers: list[logging.Handler] = [logging.StreamHandler(sys.stdout)]
    path: Path | None = None

    if log_file:
        path = Path(log_file)
        if not path.is_absolute():
            path = ROOT / path
        path.parent.mkdir(parents=True, exist_ok=True)

        # Rotate before opening the log file for writing
        if max_size_mb > 0 and path.exists():
            _rotate_log(path, max_size_mb)

        handlers.append(logging.FileHandler(path, encoding="utf-8"))

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=handlers,
        force=True,
    )

    if path and retention_days > 0:
        _cleanup_archives(path.parent, retention_days)


def _rotate_log(log_path: Path, max_size_mb: int) -> None:
    """Rename run.log to run.YYYYMMDD_HHMMSS.log if it exceeds max_size_mb."""
    max_bytes = max_size_mb * 1024 * 1024
    try:
        size = log_path.stat().st_size
    except OSError:
        return
    if size <= max_bytes:
        return

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    archived = log_path.parent / f"run.{stamp}.log"
    try:
        log_path.rename(archived)
        print(f"[LOG_ROTATE] run.log ({size / (1024 * 1024):.1f} MB) -> {archived.name}")
    except OSError as exc:
        print(f"[LOG_ROTATE] 轮转失败: {exc}")


def _cleanup_archives(log_dir: Path, retention_days: int) -> None:
    """Delete run.*.log archives older than retention_days."""
    logger = logging.getLogger(__name__)
    _ARCHIVE_PATTERN = re.compile(r"^run\.(\d{8})_\d{6}\.log$")
    cutoff = datetime.now().date() - timedelta(days=retention_days)
    deleted = 0
    freed_bytes = 0

    for f in log_dir.glob("run.*.log"):
        m = _ARCHIVE_PATTERN.match(f.name)
        if not m:
            continue
        try:
            file_date = datetime.strptime(m.group(1), "%Y%m%d").date()
        except ValueError:
            continue
        if file_date < cutoff:
            size = f.stat().st_size
            try:
                f.unlink()
                freed_bytes += size
                deleted += 1
            except OSError as exc:
                logger.warning("无法删除过期日志 %s: %s", f.name, exc)

    if deleted:
        logger.info("日志清理: 删除 %d 个过期归档, 释放 %.1f KB", deleted, freed_bytes / 1024)


def resolve_cfg_path(cfg: dict[str, Any], key: str, default_rel: str) -> Path:
    raw = cfg.get(key) or default_rel
    path = Path(raw)
    if not path.is_absolute():
        path = ROOT / path
    return path.resolve()


def write_status(
    cfg: dict[str, Any],
    *,
    ok: bool,
    message: str,
    details: dict[str, Any] | None = None,
    exit_code: int = 0,
) -> None:
    """Write human-readable + JSON status for operators / Task Scheduler."""
    finished = datetime.now().isoformat(timespec="seconds")
    status_name = "SUCCESS" if ok else "FAILED"
    details = details or {}

    status_txt = resolve_cfg_path(cfg, "status_file", "./logs/last_status.txt")
    status_json = resolve_cfg_path(cfg, "status_json", "./logs/last_status.json")
    status_txt.parent.mkdir(parents=True, exist_ok=True)

    lines = [
        f"status={status_name}",
        f"exit_code={exit_code}",
        f"finished_at={finished}",
        f"message={message}",
    ]
    if cfg.get("log_file"):
        lines.append(f"append_log={cfg['log_file']}")
    if cfg.get("master_file"):
        lines.append(f"master_file={cfg['master_file']}")
    for key, value in details.items():
        lines.append(f"{key}={value}")

    status_txt.write_text("\n".join(lines) + "\n", encoding="utf-8")
    payload = {
        "status": status_name,
        "ok": ok,
        "exit_code": exit_code,
        "finished_at": finished,
        "message": message,
        "append_log": cfg.get("log_file"),
        "master_file": cfg.get("master_file"),
        "details": details,
    }
    status_json.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    logger = logging.getLogger(__name__)
    banner = f"======== {status_name}: {message} ========"
    if ok:
        logger.info(banner)
    else:
        logger.error(banner)
    logger.info("Status file: %s", status_txt)


def load_config(path: Path) -> dict[str, Any]:
    with path.open(encoding="utf-8") as f:
        cfg = json.load(f)
    for key in ("download_dir", "master_file", "processed_file", "log_file", "status_file", "status_json"):
        if key in cfg and cfg[key]:
            p = Path(cfg[key])
            if not p.is_absolute():
                cfg[key] = str((ROOT / p).resolve())
    return cfg


def load_processed(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {"messages": {}}
    with path.open(encoding="utf-8") as f:
        data = json.load(f)
    if "messages" not in data:
        data = {"messages": data if isinstance(data, dict) else {}}
    return data


def save_processed(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def merge_one(
    cfg: dict[str, Any],
    file_path: Path,
    *,
    is_update: bool,
) -> dict[str, Any]:
    excel_cfg = cfg.get("excel", {})
    result = merge_into_master(
        file_path,
        cfg["master_file"],
        is_update=is_update,
        master_sheet=cfg.get("master_sheet", "Master"),
        sheet_name=excel_cfg.get("sheet_name", "DAILY REPORT"),
        header_keyword=excel_cfg.get("header_keyword", "ID /"),
        data_start_col=int(excel_cfg.get("data_start_col", 8)),
        data_end_col=int(excel_cfg.get("data_end_col", 23)),
    )
    return result


def run_merge_only(args: argparse.Namespace, cfg: dict[str, Any]) -> int:
    logger = logging.getLogger(__name__)
    files = [Path(f) for f in args.file]
    if not files:
        write_status(cfg, ok=False, message="--merge-only requires at least one --file", exit_code=2)
        return 2

    update_kw = cfg.get("outlook", {}).get("update_keyword", "update")
    merged_files = 0
    last_master_rows = None
    for file_path in files:
        if not file_path.is_absolute():
            file_path = (Path.cwd() / file_path).resolve()
        if not file_path.exists():
            write_status(cfg, ok=False, message=f"File not found: {file_path}", exit_code=1)
            return 1
        date_str = parse_date_from_filename(file_path.name)
        if not date_str:
            write_status(cfg, ok=False, message=f"Cannot parse date from filename: {file_path.name}", exit_code=1)
            return 1
        is_update = bool(args.update) or is_update_subject(file_path.name, update_kw)
        logger.info("Merging %s (date=%s, update=%s)", file_path.name, date_str, is_update)
        result = merge_one(cfg, file_path, is_update=is_update)
        last_master_rows = result["master_rows"]
        merged_files += 1
        logger.info(
            "Merged: added=%s removed=%s master_rows=%s -> %s",
            result["rows_added"],
            result["rows_removed"],
            result["master_rows"],
            result["master_file"],
        )

    write_status(
        cfg,
        ok=True,
        message=f"merge-only completed ({merged_files} file(s))",
        details={"merged_files": merged_files, "master_rows": last_master_rows},
        exit_code=0,
    )
    return 0


def run_outlook_pipeline(cfg: dict[str, Any]) -> int:
    logger = logging.getLogger(__name__)
    processed_path = Path(cfg["processed_file"])
    processed = load_processed(processed_path)

    # Extract message IDs that have already been fully processed (merged or errored)
    processed_msg_ids: set[str] = set()
    for dedupe_key, info in processed.get("messages", {}).items():
        status = info.get("status", "")
        if status in ("merged", "merge_error"):
            msg_id = dedupe_key.split("|", 1)[0]
            processed_msg_ids.add(msg_id)
    logger.info("Loaded %d processed message IDs", len(processed_msg_ids))

    try:
        attachments = fetch_report_attachments(cfg, processed_message_ids=processed_msg_ids)
    except RuntimeError as exc:
        logger.error("%s", exc)
        write_status(cfg, ok=False, message=str(exc), exit_code=1)
        return 1
    except Exception as exc:
        logger.error("Failed to access Outlook: %s", exc)
        write_status(
            cfg,
            ok=False,
            message=f"Failed to access Outlook: {exc}",
            details={"traceback": traceback.format_exc()},
            exit_code=1,
        )
        return 1

    if not attachments:
        logger.info("No matching attachments found.")
        write_status(
            cfg,
            ok=True,
            message="No matching attachments found",
            details={"merged": 0, "skipped": 0, "errors": 0, "attachments": 0},
            exit_code=0,
        )
        return 0

    attachments_sorted = sorted(
        attachments,
        key=lambda a: (a.received_time or datetime.min, a.path.name),
    )

    merged_count = 0
    skipped = 0
    error_count = 0
    no_date_count = 0

    for att in attachments_sorted:
        dedupe_key = f"{att.message_id}|{att.original_filename}"
        if dedupe_key in processed["messages"]:
            logger.info("Skip already processed: %s", att.path.name)
            skipped += 1
            continue

        date_str = parse_date_from_filename(att.path.name)
        if not date_str:
            logger.error("Saved but cannot parse date from filename (skip merge): %s", att.path.name)
            processed["messages"][dedupe_key] = {
                "file": str(att.path),
                "subject": att.subject,
                "status": "saved_no_date",
                "processed_at": datetime.now().isoformat(timespec="seconds"),
            }
            no_date_count += 1
            continue

        try:
            result = merge_one(cfg, att.path, is_update=att.is_update)
            logger.info(
                "Merged %s date=%s update=%s added=%s removed=%s",
                att.path.name,
                result["date"],
                att.is_update,
                result["rows_added"],
                result["rows_removed"],
            )
            processed["messages"][dedupe_key] = {
                "file": str(att.path),
                "subject": att.subject,
                "date": result["date"],
                "is_update": att.is_update,
                "rows_added": result["rows_added"],
                "status": "merged",
                "processed_at": datetime.now().isoformat(timespec="seconds"),
            }
            merged_count += 1
        except Exception as exc:
            logger.error("Merge failed for %s: %s", att.path.name, exc)
            processed["messages"][dedupe_key] = {
                "file": str(att.path),
                "subject": att.subject,
                "status": "merge_error",
                "error": str(exc),
                "processed_at": datetime.now().isoformat(timespec="seconds"),
            }
            error_count += 1

    save_processed(processed_path, processed)
    details = {
        "merged": merged_count,
        "skipped": skipped,
        "errors": error_count,
        "no_date": no_date_count,
        "attachments": len(attachments),
    }
    logger.info(
        "Finished. merged=%s skipped=%s errors=%s total_attachments=%s",
        merged_count,
        skipped,
        error_count,
        len(attachments),
    )

    if error_count > 0:
        write_status(
            cfg,
            ok=False,
            message=f"Completed with {error_count} merge error(s)",
            details=details,
            exit_code=1,
        )
        return 1

    write_status(
        cfg,
        ok=True,
        message=f"Completed: merged={merged_count}, skipped={skipped}",
        details=details,
        exit_code=0,
    )
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Download CC6 Outlook reports and merge into a master workbook.")
    parser.add_argument("--config", default=str(DEFAULT_CONFIG), help="Path to config.json")
    parser.add_argument(
        "--merge-only",
        action="store_true",
        help="Only merge local Excel files; do not touch Outlook",
    )
    parser.add_argument(
        "--file",
        action="append",
        default=[],
        help="Local xlsx to merge (repeatable). Used with --merge-only",
    )
    parser.add_argument(
        "--update",
        action="store_true",
        help="Treat --merge-only files as update (replace that date in master)",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    config_path = Path(args.config)
    if not config_path.is_absolute():
        config_path = (Path.cwd() / config_path).resolve()
    if not config_path.exists():
        print(f"Config not found: {config_path}", file=sys.stderr)
        return 1

    cfg = load_config(config_path)
    setup_logging(
        cfg.get("log_file"),
        max_size_mb=int(cfg.get("log_max_size_mb", 0)),
        retention_days=int(cfg.get("log_retention_days", 0)),
    )
    logger = logging.getLogger(__name__)
    logger.info("Config: %s", config_path)

    try:
        if args.merge_only:
            return run_merge_only(args, cfg)
        return run_outlook_pipeline(cfg)
    except Exception as exc:
        logger.error("Unhandled error: %s", exc)
        write_status(
            cfg,
            ok=False,
            message=f"Unhandled error: {exc}",
            details={"traceback": traceback.format_exc()},
            exit_code=1,
        )
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
