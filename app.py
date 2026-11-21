#!/usr/bin/env python3
"""
CAMIS Marks - desktop app.

Three clicks:
  1. choose the Excel file (or drop it / pick a recent one)
  2. open CAMIS  -> log in and open the class marks page in the browser
  3. fill marks  -> typed, saved, and a list of anything left to do by hand

The browser stays open between files, so the next class is just
"choose file -> fill".

All browser work happens on one worker thread (Playwright's sync API is not
thread-safe). The UI is an HTML page in ui/index.html rendered by pywebview;
Python owns the state and pushes it to the page with window.render(state).
"""

from __future__ import annotations

import json
import os
import queue
import subprocess
import sys
import threading
import time
import traceback
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

import webview

import camis_grader as cg
import notify

APP_NAME = "CAMIS Marks"
APP_VERSION = "2.0.0"  # set by the release workflow from CHANGELOG.md

if getattr(sys, "frozen", False):
    RES_DIR = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
else:
    RES_DIR = Path(__file__).resolve().parent

DATA_DIR = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / APP_NAME
PROFILE_DIR = DATA_DIR / "browser-profile"
REPORTS_DIR = Path.home() / "Documents" / "CAMIS Reports"
SETTINGS_FILE = DATA_DIR / "settings.json"
EXCEL_SUFFIXES = {".xls", ".xlsx", ".xlsm"}

DATA_DIR.mkdir(parents=True, exist_ok=True)
cg.DEFAULT_URL = os.environ.get("CAMIS_URL", cg.DEFAULT_URL)  # for testing against a mock

# A windowed .exe has no console: keep a log for troubleshooting instead.
if getattr(sys, "frozen", False) or sys.stdout is None:
    _log = open(DATA_DIR / "app.log", "a", encoding="utf-8", buffering=1)
    sys.stdout = sys.stderr = _log
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(errors="replace")  # never crash on a name the console can't show
    except Exception:
        pass
print(f"\n--- {APP_NAME} started {datetime.now():%Y-%m-%d %H:%M:%S} ---")


# ---------------------------------------------------------------------------
# Human wording for the report statuses
# ---------------------------------------------------------------------------

ATTENTION = {
    # status: (short label, tag level)
    "not-found": ("not found on camis", "warn"),
    "site-not-in-excel": ("not in your excel file", "warn"),
    "conflict": ("camis has a different mark", "warn"),
    "skipped-reason": ("has an absence reason", "warn"),
    "no-excel-mark": ("mark in excel is not a number", "err"),
    "invalid-excel-mark": ("mark in excel is wrong", "err"),
    "out-of-range": ("mark is above the maximum", "err"),
    "not-editable": ("cell is locked", "err"),
    "failed": ("could not type it", "err"),
    "not-entered": ("stopped before this one", "warn"),
}


def attention_rows(plan: cg.Plan) -> list[dict[str, str]]:
    rows = []
    for rec in plan.records + plan.site_only:
        if rec.status not in ATTENTION:
            continue
        why, level = ATTENTION[rec.status]
        if rec.status == "site-not-in-excel":
            detail = f"now: {rec.site_mark_before}" if rec.site_mark_before else ""
            rows.append({"name": rec.site_name, "id": rec.site_number, "mark": "",
                         "why": why, "level": level, "detail": detail})
            continue
        detail = ""
        if rec.status == "conflict":
            detail = rec.reason
        elif rec.status in {"failed", "skipped-reason", "invalid-excel-mark", "out-of-range"}:
            detail = rec.reason
        elif rec.status == "no-excel-mark":
            detail = f"excel says: {rec.excel_mark}"
        rows.append({"name": rec.excel_name, "id": rec.excel_id, "mark": rec.target_mark,
                     "why": why, "level": level, "detail": detail})
    return rows


def ago(ts: float) -> str:
    sec = max(0, time.time() - ts)
    if sec < 3600:
        return f"{max(1, int(sec // 60))} min ago"
    if sec < 86400:
        return f"{int(sec // 3600)} h ago"
    days = int(sec // 86400)
    return "yesterday" if days == 1 else f"{days} days ago"


def recent_excel_files(limit: int = 5) -> list[dict[str, Any]]:
    home = Path.home()
    folders = [home / "Downloads", home / "Desktop", home / "Documents"]
    cutoff = time.time() - 30 * 86400
    found: list[tuple[float, Path]] = []
    for folder in folders:
        try:
            for p in folder.iterdir():
                if p.suffix.lower() in EXCEL_SUFFIXES and not p.name.startswith("~$") and p.is_file():
                    mtime = p.stat().st_mtime
                    if mtime >= cutoff:
                        found.append((mtime, p))
        except OSError:
            continue
    found.sort(reverse=True)
    return [{"path": str(p), "name": p.name, "folder": p.parent.name.lower(), "when": ago(m)}
            for m, p in found[:limit]]


def open_path(path: Path) -> None:
    try:
        if os.name == "nt":
            os.startfile(str(path))  # type: ignore[attr-defined]
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except OSError as exc:
        print(f"could not open {path}: {exc}")


def friendly_error(exc: BaseException) -> str:
    text = str(exc).strip() or exc.__class__.__name__
    low = text.lower()
    if "could not open a browser" in low or "executable doesn't exist" in low:
        return ("Could not open a browser. Please install Google Chrome or Microsoft Edge, "
                "then try again.")
    if "user data directory is already in use" in low or "processsingleton" in low:
        return ("The CAMIS browser is already open from another copy of this app. "
                "Close all CAMIS Marks windows and try again.")
    if "target page, context or browser has been closed" in low or "browser has been closed" in low:
        return "The browser window was closed while working. Nothing more was typed."
    return text.splitlines()[0] if len(text) > 400 else text


def load_settings() -> dict[str, Any]:
    try:
        return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_settings(settings: dict[str, Any]) -> None:
    try:
        SETTINGS_FILE.write_text(json.dumps(settings, indent=2), encoding="utf-8")
    except OSError as exc:
        print(f"could not save settings: {exc}")


# ---------------------------------------------------------------------------
# API exposed to the page. Only public methods are visible to JavaScript;
# everything internal starts with an underscore so pywebview ignores it.
# ---------------------------------------------------------------------------


class Api:
    def __init__(self) -> None:
        self._window: Optional[webview.Window] = None
        self._hwnd = 0
        self._settings = load_settings()
        notify.set_enabled(self._settings.get("notifications", True))
        self._lock = threading.RLock()
        self._state: dict[str, Any] = {
            "phase": "idle",          # idle | waiting | review | filling | done | error
            "browser": "closed",      # closed | opening | open
            "file": None,
            "fileError": None,
            "loading": None,
            "notice": None,
            "recent": recent_excel_files(),
            "found": 0,
            "waitSame": False,
            "review": None,
            "progress": None,
            "result": None,
            "error": None,
            "version": APP_VERSION,
            "notifications": self._settings.get("notifications", True),
        }
        self._workbook: Optional[cg.MarksWorkbook] = None
        self._table: Optional[dict[str, Any]] = None
        self._site: list[cg.SiteStudent] = []
        self._columns: list[dict[str, Any]] = []
        self._selected = 0
        self._overwrite = False
        self._plan: Optional[cg.Plan] = None
        self._last_names: Optional[tuple] = None
        self._report_path: Optional[Path] = None

        self._jobs: "queue.Queue[str]" = queue.Queue()
        self._cancel = threading.Event()
        self._use_current = threading.Event()
        self._pw: Any = None
        self._ctx: Any = None
        self._ctx_closed = True
        threading.Thread(target=self._worker, name="browser", daemon=True).start()

    # ------------------------------------------------------------ state
    def _snapshot(self) -> dict[str, Any]:
        with self._lock:
            return json.loads(json.dumps(self._state, default=str))

    def _set(self, push: bool = True, **changes: Any) -> dict[str, Any]:
        with self._lock:
            self._state.update(changes)
            snap = self._snapshot()
        if push and self._window is not None:
            try:
                self._window.evaluate_js(f"window.render && window.render({json.dumps(snap)})")
            except Exception as exc:  # window closing
                print(f"push failed: {exc}")
        return snap

    def _phase(self) -> str:
        with self._lock:
            return self._state["phase"]

    # ------------------------------------------------------------ notifications
    def _notify(self, title: str, body: str, kind: str = "info") -> None:
        """Windows notification + blinking taskbar button (if the app is behind)."""
        print(f"notify [{kind}] {title}: {body}")
        notify.toast(title, body, kind)
        if not self._hwnd:
            self._hwnd = notify.find_window(APP_NAME)
        if not notify.is_foreground(self._hwnd):
            notify.flash(self._hwnd)

    def toggle_notifications(self) -> dict[str, Any]:
        on = not self._settings.get("notifications", True)
        self._settings["notifications"] = on
        save_settings(self._settings)
        notify.set_enabled(on)
        snap = self._set(notifications=on)
        if on:
            notify.toast("Notifications are on",
                         "You'll get a message here when marks are done or something needs you.")
        return snap

    # ------------------------------------------------------------ file
    def get_state(self) -> dict[str, Any]:
        return self._snapshot()

    def pick_file(self) -> dict[str, Any]:
        if self._phase() not in {"idle", "error"} or self._window is None:
            return self._snapshot()
        start = Path.home() / "Downloads"
        try:
            dialog = webview.FileDialog.OPEN  # pywebview >= 5
        except AttributeError:
            dialog = webview.OPEN_DIALOG
        result = self._window.create_file_dialog(
            dialog,
            directory=str(start if start.is_dir() else Path.home()),
            file_types=("Excel files (*.xls;*.xlsx;*.xlsm)", "All files (*.*)"),
        )
        if not result:
            return self._snapshot()
        path = result[0] if isinstance(result, (list, tuple)) else result
        return self._load_file(Path(path))

    def use_recent(self, index: int) -> dict[str, Any]:
        recent = self._state.get("recent") or []
        if 0 <= index < len(recent):
            return self._load_file(Path(recent[index]["path"]))
        return self._snapshot()

    def _on_drop(self, event: dict[str, Any]) -> None:
        try:
            files = event.get("dataTransfer", {}).get("files", [])
            paths = [f.get("pywebviewFullPath") for f in files if f.get("pywebviewFullPath")]
        except Exception:
            paths = []
        if not paths:
            return
        if self._phase() in {"idle", "error"}:
            self._load_file(Path(paths[0]))

    def _load_file(self, path: Path) -> dict[str, Any]:
        if path.suffix.lower() not in EXCEL_SUFFIXES | {".htm", ".html"}:
            return self._set(file=None, fileError=f"{path.name} is not an Excel file.", notice=None)
        self._set(loading=path.name, fileError=None, notice=None)
        try:
            wb = cg.load_workbook(path)
        except Exception as exc:
            self._workbook = None
            return self._set(file=None, fileError=str(exc), notice=None, phase="idle", loading=None)
        self._workbook = wb
        blanks = sum(1 for s in wb.students if not s.raw_mark.strip())
        col = wb.mark_column or "?"
        if wb.excel_max is not None and str(int(wb.excel_max) if wb.excel_max.is_integer() else wb.excel_max) not in col:
            col += f" (out of {cg.format_number(wb.excel_max)})"
        info = {
            "path": str(path),
            "name": path.name,
            "subject": wb.meta.get("Subject", ""),
            "klass": wb.meta.get("Class/Term", ""),
            "students": len(wb.students),
            "column": col,
            "blanks": blanks,
        }
        return self._set(file=info, fileError=None, notice=None, phase="idle", error=None, loading=None)

    def clear_file(self) -> dict[str, Any]:
        self._workbook = None
        return self._set(file=None, fileError=None, notice=None, recent=recent_excel_files())

    # ------------------------------------------------------------ flow
    def start(self) -> dict[str, Any]:
        if self._workbook is None or self._phase() != "idle":
            return self._snapshot()
        self._cancel.clear()
        self._use_current.clear()
        snap = self._set(phase="waiting", found=0, waitSame=False, notice=None)
        self._jobs.put("scan")
        return snap

    def cancel(self) -> dict[str, Any]:
        self._cancel.set()
        return self._set(phase="idle", found=0, waitSame=False)

    def use_current_page(self) -> dict[str, Any]:
        self._use_current.set()
        return self._snapshot()

    def rescan(self) -> dict[str, Any]:
        if self._phase() != "review":
            return self._snapshot()
        self._cancel.clear()
        self._use_current.set()  # the teacher asked for this page explicitly
        snap = self._set(phase="waiting", found=0, waitSame=False, notice=None)
        self._jobs.put("scan")
        return snap

    def select_column(self, index: int) -> dict[str, Any]:
        if self._phase() == "review" and 0 <= index < len(self._columns):
            self._selected = index
            self._replan()
        return self._snapshot()

    def set_overwrite(self, value: bool) -> dict[str, Any]:
        if self._phase() == "review":
            self._overwrite = bool(value)
            self._replan()
        return self._snapshot()

    def fill(self) -> dict[str, Any]:
        if self._phase() != "review" or self._plan is None:
            return self._snapshot()
        self._cancel.clear()
        total = len(self._plan.to_enter)
        snap = self._set(phase="filling", notice=None,
                         progress={"done": 0, "total": total, "current": "", "saving": False})
        self._jobs.put("fill")
        return snap

    def stop(self) -> dict[str, Any]:
        self._cancel.set()
        return self._snapshot()

    def next_file(self) -> dict[str, Any]:
        self._cancel.set()
        self._workbook = None
        self._plan = None
        return self._set(phase="idle", file=None, fileError=None, notice=None, review=None,
                         result=None, progress=None, error=None, recent=recent_excel_files())

    def dismiss(self) -> dict[str, Any]:
        return self._set(phase="idle", error=None)

    def open_report(self) -> dict[str, Any]:
        if self._report_path and self._report_path.exists():
            open_path(self._report_path)
        return self._snapshot()

    def open_reports_folder(self) -> dict[str, Any]:
        REPORTS_DIR.mkdir(parents=True, exist_ok=True)
        open_path(REPORTS_DIR)
        return self._snapshot()

    def open_help(self) -> dict[str, Any]:
        guide = RES_DIR / "HOW TO USE.txt"
        if guide.exists():
            open_path(guide)
        return self._snapshot()

    # ------------------------------------------------------------ review
    def _build_review(self, table: dict[str, Any], keep_choice: bool = False) -> None:
        self._table = table
        self._site = cg.site_students_from(table)
        columns = cg.assessment_columns(table["headers"])
        if not columns:
            raise RuntimeError("The page has a student table but no assessment columns.")
        same_columns = keep_choice and [c["title"] for c in columns] == [c["title"] for c in self._columns]
        self._columns = columns
        if not same_columns:
            self._selected = self._guess_column()
            self._overwrite = False
        self._replan()

    def _guess_column(self) -> int:
        """Prefer the editable column with the most empty cells (the one still to fill)."""
        best, best_score = 0, -1
        for i, col in enumerate(self._columns):
            score = 0
            for st in self._site:
                cell = next((c for c in st.cells if c["tdIndex"] == col["tdIndex"]), None)
                if cell and cell.get("buttonId"):
                    score += 1
                    if cg.parse_cell_text(cell.get("text", ""))[0] == "empty":
                        score += 2
            if score > best_score:
                best, best_score = i, score
        return best

    def _replan(self) -> None:
        assert self._workbook is not None
        wb = self._workbook
        target = self._columns[self._selected]
        plan = cg.plan_entries(wb, self._site, target, overwrite=self._overwrite)
        self._plan = plan

        statuses = [r.status for r in plan.records]
        conflicts = statuses.count("conflict") + sum(
            1 for r in plan.records if r.status == "to-enter" and r.reason.startswith("replaced"))
        attention = len(attention_rows(plan))
        preview = [{"name": rec.site_name, "before": rec.site_mark_before
                    if cg.parse_cell_text(rec.site_mark_before)[0] != "empty" else "",
                    "mark": rec.target_mark,
                    "note": f"excel: {rec.excel_name}" if cg.normalize_name(rec.excel_name) != cg.normalize_name(rec.site_name) else ""}
                   for rec, _site in plan.to_enter]
        excel_count = len(wb.students)
        review = {
            "columns": [{"i": i, "title": c["title"],
                         "max": cg.format_number(c["max"]) if c.get("max") else None}
                        for i, c in enumerate(self._columns)],
            "selected": self._selected,
            "siteCount": len(self._site),
            "excelCount": excel_count,
            "matched": plan.matched,
            "lowMatch": excel_count > 0 and plan.matched < max(1, excel_count * 0.5),
            "factor": plan.factor,
            "factorText": cg.format_number(plan.factor) if plan.factor else "1",
            "excelMax": cg.format_number(wb.excel_max) if wb.excel_max else "?",
            "siteMax": cg.format_number(plan.site_max) if plan.site_max else "?",
            "toEnter": len(plan.to_enter),
            "correct": statuses.count("already-correct"),
            "attention": attention,
            "conflicts": conflicts,
            "overwrite": self._overwrite,
            "preview": preview,
        }
        self._set(review=review)

    # ------------------------------------------------------------ worker thread
    def _worker(self) -> None:
        while True:
            job = self._jobs.get()
            if job == "quit":
                break
            try:
                if job == "scan":
                    self._job_scan()
                elif job == "fill":
                    self._job_fill()
            except Exception as exc:
                traceback.print_exc()
                if self._ctx_closed:
                    self._set(phase="idle", notice="The browser was closed. Press the button to open it again.",
                              browser="closed")
                    self._notify("The CAMIS browser was closed",
                                 "Marking stopped. Open the app and press the button to continue.", "warning")
                else:
                    message = friendly_error(exc)
                    self._set(phase="error", error=message)
                    self._notify("CAMIS Marks needs you", message, "error")
        self._shutdown_browser()

    def _browser_alive(self) -> bool:
        if self._ctx is None or self._ctx_closed:
            return False
        try:
            return any(not p.is_closed() for p in self._ctx.pages)
        except Exception:
            return False

    def _on_ctx_close(self, *_: Any) -> None:
        self._ctx_closed = True
        self._set(browser="closed")

    def _ensure_browser(self) -> None:
        if self._browser_alive():
            return
        if self._ctx is not None:
            try:
                self._ctx.close()
            except Exception:
                pass
            self._ctx = None
        self._set(browser="opening")
        if self._pw is None:
            from playwright.sync_api import sync_playwright
            self._pw = sync_playwright().start()
        PROFILE_DIR.mkdir(parents=True, exist_ok=True)
        try:
            self._ctx = cg.launch_browser(self._pw, PROFILE_DIR, "auto")
        except Exception:
            self._set(browser="closed")
            raise
        self._ctx_closed = False
        self._ctx.on("close", self._on_ctx_close)
        self._last_names = None
        self._set(browser="open")
        page = self._ctx.pages[0] if self._ctx.pages else self._ctx.new_page()
        page.set_default_timeout(8_000)
        try:
            page.goto(cg.DEFAULT_URL, wait_until="domcontentloaded", timeout=90_000)
        except Exception as exc:  # offline / slow: the teacher can reload by hand
            print(f"goto failed: {exc}")

    def _find_table(self) -> Optional[tuple[Any, dict[str, Any]]]:
        for page in list(self._ctx.pages):
            if page.is_closed():
                continue
            try:
                info = page.evaluate(cg.EXTRACT_TABLE_JS)
            except Exception:
                continue
            if not info:
                continue
            rows = [r for r in info.get("rows", []) if r.get("name") or r.get("number")]
            if rows and any(c.get("buttonId") for r in rows for c in r.get("cells", [])):
                return page, {**info, "rows": rows}
        return None

    @staticmethod
    def _names(info: dict[str, Any]) -> tuple:
        return tuple((r.get("name"), r.get("number")) for r in info["rows"])

    def _job_scan(self) -> None:
        if self._phase() != "waiting":
            return
        started = time.time()
        self._ensure_browser()
        last_sig = None
        stable_since = 0.0
        while True:
            if self._cancel.is_set() or self._phase() != "waiting":
                return
            if not self._browser_alive():
                self._ctx_closed = True
                self._set(phase="idle", browser="closed", found=0,
                          notice="The browser was closed. Press the button to open it again.")
                self._notify("The CAMIS browser was closed",
                             "Open the app and press the button to open it again.", "warning")
                return
            found = self._find_table()
            if found is None:
                last_sig = None
                if self._state.get("found"):
                    self._set(found=0)
            else:
                page, info = found
                sig = (page.url, tuple((r.get("name"), tuple(c.get("text") for c in r["cells"])) for r in info["rows"]))
                if sig != last_sig:
                    last_sig, stable_since = sig, time.time()
                    self._set(found=len(info["rows"]))
                elif time.time() - stable_since >= 2.0:
                    same = self._last_names is not None and self._names(info) == self._last_names
                    if same and not self._use_current.is_set():
                        if not self._state.get("waitSame"):
                            self._set(waitSame=True, found=0)
                            self._notify("Still on the class you just finished",
                                         "Open the next class in the browser, or click 'use this page' in the app.",
                                         "warning")
                    else:
                        if self._cancel.is_set():
                            return
                        self._build_review(info)
                        self._set(phase="review", waitSame=False, found=0)
                        if time.time() - started > 4:  # the teacher was busy in the browser
                            self._notify_found()
                        return
            time.sleep(0.5)

    def _job_fill(self) -> None:
        assert self._plan is not None and self._workbook is not None and self._table is not None
        if not self._browser_alive():
            self._ctx_closed = True
            self._set(phase="idle", browser="closed",
                      notice="The browser was closed. Press the button to open it again.")
            self._notify("The CAMIS browser was closed",
                         "Nothing was typed. Press the button in the app to open it again.", "warning")
            return

        # Re-read the page right before typing: if the teacher navigated
        # away since the review, the row positions would be wrong.
        found = self._find_table()
        if found is None or self._names(found[1]) != self._names(self._table):
            if found is not None:
                self._build_review(found[1])
            self._set(phase="review" if found else "idle",
                      notice="The CAMIS page changed since you checked it. Please look at the numbers again."
                      if found else "The marks table is no longer open in the browser.")
            self._notify("Nothing was typed: the CAMIS page changed",
                         "Please check the numbers in the app again before filling.", "warning")
            return
        page, info = found
        self._build_review(info, keep_choice=True)  # fresh cell values, same choices
        plan = self._plan
        wb = self._workbook
        td_index = plan.target["tdIndex"]
        total = len(plan.to_enter)
        entered = 0
        stopped = False

        for i, (record, site) in enumerate(plan.to_enter):
            if self._cancel.is_set():
                stopped = True
                break
            self._set(progress={"done": i, "total": total, "saving": False,
                                "current": f"{site.name}  →  {record.target_mark}"})
            try:
                cg.enter_mark(page, site.index, td_index, record.target_mark)
                record.status = "entered"
                entered += 1
            except Exception as exc:
                if not self._browser_alive():
                    raise
                record.status = "failed"
                record.reason = str(exc).splitlines()[0]

        for record, _ in plan.to_enter:
            if record.status == "to-enter":
                record.status = "not-entered"
                record.reason = "stopped before this student"

        if stopped:
            save_state, save_detail = "stopped", "stopped by the teacher"
        elif entered:
            self._set(progress={"done": total, "total": total, "saving": True, "current": ""})
            status, save_detail = cg.save_marks(page)
            save_state = "ok" if status == "ok" else "unsure"
        else:
            save_state, save_detail = "none", "no new marks"

        try:
            txt, _csv = cg.write_report(
                workbook=wb, records=plan.records, site_only=plan.site_only, target=plan.target,
                site_max=plan.site_max, factor=plan.factor, save_status=save_state,
                save_detail=save_detail, site_url=info.get("url", page.url),
                site_count=len(self._site), out_dir=REPORTS_DIR,
            )
            self._report_path = txt
        except Exception as exc:
            print(f"report failed: {exc}")
            self._report_path = None

        self._last_names = self._names(info)
        statuses = [r.status for r in plan.records]
        subject = wb.meta.get("Subject") or wb.path.stem
        result = {
            "title": f"{subject} · {plan.target['title']}",
            "entered": entered,
            "correct": statuses.count("already-correct"),
            "attention": attention_rows(plan),
            "saveState": save_state,
            "saveDetail": save_detail,
        }
        self._set(phase="done", result=result, progress=None)
        self._notify_done(result, statuses.count("failed"))

    def _notify_found(self) -> None:
        r = self._state.get("review") or {}
        if r.get("lowMatch"):
            self._notify("Is this the right class?",
                         f"Only {r.get('matched')} of {r.get('excelCount')} students in your file are on this "
                         "CAMIS page. Check the browser.", "warning")
        else:
            n = r.get("toEnter", 0)
            self._notify("Marks page found",
                         f"{r.get('siteCount')} students on the page. Come back to the app to fill "
                         f"{n} mark{'s' if n != 1 else ''}.")

    def _notify_done(self, result: dict[str, Any], failed: int) -> None:
        todo = len(result["attention"])
        left = (f"{todo} student{'s' if todo != 1 else ''} need you to check." if todo
                else "Nothing left to do by hand.")
        state = result["saveState"]
        if state == "ok":
            kind = "warning" if failed else "success"
            self._notify(f"{result['entered']} marks filled and saved", f"{result['title']}. {left}", kind)
        elif state == "none":
            self._notify("Nothing new to fill", f"{result['title']}. {left}", "success")
        elif state == "stopped":
            self._notify("Stopped, nothing was saved",
                         "Reload the CAMIS page to throw the typed marks away, or click Save in CAMIS to keep them.",
                         "warning")
        else:
            self._notify("Marks typed but NOT saved",
                         "Please click Save in the CAMIS page yourself.", "error")

    def _shutdown_browser(self) -> None:
        try:
            if self._ctx is not None:
                self._ctx.close()
        except Exception:
            pass
        try:
            if self._pw is not None:
                self._pw.stop()
        except Exception:
            pass

    # ------------------------------------------------------------ lifecycle
    def _attach(self, window: webview.Window) -> None:
        self._window = window
        self._hwnd = notify.find_window(APP_NAME)
        try:
            from webview.dom import DOMEventHandler

            def swallow(_e: Any) -> None:
                pass

            doc = window.dom.document
            doc.events.dragover += DOMEventHandler(swallow, True, True, debounce=500)
            doc.events.drop += DOMEventHandler(self._on_drop, True, True)
        except Exception as exc:
            print(f"drag and drop unavailable: {exc}")

    def _closing(self) -> None:
        self._cancel.set()
        self._jobs.put("quit")


def main() -> None:
    notify.register(RES_DIR / "assets" / "icon.png")
    api = Api()
    window = webview.create_window(
        APP_NAME,
        url=str(RES_DIR / "ui" / "index.html"),
        js_api=api,
        width=1040,
        height=760,
        min_size=(860, 620),
        background_color="#0a0a0a",
        text_select=False,
    )
    window.events.closing += api._closing
    webview.start(api._attach, window, private_mode=False,
                  storage_path=str(DATA_DIR / "webview"))


if __name__ == "__main__":
    main()
