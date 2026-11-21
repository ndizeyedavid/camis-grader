#!/usr/bin/env python3
"""
CAMIS marks entry assistant.

What it does
------------
1. Reads the marks workbook in this folder (supports the HTML ".xls" export
   CAMIS gives you, plus real .xls/.xlsx files).
2. Opens CAMIS in a real browser window (installed Chrome or Edge by default, or
   Playwright's Chromium) with its own persistent profile.
3. Waits while you log in and navigate to the marks-recording table.
4. Finds the right assessment column, matches students by name (fuzzy fallback),
   and types the marks into the table exactly like a human would.
5. Saves the marks, writes a report of everyone it could not enter, and leaves
   the browser open until YOU close it.

Safety defaults
---------------
* Empty cells are filled.
* Cells that already contain the same mark are left alone ("already correct").
* Cells that contain a DIFFERENT mark are reported as a conflict and left alone.
  Re-run with --overwrite if you really want to replace them.
* Cells that contain a missing-reason (ILLNESS, etc.) are left alone.
* Excel rows with an empty mark count as 0 and are entered as 0.

Requirements
------------
    pip install playwright beautifulsoup4
    (no browser install needed if Google Chrome or Microsoft Edge is installed;
     otherwise run `playwright install chromium`)

Usage
-----
    python camis_grader.py
    python camis_grader.py --file lesson.xls
    python camis_grader.py --assessment "Formative Assessment"
    python camis_grader.py --dry-run
    python camis_grader.py --overwrite
    python camis_grader.py --no-scale
    python camis_grader.py --factor 1
    python camis_grader.py --no-save

Notes
-----
* If the export has several marks columns (e.g. "Marks/10 pts" plus a converted
  "Marks /30"), the rightmost one is used. Override with --mark-column.
* Marks are still rescaled automatically when the chosen column's maximum and
  the site assessment maximum differ (e.g. 10 -> 30 means multiplied by 3).
* Some CAMIS exports are split: the .xls is only a frameset and the data lives
  in a sibling "<name>_files" folder. Keep both together, or flatten it first:
      python flatten_lesson.py lesson.xls -o lesson.xls
* The browser profile is kept in ./.camis-profile so you only log in once.
"""

from __future__ import annotations

import argparse
import csv
import html as html_module
import os
import re
import sys
import time
import traceback
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Optional
from urllib.parse import unquote

try:
    from bs4 import BeautifulSoup
except ImportError:  # pragma: no cover
    BeautifulSoup = None  # type: ignore[assignment]

try:
    from playwright.sync_api import (
        TimeoutError as PlaywrightTimeoutError,
        sync_playwright,
    )
except ImportError:  # pragma: no cover
    sync_playwright = None  # type: ignore[assignment]
    PlaywrightTimeoutError = Exception  # type: ignore[assignment]


DEFAULT_URL = "https://camis.nesa.gov.rw/"
DEFAULT_PROFILE = ".camis-profile"
META_KEYS = {
    "Subject",
    "Max Points",
    "Class/Term",
    "Teacher",
    "Entry Date",
    "Given Date",
}


# ---------------------------------------------------------------------------
# Data
# ---------------------------------------------------------------------------


@dataclass
class ExcelStudent:
    row_number: int
    student_id: str
    name: str
    raw_mark: str
    mark: Optional[float]


@dataclass
class MarksWorkbook:
    path: Path
    sheet_name: str
    meta: dict[str, str]
    excel_max: Optional[float]
    students: list[ExcelStudent]
    mark_column: str = ""
    warnings: list[str] = field(default_factory=list)


@dataclass
class SiteStudent:
    index: int
    name: str
    number: str
    cells: list[dict[str, Any]]


@dataclass
class ReportRecord:
    status: str
    reason: str = ""
    excel_row: Optional[int] = None
    excel_id: str = ""
    excel_name: str = ""
    excel_mark: str = ""
    target_mark: str = ""
    site_name: str = ""
    site_number: str = ""
    site_mark_before: str = ""


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------


def parse_number(value: Any) -> Optional[float]:
    if value is None:
        return None
    text = str(value).strip().replace(",", ".")
    if text == "":
        return None
    try:
        return float(text)
    except ValueError:
        return None


def format_number(value: float) -> str:
    value = round(value, 2)
    if abs(value - round(value)) < 1e-9:
        return str(int(round(value)))
    return f"{value:.2f}".rstrip("0").rstrip(".")


def normalize_name(value: str) -> str:
    value = unicodedata.normalize("NFKD", value or "")
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = value.upper()
    value = re.sub(r"[^A-Z0-9]+", " ", value)
    return " ".join(value.split())


def name_similarity(a: str, b: str) -> float:
    a = normalize_name(a)
    b = normalize_name(b)
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    direct = SequenceMatcher(None, a, b).ratio()
    sorted_a = " ".join(sorted(a.split()))
    sorted_b = " ".join(sorted(b.split()))
    token_sorted = SequenceMatcher(None, sorted_a, sorted_b).ratio()
    return max(direct, token_sorted)


def normalize_id(value: str) -> str:
    return re.sub(r"\D", "", value or "")


def id_keys(value: str) -> set[str]:
    digits = normalize_id(value)
    if not digits:
        return set()
    return {digits, digits.lstrip("0") or "0"}


def parse_cell_text(text: str) -> tuple[str, Any]:
    """Return (kind, value): number / empty / reason / other."""
    text = (text or "").strip()
    if not text:
        return "empty", None

    match = re.match(r"^\s*(-?\d+(?:\.\d+)?)\s*/\s*(\d+(?:\.\d+)?)", text)
    if match:
        return "number", float(match.group(1))

    if re.fullmatch(r"-?\d+(?:\.\d+)?", text):
        return "number", float(text)

    lowered = text.lower()
    if "add marks" in lowered or "no marks" in lowered:
        return "empty", None

    return "reason", text


def is_mark_header(value: str) -> bool:
    return bool(re.search(r"\bmarks?\b", value or "", re.IGNORECASE))


def is_name_header(value: str) -> bool:
    return (value or "").strip().lower() in {"name", "names", "student name", "student names"}


def pick_mark_column(headers: list[str], override: Optional[str] = None) -> int:
    """Index of the marks column to read.

    The CAMIS export can contain several marks columns (e.g. "Marks/10 pts"
    plus a hand-added "Marks /30"). By default the rightmost one is used,
    because that is the converted column added for the site scale.
    """
    candidates = [i for i, header in enumerate(headers) if is_mark_header(header)]
    if not candidates:
        raise RuntimeError(
            "Could not find a marks column in the header row: "
            + ", ".join(repr(header) for header in headers)
        )

    if override:
        wanted = override.strip()
        for index in candidates:
            if wanted.lower() in headers[index].lower():
                return index
        if wanted.isdigit():
            index = int(wanted) - 1
            if not 0 <= index < len(headers):
                raise RuntimeError(
                    f"--mark-column {wanted} is out of range (1-{len(headers)})."
                )
            return index
        raise RuntimeError(
            f"--mark-column {wanted!r} does not match any marks column "
            f"({', '.join(headers[i] for i in candidates)})."
        )

    with_max = [
        i for i in candidates if re.search(r"(\d+(?:[.,]\d+)?)", headers[i])
    ]
    return (with_max or candidates)[-1]


def column_max(header: str) -> Optional[float]:
    match = re.search(r"(\d+(?:[.,]\d+)?)", header or "")
    return parse_number(match.group(1)) if match else None


def mark_from_text(raw_mark: str) -> Optional[float]:
    """Missing marks count as 0; non-numeric text stays unknown."""
    text = " ".join((raw_mark or "").split())
    if text == "":
        return 0.0
    return parse_number(text)


def _html_links(raw: str) -> list[str]:
    pattern = re.compile(
        r'(?:<x:WorksheetSource[^>]*HRef|<frame[^>]*src|<link[^>]*href)\s*=\s*"([^"]+)"',
        re.IGNORECASE,
    )
    links = []
    for match in pattern.finditer(raw):
        href = html_module.unescape(match.group(1).strip())
        decoded = unquote(href)
        for candidate in (decoded, href):
            if candidate and candidate not in links:
                links.append(candidate)
    return links


def resolve_html_sheet(path: Path, raw: str) -> tuple[Path, str]:
    """Follow the Excel frameset / WorksheetSource link to the real sheet."""
    sheet_links = []
    for href in _html_links(raw):
        lowered = href.lower()
        if lowered.endswith((".css", ".xml", ".xls", ".xlsx")):
            continue
        if not lowered.endswith((".htm", ".html")):
            continue
        candidate = (path.parent / href).resolve()
        sheet_links.append(candidate)
        if candidate.is_file():
            return candidate, candidate.read_text(encoding="utf-8", errors="replace")

    wanted = sheet_links[0] if sheet_links else None
    if wanted is not None:
        raise RuntimeError(
            f"{path.name} is an Excel frameset that keeps its data in "
            f"{wanted.parent.name}/{wanted.name}, which is not here.\n"
            "Copy the WHOLE exported folder (the .xls file plus its "
            f"'{wanted.parent.name}' folder) next to this script, then run again."
        )
    return path, raw


def find_workbook_file(arg: Optional[str]) -> Path:
    if arg:
        path = Path(arg).expanduser()
        if not path.is_absolute():
            path = Path.cwd() / path
        if not path.exists():
            raise FileNotFoundError(f"Excel file not found: {path}")
        return path

    search_dirs = [Path.cwd(), Path(__file__).resolve().parent]
    candidates = sorted(
        {
            p
            for folder in search_dirs
            if folder.exists()
            for p in folder.iterdir()
            if p.is_file() and p.suffix.lower() in {".xls", ".xlsx", ".xlsm"}
        }
    )
    if not candidates:
        raise FileNotFoundError(
            "No .xls/.xlsx file found in this folder. Put the workbook here or pass --file."
        )
    if len(candidates) > 1:
        names = ", ".join(p.name for p in candidates)
        raise RuntimeError(
            f"Multiple workbooks found ({names}). Choose one with --file."
        )
    return candidates[0]


# ---------------------------------------------------------------------------
# Workbook parsing
# ---------------------------------------------------------------------------


def _meta_pair(tr: Any) -> tuple[str, str]:
    cells = tr.find_all(["td", "th"], recursive=False)
    if len(cells) != 2:
        return "", ""
    key = " ".join(cells[0].get_text(" ", strip=True).split())
    if key not in META_KEYS:
        return "", ""
    value = " ".join(cells[1].get_text(" ", strip=True).split())
    return key, value


def table_rows(table: Any) -> list[list[str]]:
    """Rows that belong to this table, ignoring rows of nested tables."""
    rows: list[list[str]] = []
    for tr in table.find_all("tr"):
        if tr.find_parent("table") is not table:
            continue
        cells = [
            " ".join(cell.get_text(" ", strip=True).split())
            for cell in tr.find_all(["td", "th"], recursive=False)
        ]
        if cells:
            rows.append(cells)
    return rows


def parse_html_workbook(
    path: Path,
    raw_bytes: bytes,
    mark_column: Optional[str] = None,
    sheet_name: str = "",
    display_path: Optional[Path] = None,
) -> MarksWorkbook:
    raw = raw_bytes.decode("utf-8", errors="replace")
    if BeautifulSoup is None:
        raise RuntimeError("beautifulsoup4 is required to read the HTML .xls export.")

    soup = BeautifulSoup(raw, "html.parser")

    if not sheet_name:
        name_match = re.search(r"<x:Name>(.*?)</x:Name>", raw, re.S | re.I)
        if name_match:
            sheet_name = html_module.unescape(name_match.group(1)).strip()

    rows: list[list[str]] | None = None
    for table in soup.find_all("table"):
        candidate = table_rows(table)
        for row in candidate:
            if any(is_name_header(cell) for cell in row) and any(
                is_mark_header(cell) for cell in row
            ):
                rows = candidate
                break
        if rows is not None:
            break

    if rows is None:
        sheet_path, sheet_raw = resolve_html_sheet(path, raw)
        if sheet_path != path:
            return parse_html_workbook(
                sheet_path,
                sheet_raw.encode("utf-8"),
                mark_column,
                sheet_name,
                display_path or path,
            )
        raise RuntimeError(
            f"Could not find an Id / Names / Marks table inside {path.name}."
        )

    meta = _meta_from_rows([tuple(row) for row in rows])
    for tr in soup.find_all("tr"):
        key, value = _meta_pair(tr)
        if key:
            meta[key] = value
    workbook = _rows_to_workbook(
        display_path or path, sheet_name, [tuple(row) for row in rows], meta, mark_column
    )
    return workbook


def _rows_to_workbook(
    path: Path,
    sheet_name: str,
    rows: list[tuple[Any, ...]],
    meta: dict[str, str],
    mark_column: Optional[str] = None,
) -> MarksWorkbook:
    header_index = None
    headers: list[str] = []
    for i, row in enumerate(rows):
        cells = ["" if cell is None else str(cell).strip() for cell in row]
        if any(is_name_header(cell) for cell in cells) and any(
            is_mark_header(cell) for cell in cells
        ):
            header_index = i
            headers = cells
            break

    if header_index is None:
        raise RuntimeError(f"Could not find a Names/Marks header row in sheet {sheet_name!r}.")

    name_index = next(i for i, h in enumerate(headers) if is_name_header(h))
    mark_index = pick_mark_column(headers, mark_column)
    id_index = next(
        (i for i, h in enumerate(headers) if h.strip().lower() in {"id", "student id", "student number"}),
        0,
    )

    excel_max = column_max(headers[mark_index])
    if excel_max is None:
        excel_max = parse_number(meta.get("Max Points"))

    students: list[ExcelStudent] = []
    for row in rows[header_index + 1 :]:
        cells = ["" if cell is None else str(cell).strip() for cell in row]
        if len(cells) <= max(id_index, name_index, mark_index):
            continue
        student_id = " ".join(cells[id_index].split())
        name = " ".join(cells[name_index].split())
        raw_mark = cells[mark_index]
        if not student_id and not name:
            continue
        students.append(
            ExcelStudent(
                row_number=len(students) + 1,
                student_id=student_id,
                name=name,
                raw_mark=raw_mark,
                mark=mark_from_text(raw_mark),
            )
        )

    if not students:
        detail = " ".join(
            f"{key}={meta[key]}"
            for key in ("Subject", "Class/Term", "Max Points")
            if meta.get(key)
        )
        raise RuntimeError(
            f"No student rows found in {sheet_name!r} - this export contains no "
            f"students at all ({detail or 'no metadata'}).\n"
            "CAMIS writes a sheet like this when no students are enrolled in the "
            "course/assessment yet, or when the export was generated before the "
            "marks-recording page had any rows. Re-download it from the page that "
            "actually lists the students, and check the column is named like "
            "'Marks/10 pts' in the header."
        )

    return MarksWorkbook(
        path=path,
        sheet_name=sheet_name,
        meta=meta,
        excel_max=excel_max,
        students=students,
        mark_column=headers[mark_index],
    )


def _meta_from_rows(rows: list[tuple[Any, ...]]) -> dict[str, str]:
    meta: dict[str, str] = {}
    for row in rows:
        cells = ["" if cell is None else str(cell).strip() for cell in row[:2]]
        if len(cells) == 2 and cells[0] in META_KEYS:
            meta[cells[0]] = cells[1]
    return meta


def parse_xlsx_workbook(path: Path, mark_column: Optional[str] = None) -> MarksWorkbook:
    try:
        import openpyxl
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("openpyxl is required to read .xlsx files.") from exc

    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    errors: list[str] = []
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        meta = _meta_from_rows(rows)
        try:
            return _rows_to_workbook(path, ws.title, rows, meta, mark_column)
        except RuntimeError as exc:
            errors.append(str(exc))
            continue
    raise RuntimeError("Could not find a marks table in any sheet: " + " | ".join(errors))


def parse_xls_workbook(path: Path, mark_column: Optional[str] = None) -> MarksWorkbook:
    try:
        import xlrd
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("xlrd is required to read binary .xls files.") from exc

    book = xlrd.open_workbook(str(path))
    errors: list[str] = []
    for sheet in book.sheets():
        rows = [tuple(sheet.row_values(r)) for r in range(sheet.nrows)]
        meta = _meta_from_rows(rows)
        try:
            return _rows_to_workbook(path, sheet.name, rows, meta, mark_column)
        except RuntimeError as exc:
            errors.append(str(exc))
            continue
    raise RuntimeError("Could not find a marks table in any sheet: " + " | ".join(errors))


def load_workbook(path: Path, mark_column: Optional[str] = None) -> MarksWorkbook:
    raw = path.read_bytes()
    head = raw[:4096].lstrip().lower()
    if head.startswith(b"<") or b"<table" in head or b"<html" in head:
        return parse_html_workbook(path, raw, mark_column)

    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        return parse_xlsx_workbook(path, mark_column)
    if suffix == ".xls":
        try:
            return parse_xls_workbook(path, mark_column)
        except RuntimeError:
            return parse_xlsx_workbook(path, mark_column)
    raise RuntimeError(f"Unsupported workbook type: {path.suffix}")


# ---------------------------------------------------------------------------
# Browser-side table extraction
# ---------------------------------------------------------------------------

EXTRACT_TABLE_JS = r"""
() => {
  const tables = [...document.querySelectorAll('table')];
  let table = null;
  let bestCount = -1;
  for (const candidate of tables) {
    const count = [...candidate.querySelectorAll('button[id^="cell"]')]
      .filter((button) => button.closest('table') === candidate).length;
    if (count > bestCount) {
      bestCount = count;
      table = candidate;
    }
  }
  if (!table || bestCount <= 0) return null;

  const headRow = table.querySelector('thead tr');
  if (!headRow) return null;

  const ths = [...headRow.querySelectorAll(':scope > th')];
  const headers = [];
  let column = 0;
  for (const th of ths) {
    const titleEl = th.querySelector('h4');
    const text = (th.innerText || '').trim();
    const maxMatch = text.match(/Max\s*:\s*([\d.]+)/i);
    headers.push({
      title: titleEl ? titleEl.innerText.trim() : text,
      tdIndex: column,
      max: maxMatch ? parseFloat(maxMatch[1]) : null,
    });
    column += th.colSpan || 1;
  }

  const rows = [];
  const trs = [...table.querySelectorAll('tbody > tr')];
  trs.forEach((tr, i) => {
    const tds = [...tr.querySelectorAll(':scope > td')];
    if (tds.length < 3) return;

    const isNumberSpan = (span) => /^[0-9][0-9\s-]*$/.test((span.textContent || '').trim());
    const numberTd = tds.find((td) => {
      if (td.querySelector('button[id^="cell"]')) return false;
      return [...td.querySelectorAll('span')].some(isNumberSpan);
    });
    const nameTd = numberTd || tds[2];
    const numberSpan = [...nameTd.querySelectorAll('span')].find(isNumberSpan);
    const number = numberSpan ? (numberSpan.textContent || '').trim() : '';

    let name = '';
    for (const node of nameTd.childNodes) {
      if (node.nodeType === Node.TEXT_NODE) name += node.textContent;
    }
    name = name.trim();
    if (!name) name = (nameTd.innerText || '').replace(number, '').trim();

    const cells = [];
    tds.forEach((td, idx) => {
      const btn = td.querySelector('button[id^="cell"]');
      cells.push({
        tdIndex: idx,
        buttonId: btn ? btn.id : null,
        text: btn ? (btn.innerText || '').trim() : (td.innerText || '').trim(),
      });
    });

    if (name || number) {
      rows.push({ rowIndex: i, name, number, cells });
    }
  });

  return { headers, rows, url: location.href };
}
"""

MARK_TABLE_JS = r"""
() => {
  const tables = [...document.querySelectorAll('table')];
  let best = null;
  let bestCount = -1;
  for (const candidate of tables) {
    const count = [...candidate.querySelectorAll('button[id^="cell"]')]
      .filter((button) => button.closest('table') === candidate).length;
    if (count > bestCount) {
      bestCount = count;
      best = candidate;
    }
  }

  document.querySelectorAll('[data-camis-marks-table]').forEach((element) => {
    element.removeAttribute('data-camis-marks-table');
  });

  if (!best || bestCount <= 0) return 0;
  best.setAttribute('data-camis-marks-table', '1');
  return bestCount;
}
"""


def wait_for_marks_page(
    ctx: Any, wait_minutes: float, settle_seconds: float = 2.5
) -> tuple[Any, dict[str, Any]]:
    deadline = time.time() + wait_minutes * 60
    last_note = time.time()
    candidate_page = None
    last_signature = None
    stable_since = 0.0

    print()
    print("Waiting for you to reach the marks-recording table...")
    print("  - Log in manually in the browser window.")
    print("  - Open the course / assessment page that shows the student table.")
    print("  - Do not close this terminal. This usually waits up to "
          f"{wait_minutes:g} minutes.")
    print()

    while time.time() < deadline:
        found = None
        for page in list(ctx.pages):
            if page.is_closed():
                continue
            try:
                info = page.evaluate(EXTRACT_TABLE_JS)
            except Exception:
                continue
            if not info:
                continue
            rows = [r for r in info.get("rows", []) if r.get("name") or r.get("number")]
            if not rows:
                continue
            if not any(cell.get("buttonId") for row in rows for cell in row.get("cells", [])):
                continue
            found = (page, {**info, "rows": rows})
            break

        if found is None:
            candidate_page = None
            last_signature = None
            stable_since = 0.0
        else:
            page, info = found
            signature = (
                page.url,
                len(info["rows"]),
                tuple(
                    (row.get("name"), row.get("number"), tuple(cell.get("text") for cell in row.get("cells", [])))
                    for row in info["rows"]
                ),
            )
            if candidate_page is not page or signature != last_signature:
                if candidate_page is not page:
                    print(f"Found the marks table: {page.url}")
                    print(f"  Students on page: {len(info['rows'])}")
                    print("  Waiting for the table to finish loading...")
                candidate_page = page
                last_signature = signature
                stable_since = time.time()
            elif time.time() - stable_since >= settle_seconds:
                print(f"Table is ready ({len(info['rows'])} students).")
                return page, info

        now = time.time()
        if now - last_note > 20:
            print("  still waiting...")
            last_note = now
        time.sleep(0.5)

    raise TimeoutError(
        "Timed out waiting for the marks table. Re-run when you are ready."
    )


# ---------------------------------------------------------------------------
# Matching
# ---------------------------------------------------------------------------


def match_students(
    excel_students: list[ExcelStudent],
    site_students: list[SiteStudent],
    min_similarity: float,
    margin: float,
) -> tuple[dict[int, int], dict[int, str]]:
    """Return (excel_index -> site_index, excel_index -> explanation)."""
    assigned_site: dict[int, int] = {}
    excel_to_site: dict[int, int] = {}
    notes: dict[int, str] = {}

    def assign(excel_index: int, site_index: int) -> None:
        excel_to_site[excel_index] = site_index
        assigned_site[site_index] = excel_index
        notes.pop(excel_index, None)

    # 1) Exact normalized name.
    for ei, student in enumerate(excel_students):
        key = normalize_name(student.name)
        if not key:
            continue
        candidates = [
            site
            for site in site_students
            if site.index not in assigned_site and normalize_name(site.name) == key
        ]
        if len(candidates) == 1:
            assign(ei, candidates[0].index)
        elif len(candidates) > 1:
            notes[ei] = f"{len(candidates)} students on the page have this exact name"

    # 2) Student number / ID, but only when the workbook IDs are trustworthy.
    id_use_count: dict[str, int] = {}
    for student in excel_students:
        for key in id_keys(student.student_id):
            id_use_count[key] = id_use_count.get(key, 0) + 1

    for ei, student in enumerate(excel_students):
        if ei in excel_to_site:
            continue
        wanted = id_keys(student.student_id)
        if not wanted or any(id_use_count.get(key, 0) > 1 for key in wanted):
            continue
        candidates = [
            site
            for site in site_students
            if site.index not in assigned_site and (id_keys(site.number) & wanted)
        ]
        if len(candidates) == 1:
            assign(ei, candidates[0].index)
        elif len(candidates) > 1:
            notes[ei] = f"{len(candidates)} students on the page have this ID"

    # 3) Fuzzy name, only when unambiguous.
    for ei, student in enumerate(excel_students):
        if ei in excel_to_site:
            continue
        scored = []
        for site in site_students:
            if site.index in assigned_site:
                continue
            scored.append((name_similarity(student.name, site.name), site))
        if not scored:
            notes[ei] = "no unmatched student left on the page"
            continue
        scored.sort(key=lambda item: item[0], reverse=True)
        best_score, best_site = scored[0]
        second_score = scored[1][0] if len(scored) > 1 else 0.0
        if best_score >= min_similarity and (
            second_score < min_similarity or best_score - second_score >= margin
        ):
            assign(ei, best_site.index)
        else:
            notes[ei] = (
                f"no reliable match (closest: {best_site.name!r} "
                f"{best_score:.2f}, next {second_score:.2f})"
            )

    return excel_to_site, notes


# ---------------------------------------------------------------------------
# Browser actions
# ---------------------------------------------------------------------------


def choose_assessment(
    headers: list[dict[str, Any]],
    assessment_arg: Optional[str],
    column_arg: Optional[int],
) -> dict[str, Any]:
    candidates = assessment_columns(headers)
    if not candidates:
        raise RuntimeError("No assessment columns were found in the table header.")

    if column_arg:
        if column_arg < 1 or column_arg > len(candidates):
            raise RuntimeError(
                f"--column {column_arg} is out of range; found {len(candidates)} assessment column(s)."
            )
        return candidates[column_arg - 1]

    if assessment_arg:
        needle = assessment_arg.strip().lower()
        matches = [h for h in candidates if needle in h["title"].lower()]
        if len(matches) == 1:
            return matches[0]
        if not matches:
            available = ", ".join(repr(h["title"]) for h in candidates)
            raise RuntimeError(
                f"No assessment column matched {assessment_arg!r}. Available: {available}"
            )
        raise RuntimeError(
            f"{len(matches)} assessment columns matched {assessment_arg!r}; use --column."
        )

    if len(candidates) == 1:
        return candidates[0]

    if not sys.stdin.isatty():
        available = ", ".join(
            f"{i}. {h['title']} (max {h.get('max')})"
            for i, h in enumerate(candidates, 1)
        )
        raise RuntimeError(
            "This page has several assessment columns, and no choice was given. "
            "Fill in the 'Assessment column' box with one of these and run again:\n"
            + available
        )

    print()
    print("This page has several assessment columns:")
    for i, header in enumerate(candidates, start=1):
        print(f"  {i}. {header['title']}  (max {header.get('max')})")
    while True:
        raw = input("Choose the column to fill [1]: ").strip() or "1"
        try:
            choice = int(raw)
        except ValueError:
            print("Enter a number.")
            continue
        if 1 <= choice <= len(candidates):
            return candidates[choice - 1]
        print("That number is out of range.")


def get_site_max(target: dict[str, Any], site_students: list[SiteStudent]) -> Optional[float]:
    if target.get("max"):
        return float(target["max"])
    td_index = target["tdIndex"]
    for student in site_students:
        cell = next((c for c in student.cells if c["tdIndex"] == td_index), None)
        if not cell:
            continue
        match = re.search(r"/\s*(\d+(?:\.\d+)?)", cell.get("text", ""))
        if match:
            return float(match.group(1))
    return None


def click_first_visible(page: Any, pattern: re.Pattern[str], exclude: tuple[str, ...] = ()) -> Any:
    locator = page.locator("button:visible")
    count = locator.count()
    for i in range(count):
        button = locator.nth(i)
        try:
            text = " ".join(button.inner_text().split())
            lowered = text.lower()
            if not pattern.search(text):
                continue
            if any(word in lowered for word in exclude):
                continue
            if not button.is_enabled():
                continue
            return button
        except Exception:
            continue
    return None


def enter_mark(page: Any, row_index: int, td_index: int, value: str) -> None:
    marked = page.evaluate(MARK_TABLE_JS)
    if not marked:
        raise RuntimeError("marks table is no longer on the page")
    table = page.locator("table[data-camis-marks-table]").first
    row = table.locator("tbody > tr").nth(row_index)
    cell = row.locator("td").nth(td_index)

    button = cell.locator("button[id^='cell']")
    if button.count() == 0:
        raise RuntimeError("this cell is not editable (approved/locked?)")

    button.first.click(timeout=10_000)
    number_input = cell.locator('input[type="number"]')
    number_input.wait_for(state="visible", timeout=6000)

    number_input.fill(value)
    page.wait_for_timeout(150)
    number_input.blur()
    page.wait_for_timeout(150)

    try:
        current = number_input.input_value(timeout=300)
    except Exception:
        current = None
    if current is not None and current.strip() != value:
        raise RuntimeError(f"typed {value!r} but the field shows {current!r}")


def save_marks(page: Any) -> tuple[str, str]:
    button = None
    for _ in range(12):
        button = click_first_visible(
            page,
            re.compile(r"^Save(\s+marks)?\b", re.IGNORECASE),
            exclude=("change", "draft", "template"),
        )
        if button is not None:
            break
        page.wait_for_timeout(500)
    if button is None:
        return "not-found", "Could not find an enabled Save marks button"

    try:
        with page.expect_response(
            lambda response: (
                response.request.method == "POST"
                and "/marks-record/" in response.url
            ),
            timeout=90_000,
        ) as response_info:
            button.click()
        response = response_info.value
        page.wait_for_timeout(800)
        status = f"HTTP {response.status}"
        return ("ok" if response.ok else "failed"), status
    except PlaywrightTimeoutError:
        return "unknown", "Save was clicked, but no save request was observed"
    except Exception as exc:
        return "failed", f"Save failed: {exc}"


# ---------------------------------------------------------------------------
# Planning (shared by the CLI and the desktop app)
# ---------------------------------------------------------------------------


def site_students_from(table_info: dict[str, Any]) -> list[SiteStudent]:
    return [
        SiteStudent(
            index=row["rowIndex"],
            name=row["name"],
            number=row["number"],
            cells=row["cells"],
        )
        for row in table_info["rows"]
    ]


def assessment_columns(headers: list[dict[str, Any]]) -> list[dict[str, Any]]:
    known_labels = {"no", "student number", "student id"}
    return [
        header
        for header in headers
        if header.get("title") and header["title"].strip().lower() not in known_labels
    ]


def compute_factor(excel_max: Optional[float], site_max: Optional[float]) -> float:
    if excel_max and site_max and abs(excel_max - site_max) > 1e-9:
        return site_max / excel_max
    return 1.0


@dataclass
class Plan:
    records: list[ReportRecord]
    site_only: list[ReportRecord]
    to_enter: list[tuple[ReportRecord, SiteStudent]]
    target: dict[str, Any]
    site_max: Optional[float]
    factor: float
    matched: int


def plan_entries(
    workbook: MarksWorkbook,
    site_students: list[SiteStudent],
    target: dict[str, Any],
    *,
    overwrite: bool = False,
    factor: Optional[float] = None,
    min_similarity: float = 0.80,
    margin: float = 0.08,
) -> Plan:
    """Decide what to do for every student without touching the page.

    Records that should be typed get status "to-enter" and are listed in
    ``plan.to_enter``; everything else already has its final status.
    """
    site_by_index = {student.index: student for student in site_students}
    site_max = get_site_max(target, site_students)
    if factor is None:
        factor = compute_factor(workbook.excel_max, site_max)

    excel_to_site, notes = match_students(
        workbook.students, site_students, min_similarity=min_similarity, margin=margin
    )
    matched_site_indexes = set(excel_to_site.values())
    td_index = target["tdIndex"]
    records: list[ReportRecord] = []
    to_enter: list[tuple[ReportRecord, SiteStudent]] = []

    for ei, student in enumerate(workbook.students):
        record = ReportRecord(
            status="not-entered",
            excel_row=student.row_number,
            excel_id=student.student_id,
            excel_name=student.name,
            excel_mark=student.raw_mark or "0 (blank)",
        )
        records.append(record)
        if student.mark is not None and student.mark >= 0:
            # Shown in "do these by hand" lists, already on the CAMIS scale.
            record.target_mark = format_number(student.mark * factor)

        site_index = excel_to_site.get(ei)
        site = site_by_index.get(site_index) if site_index is not None else None
        if site is None:
            record.status = "not-found"
            record.reason = notes.get(ei, "no matching student on the page")
            continue
        record.site_name = site.name
        record.site_number = site.number
        cell = next((c for c in site.cells if c["tdIndex"] == td_index), None)
        if cell is None:
            record.status = "not-editable"
            record.reason = "assessment cell not found in this row"
            continue
        record.site_mark_before = cell["text"]

        if student.mark is None:
            record.status = "no-excel-mark"
            record.reason = "the Excel mark is not a number"
            continue
        if student.mark < 0:
            record.status = "invalid-excel-mark"
            record.reason = "negative mark in Excel"
            continue
        if workbook.excel_max and student.mark > workbook.excel_max + 1e-9:
            record.status = "invalid-excel-mark"
            record.reason = f"mark is above the Excel maximum ({workbook.excel_max:g})"
            continue

        target_value = student.mark * factor
        target_text = format_number(target_value)
        record.target_mark = target_text

        if site_max is not None and target_value > site_max + 1e-9:
            record.status = "out-of-range"
            record.reason = f"converted mark {target_text} is above the site maximum {site_max:g}"
            continue
        if not cell.get("buttonId"):
            record.status = "not-editable"
            record.reason = "this cell is not editable (approved/locked?)"
            continue

        kind, existing = parse_cell_text(cell["text"])
        if kind == "reason":
            record.status = "skipped-reason"
            record.reason = f"cell already has a missing-reason ({existing}); clear it manually"
            continue
        if kind == "number":
            if abs(float(existing) - target_value) <= 1e-9:
                record.status = "already-correct"
                continue
            if not overwrite:
                record.status = "conflict"
                record.reason = f"camis has {format_number(float(existing))}, your file has {target_text}"
                continue
            record.reason = f"replaced {format_number(float(existing))}"

        record.status = "to-enter"
        to_enter.append((record, site))

    site_only = [
        ReportRecord(
            status="site-not-in-excel",
            site_name=site.name,
            site_number=site.number,
            site_mark_before=next(
                (c["text"] for c in site.cells if c["tdIndex"] == td_index), ""
            ),
            reason="student appears on the page but has no row in the Excel file",
        )
        for site in site_students
        if site.index not in matched_site_indexes
    ]

    return Plan(
        records=records,
        site_only=site_only,
        to_enter=to_enter,
        target=target,
        site_max=site_max,
        factor=factor,
        matched=len(excel_to_site),
    )


# ---------------------------------------------------------------------------
# Reporting
# ---------------------------------------------------------------------------


STATUS_LABELS = {
    "entered": "ENTERED",
    "already-correct": "ALREADY CORRECT",
    "conflict": "CONFLICT - not overwritten",
    "skipped-reason": "HAS MISSING-REASON - not overwritten",
    "not-found": "NOT FOUND ON PAGE",
    "ambiguous": "AMBIGUOUS MATCH",
    "no-excel-mark": "MARK UNREADABLE IN EXCEL",
    "invalid-excel-mark": "INVALID EXCEL MARK",
    "out-of-range": "ABOVE SITE MAXIMUM",
    "not-editable": "NOT EDITABLE",
    "failed": "FAILED",
    "site-not-in-excel": "ON PAGE, NOT IN EXCEL",
    "not-entered": "NOT ENTERED",
}


def write_report(
    workbook: MarksWorkbook,
    records: list[ReportRecord],
    site_only: list[ReportRecord],
    target: dict[str, Any],
    site_max: Optional[float],
    factor: float,
    save_status: str,
    save_detail: str,
    site_url: str,
    site_count: int,
    out_dir: Optional[Path] = None,
) -> tuple[Path, Path]:
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    out_dir = out_dir or Path.cwd()
    out_dir.mkdir(parents=True, exist_ok=True)
    txt_path = out_dir / f"camis-report-{stamp}.txt"
    csv_path = out_dir / f"camis-report-{stamp}.csv"

    counts: dict[str, int] = {}
    for record in records + site_only:
        counts[record.status] = counts.get(record.status, 0) + 1

    lines: list[str] = []
    lines.append("CAMIS MARKS ENTRY REPORT")
    lines.append("=" * 78)
    lines.append(f"Generated       : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"Excel file      : {workbook.path.name}")
    lines.append(f"Worksheet       : {workbook.sheet_name or '(unknown)'}")
    lines.append(f"Subject         : {workbook.meta.get('Subject', '(unknown)')}")
    lines.append(f"Class/Term      : {workbook.meta.get('Class/Term', '(unknown)')}")
    lines.append(f"Assessment col  : {target['title']} (site max {site_max if site_max is not None else '?'})")
    lines.append(
        f"Excel column    : {workbook.mark_column or '(unknown)'} (max "
        f"{workbook.excel_max if workbook.excel_max is not None else '?'})"
        f"  ->  scale factor {factor:g}"
    )
    lines.append(f"Site URL        : {site_url}")
    lines.append(f"Excel students  : {len(workbook.students)}")
    lines.append(f"Students on page: {site_count}")
    lines.append(f"Save            : {save_status} ({save_detail})")
    if save_status == "dry-run":
        lines.append("Mode            : DRY RUN - nothing was typed or saved")
    lines.append("")
    lines.append("SUMMARY")
    lines.append("-" * 78)
    for status, label in STATUS_LABELS.items():
        if status in counts:
            lines.append(f"  {label:38s} {counts[status]:d}")
    lines.append("")

    attention = [
        r
        for r in records + site_only
        if r.status
        in {
            "not-found",
            "ambiguous",
            "conflict",
            "skipped-reason",
            "invalid-excel-mark",
            "out-of-range",
            "not-editable",
            "failed",
            "no-excel-mark",
            "site-not-in-excel",
        }
    ]

    lines.append("STUDENTS THAT NEED YOUR ATTENTION (do these manually)")
    lines.append("-" * 78)
    if not attention:
        lines.append("  None. Everything that could be entered was entered.")
    else:
        for record in attention:
            if record.status == "site-not-in-excel":
                lines.append(
                    f"  [{STATUS_LABELS[record.status]}] "
                    f"{record.site_name} | ID: {record.site_number} | "
                    f"current cell: {record.site_mark_before or '(empty)'} | {record.reason}"
                )
            else:
                lines.append(
                    f"  [{STATUS_LABELS.get(record.status, record.status)}] "
                    f"Row {record.excel_row} | {record.excel_name} | "
                    f"Excel ID: {record.excel_id} | Excel mark: {record.excel_mark or '(blank)'} | "
                    f"page: {record.site_name or '(no match)'} | {record.reason}"
                )
    lines.append("")

    entered = [r for r in records if r.status == "entered"]
    lines.append("MARKS ENTERED")
    lines.append("-" * 78)
    if not entered:
        lines.append("  None needed entering.")
    else:
        for record in entered:
            suffix = f"  [{record.reason}]" if record.reason else ""
            lines.append(
                f"  Excel: {record.excel_name}  ->  page: {record.site_name} | "
                f"ID: {record.site_number} | Excel {record.excel_mark} -> "
                f"{record.target_mark}{suffix}"
            )
    lines.append("")

    correct = [r for r in records if r.status == "already-correct"]
    lines.append("ALREADY CORRECT (left untouched)")
    lines.append("-" * 78)
    if not correct:
        lines.append("  None.")
    else:
        for record in correct:
            lines.append(
                f"  Excel: {record.excel_name}  ->  page: {record.site_name} | "
                f"ID: {record.site_number} | {record.target_mark}"
            )
    lines.append("")

    txt_path.write_text("\n".join(lines), encoding="utf-8")

    with csv_path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(
            [
                "status",
                "excel_row",
                "excel_student_id",
                "excel_name",
                "excel_mark",
                "target_mark",
                "site_name",
                "site_student_number",
                "site_mark_before",
                "reason",
            ]
        )
        for record in records + site_only:
            writer.writerow(
                [
                    record.status,
                    record.excel_row or "",
                    record.excel_id,
                    record.excel_name,
                    record.excel_mark,
                    record.target_mark,
                    record.site_name,
                    record.site_number,
                    record.site_mark_before,
                    record.reason,
                ]
            )

    return txt_path, csv_path


# ---------------------------------------------------------------------------
# Browser selection
# ---------------------------------------------------------------------------

CHROME_PATHS = [
    Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Google/Chrome/Application/chrome.exe",
    Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "Google/Chrome/Application/chrome.exe",
    Path.home() / "AppData/Local/Google/Chrome/Application/chrome.exe",
    Path.home() / "AppData/Local/Google/Chromium/Application/chrome.exe",
    Path("/usr/bin/google-chrome"),
    Path("/usr/bin/chromium"),
    Path("/usr/bin/chromium-browser"),
]
EDGE_PATHS = [
    Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Microsoft/Edge/Application/msedge.exe",
    Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "Microsoft/Edge/Application/msedge.exe",
    Path.home() / "AppData/Local/Microsoft/Edge/Application/msedge.exe",
]

CHANNELS = {
    "chrome": ("Google Chrome", CHROME_PATHS),
    "msedge": ("Microsoft Edge", EDGE_PATHS),
}


def channel_candidates(preference: str) -> list[Optional[str]]:
    preference = (preference or "auto").lower()
    if preference == "chromium":
        return [None]
    if preference in CHANNELS:
        return [preference, None]
    candidates = [
        channel
        for channel, (_, paths) in CHANNELS.items()
        if any(path.is_file() for path in paths)
    ]
    candidates.append(None)  # Playwright's bundled Chromium as last resort
    return candidates


def launch_browser(playwright: Any, profile_dir: Path, preference: str = "auto"):
    """Launch the CAMIS browser window, preferring installed Chrome/Edge.

    Always uses its own profile folder (never the teacher's personal profile),
    because automation cannot lock a profile that is in use and CAMIS only needs
    its own cookies.
    """
    errors: list[str] = []
    for channel in channel_candidates(preference):
        if channel is None:
            label = "Playwright Chromium"
        else:
            label = CHANNELS[channel][0]
        kwargs: dict[str, Any] = dict(
            user_data_dir=str(profile_dir),
            headless=False,
            viewport=None,
            args=["--start-maximized"],
            accept_downloads=True,
        )
        if channel:
            kwargs["channel"] = channel
        try:
            ctx = playwright.chromium.launch_persistent_context(**kwargs)
        except Exception as exc:  # noqa: BLE001 - try the next browser
            errors.append(f"{label}: {exc}")
            print(f"  ! Could not launch {label} ({exc}); trying another browser...")
            continue

        if channel is None:
            print("  Browser: Playwright Chromium  (`playwright install chromium` may be required)")
        else:
            print(f"  Browser: {label} (uses its own CAMIS profile folder; your personal {label} profile is not touched)")
        return ctx

    hint = "\n".join(f"  - {e}" for e in errors)
    raise RuntimeError(
        "Could not open a browser. Make sure Google Chrome or Microsoft Edge is installed,\n"
        "or run `python -m playwright install chromium`. Details:\n" + hint
    )


# ---------------------------------------------------------------------------
# Main flow
# ---------------------------------------------------------------------------


def run(args: argparse.Namespace) -> None:
    save_enabled = getattr(args, "save", True)
    workbook = load_workbook(find_workbook_file(args.file), args.mark_column)
    print(f"Excel file : {workbook.path.name}")
    print(f"Worksheet  : {workbook.sheet_name or '(unknown)'}")
    print(f"Subject    : {workbook.meta.get('Subject', '(unknown)')}")
    print(f"Students   : {len(workbook.students)}")
    print(
        f"Marks column: {workbook.mark_column or '(unknown)'} "
        f"(max {workbook.excel_max if workbook.excel_max is not None else '?'})"
    )
    blank_marks = sum(
        1 for student in workbook.students if student.mark == 0 and not student.raw_mark.strip()
    )
    if blank_marks:
        print(f"  {blank_marks} student(s) have a blank mark -> entered as 0.")
    meta_max = parse_number(workbook.meta.get("Max Points"))
    if (
        workbook.excel_max is not None
        and meta_max is not None
        and abs(meta_max - workbook.excel_max) > 1e-9
    ):
        print(
            f"  note: the sheet metadata says Max Points={meta_max:g}, but the marks "
            f"column header says {workbook.mark_column!r}. Using it as-is."
        )

    profile_dir = Path(args.profile).expanduser()
    if not profile_dir.is_absolute():
        profile_dir = Path.cwd() / profile_dir
    profile_dir.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as playwright:
        ctx = launch_browser(playwright, profile_dir, args.browser)
        try:
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            page.set_default_timeout(8_000)
            page.goto(args.url, wait_until="domcontentloaded", timeout=90_000)

            page, table_info = wait_for_marks_page(ctx, args.wait_minutes)
            site_url = table_info.get("url", page.url)

            target = choose_assessment(
                table_info["headers"], args.assessment, args.column
            )
            site_students = site_students_from(table_info)
            if args.factor is not None:
                factor = args.factor
            elif args.no_scale:
                factor = 1.0
            else:
                factor = None

            plan = plan_entries(
                workbook,
                site_students,
                target,
                overwrite=args.overwrite,
                factor=factor,
                min_similarity=args.min_similarity,
                margin=args.match_margin,
            )
            records, site_only = plan.records, plan.site_only
            site_max, factor = plan.site_max, plan.factor

            print()
            print(f"Assessment column : {target['title']}")
            print(f"Site maximum      : {site_max if site_max is not None else '?'}")
            print(f"Scale factor      : {factor:g}")
            if factor != 1:
                print(
                    "  Marks from the workbook will be multiplied by "
                    f"{factor:g} to match the site maximum."
                )
            print(f"Matched on page   : {plan.matched} / {len(workbook.students)}")

            td_index = target["tdIndex"]
            entered_count = 0
            for record, site in plan.to_enter:
                if args.dry_run:
                    record.status = "entered"
                    record.reason = "DRY RUN - not actually typed"
                    continue
                try:
                    enter_mark(page, site.index, td_index, record.target_mark)
                    record.status = "entered"
                    entered_count += 1
                    print(f"  entered {site.name}: {record.target_mark}")
                except Exception as exc:
                    record.status = "failed"
                    record.reason = str(exc).splitlines()[0]
                    print(f"  FAILED  {site.name}: {record.reason}")
            for record in records:
                if record.status == "conflict":
                    record.reason += " (use --overwrite to replace)"

            if entered_count > 0 and not args.dry_run and save_enabled:
                print()
                print("Saving marks...")
                save_status, save_detail = save_marks(page)
                print(f"  save: {save_status} ({save_detail})")
            elif args.dry_run:
                save_status, save_detail = "dry-run", "no changes written to the site"
            elif not save_enabled:
                save_status, save_detail = "not-requested", "--no-save was used"
            else:
                save_status, save_detail = "not-needed", "no new marks were entered"

            txt_path, csv_path = write_report(
                workbook=workbook,
                records=records,
                site_only=site_only,
                target=target,
                site_max=site_max,
                factor=factor,
                save_status=save_status,
                save_detail=save_detail,
                site_url=site_url,
                site_count=len(site_students),
            )

            counts: dict[str, int] = {}
            for record in records + site_only:
                counts[record.status] = counts.get(record.status, 0) + 1

            print()
            print("=" * 78)
            print("DONE" + (" (DRY RUN - nothing was typed or saved)" if args.dry_run else ""))
            print("=" * 78)
            summary_order = [
                "entered",
                "already-correct",
                "conflict",
                "not-found",
                "no-excel-mark",
                "site-not-in-excel",
                "skipped-reason",
                "invalid-excel-mark",
                "out-of-range",
                "not-editable",
                "failed",
            ]
            for status in summary_order:
                if counts.get(status):
                    print(f"{STATUS_LABELS.get(status, status):<18}: {counts[status]}")
            print(f"Report            : {txt_path}")
            print(f"CSV               : {csv_path}")

            attention = [
                record
                for record in records + site_only
                if record.status
                in {
                    "not-found",
                    "ambiguous",
                    "conflict",
                    "skipped-reason",
                    "invalid-excel-mark",
                    "out-of-range",
                    "not-editable",
                    "failed",
                    "no-excel-mark",
                    "site-not-in-excel",
                }
            ]
            if attention:
                print()
                print("Needs manual attention:")
                for record in attention[:20]:
                    if record.status == "site-not-in-excel":
                        print(f"  - [on page, not Excel] {record.site_name} ({record.site_number})")
                    else:
                        print(
                            f"  - [{record.status}] {record.excel_name} "
                            f"({record.excel_id or 'no ID'})"
                        )
                if len(attention) > 20:
                    print(f"  ... and {len(attention) - 20} more in the report.")

            if entered_count > 0 and not save_enabled and not args.dry_run:
                print()
                print(
                    "Marks were typed but NOT saved (--no-save). "
                    "Click Save marks in the browser when you are ready."
                )

            print()
            print("Browser is still open. Close the browser window when you are done.")
            keep_browser_open(ctx)

        except KeyboardInterrupt:
            print("\nInterrupted.")
            keep_browser_open(ctx)
        except Exception:
            print("\nSomething went wrong:")
            traceback.print_exc()
            print("\nThe browser is being left open so you can finish manually.")
            keep_browser_open(ctx)


def keep_browser_open(ctx: Any) -> None:
    state = {"closed": False}

    def mark_closed(*_args: Any) -> None:
        state["closed"] = True

    try:
        ctx.on("close", mark_closed)
    except Exception:
        pass

    try:
        while not state["closed"]:
            try:
                _ = ctx.pages
            except Exception:
                break
            time.sleep(1)
    except KeyboardInterrupt:
        return


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Read marks from an Excel export and enter them into CAMIS.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("--file", help="Workbook to read (default: the only .xls/.xlsx in this folder).")
    parser.add_argument("--url", default=DEFAULT_URL, help=f"Site URL (default: {DEFAULT_URL}).")
    parser.add_argument(
        "--browser",
        default="auto",
        choices=["auto", "chrome", "msedge", "chromium"],
        help="Browser to use (default: auto — Google Chrome or Microsoft Edge if "
        "installed, otherwise Playwright's Chromium).",
    )
    parser.add_argument("--profile", default=DEFAULT_PROFILE, help="Persistent browser profile folder.")
    parser.add_argument(
        "--assessment",
        help="Assessment column title or part of it (e.g. \"Formative Assessment\").",
    )
    parser.add_argument(
        "--column",
        type=int,
        help="1-based assessment column number, for pages with several assessments.",
    )
    parser.add_argument(
        "--mark-column",
        help="Workbook marks column to use: a header fragment (e.g. \"30\") or a "
        "1-based column number (default: the rightmost marks column).",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Replace existing marks that differ from the workbook.",
    )
    parser.add_argument(
        "--no-scale",
        action="store_true",
        help="Do not scale marks when Excel and site maximums differ.",
    )
    parser.add_argument(
        "--factor",
        type=float,
        help="Use a specific scale factor instead of the detected one.",
    )
    parser.add_argument(
        "--min-similarity",
        type=float,
        default=0.80,
        help="Minimum fuzzy name similarity for an automatic match (default: 0.80).",
    )
    parser.add_argument(
        "--match-margin",
        type=float,
        default=0.08,
        help="Required gap between the best and second-best fuzzy match (default: 0.08).",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Match and report, but do not type or save anything.",
    )
    parser.add_argument(
        "--no-save",
        action="store_false",
        dest="save",
        default=True,
        help="Type the marks but leave Save for you to click.",
    )
    parser.add_argument(
        "--wait-minutes",
        type=float,
        default=60,
        help="How long to wait for you to reach the marks page (default: 60).",
    )
    return parser


def main() -> int:
    if sync_playwright is None:
        print("Playwright is not installed. Run: pip install playwright", file=sys.stderr)
        return 1

    parser = build_parser()
    args = parser.parse_args()
    try:
        run(args)
    except FileNotFoundError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
