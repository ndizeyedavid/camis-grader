#!/usr/bin/env python3
"""
Flatten a CAMIS marks export into a single self-contained .xls file.

Some CAMIS downloads are "Excel frameset" .xls files: the .xls only holds a
frameset that points at a sibling folder (<name>_files/sheet001.htm). If that
folder is lost, renamed, or not copied, the export looks empty. This tool:

1. Reads the export (frameset, plain HTML .xls, binary .xls/.xlsx).
2. Picks the marks column the grader would use (rightmost marks column).
3. Writes ONE flat HTML .xls with the metadata, the header, and every student,
   so camis_grader.py can read it with no helper folder.

Usage
-----
    python flatten_lesson.py lesson.xls
    python flatten_lesson.py lesson.xls -o lesson_flat.xls
    python flatten_lesson.py "C:\\Users\\ME\\Downloads\\camis\\fwdfyi"
    python flatten_lesson.py lesson.xls --mark-column "30"
    python flatten_lesson.py --list "C:\\Users\\ME\\Downloads\\camis\\fwdfyi"

The input can be one file or a folder; a folder processes every .xls/.xlsx in it.
"""

from __future__ import annotations

import argparse
import html
import sys
from pathlib import Path

import camis_grader as cg

EXPORT_SUFFIXES = {".xls", ".xlsx", ".xlsm"}


def collect_inputs(target: Path) -> list[Path]:
    if target.is_dir():
        return sorted(
            p for p in target.iterdir() if p.is_file() and p.suffix.lower() in EXPORT_SUFFIXES
        )
    return [target]


def render_flat_xls(workbook: cg.MarksWorkbook) -> str:
    def row(*cells: str) -> str:
        tds = "".join(
            f'<td class="v">{html.escape(str(cell))}</td>' for cell in cells
        )
        return f"<tr>{tds}</tr>"

    header = workbook.mark_column or "Marks"
    max_points = (
        cg.format_number(workbook.excel_max) if workbook.excel_max is not None else ""
    )

    lines = [
        "<html>",
        "<head><meta charset=\"utf-8\">",
        "<style>td.v { mso-number-format:\"\\@\"; }</style>",
        "</head>",
        "<body>",
        "<table border=1>",
    ]
    for key in ("Subject", "Max Points", "Class/Term", "Teacher", "Entry Date", "Given Date"):
        value = workbook.meta.get(key)
        if value:
            lines.append(row(key, value))
    if max_points:
        lines.append(row("Max Points", max_points))
    lines.append(row("Id", "Names", header))
    for student in workbook.students:
        lines.append(row(student.student_id, student.name, student.raw_mark))
    lines += ["</table>", "</body>", "</html>"]
    return "\n".join(lines)


def default_output(source: Path) -> Path:
    return source.with_name(f"{source.stem}_flat.xls")


def flatten(source: Path, output: Path, mark_column: str | None) -> cg.MarksWorkbook:
    workbook = cg.load_workbook(source, mark_column)
    output.write_text(render_flat_xls(workbook), encoding="utf-8", newline="\n")
    return workbook


def describe(workbook: cg.MarksWorkbook) -> str:
    marks = [s.mark for s in workbook.students if s.mark is not None]
    blanks = sum(1 for s in workbook.students if not s.raw_mark.strip())
    span = f"{min(marks):g}-{max(marks):g}" if marks else "n/a"
    return (
        f"{workbook.meta.get('Subject', '?')} | "
        f"{workbook.meta.get('Class/Term', '?')} | "
        f"{len(workbook.students)} students | "
        f"column {workbook.mark_column or '?'} (max "
        f"{workbook.excel_max if workbook.excel_max is not None else '?'}) | "
        f"marks {span} | blank->0: {blanks}"
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Flatten a CAMIS marks export into one self-contained .xls.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("input", help="Export file or a folder containing exports.")
    parser.add_argument(
        "-o",
        "--output",
        help="Output file (only valid for a single input). Default: <name>_flat.xls",
    )
    parser.add_argument(
        "--mark-column",
        help="Marks column to keep: a header fragment (e.g. \"30\") or 1-based number.",
    )
    parser.add_argument(
        "--list",
        action="store_true",
        help="Only inspect inputs and print a summary; write nothing.",
    )
    args = parser.parse_args(argv)

    target = Path(args.input).expanduser()
    if not target.exists():
        print(f"Not found: {target}", file=sys.stderr)
        return 2

    inputs = collect_inputs(target)
    if not inputs:
        print(f"No .xls/.xlsx files found in {target}", file=sys.stderr)
        return 2
    if args.output and len(inputs) > 1:
        print("-o only works with a single input file.", file=sys.stderr)
        return 2

    failures = 0
    for source in inputs:
        print(f"\n{source.name}")
        try:
            if args.list:
                workbook = cg.load_workbook(source, args.mark_column)
                print(f"  {describe(workbook)}")
                continue

            output = Path(args.output) if args.output else default_output(source)
            workbook = flatten(source, output, args.mark_column)
            print(f"  {describe(workbook)}")
            print(f"  wrote {output}")

            reread = cg.load_workbook(output)
            if len(reread.students) != len(workbook.students):
                raise RuntimeError(
                    f"verification failed: {len(reread.students)} rows reread"
                )
            print(f"  verified: {len(reread.students)} rows readable")
        except Exception as exc:
            failures += 1
            print(f"  FAILED: {exc}")
            hint = str(exc).lower()
            if "empty" in hint:
                print("  tip: this export has no students yet; nothing to convert.")
            elif "frameset" in hint:
                print(
                    "  tip: pass the .xls from the SAME folder as its "
                    "'..._files' helper folder, or copy both here."
                )

    print()
    print(f"Done: {len(inputs) - failures} ok, {failures} failed.")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
