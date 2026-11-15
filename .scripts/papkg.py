#!/usr/bin/env python3
"""Batch convert each .xlsx, .xls, or .txt vocabulary file into an Anki .apkg deck.

Examples:
  python .scripts/papkg.py
  python .scripts/papkg.py --root . --output-dir .anki
  python .scripts/papkg.py --root . --force

Default behavior:
  - recursively scan the project root for .xlsx, .xls, and .txt files
  - generate one deck per source file
  - save the resulting .apkg in .anki while preserving the relative directory structure
"""

from __future__ import annotations

import argparse
import hashlib
import html
import re
import subprocess
import sys
from pathlib import Path
from typing import Iterable, List, Sequence


def ensure_dependencies() -> None:
    """Install the minimum dependency set needed for parsing XLSX and writing APKG."""
    missing: list[str] = []
    for pkg in ("openpyxl", "xlrd", "genanki"):
        try:
            __import__(pkg)
        except ModuleNotFoundError:
            missing.append(pkg)

    if missing:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "--quiet", *missing],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )


ensure_dependencies()

from openpyxl import load_workbook  # type: ignore  # noqa: E402
import xlrd  # type: ignore  # noqa: E402
import genanki  # type: ignore  # noqa: E402


def safe_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return text.replace("\r\n", "\n").replace("\r", "\n")


def normalize_label(label: str) -> str:
    label = safe_text(label).lower()
    return re.sub(r"[\s\-_]+", "", label)


def sanitize_name(value: str) -> str:
    cleaned = re.sub(r"[\\/:*?\"<>|]+", "_", value).strip()
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned or "untitled"


WORD_HEADER_TOKENS = (
    "word",
    "单词",
    "词",
    "english",
    "term",
    "entry",
    "front",
    "vocabulary",
)


def detect_word_column(headers: Sequence[str]) -> int:
    for idx, header in enumerate(headers):
        label = normalize_label(header)
        if any(name in label for name in WORD_HEADER_TOKENS):
            return idx
    return 0


def has_word_column_header(headers: Sequence[str]) -> bool:
    return any(
        any(name in normalize_label(header) for name in WORD_HEADER_TOKENS)
        for header in headers
    )


def build_note_fields(raw_row: Sequence[object], headers: Sequence[str]) -> tuple[str, str] | None:
    row = [safe_text(v) for v in raw_row]
    if not any(row):
        return None

    word_col = detect_word_column(headers)
    front = row[word_col] if word_col < len(row) else ""
    if not front:
        front = next((cell for cell in row if cell), "")
    if not front:
        return None

    segments: list[str] = []
    for idx, header in enumerate(headers):
        if idx == word_col:
            continue

        value = row[idx] if idx < len(row) else ""
        if not value:
            continue

        label = header.strip() or f"Column {idx + 1}"
        display_value = html.escape(value).replace("\n", "<br>")
        segments.append(f"<div><b>{html.escape(label)}</b>: {display_value}</div>")

    if not segments:
        return front, front

    back = "".join(segments)
    return front, back


def read_xlsx_notes(path: Path) -> list[tuple[str, str]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    if not workbook.sheetnames:
        return []

    ws = workbook[workbook.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    # Skip leading empty rows until the first non-empty row is found.
    first_index = 0
    for idx, row in enumerate(rows):
        if any(safe_text(cell) for cell in row):
            first_index = idx
            break
    else:
        return []

    headers = [safe_text(cell) for cell in rows[first_index]]
    notes: list[tuple[str, str]] = []
    for row in rows[first_index + 1 :]:
        if not any(safe_text(cell) for cell in row):
            continue
        pair = build_note_fields(row, headers)
        if pair:
            notes.append(pair)

    return notes


def read_xls_notes(path: Path) -> list[tuple[str, str]]:
    workbook = xlrd.open_workbook(str(path), on_demand=True)
    if workbook.nsheets == 0:
        return []

    sheet = workbook.sheet_by_index(0)
    rows = [
        [sheet.cell_value(row_index, col_index) for col_index in range(sheet.ncols)]
        for row_index in range(sheet.nrows)
    ]
    if not rows:
        return []

    first_index = next(
        (idx for idx, row in enumerate(rows) if any(safe_text(cell) for cell in row)),
        None,
    )
    if first_index is None:
        return []

    headers = [safe_text(cell) for cell in rows[first_index]]
    first_data_index = first_index + 1
    if not has_word_column_header(headers):
        # Many legacy XLS vocabulary lists contain data immediately, without a header row.
        first_data_index = first_index
        headers = ["Word"] + ["Meaning"] + [f"Column {idx + 3}" for idx in range(max(0, len(headers) - 2))]

    notes: list[tuple[str, str]] = []
    for row in rows[first_data_index:]:
        if not any(safe_text(cell) for cell in row):
            continue
        pair = build_note_fields(row, headers)
        if pair:
            notes.append(pair)

    return notes


def read_txt_notes(path: Path) -> list[tuple[str, str]]:
    data = path.read_bytes()
    encodings = ["utf-8", "utf-8-sig", "gb18030", "gbk", "cp936"]
    text = None
    for enc in encodings:
        try:
            text = data.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        try:
            text = data.decode("latin-1")
        except UnicodeDecodeError:
            return []

    notes: list[tuple[str, str]] = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        # Split on the first space, keeping the remaining definition intact.
        match = line.split(" ", maxsplit=1)
        if len(match) == 1:
            word = line.strip()
            meaning = line.strip()
        else:
            word, meaning = match[0].strip(), match[1].strip()

        if not word:
            continue
        # TXT contents are plain text; escape HTML-sensitive characters for Anki fields.
        notes.append((html.escape(word), html.escape(meaning or word)))

    return notes


def stable_int(value: str) -> int:
    digest = hashlib.md5(value.encode("utf-8")).hexdigest()
    # SQLite INTEGER is signed 64-bit; keep deck IDs safely below 2^63 - 1.
    return int(digest[:15], 16)


def make_model(model_id: int) -> genanki.Model:
    return genanki.Model(
        model_id,
        "Vocab Basic",
        fields=[
            {"name": "Front"},
            {"name": "Back"},
        ],
        templates=[
            {
                "name": "Card 1",
                "qfmt": "{{Front}}",
                "afmt": "{{Front}}<hr id=\"answer\">{{Back}}",
            }
        ],
        css=(
            ".card { font-family: Arial, sans-serif; font-size: 20px; line-height: 1.5; text-align: left; }\n"
            "strong { font-weight: 700; }\n"
        ),
    )


def build_deck_for_xlsx(path: Path, out_path: Path, root: Path | None = None) -> bool:
    suffix = path.suffix.lower()
    if suffix == ".txt":
        notes = read_txt_notes(path)
    elif suffix == ".xls":
        notes = read_xls_notes(path)
    else:
        notes = read_xlsx_notes(path)
    if not notes:
        print(f"skip: {path} (no readable rows)")
        return False

    if root is not None:
        try:
            rel_path = path.relative_to(root)
        except ValueError:
            rel_path = path
        deck_name = sanitize_name("-".join(rel_path.with_suffix("").parts))
    else:
        deck_name = sanitize_name(path.stem)
    deck_id = stable_int(str(path.resolve()))
    model_id = deck_id + 1

    model = make_model(model_id)
    deck = genanki.Deck(deck_id, deck_name)

    for front, back in notes:
        note = genanki.Note(model=model, fields=[front, back], tags=[suffix.lstrip("."), "vocab"])
        deck.add_note(note)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    genanki.Package(deck).write_to_file(str(out_path))
    print(f"ok: {path} -> {out_path}")
    return True


def find_xlsx_files(root: Path) -> Iterable[Path]:
    for path in sorted(root.rglob("*")):
        if not path.is_file():
            continue
        if path.suffix.lower() not in {".xlsx", ".xls", ".txt"}:
            continue
        if any(part.startswith(".") for part in path.parts if part not in {".", root.name}):
            continue
        if "__MACOSX" in path.parts:
            continue
        yield path


def main() -> int:
    parser = argparse.ArgumentParser(description="Convert multiple .xlsx/.xls/.txt files into Anki .apkg decks.")
    parser.add_argument("--root", type=Path, default=Path("."), help="Root directory to scan recursively")
    parser.add_argument("--output-dir", type=Path, default=Path(".anki"), help="Directory for generated .apkg files; default is .anki")
    parser.add_argument("--force", action="store_true", help="Overwrite existing .apkg files")
    args = parser.parse_args()

    root = args.root.resolve()
    if not root.exists():
        print(f"error: root does not exist: {root}", file=sys.stderr)
        return 1

    total = 0
    failed = 0
    for xlsx_path in find_xlsx_files(root):
        rel = xlsx_path.relative_to(root)
        out_path = args.output_dir / rel.with_suffix(".apkg")

        if out_path.exists() and not args.force:
            print(f"skip: {xlsx_path} (output exists: {out_path})")
            continue

        try:
            if build_deck_for_xlsx(xlsx_path, out_path, root=root):
                total += 1
        except Exception as exc:  # pragma: no cover - command-line fallback
            failed += 1
            print(f"error: failed to process {xlsx_path}: {exc}", file=sys.stderr)

    print(f"Completed: {total} APKG file(s) generated; {failed} failure(s).")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
