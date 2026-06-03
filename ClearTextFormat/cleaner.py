#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Text Cleaner — Khử nhiễu Unicode cho văn bản tiếng Việt (và mọi ngôn ngữ)
Hỗ trợ: .xlsx, .csv, .txt, .md, .docx

Chạy 1 lệnh:
  Windows:  irm irm https://raw.githubusercontent.com/trungtdv4/Portable-Apps/refs/heads/main/ClearTextFormat/cleaner.py | python | python
  macOS:    curl -sL irm https://raw.githubusercontent.com/trungtdv4/Portable-Apps/refs/heads/main/ClearTextFormat/cleaner.py | python | python3
"""

import sys, os, subprocess, importlib, importlib.util

# ── Force UTF-8 output (quan trọng khi chạy qua PowerShell pipe) ─────────────
def _fix_encoding():
    for stream in (sys.stdout, sys.stderr):
        try:
            if stream.encoding and stream.encoding.upper().replace("-","") not in ("UTF8",):
                stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

_fix_encoding()

# ── Bootstrap: tự cài dependencies nếu thiếu ─────────────────────────────────
REQUIRED = {"openpyxl": "openpyxl", "docx": "python-docx"}

def _bootstrap():
    missing = [pip for mod, pip in REQUIRED.items()
               if importlib.util.find_spec(mod) is None]
    if not missing:
        return
    print(f"\n Cai thu vien lan dau: {', '.join(missing)} ...")
    subprocess.check_call(
        [sys.executable, "-m", "pip", "install", "--quiet", *missing],
        stdout=subprocess.DEVNULL
    )
    print("Cai xong!\n")
    importlib.invalidate_caches()

_bootstrap()

# ── Re-launch nếu chạy qua pipe (irm | python) ───────────────────────────────
def _relaunch_if_piped():
    """Khi chạy qua stdin, không có __file__ → download về temp rồi execv lại."""
    try:
        _ = __file__
        if os.path.exists(__file__):
            return  # Chạy từ file bình thường, bỏ qua
    except NameError:
        pass  # __file__ không tồn tại → đang chạy qua pipe

    import tempfile, urllib.request

    RAW_URL = "irm https://raw.githubusercontent.com/trungtdv4/Portable-Apps/refs/heads/main/ClearTextFormat/cleaner.py | python"

    tmp = tempfile.NamedTemporaryFile(
        suffix=".py", delete=False, mode="w", encoding="utf-8"
    )
    try:
        with urllib.request.urlopen(RAW_URL) as r:
            tmp.write(r.read().decode("utf-8"))
        tmp.close()
        # Truyền env đảm bảo UTF-8 cho process con
        env = os.environ.copy()
        env["PYTHONUTF8"] = "1"
        env["PYTHONIOENCODING"] = "utf-8:replace"
        os.execve(sys.executable, [sys.executable, tmp.name], env)
    except Exception as e:
        tmp.close()
        try: os.unlink(tmp.name)
        except: pass
        print(f"Loi download: {e}\nThu chay: python cleaner.py")
        sys.exit(1)

_relaunch_if_piped()

# ── Imports chính ─────────────────────────────────────────────────────────────
import re, unicodedata, shutil
from pathlib import Path

# ── Màu terminal ──────────────────────────────────────────────────────────────
def _color_ok():
    return hasattr(sys.stdout, "isatty") and sys.stdout.isatty()

C = {k: (v if _color_ok() else "") for k, v in {
    "reset": "\033[0m", "bold": "\033[1m", "green": "\033[92m",
    "yellow": "\033[93m", "cyan": "\033[96m", "red": "\033[91m", "gray": "\033[90m",
}.items()}

def c(color, text): return f"{C[color]}{text}{C['reset']}"

# ── Core: logic khử nhiễu ─────────────────────────────────────────────────────
_COMBINING_MAP = {}

def _build_combining_map():
    bases = list("aeiouyAEIOUY") + list("âêôăơưÂÊÔĂƠƯ")
    for base in bases:
        for comb in ["\u0300","\u0301","\u0309","\u0303","\u0323"]:
            composed = unicodedata.normalize("NFC", base + comb)
            if composed != base + comb:
                _COMBINING_MAP[base + comb] = composed

_build_combining_map()
_COMBINING_PAIRS = sorted(_COMBINING_MAP.items(), key=lambda x: -len(x[0]))

_JUNK_RE = re.compile(
    r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f\x80-\x9f"
    r"\u0300-\u036f\u200b-\u200d\ufeff]"
)
_SPACE_RE = re.compile(r"[ \t]{2,}")

def clean_text(text: str) -> str:
    if not text: return text
    text = text.replace("\u00a0", " ").replace("\t", " ")
    for combo, pre in _COMBINING_PAIRS:
        if combo in text:
            text = text.replace(combo, pre)
    text = _JUNK_RE.sub("", text)
    return _SPACE_RE.sub(" ", text).strip()

def count_dirty(text: str) -> int:
    if not text: return 0
    return sum(text.count(k) for k in _COMBINING_MAP) + len(_JUNK_RE.findall(text))

# ── Processors ────────────────────────────────────────────────────────────────
def process_xlsx(src, dst):
    import openpyxl
    wb = openpyxl.load_workbook(src)
    s = {"cells": 0, "dirty": 0}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    s["cells"] += 1
                    if count_dirty(cell.value): s["dirty"] += 1
                    cell.value = clean_text(cell.value)
    wb.save(dst); return s

def process_csv(src, dst):
    import csv, io
    for enc in ["utf-8-sig","utf-8","cp1258","latin-1"]:
        try: content = src.read_text(encoding=enc); break
        except UnicodeDecodeError: continue
    else: raise ValueError("Khong doc duoc encoding cua file CSV")
    s = {"cells": 0, "dirty": 0}
    rows = []
    for row in csv.reader(io.StringIO(content)):
        new = []
        for cell in row:
            s["cells"] += 1
            if count_dirty(cell): s["dirty"] += 1
            new.append(clean_text(cell))
        rows.append(new)
    with open(dst, "w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows(rows)
    return s

def process_text(src, dst):
    for enc in ["utf-8-sig","utf-8","cp1258","latin-1"]:
        try: content = src.read_text(encoding=enc); break
        except UnicodeDecodeError: continue
    else: raise ValueError("Khong doc duoc encoding")
    s = {"cells": 0, "dirty": 0}
    out = []
    for line in content.splitlines(keepends=True):
        s["cells"] += 1
        if count_dirty(line): s["dirty"] += 1
        nl, body = "", line
        for end in ["\r\n","\n","\r"]:
            if line.endswith(end): nl, body = end, line[:-len(end)]; break
        out.append(clean_text(body) + nl)
    dst.write_text("".join(out), encoding="utf-8"); return s

def process_docx(src, dst):
    from docx import Document
    doc = Document(src)
    s = {"cells": 0, "dirty": 0}
    def cp(para):
        for run in para.runs:
            if run.text:
                s["cells"] += 1
                if count_dirty(run.text): s["dirty"] += 1
                run.text = clean_text(run.text)
    for para in doc.paragraphs: cp(para)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for para in cell.paragraphs: cp(para)
    doc.save(dst); return s

PROCESSORS = {
    ".xlsx": process_xlsx, ".csv": process_csv,
    ".txt": process_text, ".md": process_text, ".docx": process_docx,
}

# ── UI ────────────────────────────────────────────────────────────────────────
BANNER = """
╔══════════════════════════════════════════╗
║        TEXT CLEANER  -  Unicode Fixer    ║
║   Ho tro: .xlsx  .csv  .txt  .md  .docx ║
╚══════════════════════════════════════════╝"""

def ask(prompt, default=""):
    try:
        ans = input(prompt).strip()
        return ans if ans else default
    except (KeyboardInterrupt, EOFError):
        print("\nDa huy.")
        sys.exit(0)

def resolve_output(src):
    print()
    print("Luu ket qua:")
    print(f"  1  Tao file moi  ({src.stem}_cleaned{src.suffix})")
    print(f"  2  Ghi de file goc")
    choice = ask("Chon (1/2) [mac dinh: 1]: ", "1")
    if choice == "2":
        ok = ask(f"Ghi de '{src.name}'? Khong the hoan tac. (y/N): ", "n")
        if ok.lower() != "y":
            print("-> Doi sang tao file moi.")
            choice = "1"
    return src if choice == "2" else src.parent / f"{src.stem}_cleaned{src.suffix}"

def main():
    print(BANNER)

    if len(sys.argv) > 1:
        file_path = Path(sys.argv[1])
    else:
        raw = ask("\nNhap duong dan file (hoac keo tha vao day): ").strip("'\"")
        file_path = Path(raw)

    if not file_path.exists():
        print(f"\nKhong tim thay file: {file_path}")
        sys.exit(1)

    ext = file_path.suffix.lower()
    if ext not in PROCESSORS:
        print(f"\nLoai file '{ext}' chua duoc ho tro.")
        print(f"Ho tro: {', '.join(PROCESSORS.keys())}")
        sys.exit(1)

    print(f"\n  File      : {file_path.name}")
    print(f"  Kich thuoc: {file_path.stat().st_size/1024:.1f} KB")

    dst = resolve_output(file_path)

    if dst == file_path:
        backup = file_path.with_suffix(file_path.suffix + ".bak")
        shutil.copy2(file_path, backup)
        print(f"\n  Backup    : {backup.name}")

    print("\n  Dang xu ly...")

    try:
        s = PROCESSORS[ext](file_path, dst)
    except Exception as e:
        print(f"\nLoi: {e}")
        sys.exit(1)

    label = "dong" if ext in (".txt",".md") else "o"
    print(f"\nHoan thanh!")
    print(f"  Da quet  : {s['cells']:,} {label}")
    print(f"  Co nhieu : {s['dirty']} {label}")
    print(f"  {'File moi' if dst != file_path else 'Da luu'} : {dst}\n")

if __name__ == "__main__":
    main()