#!/usr/bin/env python3
"""
Text Cleaner — Khử nhiễu Unicode cho văn bản tiếng Việt (và mọi ngôn ngữ)
Hỗ trợ: .xlsx, .csv, .txt, .md, .docx

Chạy 1 lệnh (sau khi đã push lên GitHub):
  Windows:  irm https://raw.githubusercontent.com/trungtdv4/text-cleauserner/main/cleaner.py | python
  macOS:    curl -sL https://raw.githubusercontent.com/trungtdv4/text-cleaner/main/cleaner.py | python3
"""

# ── Bootstrap: tự cài dependencies nếu thiếu ─────────────────────────────────
import sys, os, subprocess, importlib, importlib.util

REQUIRED = {"openpyxl": "openpyxl", "docx": "python-docx"}

def _bootstrap():
    missing = [pip for mod, pip in REQUIRED.items() if importlib.util.find_spec(mod) is None]
    if not missing:
        return
    print(f"\n📦 Cài thư viện lần đầu: {', '.join(missing)} ...")
    subprocess.check_call(
        [sys.executable, "-m", "pip", "install", "--quiet", *missing],
        stdout=subprocess.DEVNULL
    )
    print("✅ Cài xong!\n")
    # Reload sys.path để import được ngay
    import importlib as _il
    _il.invalidate_caches()

_bootstrap()

# ── Khi chạy qua pipe (irm | python), không có file trên disk ────────────────
# Script tự detect và re-launch đúng cách
def _relaunch_if_piped():
    """Nếu chạy qua stdin pipe, download về temp rồi chạy lại."""
    if os.path.exists(__file__):
        return  # Chạy từ file bình thường, không cần làm gì

    import tempfile, urllib.request, urllib.error

    RAW_URL = "https://raw.githubusercontent.com/trungtdv4/text-cleaner/main/cleaner.py"

    tmp = tempfile.NamedTemporaryFile(suffix=".py", delete=False, mode="w", encoding="utf-8")
    try:
        with urllib.request.urlopen(RAW_URL) as r:
            tmp.write(r.read().decode("utf-8"))
        tmp.close()
        os.execv(sys.executable, [sys.executable, tmp.name])
    except urllib.error.URLError:
        tmp.close()
        os.unlink(tmp.name)
        # Nếu không download được thì chạy tiếp từ stdin (chức năng cơ bản vẫn hoạt động)

try:
    _relaunch_if_piped()
except NameError:
    pass  # __file__ không tồn tại khi chạy qua pipe — đã handle ở trên

# ── Imports chính (sau bootstrap) ─────────────────────────────────────────────
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
    combinings = ["\u0300","\u0301","\u0309","\u0303","\u0323"]
    for base in bases:
        for comb in combinings:
            composed = unicodedata.normalize("NFC", base + comb)
            if composed != base + comb:
                _COMBINING_MAP[base + comb] = composed

_build_combining_map()
_COMBINING_PAIRS = sorted(_COMBINING_MAP.items(), key=lambda x: -len(x[0]))

_JUNK_RE = re.compile(
    r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f"
    r"\x80-\x9f"
    r"\u0300-\u036f"
    r"\u200b-\u200d"
    r"\ufeff]"
)
_MULTI_SPACE_RE = re.compile(r"[ \t]{2,}")

def clean_text(text: str) -> str:
    if not text:
        return text
    text = text.replace("\u00a0", " ").replace("\t", " ")
    for combo, precomposed in _COMBINING_PAIRS:
        if combo in text:
            text = text.replace(combo, precomposed)
    text = _JUNK_RE.sub("", text)
    text = _MULTI_SPACE_RE.sub(" ", text)
    return text.strip()

def count_dirty(text: str) -> int:
    if not text: return 0
    n = sum(text.count(k) for k in _COMBINING_MAP)
    n += len(_JUNK_RE.findall(text))
    return n

# ── Processors ────────────────────────────────────────────────────────────────
def process_xlsx(src, dst):
    import openpyxl
    wb = openpyxl.load_workbook(src)
    stats = {"cells": 0, "dirty": 0}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    stats["cells"] += 1
                    if count_dirty(cell.value): stats["dirty"] += 1
                    cell.value = clean_text(cell.value)
    wb.save(dst)
    return stats

def process_csv(src, dst):
    import csv, io
    for enc in ["utf-8-sig", "utf-8", "cp1258", "latin-1"]:
        try: content = src.read_text(encoding=enc); break
        except UnicodeDecodeError: continue
    else:
        raise ValueError("Không đọc được encoding của file CSV")
    stats = {"cells": 0, "dirty": 0}
    rows = []
    for row in csv.reader(io.StringIO(content)):
        new_row = []
        for cell in row:
            stats["cells"] += 1
            if count_dirty(cell): stats["dirty"] += 1
            new_row.append(clean_text(cell))
        rows.append(new_row)
    with open(dst, "w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows(rows)
    return stats

def process_text(src, dst):
    for enc in ["utf-8-sig", "utf-8", "cp1258", "latin-1"]:
        try: content = src.read_text(encoding=enc); break
        except UnicodeDecodeError: continue
    else:
        raise ValueError("Không đọc được encoding")
    stats = {"cells": 0, "dirty": 0}
    lines = content.splitlines(keepends=True)
    out = []
    for line in lines:
        stats["cells"] += 1
        if count_dirty(line): stats["dirty"] += 1
        nl = ""
        body = line
        for ending in ["\r\n", "\n", "\r"]:
            if line.endswith(ending):
                nl, body = ending, line[:-len(ending)]
                break
        out.append(clean_text(body) + nl)
    dst.write_text("".join(out), encoding="utf-8")
    return stats

def process_docx(src, dst):
    from docx import Document
    doc = Document(src)
    stats = {"cells": 0, "dirty": 0}
    def clean_para(para):
        for run in para.runs:
            if run.text:
                stats["cells"] += 1
                if count_dirty(run.text): stats["dirty"] += 1
                run.text = clean_text(run.text)
    for para in doc.paragraphs: clean_para(para)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs: clean_para(para)
    doc.save(dst)
    return stats

PROCESSORS = {".xlsx": process_xlsx, ".csv": process_csv,
              ".txt": process_text, ".md": process_text, ".docx": process_docx}

# ── UI ────────────────────────────────────────────────────────────────────────
BANNER = f"""
{C['cyan']}{C['bold']}╔══════════════════════════════════════════╗
║        TEXT CLEANER  —  Unicode Fixer    ║
║   Hỗ trợ: .xlsx  .csv  .txt  .md  .docx ║
╚══════════════════════════════════════════╝{C['reset']}"""

def ask(prompt, default=""):
    try:
        ans = input(prompt).strip()
        return ans if ans else default
    except (KeyboardInterrupt, EOFError):
        print("\n" + c("yellow", "Đã huỷ."))
        sys.exit(0)

def resolve_output(src):
    print()
    print(c("bold", "Lưu kết quả:"))
    print(f"  {c('cyan','1')}  Tạo file mới  {c('gray', f'({src.stem}_cleaned{src.suffix})')}")
    print(f"  {c('cyan','2')}  Ghi đè file gốc")
    choice = ask(c("bold", "Chọn (1/2) [mặc định: 1]: "), "1")
    if choice == "2":
        ok = ask(c("yellow", f"⚠  Ghi đè '{src.name}'? Không thể hoàn tác. (y/N): "), "n")
        if ok.lower() != "y":
            print(c("gray", "→ Đổi sang tạo file mới."))
            choice = "1"
    return src if choice == "2" else src.parent / f"{src.stem}_cleaned{src.suffix}"

def print_stats(stats, src, dst):
    label = "dòng" if src.suffix in (".txt", ".md") else "ô"
    print()
    print(c("green", "✔  Hoàn thành!"))
    print(f"   Đã quét  : {stats['cells']:,} {label}")
    dirty_str = c("yellow", str(stats["dirty"])) if stats["dirty"] else c("green", "0")
    print(f"   Có nhiễu : {dirty_str} {label}")
    print(f"   {'File mới' if dst != src else 'Đã lưu  '} : {c('cyan', str(dst))}")
    print()

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    print(BANNER)

    if len(sys.argv) > 1:
        file_path = Path(sys.argv[1])
    else:
        raw = ask(c("bold", "\nNhập đường dẫn file (hoặc kéo thả vào đây): ")).strip("'\"")
        file_path = Path(raw)

    if not file_path.exists():
        print(c("red", f"\n✘  Không tìm thấy file: {file_path}"))
        sys.exit(1)

    ext = file_path.suffix.lower()
    if ext not in PROCESSORS:
        print(c("red", f"\n✘  Loại file '{ext}' chưa được hỗ trợ."))
        print(c("gray", f"   Hỗ trợ: {', '.join(PROCESSORS.keys())}"))
        sys.exit(1)

    print(f"\n   File      : {c('cyan', file_path.name)}")
    print(f"   Kích thước: {c('gray', f'{file_path.stat().st_size/1024:.1f} KB')}")

    dst = resolve_output(file_path)

    if dst == file_path:
        backup = file_path.with_suffix(file_path.suffix + ".bak")
        shutil.copy2(file_path, backup)
        print(c("gray", f"\n   Backup    : {backup.name}"))

    print(c("gray", "\n   Đang xử lý..."))

    try:
        stats = PROCESSORS[ext](file_path, dst)
    except Exception as e:
        print(c("red", f"\n✘  Lỗi: {e}"))
        sys.exit(1)

    print_stats(stats, file_path, dst)

if __name__ == "__main__":
    main()
