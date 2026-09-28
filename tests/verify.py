"""Verify the generated workbook (handles encrypted files)."""
import io
import re
import sys
import tempfile
import os
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from openpyxl import load_workbook
from config import OUTPUT_FILE, WORKBOOK_PASSWORD, ENCRYPT_OUTPUT


COMBINING = re.compile(r"[\u064B-\u065F]")
SHEET_REF = re.compile(r"('?[^'!\s]+'?)!\$")


def _load_decrypted(path: str):
    """Load workbook; if encrypted, decrypt to a temp file first."""
    if not ENCRYPT_OUTPUT:
        return load_workbook(path)

    import msoffcrypto
    with open(path, "rb") as f:
        office_file = msoffcrypto.OfficeFile(f)
        if not office_file.is_encrypted():
            return load_workbook(path)
        office_file.load_key(password=WORKBOOK_PASSWORD)
        fd, tmp = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)
        with open(tmp, "wb") as out:
            office_file.decrypt(out)
    try:
        return load_workbook(tmp)
    finally:
        # Keep tmp around during verification; delete after
        # (openpyxl keeps the file open, so cleanup happens when process exits)
        pass


def check_sheet(ws, name):
    issues = []
    for row in ws.iter_rows():
        for cell in row:
            val = cell.value
            if not (isinstance(val, str) and val.startswith("=")):
                continue
            if COMBINING.search(val):
                issues.append((cell.coordinate, "combining mark", val[:80]))
            for m in SHEET_REF.finditer(val):
                ref = m.group(1)
                is_quoted = ref.startswith("'") and ref.endswith("'")
                needs_quotes = " " in ref or any(ord(c) > 127 for c in ref)
                if needs_quotes and not is_quoted:
                    issues.append((cell.coordinate, "sheet not quoted", ref))
            if val.count("(") != val.count(")"):
                issues.append((cell.coordinate, "unbalanced parens", val[:80]))
            if val.count('"') % 2 != 0:
                issues.append((cell.coordinate, "odd quotes", val[:80]))
    icon = "❌" if issues else "✅"
    print(f"{icon} {name}: {len(issues)} issue(s)")
    for iss in issues[:5]:
        print("   ", iss)
    return issues


def main():
    if not Path(OUTPUT_FILE).exists():
        print(f"❌ File not found: {OUTPUT_FILE}")
        print("   Run `python main.py` first to generate the workbook.")
        sys.exit(1)

    print(f"📂 Verifying: {OUTPUT_FILE}")
    wb = _load_decrypted(OUTPUT_FILE)

    total = 0
    for name in wb.sheetnames:
        total += len(check_sheet(wb[name], name))

    print(f"\nTotal issues: {total}")
    sys.exit(0 if total == 0 else 1)


if __name__ == "__main__":
    main()