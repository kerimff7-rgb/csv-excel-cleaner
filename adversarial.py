"""
ADVERSARIAL CHECKS — the inputs nobody designs for.

test_csv_to_excel.py checks what the author thought of. This file exists
because that is exactly the problem. Here the tool is fed things no sane
client would send and several things a confused one certainly will: empty
files, binary data with a .csv name, a million-character line, a thousand
columns, NUL bytes, a PNG someone renamed, a workbook that is not a
workbook.

The bar is not "produces a good result". The bar is:

    it must not crash, and it must not silently produce a wrong one.

Either finish the job, or fail with a message that says what was wrong.
A traceback in front of a client is the failure this file is here to catch.

    python adversarial.py
"""
import os
import shutil
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

import csv_to_excel as CSV

PASS, FAIL = [], []
TMP = Path(tempfile.mkdtemp())


def verdict(label, fn):
    """Run fn. Pass if it returns cleanly or raises a RuntimeError with a
    message. Fail on any other exception — that is a traceback in front of
    a client."""
    try:
        note = fn()
        PASS.append(label)
        print(f"  [OK  ] {label}" + (f"  — {note}" if note else ""))
    except RuntimeError as e:
        msg = str(e)[:70].replace("\n", " ")
        if msg.strip():
            PASS.append(label)
            print(f"  [OK  ] {label}  — refused cleanly: {msg}")
        else:
            FAIL.append(label)
            print(f"  [FAIL] {label}  — raised RuntimeError with no message")
    except SystemExit:
        PASS.append(label)
        print(f"  [OK  ] {label}  — exited with a code, not a traceback")
    except Exception as e:
        FAIL.append(label)
        print(f"  [FAIL] {label}  — {type(e).__name__}: {str(e)[:90]}")


def section(t):
    print(f"\n{t}\n" + "-" * 74)


def slug(name):
    """A folder name built from a test label, safe on every platform.

    Windows forbids : * ? " < > | in a path; Linux does not. A label like
    "ragged rows: 3 columns, then 7" made a folder on Linux and a
    NotADirectoryError on Windows - the harness itself crashed before the
    tool under test was ever called."""
    import re as _re
    s = _re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip("._-")
    return (s or "case")[:48]


def wf(name, content, encoding="utf-8", binary=False):
    d = TMP / slug(name)
    d.mkdir(exist_ok=True)
    p = d / "input.csv"
    if binary:
        p.write_bytes(content)
    else:
        p.write_text(content, encoding=encoding, newline="")
    return d, p


# ===================================================== 1. CSV -> EXCEL
section("1. csv_to_excel — files nobody should send, and someone will")

cases = [
    ("zero bytes", b"", True),
    ("only a UTF-8 BOM", "﻿", False),
    ("only newlines", "\n\n\n\n\n", False),
    ("one column, no delimiter anywhere",
     "name\nAnna\nBatyr\nCharlie\nDiana\nEmir\n", False),
    ("headers but not one data row", "a,b,c\n", False),
    ("a header row and nothing but blank lines", "a,b,c\n\n\n\n", False),
    ("every value identical", "a,b\n" + "1,2\n" * 200, False),
    ("NUL bytes in the middle", "a,b\n1,\x002\n3,4\n", False),
    ("a line one million characters long",
     "a,b\n" + "x" * 1_000_000 + ",1\n", False),
    ("ragged rows: 3 columns, then 7, then 1",
     "a,b,c\n1,2,3\n1,2,3,4,5,6,7\n9\n", False),
    ("emoji, right-to-left text and combining marks",
     "name,note\n🙂,مرحبا\né,ok\n", False),
    ("a quote that is never closed", 'a,b\n"unterminated,2\n3,4\n', False),
    ("the delimiter inside every single field",
     "a,b\n\"1,1\",\"2,2\"\n\"3,3\",\"4,4\"\n", False),
]
for label, content, binary in cases:
    d, _ = wf(label, content, binary=binary)
    verdict(label, lambda d=d: (
        CSV.build(sorted(d.glob("*.csv")), TMP / "o.xlsx") and None))

# a PNG that claims to be a CSV
png = (b"\x89PNG\r\n\x1a\n" + bytes(range(256)) * 40)
d, _ = wf("a PNG renamed to .csv", png, binary=True)
verdict("a PNG renamed to .csv",
        lambda: (CSV.build(sorted(d.glob("*.csv")), TMP / "o.xlsx") and None))

# a thousand columns
d = TMP / "wide"; d.mkdir(exist_ok=True)
(d / "input.csv").write_text(
    ",".join(f"c{i}" for i in range(1000)) + "\n" +
    "\n".join(",".join(str(i) for i in range(1000)) for _ in range(20)) + "\n",
    encoding="utf-8")
verdict("1000 columns",
        lambda: (CSV.build([d / "input.csv"], TMP / "wide.xlsx") and None))

# a folder with nothing in it
empty = TMP / "empty_dir"; empty.mkdir(exist_ok=True)
verdict("a folder containing no CSVs at all",
        lambda: (CSV.build(sorted(empty.glob("*.csv")), TMP / "e.xlsx")
                 and None))

# A name Excel will not accept as a sheet. Excel forbids [ ] : * ? / \ in a
# sheet name; Windows forbids : * ? / \ in a FILE name. The intersection is
# what can be tested through a real file on Windows - the rest is covered in
# memory by the sheet_name checks in test_csv_to_excel.py.
d = TMP / "badname"; d.mkdir(exist_ok=True)
stem = ("a[b]c-d-e-f' and a very long tail that goes past the limit"
        if os.name == "nt" else
        "a[b]c:d*e?f' and a very long tail that goes past the limit")
weird = d / (stem + ".csv")
weird.write_text("a,b\n1,2\n", encoding="utf-8")
verdict("a filename full of characters Excel forbids",
        lambda: (CSV.build([weird], TMP / "bn.xlsx") and None))


# ------------------------------------------------ 1b. workbooks, not CSVs
# The tool reads .xlsx now, which is a second front door and therefore a
# second set of ways in. A client's "Excel file" is regularly not one:
# a CSV somebody renamed, a file half-written by a crashing export, a
# password-protected book, a sheet with a title and nothing under it.
section("1b. csv_to_excel — workbooks that are not workbooks")

from openpyxl import Workbook as _WB


def xl(name, fill):
    d = TMP / slug(f"xl_{name}")
    d.mkdir(exist_ok=True)
    p = d / "book.xlsx"
    fill(p)
    return d, p


def one_sheet(build_fn):
    def fill(p):
        wb = _WB()
        build_fn(wb.active)
        wb.save(p)
    return fill


xl_cases = [
    ("random bytes wearing an .xlsx name",
     lambda p: p.write_bytes(bytes(range(256)) * 40)),
    ("a CSV somebody renamed to .xlsx",
     lambda p: p.write_text("id,city\n1,Ashgabat\n2,Dubai\n", encoding="utf-8")),
    ("a zero-byte .xlsx", lambda p: p.write_bytes(b"")),
    ("a workbook holding one cell",
     one_sheet(lambda ws: ws.cell(row=1, column=1, value="hello"))),
    ("a workbook holding only a merged title",
     one_sheet(lambda ws: (ws.cell(row=1, column=1, value="REGISTER 2026"),
                           ws.merge_cells("A1:L1")))),
    ("a sheet of 400 columns and 3 rows",
     one_sheet(lambda ws: [ws.cell(row=r, column=c, value=f"v{r}{c}")
                           for r in range(1, 4) for c in range(1, 401)])),
    ("a sheet where every cell is None",
     one_sheet(lambda ws: [ws.cell(row=r, column=c, value=None)
                           for r in range(1, 20) for c in range(1, 8)])),
    ("a header of 60 columns and data rows of 3",
     one_sheet(lambda ws: (
         [ws.cell(row=1, column=c, value=f"col{c}") for c in range(1, 61)]
         + [ws.cell(row=r, column=c, value=r * c)
            for r in range(2, 9) for c in range(1, 4)]))),
    ("a cell holding a formula that was never calculated",
     one_sheet(lambda ws: (
         [ws.cell(row=1, column=c, value=n)
          for c, n in enumerate(["id", "qty", "total"], start=1)]
         + [(ws.cell(row=r, column=1, value=r),
             ws.cell(row=r, column=2, value=r * 2),
             ws.cell(row=r, column=3, value="=B%d*10" % r))
            for r in range(2, 12)]))),
]
for label, fill in xl_cases:
    d, _ = xl(label, fill)
    verdict(label, (lambda dd: lambda: (
        CSV.build(sorted(dd.iterdir()), TMP / "xo.xlsx") and None))(d))

# A workbook of many sheets must not take the tool down, and each sheet
# must land as its own worksheet rather than quietly overwriting the last.
def many_sheets(p):
    wb = _WB()
    for i in range(30):
        ws = wb.active if i == 0 else wb.create_sheet()
        ws.title = f"Block {i:02d}"
        ws.cell(row=1, column=1, value="id")
        ws.cell(row=1, column=2, value="amount")
        for r in range(2, 25):
            ws.cell(row=r, column=1, value=r)
            ws.cell(row=r, column=2, value=f"{r * 11:,.2f}")
    wb.save(p)


d, _ = xl("thirty_sheets", many_sheets)


def thirty():
    from openpyxl import load_workbook
    out = TMP / "xthirty.xlsx"
    CSV.build(sorted(d.iterdir()), out)
    names = load_workbook(out).sheetnames
    if len(names) != 31:
        raise AssertionError(f"expected 31 sheets, got {len(names)}")
    return f"{len(names) - 1} sheets, each kept separate"


verdict("a workbook of thirty sheets", thirty)



# ============================================================== result
print("\n" + "=" * 74)
print(f"  {len(PASS)} survived, {len(FAIL)} crashed, "
      f"{len(PASS)+len(FAIL)} adversarial inputs")
if FAIL:
    print("\n  crashed with a traceback:")
    for f in FAIL:
        print(f"    - {f}")
print("=" * 74 + "\n")
shutil.rmtree(TMP, ignore_errors=True)
sys.exit(1 if FAIL else 0)
