"""
Checks for csv_to_excel.py. No network, no fixtures on disk — every case is
built in memory and asserted against a known answer.

Nine of these are named REGRESSION. Each one is a bug that was in this code
and shipped nothing, because it was caught here: a data row taken for the
header, a label column turned into numbers, an order number turned into a
year, an ISO date read day-first so 3 January became 1 March, a merged group
header in a workbook beating the real one, amounts stored as text that
stayed text, and three separate defaults that made Excel open every chart
with no axis labels at all.

None of the nine crashes. Every one of them finishes, writes a workbook
that looks correct, and is wrong — which is the only kind of failure that
reaches a client without being noticed first.

    python test_csv_to_excel.py
"""
import io
import re
import sys
import tempfile
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import csv_to_excel as M

PASS, FAIL = [], []


def check(label, got, want):
    ok = got == want
    (PASS if ok else FAIL).append(label)
    mark = "OK  " if ok else "FAIL"
    line = f"  [{mark}] {label}"
    if not ok:
        line += f"\n         expected {want!r}\n         got      {got!r}"
    print(line)


def section(title):
    print(f"\n{title}\n" + "-" * 70)


def write_tmp(text, name, encoding="utf-8"):
    p = Path(tempfile.mkdtemp()) / name
    p.write_text(text, encoding=encoding, newline="")
    return p


# ------------------------------------------------------------ delimiters
section("1. Delimiter detection")
for delim, label in ((",", "comma"), (";", "semicolon"),
                     ("\t", "tab"), ("|", "pipe")):
    sample = "\n".join(delim.join(["a", "b", "c", "d"]) for _ in range(6))
    check(f"{label} detected", M.sniff_delimiter(sample), delim)

# a single stray comma inside a semicolon file must not win
sample = "\n".join(f"id;name;note{i}" for i in range(8)) + "\nx;y,z;w"
check("stray comma does not beat semicolons", M.sniff_delimiter(sample), ";")


# ------------------------------------------------------------- encodings
section("2. Encoding detection")
p = write_tmp("name,city\nAnna,Ashgabat\n", "u.csv", "utf-8")
check("utf-8 read", M.read_bytes(p)[1] in ("utf-8", "utf-8-sig"), True)
p = write_tmp("imya,gorod\nАнна,Ашхабад\n", "w.csv", "cp1251")
text, enc = M.read_bytes(p)
check("cp1251 read", enc, "cp1251")
check("cp1251 content intact", "Ашхабад" in text, True)


# ----------------------------------------------------------- header rows
section("3. Header row detection")
t = "a,b,c\n1,2,3\n4,5,6\n"
check("header already on line 1", M.find_header_row(t, ","), 0)

t = ("Monthly export\nGenerated automatically\n\n"
     "id,city,amount\n1,Ashgabat,10\n2,Dubai,20\n3,Baku,30\n")
check("three junk lines above the header", M.find_header_row(t, ","), 3)

# REGRESSION: a quoted field containing the delimiter made data rows look
# wider than the header, so a data row was chosen and the rows above it
# were silently discarded.
t = ("Report\n\nid,city,revenue\n"
     + "".join(f'{i},Ashgabat,"{i},500.00"\n' for i in range(1, 15)))
check("REGRESSION quoted delimiter does not shift the header",
      M.find_header_row(t, ","), 2)

# a file that is all numbers has no header to find; row 0 is returned
t = "1,2,3\n4,5,6\n7,8,9\n"
check("all-numeric file falls back to row 0", M.find_header_row(t, ","), 0)


# --------------------------------------------------------- column names
section("4. Column names")
check("whitespace trimmed and collapsed",
      M.clean_names([" Order  ID ", "Date "]), ["Order ID", "Date"])
check("duplicates made unique",
      M.clean_names(["a", "a", "a"]), ["a", "a_1", "a_2"])
check("blank names replaced",
      M.clean_names(["", "b"]), ["column_1", "b"])
check("quotes stripped", M.clean_names(['"City"']), ["City"])


# -------------------------------------------------------------- numbers
section("5. Number parsing")
cases = [
    ("1234", 1234.0, "plain integer"),
    ("1234.56", 1234.56, "plain decimal"),
    ("1,234.56", 1234.56, "comma thousands, dot decimal"),
    ("1 234,56", 1234.56, "space thousands, comma decimal"),
    ("1234,56", 1234.56, "comma as decimal mark"),
    ("$ 1,234.56", 1234.56, "currency sign"),
    ("-42.5", -42.5, "negative"),
    ("1 234,00", 1234.0, "non-breaking space thousands"),
]
for raw, want, label in cases:
    check(label, M.to_number(raw), want)

# REGRESSION: an over-permissive pattern stripped any leading text, so the
# label "Manager 3" became the number 3 and the column lost its meaning.
for raw, label in (("Manager 3", "REGRESSION 'Manager 3' is not a number"),
                   ("Widget A", "text is not a number"),
                   ("Q1 2026", "'Q1 2026' is not a number"),
                   ("", "empty string is not a number"),
                   ("N/A", "'N/A' is not a number")):
    check(label, M.to_number(raw), None)


# ---------------------------------------------------------------- dates
section("6. Date parsing")

# REGRESSION: ISO dates parsed day-first turned 2026-01-03 into 1 March.
col = pd.Series(["2026-01-03", "2026-01-05", "2026-02-11", "2026-12-31"])
out, _ = M.coerce_dates(col)
check("REGRESSION ISO date keeps its month",
      out.iloc[0].strftime("%Y-%m-%d"), "2026-01-03")
check("ISO date, second value", out.iloc[2].strftime("%Y-%m-%d"), "2026-02-11")

col = pd.Series(["03.02.2026", "15.02.2026", "28.02.2026", "01.03.2026"])
out, _ = M.coerce_dates(col)
check("dd.mm.yyyy read day-first",
      out.iloc[0].strftime("%Y-%m-%d"), "2026-02-03")

# REGRESSION: a column of bare integers was read as years, so order number
# 1001 became the date 1001-01-01.
col = pd.Series(["1001", "1002", "1003", "1004"])
out, _ = M.coerce_dates(col)
check("REGRESSION bare integers are not dates", out, None)

col = pd.Series(["Ashgabat", "Dubai", "Baku"])
out, _ = M.coerce_dates(col)
check("text column is not dates", out, None)


# -------------------------------------------------- numeric column rule
section("7. Column conversion thresholds")
col = pd.Series(["10", "20", "30", "40", "50"])
out, n = M.coerce_numeric(col)
check("all-numeric column converts", n, 5)

col = pd.Series(["10", "20", "cancelled", "refunded", "pending"])
out, _ = M.coerce_numeric(col)
check("mostly-text column stays text", out, None)

col = pd.Series(["10", "20", "30", "40", "n/a"])
out, n = M.coerce_numeric(col)
check("one bad value out of five still converts", n, 4)


# ----------------------------------------------------------- end to end
section("8. End to end")
src = ("Export\n\nid,city,amount,empty\n"
       "1,Ashgabat,\"1,000.00\",\n"
       "2,Dubai,250.50,\n"
       "2,Dubai,250.50,\n"          # duplicate
       "\n"                          # blank row
       "3,Baku,75,\n")
p = write_tmp(src, "e2e.csv")
df, notes, numeric, dates = M.load_clean(p)

check("rows read", notes["rows read"], 4)
check("duplicate removed", notes["duplicate rows removed"], 1)
check("empty column dropped", notes["empty columns dropped"], 1)
check("rows kept", notes["rows kept"], 3)
check("columns kept", list(df.columns), ["id", "city", "amount"])
check("thousands separator survived", float(df["amount"].iloc[0]), 1000.0)
check("amount is numeric", "amount" in numeric, True)
check("city stays text", "city" in numeric, False)

out = Path(tempfile.mkdtemp()) / "out.xlsx"
M.build([p], out)
from openpyxl import load_workbook
wb = load_workbook(out)
check("Summary sheet exists", "Summary" in wb.sheetnames, True)
check("data sheet created", len(wb.sheetnames), 2)
ws = wb[wb.sheetnames[1]]
# Row 1 holds the table header and, further right past a gap, the heading
# of the block the charts read from. Only the table part is asserted here;
# the autofilter range below is what proves the two do not overlap.
check("header written",
      [c.value for c in ws[1]][:3], ["id", "city", "amount"])
check("the autofilter covers the table only, not the chart block",
      ws.auto_filter.ref.split(":")[1][0], "C")
check("header frozen", ws.freeze_panes, "A2")
check("filter applied", ws.auto_filter.ref is not None, True)
# Excel stores every number the same way, so openpyxl hands back an int
# when there is no fractional part. What matters is that the cell holds a
# number and carries a numeric format, not which Python type it arrives as.
check("amount cell is a number",
      isinstance(ws["C2"].value, (int, float)), True)
check("amount cell value", float(ws["C2"].value), 1000.0)
check("amount cell formatted as a number", ws["C2"].number_format, "#,##0.00")

# sheet names must survive characters Excel forbids. Slashes are left out:
# they are directory separators, and Path().stem strips them correctly.
used = set()
check("forbidden characters replaced",
      M.sheet_name("a[b]c:d*e?f.csv", used), "a-b-c-d-e-f")
check("long name truncated to 31",
      len(M.sheet_name("x" * 60 + ".csv", used)), 31)
check("colliding names made unique",
      M.sheet_name("Summary.csv", {"summary"}), "Summary_1")


# ---------------------------------------------------- Excel as the input
section("9. Excel workbooks as input")

from openpyxl import Workbook as _WB


def make_book(path, sheets):
    """A workbook shaped the way client files arrive: a merged title, a
    printed-on line, a blank row, a merged group header, and only then the
    real column names."""
    wb = _WB()
    for si, (title, n) in enumerate(sheets):
        ws = wb.active if si == 0 else wb.create_sheet()
        ws.title = title
        ws["A1"] = f"REGISTER — {title} — March 2026"
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
        ws["A2"] = "printed 01.04.2026"
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=6)
        ws.cell(row=4, column=3, value="Charges, TMT")
        ws.merge_cells(start_row=4, start_column=3, end_row=4, end_column=5)
        for j, h in enumerate(["No", "Tenant", "Fee", "Paid", "Balance",
                               "Note", "Comment"], start=1):
            ws.cell(row=5, column=j, value=h)
        for i in range(1, n + 1):
            ws.cell(row=5 + i, column=1, value=i)
            ws.cell(row=5 + i, column=2, value=f"Tenant {i}")
            # amounts typed as TEXT with thousands separators - the defect
            # that stops a client's own SUM from working
            ws.cell(row=5 + i, column=3, value=f"{1000 + i:,.2f}")
            ws.cell(row=5 + i, column=4, value=f"{500 + i:,.2f}")
            ws.cell(row=5 + i, column=5, value=f"{500:,.2f}")
            if i % 11 == 0:                      # the note column is sparse
                ws.cell(row=5 + i, column=6, value="moved out")
    wb.save(path)
    return path


# REGRESSION: the header must be found by width, and a header is WIDER than
# the typical data row because it names the columns the data leaves empty.
# Requiring an exact match to the modal width threw the real header away and
# returned the two-cell merged group header above it, so every column came
# out named column_1 ... column_6 and the first data row was lost.
grid = [
    ["REGISTER", "", "", "", "", ""],
    ["printed 01.04.2026", "", "", "", "", ""],
    ["", "", "Charges, TMT", "", "", ""],
    ["No", "Tenant", "Fee", "Paid", "Balance", "Note"],
    ["1", "Tenant 1", "1,001.00", "501.00", "500.00", ""],
    ["2", "Tenant 2", "1,002.00", "502.00", "500.00", ""],
    ["3", "Tenant 3", "1,003.00", "503.00", "500.00", ""],
]
check("REGRESSION a merged group header does not win over the real one",
      M.find_header_in_grid(grid), 3)

check("a sheet whose header is already on line 1",
      M.find_header_in_grid([["a", "b", "c"], ["1", "2", "3"],
                             ["4", "5", "6"]]), 0)
check("a sheet of bare numbers falls back to the first full row",
      M.find_header_in_grid([["1", "2", "3"], ["4", "5", "6"]]), 0)

p = make_book(Path(tempfile.mkdtemp()) / "billing.xlsx",
              [("Block A", 40), ("Block B", 25)])
sheets = M.load_sheets(p)
check("both sheets are read", len(sheets), 2)
check("sheet labels come from the workbook",
      [s[0] for s in sheets], ["Block A", "Block B"])

label, df, notes, numeric, dates = sheets[0]
check("the real header row is used",
      list(df.columns), ["No", "Tenant", "Fee", "Paid", "Balance", "Note"])
check("the header's row number is the one Excel shows, not the one left "
      "after blank rows were dropped", notes["header row"], 5)
check("every data row survives", len(df), 40)
check("the title rows are not data", str(df["Tenant"].iloc[0]), "Tenant 1")
check("REGRESSION text amounts become real numbers",
      float(df["Fee"].iloc[0]), 1001.0)
check("and the column is reported as numeric", "Fee" in numeric, True)
check("a column with a header and no values at all is dropped",
      notes["empty columns dropped"], 1)
check("but a sparsely filled column is kept", "Note" in df.columns, True)
check("the second sheet keeps its own row count", len(sheets[1][1]), 25)

out = Path(tempfile.mkdtemp()) / "book.xlsx"
M.build([p], out)
wb = load_workbook(out)
check("one output sheet per input sheet",
      wb.sheetnames, ["Summary", "Block A", "Block B"])
check("the amount lands in Excel as a number, not as text",
      isinstance(wb["Block A"]["C2"].value, (int, float)), True)

# the headline band on the Summary has to add up to the table beneath it
s = wb["Summary"]
check("the headline band counts both tables", s["B5"].value, 2)
check("the headline band's row total matches the sheets",
      s["C5"].value, 65)

# A column of formulas written by a script carries no saved result, so
# openpyxl reads it as blank and the column would be dropped as "empty" —
# silently, and it is usually the totals column. It must be named instead.
fwb = _WB()
fws = fwb.active
for j, h in enumerate(["id", "qty", "price", "total"], start=1):
    fws.cell(row=1, column=j, value=h)
for r in range(2, 22):
    fws.cell(row=r, column=1, value=r - 1)
    fws.cell(row=r, column=2, value=r)
    fws.cell(row=r, column=3, value=10.5)
    fws.cell(row=r, column=4, value=f"=B{r}*C{r}")
fp = Path(tempfile.mkdtemp()) / "totals.xlsx"
fwb.save(fp)
warn = M.read_excel_sheets(fp)[0][4]
check("an empty column that actually holds formulas is named, not dropped "
      "in silence", "total" in warn, True)
check("and the message says how to fix it", "save it once" in warn, True)

# a text file wearing an .xlsx name gets a fix, not a zip error
tp = Path(tempfile.mkdtemp()) / "renamed.xlsx"
tp.write_text("id,city\n1,Ashgabat\n2,Dubai\n", encoding="utf-8")
try:
    M.read_excel_sheets(tp)
    msg = ""
except RuntimeError as e:
    msg = str(e)
check("a CSV renamed to .xlsx is told what it is", "not an Excel workbook"
      in msg, True)
check("and told what to do about it", "Rename it to .csv" in msg, True)

# A binary wearing a .csv or .xls name should be told what it is and what
# to do, not just "this is binary". Each of these has a one-step fix and a
# client who is told the step does it; a client who is told "binary file"
# writes back and asks.
bd = Path(tempfile.mkdtemp())
for name, head, want in (
        ("old.xls", b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1", "save it as .xlsx"),
        ("book.dat", b"PK\x03\x04\x14\x00\x00\x00", "rename it to .xlsx"),
        ("photo.csv", b"\x89PNG\r\n\x1a\n", "binary file")):
    bp = bd / name
    bp.write_bytes(head + b"\x00" * 600)
    try:
        M.load_sheets(bp)
        msg = ""
    except RuntimeError as e:
        msg = str(e)
    check(f"{name} is refused with the fix, not just a label",
          want in msg, True)

# a workbook with nothing in it must refuse by name, not crash
empty = _WB()
empty.active["A1"] = "just a title"
ep = Path(tempfile.mkdtemp()) / "empty.xlsx"
empty.save(ep)
res = M.load_sheets(ep)
check("a workbook with no table is reported, not crashed",
      res[0][1] is None, True)


# ------------------------------------------------- choosing what to chart
section("10. Which column a chart is drawn from")

# A register of charges: the fee is charged, some of it is paid. Charting
# either alone hides the only question worth asking, which is the gap
# between them.
reg = pd.DataFrame({
    "No": range(1, 201),
    "Area m2": [40.0 + i % 50 for i in range(200)],
    "Rate per m2": [10.57] * 150 + [12.40] * 50,     # a tariff, not a measure
    "Monthly fee": [500.0 + i for i in range(200)],
    "Paid": [400.0 + i for i in range(200)],
    "Service": ["water", "heating"] * 100,
})
picked = M.pick_measures(reg, ["No", "Area m2", "Rate per m2",
                               "Monthly fee", "Paid"])
check("two comparable money columns are charted together",
      picked, ["Monthly fee", "Paid"])
check("an identifier is never charted as a quantity", "No" in picked, False)
check("a tariff with two distinct values is not a measure",
      "Rate per m2" in picked, False)

# Revenue and quantity are both 'money-like' by name, but a bar of 2,800
# beside a bar of 150,000 is an invisible bar. Only the larger is charted.
shop = pd.DataFrame({
    "Order ID": range(1, 101),
    "Qty": [2] * 100,
    "Revenue": [1500.0] * 100,
    "Product": ["drill", "grinder"] * 50,
})
check("columns of different magnitudes are not paired",
      M.pick_measures(shop, ["Order ID", "Qty", "Revenue"]), ["Revenue"])

# Nothing is named like money: fall back rather than draw nothing.
plain = pd.DataFrame({"city": ["a", "b"] * 50,
                      "reading": [float(i) for i in range(100)]})
check("a table with no money column still gets a measure",
      M.pick_measures(plain, ["reading"]), ["reading"])
check("a table with no numbers at all gets no chart",
      M.pick_measures(plain, []), [])

# and the two-series block has to reach the sheet as two columns
two = Path(tempfile.mkdtemp()) / "two.xlsx"
M.build([make_book(Path(tempfile.mkdtemp()) / "pair.xlsx", [("S", 30)])], two)
tws = load_workbook(two)["S"]
hdr = [tws.cell(row=1, column=c).value for c in range(1, tws.max_column + 1)]
check("the chart block names both measures",
      "Fee" in hdr and "Paid" in hdr, True)


# ------------------------------------------------ the charts must be legible
section("11. Charts Excel will actually draw")

# REGRESSION: the charts opened in Excel as bars with no labels — no
# category names underneath, no numbers beside them. Three separate defaults
# did it, and none of them raises an error, so the only way to catch this is
# to read the XML Excel will read.
#
#   * openpyxl leaves the axes' `delete` unset and Excel reads that as
#     "delete this axis";
#   * set_categories always writes a NUMERIC reference, so text labels came
#     out as 1, 2, 3;
#   * the category axis was positioned on the left, where the value axis is.
import zipfile as _zip

cp = Path(tempfile.mkdtemp()) / "charted.xlsx"
csrc = ("city,service,amount\n"
        + "".join(f"{c},{s},{100 + i}\n"
                  for i, (c, s) in enumerate(
                      [("Ashgabat", "water"), ("Dubai", "heating"),
                       ("Baku", "lift"), ("Almaty", "waste")] * 6)))
M.build([write_tmp(csrc, "charted.csv")], cp)

zf = _zip.ZipFile(cp)
charts = [n for n in zf.namelist() if n.startswith("xl/charts/chart")]
check("a chart was produced at all", len(charts) >= 1, True)
xml = zf.read(sorted(charts)[0]).decode()

check("REGRESSION neither axis is marked for deletion",
      set(re.findall(r'delete val="([^"]*)"', xml)), {"0"})
check("REGRESSION category labels are a text reference, not a numeric one",
      "<cat><strRef>" in xml, True)
check("and not a numeric one", "<cat><numRef>" in xml, False)
check("REGRESSION the category axis sits under the bars, not beside them",
      re.search(r"<catAx>.*?<axPos val=\"(.*?)\"", xml, re.S).group(1), "b")
check("the value axis stays on the left",
      re.search(r"<valAx>.*?<axPos val=\"(.*?)\"", xml, re.S).group(1), "l")
check("tick labels are placed, not left to chance",
      set(re.findall(r'tickLblPos val="([^"]*)"', xml)), {"nextTo"})
check("the category reference points at the label column",
      re.search(r"<cat><strRef><f>(.*?)</f>", xml).group(1).count("$"), 4)


# ======================================================================
# 12. MERGING  (--merge)
# Added 17.09.2026. A market survey of 75 postings found four jobs that
# needed exactly this feature and nothing else. It is opt-in: every check
# above still describes the default behaviour, untouched.
# ======================================================================
print("\n-- 12. merging ------------------------------------------------")


def _csv(dirpath, name, header, rows, delim=",", enc="utf-8"):
    p = Path(dirpath) / name
    with open(p, "w", encoding=enc, newline="") as f:
        f.write(delim.join(header) + "\n")
        for r in rows:
            f.write(delim.join(str(x) for x in r) + "\n")
    return p


HDR = ["Order ID", "City", "Amount"]
mdir = Path(tempfile.mkdtemp())
_csv(mdir, "jan.csv", HDR, [[1, "Ashgabat", "10.50"], [2, "Dubai", "20.25"]])
_csv(mdir, "feb.csv", HDR, [[3, "Almaty", "30.00"], [4, "Baku", "40.75"]],
     delim=";")
# the thousands separator needs a delimiter that is not a comma, or the
# field splits in two - which is a property of CSV, not of the merge.
_csv(mdir, "mar.csv", HDR, [[5, "Dubai", "1,050.00"]], delim=";")
_csv(mdir, "other.csv", ["Region", "Manager"], [["North", "Ann"]])

mout = mdir / "m.xlsx"
recs = M.build(sorted(mdir.glob("*.csv")), mout, merge=True)
mwb = __import__("openpyxl").load_workbook(mout)

check("a Combined sheet is written when columns match",
      "Combined" in mwb.sheetnames, True)
check("the Combined sheet sits right after the Summary",
      mwb.sheetnames[1], "Combined")

cw = mwb["Combined"]
check("every matching row is carried over, none lost",
      cw.max_row - 1, 5)
hdr = [c.value for c in cw[1]][:4]
check("the source column comes first and is named", hdr[0], "source")
check("the original columns follow it", hdr[1:4], HDR)
sources = {cw.cell(r, 1).value for r in range(2, cw.max_row + 1)}
check("every source file is named in the source column",
      sources, {"jan", "feb", "mar"})

ai = [c.value for c in cw[1]].index("Amount") + 1
amounts = [cw.cell(r, ai).value for r in range(2, cw.max_row + 1)]
check("REGRESSION types are worked out again on the merged column",
      all(isinstance(v, (int, float)) for v in amounts), True)
check("a thousands separator in one file does not turn the column to text",
      round(sum(amounts), 2), 1151.50)

check("a table whose columns differ is NOT merged",
      "other" in mwb.sheetnames, True)
left = [r for r in recs if r.get("file") == "(not combined)"]
check("the table left out is reported", len(left), 1)
check("and the reason names the difference",
      "extra Manager, Region" in left[0]["warning"], True)

# The band above the table and the table below it have to agree. Counting a
# combined sheet twice makes them disagree, and the client sees a workbook
# that contradicts itself.
sm = mwb["Summary"]
check("REGRESSION the combined sheet is not counted as a source table",
      sm["B5"].value, 4)
# four source tables: jan 2 + feb 2 + mar 1 + other 1 = 6. The Combined
# sheet holds 5 of those rows again and must not add to this number.
check("REGRESSION merged rows are not counted twice in the headline band",
      sm["D5"].value, 6)

# A client who already has a column called 'source' must not have it
# overwritten by ours.
sdir = Path(tempfile.mkdtemp())
_csv(sdir, "a.csv", ["source", "Amount"], [["shop", 5]])
_csv(sdir, "b.csv", ["source", "Amount"], [["web", 7]])
sout = sdir / "s.xlsx"
M.build(sorted(sdir.glob("*.csv")), sout, merge=True)
sw = __import__("openpyxl").load_workbook(sout)["Combined"]
shdr = [c.value for c in sw[1]][:3]
check("REGRESSION a client column named 'source' is not overwritten",
      shdr[0], "source file")
check("and the client's own column survives intact", shdr[1], "source")
check("with its values", {sw.cell(r, 2).value for r in range(2, 4)},
      {"shop", "web"})

# The same invoice in two monthly exports is not a duplicate. Counting is
# useful; deleting is a decision the client never asked for.
rdir = Path(tempfile.mkdtemp())
_csv(rdir, "m1.csv", ["ID", "Amount"], [[1, 10], [2, 20]])
_csv(rdir, "m2.csv", ["ID", "Amount"], [[1, 10], [3, 30]])
rout = rdir / "r.xlsx"
rrecs = M.build(sorted(rdir.glob("*.csv")), rout, merge=True)
rw = __import__("openpyxl").load_workbook(rout)["Combined"]
check("rows repeated across sources are KEPT, not silently removed",
      rw.max_row - 1, 4)
comb = [r for r in rrecs if r.get("file") == "(combined)"][0]
check("but they are counted and reported",
      "1 rows repeat across sources" in comb["warning"], True)

# Default behaviour is untouched: no flag, no Combined sheet.
dout = mdir / "d.xlsx"
M.build(sorted(mdir.glob("*.csv")), dout)
check("REGRESSION without the flag nothing changes",
      "Combined" in __import__("openpyxl").load_workbook(dout).sheetnames,
      False)

# One table alone has nothing to merge with, and must not crash or produce
# a one-source 'Combined' sheet.
odir = Path(tempfile.mkdtemp())
_csv(odir, "only.csv", ["A", "B"], [[1, 2]])
oout = odir / "o.xlsx"
M.build(sorted(odir.glob("*.csv")), oout, merge=True)
check("a single table produces no Combined sheet",
      "Combined" in __import__("openpyxl").load_workbook(oout).sheetnames,
      False)

check("column order does not decide whether two tables match",
      M.merge_key(pd.DataFrame(columns=["b", "a"])),
      M.merge_key(pd.DataFrame(columns=["a", "b"])))
check("but column names do",
      M.merge_key(pd.DataFrame(columns=["a", "b"]))
      == M.merge_key(pd.DataFrame(columns=["a", "c"])), False)



# Excel cannot hold more than 1,048,576 rows in a sheet. A merge that goes
# past it must refuse in words, not write a file missing rows that the
# client discovers from their own totals weeks later.
_limit = M.EXCEL_MAX_ROWS
M.EXCEL_MAX_ROWS = 4
_a = pd.DataFrame({"ID": [1, 2], "Amount": [10, 20]})
_b = pd.DataFrame({"ID": [3, 4], "Amount": [30, 40]})
_m, _l = M.merge_tables([("a", _a), ("b", _b)])
check("REGRESSION a merge past Excel's row limit is refused, not truncated",
      len(_m), 0)
check("and the refusal says why", "exceeds the 4-row limit" in _l[0][1], True)
M.EXCEL_MAX_ROWS = _limit
_m, _l = M.merge_tables([("a", _a), ("b", _b)])
check("under the limit the same tables merge normally", len(_m), 1)



# ======================================================================
# 13. ЧИСЛА, КОТОРЫЕ ЧУТЬ НЕ СТАЛИ ДАТАМИ
# Найдено 17.09.2026 на демонстрационных данных для галереи, НЕ тестами.
# Колонка сумм вида 4141.98 - точка, без разделителя тысяч, целая часть
# в 1-4 знака - целиком превращалась в даты: 4141.98 -> 4141-01-01.
# Уже типизированный float 4821.55 превращался в 1970-01-01 плюс
# наносекунды. Ничего не падало. Отчёт открывался. Деньги исчезали.
#
# Баг был в опубликованном инструменте, а не в новом слиянии: DATE_HINT
# считал одну точку между цифрами признаком даты.
# ======================================================================
print("\n-- 13. числа против дат ---------------------------------------")

_num = [("суммы с точкой, 4 знака", ["4141.98", "9787.80", "588.40", "1633.25"]),
        ("суммы с запятой",         ["4,141.98", "9,787.80", "2,940.75"]),
        ("уже float",               [4141.98, 9787.80, 588.40]),
        ("мелкие float",            [35.5, 9.25, 11.75, 4.5]),
        ("проценты",                ["12.5", "7.25", "99.9"]),
        ("номера заказов",          ["100001", "100002", "100003"])]
for _label, _vals in _num:
    _got, _ = M.coerce_dates(pd.Series(_vals))
    check(f"REGRESSION не дата: {_label}", _got is None, True)

_dat = [("ISO",              ["2026-01-12", "2026-02-28", "2026-03-07"]),
        ("европейские",      ["03.02.2026", "28.02.2026", "17.03.2026"]),
        ("американские",     ["12/31/2026", "01/15/2026", "06/30/2026"]),
        ("с месяцем словом", ["12 Jan 2026", "28 Feb 2026", "7 Mar 2026"]),
        ("со временем",      ["2026-01-12 14:30", "2026-02-01 09:05"])]
for _label, _vals in _dat:
    _got, _ = M.coerce_dates(pd.Series(_vals))
    check(f"по-прежнему дата: {_label}", _got is not None, True)

# Сквозная проверка: файл, на котором баг и проявился.
_bd = Path(tempfile.mkdtemp())
with open(_bd / "invoices.csv", "w", encoding="utf-8", newline="") as _f:
    _f.write("Invoice,Date,Amount\n")
    for _i, _a in enumerate(["4141.98", "9787.80", "588.40", "1633.25",
                             "2940.75", "8115.20"], start=1):
        _f.write(f"INV-{_i:04d},2026-0{_i}-1{_i},{_a}\n")
_bo = _bd / "o.xlsx"
M.build([_bd / "invoices.csv"], _bo)
_bw = __import__("openpyxl").load_workbook(_bo)["invoices"]
_bh = [c.value for c in _bw[1]]
_bamt = [_bw.cell(r, _bh.index("Amount") + 1).value for r in range(2, 8)]
_bdat = [_bw.cell(r, _bh.index("Date") + 1).value for r in range(2, 8)]
check("REGRESSION сквозной прогон: Amount остался числом",
      all(isinstance(v, (int, float)) for v in _bamt), True)
check("и сумма не потерялась", round(sum(_bamt), 2), 27207.38)
check("а Date всё же стал датой",
      all(hasattr(v, "year") for v in _bdat), True)



# ======================================================================
# 14. ФАЙЛЫ, КОТОРЫХ ЗАКАЗЧИК НЕ ВИДИТ
# Найдено 17.09.2026 на репетиции заказа, НЕ тестами.
# turkmenabat.xlsx был открыт в Excel. Excel держит рядом скрытый
# ~$turkmenabat.xlsx - это блокировка, а не данные. Инструмент посчитал
# его шестым файлом, попытался прочитать, получил Permission denied и
# записал SKIPPED в Summary. Заказчик видит в отчёте строку про файл,
# которого у него в папке не видно, и решает, что инструмент сломан.
# macOS оставляет ._name.xlsx по той же схеме.
# ======================================================================
print("\n-- 14. служебные файлы ----------------------------------------")

check("REGRESSION ~$ - это блокировка Excel, не данные",
      M.is_junk("~$turkmenabat.xlsx"), True)
check("REGRESSION ._ - это огрызок от macOS",
      M.is_junk("._report.xlsx"), True)
check("а обычное имя - данные", M.is_junk("turkmenabat.xlsx"), False)
check("и имя с ~ не в начале тоже данные",
      M.is_junk("otchet~$avgust.csv"), False)
check("путь, а не только имя", M.is_junk(Path("a") / "b" / "~$x.xlsx"), True)

_jd = Path(tempfile.mkdtemp())
(_jd / "sales.csv").write_text("A,B\n1,2\n", encoding="utf-8")
(_jd / "~$sales.xlsx").write_bytes(b"\x00\x01lock")
(_jd / "._sales.csv").write_text("junk", encoding="utf-8")
(_jd / "notes.md").write_text("not a table", encoding="utf-8")
_got = [p.name for p in M.collect_inputs([str(_jd)])]
check("REGRESSION папка отдаёт только настоящие файлы", _got, ["sales.csv"])

_named = M.collect_inputs([str(_jd / "~$sales.xlsx")])
check("и явно названный lock-файл тоже не берём", _named, [])


# ======================================================================
# 15. ЗАГОЛОВКИ, КОТОРЫЕ ВИДНО
# Найдено 17.09.2026 глазами, при первом открытии отчёта в LibreOffice.
# Верхняя полоса итогов стоит в тех же столбцах, что и таблица под ней,
# но подписи у неё свои. "empty columns dropped" попадало в столбец,
# ширина которого посчитана под "rows kept". Высота строки 4 была жёстко
# задана в 26 пунктов, и заказчик читал "columns dropped" в самом
# заметном месте отчёта. Ни один из 144 тестов этого не видел: число
# было правильным, подпись над ним - обрезанной.
# ======================================================================
print("\n-- 15. высота полосы итогов ------------------------------------")

_wb15 = __import__("openpyxl").Workbook()
_ws15 = _wb15.active
for _L, _w in (("A", 12), ("B", 40)):
    _ws15.column_dimensions[_L].width = _w

check("короткие подписи - минимальная высота",
      M.band_height(_ws15, [("rows", 1), ("files", 2)]), 26)
_tall = M.band_height(_ws15, [("empty columns dropped", 1), ("files", 2)])
check("REGRESSION длинная подпись в узком столбце поднимает строку",
      _tall > 26, True)
check("а в широком столбце та же подпись помещается",
      M.band_height(_ws15, [("files", 1), ("empty columns dropped", 2)]), 26)
check("одно длинное слово в узком столбце тоже переносится",
      M.band_height(_ws15, [("supercalifragilistic", 1)]) > 26, True)

_hd = Path(tempfile.mkdtemp())
(_hd / "h.csv").write_text("Invoice,Amount\nA-1,10.50\nA-2,20.25\n",
                           encoding="utf-8")
_ho = _hd / "h.xlsx"
M.build([_hd / "h.csv"], _ho)
_hs = __import__("openpyxl").load_workbook(_ho)["Summary"]
_hh = _hs.row_dimensions[4].height
check("REGRESSION в готовом отчёте строка 4 выше жёстких 26",
      _hh > 26, True)
check("и подпись на месте",
      _hs.cell(4, 7).value, "empty columns dropped")


# --------------------------------------------------------------- result
print("\n" + "=" * 70)
print(f"  {len(PASS)} passed, {len(FAIL)} failed, {len(PASS)+len(FAIL)} total")
if FAIL:
    print("\n  failures:")
    for f in FAIL:
        print(f"    - {f}")
print("=" * 70 + "\n")
sys.exit(1 if FAIL else 0)
