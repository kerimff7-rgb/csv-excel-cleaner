"""
CSV TO EXCEL — turn a folder of messy CSV exports into one clean workbook.

The problem this solves. Reports arrive as CSV dumps from whatever system
produced them, and they are never uniform: one file is comma separated and
another semicolon, one was saved in Windows-1251, the header sits three
rows down under a title nobody removed, the revenue column is text because
somebody formatted it with thousands separators, and the same row appears
twice because an export was run twice.

Opening those by hand in Excel is where an afternoon goes. This does it in
one command and, more importantly, tells you what it changed: every fix is
counted and reported, so the output can be checked rather than trusted.

What it produces: one .xlsx with a Summary sheet describing what was found
in each file, one sheet per file holding the cleaned table with frozen
headers, filters and real number formats, and a chart wherever the data
supports one.

Usage:
    python csv_to_excel.py data/                    all CSVs in a folder
    python csv_to_excel.py data/ -o report.xlsx     choose the output name
    python csv_to_excel.py a.csv b.csv              specific files
    python csv_to_excel.py data/ --no-dedupe        keep duplicate rows
"""
import argparse
import csv
import io
import re
import sys
from pathlib import Path

import pandas as pd
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.data_source import AxDataSource, StrRef
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.utils import get_column_letter

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception:
    pass

ENCODINGS = ("utf-8-sig", "utf-8", "cp1251", "cp1252", "latin-1")
DELIMITERS = (",", ";", "\t", "|")
SUFFIXES = (".csv", ".tsv", ".txt", ".xlsx", ".xlsm")

# Names that carry a data extension but are not data. Excel leaves a hidden
# "~$name.xlsx" beside every workbook it currently has open, and macOS leaves
# "._name.xlsx" on anything that travelled through a USB stick. Both end in
# .xlsx, neither can be read - the first is locked by Excel, the second is a
# fragment. A client who runs this on a folder while one file is open in
# Excel would otherwise get a SKIPPED line about a file they cannot even see,
# which reads as a broken tool. Found on 17.09.2026, on a real folder.
JUNK_PREFIXES = ("~$", "._")


def is_junk(path):
    """True for an editor lock file or a resource-fork stub."""
    return Path(path).name.startswith(JUNK_PREFIXES)

HEAD_FILL = PatternFill("solid", fgColor="1F3864")
HEAD_FONT = Font(color="FFFFFF", bold=True)
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# Chart size in centimetres, and how many rows one chart occupies.
# 17 rows at the default row height is a little taller than 9 cm,
# which leaves a gap between chart rows instead of an overlap.
CHART_W_CM, CHART_H_CM, ROW_STEP = 16, 9, 19
# Excel's hard ceiling. A sheet cannot hold more, and a merge that goes
# past it must say so rather than write a file the client cannot open.
EXCEL_MAX_ROWS = 1_048_576

LEFT_PAD = Alignment(horizontal="left", indent=1)
RIGHT_PAD = Alignment(horizontal="right", indent=1)


# --------------------------------------------------------------------- read

def read_bytes(path):
    """Decode the file, trying the encodings that actually turn up in
    exports. Returns the text and the encoding that worked, because the
    encoding is itself worth reporting — it tells the client which system
    produced the file."""
    raw = Path(path).read_bytes()
    for enc in ENCODINGS:
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError:
            continue
    # latin-1 cannot fail, so this is unreachable in practice
    return raw.decode("latin-1", errors="replace"), "latin-1"


def sniff_delimiter(sample):
    """csv.Sniffer is right most of the time and confidently wrong the rest,
    so its answer is only accepted when it is one we expect. Otherwise fall
    back to counting: the delimiter is the candidate whose count per line is
    both highest and most consistent."""
    try:
        d = csv.Sniffer().sniff(sample, delimiters="".join(DELIMITERS)).delimiter
        if d in DELIMITERS:
            return d
    except csv.Error:
        pass
    lines = [ln for ln in sample.splitlines() if ln.strip()][:20]
    best, best_score = ",", -1.0
    for d in DELIMITERS:
        counts = [ln.count(d) for ln in lines]
        if not counts or max(counts) == 0:
            continue
        # consistent count across lines beats a high but erratic one
        spread = max(counts) - min(counts)
        score = sum(counts) / len(counts) - spread
        if score > best_score:
            best, best_score = d, score
    return best


def split_line(ln, delim):
    """Split one line the way the CSV dialect says, not by str.split. A
    field like "5,800.61" contains the delimiter inside quotes; splitting
    naively makes a data row look wider than the header, and the header
    detector below then picks a data row. That bug silently discarded the
    rows above it and used data as column names."""
    try:
        return next(csv.reader([ln], delimiter=delim))
    except (csv.Error, StopIteration):
        return ln.split(delim)


def find_header_row(text, delim, look=40):
    """Exports often carry a title, a timestamp and a blank line above the
    real header.

    The table's width is taken to be the most common field count among the
    first lines — data rows are many, junk lines are few, so the mode is the
    table. The header is then the FIRST line of that width whose fields are
    mostly non-numeric, because a row of numbers is data, not a header.
    First, not best: once the width matches, a later match is a data row."""
    lines = text.splitlines()[:look]
    idx = [i for i, ln in enumerate(lines) if ln.strip()]
    if not idx:
        return 0

    widths = {}
    for i in idx:
        n = len(split_line(lines[i], delim))
        if n >= 2:
            widths[i] = n
    if not widths:
        return 0

    counts = {}
    for n in widths.values():
        counts[n] = counts.get(n, 0) + 1
    table_w = max(counts, key=lambda n: (counts[n], n))

    for i in idx:
        if widths.get(i) != table_w:
            continue
        fields = [f.strip().strip('"') for f in split_line(lines[i], delim)]
        nonempty = [f for f in fields if f]
        if len(nonempty) < 2:
            continue
        numeric = sum(1 for f in nonempty
                      if re.fullmatch(r"[-+]?[\d\s.,]+", f))
        if numeric > len(nonempty) / 2:
            continue
        return i
    return idx[0]


def find_header_in_grid(rows, look=40):
    """The same question as find_header_row, asked of a sheet instead of a
    text file. A spreadsheet arrives with a title, a date, a blank row and
    only then the header - and often the header itself spans two rows with
    merged cells above it.

    The table's width is the most common count of filled cells; the header is
    the FIRST row about that wide whose cells are mostly non-numeric.

    "About that wide" rather than "exactly that wide" is the whole trick. A
    header names every column, so it is usually WIDER than the typical data
    row, which leaves a note or a comment empty - demanding an exact match
    threw the real header away. Everything above it is narrower: a merged
    title fills one cell, and a merged group header two or three. The floor
    sits between the two."""
    widths = {}
    for i, r in enumerate(rows[:look]):
        n = sum(1 for v in r if str(v).strip() not in ("", "None", "nan"))
        if n >= 2:
            widths[i] = n
    if not widths:
        return 0
    counts = {}
    for n in widths.values():
        counts[n] = counts.get(n, 0) + 1
    table_w = max(counts, key=lambda n: (counts[n], n))
    floor = max(2, round(table_w * 0.7))
    first_wide = None
    for i in sorted(widths):
        if widths[i] < floor:
            continue                      # a title or a group header
        if first_wide is None:
            first_wide = i
        cells = [str(v).strip() for v in rows[i]
                 if str(v).strip() not in ("", "None", "nan")]
        numeric = sum(1 for c in cells if re.fullmatch(r"[-+]?[\d\s.,]+", c))
        if numeric > len(cells) / 2:
            continue                      # a data row, not a header
        return i
    # No row reads as a header - a sheet of bare numbers. Take the first full
    # row rather than a stray two-cell note above it.
    return first_wide if first_wide is not None else min(widths)


def formula_warnings(path, empty_cols):
    """Why a column came back empty.

    openpyxl reads the value Excel last SAVED for a formula cell. A workbook
    written by a script - an export, a report generator, another tool - has
    no saved values, so every formula column reads as blank. Dropping it
    quietly would lose the column the client cares most about, usually the
    totals. Here the file is opened a second time, formulas and all, and the
    columns that hold formulas are named.

    The second read happens only when something was blank, so an ordinary
    workbook pays nothing for it."""
    if not empty_cols:
        return {}
    import openpyxl
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    except Exception:
        return {}
    warn = {}
    try:
        for ws in wb.worksheets:
            if ws.title not in empty_cols:
                continue
            idx, header_row, names = empty_cols[ws.title]
            hit = set()
            for row in ws.iter_rows(min_row=header_row + 1,
                                    max_row=header_row + 200):
                for j in idx:
                    if j - 1 < len(row):
                        v = row[j - 1].value
                        if isinstance(v, str) and v.startswith("="):
                            hit.add(names[j - 1])
            if hit:
                warn[ws.title] = (
                    "formulas with no saved result in: "
                    + ", ".join(sorted(hit))
                    + " - open the file in Excel and save it once")
    finally:
        try:
            wb.close()
        except Exception:
            pass
    return warn


def read_excel_sheets(path):
    """Every sheet of a workbook, as text, with its header row found.

    Client data arrives as .xlsx at least as often as .csv - a billing
    register, a stock list, a payroll sheet. Refusing those and saying "no
    CSVs found" would turn the most common case away at the door."""
    import openpyxl
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        # A file that is not a zip is not a workbook. Say which kind of
        # not-a-workbook it is, because the usual answer is "somebody
        # renamed a CSV" and that has a one-step fix.
        head = b""
        try:
            with open(path, "rb") as f:
                head = f.read(400)
        except OSError:
            pass
        if head and head[:2] != b"PK":
            if not head.strip():
                raise RuntimeError(f"{Path(path).name} is empty") from None
            try:
                sample = head.decode("utf-8")
                if sample.isprintable() or "\n" in sample:
                    raise RuntimeError(
                        f"{Path(path).name} is not an Excel workbook - it is "
                        f"plain text with an .xlsx name. Rename it to .csv "
                        f"and it will be read") from None
            except UnicodeDecodeError:
                pass
        raise RuntimeError(
            f"{Path(path).name} could not be opened as a workbook: {e}") from None

    out = []
    empty_cols = {}          # sheet title -> header names with no values
    for ws in wb.worksheets:
        raw = [["" if c is None else c for c in r]
               for r in ws.iter_rows(values_only=True)]
        # Blank rows are dropped, but their numbers are kept: the summary
        # reports the header's row as the user sees it in Excel, not its
        # position after the blanks were removed.
        keep = [(n, r) for n, r in enumerate(raw, start=1)
                if any(str(v).strip() not in ("", "None") for v in r)]
        rows = [r for _, r in keep]
        if len(rows) < 2:
            out.append((ws.title, None, 0, "sheet is empty", ""))
            continue
        h = find_header_in_grid(rows)
        real_row = keep[h][0]
        header = clean_names(rows[h])
        body = rows[h + 1:]
        width = len(header)
        body = [list(r[:width]) + [""] * max(0, width - len(r)) for r in body]
        df = pd.DataFrame(body, columns=header).astype(str)
        blank = [j for j, c in enumerate(df.columns, start=1)
                 if not df[c].str.strip().replace("None", "").any()]
        if blank:
            empty_cols[ws.title] = (blank, real_row, list(df.columns))
        out.append((ws.title, df, real_row, None, ""))
    try:
        wb.close()
    except Exception:
        pass
    if not out:
        raise RuntimeError(f"{Path(path).name} has no sheets")

    # A column that came back completely empty may not be empty at all: if
    # the workbook was written by a script rather than by Excel, its
    # formulas carry no cached result and openpyxl hands back nothing. That
    # is silent data loss, which is the one outcome this tool must not have.
    warn = formula_warnings(path, empty_cols)
    if warn:
        out = [(t, d, r, e, warn.get(t, "")) for t, d, r, e, _ in out]
    return out


# ------------------------------------------------------------------- clean

def clean_names(cols):
    """Strip, collapse inner whitespace, name the unnamed, and make
    duplicates unique — a duplicate column name silently loses data later."""
    out, seen = [], {}
    for i, c in enumerate(cols):
        name = re.sub(r"\s+", " ", str(c)).strip().strip('"').strip()
        if not name or name.lower().startswith("unnamed"):
            name = f"column_{i+1}"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        out.append(name)
    return out


# Only a currency sign or a percent may sit beside the digits. An earlier
# version allowed any non-digit prefix, which turned "Manager 3" into 3 and
# destroyed the labels in that column. A number is digits with a currency
# mark, not digits hiding inside a word.
NUM_RE = re.compile(
    r"^\s*[$€£¥₽₺]?\s*([-+]?[\d\s .,]+)\s*(?:%|[$€£¥₽₺]|"
    r"USD|EUR|RUB|TMT)?\s*$", re.I)

# A date needs a separator between digits, a time, or a month name. Without
# this guard pandas reads the bare integer 1001 as the year 1001, which is
# how an order number becomes a date.
DATE_HINT = re.compile(
    # A date written in digits needs TWO separators - 03.02.2026, 2026-01-12,
    # 12/31/26. One is not enough: "4141.98" is a sum of money, and reading it
    # as the year 4141 destroys the column silently. This cost a real bug on
    # 17.09.2026, found only because a demo invoice happened to be under
    # 10,000 and written without a thousands separator.
    # Sharpened again on 17.09.2026: two separators are necessary but not
    # sufficient. A size "10.5.2", an article number "10.20.30" and a
    # version "1.2.3" all have two, and all three were being read as dates.
    # The rule that separates them: with DOTS the year must be four digits.
    # 03.02.2026 is a date; 10.5.2 is a size. Slashes and dashes keep the
    # two-digit year, because 12/31/26 is how half the world writes it.
    #
    # 31/12/26, 31-12-2026
    r"\d{1,2}[-/]\d{1,2}[-/]\d{2,4}(?!\d)"
    # 2026-01-12, 2026/01/12, 2026.01.12
    r"|\d{4}[-/.]\d{1,2}[-/.]\d{1,2}(?!\d)"
    # 03.02.2026 - dots demand a full year
    r"|\d{1,2}\.\d{1,2}\.\d{4}(?!\d)"
    # a clock time
    r"|\d{1,2}:\d{2}"
    # or a word, which is how month names arrive
    r"|[A-Za-zА-Яа-я]{3,}")

# A bare decimal number - 4141.98, -12,5, 1 234.50 - is never a date, whatever
# else it may look like. Checked before any parsing is attempted.
BARE_NUMBER = re.compile(r"^[-+]?[\d\s]*[\d][\s\d]*(?:[.,]\d+)?$")


def to_number(s):
    """Parse a number written the way people write numbers: with a currency
    sign, with spaces or non-breaking spaces as thousands separators, and
    with either a comma or a dot as the decimal mark. Returns None when the
    value is not a number, so the caller can decide what that means."""
    if s is None:
        return None
    t = str(s).strip()
    if not t:
        return None
    m = NUM_RE.match(t)
    if not m:
        return None
    t = m.group(1).replace(" ", "").replace(" ", "")
    if "," in t and "." in t:
        # whichever comes last is the decimal mark
        t = (t.replace(",", "") if t.rfind(".") > t.rfind(",")
             else t.replace(".", "").replace(",", "."))
    elif "," in t:
        # a single comma with 1-2 trailing digits is a decimal mark,
        # otherwise it separates thousands
        t = t.replace(",", "." if re.search(r",\d{1,2}$", t) else "")
    try:
        return float(t)
    except ValueError:
        return None


# An unbroken run of digits long enough to be an identifier rather than a
# quantity, and a value whose first digit is a zero - which a number never
# has and an identifier very often does.
LONG_DIGITS = re.compile(r"^\d{7,}$")
LEADING_ZERO = re.compile(r"^0\d+$")


def is_identifier(vals, threshold=0.8):
    """True when a column holds identifiers, not measures.

    Found on 18.09.2026, after the same defect turned up in the API tool:
    a postal code 05401 became the number 5401 and the leading zero was
    gone for good, a phone 8025285988 became 8,025,285,988, and a barcode
    became a float. None of it crashed and none of it was reported.

    Two signals, both cheap and both hard to argue with:

      - a leading zero. No quantity is written 05401; an identifier very
        often is. Length does not matter here.
      - a long unbroken run of digits, ALL THE SAME LENGTH. Phones,
        accounts and barcodes are a fixed width; populations and amounts
        are not, which is what keeps this from eating real numbers."""
    zeros = vals.map(lambda v: bool(LEADING_ZERO.match(str(v).strip())))
    if zeros.mean() >= threshold:
        return True
    longs = vals.map(lambda v: bool(LONG_DIGITS.match(str(v).strip())))
    if longs.mean() >= threshold:
        widths = {len(str(v).strip()) for v, k in zip(vals, longs) if k}
        if len(widths) == 1:
            return True
    return False


def coerce_numeric(col, threshold=0.8):
    """Convert a text column to numbers only when almost all of its
    non-empty values parse. A column that is 60% numeric is usually a text
    column with some numbers in it, and converting it would destroy the
    rest."""
    vals = col.dropna().astype(str)
    vals = vals[vals.str.strip() != ""]
    if len(vals) == 0:
        return None, 0
    if is_identifier(vals, threshold):
        return None, 0
    parsed = vals.map(to_number)
    ok = parsed.notna().sum()
    if ok / len(vals) < threshold:
        return None, 0
    return col.map(lambda v: to_number(v) if str(v).strip() else None), int(ok)


def coerce_dates(col, threshold=0.8):
    """Same rule for dates. Day-first is tried first: exports written for
    Europe and Central Asia use it, and '03.02.2026' is otherwise silently
    read as March.

    The DATE_HINT guard comes first: a column of bare integers is never a
    date column, whatever pandas is willing to make of it."""
    vals = col.dropna().astype(str)
    vals = vals[vals.str.strip() != ""]
    if len(vals) == 0:
        return None, 0
    if vals.map(lambda v: bool(DATE_HINT.search(v))).mean() < threshold:
        return None, 0

    # Second guard, and the one that matters. A column of plain numbers must
    # never become dates. pandas will happily read "4141.98" as the year 4141
    # and a float 4821.55 as nanoseconds past 1970 - both look like a working
    # conversion and are a destroyed money column.
    if vals.map(lambda v: bool(BARE_NUMBER.match(v))).mean() >= threshold:
        return None, 0

    # ISO dates lead with a four-digit year and are never ambiguous. Trying
    # day-first on them turns 2026-01-03 into 1 March — a silent swap of day
    # and month that nobody notices until the monthly totals are wrong.
    iso = vals.str.match(r"^\s*\d{4}[-/]\d{1,2}[-/]\d{1,2}").mean() >= threshold
    order = (False,) if iso else (True, False)

    for dayfirst in order:
        p = pd.to_datetime(vals, errors="coerce", dayfirst=dayfirst, format="mixed")
        ok = p.notna().sum()
        if ok / len(vals) >= threshold:
            full = pd.to_datetime(col, errors="coerce", dayfirst=dayfirst,
                                  format="mixed")
            return full, int(ok)
    return None, 0


def clean_frame(df, notes, dedupe=True):
    """The cleaning every source shares, once the rows are in hand: names
    tidied, empty rows and columns dropped, duplicates removed, columns
    given their real types. Counted as it goes, because the count is the
    product."""
    df.columns = clean_names(df.columns)
    df = df.replace(r"^\s*$", pd.NA, regex=True)

    before_c = df.shape[1]
    df = df.dropna(axis=1, how="all")
    notes["empty columns dropped"] = before_c - df.shape[1]

    before_r = len(df)
    df = df.dropna(axis=0, how="all")
    notes["empty rows dropped"] = before_r - len(df)

    if dedupe:
        before_r = len(df)
        df = df.drop_duplicates()
        notes["duplicate rows removed"] = before_r - len(df)
    else:
        notes["duplicate rows removed"] = 0

    numeric, dates = [], []
    for c in df.columns:
        conv, _ = coerce_dates(df[c])
        if conv is not None:
            df[c] = conv
            dates.append(c)
            continue
        conv, _ = coerce_numeric(df[c])
        if conv is not None:
            df[c] = conv
            numeric.append(c)
        else:
            df[c] = df[c].astype(str).str.strip().replace("<NA>", "")

    notes["columns to numbers"] = len(numeric)
    notes["columns to dates"] = len(dates)
    notes["rows kept"] = len(df)
    return df.reset_index(drop=True), notes, numeric, dates


def load_sheets(path, dedupe=True):
    """Everything readable in one file: a CSV gives one table, a workbook
    gives one per sheet. Returns a list of (label, df, notes, numeric,
    dates) so the caller does not care which it was."""
    if Path(path).suffix.lower() in (".xlsx", ".xlsm"):
        out = []
        for name, raw, hrow, err, warn in read_excel_sheets(path):
            if err or raw is None or raw.empty:
                out.append((name, None, {"file": Path(path).name,
                                         "encoding": "xlsx",
                                         "delimiter": "sheet",
                                         "header row": hrow,
                                         "rows read": 0,
                                         "warning": err or "no data rows"},
                            [], []))
                continue
            notes = {"file": Path(path).name, "encoding": "xlsx",
                     "delimiter": f"sheet '{name}'", "header row": hrow,
                     "rows read": len(raw)}
            df, notes, num, dat = clean_frame(raw, notes, dedupe)
            if warn:
                notes["warning"] = warn
            out.append((name, df, notes, num, dat))
        return out
    df, notes, num, dat = load_clean(path, dedupe)
    return [(Path(path).stem, df, notes, num, dat)]


def load_clean(path, dedupe=True):
    """Read one delimited text file and return the cleaned frame plus a
    record of every change made to it."""
    text, enc = read_bytes(path)

    # Refuse before parsing, with a reason. A client who renames a PNG to
    # .csv, or sends a file the export wrote nothing into, should be told
    # that - not shown a traceback from inside pandas.
    if "\x00" in text:
        # Name the two binaries that actually turn up wearing the wrong
        # extension, because each has a one-step fix and "this is binary"
        # does not tell anyone what to do next.
        head = Path(path).read_bytes()[:8]
        if head.startswith(b"\xd0\xcf\x11\xe0"):
            raise RuntimeError(
                f"{Path(path).name} is an old Excel .xls file (the 1997 "
                f"format). Open it in Excel and save it as .xlsx - the data "
                f"does not change - then send it again")
        if head.startswith(b"PK\x03\x04"):
            raise RuntimeError(
                f"{Path(path).name} is a zip archive or an Office file with "
                f"the wrong extension. If it is a workbook, rename it "
                f"to .xlsx")
        raise RuntimeError(
            f"{Path(path).name} contains NUL bytes - this looks like a "
            f"binary file (an image, a zip, a database), not a CSV")
    if not text.strip():
        raise RuntimeError(f"{Path(path).name} is empty")

    delim = sniff_delimiter(text[:8000])
    skip = find_header_row(text, delim)

    try:
        df = pd.read_csv(io.StringIO(text), sep=delim, skiprows=skip,
                         dtype=str, keep_default_na=False, engine="python",
                         skip_blank_lines=True, on_bad_lines="skip")
    except pd.errors.EmptyDataError:
        raise RuntimeError(f"{Path(path).name} has no columns to read") from None
    except pd.errors.ParserError as e:
        raise RuntimeError(
            f"{Path(path).name} could not be parsed as {delim!r}-separated "
            f"text: {e}") from None

    notes = {"file": Path(path).name, "encoding": enc,
             "delimiter": {",": "comma", ";": "semicolon", "\t": "tab",
                           "|": "pipe"}.get(delim, delim),
             "header row": skip + 1, "rows read": len(df)}

    return clean_frame(df, notes, dedupe)


# ------------------------------------------------------------------- write

def sheet_name(name, used):
    """Excel forbids []:*?/\\ in sheet names and caps them at 31 chars."""
    s = re.sub(r"[\[\]:*?/\\]", "-", Path(name).stem)[:31] or "sheet"
    base, i = s, 1
    while s.lower() in used:
        suffix = f"_{i}"
        s = base[:31 - len(suffix)] + suffix
        i += 1
    used.add(s.lower())
    return s


def safe(v):
    """Excel rejects most control characters outright. A value that
    contains one is not worth losing the whole workbook over, so the
    character is dropped and the rest of the value kept."""
    if isinstance(v, str):
        return ILLEGAL_CHARACTERS_RE.sub("", v)
    return v


def write_table(ws, df, numeric, dates):
    ws.append([safe(c) for c in df.columns])
    for cell in ws[1]:
        cell.fill, cell.font = HEAD_FILL, HEAD_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row in df.itertuples(index=False):
        ws.append([None if pd.isna(v) else v for v in row])

    id_like = re.compile(r"\b(id|no|num|nr|code|nomer|number|order)\b|№", re.I)

    for j, col in enumerate(df.columns, start=1):
        letter = get_column_letter(j)
        if col in numeric:
            vals = pd.to_numeric(df[col], errors="coerce").dropna()
            whole = len(vals) > 0 and bool((vals % 1 == 0).all())
            if whole and id_like.search(str(col)):
                fmt = "0"            # an identifier: no grouping, no decimals
            elif whole:
                fmt = "#,##0"        # counts stay counts
            else:
                fmt = "#,##0.00"
        elif col in dates:
            fmt = "yyyy-mm-dd"
        else:
            fmt = None
        # A one-character indent on every data cell. Without it the values
        # sit flush against the cell edge, and on a printed sheet with
        # gridlines a right-aligned number touches the left-aligned text
        # beside it: "249.45heating". The two Alignment objects are made
        # once and shared, so this costs almost nothing on a large sheet.
        align = RIGHT_PAD if (col in numeric or col in dates) else LEFT_PAD
        for cell in ws[letter][1:]:
            cell.alignment = align
            if fmt:
                cell.number_format = fmt
        # A column that is entirely empty has no 95th percentile, and
        # int(NaN) raises. That is exactly the column a wrong selector
        # produces, so the crash landed precisely where the tool was
        # supposed to report the problem instead.
        body = df[col].dropna().astype(str)
        q = body.str.len().quantile(0.95) if len(body) else 0
        width = max(len(str(col)) + 4,
                    (int(q) + 3 if pd.notna(q) else 10))
        ws.column_dimensions[letter].width = min(max(width, 10), 42)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    # Print the table and nothing else. The chart's aggregation block sits
    # further right on this sheet; including it in the print range makes
    # "fit to one page wide" squeeze the real columns until the values
    # touch - "249.45heating" - and prints a block of helper numbers the
    # reader did not ask for.
    page_setup(ws, repeat_header=True,
               area=f"A1:{get_column_letter(len(df.columns))}{len(df) + 1}")


def page_setup(ws, repeat_header=False, area=None):
    """Make the sheet printable without anyone touching Page Layout.

    A report that comes out over nine pages with the columns split down the
    middle is half a report. Landscape, one page wide, header repeated on
    every page - the settings a person would have set by hand."""
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    if repeat_header:
        ws.print_title_rows = "1:1"
        # Gridlines on a printed data sheet. Without them a wide table
        # prints as columns of text with no boundaries, and neighbouring
        # values run into each other - a right-aligned number touching a
        # left-aligned one reads as a single figure.
        ws.print_options.gridLines = True
    if area:
        ws.print_area = area


# An identifier is a number but not a quantity. Summing order numbers by
# product gives a chart that means nothing, and the id column is often the
# first numeric one. No chart beats a meaningless chart.
ID_LIKE = re.compile(
    r"\b(id|no|num|nr|code|nomer|number|order|key|ref)\b|№"
    r"|\b(номер|ном|код|артикул|счет|счёт|ид)\b", re.I)

# Columns whose name says they hold money or a count. A chart of these is
# about the business; a chart of "area" or "rate per m2" is about the
# spreadsheet.
MONEY_LIKE = re.compile(
    r"\b(amount|total|sum|revenue|sales|turnover|paid|payment|fee|charge[ds]?|"
    r"cost|price|balance|due|debt|income|profit|qty|quantity|volume|deals)\b"
    r"|\b(сумма|итого|оплачено|начислено|выручка|оборот|долг|остаток|цена|"
    r"стоимость|количество|объем|объём|продаж\w*)\b", re.I)


def pick_measures(df, numeric):
    """Which column or columns a chart should show.

    Two rules, both learned from charts that were technically correct and
    told the reader nothing.

    A column with two or three distinct values across hundreds of rows is a
    tariff or a rate, not a measure - charting it draws the price list.

    And when a table holds two comparable money columns - charged against
    collected, ordered against shipped - the interesting chart is the two
    side by side, not either one alone. They are only paired when they are
    the same size: revenue beside a quantity makes the quantity a flat line
    at the bottom of the axis."""
    pool = [c for c in numeric if not ID_LIKE.search(str(c))]
    if len(df) > 20:
        pool = [c for c in pool if df[c].nunique(dropna=True) > 3] or pool
    if not pool:
        return []

    named = [c for c in pool if MONEY_LIKE.search(str(c))]
    if len(named) >= 2:
        sums = {c: abs(float(pd.to_numeric(df[c], errors="coerce").sum()))
                for c in named}
        top = sorted(named, key=lambda c: -sums[c])[:2]
        big, small = sums[top[0]], sums[top[1]]
        if big > 0 and small / big >= 0.15:
            return sorted(top, key=lambda c: list(df.columns).index(c))
        return [top[0]]
    if named:
        return named
    return [pool[-1]]


def chart_data(ws, df, numeric, dates):
    """Write the aggregated series a chart needs, to the right of the table,
    and return what was written. The numbers have to live on a sheet for a
    chart to reference them; they sit out of the way past a blank column,
    outside the autofilter range."""
    # An identifier is a number but not a quantity. Summing order numbers by
    # product gives a chart that means nothing, and the id column is often
    # the first numeric one. No chart beats a meaningless chart.
    vals = pick_measures(df, numeric)
    if not vals:
        return []
    val = " / ".join(vals)

    side = df.shape[1] + 2
    ws.column_dimensions[get_column_letter(side)].width = 20
    for j in range(len(vals)):
        ws.column_dimensions[get_column_letter(side + 1 + j)].width = 14
    blocks, row = [], 1

    def write(kind, label, keys, series):
        """keys: the category labels. series: one list of numbers per measure,
        in the same order as `vals`."""
        nonlocal row
        ws.cell(row=row, column=side, value=label).font = Font(bold=True)
        for j, name in enumerate(vals):
            ws.cell(row=row, column=side + 1 + j,
                    value=name).font = Font(bold=True)
        for i, k in enumerate(keys):
            r = row + 1 + i
            ws.cell(row=r, column=side, value=str(k))
            for j, col in enumerate(series):
                c = ws.cell(row=r, column=side + 1 + j, value=float(col[i]))
                c.number_format = "#,##0.00"
        blocks.append({"kind": kind, "label": label, "value": val,
                       "head": row, "first": row + 1,
                       "last": row + len(keys), "col": side,
                       "ncols": len(vals)})
        row += len(keys) + 3

    # Which column to group by. Taking the one with the FEWEST distinct
    # values always lands on "status" or "payment method" - three bars that
    # say nothing. Taking the richest column that still fits on an axis
    # lands on product, category or region, which is what a report is for.
    text_cols = [c for c in df.columns
                 if c not in numeric and c not in dates
                 and 2 <= df[c].nunique() <= 20]
    if text_cols:
        key = max(text_cols, key=lambda c: (df[c].nunique(),
                                            -list(df.columns).index(c)))
        agg = df.groupby(key, dropna=True)[vals].sum()
        agg = agg.sort_values(vals[0], ascending=False).head(12)
        write("bar", f"{val} by {key}", list(agg.index),
              [agg[c].tolist() for c in vals])

    if dates:
        tmp = df[[dates[0]] + vals].dropna()
        if len(tmp) > 2:
            m = tmp.set_index(dates[0])[vals].resample("MS").sum()
            m = m[m.index.notna()]
            if 2 <= len(m) <= 60:
                write("line", f"{val} by month",
                      [k.strftime("%Y-%m") for k in m.index],
                      [m[c].tolist() for c in vals])
    return blocks


def place_charts(summary, sheets):
    """All the charts go on the Summary sheet, under the cleaning table.

    Put them beside the data instead and a wide table pushes them off the
    screen, which is the same failure as putting them underneath it. The
    Summary is the sheet that opens first and it is nearly empty, so that is
    where a reader will actually meet them.
    """
    if not any(b for _, b in sheets):
        return 0

    # Which column the right-hand chart starts in. Anchoring it to a fixed
    # column letter leaves a gap as wide as the summary table's columns
    # happen to be - on a wide table the two charts end up an arm's length
    # apart. Walk the actual widths instead and start where the left chart
    # ends. A width unit is one character, about seven pixels.
    need = CHART_W_CM / 2.54 * 96 / 7.0
    total, right_col = 0.0, 10
    for j in range(1, 80):
        w = summary.column_dimensions[get_column_letter(j)].width or 8.43
        total += w
        if total >= need + 1:
            right_col = j + 1
            break

    top = summary.max_row + 3
    summary.cell(row=top, column=1, value="Charts").font = Font(bold=True,
                                                                size=13)
    top += 2
    col, i = 1, 0
    for ws, blocks in sheets:
        for b in blocks:
            ch = BarChart() if b["kind"] == "bar" else LineChart()
            ch.style = 10 if b["kind"] == "bar" else 12
            if b["kind"] == "bar":
                ch.type = "col"
            n = b.get("ncols", 1)
            ch.title = f'{b["label"]}  ({ws.title})'
            # No axis titles. The chart title already names the measure, and
            # a rotated title on a chart this narrow is drawn straight
            # through the tick numbers - "Monthly fee / Paid" printed over
            # 30 000, 40 000, 50 000.
            ch.y_axis.title, ch.x_axis.title = None, None
            ch.height, ch.width = CHART_H_CM, CHART_W_CM

            # Excel hides both axes unless it is told not to. openpyxl leaves
            # `delete` unset, Excel reads that as "delete", and the chart
            # opens with bars and no labels - no category names under them
            # and no numbers beside them. A chart nobody can read is worse
            # than no chart, because it looks finished.
            for ax in (ch.x_axis, ch.y_axis):
                ax.delete = False
                ax.tickLblPos = "nextTo"
                ax.majorTickMark = "out"
            ch.y_axis.numFmt = "#,##0"
            # One series needs no legend - the title already names it. Two
            # series are useless without one.
            if n == 1:
                ch.legend = None
            else:
                ch.legend.position = "b"
                # Without this the legend is drawn ON the plot, straight
                # across the category labels - "waste" disappeared behind
                # "Monthly fee / Paid". Excel only keeps them apart when it
                # is told the legend does not overlay.
                ch.legend.overlay = False
                if b["kind"] == "bar":
                    ch.grouping, ch.overlap = "clustered", -10
            ch.add_data(Reference(ws, min_col=b["col"] + 1,
                                  max_col=b["col"] + n, min_row=b["head"],
                                  max_row=b["last"]), titles_from_data=True)
            ch.set_categories(Reference(ws, min_col=b["col"],
                                        min_row=b["first"], max_row=b["last"]))

            # Two more things Excel needs said out loud.
            #
            # set_categories always writes a NUMERIC reference, whatever the
            # cells hold. These hold names - "water", "Ashgabat" - so Excel
            # finds no numbers and labels the bars 1, 2, 3 instead.
            #
            # And openpyxl puts the category axis on the LEFT of a column
            # chart, where the value axis already is. It belongs underneath.
            letter = get_column_letter(b["col"])
            ref = (f"'{ws.title}'!${letter}${b['first']}"
                   f":${letter}${b['last']}")
            for s in ch.series:
                s.cat = AxDataSource(strRef=StrRef(f=ref))
            ch.x_axis.axPos = "b"
            col = 1 if i % 2 == 0 else right_col
            anchor = f"{get_column_letter(col)}{top + (i // 2) * ROW_STEP}"
            summary.add_chart(ch, anchor)
            i += 1
    # The row the last chart reaches, so the caller can set a print range
    # that actually contains them. max_row only knows about cells.
    return top + ((i - 1) // 2) * ROW_STEP + ROW_STEP if i else top


def totals(records):
    """The six numbers a client looks at before reading anything else.

    Combined sheets are excluded: their rows were already counted once in
    the tables they came from, and counting them twice would make the
    headline band disagree with the table under it."""
    records = [r for r in records if not r.get("merged")]

    def s(key):
        return sum(int(r.get(key) or 0) for r in records)
    return [
        ("source files", len({r.get("file") for r in records})),
        ("tables", sum(1 for r in records if r.get("sheet")
                       not in (None, "", "(skipped)"))),
        ("rows read", s("rows read")),
        ("rows kept", s("rows kept")),
        ("duplicates removed", s("duplicate rows removed")),
        ("empty rows dropped", s("empty rows dropped")),
        ("empty columns dropped", s("empty columns dropped")),
        ("columns repaired", s("columns to numbers") + s("columns to dates")),
    ]


def band_height(ws, band, font_size=9, line=11.5, pad=7, floor=26):
    """How tall row 4 must be for every headline label to be readable.

    A wrapped label needs one line per (roughly) column-width characters,
    and the word cannot be split, so a long word forces its own line.
    Excel's own autofit does not run on a file written by a program, which
    is why this is computed rather than left to the reader's Excel."""
    most = 1
    for j, (label, _) in enumerate(band, start=1):
        width = ws.column_dimensions[get_column_letter(j)].width or 12
        # Deliberately pessimistic: the column width is measured in
        # characters of the default font, and a bold label is wider per
        # character. A row one line too tall costs nothing; a label cut in
        # half is what the client sees first.
        per_line = max(int(width), 1)
        lines, current = 1, 0
        for word in str(label).split():
            need = len(word) if current == 0 else current + 1 + len(word)
            if need <= per_line:
                current = need
            else:
                lines += 1
                current = len(word)
            # a single word longer than the column wraps inside itself
            while current > per_line:
                lines += 1
                current -= per_line
        most = max(most, lines)
    return max(floor, most * line + pad)


def write_summary(ws, records):
    ws["A1"] = "CSV TO EXCEL — cleaning summary"
    ws["A1"].font = Font(bold=True, size=16)
    ws["A2"] = ("Every number below is a change made to the source data. "
                "Nothing was altered silently.")
    ws["A2"].font = Font(italic=True, color="595959")

    # ---- the headline band. A reader who opens the workbook and reads one
    # thing reads this, so it carries the totals rather than a decoration.
    band = totals(records)
    for j, (label, value) in enumerate(band, start=1):
        lab = ws.cell(row=4, column=j, value=label)
        lab.font = Font(color="FFFFFF", bold=True, size=9)
        lab.fill = HEAD_FILL
        lab.alignment = Alignment(horizontal="center", wrap_text=True)
        val = ws.cell(row=5, column=j, value=value)
        val.font = Font(bold=True, size=14, color="1F3864")
        val.fill = PatternFill("solid", fgColor="DCE6F1")
        val.alignment = Alignment(horizontal="center")
        val.number_format = "#,##0"
        lab.border = val.border = BORDER
    ws.row_dimensions[4].height = 26
    ws.row_dimensions[5].height = 22

    head_row = 7
    cols = ["file", "sheet", "encoding", "delimiter", "header row", "rows read",
            "rows kept", "duplicate rows removed", "empty rows dropped",
            "empty columns dropped", "columns to numbers", "columns to dates",
            "chart", "warning"]
    for j, c in enumerate(cols, start=1):
        cell = ws.cell(row=head_row, column=j, value=c)
        cell.fill, cell.font = HEAD_FILL, HEAD_FONT
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    for i, r in enumerate(records):
        row = head_row + 1 + i
        for j, c in enumerate(cols, start=1):
            cell = ws.cell(row=row, column=j, value=r.get(c, ""))
            cell.border = BORDER
            if i % 2:
                cell.fill = PatternFill("solid", fgColor="F2F5FA")
    for j, c in enumerate(cols, start=1):
        letter = get_column_letter(j)
        widest = max([len(str(r.get(c, ""))) for r in records] or [0])
        ws.column_dimensions[letter].width = min(
            max(len(c) + 3, widest + 3, 12), 40)
        ws.cell(row=head_row, column=j).border = BORDER

    # The headline band sits in the same columns as the table below, but its
    # labels are not the table's headers - "empty columns dropped" can land
    # in a column sized for "rows kept". Row 4 was a fixed 26 points, so the
    # label was cut off and the client read "columns dropped" in the most
    # prominent part of the report. Found on 17.09.2026 by opening the file
    # and looking at it; no test sees this.
    ws.row_dimensions[4].height = band_height(ws, band)
    ws.freeze_panes = f"A{head_row + 1}"


# ---------------------------------------------------------------- merging
# Added 17.09.2026. Four jobs in a 75-posting market survey needed exactly
# this and nothing else: "several workbooks, same headers, merged into one
# sheet". The feature was written down as missing on 16.09 and postponed
# because nobody knew whether it was wanted. It was.

def merge_key(df):
    """Two tables merge only when their column names match exactly. Order
    does not matter, spelling does."""
    return tuple(sorted(str(c) for c in df.columns))


def source_column_name(cols):
    """The source column must not collide with a column the client already
    has. Overwriting their own 'source' column would be exactly the kind of
    quiet damage this tool exists to prevent."""
    taken = {str(c).strip().lower() for c in cols}
    for name in ("source", "source file", "source table", "origin"):
        if name not in taken:
            return name
    n = 2
    while f"source_{n}" in taken:
        n += 1
    return f"source_{n}"


def describe_difference(cols, ref):
    """Why a table was left out, in words the client can act on."""
    cols = {str(c) for c in cols}
    if not ref:
        return "no other table has the same columns"
    missing = sorted(ref - cols)
    extra = sorted(cols - ref)
    bits = []
    if missing:
        bits.append("missing " + ", ".join(missing[:4])
                    + (" ..." if len(missing) > 4 else ""))
    if extra:
        bits.append("extra " + ", ".join(extra[:4])
                    + (" ..." if len(extra) > 4 else ""))
    if not bits:
        return "columns match but no second table shares them"
    return "not merged - " + "; ".join(bits)


def merge_tables(tables):
    """Combine tables that share a column set into one frame each, with a
    column naming where every row came from.

    STRICT BY DESIGN. The columns must match exactly. A near match - eleven
    of twelve columns the same - is NOT merged, because guessing which
    column corresponds to which is how two months of a client's data get
    silently mixed. What did not merge is reported instead, with the
    difference named, so the client decides rather than the script.

    Types are worked out again on the combined column, not inherited: a
    column that was numeric in one file and text in another must not end up
    as text for everybody just because it was concatenated.

    Rows that repeat across sources are counted but NOT removed. The same
    invoice legitimately appears in two monthly exports; deleting it would
    be a judgement the client never asked for.

    Returns (merged, leftovers):
      merged    - list of dicts: frame, labels, source, numeric, dates,
                  repeats
      leftovers - list of (label, reason)
    """
    groups = {}
    for label, df in tables:
        if df is None or df.empty or not len(df.columns):
            continue
        groups.setdefault(merge_key(df), []).append((label, df))
    if not groups:
        return [], []

    # The reference set is the largest real group; everything else is
    # described as a difference from it, which is what a client wants to
    # read.
    real = [m for m in groups.values() if len(m) > 1]
    ref = set(str(c) for c in real[0][0][1].columns) if real else set()
    if real:
        ref = set(str(c) for c in max(real, key=len)[0][1].columns)

    merged, leftovers = [], []
    for members in sorted(groups.values(), key=lambda m: -len(m)):
        if len(members) < 2:
            label, df = members[0]
            leftovers.append((label, describe_difference(df.columns, ref)))
            continue

        src = source_column_name(members[0][1].columns)
        parts = []
        for label, df in members:
            piece = df.copy()
            piece.insert(0, src, str(label))
            parts.append(piece)
        out = pd.concat(parts, ignore_index=True)

        numeric, dates = [], []
        for c in out.columns:
            if c == src:
                continue
            conv, _ = coerce_dates(out[c])
            if conv is not None:
                out[c] = conv
                dates.append(c)
                continue
            conv, _ = coerce_numeric(out[c])
            if conv is not None:
                out[c] = conv
                numeric.append(c)
            else:
                out[c] = (out[c].astype(str).str.strip()
                          .replace("<NA>", "").replace("nan", ""))

        # Refuse loudly instead of writing a sheet Excel will truncate.
        # A client who paid for a merge and received a file missing rows
        # would find out from their own totals, weeks later.
        if len(out) > EXCEL_MAX_ROWS - 1:
            leftovers.append((", ".join(l for l, _ in members),
                              f"not merged - {len(out):,} rows exceeds the "
                              f"{EXCEL_MAX_ROWS:,}-row limit of one Excel "
                              f"sheet; the source sheets are unchanged"))
            continue

        body = [c for c in out.columns if c != src]
        repeats = int(out.duplicated(subset=body).sum()) if body else 0
        merged.append({"frame": out, "labels": [l for l, _ in members],
                       "source": src, "numeric": numeric, "dates": dates,
                       "repeats": repeats})
    return merged, leftovers


def build(paths, out, dedupe=True, merge=False):
    from openpyxl import Workbook
    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"

    records, used, chart_sheets = [], {"summary"}, []
    for_merge = []
    for p in paths:
        print(f"  reading {Path(p).name} ...", end=" ")
        try:
            tables = load_sheets(p, dedupe=dedupe)
        except RuntimeError as e:
            # One unreadable file must not cost the client the other nine.
            print(f"SKIPPED: {e}")
            records.append({"file": Path(p).name, "sheet": "(skipped)",
                            "rows read": 0, "rows kept": 0,
                            "chart": "", "warning": str(e)})
            continue

        kept = []
        for label, df, notes, numeric, dates in tables:
            if df is None or df.empty:
                notes["sheet"] = "(skipped)"
                records.append(notes)
                continue
            name = sheet_name(label, used)
            ws = wb.create_sheet(name)
            write_table(ws, df, numeric, dates)
            blocks = chart_data(ws, df, numeric, dates)
            chart_sheets.append((ws, blocks))
            notes["chart"] = ", ".join(b["label"] for b in blocks)
            notes["sheet"] = name
            records.append(notes)
            kept.append((name, len(df)))
            for_merge.append((name, df))
        if not kept:
            print("no data rows, skipped")
        elif len(kept) == 1:
            print(f"{kept[0][1]:,} rows -> sheet '{kept[0][0]}'")
        else:
            print(f"{sum(n for _, n in kept):,} rows across "
                  f"{len(kept)} sheets -> "
                  + ", ".join(f"'{n}'" for n, _ in kept))
        # A warning belongs on its own line, after the count, where it can
        # be read. Buried at the end of the reading line it is missed.
        # A sheet that was already reported as skipped needs no second line
        # saying the same thing.
        for label, frame, notes, _, _ in tables:
            if notes.get("warning") and frame is not None and not frame.empty:
                print(f"    WARNING  {label}: {notes['warning']}")

    if merge:
        merged, leftovers = merge_tables(for_merge)
        if not merged:
            print("\n  --merge: nothing to combine - no two tables share "
                  "the same columns")
        for k, g in enumerate(merged):
            name = sheet_name("Combined" if len(merged) == 1
                              else f"Combined {k + 1}", used)
            ws = wb.create_sheet(name)
            write_table(ws, g["frame"], g["numeric"], g["dates"])
            blocks = chart_data(ws, g["frame"], g["numeric"], g["dates"])
            chart_sheets.append((ws, blocks))
            note = (f"{len(g['labels'])} tables with identical columns, "
                    f"joined; '{g['source']}' names the origin of each row")
            if g["repeats"]:
                note += (f"; {g['repeats']:,} rows repeat across sources - "
                         f"kept, not removed")
            records.append({"merged": True, "file": "(combined)",
                            "sheet": name, "encoding": "",
                            "delimiter": ", ".join(g["labels"][:6])
                            + (" ..." if len(g["labels"]) > 6 else ""),
                            "header row": "", "rows kept": len(g["frame"]),
                            "chart": ", ".join(b["label"] for b in blocks),
                            "warning": note})
            print(f"  combined {len(g['labels'])} tables -> "
                  f"'{name}', {len(g['frame']):,} rows")
        for label, reason in leftovers:
            records.append({"merged": True, "file": "(not combined)",
                            "sheet": label, "warning": reason})
        # The combined sheet is the deliverable; it belongs next to the
        # Summary, not after nine source sheets the client scrolls past.
        for k in range(len(merged), 0, -1):
            wb.move_sheet(wb.worksheets[-1], offset=-(len(wb.worksheets) - 2))

    write_summary(summary, records)
    bottom = place_charts(summary, chart_sheets)
    # The Summary's print range stops where the charts stop; without one,
    # Excel prints every empty column the charts happen to touch, and with
    # one that is too short it cuts the charts off.
    page_setup(summary, area=f"A1:R{max(bottom, summary.max_row)}")
    wb.save(out)
    return records


def collect_inputs(items):
    """Turn the command line into the list of files that will be read.

    Kept separate from main() so the rule about lock files is covered by
    the test suite rather than by whoever runs it next."""
    paths = []
    for item in items:
        p = Path(item)
        if p.is_dir():
            paths += sorted(q for q in p.iterdir()
                            if q.suffix.lower() in SUFFIXES and not is_junk(q))
        elif p.is_file():
            if is_junk(p):
                print(f"  ignored: {p.name} is a lock file left by an open "
                      f"editor, not data")
            else:
                paths.append(p)
        else:
            print(f"  not found: {item}")
    return paths


def main():
    ap = argparse.ArgumentParser(
        description="Clean a folder of CSV exports into one Excel workbook.")
    ap.add_argument("inputs", nargs="+", help="CSV/TSV files, or a folder")
    ap.add_argument("-o", "--output", default="report.xlsx")
    ap.add_argument("--no-dedupe", action="store_true",
                    help="keep duplicate rows instead of removing them")
    ap.add_argument("--merge", action="store_true",
                    help="also write a Combined sheet joining every table "
                         "whose columns match exactly, with a column naming "
                         "the source of each row")
    a = ap.parse_args()

    paths = collect_inputs(a.inputs)
    if not paths:
        print("No input files found.")
        sys.exit(1)

    print(f"\n{len(paths)} file(s) -> {a.output}\n")
    records = build(paths, a.output, dedupe=not a.no_dedupe,
                    merge=a.merge)

    # Combined sheets are excluded here for the same reason as in
    # totals(): their rows were counted once already, in the tables
    # they were built from.
    src_records = [r for r in records if not r.get("merged")]
    total_in = sum(r.get("rows read", 0) for r in src_records)
    total_out = sum(r.get("rows kept", 0) for r in src_records)
    dupes = sum(r.get("duplicate rows removed", 0) for r in src_records)
    print(f"\n  {total_in:,} rows read, {total_out:,} kept, "
          f"{dupes:,} duplicates removed")
    print(f"  written to {a.output}")
    print("  the Summary sheet lists every change, file by file.\n")


if __name__ == "__main__":
    main()
