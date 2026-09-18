# csv-excel-cleaner

Turns a folder of inconsistent exports — CSV, TSV or Excel workbooks —
into one clean Excel report, and writes down every change it made.

The point is not the cleaning. The point is the **Summary sheet**: for
every file it records the encoding, the delimiter or sheet name, which
row the header was really on, and every alteration made to the data. If
eleven rows were dropped, you can see that eleven rows were dropped and
decide whether that was right.

Nothing is changed silently.

```bash
python csv_to_excel.py data/ -o report.xlsx
```

```
4 file(s) -> report.xlsx

  reading billing_march.xlsx ... 1,683 rows across 4 sheets -> 'Block A', 'Block B', 'Block C', 'Block D'
  reading oborot_1c.csv ... 900 rows -> sheet 'oborot_1c'
  reading orders_webshop.csv ... 1,400 rows -> sheet 'orders_webshop'
  reading regions_q1.tsv ... 320 rows -> sheet 'regions_q1'

  4,343 rows read, 4,303 kept, 40 duplicates removed
  written to report.xlsx
```

---

## What it handles

Real exports are not tidy. These are the cases the tool was built
against, each one a bug that reached the test suite:

| problem | what happens |
|---|---|
| Every file a different delimiter | detected per file — comma, semicolon, tab, pipe |
| One file opens as mojibake | encoding detected per file, CP1251 and UTF-8 both read |
| Three junk lines above the header | the real header row is found and reported |
| A merged title over group headers in a workbook | the group header does not win over the real one |
| Amounts stored as text: `1 234,56`, `1,234.56`, `$1,234` | become real numbers |
| Dates as `03.02.2026` | read day-first, as Central Asia and Europe write them |
| Dates as `2026-01-03` | read ISO, **not** flipped to 3 March |
| A totals column that looks empty | it holds formulas with no cached value — named, not dropped |
| An old `.xls`, a renamed PNG, a zip with the wrong extension | refused with a sentence saying what to do |
| One file in the folder is unreadable | it is skipped with the reason recorded; the other nine still run |
| A column of postal codes, phones or barcodes | kept as text — a leading zero is never dropped and a phone is never formatted as money |
| A file is open in Excel while it runs | the hidden `~$` lock file beside it is ignored, not counted and not reported as a failure |

Every sheet of a workbook is read as its own table.

---

## Joining files

```bash
python csv_to_excel.py data/ --merge
```

Adds a **Combined** sheet holding every table whose columns match, with
a first column naming which file each row came from.

```
  combined 4 tables -> 'Combined', 1,683 rows
```

**The columns must match exactly.** Order does not matter, spelling
does. This is deliberate: guessing that `Amount` in one file means `Sum`
in another is how two months of data get quietly mixed, and nobody finds
out until the totals are already wrong. A table that does not match is
left as its own sheet and named in the Summary with the difference
spelled out:

```
  regions_q1   not merged - missing Area m2, Balance, Monthly fee, No ...; extra amount, deals, manager, region
```

Duplicates are treated differently inside and across files. A row
repeated **inside** one file is removed and counted. A row appearing in
**two** files is counted but kept — the same invoice legitimately turns
up in two monthly exports, and deleting it is not the script's decision.

---

## What it does not do

Stated plainly, because a tool that hides its limits costs more than one
that names them:

- It does not merge files whose columns are named differently. Renaming
  is a human decision and stays one.
- It does not build pivot tables or custom totals.
- It does not join tables on a key, the way `VLOOKUP` does.
- It does not reorganise data by business logic — one tab per month, and
  so on.

---

## Tested

```bash
python test_csv_to_excel.py
```

**167 checks**, no network and no fixtures on disk — every case is built
in memory and asserted against a known answer.

Thirty-five of them are marked `REGRESSION`. Each is a bug that was in this
code and shipped nothing, because it was caught here. None of them
crashed. Every one finished, wrote a workbook that looked correct, and
was wrong — which is the only kind of failure that reaches a user
without being noticed.

A few, so the word is not just decoration:

- A data row taken for the header.
- A label column turned into numbers.
- An order number read as a year.
- An ISO date read day-first, so 3 January became 1 March.
- **A money column read as dates**: `4141.98` became the year 4141, and
  a float `4821.55` became 1 January 1970 plus nanoseconds. One dot
  between digits was enough to make the parser try. Two separators are
  now required, and a column of plain numbers is vetoed outright.
- Charts drawn by openpyxl with no axis labels at all, because three
  independent defaults each remove them.
- A postal code `05401` turned into the number 5401, losing its leading
  zero for good; a phone `8025285988` rendered as `8,025,285,988`, money
  formatting on a telephone; a barcode turned into a float. An identifier
  is not a measure, and a leading zero or a fixed-width run of digits says
  so.
- A size written `10.5.2` read as 10 May 2002, and an article
  number `10.20.30` as 20 October 2030. Two separators are necessary for
  a date but not sufficient: with dots, the year must be four digits.
- A headline label cut in half — `empty columns dropped` rendered as
  `columns dropped` — because the summary band shares its columns with
  the table below it and the row height was a fixed 26 points. The
  number under it was right. No test sees a label that does not fit;
  this one was found by opening the file and looking.
- A hidden `~$` lock file, which Excel keeps beside any workbook it has
  open, counted as a sixth input file and then reported to the client as
  a file that failed to open. The report was fine; the Summary accused a
  file the client cannot even see.

The suite above covers correctness. A second one covers survival:

```bash
python adversarial.py
```

**27 deliberately hostile inputs** — a zero-byte file, a file of only
newlines, NUL bytes in the middle, a line a million characters long,
ragged rows, a quote that is never closed, a PNG renamed to `.csv`,
1000 columns, a filename full of characters Excel forbids, random bytes
wearing an `.xlsx` name, a workbook of thirty sheets.

The bar is not "produces a good result". The bar is: **it must not
crash, and it must not silently produce a wrong one.** Either finish the
job, or fail with a message saying what was wrong. All 27 pass.

---

## Measured, not estimated

On a 16 GB Windows machine, every run checked against the source to the
cent:

| rows | files | time |
|---|---|---|
| 10,000 | 5 | 5 seconds |
| 100,000 | 5 | 49 seconds |
| 250,000 | 5 | 2 minutes |
| 500,000 | 10 | 4 minutes |

Memory is the limit, not speed: half a million rows needs about 3.5 GB
free. Above 1,048,576 rows a merge is refused rather than truncated —
that is Excel's own ceiling for one sheet.

---

## Install and run

```bash
pip install -r requirements.txt
python csv_to_excel.py data/ -o report.xlsx
```

Python 3.9 or newer. Two dependencies: `pandas` 2.0+ and `openpyxl`.
`pandas` 2.0 is the floor, not a preference: the date parser uses
`format="mixed"`, which does not exist before it.

```
python csv_to_excel.py data/                  everything in a folder
python csv_to_excel.py a.csv b.xlsx           named files
python csv_to_excel.py data/ -o march.xlsx    choose the output name
python csv_to_excel.py data/ --no-dedupe      keep duplicate rows
python csv_to_excel.py data/ --merge          add a Combined sheet
```

### Try it without your own data

```bash
python make_messy_data.py
python csv_to_excel.py demo_data/ -o report.xlsx --merge
```

`make_messy_data.py` generates deliberately awful sample files —
different delimiters, mixed encodings, merged titles, amounts stored as
text, phone numbers crammed into one cell.

---

## Why this exists

I sell this as a service, and the Summary sheet is the reason clients
come back: they can check the work instead of trusting it.

Everything in this repository is a work sample. If you want to use it
for something, ask me.

Written up: [pandas turned my money column into dates, and 123 tests said nothing](https://dev.to/kerimff7rgb/pandas-turned-my-money-column-into-dates-and-123-tests-said-nothing-3gbh)

**Annamyrat Hallyyev** — Python, data cleaning, Excel automation.
