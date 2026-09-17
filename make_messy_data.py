"""
Builds the three files a small online shop actually ends up with.

The scenario is the ordinary one. A shop sells hardware online. Every month
the owner needs one report, and the numbers live in three places:

  orders_webshop.csv    the web platform's export - comma separated, with
                        two lines of title above the header nobody removes,
                        revenue formatted with thousands separators so it
                        will not sum, and a block of rows duplicated
                        because the export was run twice

  oborot_1c.csv         the accountant's export - semicolons, Windows-1251,
                        Russian column names, comma as the decimal mark,
                        and an empty "note" column nobody ever filled

  regions_q1.tsv        the sales manager's own file - tab separated, blank
                        lines where he grouped things by eye, amounts typed
                        with a dollar sign, and a column of notes that is
                        entirely empty

Every defect here is one that turns up in real exports. The point of the
fixture is that the cleaner has to be shown working on mess.
"""
import os
import random
from datetime import date, timedelta

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "demo_data")
os.makedirs(OUT, exist_ok=True)
rng = random.Random(20260916)

PRODUCTS = [
    ("Cordless Drill 18V", "Power tools", 189.00),
    ("Impact Driver", "Power tools", 142.50),
    ("Angle Grinder 125mm", "Power tools", 96.00),
    ("Socket Set 108pc", "Hand tools", 78.40),
    ("Torque Wrench", "Hand tools", 134.90),
    ("Measuring Tape 8m", "Hand tools", 11.20),
    ("Safety Goggles", "Protection", 8.75),
    ("Work Gloves", "Protection", 6.40),
    ("Steel Toe Boots", "Protection", 87.00),
    ("LED Work Light", "Site equipment", 54.30),
    ("Extension Reel 25m", "Site equipment", 41.60),
    ("Tool Chest", "Site equipment", 260.00),
]
CITIES_EN = ["Ashgabat", "Istanbul", "Dubai", "Almaty", "Baku", "Tbilisi"]
CITIES_RU = ["Ашхабад", "Стамбул", "Дубай", "Алматы", "Баку", "Тбилиси"]
PAY = ["card", "bank transfer", "cash on delivery"]
STATUS = ["shipped", "shipped", "shipped", "returned", "cancelled"]


def w(name, text, encoding="utf-8"):
    with open(os.path.join(OUT, name), "w", encoding=encoding, newline="") as f:
        f.write(text)
    print(f"  {name:22s} {encoding:10s} {len(text):>8,} bytes")


def season(d):
    """Sales climb toward the end of the period, the way a shop's do. A flat
    series makes a flat chart, and a flat chart illustrates nothing."""
    return 0.75 + 0.55 * (d.month - 1) / 5.0


# ------------------------------------------------------------------ file 1
# The web platform's export: two title lines, a blank line, then the header.
rows = ["Sales export - Webshop", "Generated automatically, do not edit", ""]
rows.append("Order ID,Date ,Customer,City,Product,Category,Qty, Revenue ,"
            "Payment,Status")

body, d0 = [], date(2026, 1, 1)
for i in range(1400):
    # Jan 1 to Jun 30 exactly. A series that stops on the 2nd of a month
    # draws a cliff on the monthly chart - honest, and read as a bug.
    d = d0 + timedelta(days=(i * 181) // 1400)
    name, cat, price = PRODUCTS[rng.randrange(len(PRODUCTS))]
    qty = rng.choices([1, 1, 2, 3, 5], [5, 4, 3, 2, 1])[0]
    rev = price * qty * rng.uniform(0.92, 1.08) * season(d)
    body.append(
        f"{50000+i},{d:%Y-%m-%d},Customer {100+i%430},"
        f"{rng.choice(CITIES_EN)},{name},{cat},{qty},"
        f"\"{rev:,.2f}\",{rng.choice(PAY)},{rng.choice(STATUS)}")

# the export was run twice and a block of rows came through again
body[620:620] = body[580:620]
rows.extend(body)
w("orders_webshop.csv", "\n".join(rows) + "\n")


# ------------------------------------------------------------------ file 2
# The accountant's export from the bookkeeping system.
rows = ["Номер;Дата;Город;Категория;Сумма;Примечание;"]
d0 = date(2026, 2, 1)
for i in range(900):
    d = d0 + timedelta(days=(i * 120) // 900)      # Feb 1 to May 31 exactly
    _, cat, price = PRODUCTS[rng.randrange(len(PRODUCTS))]
    amount = f"{price * rng.uniform(1, 4) * season(d):.2f}".replace(".", ",")
    rows.append(f"{70000+i};{d:%d.%m.%Y};{rng.choice(CITIES_RU)};"
                f"{cat};{amount};;")
w("oborot_1c.csv", "\n".join(rows) + "\n", encoding="cp1251")


# ------------------------------------------------------------------ file 3
# The sales manager's own file, kept by hand.
rows = ["region\tmanager\tdeals\tamount\tnotes"]
for i in range(320):
    if i % 13 == 0:
        rows.append("")                 # he leaves a blank line between groups
    _, _, price = PRODUCTS[rng.randrange(len(PRODUCTS))]
    amount = price * rng.uniform(2, 9)
    rows.append(f"{rng.choice(CITIES_EN)}\tManager {i % 6 + 1}\t"
                f"{rng.randint(1, 18)}\t$ {amount:,.2f}\t")
w("regions_q1.tsv", "\n".join(rows) + "\n")

print()


# ------------------------------------------------------------------ file 4
# The shape client data actually arrives in: a workbook, not a CSV. A title
# above the header, a printed-on line, a blank row, a merged group header,
# and only then the real column names - on two sheets, because the register
# is kept per building.
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

SERVICES = ["cleaning", "water", "heating", "lift", "waste"]
FIRST = ["Aman", "Merdan", "Gözel", "Bahar", "Serdar", "Maral", "Kerim",
         "Jemal", "Batyr", "Oguljan", "Dövlet", "Selbi"]
LAST = ["Amanow", "Berdiyew", "Gurbanow", "Hojayew", "Myradow", "Ashirow",
        "Nurgeldiyew", "Ataýewa", "Kurbanowa", "Saparow"]

HEAD = ["No", "Tenant", "Area m2", "Rate per m2", "Monthly fee",
        "Previous balance", "Paid", "Balance", "Service", "Payment date",
        "Payment", "Phone", "Note"]
NC = len(HEAD)

BLOCKS = [("Block A", 430), ("Block B", 385), ("Block C", 466),
          ("Block D", 402)]

wb = Workbook()
total_rows = 0
for si, (block, n) in enumerate(BLOCKS):
    ws = wb.active if si == 0 else wb.create_sheet()
    ws.title = block

    ws["A1"] = f"SERVICE CHARGES REGISTER — {block} — March 2026"
    ws["A1"].font = Font(bold=True, size=13)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=NC)
    ws["A2"] = "printed 01.04.2026"
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=NC)

    # a merged group header sitting above the real one
    ws.cell(row=4, column=3, value="Rates").font = Font(bold=True)
    ws.merge_cells(start_row=4, start_column=3, end_row=4, end_column=4)
    ws.cell(row=4, column=5, value="Charges, TMT").font = Font(bold=True)
    ws.merge_cells(start_row=4, start_column=5, end_row=4, end_column=8)
    for c in (3, 5):
        ws.cell(row=4, column=c).alignment = Alignment(horizontal="center")

    for j, h in enumerate(HEAD, start=1):
        ws.cell(row=5, column=j, value=h).font = Font(bold=True)

    r = 6
    for i in range(1, n + 1):
        area = round(rng.uniform(28, 96), 1)
        rate = rng.choice([10.57, 10.57, 10.57, 12.40])
        fee = round(area * rate, 2)
        prev = round(rng.choice([0, 0, 0, rng.uniform(-400, 900)]), 2)
        paid = round(fee * rng.choice([0, 0.5, 1, 1, 1, 1.2]), 2)
        pay_day = date(2026, 3, rng.randint(1, 28))
        ws.cell(row=r, column=1, value=i)
        ws.cell(row=r, column=2,
                value=f"{rng.choice(LAST)} {rng.choice(FIRST)}")
        ws.cell(row=r, column=3, value=area)
        ws.cell(row=r, column=4, value=rate)
        # the amount columns are TEXT, formatted by hand - so they will not sum
        ws.cell(row=r, column=5, value=f"{fee:,.2f}")
        ws.cell(row=r, column=6, value=f"{prev:,.2f}")
        ws.cell(row=r, column=7, value=f"{paid:,.2f}")
        ws.cell(row=r, column=8, value=f"{fee + prev - paid:,.2f}")
        ws.cell(row=r, column=9, value=rng.choice(SERVICES))
        # the date is text in dd.mm.yyyy, which Excel will not sort as a date
        ws.cell(row=r, column=10, value=f"{pay_day:%d.%m.%Y}")
        ws.cell(row=r, column=11, value=rng.choice(PAY))
        # several numbers crammed into one cell, as people do
        tel = (f"8{rng.randint(10,99)} {rng.randint(10,99)}"
               f"-{rng.randint(10,99)}-{rng.randint(10,99)}")
        if i % 7 == 0:
            tel += (f"   6{rng.randint(10,99)}-{rng.randint(10,99)}"
                    f"-{rng.randint(10,99)}")
        ws.cell(row=r, column=12, value=tel)
        if i % 11 == 0:
            ws.cell(row=r, column=13, value=rng.choice(
                ["moved out", "disputes the charge", "pays in person", "-"]))
        r += 1
    total_rows += n

    for j in range(1, NC + 1):
        ws.column_dimensions[get_column_letter(j)].width = \
            max(10, len(HEAD[j - 1]) + 4)

path = os.path.join(OUT, "billing_march.xlsx")
wb.save(path)
print(f"  {'billing_march.xlsx':22s} {'xlsx':10s} "
      f"{os.path.getsize(path):>8,} bytes   "
      f"({len(BLOCKS)} sheets, {total_rows:,} rows, merged header)")
