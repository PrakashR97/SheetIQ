from openpyxl import Workbook
from datetime import datetime
from pathlib import Path
import random

random.seed(42)

output_file = Path("test_data.xlsx")
wb = Workbook()

# Sheet 1: Customers (40 data rows)
ws = wb.active
ws.title = "Customers"
ws.append(["Customer ID", "Customer Name", "Email", "Amount", "Join Date", "Status"])

names = ["Alice", "Bob", "Charlie", "David", "Priya", "Kumar", None, "   "]
statuses = ["Active", "Inactive", "Completed", "UNKNOWN_STATUS"]

for i in range(1, 41):
    customer_id = f"{i:05d}"
    name = random.choice(names)
    email = f"{name.lower()}@example.com" if name and name.strip() else None
    amount = random.choice([100.50, 250, 0, -500, None, "1,200.50"])
    join_date = random.choice([
        datetime(2026, 1, 10),
        datetime(2026, 2, 15),
        "not-a-date",
        None
    ])
    status = random.choice(statuses)
    ws.append([customer_id, name, email, amount, join_date, status])

# Add exact duplicates and a fully blank row
ws.append(list(ws[2]))
ws.append([None] * 6)

# Sheet 2: Orders (35 data rows)
ws2 = wb.create_sheet("Orders")
ws2.append(["Order ID", "Customer Name", "Quantity", "Price", "Order Date"])

for i in range(101, 136):
    ws2.append([
        i,
        random.choice(["Alice", "Bob", "Priya", None, "   "]),
        random.choice([1, 2, 3, 0, -1, None]),
        random.choice([50, 75.5, 100, -20, None]),
        random.choice([datetime(2026, 5, 1), datetime(2026, 5, 2), None])
    ])

ws2.append(list(ws2[2]))  # Exact duplicate
ws2.append([None] * 5)

# Sheet 3: Mixed data (25 data rows)
ws3 = wb.create_sheet("MixedData")
ws3.append(["Record ID", "Value", "Is Active", "Notes"])

for i in range(1, 26):
    ws3.append([
        f"{i:04d}",
        random.choice([100, "200", "N/A", None, 3.14]),
        random.choice([True, False, None]),
        random.choice(["Valid record", "  Needs trimming  ", None, "Check this value"])
    ])

# Format headers and freeze the first row
for sheet in wb.worksheets:
    sheet.freeze_panes = "A2"
    for cell in sheet[1]:
        cell.font = cell.font.copy(bold=True)

wb.save(output_file)
wb.close()

print(f"Test workbook created: {output_file.resolve()}")
print("Data rows: 40 Customers + 35 Orders + 25 MixedData = 100")
