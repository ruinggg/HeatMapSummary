from openpyxl import load_workbook, Workbook
import os

# === Configuration ===
tower_list = ["P1", "P2", "P3", "P4"]
base_filename = "20250411_M46_JV3.1_{}.xlsm"
source_sheet = "HeatMap"
target_file = "Summary.xlsx"

# Each tower occupies 37 piers + 1 blank column = 39 columns
tower_column_span = 39
target_start_row = 2  # Paste begins at row 2 (B2)
target_start_col = 2  # Column B = 2
data_row_count = 113  # Number of data rows

# === Field settings ===
fields = [
    {
        "name": "V_DCR",
        "range": ("DR3", "FC115"),
        "sheet": "SummaryVDCR"
    },
    {
        "name": "Vmax_DCR",
        "range": ("FE3", "GP115"),
        "sheet": "SummaryVmax"
    },
    {
        "name": "PC_DCR",
        "range": ("GR3", "IC115"),
        "sheet": "SummaryPCDCR"
    },
    {
        "name": "PT_DCR",
        "range": ("IE3", "JP115"),
        "sheet": "SummaryPTDCR"
    }
]

# === Load or create target workbook ===
if os.path.exists(target_file):
    print("📄 Loading existing Summary.xlsx...")
    tgt_wb = load_workbook(target_file)
else:
    print("🆕 Creating new Summary.xlsx...")
    tgt_wb = Workbook()
    # Clear the default sheet if it exists
    default_sheet = tgt_wb.active
    if default_sheet and default_sheet.title == "Sheet":
        tgt_wb.remove(default_sheet)

# === Process each field (per sheet) ===
for field in fields:
    sheet_name = field["sheet"]
    source_range = field["range"]

    # Create or get the worksheet
    if sheet_name in tgt_wb.sheetnames:
        tgt_ws = tgt_wb[sheet_name]
    else:
        print(f"🆕 Creating sheet: {sheet_name}")
        tgt_ws = tgt_wb.create_sheet(title=sheet_name)

    # Clear header (row 1-2) and body (row 2-114) before pasting
    print(f"\n🧹 Clearing old data in {sheet_name}...")
    for col in range(target_start_col, target_start_col + len(tower_list) * tower_column_span + 1):
        tgt_ws.cell(row=1, column=col).value = None
        tgt_ws.cell(row=2, column=col).value = None
        for row in range(target_start_row, target_start_row + data_row_count):
            tgt_ws.cell(row=row, column=col).value = None

    # === Process each tower ===
    for idx, tower in enumerate(tower_list):
        source_file = base_filename.format(tower)
        start_col = target_start_col + idx * tower_column_span

        print(f"\n🚧 Processing {field['name']} for Tower {tower} → starting column {start_col}...")

        if not os.path.exists(source_file):
            print(f"❌ File not found: {source_file} → Skipping.")
            continue

        try:
            src_wb = load_workbook(source_file, data_only=True, read_only=True)
            src_ws = src_wb[source_sheet]
        except Exception as e:
            print(f"❌ Failed to open {source_file}: {e}")
            continue

        print(f"📊 Reading range {source_range[0]}:{source_range[1]}...")
        dcr_values = src_ws[source_range[0]:source_range[1]]

        # Add header "Tower Px"
        tgt_ws.cell(row=1, column=start_col).value = f"Tower {tower}"

        print(f"✍️  Pasting data into {sheet_name}...")
        for i, row in enumerate(dcr_values):
            for j, cell in enumerate(row):
                tgt_ws.cell(row=target_start_row + i, column=start_col + j).value = cell.value

        print(f"✅ Finished pasting {tower} for {field['name']}.")

# === Save the workbook ===
print("\n💾 Saving all results to Summary.xlsx...")
tgt_wb.save(target_file)
print("🎉 All data pasted successfully!")
