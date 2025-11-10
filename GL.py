# ===========================================
# FINAL AGGREGATOR — COLOR GROUP LOGIC
# ===========================================
# Aggregates Sheet2[K,L] by Διάσταση 2 groups
# Adds totals into Sheet1 (col J,K) using mapping via column F
# Keeps formatting, keeps Διάσταση, sets 0 for specified accounts
# ===========================================

import sys
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

if len(sys.argv) < 2:
    print("Usage: python aggregate_final.py <excel_path>")
    sys.exit()

file_path = sys.argv[1]
wb = load_workbook(file_path)
ws1, ws2 = wb.worksheets[:2]

# =====================================================
# CONFIG
# =====================================================
SHEET1_COL_TITLE = 6   # Column F for mapping
SHEET1_COL_J = 10      # Χρεωστικό Υπόλοιπο
SHEET1_COL_K = 11      # Πιστωτικό Υπόλοιπο
SHEET2_COL_D2 = 2      # Διάσταση 2
SHEET2_COL_K = 11
SHEET2_COL_L = 12

ZERO_ACCOUNTS = {
    "50.00.00.0000","50.00.00.0001","50.00.00.0002","50.00.00.0003",
    "50.01.00.0000","50.01.01.0000","50.05.00.0000"
}

# Mapping between Διάσταση 2 and Sheet1 Τίτλος (column F)
D2_TO_TITLE = {
    "--": "Προμηθευτές Capex πιστωτικά υπόλοιπα τέλους περιόδου",
    "01 - OpEx Payables": "Προμηθευτές Capex πιστωτικά υπόλοιπα τέλους περιόδου",
    "02 - CapEx Payables": "Προμηθευτές πιστωτικά υπόλοιπα τέλους περιόδου",
    "03 - Other Payables": "Προμηθευτές Capex πιστωτικά υπόλοιπα τέλους περιόδου",
    "04 - OpEx Advances": "Προμηθευτές χρεωστικά (προκαταβολές) υπόλοιπα τέλους περιόδου - Χρεώστες",
    "05 - CapEx Advances": "Προμηθευτές χρεωστικά (προκαταβολές) υπόλοιπα τέλους περιόδου - Προκαταβολές για αγορές Παγίων",
    "06 - Other Advances": "Προμηθευτές χρεωστικά (προκαταβολές) υπόλοιπα τέλους περιόδου - Χρεώστες",
    "100 - General B2B Invoices – Payments": "Προμηθευτές Capex πιστωτικά υπόλοιπα τέλους περιόδου",
    "110 - B2B Aging collections": "Προμηθευτές Capex πιστωτικά υπόλοιπα τέλους περιόδου",
    "2200 - Development Capex": "Προμηθευτές Capex πιστωτικά υπόλοιπα τέλους περιόδου",
    "300 - Financing Cashflows": "Προμηθευτές Capex πιστωτικά υπόλοιπα τέλους περιόδου"
}

# COLOR-GROUP SIMULATION (like you showed)
COLOR_GROUPS = {
    "GROUP_A": ["01 - OpEx Payables", "03 - Other Payables", "--"],
    "GROUP_B": ["02 - CapEx Payables"],
    "GROUP_C": ["04 - OpEx Advances", "05 - CapEx Advances", "06 - Other Advances"],
    "GROUP_D": ["100 - General B2B Invoices – Payments", "110 - B2B Aging collections", "2200 - Development Capex", "300 - Financing Cashflows"]
}

# =====================================================
# STEP 1: Aggregate Sheet2 by color group
# =====================================================
group_sums = {g: {"K": 0.0, "L": 0.0} for g in COLOR_GROUPS}

for row in range(2, ws2.max_row + 1):
    d2 = ws2.cell(row, SHEET2_COL_D2).value
    if not d2:
        continue
    d2 = str(d2).strip()
    try:
        val_k = float(ws2.cell(row, SHEET2_COL_K).value or 0)
        val_l = float(ws2.cell(row, SHEET2_COL_L).value or 0)
    except:
        continue

    # Find which group this Διάσταση belongs to
    for group, d2_list in COLOR_GROUPS.items():
        if d2 in d2_list:
            group_sums[group]["K"] += val_k
            group_sums[group]["L"] += val_l
            break

# =====================================================
# STEP 2: Push totals to Sheet1
# =====================================================
# Helper: find Διάσταση 2 key for a given title
def find_group_for_title(title):
    for d2, greek in D2_TO_TITLE.items():
        if greek == title:
            for group, members in COLOR_GROUPS.items():
                if d2 in members:
                    return group
    return None

for r in range(2, ws1.max_row + 1):
    acc = str(ws1.cell(r, 4).value or "").strip()  # column D = account
    title = str(ws1.cell(r, SHEET1_COL_TITLE).value or "").strip()

    # Zero accounts → force 0
    if acc in ZERO_ACCOUNTS:
        ws1.cell(r, SHEET1_COL_J, 0)
        ws1.cell(r, SHEET1_COL_K, 0)
        continue

    group = find_group_for_title(title)
    if not group:
        continue

    cur_j = float(ws1.cell(r, SHEET1_COL_J).value or 0)
    cur_k = float(ws1.cell(r, SHEET1_COL_K).value or 0)
    ws1.cell(r, SHEET1_COL_J, cur_j + group_sums[group]["K"])
    ws1.cell(r, SHEET1_COL_K, cur_k + group_sums[group]["L"])

# =====================================================
# STEP 3: Auto-fit width (visual cleanup)
# =====================================================
for ws in (ws1, ws2):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max_len + 2, 40)

wb.save(file_path)
print("✅ Done! Aggregation successful and formatting preserved.")
