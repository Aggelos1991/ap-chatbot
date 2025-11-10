# file: fix_aggregate.py
# Usage: python fix_aggregate.py "path\to\your.xlsx"

import sys
from openpyxl import load_workbook

# ========= CONFIG (change only if needed) =========
ADD_MODE = True          # True = add to existing; False = overwrite
SHEET1_COL_TITLE = 6     # Sheet1 column F (Greek title)
SHEET1_COL_K = 11        # Sheet1 column K (Πιστωτικό Υπόλοιπο)
SHEET1_COL_L = 12        # Sheet1 column L (Χρεωστικό Υπόλοιπο) if you want J use 10
SHEET2_COL_D2 = 2        # Sheet2 column B (Διάσταση 2)
SHEET2_COL_K = 11        # Sheet2 column K
SHEET2_COL_L = 12        # Sheet2 column L

ZERO_ACCOUNTS = {
    "50.00.00.0000","50.00.00.0001","50.00.00.0002","50.00.00.0003",
    "50.01.00.0000","50.01.01.0000","50.05.00.0000"
}
# Διάσταση 2 -> Greek title (what appears in Sheet1 column F)
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
    "300 - Financing Cashflows": "Προμηθευτές Capex πιστωτικά υπόλοιπα τέλους περιόδου",
}
TITLE_TO_D2 = {v: k for k, v in D2_TO_TITLE.items()}
# ==================================================

if len(sys.argv) < 2:
    print("Give path to .xlsx, e.g. python fix_aggregate.py \"C:\\path\\file.xlsx\"")
    sys.exit(1)

path = sys.argv[1]
wb = load_workbook(path)
ws1 = wb.worksheets[0]   # target
ws2 = wb.worksheets[1]   # source

# ---- 1) Aggregate Sheet2 by Διάσταση 2 (B) for K and L ----
aggK, aggL = {}, {}
for r in range(2, ws2.max_row + 1):
    d2 = ws2.cell(r, SHEET2_COL_D2).value
    if not d2: 
        continue
    d2 = str(d2).strip()
    try: k_val = float(ws2.cell(r, SHEET2_COL_K).value or 0)
    except: k_val = 0.0
    try: l_val = float(ws2.cell(r, SHEET2_COL_L).value or 0)
    except: l_val = 0.0
    aggK[d2] = aggK.get(d2, 0.0) + k_val
    aggL[d2] = aggL.get(d2, 0.0) + l_val

# ---- 2) Push into Sheet1 (by column positions only) ----
# Account column: try to find one with dots; fallback to column 2 (B)
def account_val(row_idx):
    # Try common places (B/D/E). We do NOT insert/delete anything.
    for c in (2,4,5):
        v = ws1.cell(row_idx, c).value
        if v and isinstance(v, str) and v.count(".") >= 3:
            return v.strip()
    # fallback
    v = ws1.cell(row_idx, 2).value
    return str(v).strip() if v else ""

for r in range(2, ws1.max_row + 1):
    acc = account_val(r)
    if acc in ZERO_ACCOUNTS:
        ws1.cell(r, SHEET1_COL_K, 0)   # K
        ws1.cell(r, SHEET1_COL_L, 0)   # L
        continue

    title = ws1.cell(r, SHEET1_COL_TITLE).value
    if not title:
        continue
    title = str(title).strip()
    d2_key = TITLE_TO_D2.get(title, None)
    if not d2_key:
        continue

    addK = aggK.get(d2_key, 0.0)
    addL = aggL.get(d2_key, 0.0)

    # read current
    try: curK = float(ws1.cell(r, SHEET1_COL_K).value or 0)
    except: curK = 0.0
    try: curL = float(ws1.cell(r, SHEET1_COL_L).value or 0)
    except: curL = 0.0

    if ADD_MODE:
        ws1.cell(r, SHEET1_COL_K, curK + addK)
        ws1.cell(r, SHEET1_COL_L, curL + addL)
    else:
        ws1.cell(r, SHEET1_COL_K, addK)
        ws1.cell(r, SHEET1_COL_L, addL)

wb.save(path)
print("Done ✔  — Sheet2[B,K,L] aggregated and pushed to Sheet1[K,L]. Διάσταση untouched.")
