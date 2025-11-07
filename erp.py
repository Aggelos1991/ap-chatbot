import pandas as pd
from openai import OpenAI
import math, time, re

client = OpenAI(api_key="YOUR_OPENAI_KEY")

# === Load your Excel export ===
df = pd.read_excel("entersoft_dictionary.xlsx")

# --- Clean up ---
df = df.dropna(subset=["Greek", "English"]).reset_index(drop=True)

BATCH_SIZE = 50   # process 50 rows at a time
results = []

def parse_ai_output(text):
    """
    Expecting lines like:
    OBJECTID | Greek | English | Status | Description | Reason
    """
    out = []
    for line in text.strip().splitlines():
        parts = [p.strip() for p in line.split("|")]
        if len(parts) >= 6:
            out.append({
                "OBJECTID": parts[0],
                "Greek": parts[1],
                "English": parts[2],
                "Status": parts[3],
                "Status Description": parts[4],
                "Reason": "|".join(parts[5:])   # join if '|' exists in reason
            })
    return out

for i in range(0, len(df), BATCH_SIZE):
    batch = df.iloc[i:i+BATCH_SIZE]

    # --- Build compact prompt ---
    prompt_rows = []
    for _, row in batch.iterrows():
        prompt_rows.append(f"{row['OBJECTID']} | {row['Greek']} | {row['English']}")
    joined = "\n".join(prompt_rows)

    prompt = f"""
You are an ERP translation auditor. For each line below (OBJECTID | Greek | English),
decide which status applies and explain briefly.

Statuses:
1 = Translated_Correct
2 = Translated_Not Accurate
3 = Field Not Translated
4 = Field Not Found on the Report View
0 = Pending Review (unclear)

Return one line per input, formatted exactly as:
OBJECTID | Greek | English | Status | Status Description | Reason

Now analyze:

{joined}
"""

    try:
        resp = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are a strict ERP translation checker. Output only in pipe-separated table lines."},
                {"role": "user", "content": prompt}
            ],
            temperature=0
        )

        text = resp.choices[0].message.content
        batch_results = parse_ai_output(text)
        results.extend(batch_results)

        print(f"✅ Processed rows {i+1}–{i+len(batch)}")
        time.sleep(0.3)   # gentle rate-limit safety

    except Exception as e:
        print(f"⚠️ Batch {i} failed: {e}")
        # mark all rows in that batch as pending review
        for _, row in batch.iterrows():
            results.append({
                "OBJECTID": row["OBJECTID"],
                "Greek": row["Greek"],
                "English": row["English"],
                "Status": 0,
                "Status Description": "Pending Review",
                "Reason": f"Error: {e}"
            })

# === Export results ===
out = pd.DataFrame(results)
out.to_excel("translation_audit_results.xlsx", index=False)
print("\n✅ translation_audit_results.xlsx created successfully.")
