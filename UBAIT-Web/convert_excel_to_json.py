import json
from pathlib import Path

import pandas as pd

INPUT_FILE = Path("Supplemental Table 1.xlsx")
OUTPUT_FILE = Path("data.json")

def to_float(value):
    if pd.isna(value):
        return None
    try:
        return float(value)
    except Exception:
        return None

def to_bool(value):
    if pd.isna(value):
        return False
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {"true", "t", "1", "yes", "y"}

def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()

def main():
    xls = pd.ExcelFile(INPUT_FILE)
    experiments = []

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(INPUT_FILE, sheet_name=sheet_name)
        df.columns = [str(c).strip() for c in df.columns]

        if "Accession" not in df.columns or "Description" not in df.columns:
            continue

        proteins = []
        for _, row in df.iterrows():
            accession = clean_text(row.get("Accession"))
            if not accession:
                continue

            proteins.append({
                "accession": accession,
                "description": clean_text(row.get("Description")),
                "log2_fc": to_float(row.get("log2FC")),
                "adj_p_value": to_float(row.get("adj.P.Value")),
                "neg_log10_p": to_float(row.get("negLog10P")),
                "neg_log10_adj_p": to_float(row.get("negLog10adj.P")),
                "status": clean_text(row.get("status")),
                "was_unique": to_bool(row.get("was_unique")),
                "imputed_for_volcano": to_bool(row.get("was_unique"))
            })

        experiments.append({
            "driver_name": sheet_name,
            "proteins": proteins
        })

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump({"experiments": experiments}, f, indent=2)

    print(f"Wrote {OUTPUT_FILE} with {len(experiments)} sheets.")

if __name__ == "__main__":
    main()
