import json
import pandas as pd
import numpy as np
import argparse
from pathlib import Path

REQUIRED = [
    "SampleID",
    "iso87Sr86Sr",
    "iso143Nd144Nd",
    "iso176Hf177Hf",
    "size", "color", "marker", "alpha", "edgecolor", "edgewidth",
]

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("excel_path", help="Path to RCD_master_excel_1.xlsx")
    ap.add_argument("--sheet", default="RCD")
    ap.add_argument("--out", default="rcd_points.json")
    args = ap.parse_args()

    df = pd.read_excel(args.excel_path, sheet_name=args.sheet)

    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(f"Missing columns in sheet '{args.sheet}': {missing}")

    # Keep only required columns
    out = df[REQUIRED].copy()

    # Coerce numeric columns (ratios + style numerics)
    num_cols = ["iso87Sr86Sr", "iso143Nd144Nd", "iso176Hf177Hf", "size", "alpha", "edgewidth"]
    for c in num_cols:
        out[c] = pd.to_numeric(out[c], errors="coerce")

    # Replace NaN with None so JSON is clean
    out = out.replace({np.nan: None})

    records = out.to_dict(orient="records")
    Path(args.out).write_text(json.dumps(records, indent=2), encoding="utf-8")

    print(f"Wrote {args.out} with {len(records)} rows.")

if __name__ == "__main__":
    main()