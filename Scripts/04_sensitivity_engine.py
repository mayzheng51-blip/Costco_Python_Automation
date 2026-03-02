import xlwings as xw
import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime

ROOT = Path(__file__).resolve().parents[1]
MODEL = ROOT / "Models" / "Costco – Elite Financial Model & Valuation.xlsx"
OUT = ROOT / "outputs"
OUT.mkdir(exist_ok=True)

WACC_CELL = ("DCF Valuation", "B12")
G_CELL = ("DCF Valuation", "B13")
EQUITY_CELL = ("DCF Valuation", "B43")

# Sensitivity ranges
wacc_range = np.arange(0.07, 0.091, 0.005)      # 7.0% to 9.0%
g_range = np.arange(0.02, 0.031, 0.005)         # 2.0% to 3.0%

def run_sensitivity():
    app = xw.App(visible=False)
    try:
        wb = app.books.open(str(MODEL))
        sheet = wb.sheets["DCF Valuation"]

        results = []

        for w in wacc_range:
            for g in g_range:
                sheet.range(WACC_CELL[1]).value = w
                sheet.range(G_CELL[1]).value = g

                app.calculate()

                equity = sheet.range(EQUITY_CELL[1]).value

                results.append({
                    "WACC": w,
                    "Terminal Growth": g,
                    "Equity Value": equity
                })

        wb.close()

        return pd.DataFrame(results)

    finally:
        app.quit()

if __name__ == "__main__":
    df = run_sensitivity()

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = OUT / f"costco_sensitivity_{ts}.csv"
    df.to_csv(path, index=False)

    print("Saved sensitivity grid:", path)
    