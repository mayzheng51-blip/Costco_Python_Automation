import xlwings as xw
import pandas as pd
from pathlib import Path
from datetime import datetime

ROOT = Path(__file__).resolve().parents[1]
MODEL = ROOT / "Models" / "Costco – Elite Financial Model & Valuation.xlsx"
OUT = ROOT / "outputs"
OUT.mkdir(exist_ok=True)

SCENARIO_SHEET = "Executive Dashboard"
SCENARIO_CELL = "E4"

CELL_MAP = {
    "WACC": ("DCF Valuation", "B12"),
    "Terminal Growth": ("DCF Valuation", "B13"),
    "PV of FCFF": ("DCF Valuation", "B31"),
    "PV of Terminal Value": ("DCF Valuation", "B34"),
    "Enterprise Value": ("DCF Valuation", "B36"),
    "Equity Value": ("DCF Valuation", "B43"),
    "2026 Revenue": ("DCF Valuation", "B17"),
    "2026 EBIT": ("DCF Valuation", "B22"),
    "2026 FCFF": ("DCF Valuation", "B28"),
}

def snapshot_scenario(app, scenario: str):
    wb = app.books.open(str(MODEL))
    try:
        wb.sheets[SCENARIO_SHEET].range(SCENARIO_CELL).value = scenario
        app.calculate()

        row = {"Scenario": scenario}
        for name, (sheet, addr) in CELL_MAP.items():
            row[name] = wb.sheets[sheet].range(addr).value

        # Derived metric
        pv_term = row["PV of Terminal Value"]
        ev = row["Enterprise Value"]
        row["Terminal % of EV"] = (pv_term / ev) if ev else None

        return row
    finally:
        wb.close()

def main():
    app = xw.App(visible=False)
    try:
        rows = [snapshot_scenario(app, s) for s in ["Bear", "Base", "Bull"]]
    finally:
        app.quit()

    df = pd.DataFrame(rows)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_path = OUT / f"costco_scenarios_{ts}.csv"
    df.to_csv(csv_path, index=False)

    print("Saved CSV:", csv_path)

if __name__ == "__main__":
    main()
    