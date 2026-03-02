import xlwings as xw
from pathlib import Path
from datetime import datetime

ROOT = Path(__file__).resolve().parents[1]
MODEL = ROOT / "Models" / "Costco – Elite Financial Model & Valuation.xlsx"
OUT = ROOT / "outputs"

SCENARIO_SHEET = "Executive Dashboard"
SCENARIO_CELL = "E4"

def recalc_and_save(scenario):
    app = xw.App(visible=False)
    try:
        wb = app.books.open(str(MODEL))
        ws = wb.sheets[SCENARIO_SHEET]

        ws.range(SCENARIO_CELL).value = scenario

        app.calculate()  # force full recalculation

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        save_path = OUT / f"recalc_{scenario}_{ts}.xlsx"
        wb.save(str(save_path))
        wb.close()

        print(f"Recalculated and saved: {scenario}")

    finally:
        app.quit()

if __name__ == "__main__":
    OUT.mkdir(exist_ok=True)
    for s in ["Bear", "Base", "Bull"]:
        recalc_and_save(s)