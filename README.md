# Financial Model Automation – Python + Excel

## Overview

This project automates scenario analysis and sensitivity testing for an Excel-based financial valuation model.

Instead of manually updating assumptions and recalculating outputs, this system uses Python to:

- Update scenario inputs
- Trigger Excel recalculations
- Extract key outputs
- Export structured CSV results
- Run automated sensitivity analysis

## Technologies Used

- Python
- Pandas
- OpenPyXL
- Excel integration
- Financial modeling logic

## Project Structure

Scripts:
- `01_set_scenario_save.py` – updates financial assumptions
- `02_recalc_with_excel.py` – forces model recalculation
- `03_snapshot_to_csv.py` – exports model outputs
- `04_sensitivity_engine.py` – runs multi-scenario sensitivity tests

## Why This Matters

This project demonstrates:

- Automation of financial workflows
- Python + Excel integration
- Scenario modeling
- Data extraction for analytics
- Process optimization

## Future Improvements

- Tableau dashboard integration
- Monte Carlo simulation module
- Web-based input interface
