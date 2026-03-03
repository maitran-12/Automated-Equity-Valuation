# Automated Equity Valuation & Stress Test Engine

## 🚀 Overview
This project is an automated financial decision-making tool. It bridges the gap between raw market data and actionable valuation insights by combining Python's data processing power with Excel's flexibility through VBA.

## 🛠️ Tech Stack
- **Data engineering:** Python (`yfinance`, `xlwings`, `pandas`).
- **Financial Modeling:** MS Excel (Advanced DCF, 3-Statement Linking).
- **Automation:** VBA (Visual Basic for Applications) for Scenario management.

## 💡 Key Features
- **Real-time Data updated:** Automatically fetches 3-year historical financials, Beta, and Risk-free rates directly into the model.
- **Dynamic Scenario calculation:** Built custom VBA macros to toggle between **Bull, Base, and Bear** cases. The UI automatically updates font colors and recalculates the intrinsic value instantly.
- **WACC stress testing:** Includes a WACC adjustment toggle to simulate market volatility in each scenario.

## 📁 Project Structure
- `Fetch_data.py`: Connects to API, cleans data, and dumps into Excel.
- `Valuation_model.xlsm`: The core engine containing the DCF logic and VBA UI.

## ⚙️ System Workflow
The integration follows a structured data pipeline to ensure accuracy and automation:

1. **Data Acquisition (Python):** Uses `yfinance` to fetch TTM (Trailing Twelve Months) financials and historical data.
   - Retrieves real-time Market Cap, Beta, and Risk-Free Rate (10Y Treasury).
2. **Data Processing & Bridge:** Cleans raw API data using `Pandas`.
   - Uses `xlwings` to push structured data into the `Data_dump` sheet without manual copy-pasting.
3. **Valuation Engine (Excel):** Formulas link `Data_dump` to the DCF model.
   - Calculates WACC, Free Cash Flow projections, and Terminal Value.
4. **Interactive UI (VBA):** User triggers "Scenario Buttons" on the Dashboard.
   - VBA updates assumptions and re-calculates the Intrinsic Value instantly.

## 📖 How to run
1. Ensure the ticker is entered in the `Ticker_input` range on the Dashboard.
2. Run `Fetch_data.py` to refresh all financial data.

3. Open the Excel file and use the **Scenario Buttons** to analyze different valuation outcomes.

## DEMO
<div align="center">
<video src="https://github.com/maitran-12/Automated-Equity-Valuation/raw/c1fc05bbdba97109d4a81c201d1691f22236acf9/Video%20Project.mp4" width="100%" controls></video>
</div>

## !NOTE
* The Python script (Fetch_data.py) fetches real-time market data directly from Yahoo Finance.
* It automatically updates the Risk-free rate via the 10-year Treasury Note (^TNX) and the latest Beta/Market Cap.
* Why the numbers might vary: If you see a slight difference between the recorded demo and the latest execution (e.g., $321.3 vs $320.5), it is because the model reflects the most current market conditions at the exact moment the script is run.



