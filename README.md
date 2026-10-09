# 📊 Excel MIS Report Automation System

Ek powerful aur intelligent **Management Information System (MIS) Report Automation Tool** jo kisi bhi Excel (`.xlsx`, `.xls`) ya CSV dataset ko automatically analyse karke:
1. **Interactive Web Dashboard** (Streamlit) provide karta hai.
2. **Professional Multi-Tab Formatted Excel MIS Report** (`.xlsx`) generate karta hai jisme executive KPI scorecards, multi-dimensional breakdowns, monthly run-rate trends, formulas, freeze panes, aur data quality audit shamil hain.
3. **CLI / Terminal Runner** provide karta hai background ya batch automation ke liye.

---

## 🚀 Key Features (Khasiyat)

- **Smart Auto-Detection**: Aapki file ke column names ya structure kuch bhi ho, ye automatically pehchanta hai:
  - **Metrics / Values**: (Gross Sales, Net Sales, Amount, Profit, Cost, Quantity, etc.)
  - **Dimensions / Categories**: (Region, Branch, Category, Product, Sales Rep, Status, etc.)
  - **Timeline / Dates**: (Order Date, Invoice Date, Month, Timestamp, etc.)
- **Executive KPI Cards**: Total Volume, Cumulative Totals, Average Ticket Size, Max Peaks.
- **Multi-Dimensional Breakdowns**:
  - Category-wise Sum, Average, Count, aur **% Contribution Share**.
  - **Pareto (80/20 Rule) Analysis**: Cumulative percentage tracking.
- **Time-Series & Trend Analysis**:
  - Month-over-Month (MoM) Growth Rate (%).
  - Peak activity period aur daily run-rate.
- **Interactive 2-Way Cross-Tabulation (Pivot Table)**: Kisi bhi 2 categories ka live matrix breakdown.
- **Data Quality & Exception Audit**:
  - Missing values / Null cells check.
  - Duplicate records detection.
  - Statistical Outlier detection using **IQR Method**.
- **1-Click Executive Excel Download**: Ek formatted, corporate-styled Excel workbook jisme 5 sheets hoti hain:
  1. `Executive_Summary`: KPI cards, automated business insights narrative, key dimension snapshot.
  2. `Category_Analysis`: Top categories ki formatted tables with Excel `=SUM()` formulas aur percentage formatting.
  3. `Trends_Monthly`: Month-by-month run-rate with MoM growth rate %.
  4. `Data_Quality_Audit`: Missing values aur outlier analysis audit.
  5. `Cleaned_Data`: Freeze panes aur AutoFilter ke saath structured raw data.

---

## 📁 Project Structure

```text
excel_mis_automation/
├── app.py                  # Interactive Streamlit Web Dashboard
├── mis_engine.py           # Core analytics, KPI & machine intelligence engine
├── excel_reporter.py       # OpenPyXL-based corporate multi-tab Excel generator
├── cli_runner.py           # Command line runner for quick batch automation
├── generate_sample_data.py # Realistic enterprise sales dataset generator
├── run_app.bat             # Windows 1-click launcher
├── Sample_Sales_MIS.xlsx   # Pre-generated sample dataset for instant testing
└── README.md               # Complete documentation
```

---

## 💻 Kaise Chalayein (How to Run)

### Method 1: Interactive Web Dashboard (Recommended)
Terminal mein ye command chalayein:
```bash
streamlit run app.py
```
*Ya simply `run_app.bat` file par double-click karein.*

Browser mein dashboard khul jayega:
- **Apni Excel ya CSV file drag & drop karein**, ya sidebar se **"⚡ Load Demo Sales MIS Data"** button dabakar sample data test karein.
- Filters aur analytical mappings adjust karein.
- **"📥 Download Formatted Excel MIS Report"** button dabakar complete styled Excel workbook download karein!

---

### Method 2: Command Line (CLI / Batch Automation)
Agar aapko bina browser khole directly kisi file ki MIS report generate karni hai:
```bash
python cli_runner.py "path/to/your_file.xlsx"
```
Ye automatic format karke output file save kar dega:
```bash
# Example Output:
# [SUCCESS] Executive MIS Report generated at: your_file_MIS_Report.xlsx
```

---

### Method 3: Naya Demo Dataset Generate Karna
Agar aapko naya testing dataset generate karna ho:
```bash
python generate_sample_data.py
```
Ye 500 transactions waala realistic enterprise dataset `Sample_Sales_MIS.xlsx` create karega.
