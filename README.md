# 🟦 StockDataManager

**StockDataManager** is a powerful Python tool for fetching, processing, and visualizing stock market data.  
It automatically downloads stock data from Yahoo Finance, saves styled Excel reports, and shows interactive closing price charts. Perfect for traders, analysts, and coders who want fast, organized stock insights.  

---

## ⚡ Features

- Fetch stock data by symbol (e.g., `TSLA`, `AAPL`)  
- Support for custom date ranges or predefined periods (`1d`, `1mo`, `6mo`, `1y`, `max`, etc.)  
- Configurable interval (`1m`, `5m`, `1d`, `1wk`, `1mo`, …)  
- Automatically updates Excel files with latest stock data  
- Generates clean, styled Excel sheets (Date, Time, Open, High, Low, Close, Volume)  
- Interactive chart visualization using `matplotlib` and `Tkinter`  
- Fully customizable and open-source  

---

## 🛠 Installation

1. Clone the repo:  
```bash
git clone https://github.com/username/StockDataManager.git
cd StockDataManager
Install dependencies:

pip install -r requirements.txt


requirements.txt should include: yfinance, pandas, matplotlib, openpyxl, tkinter

🚀 How to Use

Open stock_data_manager.py

Configure your settings at the top:

stock = 'TSLA'           # Stock symbol
period = "6mo"           # e.g., "1d", "1mo", "6mo", "1y", "max"
interval = "1d"           # e.g., "1m", "5m", "1d", "1wk"
use_custom_date = False   # Set True to use start/end dates
start_date = "2024-01-01"
end_date = "2024-06-01"
show_chart = True         # Show interactive chart
auto_update = True        # Automatically update Excel files


Run the script:

python stock_data_manager.py


The Excel file will be saved in the stocks/ folder:

TSLA-1d-6mo.xlsx


If show_chart = True, an interactive chart window will appear with closing prices.

📊 Example Output

Excel Sheet Columns:

Date – Trading date

Time – Always 00:00:00

Open – Opening price

High – Highest price

Low – Lowest price

Close – Closing price

Volume – Trading volume

Chart Example:
Closing price over the last 6 months (interactive, zoomable via Tkinter/Matplotlib)

⚙️ Notes

If auto_update = True, existing Excel files in the folder will be updated automatically

Supports all major Yahoo Finance stock symbols

Designed for Windows, Linux, and MacOS with Python 3.9+

💡 Future Improvements

Add versioning for Excel files instead of overwriting

Add multiple stock comparison charts

Export charts as PNG or PDF automatically

Include candlestick charts for better analysis

📝 License

MIT License – free to use, modify,
