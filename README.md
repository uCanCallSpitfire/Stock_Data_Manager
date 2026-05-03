# Stock Data Manager

Stock Data Manager is a lightweight Python CLI for downloading Yahoo Finance market data, exporting polished Excel reports, and optionally displaying a closing-price chart.

It is built for simple personal analysis workflows: choose a ticker, select a period or custom date range, generate a clean workbook, and keep previous generated reports refreshed.

## Features

- Download OHLCV market data with `yfinance`
- Export styled Excel reports with filters, frozen headers, borders, and number formatting
- Use predefined periods such as `1mo`, `6mo`, `1y`, `ytd`, and `max`
- Use custom date ranges with configurable intervals
- Refresh existing generated reports in the output folder
- Display a Tkinter/Matplotlib closing-price chart
- Run as a plain Python script or install as a command-line tool

## Requirements

- Python 3.10 or newer
- Internet connection for Yahoo Finance data
- Tkinter support if you want to open the chart window

## Installation

Clone the repository:

```bash
git clone <repository-url>
cd Stock_Data_Manager
```

Create a virtual environment:

```bash
python -m venv .venv
.venv\Scripts\activate
```

Install dependencies:

```bash
pip install -r requirements.txt
```

Optional editable installation:

```bash
pip install -e .
```

After editable installation, you can run the tool with:

```bash
stock-data-manager --stock TSLA --period 6mo --interval 1d
```

## Usage

Run with the default configuration:

```bash
python Stock_Data_Manager.py
```

This downloads `TSLA` data for the last `6mo` using a `1d` interval, saves an Excel report to `stocks/`, refreshes existing generated reports, and opens a chart window.

Download a different ticker:

```bash
python Stock_Data_Manager.py --stock AAPL
```

Change period and interval:

```bash
python Stock_Data_Manager.py --stock MSFT --period 1y --interval 1d
```

Use a custom date range:

```bash
python Stock_Data_Manager.py --stock NVDA --start-date 2024-01-01 --end-date 2024-06-01 --custom-interval 1d
```

Skip the chart window:

```bash
python Stock_Data_Manager.py --stock TSLA --no-chart
```

Skip automatic refresh of existing reports:

```bash
python Stock_Data_Manager.py --stock TSLA --no-auto-update
```

Choose another output folder:

```bash
python Stock_Data_Manager.py --stock TSLA --output-dir reports
```

## Command-Line Options

| Option | Description | Default |
| --- | --- | --- |
| `--stock` | Ticker symbol, for example `TSLA` or `AAPL` | `TSLA` |
| `--period` | Yahoo Finance period | `6mo` |
| `--interval` | Yahoo Finance interval for period-based downloads | `1d` |
| `--start-date` | Custom start date in `YYYY-MM-DD` format | none |
| `--end-date` | Custom end date in `YYYY-MM-DD` format | none |
| `--custom-interval` | Interval for custom date ranges | `1d` |
| `--output-dir` | Folder for generated Excel files | `stocks` |
| `--no-chart` | Skip the chart window | disabled |
| `--no-auto-update` | Skip refreshing existing generated reports | disabled |

Supported periods:

```text
1d, 5d, 1mo, 3mo, 6mo, 1y, 2y, 5y, 10y, ytd, max
```

Supported intervals:

```text
1m, 2m, 5m, 15m, 30m, 60m, 90m, 1d, 5d, 1wk, 1mo, 3mo
```

Note: Yahoo Finance limits very short intervals such as `1m` to recent date ranges.

## Output

Generated workbooks contain:

| Column | Description |
| --- | --- |
| Date | Trading date |
| Time | Trading time |
| Open | Opening price |
| High | Highest price |
| Low | Lowest price |
| Close | Closing price |
| Volume | Trading volume |

Example output file:

```text
stocks/TSLA-1d-6mo.xlsx
```

Generated Excel files are ignored by Git so the repository stays clean.

## Project Structure

```text
Stock_Data_Manager/
  .github/
    ISSUE_TEMPLATE/
    workflows/
  CHANGELOG.md
  CONTRIBUTING.md
  LICENSE
  README.md
  Stock_Data_Manager.py
  pyproject.toml
  requirements.txt
```

## Development

Run basic checks:

```bash
python -m py_compile Stock_Data_Manager.py
python Stock_Data_Manager.py --help
```

Run a real export test after installing dependencies:

```bash
python Stock_Data_Manager.py --stock TSLA --period 1mo --interval 1d --no-chart --no-auto-update
```

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE).

## Disclaimer

This project is for educational and personal analysis purposes only. It does not provide financial advice.
