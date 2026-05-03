from __future__ import annotations

import argparse
import logging
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    import pandas as pd


LOGGER = logging.getLogger("stock-data-manager")
OUTPUT_DIR = Path("stocks")

PERIOD_OPTIONS = (
    "1d",
    "5d",
    "1mo",
    "3mo",
    "6mo",
    "1y",
    "2y",
    "5y",
    "10y",
    "ytd",
    "max",
)

INTERVAL_OPTIONS = (
    "1m",
    "2m",
    "5m",
    "15m",
    "30m",
    "60m",
    "90m",
    "1d",
    "5d",
    "1wk",
    "1mo",
    "3mo",
)


@dataclass(frozen=True)
class StockRequest:
    symbol: str
    period: str = "6mo"
    interval: str = "1d"
    start_date: str | None = None
    end_date: str | None = None
    custom_interval: str = "1d"

    @property
    def uses_custom_dates(self) -> bool:
        return bool(self.start_date and self.end_date)

    @property
    def effective_interval(self) -> str:
        return self.custom_interval if self.uses_custom_dates else self.interval


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Download Yahoo Finance market data and export a polished Excel report."
    )
    parser.add_argument("--stock", default="TSLA", help="Ticker symbol, for example TSLA or AAPL.")
    parser.add_argument("--period", default="6mo", choices=PERIOD_OPTIONS)
    parser.add_argument("--interval", default="1d", choices=INTERVAL_OPTIONS)
    parser.add_argument("--start-date", help="Custom start date in YYYY-MM-DD format.")
    parser.add_argument("--end-date", help="Custom end date in YYYY-MM-DD format.")
    parser.add_argument("--custom-interval", default="1d", choices=INTERVAL_OPTIONS)
    parser.add_argument("--output-dir", default=str(OUTPUT_DIR), help="Folder for generated Excel files.")
    parser.add_argument("--no-chart", action="store_true", help="Skip the chart window.")
    parser.add_argument(
        "--no-auto-update",
        action="store_true",
        help="Skip refreshing existing Excel files in the output folder.",
    )
    return parser.parse_args()


def configure_logging() -> None:
    logging.basicConfig(format="%(levelname)s: %(message)s", level=logging.INFO)


def normalize_history_index(df: pd.DataFrame) -> pd.DataFrame:
    import pandas as pd

    if isinstance(df.index, pd.DatetimeIndex) and df.index.tz is not None:
        df = df.copy()
        df.index = df.index.tz_localize(None)
    return df


def fetch_history(request: StockRequest) -> pd.DataFrame:
    import yfinance as yf

    ticker = yf.Ticker(request.symbol)

    if request.uses_custom_dates:
        LOGGER.info(
            "Fetching %s from %s to %s at %s interval",
            request.symbol,
            request.start_date,
            request.end_date,
            request.custom_interval,
        )
        df = ticker.history(
            start=request.start_date,
            end=request.end_date,
            interval=request.custom_interval,
        )
    else:
        LOGGER.info(
            "Fetching %s for %s at %s interval",
            request.symbol,
            request.period,
            request.interval,
        )
        df = ticker.history(period=request.period, interval=request.interval)

    df = normalize_history_index(df)
    if df.empty:
        raise ValueError(f"No market data returned for {request.symbol}. Check the ticker and date range.")

    return df


def build_ohlc_table(df: pd.DataFrame) -> pd.DataFrame:
    required_columns = ["Open", "High", "Low", "Close", "Volume"]
    missing_columns = [column for column in required_columns if column not in df.columns]

    if missing_columns:
        raise ValueError(f"Missing expected market data columns: {', '.join(missing_columns)}")

    ohlc = df[required_columns].copy()
    ohlc[["Open", "High", "Low", "Close"]] = ohlc[["Open", "High", "Low", "Close"]].round(2)
    ohlc["Date"] = df.index.strftime("%Y-%m-%d")
    ohlc["Time"] = df.index.strftime("%H:%M:%S")
    return ohlc[["Date", "Time", "Open", "High", "Low", "Close", "Volume"]]


def apply_excel_style(ws) -> None:
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    header_fill = PatternFill("solid", fgColor="1F2937")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    body_font = Font(name="Calibri", size=11, color="111827")
    thin_gray = Side(style="thin", color="D1D5DB")
    border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for row_index, row in enumerate(ws.iter_rows(), start=1):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(
                horizontal="left" if row_index == 1 or cell.column <= 2 else "right",
                vertical="center",
            )
            if row_index == 1:
                cell.fill = header_fill
                cell.font = header_font
            else:
                cell.font = body_font

            if row_index > 1 and cell.column in (3, 4, 5, 6):
                cell.number_format = "#,##0.00"
            elif row_index > 1 and cell.column == 7:
                cell.number_format = "#,##0"

    column_widths = {
        "A": 14,
        "B": 12,
        "C": 12,
        "D": 12,
        "E": 12,
        "F": 12,
        "G": 16,
    }
    for column, width in column_widths.items():
        ws.column_dimensions[column].width = width


def export_to_excel(ohlc: pd.DataFrame, destination: Path) -> Path:
    from openpyxl import Workbook
    from openpyxl.utils.dataframe import dataframe_to_rows

    destination.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Market Data"

    for row in dataframe_to_rows(ohlc, index=False, header=True):
        ws.append(row)

    apply_excel_style(ws)
    wb.save(destination)
    return destination


def build_output_filename(request: StockRequest) -> str:
    symbol = sanitize_filename_part(request.symbol.upper())
    if request.uses_custom_dates:
        return f"{symbol}-{request.start_date}-{request.end_date}-{request.custom_interval}.xlsx"
    return f"{symbol}-{request.interval}-{request.period}.xlsx"


def sanitize_filename_part(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9_.=-]+", "_", value)


def refresh_existing_reports(output_dir: Path) -> None:
    if not output_dir.exists():
        LOGGER.info("No existing report folder found. Skipping auto-update.")
        return

    report_pattern = re.compile(
        r"^(?P<symbol>[A-Za-z0-9_.=]+)-(?P<interval>[A-Za-z0-9]+)-(?P<period>[A-Za-z0-9]+)\.xlsx$"
    )

    updated_count = 0
    for report_path in output_dir.glob("*.xlsx"):
        match = report_pattern.match(report_path.name)
        if not match:
            LOGGER.debug("Skipping unrecognized report name: %s", report_path.name)
            continue

        request = StockRequest(
            symbol=match.group("symbol"),
            period=match.group("period"),
            interval=match.group("interval"),
        )

        try:
            df = fetch_history(request)
            export_to_excel(build_ohlc_table(df), report_path)
            updated_count += 1
        except Exception as exc:
            LOGGER.warning("Could not update %s: %s", report_path.name, exc)

    LOGGER.info("Updated %s existing report(s).", updated_count)


def show_price_chart(df: pd.DataFrame, request: StockRequest) -> None:
    import matplotlib.pyplot as plt
    import tkinter as tk
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

    root = tk.Tk()
    root.title(f"{request.symbol.upper()} Closing Price")
    root.geometry("1100x650")
    root.minsize(850, 500)

    fig, ax = plt.subplots(figsize=(10, 5.5))
    ax.plot(df.index, df["Close"], label="Close", color="#2563EB", linewidth=2)
    ax.set_title(f"{request.symbol.upper()} Closing Price", fontsize=14, fontweight="bold")
    ax.set_xlabel("Date")
    ax.set_ylabel("Price")
    ax.grid(True, linestyle="--", linewidth=0.5, alpha=0.45)
    ax.legend()
    fig.autofmt_xdate()
    fig.tight_layout()

    canvas = FigureCanvasTkAgg(fig, master=root)
    canvas.draw()
    canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    root.mainloop()


def validate_request(args: argparse.Namespace) -> StockRequest:
    symbol = args.stock.strip().upper()
    if not symbol:
        raise ValueError("Stock symbol cannot be empty.")

    if bool(args.start_date) != bool(args.end_date):
        raise ValueError("Use --start-date and --end-date together.")

    if args.start_date and args.end_date:
        start_date = parse_iso_date(args.start_date, "--start-date")
        end_date = parse_iso_date(args.end_date, "--end-date")
        if start_date >= end_date:
            raise ValueError("--start-date must be earlier than --end-date.")

    return StockRequest(
        symbol=symbol,
        period=args.period,
        interval=args.interval,
        start_date=args.start_date,
        end_date=args.end_date,
        custom_interval=args.custom_interval,
    )


def parse_iso_date(value: str, option_name: str) -> date:
    try:
        return date.fromisoformat(value)
    except ValueError as exc:
        raise ValueError(f"{option_name} must use YYYY-MM-DD format.") from exc


def main() -> int:
    configure_logging()
    args = parse_args()
    output_dir = Path(args.output_dir)

    try:
        request = validate_request(args)

        if not args.no_auto_update:
            refresh_existing_reports(output_dir)

        df = fetch_history(request)
        ohlc = build_ohlc_table(df)
        output_path = export_to_excel(ohlc, output_dir / build_output_filename(request))
        LOGGER.info("Excel report saved: %s", output_path)

        if not args.no_chart:
            show_price_chart(df, request)

        return 0
    except Exception as exc:
        LOGGER.error("%s", exc)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
