#!/usr/bin/env python3
import argparse
import csv
import os
import sys
import time
import webbrowser
from datetime import datetime, timedelta, timezone

import numpy as np
import requests
from tqdm import tqdm
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ----------------------------
# Self-contained script setup
# ----------------------------
# Ensure script runs from its own directory (useful for StreamDeck, etc.)
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

# ----------------------------
# Config
# ----------------------------
BASE_URL = "https://api.exchange.coinbase.com"
HEADERS = {"User-Agent": "Mozilla/5.0"}
GRANULARITY = 86400  # daily candles
RATE_LIMIT_SLEEP = 0.35  # ~3 requests/sec safety

# ----------------------------
# Date window: will be calculated in main() based on --days parameter
# ----------------------------


# ----------------------------
# Helpers
# ----------------------------
def iso_format(dt: datetime) -> str:
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.isoformat()


def calculate_percentage_change(low: float, high: float) -> float:
    if low <= 0:
        return 0.0
    return (high - low) / low * 100.0


def safe_file_operation(operation_func, file_path: str, operation_name: str, max_retries: int = 10):
    """
    Safely perform file operations with user prompts when file is in use.
    
    Args:
        operation_func: Function to execute (should return True on success)
        file_path: Path to the file
        operation_name: Description of the operation for user prompts
        max_retries: Maximum number of retry attempts
    
    Returns:
        True if operation succeeded, False if user cancelled
    """
    for attempt in range(max_retries):
        try:
            return operation_func()
        except (PermissionError, OSError) as e:
            if attempt < max_retries - 1:
                print(f"\n⚠️  Cannot {operation_name} '{file_path}' - file may be open in another program.")
                print(f"   Error: {e}")
                print(f"\nPlease close the file and press Enter to retry, or type 'q' and Enter to quit:")
                
                user_input = input().strip().lower()
                if user_input == 'q':
                    print("Operation cancelled by user.")
                    return False
                print("Retrying...")
            else:
                print(f"\n❌ Failed to {operation_name} '{file_path}' after {max_retries} attempts.")
                print("Please close the file manually and run the script again.")
                return False
        except KeyboardInterrupt:
            print(f"\n\nOperation cancelled by user (Ctrl+C) during {operation_name}.")
            return False
        except Exception as e:
            print(f"❌ Unexpected error during {operation_name}: {e}")
            return False
    
    return False


def safe_remove_file(file_path: str) -> bool:
    """Safely remove a file with user prompts if it's in use."""
    if not os.path.exists(file_path):
        return True
    
    def remove_operation():
        os.remove(file_path)
        return True
    
    return safe_file_operation(remove_operation, file_path, "remove file")


def safe_write_file(file_path: str, write_func, operation_name: str) -> bool:
    """Safely write to a file with user prompts if it's in use."""
    def write_operation():
        write_func(file_path)
        return True
    
    return safe_file_operation(write_operation, file_path, operation_name)


def open_file(file_path: str) -> bool:
    """Open file with the default application."""
    try:
        # Convert to absolute path for better compatibility
        abs_path = os.path.abspath(file_path)
        
        # Use webbrowser to open files (works cross-platform)
        webbrowser.open(f"file://{abs_path}")
        return True
    except Exception as e:
        print(f"Warning: Could not auto-open file: {e}")
        return False


def create_excel_file(output_file: str) -> Workbook:
    """Create a new Excel workbook with headers and formatting."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Volatility Analysis"
    
    # Add headers
    headers = ["Pair", "Volatility", "Volume", "MinFunds"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    return wb


def save_to_excel(pair: str, median_pct_change: float, volume: float, min_order_size, wb: Workbook) -> None:
    """Add data row to Excel workbook."""
    ws = wb.active
    row = ws.max_row + 1
    
    # Add data as numeric values (Excel will handle formatting)
    ws.cell(row=row, column=1, value=pair)
    ws.cell(row=row, column=2, value=median_pct_change)
    ws.cell(row=row, column=3, value=int(volume) if volume > 0 else 0)
    
    # Ensure MinFunds is stored as a number
    if min_order_size == "N/A":
        ws.cell(row=row, column=4, value=0)
    else:
        try:
            # Convert to float to ensure it's a number
            ws.cell(row=row, column=4, value=float(min_order_size))
        except (ValueError, TypeError):
            ws.cell(row=row, column=4, value=0)


def format_excel_file(wb: Workbook) -> None:
    """Format Excel file with auto-sized columns and styling."""
    ws = wb.active
    
    # Set specific column widths for better appearance
    # Pair column (A) - auto-size based on content
    max_pair_length = 0
    for row in range(1, ws.max_row + 1):
        cell_value = str(ws.cell(row=row, column=1).value or "")
        if len(cell_value) > max_pair_length:
            max_pair_length = len(cell_value)
    ws.column_dimensions['A'].width = min(max_pair_length + 2, 20)
    
    # Volatility column (B) - fixed width for 2 decimal places
    ws.column_dimensions['B'].width = 12
    
    # Volume column (C) - wider for large numbers with commas
    max_volume_length = 0
    for row in range(1, ws.max_row + 1):
        cell_value = str(ws.cell(row=row, column=3).value or "")
        if len(cell_value) > max_volume_length:
            max_volume_length = len(cell_value)
    ws.column_dimensions['C'].width = min(max_volume_length + 3, 25)
    
    # MinFunds column (D) - fixed width for currency
    ws.column_dimensions['D'].width = 12
    
    # Format specific columns
    # Volatility column (B) - 2 decimal places
    for row in range(2, ws.max_row + 1):
        cell = ws.cell(row=row, column=2)
        cell.number_format = '0.00'
    
    # Volume column (C) - Number format with commas (no decimals)
    for row in range(2, ws.max_row + 1):
        cell = ws.cell(row=row, column=3)
        cell.number_format = '#,##0'
    
    # MinFunds column (D) - Simple number format to avoid green triangles
    for row in range(2, ws.max_row + 1):
        cell = ws.cell(row=row, column=4)
        # Use simple number format without currency symbol
        cell.number_format = '0.00'


def sort_excel_by_volatility(wb: Workbook) -> None:
    """Sort Excel data by volatility (highest first)."""
    ws = wb.active
    
    # Get all data rows (excluding header)
    data_rows = []
    for row in range(2, ws.max_row + 1):
        row_data = []
        for col in range(1, 5):  # 4 columns
            row_data.append(ws.cell(row=row, column=col).value)
        data_rows.append(row_data)
    
    # Sort by volatility (column 1) in descending order
    data_rows.sort(key=lambda x: float(x[1]) if x[1] is not None else 0, reverse=True)
    
    # Clear data rows and rewrite sorted data
    for row in range(2, ws.max_row + 1):
        for col in range(1, 5):
            ws.cell(row=row, column=col, value="")
    
    # Write sorted data
    for i, row_data in enumerate(data_rows, start=2):
        for j, value in enumerate(row_data, start=1):
            ws.cell(row=i, column=j, value=value)


def save_to_csv(pair: str, median_pct_change: float, volume: float, min_order_size, output_file: str) -> bool:
    """Save data to CSV file with safe file handling. Returns True if successful, False if cancelled."""
    header = ["Pair", "Volatility", "Volume", "MinFunds"]
    file_exists = os.path.exists(output_file)
    
    def write_operation():
        with open(output_file, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            if not file_exists:
                writer.writerow(header)
            # Format: Volatility to 2 decimal places, Volume as integer with commas
            formatted_volume = f"{int(volume):,}" if volume > 0 else "0"
            writer.writerow([pair, f"{median_pct_change:.2f}", formatted_volume, min_order_size])
        return True
    
    return safe_file_operation(write_operation, output_file, "write to CSV file")


def sort_csv_by_median(output_file: str) -> bool:
    """Sort CSV file by volatility with safe file handling. Returns True if successful, False if cancelled."""
    if not os.path.exists(output_file):
        return True
    
    def sort_operation():
        with open(output_file, "r", encoding="utf-8") as f:
            rows = list(csv.reader(f))
        if not rows:
            return True
        header, data = rows[0], rows[1:]
        # Sort by volatility (column 1), parsing the formatted string back to float
        data.sort(key=lambda r: float(r[1]), reverse=True)
        with open(output_file, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(header)
            writer.writerows(data)
        return True
    
    return safe_file_operation(sort_operation, output_file, "sort CSV file")


# ----------------------------
# Coinbase API
# ----------------------------
def get_active_pairs(quote_currency: str = "USD"):
    url = f"{BASE_URL}/products"
    resp = requests.get(url, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    products = resp.json()

    active = []
    # Convert to uppercase for consistent matching
    quote_currency_upper = quote_currency.upper()
    quote_suffix = f"-{quote_currency_upper}"
    
    for p in products:
        if p.get("trading_disabled"):
            continue
        if p.get("cancel_only"):
            continue
        if p.get("status") not in (None, "online", "active", "online_trading"):
            continue

        product_id = p.get("id") or p.get("product_id")
        if not product_id:
            continue

        if not product_id.endswith(quote_suffix):
            continue

        active.append(product_id)

    return sorted(set(active))


def get_pair_info(pair: str):
    url = f"{BASE_URL}/products/{pair}"
    resp = requests.get(url, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    product_info = resp.json()
    return product_info.get("min_market_funds", None)


def get_daily_volume(pair: str, start: datetime, end: datetime):
    """Get the median daily volume over the specified period"""
    url = f"{BASE_URL}/products/{pair}/candles"
    params = {
        "start": iso_format(start),
        "end": iso_format(end),
        "granularity": GRANULARITY,
    }
    resp = requests.get(url, headers=HEADERS, params=params, timeout=30)
    resp.raise_for_status()
    data = resp.json()

    if not isinstance(data, list):
        msg = data.get("message", "Unknown error format")
        raise RuntimeError(f"Candles error for {pair}: {msg}")

    if not data:
        return 0.0

    # Calculate median daily volume (volume is at index 5 in the candles data)
    volumes = [float(row[5]) for row in data if len(row) > 5]
    return float(np.median(volumes)) if volumes else 0.0


def get_daily_ohlc(pair: str, start: datetime, end: datetime):
    url = f"{BASE_URL}/products/{pair}/candles"
    params = {
        "start": iso_format(start),
        "end": iso_format(end),     # end at UTC midnight -> no partial candle
        "granularity": GRANULARITY,
    }
    resp = requests.get(url, headers=HEADERS, params=params, timeout=30)
    resp.raise_for_status()
    data = resp.json()

    if not isinstance(data, list):
        msg = data.get("message", "Unknown error format")
        raise RuntimeError(f"Candles error for {pair}: {msg}")

    data.sort(key=lambda r: r[0])  # oldest -> newest
    return data


# ----------------------------
# Main
# ----------------------------
def main(volatility_threshold: float = 2.0, days: int = 90, output_file: str = "volatility.xlsx", volume_threshold: float = 1000000.0, output_format: str = "excel", quote_currency: str = "USD"):
    try:
        # Normalize quote currency to uppercase for consistency
        quote_currency = quote_currency.upper()
        
        # Calculate date window based on days parameter
        end_date = datetime.now(timezone.utc).replace(hour=0, minute=0, second=0, microsecond=0)
        start_date = end_date - timedelta(days=days)  # inclusive start
        
        print(f"Window: {start_date.isoformat()} to {end_date.isoformat()} (UTC), {days} FULL daily candles expected.")
        print(f"Using volatility threshold: {volatility_threshold}%")
        print(f"Using volume threshold: {volume_threshold:,.0f}")
        print(f"Quote currency: {quote_currency}")
        print(f"Output format: {output_format.upper()}")
        print(f"Output file: {output_file}")
        print("Press Ctrl+C at any time to cancel the operation, or type 'q' when prompted.\n")
        
        # Safely remove existing file if it exists
        if not safe_remove_file(output_file):
            print("❌ Cannot proceed without removing the existing file.")
            return

        active_pairs = get_active_pairs(quote_currency)
        print(f"Found {len(active_pairs)} active -{quote_currency} pairs.")

        # Initialize output file based on format
        if output_format == "excel":
            wb = create_excel_file(output_file)
        else:
            wb = None  # CSV mode

        # Create progress bar
        progress_bar = tqdm(
            total=len(active_pairs),
            desc="Processing pairs",
            unit="pair",
            bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}, {rate_fmt}]",
            position=0,
            leave=True,
            dynamic_ncols=True
        )

        for i, pair in enumerate(active_pairs, start=1):
            try:
                tqdm.write(f"[{i}/{len(active_pairs)}] Fetching {pair} ...")
                ohlc = get_daily_ohlc(pair, start_date, end_date)

                # Enforce at least the required number of FULL daily candles
                if len(ohlc) < days:
                    tqdm.write(f"  Skipping {pair}: only {len(ohlc)} full day(s) of history in the {days}-day window.")
                    progress_bar.update(1)
                    time.sleep(RATE_LIMIT_SLEEP)
                    continue

                # Compute per-day % range, then the median
                pct_changes = []
                for row in ohlc[-days:]:
                    _, low, high, *_ = row
                    pct_changes.append(calculate_percentage_change(float(low), float(high)))

                median_pct = float(np.median(pct_changes)) if pct_changes else 0.0

                # NEW: Exclude under-threshold medians
                if median_pct < volatility_threshold:
                    tqdm.write(f"  Skipping {pair}: median {median_pct:.4f}% < {volatility_threshold:.2f}% threshold.")
                    progress_bar.update(1)
                    time.sleep(RATE_LIMIT_SLEEP)
                    continue

                tqdm.write(f"  Median daily range% ({days}d): {median_pct:.4f}")

                # Get median daily volume
                try:
                    median_volume = get_daily_volume(pair, start_date, end_date)
                    tqdm.write(f"  Median daily volume: {median_volume:,.2f}")
                except Exception as e:
                    tqdm.write(f"  Warning: could not fetch volume for {pair}: {e}")
                    median_volume = 0.0

                # NEW: Exclude under-threshold volumes
                if median_volume < volume_threshold:
                    tqdm.write(f"  Skipping {pair}: median volume {median_volume:,.2f} < {volume_threshold:,.0f} threshold.")
                    progress_bar.update(1)
                    time.sleep(RATE_LIMIT_SLEEP)
                    continue

                # Get minimum market funds
                try:
                    min_mkt_funds = get_pair_info(pair)
                except Exception as e:
                    tqdm.write(f"  Warning: could not fetch min_market_funds for {pair}: {e}")
                    min_mkt_funds = "N/A"

                # Save data based on format
                if output_format == "excel":
                    save_to_excel(pair, median_pct, median_volume, min_mkt_funds, wb)
                else:
                    if not save_to_csv(pair, median_pct, median_volume, min_mkt_funds, output_file):
                        tqdm.write("❌ Failed to save data. Operation cancelled.")
                        progress_bar.close()
                        return

                tqdm.write(f"  ✅ {pair} added to results")

            except requests.HTTPError as e:
                tqdm.write(f"HTTP error for {pair}: {e} | Response: {getattr(e, 'response', None)}")
            except Exception as e:
                tqdm.write(f"Error for {pair}: {e}")

            progress_bar.update(1)
            time.sleep(RATE_LIMIT_SLEEP)

        # Close progress bar
        progress_bar.close()
        
        # Finalize output based on format
        if output_format == "excel":
            # Sort and format Excel file
            sort_excel_by_volatility(wb)
            format_excel_file(wb)
            
            # Save Excel file
            def save_excel_operation():
                wb.save(output_file)
                return True
            
            if not safe_file_operation(save_excel_operation, output_file, "save Excel file"):
                print("❌ Failed to save Excel file.")
                return
        else:
            # Sort CSV file
            if not sort_csv_by_median(output_file):
                print("❌ Failed to sort the CSV file. Data may be saved but not sorted.")
                return
        
        print(f"\n✅ Done. Results saved (and sorted) in {output_file}")
        
        # Auto-open the file
        print("📂 Opening file...")
        if open_file(output_file):
            print("✅ File opened successfully!")
        else:
            print("ℹ️  File saved but could not be auto-opened.")
        
    except KeyboardInterrupt:
        print(f"\n\n🛑 Operation cancelled by user (Ctrl+C).")
        print("Partial data may have been saved to the output file.")
        if 'progress_bar' in locals():
            progress_bar.close()
        sys.exit(0)


def parse_arguments():
    parser = argparse.ArgumentParser(
        description="Fetch Coinbase coins with volatility above threshold over specified number of days"
    )
    parser.add_argument(
        "--volatility",
        type=float,
        default=2.0,
        help="Minimum daily median volatility percentage threshold (default: 2.0)"
    )
    parser.add_argument(
        "--days",
        type=int,
        default=90,
        help="Number of days to analyze for volatility (default: 90)"
    )
    parser.add_argument(
        "--output",
        type=str,
        default="volatility.xlsx",
        help="Output file name and path (default: volatility.xlsx)"
    )
    parser.add_argument(
        "--volume",
        type=float,
        default=1000000.0,
        help="Minimum median daily volume threshold (default: 1000000.0)"
    )
    parser.add_argument(
        "--format",
        choices=["excel", "csv"],
        default="excel",
        help="Output format: excel or csv (default: excel)"
    )
    parser.add_argument(
        "--quote",
        type=str,
        default="USD",
        help="Quote currency for trading pairs (e.g., USD, BTC, ETH, USDC, USDT) (default: USD)"
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_arguments()
    main(volatility_threshold=args.volatility, days=args.days, output_file=args.output, volume_threshold=args.volume, output_format=args.format, quote_currency=args.quote)
