# Coinbase Volatility Analyzer

A Python script that analyzes Coinbase trading pairs to find coins with high volatility over a specified time period. The script fetches historical data, calculates median daily volatility and volume, and exports results to Excel or CSV files with professional formatting.

## Features

- **Volatility Analysis**: Finds trading pairs with median daily volatility above a specified threshold
- **Volume Analysis**: Calculates median daily volume for each pair with filtering options
- **Flexible Time Periods**: Analyze any number of days (30, 90, 365, etc.)
- **Multiple Quote Currencies**: Support for USD, BTC, ETH, USDC, USDT, and more
- **Excel & CSV Output**: Professional Excel formatting with auto-sized columns or CSV export
- **Volume Filtering**: Filter by minimum median daily volume
- **Progress Tracking**: Beautiful progress bar with time estimates
- **Robust File Handling**: Smart file handling with user prompts when files are in use
- **Auto-Open Results**: Automatically opens output files when complete
- **Rate Limiting**: Built-in API rate limiting to respect Coinbase's limits

## Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/coinbase-volatility-analyzer.git
cd coinbase-volatility-analyzer
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

```bash
# Analyze last 90 days with 2% volatility threshold (default Excel output)
python CoinbaseVolatility.py

# Analyze last 30 days
python CoinbaseVolatility.py --days 30

# Use 5% volatility threshold
python CoinbaseVolatility.py --volatility 5

# Bitcoin-quoted pairs
python CoinbaseVolatility.py --quote BTC

# Custom output file
python CoinbaseVolatility.py --output my_analysis.xlsx

# Combine options
python CoinbaseVolatility.py --volatility 3 --days 60 --volume 2000000 --quote ETH
```

### Command Line Options

| Option | Description | Default |
|--------|-------------|---------|
| `--volatility` | Minimum daily median volatility percentage threshold | 2.0 |
| `--days` | Number of days to analyze for volatility | 90 |
| `--volume` | Minimum median daily volume threshold | 1000000.0 |
| `--quote` | Quote currency (USD, BTC, ETH, USDC, USDT, etc.) | USD |
| `--format` | Output format: excel or csv | excel |
| `--output` | Output file name and path | volatility.xlsx |
| `--help` | Show help message | - |

### Examples

```bash
# Find very volatile coins (5%+ daily volatility) over 30 days
python CoinbaseVolatility.py --volatility 5 --days 30 --output high_vol_30d.xlsx

# Bitcoin-quoted pairs with high volume
python CoinbaseVolatility.py --quote BTC --volume 5000000 --volatility 3

# Conservative analysis - 1% volatility over 1 year
python CoinbaseVolatility.py --volatility 1 --days 365 --output conservative_1y.xlsx

# USDC pairs with CSV output
python CoinbaseVolatility.py --quote USDC --format csv --output usdc_pairs.csv

# Quick 7-day analysis of Ethereum ecosystem
python CoinbaseVolatility.py --days 7 --quote ETH --output eth_weekly.xlsx
```

## Output

The script generates Excel (.xlsx) or CSV files with the following columns:

| Column | Description |
|--------|-------------|
| `Pair` | Trading pair (e.g., BTC-USD, ETH-BTC) |
| `Volatility` | Median daily volatility percentage (2 decimal places) |
| `Volume` | Median daily volume (formatted with commas) |
| `MinFunds` | Minimum trade size |

### Excel Output Features

- **Professional formatting** with styled headers
- **Auto-sized columns** for optimal readability
- **Number formatting** with proper decimal places and commas
- **Sorted by volatility** (highest first)
- **No Excel warnings** - clean, error-free display

### Example Output

**Excel Format:**
```
Pair        | Volatility | Volume      | MinFunds
BTC-USD     | 2.46       | 1,234,568   | 10.00
ETH-USD     | 3.12       | 987,654     | 10.00
ADA-USD     | 4.57       | 456,789     | 10.00
```

**CSV Format:**
```csv
Pair,Volatility,Volume,MinFunds
BTC-USD,2.46,1,234,568,10.00
ETH-USD,3.12,987,654,10.00
ADA-USD,4.57,456,789,10.00
```

## How It Works

1. **Fetches Active Pairs**: Gets all active trading pairs for the specified quote currency from Coinbase
2. **Historical Data**: Downloads daily OHLC data for the specified time period
3. **Volatility Calculation**: Calculates median daily volatility (high-low range percentage)
4. **Volume Analysis**: Calculates median daily volume
5. **Filtering**: Excludes pairs with insufficient data, low volatility, or low volume
6. **Export**: Saves results to Excel or CSV, sorted by volatility (highest first)
7. **Auto-Open**: Automatically opens the output file when complete

## File Handling

The script includes intelligent file handling:

- **File in Use Detection**: Prompts user when output file is open in another program
- **User-Friendly Prompts**: Simple "q" to quit or Enter to retry
- **Graceful Cancellation**: Ctrl+C support throughout the entire process
- **Progress Tracking**: Visual progress bar with time estimates

## API Rate Limiting

The script includes built-in rate limiting to respect Coinbase's API limits:
- ~3 requests per second (configurable)
- Automatic delays between requests
- Error handling for rate limit responses

## Requirements

- Python 3.7+
- Internet connection
- Coinbase API access (public endpoints)

## Dependencies

- `numpy` - Statistical calculations
- `requests` - HTTP requests to Coinbase API
- `tqdm` - Progress bar display
- `openpyxl` - Excel file creation and formatting

## Error Handling

The script handles various error conditions:

- **Network Issues**: Retries and graceful error messages
- **File Access**: User prompts when files are in use
- **API Errors**: Clear error messages for API issues
- **Data Issues**: Skips pairs with insufficient historical data

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Disclaimer

This tool is for educational and research purposes only. Always do your own research before making investment decisions. Cryptocurrency trading involves significant risk.

## Support

If you encounter any issues or have questions:

1. Check the [Issues](https://github.com/yourusername/coinbase-volatility-analyzer/issues) page
2. Create a new issue with detailed information
3. Include your Python version and error messages

## Changelog

### v2.0.0
- **Excel output as default** with professional formatting
- **Multiple quote currencies** (USD, BTC, ETH, USDC, USDT, etc.)
- **Volume filtering** with configurable thresholds
- **Case-insensitive quote currency** input
- **Auto-sized Excel columns** with proper number formatting
- **Auto-open results** when analysis completes
- **Enhanced progress tracking** with tqdm
- **Improved file handling** with user prompts

### v1.0.0
- Initial release
- Volatility analysis with configurable thresholds
- Volume analysis
- Flexible time periods
- Progress tracking
- Robust file handling
- Command-line interface
