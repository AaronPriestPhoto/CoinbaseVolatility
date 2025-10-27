# Coinbase Volatility Analyzer

A Python script that analyzes Coinbase trading pairs to find coins with high volatility over a specified time period. The script fetches historical data, calculates median daily volatility and volume, and exports results to Excel or CSV files with professional formatting.

## Features

- **Volatility Analysis**: Finds trading pairs with median daily volatility above a specified threshold
- **Volume Analysis**: Calculates median daily volume for each pair with filtering options
- **SuperTrend Analysis**: Analyzes 30-minute SuperTrend sessions with Factor 3 and ATR Length 10
- **Trading Session Metrics**: Calculates median and maximum % changes in SuperTrend sessions
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
git clone https://github.com/AaronPriestPhoto/CoinbaseVolatility.git
cd CoinbaseVolatility
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
| `--supertrend` | Number of top volatile coins to analyze with SuperTrend (0 = disabled) | 0 |
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

# SuperTrend analysis for top 20 most volatile pairs
python CoinbaseVolatility.py --supertrend 20

# SuperTrend analysis for top 50 pairs with custom thresholds
python CoinbaseVolatility.py --volatility 3 --days 30 --supertrend 50 --output supertrend_analysis.xlsx

# Bitcoin pairs with SuperTrend for top 10
python CoinbaseVolatility.py --quote BTC --supertrend 10 --volume 2000000
```

## Output

The script generates Excel (.xlsx) or CSV files with the following columns:

| Column | Description |
|--------|-------------|
| `Pair` | Trading pair (e.g., BTC-USD, ETH-BTC) |
| `Volatility` | Median daily volatility percentage (2 decimal places) |
| `Volume` | Median daily volume (formatted with commas) |
| `MinFunds` | Minimum trade size |
| `MedLong%` | Median % rise in SuperTrend long sessions |
| `MaxLong%` | Maximum % rise in any SuperTrend long session |
| `MedShort%` | Median % fall in SuperTrend short sessions |
| `MaxShort%` | Maximum % fall in any SuperTrend short session |
| `Sessions` | Total number of SuperTrend sessions in period |

### Excel Output Features

- **Professional formatting** with styled headers
- **Auto-sized columns** for optimal readability
- **Number formatting** with proper decimal places and commas
- **Sorted by volatility** (highest first)
- **No Excel warnings** - clean, error-free display

### Example Output

**Excel Format:**
```
Pair        | Volatility | Volume      | MinFunds | MedLong% | MaxLong% | MedShort% | MaxShort% | Sessions
BTC-USD     | 2.46       | 1,234,568   | 10.00    | 3.25     | 8.45      | 2.15      | 5.80      | 12
ETH-USD     | 3.12       | 987,654     | 10.00    | 4.20     | 9.80      | 2.85      | 6.20      | 15
ADA-USD     | 4.57       | 456,789     | 10.00    | 5.10     | 12.30     | 3.40      | 7.90      | 18
```

**CSV Format:**
```csv
Pair,Volatility,Volume,MinFunds,MedLong%,MaxLong%,MedShort%,MaxShort%,Sessions
BTC-USD,2.46,1,234,568,10.00,3.25,8.45,2.15,5.80,12
ETH-USD,3.12,987,654,10.00,4.20,9.80,2.85,6.20,15
ADA-USD,4.57,456,789,10.00,5.10,12.30,3.40,7.90,18
```

## How It Works

1. **Fetches Active Pairs**: Gets all active trading pairs for the specified quote currency from Coinbase
2. **Historical Data**: Downloads daily OHLC data for the specified time period
3. **Volatility Calculation**: Calculates median daily volatility (high-low range percentage)
4. **Volume Analysis**: Calculates median daily volume
5. **Filtering**: Excludes pairs with insufficient data, low volatility, or low volume
6. **SuperTrend Analysis**: Downloads 30-minute candles and calculates SuperTrend sessions (Factor 3, ATR Length 10)
7. **Session Metrics**: Calculates median and maximum % changes in SuperTrend long/short sessions
8. **Export**: Saves results to Excel or CSV, sorted by volatility (highest first)
9. **Auto-Open**: Automatically opens the output file when complete

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

## SuperTrend Analysis

The script includes advanced SuperTrend analysis for qualifying trading pairs:

### SuperTrend Configuration
- **Timeframe**: 30-minute candles
- **Factor**: 3 (SuperTrend multiplier)
- **ATR Length**: 10 periods
- **ATR Method**: RMA (Wilder's smoothing) - matches TradingView default
- **Analysis Period**: Same as volatility analysis (90 days default)

### Session Analysis
- **Long Sessions**: Tracks % rise from SuperTrend buy signal to highest point in session
- **Short Sessions**: Tracks % fall from SuperTrend sell signal to lowest point in session
- **Session Metrics**: Calculates median and maximum % changes for each session type
- **Total Sessions**: Counts all SuperTrend signal changes in the analysis period

### Trading Insights
- **MedLong%**: Median percentage gain in long SuperTrend sessions
- **MaxLong%**: Maximum percentage gain in any single long session
- **MedShort%**: Median percentage gain in short SuperTrend sessions (fall protection)
- **MaxShort%**: Maximum percentage gain in any single short session
- **Sessions**: Total number of SuperTrend signals (indicates trading frequency)

### Data Efficiency
- SuperTrend analysis is only performed on pairs that pass volatility and volume filters
- 30-minute data is downloaded only for qualifying pairs to minimize bandwidth usage
- Analysis uses the same time period as volatility calculations for consistency

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

### v3.0.0
- **SuperTrend Analysis** with 30-minute candles (Factor 3, ATR Length 10)
- **TradingView Compatible** - Uses RMA (Wilder's smoothing) for ATR calculation
- **Trading Session Metrics** - Median and maximum % changes in SuperTrend sessions
- **Enhanced Output** with 5 new SuperTrend columns (MedLong%, MaxLong%, MedShort%, MaxShort%, Sessions)
- **Data Efficiency** - SuperTrend analysis only for qualifying pairs
- **Self-contained Script** - Works from any directory (StreamDeck compatible)
- **Comprehensive Documentation** - Detailed SuperTrend analysis explanation

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
