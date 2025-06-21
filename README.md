# OpenAlgo xlwings Lite Edition

A cross-platform Excel add-in for OpenAlgo algorithmic trading, powered by xlwings Lite and Python WebAssembly (Pyodide).

## Overview

This is the xlwings Lite port of the OpenAlgo Excel Add-in, providing seamless integration with the OpenAlgo API for algorithmic trading. Unlike the original Windows-only Excel-DNA version, this implementation works on **Windows, macOS, and Excel on the web** without requiring local Python installation.

## Key Advantages

✅ **Cross-Platform**: Works on Windows, macOS, and Excel on the web  
✅ **No Installation**: Python runs in browser via WebAssembly  
✅ **Simple Distribution**: Single Excel file contains everything  
✅ **Auto-Updates**: Code updates when workbook is shared  
✅ **Secure**: Runs in browser sandbox  

## Features

- **Account Management**: Retrieve funds, order books, trade books, and position books
- **Market Data**: Fetch real-time quotes, depth, historical data, and available intervals
- **Order Management**: Place, modify, cancel, and retrieve order statuses
- **Smart & Basket Orders**: Execute split, smart, and bulk orders
- **Risk Management**: Close all open positions for a given strategy
- **Automation Scripts**: Helper functions for dashboard creation and data refresh

## Prerequisites

- Microsoft Excel 2016+ (Windows/Mac) or Excel on the web
- xlwings Lite add-in installed from Office Add-in Store
- OpenAlgo server running and accessible

## Installation

### Step 1: Install xlwings Lite
1. Open Excel
2. Go to **Insert > Add-ins > Get Add-ins**
3. Search for "xlwings Lite"
4. Install the xlwings Lite add-in

### Step 2: Load OpenAlgo Functions
1. Download the `main.py` file from this repository
2. Open Excel and activate xlwings Lite add-in
3. In xlwings Lite editor, paste the contents of `main.py`
4. Save the workbook as `.xlsm` format

### Step 3: Configure API
```excel
=oa_api("YOUR_API_KEY", "v1", "http://127.0.0.1:5000")
```

## Available Functions

### 📌 Configuration
| Function | Description |
|----------|-------------|
| `=oa_api(api_key, version, host_url)` | Set OpenAlgo API credentials |
| `=oa_get_config()` | View current configuration |
| `=oa_test_connection()` | Test API connection |

### 📌 Market Data
| Function | Description |
|----------|-------------|
| `=oa_quotes("SYMBOL", "EXCHANGE")` | Retrieve market quotes |
| `=oa_depth("SYMBOL", "EXCHANGE")` | Retrieve bid/ask depth |
| `=oa_history("SYMBOL", "EXCHANGE", "1m", "2024-01-01", "2024-01-31")` | Fetch historical data |
| `=oa_intervals()` | Retrieve available time intervals |

### 📌 Account Management
| Function | Description |
|----------|-------------|
| `=oa_funds()` | Retrieve available funds |
| `=oa_orderbook()` | Fetch open order book |
| `=oa_tradebook()` | Fetch trade book |
| `=oa_positionbook()` | Fetch position book |
| `=oa_holdings()` | Fetch holdings data |

### 📌 Order Management
| Function | Description |
|----------|-------------|
| `=oa_placeorder("Strategy", "SYMBOL", "BUY/SELL", "EXCHANGE", "LIMIT", "MIS", "10", "100", "0", "0")` | Place an order |
| `=oa_placesmartorder("Strategy", "SYMBOL", "BUY/SELL", "EXCHANGE", "LIMIT", "MIS", "10", "100", "0", "0", "0")` | Place a smart order |
| `=oa_basketorder("Strategy", A1:I10)` | Place multiple orders in a basket |
| `=oa_splitorder("Strategy", "SYMBOL", "BUY/SELL", "EXCHANGE", "100", "10", "LIMIT", "MIS", "100", "0", "0")` | Place split order |
| `=oa_modifyorder("Strategy", "ORDER_ID", "SYMBOL", "BUY", "NSE", 1, "LIMIT", "MIS", 2500, 0, 0)` | Modify an order |
| `=oa_cancelorder("Strategy", "ORDER_ID")` | Cancel a specific order |
| `=oa_cancelallorder("Strategy")` | Cancel all orders for a strategy |
| `=oa_closeposition("Strategy")` | Close all open positions for a strategy |
| `=oa_orderstatus("Strategy", "ORDER_ID")` | Retrieve order status |
| `=oa_openposition("Strategy", "SYMBOL", "EXCHANGE", "MIS")` | Fetch open positions |

## Automation Scripts

xlwings Lite also provides automation scripts that can be run from the xlwings editor:

- `refresh_all_quotes()` - Refresh all quote formulas in active sheet
- `setup_dashboard()` - Create sample trading dashboard
- `refresh_all_data()` - Refresh all OpenAlgo formulas

## Usage Examples

### Basic Configuration
```excel
' Set API credentials
=oa_api("your_api_key_here", "v1", "http://127.0.0.1:5000")

' Test connection
=oa_test_connection()
```

### Market Data
```excel
' Get live quotes
=oa_quotes("RELIANCE", "NSE")

' Get market depth
=oa_depth("INFY", "NSE")

' Get historical data
=oa_history("TATASTEEL", "NSE", "1h", "2024-01-01", "2024-01-31")
```

### Account Information
```excel
' Check available funds
=oa_funds()

' View current positions
=oa_positionbook()

' Check order book
=oa_orderbook()
```

### Order Placement
```excel
' Place a market order
=oa_placeorder("MyStrategy", "RELIANCE", "BUY", "NSE", "MARKET", "MIS", "1", "0", "0", "0")

' Place a limit order
=oa_placeorder("MyStrategy", "INFY", "SELL", "NSE", "LIMIT", "MIS", "10", "1500", "0", "0")
```

## Differences from Excel-DNA Version

### Similarities
- **Identical Function Names**: All `oa_*` functions work exactly the same
- **Same Parameters**: Function signatures are maintained
- **Same Output Format**: Returns identical Excel-friendly tables
- **Same Error Handling**: Consistent error messages

### Key Differences
- **Platform Support**: Works on Windows, macOS, and web
- **No Installation**: Runs in browser without local Python
- **Code Storage**: Python code stored in Excel workbook
- **Deployment**: Single `.xlsm` file distribution
- **Performance**: Slight overhead due to WebAssembly

## xlwings Lite Specific Features

### WebAssembly Compatibility
- Uses `urllib` instead of `requests` for HTTP calls
- Automatic `pyodide_http` patching for browser compatibility
- Pure Python implementation - no compiled extensions

### Error Handling
All functions return Excel-friendly error messages:
```
Error: OpenAlgo API Key is not set. Use oa_api()
Error: HTTP Error 401: Unauthorized
Error: No data found
```

### Data Formatting
- Timestamps automatically converted to IST
- All values converted to strings for Excel compatibility
- Consistent 2D array output format
- Proper headers for tabular data

## Development

### File Structure
```
OpenAlgo_xlwings_lite/
├── main.py              # Complete xlwings Lite implementation
├── requirements.txt     # Python dependencies
└── README.md           # This documentation
```

### Dependencies
- xlwings==0.33.14 (required)
- python-dotenv==1.1.0 (required)
- pandas (data manipulation)
- black (code formatting)

### Testing
Use the built-in test functions:
```excel
=oa_test_connection()    # Test API connectivity
=oa_get_config()         # View current settings
```

## Migration from Excel-DNA

If you're migrating from the Excel-DNA version:

1. **Function Compatibility**: All functions work identically
2. **Formula Updates**: No changes needed to existing formulas
3. **Parameter Handling**: Same parameter requirements
4. **Output Format**: Identical table structures

Simply replace your Excel-DNA add-in with this xlwings Lite version.

## Troubleshooting

### Common Issues

**Functions return `#NAME?`**
- Ensure xlwings Lite add-in is installed and active
- Verify `main.py` is loaded in xlwings editor

**API Connection Errors**
- Check OpenAlgo server is running
- Verify API key is correct using `=oa_test_connection()`
- Ensure firewall allows connections to OpenAlgo server

**Slow Performance**
- WebAssembly has slight overhead vs native code
- Minimize large data transfers
- Use pagination for historical data

### Browser Compatibility
- **Edge**: Full support ✅
- **Chrome**: Full support ✅  
- **Safari**: Full support ✅
- **Firefox**: Limited WebAssembly support ⚠️

## Security Considerations

- Code runs in browser sandbox for security
- API keys stored in Excel workbook (use with caution)
- HTTPS recommended for OpenAlgo server connections
- Test in demo mode before live trading

## Support

For issues specific to this xlwings Lite implementation:
1. Check xlwings Lite documentation
2. Verify Pyodide package compatibility
3. Test with `oa_test_connection()` function

For OpenAlgo API issues:
- Refer to [OpenAlgo API Documentation](https://docs.openalgo.in/api-documentation/v1/)

## License

This xlwings Lite implementation follows the same license as the original OpenAlgo Excel Add-in.

## Disclaimer

This add-in is provided as-is. Test thoroughly in demo/paper trading mode before using with real money. The creators are not responsible for any trading losses.

🚀 **Happy Trading with Cross-Platform Support!**
