# OpenAlgo xlwings Lite Edition

A cross-platform Excel add-in for OpenAlgo algorithmic trading, powered by xlwings Lite and Python WebAssembly (Pyodide). **Now featuring Dynamic Response Processing for automatic API format adaptation!**

## Overview

This is the xlwings Lite port of the OpenAlgo Excel Add-in, providing seamless integration with the OpenAlgo API for algorithmic trading. Unlike the original Windows-only Excel-DNA version, this implementation works on **Windows, macOS, and Excel on the web** without requiring local Python installation.

## Key Advantages

✅ **Cross-Platform**: Works on Windows, macOS, and Excel on the web  
✅ **No Installation**: Python runs in browser via WebAssembly  
✅ **Simple Distribution**: Single Excel file contains everything  
✅ **Auto-Updates**: Code updates when workbook is shared  
✅ **Secure**: Runs in browser sandbox  
🆕 **Dynamic API Processing**: Automatically adapts to API response format changes  
🆕 **Smart Formatting**: Intelligent field ordering and data presentation  
🆕 **Excel IntelliSense**: Professional function help text and parameter hints  

## 🆕 NEW: Dynamic Response Features

### Auto-Adaptive Formatting
- **Smart Format Detection**: Automatically chooses optimal display format
- **List/Dict Handling**: Seamlessly handles API format inconsistencies
- **Field Prioritization**: Important fields (symbol, price, quantity) displayed first
- **Smart Value Formatting**: Context-aware formatting for different data types
  - **Price Fields**: Currency formatting with 2 decimals (2,500.00)
  - **Quantity Fields**: Integer formatting with thousands separators (10,000)
  - **Currency Fields**: Large amount formatting (₹1,50,000.00)
  - **Percentage Fields**: Automatic % suffix (5.25%)
  - **Options Greeks**: High-precision formatting (0.1234)
  - **Timestamps**: Readable date-time format (2024-06-22 14:30:00)
- **Enhanced Field Mappings**: 90+ technical field names converted to user-friendly labels

### Configuration Functions
- **`=oa_set_format("auto"|"table"|"key_value")`** - Set display preference
- **`=oa_response_info()`** - Learn about dynamic features
- **`=oa_all_functions()`** - Enhanced function list with new features

### 🆕 Excel IntelliSense Features
- **Function Wizard Integration**: Detailed descriptions appear in Excel's Insert Function dialog
- **Parameter Hints**: IntelliSense shows parameter names and descriptions while typing
- **Help URLs**: Direct links to OpenAlgo API documentation for each function
- **Usage Examples**: Real-world examples included in function descriptions
- **Professional Documentation**: Comprehensive help text like built-in Excel functions

## Prerequisites

- Microsoft Excel 2016+ (Windows/Mac) or Excel on the web
- xlwings Lite add-in installed from Office Add-in Store
- OpenAlgo server running and accessible

## Installation & Setup

### Step 1: Install xlwings Lite
1. Open Excel
2. Go to **Insert > Add-ins > Get Add-ins**
3. Search for "xlwings Lite"
4. Install the xlwings Lite add-in

### Step 2: Configure OpenAlgo Server (CORS Settings)

**IMPORTANT**: To enable Excel Online and cross-origin requests, update your OpenAlgo `.env` file:

```env
# Add this line to your OpenAlgo .env file
CORS_ALLOWED_ORIGINS=http://127.0.0.1:5000,https://addin.xlwings.org

# For custom domains, add them comma-separated:
# CORS_ALLOWED_ORIGINS=http://127.0.0.1:5000,https://addin.xlwings.org,https://yourdomain.com
```

**Without this CORS configuration, you will get connection errors in Excel Online.**

### Step 3: Load OpenAlgo Functions
1. Download the `main.py` file from this repository
2. Open Excel and activate xlwings Lite add-in
3. In xlwings Lite editor, paste the contents of `main.py`
4. Save the workbook as `.xlsm` format

### Step 4: Configure API
```excel
=oa_api("YOUR_API_KEY", "v1", "http://127.0.0.1:5000")
```

### Step 5: Test Connection
```excel
=oa_test_connection()
```

## Available Functions

### 📌 Configuration & Setup
| Function | Description | Dynamic Features |
|----------|-------------|------------------|
| `=oa_api(api_key, version, host_url)` | Set OpenAlgo API credentials | |
| `=oa_get_config()` | View current configuration | Enhanced display |
| `=oa_set_format("auto")` | 🆕 Set response format preference | New feature |
| `=oa_response_info()` | 🆕 Learn about dynamic features | New feature |
| `=oa_test_connection()` | Test API connection | |
| `=oa_force_http(True)` | 🆕 Force HTTP for HTTPS compatibility | Protocol management |
| `=oa_test_https_support()` | 🆕 Test HTTPS support in environment | Diagnostics |
| `=oa_connection_help()` | 🆕 Get help for connection issues | Troubleshooting |

### 📌 Market Data (🔄 Auto-Formatted)
| Function | Description | Dynamic Features |
|----------|-------------|------------------|
| `=oa_quotes("SYMBOL", "EXCHANGE")` | Retrieve market quotes | 🔄 Auto-format, smart field ordering |
| `=oa_depth("SYMBOL", "EXCHANGE")` | Retrieve bid/ask depth | Enhanced price formatting |
| `=oa_history("SYMBOL", "EXCHANGE", "1m", "2024-01-01", "2024-01-31")` | Fetch historical data | Timestamp conversion |
| `=oa_intervals()` | Retrieve available time intervals | 🔄 Auto-format |

### 📌 Account Management (🔄 Auto-Formatted)
| Function | Description | Dynamic Features |
|----------|-------------|------------------|
| `=oa_funds()` | Retrieve available funds | 🔄 Smart key-value display |
| `=oa_orderbook()` | Fetch open order book | 🔄 Auto table format, timestamp conversion |
| `=oa_tradebook()` | Fetch trade book | 🔄 Auto table format, price formatting |
| `=oa_positionbook()` | Fetch position book | 🔄 Auto table format |
| `=oa_holdings()` | Fetch holdings data | 🔄 P&L formatting, percentage display |

### 📌 Order Management
| Function | Description |
|----------|-------------|
| `=oa_placeorder("Strategy", "SYMBOL", "BUY/SELL", "EXCHANGE", "LIMIT", "MIS", "10", "100", "0", "0")` | Place an order |
| `=oa_modifyorder("Strategy", "ORDER_ID", "SYMBOL", "BUY", "NSE", 1, "LIMIT", "MIS", 2500, 0, 0)` | Modify an order |
| `=oa_cancelorder("Strategy", "ORDER_ID")` | Cancel a specific order |
| `=oa_orderstatus("Strategy", "ORDER_ID")` | Retrieve order status |

### 📌 Debug & Diagnostics
| Function | Description |
|----------|-------------|
| `=oa_debug_last_request()` | Show last HTTP request details |
| `=oa_debug_last_response()` | Show last HTTP response details |
| `=oa_debug_full_log()` | Show complete request/response log |

## Usage Examples

### Basic Configuration
```excel
' Set API credentials
=oa_api("your_api_key_here", "v1", "http://127.0.0.1:5000")

' Test connection
=oa_test_connection()

' Set response format preference (optional)
=oa_set_format("auto")
```

### Market Data with Auto-Formatting
```excel
' Get live quotes (auto-formatted with smart field ordering)
=oa_quotes("RELIANCE", "NSE")

' Get market depth (enhanced price formatting)
=oa_depth("INFY", "NSE")

' Get historical data (automatic timestamp conversion)
=oa_history("TATASTEEL", "NSE", "1h", "2024-01-01", "2024-01-31")
```

### Account Information (Enhanced Display)
```excel
' Check available funds (smart key-value display)
=oa_funds()

' View current positions (auto table format)
=oa_positionbook()

' Check order book (timestamp conversion, field prioritization)
=oa_orderbook()
```

### Response Format Customization
```excel
' Force table format for all functions
=oa_set_format("table")

' Force key-value format
=oa_set_format("key_value")

' Auto-detect best format (default)
=oa_set_format("auto")

' Learn about dynamic features
=oa_response_info()
```

## CORS Configuration Details

### Why CORS Configuration is Needed

xlwings Lite runs Python in the browser using WebAssembly. When Excel Online makes API requests to your OpenAlgo server, browsers enforce CORS (Cross-Origin Resource Sharing) policies. Without proper CORS headers, requests will be blocked.

### OpenAlgo .env File Setup

Add or update this line in your OpenAlgo `.env` file:

```env
# For local development and Excel Online
CORS_ALLOWED_ORIGINS=http://127.0.0.1:5000,https://addin.xlwings.org

# For production with custom domains
CORS_ALLOWED_ORIGINS=http://127.0.0.1:5000,https://addin.xlwings.org,https://yourdomain.com,https://your-openalgo-domain.com

# For localhost variations (if needed)
CORS_ALLOWED_ORIGINS=http://127.0.0.1:5000,http://localhost:5000,https://addin.xlwings.org
```

### Restart OpenAlgo Server
After updating the `.env` file, restart your OpenAlgo server:
```bash
# Stop the server (Ctrl+C)
# Then restart
python app.py  # or however you run OpenAlgo
```

### Testing CORS Configuration
```excel
' This should return "SUCCESS" if CORS is properly configured
=oa_test_connection()
```

## 🆕 Dynamic Response System

### How It Works
The new dynamic response system automatically:
1. **Detects Response Structure**: Identifies if API returns list or dictionary format
2. **Chooses Optimal Display**: Selects table or key-value format based on data
3. **Orders Fields Intelligently**: Prioritizes important fields like symbol, price, quantity
4. **Formats Values**: Applies currency formatting, percentage signs, readable timestamps
5. **Handles API Changes**: Adapts automatically if OpenAlgo changes response formats

### Benefits Over Manual Formatting
- **85% Less Code**: Functions are now 3-10 lines instead of 50+
- **Automatic Adaptation**: No manual updates needed for API changes
- **Consistent Display**: Professional formatting across all functions
- **Better User Experience**: Readable field names and proper value formatting

### Enhanced Field Mappings
**90+ technical field names** automatically converted to user-friendly labels:

#### Core Trading Fields:
- `ltp` → `Last Trade Price`
- `prev_close` → `Previous Close`
- `pnl` → `P&L`
- `orderid` → `Order ID`
- `tradingsymbol` → `Trading Symbol`

#### Account & Fund Fields:
- `availablecash` → `Available Cash`
- `m2mrealized` → `Realized M2M`
- `m2munrealized` → `Unrealized M2M`
- `utiliseddebits` → `Used Debits`
- `collateral` → `Collateral Value`

#### Order Management Fields:
- `triggerprice` → `Trigger Price`
- `averageprice` → `Average Price`
- `remainingquantity` → `Remaining Qty`
- `filledquantity` → `Filled Qty`
- `order_status` → `Order Status`

#### Market Data Fields:
- `bid_price` → `Bid Price`
- `ask_price` → `Ask Price`
- `total_traded_volume` → `Total Volume`
- `upper_circuit` → `Upper Circuit`
- `day_high` → `Day High`

#### Options Trading Fields:
- `strikeprice` → `Strike Price`
- `optiontype` → `Option Type`
- `implied_volatility` → `IV`
- `days_to_expiry` → `Days to Expiry`
- Greeks: `delta`, `gamma`, `theta`, `vega`, `rho`

#### Position & P&L Fields:
- `unrealized_pnl` → `Unrealized P&L`
- `net_quantity` → `Net Quantity`
- `buy_value` → `Buy Value`
- `margin_required` → `Margin Required`

## Error Handling & Debugging

### Enhanced Error Messages
```excel
' Clear, actionable error messages
Error: OpenAlgo API Key is not set. Use oa_api()
Error: HTTP Error 401: Unauthorized
Error: No data received from API
```

### Debug Functions
```excel
' See exactly what was sent to API
=oa_debug_last_request()

' See exactly what API returned
=oa_debug_last_response()

' Complete request/response cycle
=oa_debug_full_log()
```

### Connection Troubleshooting
1. **Test Connection**: `=oa_test_connection()`
2. **Check CORS**: Ensure `.env` file is updated
3. **Verify API Key**: Use `=oa_get_config()`
4. **Check Server**: Ensure OpenAlgo is running
5. **HTTPS Issues**: Use `=oa_test_https_support()` for diagnostics
6. **Protocol Problems**: Try `=oa_force_http(True)` for compatibility

## File Structure

```
OpenAlgo_xlwings_lite/
├── main.py              # Complete xlwings Lite implementation with dynamic features
├── requirements.txt     # Python dependencies
└── README.md           # This documentation
```

## Dependencies

Located in `requirements.txt`:
- xlwings==0.33.14 (required)
- python-dotenv==1.1.0 (required)

Optional (loaded dynamically if available):
- pandas (enhanced data manipulation)
- pyodide_http (WebAssembly HTTP patching)

## Migration from Excel-DNA

### Function Compatibility
All functions work identically with enhanced features:
- **Same Function Names**: All `oa_*` functions
- **Same Parameters**: Identical function signatures  
- **Enhanced Output**: Better formatting and display
- **Same Error Handling**: Consistent error messages

### New Features Available
- Dynamic response formatting
- Smart field ordering
- Automatic value formatting
- User-configurable display preferences

## Performance & Compatibility

### Browser Support
- **Edge**: Full support ✅
- **Chrome**: Full support ✅  
- **Safari**: Full support ✅
- **Firefox**: Full support ✅

### Performance Considerations
- WebAssembly adds ~100-200ms overhead
- Dynamic formatting adds minimal processing time
- Recommended for normal trading operations
- Use pagination for large historical data requests

## Security Considerations

- Code runs in browser sandbox for security
- API keys stored in Excel workbook (use with caution)
- HTTPS recommended for OpenAlgo server connections
- CORS configuration restricts access to authorized domains
- Test in demo mode before live trading

## Troubleshooting

### Common CORS Issues

**Error: Network request failed**
```
Solution: Add https://addin.xlwings.org to CORS_ALLOWED_ORIGINS in OpenAlgo .env file
```

**Functions work locally but not in Excel Online**
```
Solution: Ensure CORS_ALLOWED_ORIGINS includes https://addin.xlwings.org
```

### HTTPS Compatibility Issues

**Error: URL Error: unknown url type: https**
```excel
' This error occurs when HTTPS is not supported in the xlwings Lite environment
' Solution 1: Enable automatic HTTP fallback (recommended)
=oa_force_http(True)

' Solution 2: Use HTTP URL in configuration
=oa_api("API_KEY", "v1", "http://127.0.0.1:5000")

' Solution 3: Test HTTPS support first
=oa_test_https_support()
```

**HTTPS works sometimes but not always**
```excel
' The system automatically falls back to HTTP when HTTPS fails
' Check configuration and diagnostics
=oa_get_config()
=oa_connection_help()
```

### Function Issues

**Functions return `#NAME?`**
- Ensure xlwings Lite add-in is installed and active
- Verify `main.py` is loaded in xlwings editor

**Functions return `#VALUE!`**
- Check API connection with `=oa_test_connection()`
- Verify OpenAlgo server is running
- Check CORS configuration

**Slow Performance**
- Use `=oa_set_format("table")` for large datasets
- Minimize real-time data refresh frequency

## Advanced Configuration

### Custom Response Formatting
```excel
' Set global format preference
=oa_set_format("table")     ' Always use table format
=oa_set_format("key_value") ' Always use key-value format  
=oa_set_format("auto")      ' Let system decide (default)
```

### Environment-Specific Settings
```excel
' For local development
=oa_api("API_KEY", "v1", "http://127.0.0.1:5000")

' For production server
=oa_api("API_KEY", "v1", "https://your-openalgo-server.com")
```

## Support

For issues specific to this xlwings Lite implementation:
1. Check CORS configuration first
2. Test with `=oa_test_connection()`
3. Use debug functions to inspect requests/responses
4. Verify xlwings Lite add-in is active

For OpenAlgo API issues:
- Refer to [OpenAlgo API Documentation](https://docs.openalgo.in/api-documentation/v1/)

## License

This xlwings Lite implementation follows the same license as the original OpenAlgo Excel Add-in.

## Disclaimer

This add-in is provided as-is. Test thoroughly in demo/paper trading mode before using with real money. The creators are not responsible for any trading losses.

---

🚀 **Happy Trading with Cross-Platform Support and Dynamic API Processing!**

### What's New in This Version
- ✨ **Dynamic Response Processing**: Auto-adapts to API format changes
- 🎯 **Smart Field Ordering**: Important fields displayed first  
- 💰 **Enhanced Formatting**: Automatic price, percentage, and timestamp formatting
- 🔧 **User Configuration**: Control display preferences with `oa_set_format()`
- 🐛 **Better Debugging**: Comprehensive request/response logging
- 🌐 **CORS Guide**: Complete setup instructions for Excel Online compatibility
- 📚 **Excel IntelliSense**: Professional function documentation with parameter hints and help links
- 🏷️ **90+ Field Mappings**: Comprehensive technical field name conversion to user-friendly labels