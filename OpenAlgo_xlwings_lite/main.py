"""OpenAlgo Excel Integration via xlwings Lite

This module provides all OpenAlgo trading functions for Excel via xlwings Lite.
xlwings Lite runs Python in the browser using WebAssembly (Pyodide), providing
cross-platform support for Windows, macOS, and Excel on the web.
"""

import xlwings as xw
from xlwings import func, script
import json
import urllib.request
import urllib.parse
from datetime import datetime

# Optional imports with proper error handling
try:
    from typing import List, Any, Optional
    TYPING_AVAILABLE = True
except ImportError:
    TYPING_AVAILABLE = False
    # Fallback for environments without typing
    List = list
    Any = object
    Optional = object

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False
    pd = None

# Try to import pyodide_http for WebAssembly compatibility
try:
    import pyodide_http
    pyodide_http.patch_all()
    PYODIDE_AVAILABLE = True
except ImportError:
    # Running in standard Python environment
    PYODIDE_AVAILABLE = False

# Global Configuration Storage
class OpenAlgoConfig:
    """Global configuration for OpenAlgo API"""
    api_key = ""
    version = "v1"
    host_url = "http://127.0.0.1:5000"

# Global debug storage for request/response logging
class DebugLog:
    """Store request/response logs for debugging"""
    last_request = None
    last_response = None
    request_count = 0

# Response formatting configuration
class ResponseConfig:
    """Configuration for dynamic response formatting"""
    # Display preferences
    preferred_format = "auto"  # auto, table, key_value
    max_nested_depth = 3
    timestamp_format = '%Y-%m-%d %H:%M:%S'
    
    # Field mappings for better display names
    field_labels = {
        'ltp': 'Last Trade Price',
        'prev_close': 'Previous Close',
        'pnl': 'P&L',
        'pnl_percent': 'P&L %',
        'orderid': 'Order ID',
        'tradingsymbol': 'Trading Symbol'
    }
    
    # Fields to prioritize in display order
    priority_fields = ['symbol', 'ltp', 'price', 'quantity', 'status', 'orderid']
    
    # Endpoints with known response patterns
    endpoint_schemas = {
        'quotes': {'format': 'key_value', 'title_field': 'symbol'},
        'funds': {'format': 'key_value', 'title': 'Account Funds'},
        'orderbook': {'format': 'table', 'sort_by': 'timestamp'},
        'tradebook': {'format': 'table', 'sort_by': 'timestamp'},
        'positionbook': {'format': 'table'},
        'holdings': {'format': 'table'}
    }

# Utility Functions
def post_request(endpoint, payload):
    """Make HTTP POST request using urllib (Pyodide compatible)"""
    try:
        # Log the request
        DebugLog.request_count += 1
        DebugLog.last_request = {
            "endpoint": endpoint,
            "payload": payload,
            "timestamp": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            "request_id": DebugLog.request_count
        }
        
        print(f"[REQUEST {DebugLog.request_count}] {endpoint}")
        print(f"[PAYLOAD {DebugLog.request_count}] {json.dumps(payload, indent=2)}")
        
        data = json.dumps(payload).encode('utf-8')
        headers = {
            'Content-Type': 'application/json',
            'Accept': 'application/json'
        }
        
        request = urllib.request.Request(endpoint, data=data, headers=headers)
        response = urllib.request.urlopen(request, timeout=30)
        
        response_data = json.loads(response.read().decode('utf-8'))
        
        # Log the response
        DebugLog.last_response = {
            "status_code": response.getcode(),
            "data": response_data,
            "timestamp": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            "request_id": DebugLog.request_count
        }
        
        print(f"[RESPONSE {DebugLog.request_count}] Status: {response.getcode()}")
        print(f"[DATA {DebugLog.request_count}] {json.dumps(response_data, indent=2)}")
        
        return response_data
        
    except urllib.error.HTTPError as e:
        error_msg = f"HTTP Error {e.code}: {e.reason}"
        DebugLog.last_response = {
            "error": error_msg,
            "status_code": e.code,
            "timestamp": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            "request_id": DebugLog.request_count
        }
        print(f"[ERROR {DebugLog.request_count}] {error_msg}")
        return {"error": error_msg}
        
    except urllib.error.URLError as e:
        error_msg = f"URL Error: {e.reason}"
        DebugLog.last_response = {
            "error": error_msg,
            "timestamp": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            "request_id": DebugLog.request_count
        }
        print(f"[ERROR {DebugLog.request_count}] {error_msg}")
        return {"error": error_msg}
        
    except json.JSONDecodeError as e:
        error_msg = f"JSON Decode Error: {str(e)}"
        DebugLog.last_response = {
            "error": error_msg,
            "timestamp": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            "request_id": DebugLog.request_count
        }
        print(f"[ERROR {DebugLog.request_count}] {error_msg}")
        return {"error": error_msg}
        
    except Exception as e:
        error_msg = str(e)
        DebugLog.last_response = {
            "error": error_msg,
            "timestamp": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            "request_id": DebugLog.request_count
        }
        print(f"[ERROR {DebugLog.request_count}] {error_msg}")
        return {"error": error_msg}

def detect_endpoint_type(endpoint):
    """Extract endpoint type from URL for schema detection"""
    if not endpoint:
        return 'unknown'
    
    # Extract the last part of the API path
    parts = endpoint.split('/')
    if len(parts) >= 2:
        endpoint_type = parts[-1].lower()
        return endpoint_type
    return 'unknown'

def smart_format_value(key, value):
    """Apply intelligent formatting to field values"""
    if value is None or value == "":
        return ""
    
    # Handle timestamps
    if key.lower() in ['timestamp', 'date', 'time'] and isinstance(value, (int, float)):
        try:
            dt = datetime.fromtimestamp(value)
            return dt.strftime(ResponseConfig.timestamp_format)
        except (ValueError, TypeError, OSError):
            pass
    
    # Handle numeric formatting
    if key.lower() in ['price', 'ltp', 'high', 'low', 'open', 'close', 'trigger_price']:
        try:
            num_val = float(value)
            return f"{num_val:.2f}" if num_val != 0 else "0.00"
        except (ValueError, TypeError):
            pass
    
    # Handle percentage fields
    if 'percent' in key.lower() or key.lower().endswith('_pct'):
        try:
            num_val = float(value)
            return f"{num_val:.2f}%"
        except (ValueError, TypeError):
            pass
    
    return str(value)

def get_display_label(field_name):
    """Get user-friendly display label for field"""
    return ResponseConfig.field_labels.get(field_name, field_name.replace('_', ' ').title())

def sort_fields_by_priority(fields):
    """Sort fields by priority order for better display"""
    priority_set = set(ResponseConfig.priority_fields)
    priority_fields = [f for f in ResponseConfig.priority_fields if f in fields]
    other_fields = sorted([f for f in fields if f not in priority_set])
    return priority_fields + other_fields

def process_api_response(response, endpoint="", custom_title=""):
    """
    Dynamically process API response and format for Excel display
    
    Args:
        response: Raw API response dictionary
        endpoint: API endpoint URL for schema detection
        custom_title: Custom title for the data
    
    Returns:
        2D array formatted for Excel display
    """
    if "error" in response:
        return format_error(response["error"])
    
    # Extract data from response
    data = response.get("data", response)
    if not data:
        return [["No data received"]]
    
    # Detect endpoint type for formatting hints
    endpoint_type = detect_endpoint_type(endpoint)
    schema = ResponseConfig.endpoint_schemas.get(endpoint_type, {})
    
    # Determine format type
    format_type = schema.get('format', ResponseConfig.preferred_format)
    if format_type == 'auto':
        format_type = 'table' if isinstance(data, list) else 'key_value'
    
    # Handle list vs dict format ambiguity
    if isinstance(data, list) and len(data) == 1 and isinstance(data[0], dict):
        # Single item list - treat as key-value if schema suggests it
        if format_type == 'key_value':
            data = data[0]
    
    # Process based on determined format
    if format_type == 'key_value':
        return format_key_value_data(data, endpoint_type, custom_title)
    elif format_type == 'table':
        return format_table_data(data, endpoint_type, schema)
    else:
        # Fallback to enhanced format_for_excel
        return format_for_excel(data)

def format_key_value_data(data, endpoint_type="", custom_title=""):
    """Format data as key-value pairs with intelligent ordering"""
    if not isinstance(data, dict):
        if isinstance(data, list) and data and isinstance(data[0], dict):
            data = data[0]  # Take first item if it's a single-item list
        else:
            return [["Invalid data format for key-value display"]]
    
    # Create title
    title = custom_title
    if not title:
        schema = ResponseConfig.endpoint_schemas.get(endpoint_type, {})
        if 'title' in schema:
            title = schema['title']
        elif 'title_field' in schema and schema['title_field'] in data:
            symbol = data.get(schema['title_field'], '')
            exchange = data.get('exchange', '')
            title = f"{symbol} ({exchange})" if exchange else str(symbol)
        else:
            title = endpoint_type.title() + " Data"
    
    # Sort fields by priority
    fields = sort_fields_by_priority(list(data.keys()))
    
    # Build result
    result = [[title, "Value"]] if title else [["Field", "Value"]]
    
    for field in fields:
        label = get_display_label(field)
        value = smart_format_value(field, data[field])
        result.append([label, value])
    
    return result

def format_table_data(data, endpoint_type="", schema=None):
    """Format data as a table with intelligent column ordering"""
    if not isinstance(data, list):
        return [["Expected list data for table format"]]
    
    if not data:
        return [["No data available"]]
    
    if not isinstance(data[0], dict):
        # Simple list - convert to single column
        return [["Items"]] + [[str(item)] for item in data]
    
    # Get all unique fields from all records
    all_fields = set()
    for item in data:
        if isinstance(item, dict):
            all_fields.update(item.keys())
    
    # Sort fields by priority
    ordered_fields = sort_fields_by_priority(list(all_fields))
    
    # Create headers with display labels
    headers = [get_display_label(field) for field in ordered_fields]
    result = [headers]
    
    # Process each row
    for item in data:
        row = []
        for field in ordered_fields:
            value = item.get(field, "")
            formatted_value = smart_format_value(field, value)
            row.append(formatted_value)
        result.append(row)
    
    # Sort by timestamp if specified in schema
    if schema and 'sort_by' in schema:
        sort_field = schema['sort_by']
        if sort_field in ordered_fields:
            field_index = ordered_fields.index(sort_field)
            # Sort data rows (skip header)
            result[1:] = sorted(result[1:], key=lambda x: x[field_index], reverse=True)
    
    return result

def format_for_excel(data, headers=None):
    """Enhanced Excel formatter with fallback support"""
    if isinstance(data, dict):
        # Convert dict to 2D array (key-value pairs)
        if headers:
            result = [headers]
        else:
            result = []
        for key, value in data.items():
            formatted_value = smart_format_value(key, value)
            result.append([str(key), formatted_value])
        return result
    
    elif isinstance(data, list) and data:
        if isinstance(data[0], dict):
            # Use the new table formatter
            return format_table_data(data)
        else:
            # List of simple values
            return [[str(item)] for item in data]
    
    elif PANDAS_AVAILABLE and pd and isinstance(data, pd.DataFrame):
        # Pandas DataFrame (only if pandas is available)
        result = [data.columns.tolist()]
        result.extend(data.values.tolist())
        return result
    
    else:
        # Single value
        return [[str(data)]]

def validate_api_key():
    """Check if API key is configured"""
    return bool(OpenAlgoConfig.api_key and OpenAlgoConfig.api_key.strip())

def format_error(message):
    """Return error in Excel-compatible format"""
    return [[f"Error: {message}"]]

# xlwings Lite Implementation - Direct HTTP Requests
@func
def test_xlwings():
    """Test if xlwings Lite is working properly"""
    return "xlwings Lite is working! ✓"

@func
def get_status():
    """Get current system status"""
    return [
        ["xlwings Lite", "✓ Working"],
        ["API Key", "✓ Set" if OpenAlgoConfig.api_key else "✗ Not Set"],
        ["HTTP Method", "Direct API Calls"],
        ["OpenAlgo Host", OpenAlgoConfig.host_url],
        ["API Version", OpenAlgoConfig.version],
        ["Requests Made", str(DebugLog.request_count)]
    ]

@func
def oa_debug_last_request():
    """Get details of the last HTTP request made"""
    if not DebugLog.last_request:
        return [["No requests made yet"]]
    
    req = DebugLog.last_request
    result = [
        ["Property", "Value"],
        ["Request ID", str(req["request_id"])],
        ["Timestamp", req["timestamp"]],
        ["Endpoint", req["endpoint"]],
        ["API Key", "***" + req["payload"]["apikey"][-4:] if "apikey" in req["payload"] else "Not Found"]
    ]
    
    # Add other payload parameters (excluding API key)
    for key, value in req["payload"].items():
        if key != "apikey":
            result.append([f"Param: {key}", str(value)])
    
    return result

@func
def oa_debug_last_response():
    """Get details of the last HTTP response received"""
    if not DebugLog.last_response:
        return [["No responses received yet"]]
    
    resp = DebugLog.last_response
    result = [
        ["Property", "Value"],
        ["Request ID", str(resp["request_id"])],
        ["Timestamp", resp["timestamp"]]
    ]
    
    if "status_code" in resp:
        result.append(["Status Code", str(resp["status_code"])])
    
    if "error" in resp:
        result.append(["Error", resp["error"]])
    elif "data" in resp:
        result.append(["Response Type", "Success"])
        # Show first few keys of response data
        if isinstance(resp["data"], dict):
            result.append(["Response Keys", ", ".join(list(resp["data"].keys())[:5])])
            if "status" in resp["data"]:
                result.append(["API Status", str(resp["data"]["status"])])
            if "message" in resp["data"]:
                result.append(["API Message", str(resp["data"]["message"])[:100]])
    
    return result

@func
def oa_debug_full_log():
    """Get a combined view of the last request and response"""
    if not DebugLog.last_request and not DebugLog.last_response:
        return [["No API calls made yet"]]
    
    result = [["Debug Log", "Details"]]
    
    if DebugLog.last_request:
        req = DebugLog.last_request
        result.extend([
            ["", ""],
            ["REQUEST INFO", ""],
            ["Request ID", str(req["request_id"])],
            ["Timestamp", req["timestamp"]],
            ["Endpoint", req["endpoint"]],
            ["Method", "POST"],
            ["Content-Type", "application/json"]
        ])
        
        # Add payload details
        for key, value in req["payload"].items():
            if key == "apikey":
                result.append([f"Payload: {key}", "***" + str(value)[-4:]])
            else:
                result.append([f"Payload: {key}", str(value)])
    
    if DebugLog.last_response:
        resp = DebugLog.last_response
        result.extend([
            ["", ""],
            ["RESPONSE INFO", ""],
            ["Response ID", str(resp["request_id"])],
            ["Timestamp", resp["timestamp"]]
        ])
        
        if "status_code" in resp:
            result.append(["HTTP Status", str(resp["status_code"])])
        
        if "error" in resp:
            result.append(["Error", resp["error"]])
        elif "data" in resp:
            result.append(["Status", "Success"])
            if isinstance(resp["data"], dict):
                # Show key API response fields
                for key in ["status", "message", "orderid"]:
                    if key in resp["data"]:
                        result.append([f"API {key}", str(resp["data"][key])])
    
    return result

# Configuration Functions
@func
def oa_api(api_key, version="v1", host_url="http://127.0.0.1:5000"):
    """Set the OpenAlgo API Key, API Version, and Host URL globally.
    
    Args:
        api_key: API Key for authentication (required)
        version: API Version (default: v1)
        host_url: Base API URL (default: http://127.0.0.1:5000)
    
    Returns:
        Configuration confirmation message
    """
    if not api_key or not api_key.strip():
        return "Error: API Key is required."
    
    OpenAlgoConfig.api_key = str(api_key).strip()
    OpenAlgoConfig.version = str(version)
    OpenAlgoConfig.host_url = str(host_url).rstrip('/')
    
    return f"Configuration updated: API Key Set, Version = {OpenAlgoConfig.version}, Host = {OpenAlgoConfig.host_url}"

@func
def oa_get_config():
    """Get current OpenAlgo configuration"""
    api_key_display = "***" + OpenAlgoConfig.api_key[-4:] if len(OpenAlgoConfig.api_key) > 4 else "Not Set"
    
    return [
        ["Configuration", "Value"],
        ["API Key", api_key_display],
        ["Version", OpenAlgoConfig.version],
        ["Host URL", OpenAlgoConfig.host_url],
        ["Response Format", ResponseConfig.preferred_format],
        ["Status", "Ready for dynamic API calls"]
    ]

@func
def oa_set_format(format_type="auto"):
    """Set preferred response format for all functions
    
    Args:
        format_type: Format preference ('auto', 'table', 'key_value')
    
    Returns:
        Confirmation message
    """
    valid_formats = ["auto", "table", "key_value"]
    if format_type not in valid_formats:
        return f"Error: format_type must be one of {valid_formats}"
    
    ResponseConfig.preferred_format = format_type
    return f"Response format set to: {format_type}"

@func
def oa_response_info():
    """Get information about the dynamic response system"""
    return [
        ["Feature", "Description"],
        ["Auto Format Detection", "Automatically chooses best display format"],
        ["Smart Field Ordering", "Prioritizes important fields first"],
        ["Price Formatting", "Formats currency values with 2 decimals"],
        ["Timestamp Conversion", "Converts Unix timestamps to readable dates"],
        ["Percentage Formatting", "Adds % suffix to percentage fields"],
        ["Field Labels", "Uses user-friendly column names"],
        ["Schema Learning", "Adapts to different API response patterns"],
        ["List/Dict Handling", "Handles response format inconsistencies"],
        ["Error Processing", "Provides clear error messages"]
    ]

# Market Data Functions
@func
def oa_quotes(symbol, exchange):
    """Get real-time quotes from OpenAlgo API"""
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/quotes"
    payload = {
        "apikey": OpenAlgoConfig.api_key,
        "symbol": str(symbol),
        "exchange": str(exchange)
    }
    
    response = post_request(endpoint, payload)
    
    # Use dynamic response processor with custom title
    custom_title = f"{symbol} ({exchange})"
    return process_api_response(response, endpoint, custom_title)

@func
def oa_depth(symbol, exchange):
    """Get market depth from OpenAlgo API"""
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/depth"
    payload = {
        "apikey": OpenAlgoConfig.api_key,
        "symbol": str(symbol),
        "exchange": str(exchange)
    }
    
    response = post_request(endpoint, payload)
    if "error" in response:
        return format_error(response["error"])
    
    data = response.get("data", {})
    if not data:
        return format_error("No depth data received")
    
    # Special handling for depth data (asks/bids structure)
    asks = data.get("asks", [])
    bids = data.get("bids", [])
    
    result = [["Ask Price", "Ask Qty", "Bid Price", "Bid Qty"]]
    
    max_depth = max(len(asks), len(bids))
    for i in range(max_depth):
        ask_price = smart_format_value("price", asks[i]["price"]) if i < len(asks) else ""
        ask_qty = asks[i]["quantity"] if i < len(asks) else ""
        bid_price = smart_format_value("price", bids[i]["price"]) if i < len(bids) else ""
        bid_qty = bids[i]["quantity"] if i < len(bids) else ""
        
        result.append([ask_price, str(ask_qty), bid_price, str(bid_qty)])
    
    return result

@func
def oa_intervals():
    """Get available time intervals from OpenAlgo API"""
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/intervals"
    payload = {"apikey": OpenAlgoConfig.api_key}
    
    response = post_request(endpoint, payload)
    
    # Try dynamic processing first
    result = process_api_response(response, endpoint)
    
    # If no data from API, return default intervals
    if result == [["No data received"]]:
        return [
            ["Category", "Interval"],
            ["Minutes", "1m"],
            ["Minutes", "5m"],
            ["Minutes", "15m"],
            ["Minutes", "30m"],
            ["Hours", "1h"],
            ["Hours", "4h"],
            ["Daily", "1d"],
            ["Weekly", "1w"],
            ["Monthly", "1M"]
        ]
    
    return result

# Account Management Functions
@func
def oa_funds():
    """Get account funds from OpenAlgo API"""
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/funds"
    payload = {"apikey": OpenAlgoConfig.api_key}
    
    response = post_request(endpoint, payload)
    
    # Use dynamic response processor
    return process_api_response(response, endpoint)

@func
def oa_orderbook():
    """Get order book from OpenAlgo API"""
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/orderbook"
    payload = {"apikey": OpenAlgoConfig.api_key}
    
    response = post_request(endpoint, payload)
    
    # Use dynamic response processor
    return process_api_response(response, endpoint)

@func
def oa_tradebook():
    """Get trade book from OpenAlgo API"""
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/tradebook"
    payload = {"apikey": OpenAlgoConfig.api_key}
    
    response = post_request(endpoint, payload)
    
    # Use dynamic response processor
    return process_api_response(response, endpoint)

@func
def oa_positionbook():
    """Get position book from OpenAlgo API"""
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/positionbook"
    payload = {"apikey": OpenAlgoConfig.api_key}
    
    response = post_request(endpoint, payload)
    
    # Use dynamic response processor
    return process_api_response(response, endpoint)

@func
def oa_holdings():
    """Get holdings from OpenAlgo API"""
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/holdings"
    payload = {"apikey": OpenAlgoConfig.api_key}
    
    response = post_request(endpoint, payload)
    
    # Use dynamic response processor
    return process_api_response(response, endpoint)

# Order Management Functions
def handle_optional_param(param, default="0"):
    """Handle Excel optional parameters - convert None to default"""
    if param is None or param == "":
        return default
    return str(param)

@func
def oa_placeorder(strategy, symbol, action, exchange, pricetype, product, quantity, price=0, trigger_price=0, disclosed_quantity=0):
    """Place an order via OpenAlgo API - CAUTION: REAL ORDERS!"""
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/placeorder"
    payload = {
        "apikey": OpenAlgoConfig.api_key,
        "strategy": str(strategy),
        "symbol": str(symbol),
        "action": str(action),
        "exchange": str(exchange),
        "pricetype": str(pricetype),
        "product": str(product),
        "quantity": handle_optional_param(quantity, "0"),
        "price": handle_optional_param(price, "0"),
        "trigger_price": handle_optional_param(trigger_price, "0"),
        "disclosed_quantity": handle_optional_param(disclosed_quantity, "0")
    }
    
    response = post_request(endpoint, payload)
    if "error" in response:
        return format_error(response["error"])
    
    order_id = response.get("orderid", "Unknown")
    return [["⚠️ ORDER PLACED", "Order ID"], ["Result", str(order_id)]]

@func
def oa_modifyorder(strategy, order_id, symbol, action, exchange, quantity, pricetype="MARKET", product="MIS", price=0, trigger_price=0, disclosed_quantity=0):
    """Modify an existing order"""
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/modifyorder"
    payload = {
        "apikey": OpenAlgoConfig.api_key,
        "strategy": str(strategy),
        "orderid": str(order_id),
        "symbol": str(symbol),
        "action": str(action),
        "exchange": str(exchange),
        "quantity": handle_optional_param(quantity, "0"),
        "pricetype": str(pricetype),
        "product": str(product),
        "price": handle_optional_param(price, "0"),
        "trigger_price": handle_optional_param(trigger_price, "0"),
        "disclosed_quantity": handle_optional_param(disclosed_quantity, "0")
    }
    
    response = post_request(endpoint, payload)
    if "error" in response:
        return format_error(response["error"])
    
    status = response.get("status", "Unknown")
    message = response.get("message", "Order modification request sent")
    return [["Status", str(status)], ["Message", str(message)]]

@func
def oa_cancelorder(strategy, order_id):
    """Cancel a specific order"""
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/cancelorder"
    payload = {
        "apikey": OpenAlgoConfig.api_key,
        "strategy": str(strategy),
        "orderid": str(order_id)
    }
    
    response = post_request(endpoint, payload)
    if "error" in response:
        return format_error(response["error"])
    
    status = response.get("status", "Unknown")
    message = response.get("message", "Order cancellation request sent")
    return [["Status", str(status)], ["Message", str(message)]]

@func
def oa_orderstatus(strategy, order_id):
    """Get order status and details"""
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/orderstatus"
    payload = {
        "apikey": OpenAlgoConfig.api_key,
        "strategy": str(strategy),
        "orderid": str(order_id)
    }
    
    response = post_request(endpoint, payload)
    if "error" in response:
        return format_error(response["error"])
    
    data = response.get("data", {})
    if not data:
        return format_error("No order status data found")
    
    # Convert order details to key-value format
    result = []
    for key, value in data.items():
        # Handle timestamp conversion
        if key.lower() in ['timestamp', 'date', 'time'] and isinstance(value, (int, float)):
            try:
                dt = datetime.fromtimestamp(value)
                value = dt.strftime('%Y-%m-%d %H:%M:%S')
            except (ValueError, TypeError, OSError):
                pass
        result.append([str(key), str(value)])
    
    return result

# Utility Functions
@func
def oa_all_functions():
    """List all available OpenAlgo functions with new dynamic features"""
    return [
        ["Category", "Function", "Description"],
        ["Setup", "oa_api(api_key, version, host_url)", "Set API configuration"],
        ["Setup", "oa_get_config()", "View current configuration"],
        ["Setup", "oa_set_format(format_type)", "🆕 Set response format preference"],
        ["Setup", "oa_response_info()", "🆕 Learn about dynamic response features"],
        ["Setup", "get_status()", "Check system status"],
        ["Debug", "oa_debug_last_request()", "Show last HTTP request details"],
        ["Debug", "oa_debug_last_response()", "Show last HTTP response details"],
        ["Debug", "oa_debug_full_log()", "Show complete request/response log"],
        ["Market Data", "oa_quotes(symbol, exchange)", "🔄 Get real-time quotes - AUTO FORMAT"],
        ["Market Data", "oa_depth(symbol, exchange)", "Get market depth"],
        ["Market Data", "oa_history(symbol, exchange, interval, start, end)", "Get historical data"],
        ["Market Data", "oa_intervals()", "🔄 Get available intervals - AUTO FORMAT"],
        ["Account", "oa_funds()", "🔄 Get account funds - AUTO FORMAT"],
        ["Account", "oa_orderbook()", "🔄 Get order book - AUTO FORMAT"],
        ["Account", "oa_tradebook()", "🔄 Get trade book - AUTO FORMAT"],
        ["Account", "oa_positionbook()", "🔄 Get position book - AUTO FORMAT"],
        ["Account", "oa_holdings()", "🔄 Get holdings - AUTO FORMAT"],
        ["Orders", "oa_placeorder(...)", "Place order"],
        ["Orders", "oa_modifyorder(...)", "Modify order"],
        ["Orders", "oa_cancelorder(strategy, order_id)", "Cancel order"],
        ["Orders", "oa_orderstatus(strategy, order_id)", "Get order status"],
        ["Help", "oa_all_functions()", "This enhanced function list"],
        ["Help", "oa_test_connection()", "Test API connection"],
        ["", "", ""],
        ["🆕 NEW FEATURES", "", ""],
        ["Dynamic Formatting", "All functions auto-adapt", "Handles list/dict format changes"],
        ["Smart Field Ordering", "Important fields first", "Symbol, price, quantity prioritized"],
        ["Price Formatting", "Auto currency format", "Prices show as 123.45"],
        ["Timestamp Conversion", "Readable dates", "Unix timestamps → 2024-06-22 14:30:00"],
        ["Field Labels", "User-friendly names", "ltp → Last Trade Price"],
        ["Error Handling", "Clear error messages", "Better validation and feedback"]
    ]

@func
def oa_test_connection():
    """Test connection to OpenAlgo API"""
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    try:
        endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/funds"
        payload = {"apikey": OpenAlgoConfig.api_key}
        
        response = post_request(endpoint, payload)
        
        if "error" in response:
            return [
                ["Connection Test", "FAILED"],
                ["Error", response["error"]],
                ["Host", OpenAlgoConfig.host_url]
            ]
        else:
            return [
                ["Connection Test", "SUCCESS"],
                ["Host", OpenAlgoConfig.host_url],
                ["Version", OpenAlgoConfig.version]
            ]
            
    except Exception as e:
        return [
            ["Connection Test", "FAILED"],
            ["Error", str(e)]
        ]

@func
def oa_history(symbol, exchange, interval, start_date, end_date):
    """Get historical data from OpenAlgo API"""
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/history"
    payload = {
        "apikey": OpenAlgoConfig.api_key,
        "symbol": str(symbol),
        "exchange": str(exchange),
        "interval": str(interval),
        "start_date": str(start_date),
        "end_date": str(end_date)
    }
    
    response = post_request(endpoint, payload)
    
    # For historical data, we need special handling to include symbol and split timestamp
    if "error" in response:
        return format_error(response["error"])
    
    data = response.get("data", [])
    if not data:
        return format_error("No historical data found")
    
    # Historical data needs special formatting with symbol and split timestamp
    result = [["Ticker", "Date", "Time", "Open", "High", "Low", "Close", "Volume"]]
    
    for item in data:
        # Convert timestamp to IST date and time
        try:
            timestamp = item.get("timestamp", 0)
            dt = datetime.fromtimestamp(timestamp)
            date_str = dt.strftime('%Y-%m-%d')
            time_str = dt.strftime('%H:%M:%S')
        except (ValueError, TypeError, OSError):
            date_str = "N/A"
            time_str = "N/A"
        
        result.append([
            str(symbol),
            date_str,
            time_str,
            smart_format_value("open", item.get("open", "")),
            smart_format_value("high", item.get("high", "")),
            smart_format_value("low", item.get("low", "")),
            smart_format_value("close", item.get("close", "")),
            str(item.get("volume", ""))
        ])
    
    return result


# Order Management Functions (still using manual formatting for complex order responses)
def handle_optional_param(param, default="0"):
    """Handle Excel optional parameters - convert None to default"""
    if param is None or param == "":
        return default
    return str(param)

