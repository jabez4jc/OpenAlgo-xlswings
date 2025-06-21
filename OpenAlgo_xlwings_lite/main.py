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
from typing import List, Any, Optional, Union
import pandas as pd

# Try to import pyodide_http for WebAssembly compatibility
try:
    import pyodide_http
    pyodide_http.patch_all()
except ImportError:
    # Running in standard Python environment
    pass

# Global Configuration Storage
class OpenAlgoConfig:
    """Global configuration for OpenAlgo API"""
    api_key: str = ""
    version: str = "v1"
    host_url: str = "http://127.0.0.1:5000"

# Utility Functions
def post_request(endpoint: str, payload: dict) -> dict:
    """Make HTTP POST request using urllib (Pyodide compatible)"""
    try:
        data = json.dumps(payload).encode('utf-8')
        headers = {
            'Content-Type': 'application/json',
            'Accept': 'application/json'
        }
        
        request = urllib.request.Request(endpoint, data=data, headers=headers)
        response = urllib.request.urlopen(request, timeout=30)
        
        return json.loads(response.read().decode('utf-8'))
    except urllib.error.HTTPError as e:
        return {"error": f"HTTP Error {e.code}: {e.reason}"}
    except urllib.error.URLError as e:
        return {"error": f"URL Error: {e.reason}"}
    except json.JSONDecodeError as e:
        return {"error": f"JSON Decode Error: {str(e)}"}
    except Exception as e:
        return {"error": str(e)}

def format_for_excel(data: Any, headers: Optional[List[str]] = None) -> List[List[Any]]:
    """Convert various data types to Excel-friendly 2D arrays"""
    if isinstance(data, dict):
        # Convert dict to 2D array (key-value pairs)
        if headers:
            result = [headers]
        else:
            result = []
        for key, value in data.items():
            result.append([str(key), str(value)])
        return result
    
    elif isinstance(data, list) and data:
        if isinstance(data[0], dict):
            # List of dictionaries - create table with headers
            if not data:
                return [["No data available"]]
            
            headers = list(data[0].keys())
            rows = [headers]
            for item in data:
                row = []
                for header in headers:
                    value = item.get(header, "")
                    # Handle special formatting for timestamps
                    if header.lower() in ['timestamp', 'date', 'time'] and isinstance(value, (int, float)):
                        try:
                            # Convert Unix timestamp to IST
                            dt = datetime.fromtimestamp(value)
                            value = dt.strftime('%Y-%m-%d %H:%M:%S')
                        except:
                            pass
                    row.append(str(value))
                rows.append(row)
            return rows
        else:
            # List of simple values
            return [[str(item)] for item in data]
    
    elif isinstance(data, pd.DataFrame):
        # Pandas DataFrame
        result = [data.columns.tolist()]
        result.extend(data.values.tolist())
        return result
    
    else:
        # Single value
        return [[str(data)]]

def validate_api_key() -> bool:
    """Check if API key is configured"""
    return bool(OpenAlgoConfig.api_key and OpenAlgoConfig.api_key.strip())

def format_error(message: str) -> List[List[str]]:
    """Return error in Excel-compatible format"""
    return [[f"Error: {message}"]]

# Configuration Function
@func
def oa_api(api_key: str, version: str = "v1", host_url: str = "http://127.0.0.1:5000") -> str:
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

# Market Data Functions
@func
def oa_quotes(symbol: str, exchange: str) -> List[List[str]]:
    """Retrieve market quotes for a given symbol.
    
    Args:
        symbol: Trading symbol
        exchange: Exchange code (NSE/BSE/NFO/MCX)
    
    Returns:
        2D array with quote data
    """
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/quotes"
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
        return format_error("No quote data found")
    
    # Format as key-value table with header
    result = [[f"{symbol} ({exchange})", "Value"]]
    for key, value in data.items():
        result.append([str(key), str(value)])
    
    return result

@func
def oa_depth(symbol: str, exchange: str) -> List[List[str]]:
    """Retrieve market depth for a given symbol.
    
    Args:
        symbol: Trading symbol
        exchange: Exchange code (NSE/BSE/NFO/MCX)
    
    Returns:
        2D array with market depth data
    """
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
        return format_error("No depth data found")
    
    # Format asks and bids data
    asks = data.get("asks", [])
    bids = data.get("bids", [])
    
    # Create table with Ask Price, Ask Qty, Bid Price, Bid Qty columns
    result = [["Ask Price", "Ask Qty", "Bid Price", "Bid Qty"]]
    
    max_depth = max(len(asks), len(bids))
    for i in range(max_depth):
        ask_price = asks[i]["price"] if i < len(asks) else ""
        ask_qty = asks[i]["quantity"] if i < len(asks) else ""
        bid_price = bids[i]["price"] if i < len(bids) else ""
        bid_qty = bids[i]["quantity"] if i < len(bids) else ""
        
        result.append([str(ask_price), str(ask_qty), str(bid_price), str(bid_qty)])
    
    return result

@func
def oa_history(symbol: str, exchange: str, interval: str, start_date: str, end_date: str) -> List[List[str]]:
    """Retrieve historical data for a given symbol.
    
    Args:
        symbol: Trading symbol
        exchange: Exchange code (NSE/BSE/NFO/MCX)
        interval: Time interval (1m, 5m, 15m, 1h, 1d)
        start_date: Start date (YYYY-MM-DD)
        end_date: End date (YYYY-MM-DD)
    
    Returns:
        2D array with historical OHLCV data
    """
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
    if "error" in response:
        return format_error(response["error"])
    
    data = response.get("data", [])
    if not data:
        return format_error("No historical data found")
    
    # Format with headers: Ticker, Date, Time, Open, High, Low, Close, Volume
    result = [["Ticker", "Date", "Time", "Open", "High", "Low", "Close", "Volume"]]
    
    for item in data:
        # Convert timestamp to IST date and time
        try:
            timestamp = item.get("timestamp", 0)
            dt = datetime.fromtimestamp(timestamp)
            date_str = dt.strftime('%Y-%m-%d')
            time_str = dt.strftime('%H:%M:%S')
        except:
            date_str = "N/A"
            time_str = "N/A"
        
        result.append([
            str(symbol),
            date_str,
            time_str,
            str(item.get("open", "")),
            str(item.get("high", "")),
            str(item.get("low", "")),
            str(item.get("close", "")),
            str(item.get("volume", ""))
        ])
    
    return result

@func
def oa_intervals() -> List[List[str]]:
    """Retrieve available time intervals.
    
    Returns:
        2D array with available intervals categorized
    """
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/intervals"
    payload = {"apikey": OpenAlgoConfig.api_key}
    
    response = post_request(endpoint, payload)
    if "error" in response:
        return format_error(response["error"])
    
    data = response.get("data", {})
    if not data:
        # Return default intervals if API doesn't provide them
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
    
    return format_for_excel(data, ["Category", "Interval"])

# Account Management Functions
@func
def oa_funds() -> List[List[str]]:
    """Retrieve funds from OpenAlgo API.
    
    Returns:
        2D array with account funds information
    """
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/funds"
    payload = {"apikey": OpenAlgoConfig.api_key}
    
    response = post_request(endpoint, payload)
    if "error" in response:
        return format_error(response["error"])
    
    data = response.get("data", {})
    if not data:
        return format_error("No funds data found")
    
    # Convert funds data to key-value format
    result = []
    for key, value in data.items():
        result.append([str(key), str(value)])
    
    return result

@func
def oa_orderbook() -> List[List[str]]:
    """Retrieve the order book from OpenAlgo API.
    
    Returns:
        2D array with order book data (11 columns)
    """
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/orderbook"
    payload = {"apikey": OpenAlgoConfig.api_key}
    
    response = post_request(endpoint, payload)
    if "error" in response:
        return format_error(response["error"])
    
    data = response.get("data", [])
    if not data:
        return [["No orders found"]]
    
    # Headers matching C# implementation
    headers = [
        "Symbol", "Action", "Exchange", "Quantity", "Order Status", 
        "Order ID", "Price", "Price Type", "Trigger Price", "Product", "Timestamp"
    ]
    
    result = [headers]
    for order in data:
        # Convert timestamp if present
        timestamp = order.get("timestamp", "")
        if timestamp and isinstance(timestamp, (int, float)):
            try:
                dt = datetime.fromtimestamp(timestamp)
                timestamp = dt.strftime('%Y-%m-%d %H:%M:%S')
            except:
                pass
        
        row = [
            str(order.get("symbol", "")),
            str(order.get("action", "")),
            str(order.get("exchange", "")),
            str(order.get("quantity", "")),
            str(order.get("status", "")),
            str(order.get("orderid", "")),
            str(order.get("price", "")),
            str(order.get("pricetype", "")),
            str(order.get("trigger_price", "")),
            str(order.get("product", "")),
            str(timestamp)
        ]
        result.append(row)
    
    return result

@func
def oa_tradebook() -> List[List[str]]:
    """Retrieve the trade book from OpenAlgo API.
    
    Returns:
        2D array with trade book data (9 columns)
    """
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/tradebook"
    payload = {"apikey": OpenAlgoConfig.api_key}
    
    response = post_request(endpoint, payload)
    if "error" in response:
        return format_error(response["error"])
    
    data = response.get("data", [])
    if not data:
        return [["No trades found"]]
    
    # Headers matching C# implementation
    headers = [
        "Symbol", "Exchange", "Action", "Quantity", "Product", 
        "Timestamp", "Trade Value", "Average Price", "Order ID"
    ]
    
    result = [headers]
    for trade in data:
        # Convert timestamp if present
        timestamp = trade.get("timestamp", "")
        if timestamp and isinstance(timestamp, (int, float)):
            try:
                dt = datetime.fromtimestamp(timestamp)
                timestamp = dt.strftime('%Y-%m-%d %H:%M:%S')
            except:
                pass
        
        row = [
            str(trade.get("symbol", "")),
            str(trade.get("exchange", "")),
            str(trade.get("action", "")),
            str(trade.get("quantity", "")),
            str(trade.get("product", "")),
            str(timestamp),
            str(trade.get("trade_value", "")),
            str(trade.get("average_price", "")),
            str(trade.get("orderid", ""))
        ]
        result.append(row)
    
    return result

@func
def oa_positionbook() -> List[List[str]]:
    """Retrieve the position book from OpenAlgo API.
    
    Returns:
        2D array with position book data (5 columns)
    """
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/positionbook"
    payload = {"apikey": OpenAlgoConfig.api_key}
    
    response = post_request(endpoint, payload)
    if "error" in response:
        return format_error(response["error"])
    
    data = response.get("data", [])
    if not data:
        return [["No positions found"]]
    
    # Headers matching C# implementation
    headers = ["Symbol", "Exchange", "Quantity", "Product", "Average Price"]
    
    result = [headers]
    for position in data:
        row = [
            str(position.get("symbol", "")),
            str(position.get("exchange", "")),
            str(position.get("quantity", "")),
            str(position.get("product", "")),
            str(position.get("average_price", ""))
        ]
        result.append(row)
    
    return result

@func
def oa_holdings() -> List[List[str]]:
    """Retrieve holdings from OpenAlgo API.
    
    Returns:
        2D array with holdings data (6 columns)
    """
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/holdings"
    payload = {"apikey": OpenAlgoConfig.api_key}
    
    response = post_request(endpoint, payload)
    if "error" in response:
        return format_error(response["error"])
    
    data = response.get("data", [])
    if not data:
        return [["No holdings found"]]
    
    # Headers matching C# implementation
    headers = ["Symbol", "Exchange", "Quantity", "Product", "PnL", "PnL Percent"]
    
    result = [headers]
    for holding in data:
        row = [
            str(holding.get("symbol", "")),
            str(holding.get("exchange", "")),
            str(holding.get("quantity", "")),
            str(holding.get("product", "")),
            str(holding.get("pnl", "")),
            str(holding.get("pnl_percent", ""))
        ]
        result.append(row)
    
    return result

# Order Management Functions
def handle_optional_param(param: Any, default: str = "0") -> str:
    """Handle Excel optional parameters - convert None to default"""
    if param is None or param == "":
        return default
    return str(param)

@func
def oa_placeorder(
    strategy: str,
    symbol: str, 
    action: str,
    exchange: str,
    pricetype: str,
    product: str,
    quantity: Union[str, int, float] = 0,
    price: Union[str, int, float] = 0,
    trigger_price: Union[str, int, float] = 0,
    disclosed_quantity: Union[str, int, float] = 0
) -> List[List[str]]:
    """Place an order via OpenAlgo API.
    
    Args:
        strategy: Trading strategy name
        symbol: Trading symbol
        action: Order action (BUY/SELL)
        exchange: Exchange code (NSE/BSE/NFO/MCX)
        pricetype: Price type (MARKET/LIMIT)
        product: Product type (MIS/CNC/NRML)
        quantity: Order quantity
        price: Order price (optional)
        trigger_price: Trigger price (optional)
        disclosed_quantity: Disclosed quantity (optional)
    
    Returns:
        2D array with Order ID or error message
    """
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    # Convert all parameters to strings as required by OpenAlgo API
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
    return [["Order ID", str(order_id)]]

@func
def oa_modifyorder(
    strategy: str,
    order_id: str,
    symbol: str,
    action: str,
    exchange: str,
    quantity: Union[str, int, float] = 0,
    pricetype: str = "MARKET",
    product: str = "MIS",
    price: Union[str, int, float] = 0,
    trigger_price: Union[str, int, float] = 0,
    disclosed_quantity: Union[str, int, float] = 0
) -> List[List[str]]:
    """Modify an existing order.
    
    Args:
        strategy: Trading strategy name
        order_id: Order ID to modify
        symbol: Trading symbol
        action: Order action (BUY/SELL)
        exchange: Exchange code (NSE/BSE/NFO/MCX)
        quantity: New order quantity
        pricetype: Price type (MARKET/LIMIT)
        product: Product type (MIS/CNC/NRML)
        price: New order price (optional)
        trigger_price: New trigger price (optional)
        disclosed_quantity: New disclosed quantity (optional)
    
    Returns:
        2D array with modification status
    """
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
def oa_cancelorder(strategy: str, order_id: str) -> List[List[str]]:
    """Cancel a specific order.
    
    Args:
        strategy: Trading strategy name
        order_id: Order ID to cancel
    
    Returns:
        2D array with cancellation status
    """
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
def oa_orderstatus(strategy: str, order_id: str) -> List[List[str]]:
    """Get order status and details.
    
    Args:
        strategy: Trading strategy name
        order_id: Order ID to check
    
    Returns:
        2D array with order details
    """
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
            except:
                pass
        result.append([str(key), str(value)])
    
    return result

@func
def oa_openposition(strategy: str, symbol: str, exchange: str, product: str) -> List[List[str]]:
    """Get open position details for specific instrument.
    
    Args:
        strategy: Trading strategy name
        symbol: Trading symbol
        exchange: Exchange code (NSE/BSE/NFO/MCX)
        product: Product type (MIS/CNC/NRML)
    
    Returns:
        2D array with position details
    """
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/openposition"
    payload = {
        "apikey": OpenAlgoConfig.api_key,
        "strategy": str(strategy),
        "symbol": str(symbol),
        "exchange": str(exchange),
        "product": str(product)
    }
    
    response = post_request(endpoint, payload)
    if "error" in response:
        return format_error(response["error"])
    
    data = response.get("data", {})
    if not data:
        return format_error("No position data found")
    
    # Convert position details to key-value format
    result = []
    for key, value in data.items():
        result.append([str(key), str(value)])
    
    return result

@func
def oa_closeposition(strategy: str) -> List[List[str]]:
    """Close all open positions for a strategy.
    
    Args:
        strategy: Trading strategy name
    
    Returns:
        2D array with close position results
    """
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/closeposition"
    payload = {
        "apikey": OpenAlgoConfig.api_key,
        "strategy": str(strategy)
    }
    
    response = post_request(endpoint, payload)
    if "error" in response:
        return format_error(response["error"])
    
    # Handle different response formats
    if "data" in response:
        data = response["data"]
        if isinstance(data, list):
            # List of closed positions
            headers = ["Symbol", "Action", "Quantity", "Status", "Message"]
            result = [headers]
            for item in data:
                if isinstance(item, dict):
                    result.append([
                        str(item.get("symbol", "")),
                        str(item.get("action", "")),
                        str(item.get("quantity", "")),
                        str(item.get("status", "")),
                        str(item.get("message", ""))
                    ])
            return result
        else:
            # Simple status response
            return [["Status", str(data)]]
    else:
        status = response.get("status", "Unknown")
        message = response.get("message", "Close positions request sent")
        return [["Status", str(status)], ["Message", str(message)]]

# Advanced Order Functions
@func
def oa_placesmartorder(
    strategy: str,
    symbol: str,
    action: str,
    exchange: str,
    pricetype: str,
    product: str,
    quantity: Union[str, int, float] = 0,
    position_size: Union[str, int, float] = 0,
    price: Union[str, int, float] = 0,
    trigger_price: Union[str, int, float] = 0,
    disclosed_quantity: Union[str, int, float] = 0
) -> List[List[str]]:
    """Place a smart order with position sizing.
    
    Args:
        strategy: Trading strategy name
        symbol: Trading symbol
        action: Order action (BUY/SELL)
        exchange: Exchange code (NSE/BSE/NFO/MCX)
        pricetype: Price type (MARKET/LIMIT)
        product: Product type (MIS/CNC/NRML)
        quantity: Order quantity
        position_size: Position size for smart order
        price: Order price (optional)
        trigger_price: Trigger price (optional)
        disclosed_quantity: Disclosed quantity (optional)
    
    Returns:
        2D array with Order ID or error message
    """
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/placesmartorder"
    payload = {
        "apikey": OpenAlgoConfig.api_key,
        "strategy": str(strategy),
        "symbol": str(symbol),
        "action": str(action),
        "exchange": str(exchange),
        "pricetype": str(pricetype),
        "product": str(product),
        "quantity": handle_optional_param(quantity, "0"),
        "position_size": handle_optional_param(position_size, "0"),
        "price": handle_optional_param(price, "0"),
        "trigger_price": handle_optional_param(trigger_price, "0"),
        "disclosed_quantity": handle_optional_param(disclosed_quantity, "0")
    }
    
    response = post_request(endpoint, payload)
    if "error" in response:
        return format_error(response["error"])
    
    order_id = response.get("orderid", "Unknown")
    return [["Order ID", str(order_id)]]

@func
def oa_basketorder(strategy: str, orders: List[List[Any]]) -> List[List[str]]:
    """Place multiple orders in a basket.
    
    Args:
        strategy: Trading strategy name
        orders: 2D array of order parameters (from Excel range)
    
    Returns:
        2D array with basket order results
    """
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    if not orders or not isinstance(orders, list):
        return format_error("Invalid orders data. Provide a range of order parameters.")
    
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/basketorder"
    
    # Process orders array - assume first row might be headers
    order_list = []
    start_row = 1 if len(orders) > 1 and isinstance(orders[0][0], str) and "symbol" in str(orders[0][0]).lower() else 0
    
    for i in range(start_row, len(orders)):
        row = orders[i]
        if len(row) >= 6:  # Minimum required columns
            order_dict = {
                "symbol": str(row[0]) if len(row) > 0 else "",
                "action": str(row[1]) if len(row) > 1 else "",
                "exchange": str(row[2]) if len(row) > 2 else "",
                "pricetype": str(row[3]) if len(row) > 3 else "MARKET",
                "product": str(row[4]) if len(row) > 4 else "MIS",
                "quantity": str(row[5]) if len(row) > 5 else "0",
                "price": str(row[6]) if len(row) > 6 else "0",
                "trigger_price": str(row[7]) if len(row) > 7 else "0",
                "disclosed_quantity": str(row[8]) if len(row) > 8 else "0"
            }
            order_list.append(order_dict)
    
    if not order_list:
        return format_error("No valid orders found in the provided data")
    
    payload = {
        "apikey": OpenAlgoConfig.api_key,
        "strategy": str(strategy),
        "orders": order_list
    }
    
    response = post_request(endpoint, payload)
    if "error" in response:
        return format_error(response["error"])
    
    # Format results
    data = response.get("data", [])
    if isinstance(data, list):
        headers = ["Symbol", "Order ID", "Status", "Message"]
        result = [headers]
        for item in data:
            if isinstance(item, dict):
                result.append([
                    str(item.get("symbol", "")),
                    str(item.get("orderid", "")),
                    str(item.get("status", "")),
                    str(item.get("message", ""))
                ])
        return result
    else:
        return [["Status", str(data)]]

@func
def oa_splitorder(
    strategy: str,
    symbol: str,
    action: str,
    exchange: str,
    quantity: Union[str, int, float] = 0,
    splitsize: Union[str, int, float] = 0,
    pricetype: str = "MARKET",
    product: str = "MIS",
    price: Union[str, int, float] = 0,
    trigger_price: Union[str, int, float] = 0,
    disclosed_quantity: Union[str, int, float] = 0
) -> List[List[str]]:
    """Place a split order (break large order into smaller chunks).
    
    Args:
        strategy: Trading strategy name
        symbol: Trading symbol
        action: Order action (BUY/SELL)
        exchange: Exchange code (NSE/BSE/NFO/MCX)
        quantity: Total order quantity
        splitsize: Size of each split order
        pricetype: Price type (MARKET/LIMIT)
        product: Product type (MIS/CNC/NRML)
        price: Order price (optional)
        trigger_price: Trigger price (optional)
        disclosed_quantity: Disclosed quantity (optional)
    
    Returns:
        2D array with split order results
    """
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/splitorder"
    payload = {
        "apikey": OpenAlgoConfig.api_key,
        "strategy": str(strategy),
        "symbol": str(symbol),
        "action": str(action),
        "exchange": str(exchange),
        "quantity": handle_optional_param(quantity, "0"),
        "splitsize": handle_optional_param(splitsize, "0"),
        "pricetype": str(pricetype),
        "product": str(product),
        "price": handle_optional_param(price, "0"),
        "trigger_price": handle_optional_param(trigger_price, "0"),
        "disclosed_quantity": handle_optional_param(disclosed_quantity, "0")
    }
    
    response = post_request(endpoint, payload)
    if "error" in response:
        return format_error(response["error"])
    
    # Format results
    data = response.get("data", [])
    if isinstance(data, list):
        headers = ["Order Number", "Order ID", "Quantity", "Status", "Message"]
        result = [headers]
        for i, item in enumerate(data):
            if isinstance(item, dict):
                result.append([
                    str(i + 1),
                    str(item.get("orderid", "")),
                    str(item.get("quantity", "")),
                    str(item.get("status", "")),
                    str(item.get("message", ""))
                ])
        return result
    else:
        return [["Status", str(data)]]
# Automation Scripts and Helper Functions
@script
def refresh_all_quotes(book: xw.Book):
    """Refresh all quote formulas in the active sheet."""
    sheet = book.sheets.active
    cells_updated = 0
    
    try:
        for cell in sheet.used_range:
            if cell.formula and "oa_quotes" in str(cell.formula):
                temp_formula = cell.formula
                cell.clear()
                cell.formula = temp_formula
                cells_updated += 1
        
        sheet["A1"].value = f"Refreshed {cells_updated} quote cells at {datetime.now().strftime('%H:%M:%S')}"
    except Exception as e:
        sheet["A1"].value = f"Error refreshing quotes: {str(e)}"

@script
def setup_dashboard(book: xw.Book):
    """Create a sample OpenAlgo trading dashboard."""
    try:
        try:
            sheet = book.sheets["Dashboard"]
            sheet.clear()
        except:
            sheet = book.sheets.add("Dashboard")
        
        sheet["A1"].value = "OpenAlgo Trading Dashboard"
        sheet["A1"].font.size = 18
        sheet["A1"].font.bold = True
        
        sheet["A3"].value = "Configuration"
        sheet["A3"].font.bold = True
        sheet["A4"].value = "API Key:"
        sheet["B4"].value = '=oa_api("your_api_key_here")'
        
        sheet["A6"].value = "Market Data"
        sheet["A6"].font.bold = True
        sheet["A7"].value = "Symbol"
        sheet["B7"].value = "Exchange"
        sheet["C7"].value = "Quote Data"
        
        sheet["A8"].value = "RELIANCE"
        sheet["B8"].value = "NSE"
        sheet["C8"].value = '=oa_quotes(A8,B8)'
        
        sheet["A11"].value = "Account Information"
        sheet["A11"].font.bold = True
        sheet["A12"].value = "Funds:"
        sheet["B12"].value = '=oa_funds()'
        
        sheet.activate()
        
    except Exception as e:
        book.sheets.active["A1"].value = f"Error creating dashboard: {str(e)}"

# Testing Functions
@func
def oa_test_connection() -> List[List[str]]:
    """Test connection to OpenAlgo API."""
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
def oa_get_config() -> List[List[str]]:
    """Get current OpenAlgo configuration."""
    api_key_display = "***" + OpenAlgoConfig.api_key[-4:] if len(OpenAlgoConfig.api_key) > 4 else "Not Set"
    
    return [
        ["Configuration", "Value"],
        ["API Key", api_key_display],
        ["Version", OpenAlgoConfig.version],
        ["Host URL", OpenAlgoConfig.host_url]
    ]
EOF < /dev/null
