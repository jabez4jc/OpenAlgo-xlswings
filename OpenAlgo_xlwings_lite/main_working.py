"""OpenAlgo Excel Integration via xlwings Lite - Working Version

This version generates curl commands and API request details since direct HTTP
requests don't work in xlwings Lite environment. Use the generated curl commands
or Power Query to access OpenAlgo API data.
"""

import xlwings as xw
from xlwings import func, script
import json
from datetime import datetime

# Global Configuration Storage
class OpenAlgoConfig:
    """Global configuration for OpenAlgo API"""
    api_key = ""
    version = "v1"
    host_url = "https://openalgo.simplifyed.in"

def validate_api_key():
    """Check if API key is configured"""
    return bool(OpenAlgoConfig.api_key and OpenAlgoConfig.api_key.strip())

def format_error(message):
    """Return error in Excel-compatible format"""
    return [[f"Error: {message}"]]

# System Status and Configuration Functions
@func
def test_xlwings():
    """Test if xlwings Lite is working properly"""
    return "xlwings Lite is working! ✓"

@func
def get_status():
    """Get current system status"""
    return [
        ["Component", "Status"],
        ["xlwings Lite", "✓ Working"],
        ["API Key", "✓ Set" if OpenAlgoConfig.api_key else "✗ Not Set"],
        ["HTTP Method", "Manual (curl/Power Query)"],
        ["OpenAlgo Host", OpenAlgoConfig.host_url],
        ["API Version", OpenAlgoConfig.version]
    ]

@func
def oa_api(api_key, version="v1", host_url="https://openalgo.simplifyed.in"):
    """Set the OpenAlgo API Key, API Version, and Host URL globally"""
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
        ["Status", "Ready for manual requests"]
    ]

# Request Generator Functions - Generate curl commands and API details
@func
def oa_funds_request():
    """Generate funds API request details"""
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    payload = {"apikey": OpenAlgoConfig.api_key}
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/funds"
    
    return [
        ["URL", endpoint],
        ["Method", "POST"],
        ["Content-Type", "application/json"],
        ["JSON Body", json.dumps(payload)],
        ["Curl Command", f'curl -X POST "{endpoint}" -H "Content-Type: application/json" -d \'{json.dumps(payload)}\'']
    ]

@func
def oa_quotes_request(symbol, exchange):
    """Generate quotes API request details"""
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    payload = {
        "apikey": OpenAlgoConfig.api_key,
        "symbol": str(symbol),
        "exchange": str(exchange)
    }
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/quotes"
    
    return [
        ["URL", endpoint],
        ["Method", "POST"], 
        ["Content-Type", "application/json"],
        ["JSON Body", json.dumps(payload)],
        ["Symbol", str(symbol)],
        ["Exchange", str(exchange)],
        ["Curl Command", f'curl -X POST "{endpoint}" -H "Content-Type: application/json" -d \'{json.dumps(payload)}\'']
    ]

@func
def oa_depth_request(symbol, exchange):
    """Generate market depth API request details"""
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    payload = {
        "apikey": OpenAlgoConfig.api_key,
        "symbol": str(symbol),
        "exchange": str(exchange)
    }
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/depth"
    
    return [
        ["URL", endpoint],
        ["Method", "POST"],
        ["Content-Type", "application/json"],
        ["JSON Body", json.dumps(payload)],
        ["Symbol", str(symbol)],
        ["Exchange", str(exchange)],
        ["Curl Command", f'curl -X POST "{endpoint}" -H "Content-Type: application/json" -d \'{json.dumps(payload)}\'']
    ]

@func
def oa_history_request(symbol, exchange, interval, start_date, end_date):
    """Generate historical data API request details"""
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    payload = {
        "apikey": OpenAlgoConfig.api_key,
        "symbol": str(symbol),
        "exchange": str(exchange),
        "interval": str(interval),
        "start_date": str(start_date),
        "end_date": str(end_date)
    }
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/history"
    
    return [
        ["URL", endpoint],
        ["Method", "POST"],
        ["Content-Type", "application/json"],
        ["JSON Body", json.dumps(payload)],
        ["Curl Command", f'curl -X POST "{endpoint}" -H "Content-Type: application/json" -d \'{json.dumps(payload)}\'']
    ]

@func
def oa_intervals_request():
    """Generate intervals API request details"""
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    payload = {"apikey": OpenAlgoConfig.api_key}
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/intervals"
    
    return [
        ["URL", endpoint],
        ["Method", "POST"],
        ["Content-Type", "application/json"],
        ["JSON Body", json.dumps(payload)],
        ["Curl Command", f'curl -X POST "{endpoint}" -H "Content-Type: application/json" -d \'{json.dumps(payload)}\'']
    ]

# Account Management Request Generators
@func
def oa_orderbook_request():
    """Generate order book API request details"""
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    payload = {"apikey": OpenAlgoConfig.api_key}
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/orderbook"
    
    return [
        ["URL", endpoint],
        ["Method", "POST"],
        ["Content-Type", "application/json"],
        ["JSON Body", json.dumps(payload)],
        ["Curl Command", f'curl -X POST "{endpoint}" -H "Content-Type: application/json" -d \'{json.dumps(payload)}\'']
    ]

@func
def oa_tradebook_request():
    """Generate trade book API request details"""
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    payload = {"apikey": OpenAlgoConfig.api_key}
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/tradebook"
    
    return [
        ["URL", endpoint],
        ["Method", "POST"],
        ["Content-Type", "application/json"],
        ["JSON Body", json.dumps(payload)],
        ["Curl Command", f'curl -X POST "{endpoint}" -H "Content-Type: application/json" -d \'{json.dumps(payload)}\'']
    ]

@func
def oa_positionbook_request():
    """Generate position book API request details"""
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    payload = {"apikey": OpenAlgoConfig.api_key}
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/positionbook"
    
    return [
        ["URL", endpoint],
        ["Method", "POST"],
        ["Content-Type", "application/json"],
        ["JSON Body", json.dumps(payload)],
        ["Curl Command", f'curl -X POST "{endpoint}" -H "Content-Type: application/json" -d \'{json.dumps(payload)}\'']
    ]

@func
def oa_holdings_request():
    """Generate holdings API request details"""
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    payload = {"apikey": OpenAlgoConfig.api_key}
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/holdings"
    
    return [
        ["URL", endpoint],
        ["Method", "POST"],
        ["Content-Type", "application/json"],
        ["JSON Body", json.dumps(payload)],
        ["Curl Command", f'curl -X POST "{endpoint}" -H "Content-Type: application/json" -d \'{json.dumps(payload)}\'']
    ]

# Order Management Request Generators  
@func
def oa_placeorder_request(strategy, symbol, action, exchange, pricetype, product, quantity, price=0, trigger_price=0, disclosed_quantity=0):
    """Generate place order API request details - CAUTION: REAL ORDERS!"""
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    payload = {
        "apikey": OpenAlgoConfig.api_key,
        "strategy": str(strategy),
        "symbol": str(symbol),
        "action": str(action),
        "exchange": str(exchange),
        "pricetype": str(pricetype),
        "product": str(product),
        "quantity": str(quantity),
        "price": str(price),
        "trigger_price": str(trigger_price),
        "disclosed_quantity": str(disclosed_quantity)
    }
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/placeorder"
    
    return [
        ["⚠️ WARNING", "THIS WILL PLACE A REAL ORDER!"],
        ["URL", endpoint],
        ["Method", "POST"],
        ["Content-Type", "application/json"],
        ["JSON Body", json.dumps(payload)],
        ["Strategy", str(strategy)],
        ["Symbol", str(symbol)],
        ["Action", str(action)],
        ["Quantity", str(quantity)],
        ["Curl Command", f'curl -X POST "{endpoint}" -H "Content-Type: application/json" -d \'{json.dumps(payload)}\'']
    ]

@func
def oa_modifyorder_request(strategy, order_id, symbol, action, exchange, quantity, pricetype="MARKET", product="MIS", price=0, trigger_price=0, disclosed_quantity=0):
    """Generate modify order API request details"""
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    payload = {
        "apikey": OpenAlgoConfig.api_key,
        "strategy": str(strategy),
        "orderid": str(order_id),
        "symbol": str(symbol),
        "action": str(action),
        "exchange": str(exchange),
        "quantity": str(quantity),
        "pricetype": str(pricetype),
        "product": str(product),
        "price": str(price),
        "trigger_price": str(trigger_price),
        "disclosed_quantity": str(disclosed_quantity)
    }
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/modifyorder"
    
    return [
        ["URL", endpoint],
        ["Method", "POST"],
        ["Content-Type", "application/json"],
        ["JSON Body", json.dumps(payload)],
        ["Order ID", str(order_id)],
        ["Curl Command", f'curl -X POST "{endpoint}" -H "Content-Type: application/json" -d \'{json.dumps(payload)}\'']
    ]

@func
def oa_cancelorder_request(strategy, order_id):
    """Generate cancel order API request details"""
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    payload = {
        "apikey": OpenAlgoConfig.api_key,
        "strategy": str(strategy),
        "orderid": str(order_id)
    }
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/cancelorder"
    
    return [
        ["URL", endpoint],
        ["Method", "POST"],
        ["Content-Type", "application/json"],
        ["JSON Body", json.dumps(payload)],
        ["Order ID", str(order_id)],
        ["Curl Command", f'curl -X POST "{endpoint}" -H "Content-Type: application/json" -d \'{json.dumps(payload)}\'']
    ]

@func
def oa_orderstatus_request(strategy, order_id):
    """Generate order status API request details"""
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    payload = {
        "apikey": OpenAlgoConfig.api_key,
        "strategy": str(strategy),
        "orderid": str(order_id)
    }
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/orderstatus"
    
    return [
        ["URL", endpoint],
        ["Method", "POST"],
        ["Content-Type", "application/json"],
        ["JSON Body", json.dumps(payload)],
        ["Order ID", str(order_id)],
        ["Curl Command", f'curl -X POST "{endpoint}" -H "Content-Type: application/json" -d \'{json.dumps(payload)}\'']
    ]

# Legacy function names for compatibility (redirect to request generators)
@func
def oa_quotes(symbol, exchange):
    """Generate quotes request (use oa_quotes_request for detailed output)"""
    return oa_quotes_request(symbol, exchange)

@func  
def oa_depth(symbol, exchange):
    """Generate market depth request (use oa_depth_request for detailed output)"""
    return oa_depth_request(symbol, exchange)

@func
def oa_history(symbol, exchange, interval, start_date, end_date):
    """Generate historical data request (use oa_history_request for detailed output)"""
    return oa_history_request(symbol, exchange, interval, start_date, end_date)

@func
def oa_intervals():
    """Generate intervals request (use oa_intervals_request for detailed output)"""
    return oa_intervals_request()

@func
def oa_funds():
    """Generate funds request (use oa_funds_request for detailed output)"""
    return oa_funds_request()

@func
def oa_orderbook():
    """Generate orderbook request (use oa_orderbook_request for detailed output)"""
    return oa_orderbook_request()

@func
def oa_tradebook():
    """Generate tradebook request (use oa_tradebook_request for detailed output)"""
    return oa_tradebook_request()

@func
def oa_positionbook():
    """Generate positionbook request (use oa_positionbook_request for detailed output)"""
    return oa_positionbook_request()

@func
def oa_holdings():
    """Generate holdings request (use oa_holdings_request for detailed output)"""
    return oa_holdings_request()

@func
def oa_placeorder(strategy, symbol, action, exchange, pricetype, product, quantity, price=0, trigger_price=0, disclosed_quantity=0):
    """Generate place order request - CAUTION: REAL ORDERS!"""
    return oa_placeorder_request(strategy, symbol, action, exchange, pricetype, product, quantity, price, trigger_price, disclosed_quantity)

@func
def oa_test_connection():
    """Test configuration and generate funds request"""
    if not validate_api_key():
        return format_error("OpenAlgo API Key is not set. Use oa_api()")
    
    return [
        ["Status", "API Key Set"],
        ["Test", "Use oa_funds_request() and run curl command"],
        ["Host", OpenAlgoConfig.host_url],
        ["Version", OpenAlgoConfig.version]
    ]

# Utility Functions
@func
def oa_all_functions():
    """List all available OpenAlgo functions"""
    return [
        ["Category", "Function", "Description"],
        ["Setup", "oa_api(api_key, version, host_url)", "Set API configuration"],
        ["Setup", "oa_get_config()", "View current configuration"],
        ["Setup", "get_status()", "Check system status"],
        ["Market Data", "oa_quotes_request(symbol, exchange)", "Get quotes"],
        ["Market Data", "oa_depth_request(symbol, exchange)", "Get market depth"],
        ["Market Data", "oa_history_request(symbol, exchange, interval, start_date, end_date)", "Get historical data"],
        ["Market Data", "oa_intervals_request()", "Get available intervals"],
        ["Account", "oa_funds_request()", "Get account funds"],
        ["Account", "oa_orderbook_request()", "Get order book"],
        ["Account", "oa_tradebook_request()", "Get trade book"],
        ["Account", "oa_positionbook_request()", "Get position book"],
        ["Account", "oa_holdings_request()", "Get holdings"],
        ["Orders", "oa_placeorder_request(strategy, symbol, action, exchange, pricetype, product, quantity, price)", "Place order"],
        ["Orders", "oa_modifyorder_request(strategy, order_id, symbol, action, exchange, quantity)", "Modify order"],
        ["Orders", "oa_cancelorder_request(strategy, order_id)", "Cancel order"],
        ["Orders", "oa_orderstatus_request(strategy, order_id)", "Get order status"],
        ["Help", "oa_power_query_guide()", "Power Query setup instructions"],
        ["Help", "oa_all_functions()", "This function list"]
    ]

@func
def oa_power_query_guide():
    """Instructions for using Power Query with OpenAlgo requests"""
    return [
        ["Step", "Action"],
        ["1", "Go to Data > Get Data > From Other Sources > From Web"],
        ["2", "Select 'Advanced' option"],
        ["3", "Copy URL from any oa_*_request() function"],
        ["4", "Change HTTP method to POST"],
        ["5", "Add header: Content-Type = application/json"],
        ["6", "Copy JSON Body from the function result"],
        ["7", "Paste JSON in request body section"],
        ["8", "Click OK to execute the request"],
        ["9", "Excel will import the API response as a table"],
        ["Note", "Save the query to refresh data easily"]
    ]