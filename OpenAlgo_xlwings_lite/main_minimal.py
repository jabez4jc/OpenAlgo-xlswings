"""OpenAlgo Excel Integration via xlwings Lite - Minimal Version

This is a simplified version for troubleshooting #BUSY! errors.
Start with this version to verify xlwings Lite is working properly.
"""

import xlwings as xw
from xlwings import func, script
import json
import urllib.request
from datetime import datetime

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

# Test Functions First
@func
def test_xlwings():
    """Test if xlwings Lite is working properly"""
    return "xlwings Lite is working! ✓"

@func
def test_imports():
    """Test which packages are available"""
    results = [["Package", "Status"]]
    
    # Test core imports
    try:
        import json
        results.append(["json", "✓ Available"])
    except ImportError:
        results.append(["json", "✗ Missing"])
    
    try:
        import urllib.request
        results.append(["urllib.request", "✓ Available"])
    except ImportError:
        results.append(["urllib.request", "✗ Missing"])
    
    try:
        from datetime import datetime
        results.append(["datetime", "✓ Available"])
    except ImportError:
        results.append(["datetime", "✗ Missing"])
    
    # Test pyodide
    if PYODIDE_AVAILABLE:
        results.append(["pyodide_http", "✓ Available"])
    else:
        results.append(["pyodide_http", "✗ Missing"])
    
    # Test pandas (optional)
    try:
        import pandas
        results.append(["pandas", "✓ Available"])
    except ImportError:
        results.append(["pandas", "✗ Missing"])
    
    return results

@func
def test_config():
    """Test configuration system"""
    return [
        ["Setting", "Value"],
        ["API Key", "Not Set" if not OpenAlgoConfig.api_key else "Set"],
        ["Version", OpenAlgoConfig.version],
        ["Host URL", OpenAlgoConfig.host_url],
        ["Pyodide", "Available" if PYODIDE_AVAILABLE else "Not Available"]
    ]

# Simple HTTP function
def simple_post_request(endpoint, payload):
    """Simplified HTTP POST request"""
    try:
        data = json.dumps(payload).encode('utf-8')
        headers = {'Content-Type': 'application/json'}
        
        request = urllib.request.Request(endpoint, data=data, headers=headers)
        response = urllib.request.urlopen(request, timeout=10)
        
        return json.loads(response.read().decode('utf-8'))
    except Exception as e:
        return {"error": str(e)}

# Basic Configuration Function
@func
def oa_api_simple(api_key, version="v1", host_url="http://127.0.0.1:5000"):
    """Simplified version of oa_api for testing"""
    if not api_key or not api_key.strip():
        return "Error: API Key is required"
    
    OpenAlgoConfig.api_key = str(api_key).strip()
    OpenAlgoConfig.version = str(version)
    OpenAlgoConfig.host_url = str(host_url).rstrip('/')
    
    return f"Config set: Key={api_key[:4]}***, Version={version}, Host={host_url}"

# Basic Test Function
@func
def test_api_connection():
    """Test API connection with simplified error handling"""
    if not OpenAlgoConfig.api_key:
        return [["Status", "Error: API Key not set"]]
    
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/funds"
    payload = {"apikey": OpenAlgoConfig.api_key}
    
    response = simple_post_request(endpoint, payload)
    
    if "error" in response:
        return [
            ["Status", "Connection Failed"],
            ["Error", response["error"]],
            ["Endpoint", endpoint]
        ]
    else:
        return [
            ["Status", "Connection Success"],
            ["Endpoint", endpoint],
            ["Response", "OK"]
        ]

# Simple Market Data Function
@func
def oa_quotes_simple(symbol, exchange):
    """Simplified quotes function for testing"""
    if not OpenAlgoConfig.api_key:
        return [["Error", "API Key not set. Use oa_api_simple()"]]
    
    endpoint = f"{OpenAlgoConfig.host_url}/api/{OpenAlgoConfig.version}/quotes"
    payload = {
        "apikey": OpenAlgoConfig.api_key,
        "symbol": str(symbol),
        "exchange": str(exchange)
    }
    
    response = simple_post_request(endpoint, payload)
    
    if "error" in response:
        return [["Error", response["error"]]]
    
    data = response.get("data", {})
    if not data:
        return [["Error", "No data received"]]
    
    # Simple key-value format
    result = [[f"{symbol} ({exchange})", "Value"]]
    for key, value in data.items():
        result.append([str(key), str(value)])
    
    return result

# Debug Script
@script
def debug_xlwings(book: xw.Book):
    """Debug xlwings Lite setup"""
    try:
        sheet = book.sheets.active
        
        # Clear area for debug info
        sheet.range("A1:C20").clear()
        
        # Header
        sheet["A1"].value = "xlwings Lite Debug Information"
        sheet["A1"].font.bold = True
        
        # Test basic xlwings
        sheet["A3"].value = "Basic Test:"
        sheet["B3"].value = test_xlwings()
        
        # Test imports
        sheet["A5"].value = "Package Status:"
        import_results = test_imports()
        for i, row in enumerate(import_results):
            sheet[f"A{6+i}"].value = row[0]
            sheet[f"B{6+i}"].value = row[1]
        
        # Test config
        config_row = 6 + len(import_results) + 2
        sheet[f"A{config_row}"].value = "Configuration:"
        config_results = test_config()
        for i, row in enumerate(config_results):
            sheet[f"A{config_row+1+i}"].value = row[0]
            sheet[f"B{config_row+1+i}"].value = row[1]
            
        sheet[f"A{config_row+len(config_results)+2}"].value = f"Debug completed at {datetime.now().strftime('%H:%M:%S')}"
        
    except Exception as e:
        sheet["A1"].value = f"Debug Error: {str(e)}"