"""Ultra-Simple Test for xlwings Lite

Copy this code into xlwings Lite editor to test #BUSY! error fixes.
Test functions one by one to identify exactly where the issue occurs.
"""

import xlwings as xw
from xlwings import func, script
import json
import urllib.request

# Try pyodide_http import
try:
    import pyodide_http
    pyodide_http.patch_all()
except ImportError:
    pass

# Simple config class
class Config:
    api_key = ""
    version = "v1"
    host_url = "http://127.0.0.1:5000"

# Test 1: Basic xlwings function (should work)
@func
def test_basic():
    return "Basic test works!"

# Test 2: Return array
@func
def test_array():
    return [["Test", "Array"], ["Works", "Fine"]]

# Test 3: Configuration function
@func
def set_config(api_key, version="v1", host_url="https://openalgo.simplifyed.in"):
    Config.api_key = str(api_key)
    Config.version = str(version) 
    Config.host_url = str(host_url)
    return f"Config set: {api_key[:4]}***, {version}, {host_url}"

# Test 4: Check config
@func
def get_config():
    return [
        ["Setting", "Value"],
        ["API Key", "Set" if Config.api_key else "Not Set"],
        ["Version", Config.version],
        ["Host URL", Config.host_url]
    ]

# Test 5: Simple HTTP request
@func
def test_http():
    if not Config.api_key:
        return [["Error", "Set API key first with set_config()"]]
    
    try:
        endpoint = f"{Config.host_url}/api/{Config.version}/funds"
        payload = {"apikey": Config.api_key}
        
        data = json.dumps(payload).encode('utf-8')
        headers = {'Content-Type': 'application/json'}
        
        request = urllib.request.Request(endpoint, data=data, headers=headers)
        response = urllib.request.urlopen(request, timeout=10)
        result = json.loads(response.read().decode('utf-8'))
        
        if "error" in result:
            return [["Status", "Error"], ["Message", result["error"]]]
        else:
            return [["Status", "Success"], ["Data", "Received"]]
            
    except Exception as e:
        return [["Status", "Failed"], ["Error", str(e)]]

# Test 6: Market data
@func
def test_quotes(symbol, exchange):
    if not Config.api_key:
        return [["Error", "Set API key first"]]
    
    try:
        endpoint = f"{Config.host_url}/api/{Config.version}/quotes"
        payload = {
            "apikey": Config.api_key,
            "symbol": str(symbol),
            "exchange": str(exchange)
        }
        
        data = json.dumps(payload).encode('utf-8')
        headers = {'Content-Type': 'application/json'}
        
        request = urllib.request.Request(endpoint, data=data, headers=headers)
        response = urllib.request.urlopen(request, timeout=10)
        result = json.loads(response.read().decode('utf-8'))
        
        if "error" in result:
            return [["Error", result["error"]]]
        
        quote_data = result.get("data", {})
        if not quote_data:
            return [["Error", "No data received"]]
        
        # Format as simple table
        output = [[f"{symbol} ({exchange})", "Value"]]
        for key, value in quote_data.items():
            output.append([str(key), str(value)])
        
        return output
        
    except Exception as e:
        return [["Error", str(e)]]