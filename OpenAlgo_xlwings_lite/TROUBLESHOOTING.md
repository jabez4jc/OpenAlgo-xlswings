# OpenAlgo xlwings Lite Troubleshooting Guide

## #BUSY! Error Solutions

The `#BUSY!` error in Excel indicates that xlwings Lite cannot execute the Python function. Follow these steps to diagnose and fix the issue:

### Step 1: Verify xlwings Lite Installation

1. **Check xlwings Lite is installed:**
   - Go to `Insert > Add-ins > My Add-ins`
   - Ensure "xlwings Lite" is listed and enabled
   - If not installed, get it from Office Add-in Store

2. **Restart Excel** after installing xlwings Lite

### Step 2: Use Minimal Test Version

Instead of the full `main.py`, start with `main_minimal.py`:

1. **Copy `main_minimal.py` contents**
2. **Open xlwings Lite editor** (xlwings tab in Excel ribbon)
3. **Paste the minimal code** and save
4. **Test with simple functions first:**

```excel
=test_xlwings()          # Should return "xlwings Lite is working! ✓"
=test_imports()          # Shows package availability
=test_config()           # Shows configuration status
```

### Step 3: Common Fixes

#### Fix 1: Clear and Reload Code
1. Open xlwings Lite editor
2. Select all code (Ctrl+A) and delete
3. Paste fresh code from `main_minimal.py`
4. Save and try functions again

#### Fix 2: Check Excel Calculation Mode
1. Go to `Formulas > Calculation Options`
2. Ensure it's set to "Automatic"
3. Try manual calculation: `Ctrl+Shift+F9`

#### Fix 3: Function Name Issues
- Use exact function names: `=test_xlwings()` not `=TEST_XLWINGS()`
- Ensure no spaces: `=test_xlwings()` not `= test_xlwings()`

#### Fix 4: Excel File Format
- Save workbook as `.xlsm` (Excel Macro-Enabled Workbook)
- Regular `.xlsx` files may not work with xlwings Lite

#### Fix 5: Trust Center Settings
1. Go to `File > Options > Trust Center > Trust Center Settings`
2. Go to `Add-ins` and ensure xlwings Lite is trusted
3. Enable "Trust access to the VBA project object model"

### Step 4: Diagnostic Tests

Run these tests in order:

```excel
# Test 1: Basic xlwings functionality
=test_xlwings()

# Test 2: Package availability 
=test_imports()

# Test 3: Configuration system
=test_config()

# Test 4: Simple API configuration
=oa_api_simple("test_key", "v1", "http://127.0.0.1:5000")

# Test 5: API connection (requires real API key)
=test_api_connection()

# Test 6: Simple market data (requires real API key)
=oa_quotes_simple("RELIANCE", "NSE")
```

### Step 5: Use Debug Script

1. In xlwings Lite editor, run the debug script:
   - Type `debug_xlwings` in the script box
   - Click "Run" button
2. This will populate your worksheet with diagnostic information

### Step 6: Check xlwings Console

1. In xlwings Lite editor, check the **Console** tab
2. Look for error messages or import failures
3. Common issues:
   - `ModuleNotFoundError`: Missing packages
   - `SyntaxError`: Code syntax issues
   - `ImportError`: Package compatibility issues

### Common Error Messages and Solutions

#### Error: "ModuleNotFoundError: No module named 'pandas'"
**Solution:** Remove pandas import or use the minimal version without pandas

#### Error: "Function not found"
**Solution:** 
- Ensure code is saved in xlwings editor
- Check function has `@func` decorator
- Restart Excel and reload code

#### Error: "pyodide_http not available"
**Solution:** This is normal - the code handles this gracefully

#### Error: "HTTP Error 500/503"
**Solution:** 
- Check OpenAlgo server is running
- Verify API endpoint URL
- Test with `=test_api_connection()`

### Platform-Specific Issues

#### Windows
- Ensure Windows Defender isn't blocking xlwings Lite
- Check firewall settings for Excel

#### macOS
- Grant Excel permissions in System Preferences > Security & Privacy
- Ensure Gatekeeper allows xlwings Lite

#### Excel on Web
- Some advanced formatting may not work
- Ensure internet connection for WebAssembly packages
- Clear browser cache if functions stop working

### Advanced Troubleshooting

#### Enable Developer Mode
1. `File > Options > Customize Ribbon`
2. Check "Developer" tab
3. Use Developer tools for advanced debugging

#### Check Event Viewer (Windows)
1. Open Event Viewer
2. Go to Windows Logs > Application
3. Look for Excel or xlwings errors

#### Browser Console (Excel on Web)
1. Press F12 to open developer tools
2. Check Console tab for JavaScript errors
3. Look for WebAssembly or Pyodide errors

### Getting Help

If none of these solutions work:

1. **Create a test workbook** with just the minimal functions
2. **Document the exact error** (screenshot + any console messages)
3. **Note your environment:**
   - Excel version
   - Operating system
   - xlwings Lite version
4. **Test on different platforms** if possible (Windows/Mac/Web)

### Migration from Full Version

Once minimal version works:

1. **Test each function group individually**
2. **Add one function at a time** from full `main.py`
3. **Test after each addition** to identify problematic functions
4. **Use simplified versions** of functions that cause issues

### Success Indicators

When everything is working correctly, you should see:

- `=test_xlwings()` returns "xlwings Lite is working! ✓"
- `=test_imports()` shows all packages available
- Functions execute without `#BUSY!` errors
- xlwings console shows no error messages
- Debug script populates worksheet with information

### Contact Information

For persistent issues:
- Check xlwings Lite documentation
- Review Pyodide package compatibility
- Test with OpenAlgo API documentation examples