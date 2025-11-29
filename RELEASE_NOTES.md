# AlSawife Factory Release Notes

## Version 1.0.0 (2025-11-29)

This is the first official release of the AlSawife Factory application, a comprehensive solution for managing factory operations and invoices.

### Key Features

1. **Invoice Management**
   - Create and manage invoices with detailed product information
   - Automatic calculation of area and total prices
   - Professional Excel export with proper formatting

2. **Client Management**
   - Maintain client information and history
   - Automatic ledger updates with each invoice
   - Easy client lookup with autocomplete functionality

3. **Product Pricing**
   - Flexible pricing system based on JSON configuration
   - Support for complex pricing structures with ranges
   - Automatic price lookup based on product selection

4. **Excel Integration**
   - Professional invoice templates with Arabic support
   - Automatic calculation formulas in Excel sheets
   - Summary tables with proper aggregation

### Major Fixes

- Fixed critical SUM formula issues in invoice summary tables
- Resolved attribute access errors throughout the application
- Fixed import issues with utility modules
- Corrected number input formatting problems
- Addressed Excel export issues including missing prefixes

### Technical Improvements

- Robust error handling and user feedback
- Improved UI with better scaling support
- Enhanced data validation and input sanitization
- Proper file organization and modular structure

### Known Limitations

- Application is currently Windows-focused
- Requires manual installation of dependencies
- No automatic update mechanism

### Installation Instructions

1. Ensure Python 3.8 or higher is installed
2. Install required packages: `pip install -r requirements.txt`
3. Run the application: `python main.py`

For standalone executable, run: `python build.py` to create a distributable .exe file.