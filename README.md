# Invoice Creator Application

A GUI application for creating invoices and exporting them to Excel format.

## Features

- User-friendly interface built with CustomTkinter
- Automatic invoice numbering with SQLite database storage
- Product database with JSON file storage
- Excel export with professional formatting
- Date selection with calendar widget
- Automatic calculations for area and total prices

## Requirements

- Python 3.7+
- CustomTkinter
- tkcalendar
- xlsxwriter

## Installation

1. Install the required packages:
   ```
   pip install -r requirements.txt
   ```

## Usage

Run the application:
```
python main.py
```

## Project Structure

- `main.py` - Entry point of the application
- `ui.py` - User interface implementation
- `db_utils.py` - Database utilities for invoice numbering
- `excel_utils.py` - Excel file generation
- `res/products.json` - Product database
- `res/invoice.db` - SQLite database for invoice counters

## How to Use

1. Enter invoice details (operation number, client, driver, date, phone)
2. Add items to the invoice using the "إضافة بند" button
3. Fill in item details (description, block number, thickness, material, count, length, height, price)
4. Save the invoice to Excel using the "حفظ إلى Excel" button
5. Start a new invoice with the "عملية جديدة" button

## License

This project is licensed under the MIT License.