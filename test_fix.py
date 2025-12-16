import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from utils.blocks_utils import export_simple_blocks_excel

# Test data
test_data = [{
    "trip_number": "T001",
    "trip_count": "2",
    "date": "2023-01-01",
    "quarry": "Q001",
    "machine_number": "M001",
    "block_number": "B001",
    "material": "Granite",
    "length": "2.5",
    "weight": "1000",
    "price_per_ton": "50",
    "trip_price": "50000",
    "thickness_dropdown": "2سم",
    "quantity": "3",
    "publish_height": "1.8",
    "height": "1.5"
}]

# Test the export function
try:
    filepath = export_simple_blocks_excel(test_data)
    print(f"Successfully created file: {filepath}")
except Exception as e:
    print(f"Error: {e}")
    import traceback
    traceback.print_exc()