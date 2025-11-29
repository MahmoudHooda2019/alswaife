#!/usr/bin/env python3
# Debug the summary table formulas

# Let's trace through the logic with concrete numbers
print("Debugging summary table formulas")
print("=" * 40)

# Example scenario:
summary_start_row = 20      # Row with "اجمالي الفاتورة" (0-indexed)
summary_header_row = 21     # Row with column headers (0-indexed)
summary_data_start_row = 22 # First data row (0-indexed)
row_counter = 3             # Number of data rows
summary_total_row = 25      # Row where we put the total (0-indexed)

print(f"summary_start_row: {summary_start_row} (title row)")
print(f"summary_header_row: {summary_header_row} (header row)")
print(f"summary_data_start_row: {summary_data_start_row} (first data row)")
print(f"row_counter: {row_counter} (number of data rows)")
print(f"summary_total_row: {summary_total_row} (total row)")

print("\nData rows:")
for i in range(row_counter):
    data_row_0indexed = summary_data_start_row + i
    data_row_1indexed = data_row_0indexed + 1
    print(f"  Data row {i+1}: 0-indexed {data_row_0indexed}, 1-indexed {data_row_1indexed}")

print(f"\nTotal row: 0-indexed {summary_total_row}, 1-indexed {summary_total_row + 1}")

print("\nCurrent formula (what we're generating):")
area_formula = f"=SUM(D{summary_data_start_row+1}:D{summary_total_row})"
price_formula = f"=SUM(E{summary_data_start_row+1}:E{summary_total_row})"
print(f"Area formula: {area_formula}")
print(f"Price formula: {price_formula}")

print("\nThis includes the total row in the sum, which is WRONG!")

print("\nCorrect formula should be:")
correct_area_formula = f"=SUM(D{summary_data_start_row+1}:D{summary_total_row})"
correct_price_formula = f"=SUM(E{summary_data_start_row+1}:E{summary_total_row})"
print(f"Area formula: {correct_area_formula}")
print(f"Price formula: {correct_price_formula}")

print("\nWait, that's the same. Let me think...")

print("\nThe issue is that summary_total_row is where we're PUTTING the total,")
print("so we shouldn't INCLUDE it in the sum.")
print("We want to sum from the first data row to the LAST DATA ROW.")

last_data_row_0indexed = summary_data_start_row + row_counter - 1
last_data_row_1indexed = last_data_row_0indexed + 1

print(f"\nLast data row: 0-indexed {last_data_row_0indexed}, 1-indexed {last_data_row_1indexed}")

correct_area_formula = f"=SUM(D{summary_data_start_row+1}:D{last_data_row_1indexed})"
correct_price_formula = f"=SUM(E{summary_data_start_row+1}:E{last_data_row_1indexed})"
print(f"Correct area formula: {correct_area_formula}")
print(f"Correct price formula: {correct_price_formula}")

# Alternative way to express it:
correct_area_formula_alt = f"=SUM(D{summary_data_start_row+1}:D{summary_total_row})"
correct_price_formula_alt = f"=SUM(E{summary_data_start_row+1}:E{summary_total_row})"
print(f"\nAlternative (should be same): {correct_area_formula_alt}")
print(f"Alternative (should be same): {correct_price_formula_alt}")