import pandas as pd
import re
from collections import defaultdict

# Load the Excel file
file_path = 'vivo/2025-01-27_at_02_16_24516PM_culture.xlsx'  # Replace with the actual path to your file
df = pd.read_excel(file_path, sheet_name=0)
output_path = 'vivo_results/Culture_results.xlsx'
# Exclude rows where the year in `Stored / frozen on` is 2025
df['Stored / frozen on'] = pd.to_datetime(df['Stored / frozen on'], errors='coerce')  # Convert to datetime
df = df[df['Stored / frozen on'].dt.year != 2025]  # Filter out rows where year is 2025

# Extract amount and unit from `Units remarks`
def parse_unit_remarks(remarks):
    match = re.match(r'(\d+\.?\d*)\s*([a-zA-Zµ]+)', str(remarks))
    if match:
        amount = float(match.group(1))
        unit = match.group(2)
        return amount, unit
    return None, None

df[['Units_remarks_amount', 'Units_remarks_unit']] = df['Units remarks'].apply(
    lambda x: pd.Series(parse_unit_remarks(x))
)

# Determine final amount and unit for each row
def get_final_amount_unit(row):
    if row['Stock volume'] == 0 and pd.notnull(row['Stock weight']):
        return row['Stock weight'], row['Weight units']
    elif row['Stock weight'] == 0 and pd.notnull(row['Stock volume']):
        return row['Stock volume'], row['Volume units']
    elif pd.notnull(row['Stock volume']) and pd.notnull(row['Volume units']):
        return row['Stock volume'], row['Volume units']
    elif pd.notnull(row['Stock weight']) and pd.notnull(row['Weight units']):
        return row['Stock weight'], row['Weight units']
    return None, None

df[['Amount_final', 'Unit_final']] = df.apply(
    lambda row: pd.Series(get_final_amount_unit(row)), axis=1
)

# Group by `Stock name` and calculate totals for each unit
grouped = df.groupby('Stock name')
results = []

for stock_name, group in grouped:
    unit_totals = defaultdict(float)
    total_price = 0  # Initialize total price
    
    for _, row in group.iterrows():
        amount = row['Amount_final']
        unit = row['Unit_final']
        price_per_unit = row['Price']
        unit_remarks_amount = row['Units_remarks_amount']
        unit_remarks_unit = row['Units_remarks_unit']
        
        if pd.notnull(amount) and pd.notnull(unit):
            unit_totals[unit] += amount
            
            # Calculate price based on `Units remarks`
            if pd.notnull(unit_remarks_amount) and unit == unit_remarks_unit:
                total_price += (amount / unit_remarks_amount) * price_per_unit
    
    # Combine totals into a single string
    total_amounts = ', '.join(f"{value} {unit}" for unit, value in unit_totals.items())
    
    # Collect other data
    results.append({
        'Stock name': stock_name,
        'Count': len(group),
        'Manufacturer': group['Manufacturer'].iloc[0],
        'Catalog no.': group['Catalog no.'].iloc[0],
        'Price': group['Price'].iloc[0],
        'Units remarks': group['Units remarks'].iloc[0],  # Include `Units remarks`
        'Total Amount & Unit': total_amounts,
        'Total Price': total_price
    })

# Convert results into a DataFrame
result_df = pd.DataFrame(results)

# Reorder columns
result_df = result_df[['Stock name', 'Count', 'Manufacturer', 'Catalog no.', 'Price', 'Units remarks', 'Total Amount & Unit', 'Total Price']]

# Apply conditional formatting
def highlight_zero_price(row):
    if row['Total Price'] == 0 or pd.isnull(row['Total Price']):
        return ['background-color: #ffcccc'] * len(row)
    return [''] * len(row)

styled_result = result_df.style.apply(highlight_zero_price, axis=1)

# Save the styled DataFrame as an Excel file

styled_result.to_excel(output_path, index=False, engine='openpyxl')

print(f"Styled results have been saved to {output_path}")
