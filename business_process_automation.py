import os
import sys
from datetime import date
import pandas as pd

def main():
    csv_path = get_csv_path()
    output_dir = create_orders_dir(os.path.dirname(csv_path))
    process_sales_data(csv_path, output_dir)
    print(f"Successfully processed all orders. Files saved to: {output_dir}")

def get_csv_path():
    if len(sys.argv) < 2:
        print("Error: Missing command line parameter. Please provide the path to the sales data CSV file.", file=sys.stderr)
        sys.exit(1)
        
    csv_path = sys.argv[1]
    
    if not os.path.isfile(csv_path):
        print(f"Error: The provided path '{csv_path}' does not specify an existing file.", file=sys.stderr)
        sys.exit(1)
        
    return os.path.abspath(csv_path)

def create_orders_dir(base_path):
    today_str = date.today().isoformat()
    dir_name = f"Orders_{today_str}"
    full_output_path = os.path.join(base_path, dir_name)
    
    if not os.path.exists(full_output_path):
        os.makedirs(full_output_path)
        
    return full_output_path

def process_sales_data(csv_path, output_dir):
    df = pd.read_csv(csv_path)
    df['TOTAL PRICE'] = df['ITEM QUANTITY'] * df['ITEM PRICE']
    df = df.sort_values(by=['ORDER ID', 'ITEM NUMBER'])
    
    cols_to_keep = [
        'ORDER DATE', 'ITEM NUMBER', 'PRODUCT LINE', 'PRODUCT CODE',
        'ITEM QUANTITY', 'ITEM PRICE', 'TOTAL PRICE', 'STATUS', 'CUSTOMER NAME'
    ]
    
    grouped = df.groupby('ORDER ID')
    
    for order_id, order_frame in grouped:
        final_order_df = order_frame[cols_to_keep]
        save_order_excel(order_id, final_order_df, output_dir)

def save_order_excel(order_id, order_df, output_dir):
    file_name = f"Order_{order_id}.xlsx"
    file_path = os.path.join(output_dir, file_name)
    
    writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
    order_df.to_excel(writer, sheet_name='Order Summary', index=False)
    
    workbook  = writer.book
    worksheet = writer.sheets['Order Summary']
    
    money_format = workbook.add_format({'num_format': '$#,##0.00'})
    
    grand_total = order_df['TOTAL PRICE'].sum()
    total_row_idx = len(order_df) + 1
    
    worksheet.write(total_row_idx, 5, "GRAND TOTAL:")
    worksheet.write(total_row_idx, 6, grand_total, money_format)
    
    column_configs = [
        (0, 11, None),
        (1, 13, None),
        (2, 15, None),
        (3, 15, None),
        (4, 15, None),
        (5, 13, money_format),
        (6, 13, money_format),
        (7, 10, None),
        (8, 30, None)
    ]
    
    for col_idx, width, fmt in column_configs:
        if fmt is not None:
            worksheet.set_column(col_idx, col_idx, width, fmt)
        else:
            worksheet.set_column(col_idx, col_idx, width)
        
    writer.close()

if __name__ == "__main__":
    main()