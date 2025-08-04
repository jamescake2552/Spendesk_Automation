import os
import pandas as pd
from datetime import datetime

def clean_csv_file(input_path, output_path):
    """Cleans the CSV file by removing quotes and trailing semicolons."""
    try:
        with open(input_path, 'r', encoding='utf-8') as file:
            content = file.read()

        cleaned_content = content.replace('"', '')
        lines = cleaned_content.strip().split('\n')
        cleaned_lines = [line.strip().rstrip(';') for line in lines if line.strip()]

        with open(output_path, 'w', encoding='utf-8', newline='') as file:
            file.write('\n'.join(cleaned_lines))

        print(f"✓ Cleaned and saved to: {output_path}")
    except FileNotFoundError:
        print(f"❌ File not found: {input_path}")
    except Exception as e:
        print(f"❌ Error: {e}")

def enrich_and_save_excel(input_csv_path, output_excel_path, reference_excel_path):
    """Enriches the CSV with Department and Location, adds a Summary sheet."""
    try:
        df = pd.read_csv(input_csv_path, sep=';')

        # === Load Reference Sheets ===
        reference_sheets = pd.read_excel(reference_excel_path, sheet_name=["Employee", "Account"])

        # --- Employee Sheet (for Department) ---
        emp_df = reference_sheets["Employee"]
        emp_df.columns = [col.strip().lower() for col in emp_df.columns]
        emp_df.rename(columns={'spendesk names': 'Payer'}, inplace=True)

        if 'netsuite department' in emp_df.columns:
            df = df.merge(emp_df[['Payer', 'netsuite department']], on='Payer', how='left')
            df.insert(0, 'Department', df.pop('netsuite department'))
        else:
            df.insert(0, 'Department', "")

        # --- Location (from Signed Total Amount) ---
        amount_col = [col for col in df.columns if 'signed total amount' in col.lower()]
        if amount_col:
            signed_col = amount_col[0]
            # -----------------------------------------------
            # -----------------------------------------------
            # UPDATE VALUE OF £250 HERE

            df.insert(1, 'Location', df[signed_col].apply(lambda x: "Central" if pd.to_numeric(x, errors='coerce') < 250 else ""))
            
            # -----------------------------------------------
            # -----------------------------------------------
        else:
            df.insert(1, 'Location', "")

        # Save the enriched data
        with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Data', index=False)

        print(f"✓ Saved cleaned Excel to: {output_excel_path}")

        # === Create Summary Sheet ===
        today = datetime.today()
        current_date_str = today.strftime('%d/%m/%Y')
        prev_month = today.month - 1 if today.month > 1 else 12
        prev_year = today.year if today.month > 1 else today.year - 1
        prev_month_str = datetime(prev_year, prev_month, 1).strftime('%b')
        posting_period = f"{prev_month_str}-{str(prev_year)[-2:]}"  # e.g., "Jun-25"

        template_header = {
            "REFERENCE": f"Spendesk {posting_period}",
            "VENDOR": "Spendesk",
            "ACCOUNT GENERAL": "20001 Accounts Payable: Trade Creditors",
            "MEMO": f"Spendesk {posting_period}",
            "DATE": current_date_str,
            "POSTING PERIOD": posting_period
        }

        df = pd.read_excel(output_excel_path, sheet_name='Data')

        # Column mapping
        expense_col = next((col for col in df.columns if 'expense account' in col.lower()), None)
        net_col = next((col for col in df.columns if 'net amount' in col.lower()), None)
        tax_col = next((col for col in df.columns if 'tax amount' in col.lower()), None)
        total_col = next((col for col in df.columns if 'signed total amount' in col.lower()), None)

        if not all([expense_col, net_col, tax_col, total_col]):
            print("❌ Could not find all required columns for summary sheet.")
            return

        # Ensure blanks are consistent for grouping
        df['Department'] = df['Department'].fillna("Unassigned")
        df['Location'] = df['Location'].replace('', 'Blank').fillna("Blank")

        # Group by Expense Account, Department, and Location
        grouped = df.groupby([expense_col, 'Department', 'Location'], dropna=False).agg({
            net_col: 'sum',
            tax_col: 'sum',
            total_col: 'sum'
        }).reset_index()
        grouped.rename(columns={expense_col: 'Expense Account Number'}, inplace=True)


        # Lookup Display Name for each account
        acct_df = reference_sheets["Account"]
        acct_df.columns = [col.strip() for col in acct_df.columns]
        lookup_df = acct_df[['Expense Account Number', 'Display Name']]

        merged = grouped.merge(lookup_df, on='Expense Account Number', how='left')

        summary_data = []

        for _, row in merged.iterrows():
            tax_code = "VAT:S-GB" if row[tax_col] > 0 else "VAT:Z-GB"

            summary_data.append({
                'REFERENCE': template_header['REFERENCE'],
                'VENDOR': template_header['VENDOR'],
                'ACCOUNT GENERAL': template_header['ACCOUNT GENERAL'],
                'MEMO': template_header['MEMO'],
                'DATE': template_header['DATE'],
                'POSTING PERIOD': template_header['POSTING PERIOD'],
                'ACCOUNT SPECIFIC': row['Display Name'] if pd.notna(row['Display Name']) else row['Expense Account Number'],
                'AMOUNT': row[net_col],
                'TAX CODE': tax_code,
                'TAX AMOUNT': row[tax_col],
                'GROSS AMOUNT': row[total_col],
                'DEPARTMENT': row['Department'],
                'LOCATION': row['Location']
            })

        summary_df = pd.DataFrame(summary_data)

        # Append to Excel
        with pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            summary_df.to_excel(writer, sheet_name='Summary', index=False)

        print("✓ Added 'Summary' sheet grouped by account, department, and location.")

    except Exception as e:
        print(f"❌ Error during Excel processing: {e}")

def main():
    print("=== CSV Cleaner and Excel Generator ===")

    input_csv = input("Enter path to the CSV file: ").strip().strip('"').strip("'")
    output_excel = input("Enter path to save Excel (.xlsx): ").strip().strip('"').strip("'")
    ref_excel = input("Enter path to employee + account reference Excel file: ").strip().strip('"').strip("'")

    if not output_excel.endswith('.xlsx'):
        output_excel += '.xlsx'

    temp_cleaned = input_csv.rsplit('.', 1)[0] + '_cleaned.csv'

    clean_csv_file(input_csv, temp_cleaned)
    enrich_and_save_excel(temp_cleaned, output_excel, ref_excel)

    try:
        os.remove(temp_cleaned)
    except:
        print(f"Note: Could not remove temporary file: {temp_cleaned}")

if __name__ == "__main__":
    main()
