import pandas as pd
import os
from datetime import datetime
import logging

# Set up logging for better debugging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def get_file_paths():
    """Get file paths from user input"""
    print("📁 Please provide the required file paths:")
    print("=" * 50)
    
    # Get bookkeeping file path
    bookkeeping_path = input("Enter the path to the bookkeeping Excel file: ").strip().strip('"')
    
    # Get statement file path
    statement_path = input("Enter the path to the statement Excel file: ").strip().strip('"')
    
    # Get output file path
    output_path = input("Enter the path where you want to save the reconciliation results: ").strip().strip('"')
    
    return bookkeeping_path, statement_path, output_path

def validate_file_exists(file_path, file_type):
    """Validate that the Excel file exists"""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"{file_type} file not found: {file_path}")
    print(f"✓ {file_type} file found: {file_path}")

def validate_output_directory(output_path):
    """Validate that the output directory exists"""
    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        raise FileNotFoundError(f"Output directory not found: {output_dir}")
    print(f"✓ Output directory validated: {output_dir if output_dir else 'Current directory'}")

def load_workbook_data(file_path, file_type, expected_columns):
    """Load data from a single workbook (assumes data is on the first sheet)"""
    print(f"\n📊 Loading {file_type} data from: {os.path.basename(file_path)}")
    
    try:
        # Read the first sheet
        df = pd.read_excel(file_path, sheet_name=0)
        
        initial_rows = len(df)
        print(f"  - Initial rows: {initial_rows}")
        print(f"  - Available columns: {list(df.columns)}")
        
        # Check if expected columns exist (with some flexibility)
        available_columns = list(df.columns)
        column_mapping = {}
        
        for expected_col in expected_columns:
            # Try to find the column (exact match first, then case-insensitive)
            if expected_col in available_columns:
                column_mapping[expected_col] = expected_col
            else:
                # Try case-insensitive match
                for col in available_columns:
                    if col.lower() == expected_col.lower():
                        column_mapping[expected_col] = col
                        break
        
        # Check if we found all required columns
        missing_columns = [col for col in expected_columns if col not in column_mapping]
        if missing_columns:
            print(f"  ❌ Missing columns: {missing_columns}")
            print(f"  📋 Available columns: {available_columns}")
            raise ValueError(f"Missing required columns in {file_type}: {missing_columns}")
        
        # Select and rename columns
        df = df[[column_mapping[col] for col in expected_columns]]
        df.columns = expected_columns
        
        print(f"  - Selected columns: {expected_columns}")
        
        # Clean the data
        df = clean_data(df, file_type)
        
        return df
        
    except Exception as e:
        print(f"❌ Error loading {file_type}: {str(e)}")
        raise

def clean_data(df, file_type):
    """Clean data with aggressive blank removal and text normalization"""
    initial_rows = len(df)
    
    # Step 1: Remove rows where all columns are null
    df = df.dropna(how='all')
    after_all_null = len(df)
    print(f"  - After removing all-null rows: {after_all_null}")
    
    # Step 2: Clean and normalize text columns (Payer and Description)
    print(f"  - Cleaning and normalizing text columns...")
    
    # Clean Payer column
    if 'Payer' in df.columns:
        df['Payer'] = df['Payer'].astype(str).str.strip()
        # Normalize whitespace in Payer column (replace newlines and multiple spaces with single space)
        df['Payer'] = df['Payer'].str.replace(r'\s+', ' ', regex=True)
        
        # Define what counts as "blank" for Payer
        payer_blank_conditions = (
            (df['Payer'] == 'nan') |
            (df['Payer'] == 'NaN') |
            (df['Payer'] == 'None') |
            (df['Payer'] == '') |
            (df['Payer'].str.lower() == 'nan') |
            (df['Payer'].str.contains(r'^\s*$', regex=True, na=False))
        )
        
        # Clean Description column and define blank conditions
        if 'Description' in df.columns:
            df['Description'] = df['Description'].astype(str).str.strip()
            # Normalize whitespace in Description column (replace newlines and multiple spaces with single space)
            df['Description'] = df['Description'].str.replace(r'\s+', ' ', regex=True)
            
            description_blank_conditions = (
                (df['Description'] == 'nan') |
                (df['Description'] == 'NaN') |
                (df['Description'] == 'None') |
                (df['Description'] == '') |
                (df['Description'].str.lower() == 'nan') |
                (df['Description'].str.contains(r'^\s*$', regex=True, na=False))
            )
        else:
            description_blank_conditions = pd.Series(False, index=df.index)
        
        # Remove rows where BOTH Payer AND Description are blank
        both_blank_conditions = payer_blank_conditions & description_blank_conditions
        rows_with_both_blank = both_blank_conditions.sum()
        print(f"    - Rows with BOTH payer and description blank: {rows_with_both_blank}")
        
        # Remove only rows where both are blank
        df = df[~both_blank_conditions]
        
        after_both_blank_filter = len(df)
        print(f"  - After removing rows with both blank payer AND description: {after_both_blank_filter}")
        
        # Show what we kept
        payer_only_blank = (~payer_blank_conditions & description_blank_conditions).sum()
        description_only_blank = (payer_blank_conditions & ~description_blank_conditions).sum()
        if payer_only_blank > 0:
            print(f"    - Kept {payer_only_blank} rows with blank payer but valid description")
        if description_only_blank > 0:
            print(f"    - Kept {description_only_blank} rows with blank description but valid payer")
    
    # Step 3: Remove rows where amount is null
    amount_col = 'Signed Total Amount' if file_type == 'Bookkeeping' else 'Debit'
    if amount_col in df.columns:
        df = df.dropna(subset=[amount_col])
        after_amount_filter = len(df)
        print(f"  - After removing null amounts: {after_amount_filter}")
    
    cleaned_rows = len(df)
    total_removed = initial_rows - cleaned_rows
    
    print(f"  - Final rows: {cleaned_rows}")
    print(f"  - Total removed: {total_removed} rows ({total_removed/initial_rows*100:.1f}%)")
    
    # Show sample data for verification
    if len(df) > 0:
        unique_payers = df['Payer'].nunique()
        print(f"  - Unique payers: {unique_payers}")
        
        # Show samples of different types of rows
        has_payer_and_desc = df[(df['Payer'] != '') & (df['Description'] != '')]
        has_payer_only = df[(df['Payer'] != '') & (df['Description'] == '')]
        has_desc_only = df[(df['Payer'] == '') & (df['Description'] != '')]
        
        print(f"  - Rows with both payer and description: {len(has_payer_and_desc)}")
        print(f"  - Rows with payer only: {len(has_payer_only)}")
        print(f"  - Rows with description only: {len(has_desc_only)}")
        
        # Show sample payers (excluding blank ones)
        sample_payers = df[df['Payer'] != '']['Payer'].unique()[:5]
        if len(sample_payers) > 0:
            print(f"  - Sample payers: {list(sample_payers)}")
        
        # Show sample descriptions after normalization (for debugging)
        sample_descriptions = df[df['Description'] != '']['Description'].unique()[:3]
        if len(sample_descriptions) > 0:
            print(f"  - Sample normalized descriptions: {list(sample_descriptions)}")
    
    # Additional validation
    if df.empty:
        raise ValueError(f"No data found in {file_type} after cleaning")
    
    return df

def prepare_comparison_data(bookkeeping_df, statement_df):
    """Prepare data for comparison with consistent column names"""
    print("\n🔄 Preparing data for comparison...")
    
    # Create comparison dataframes with standardized column names
    comparison_bookkeeping = bookkeeping_df.rename(
        columns={'Signed Total Amount': 'Amount'}
    ).copy()
    
    comparison_statement = statement_df.rename(
        columns={'Debit': 'Amount'}
    ).copy()
    
    # Add original indices to track which rows to remove
    comparison_bookkeeping = comparison_bookkeeping.reset_index().rename(columns={'index': 'original_idx'})
    comparison_statement = comparison_statement.reset_index().rename(columns={'index': 'original_idx'})
    
    print(f"  - Bookkeeping records prepared: {len(comparison_bookkeeping)}")
    print(f"  - Statement records prepared: {len(comparison_statement)}")
    
    return comparison_bookkeeping, comparison_statement

def find_matches(comparison_bookkeeping, comparison_statement):
    """Find matching transactions between bookkeeping and statement"""
    print("\n🔍 Finding matching transactions...")
    
    # Find matches using merge with indicator to handle duplicates properly
    merged = pd.merge(
        comparison_bookkeeping,
        comparison_statement,
        on=['Payer', 'Description', 'Amount'],
        how='inner',
        suffixes=('_bookkeeping', '_statement')
    )
    
    print(f"  - Found {len(merged)} matching transactions")
    
    if len(merged) > 0:
        print("\n📋 Sample of matched transactions:")
        print(merged[['Payer', 'Description', 'Amount']].head(10).to_string(index=False))
        
        # Show summary statistics
        total_matched_amount = merged['Amount'].sum()
        print(f"\n💰 Total matched amount: £{total_matched_amount:,.2f}")
    else:
        print("  ⚠️  No matching transactions found")
    
    return merged

def generate_outlier_report(bookkeeping_df, statement_df, merged):
    """Generate report of unmatched transactions"""
    print("\n📊 Generating outlier report...")
    
    # Get the indices of matched rows to remove
    bookkeeping_indices_to_remove = merged['original_idx_bookkeeping'].tolist()
    statement_indices_to_remove = merged['original_idx_statement'].tolist()
    
    print(f"  - Removing {len(bookkeeping_indices_to_remove)} matched bookkeeping records")
    print(f"  - Removing {len(statement_indices_to_remove)} matched statement records")
    
    # Remove matched rows from both dataframes
    outlier_bookkeeping = bookkeeping_df.drop(bookkeeping_indices_to_remove).dropna(how='all')
    outlier_statement = statement_df.drop(statement_indices_to_remove).dropna(how='all')
    
    print(f"  - Unmatched bookkeeping records: {len(outlier_bookkeeping)}")
    print(f"  - Unmatched statement records: {len(outlier_statement)}")
    
    # Standardize column names for combining
    outlier_bookkeeping = outlier_bookkeeping.rename(columns={'Signed Total Amount': 'Amount'}).copy()
    outlier_statement = outlier_statement.rename(columns={'Debit': 'Amount'}).copy()
    
    # Create the side-by-side comparison format
    combined_report = create_side_by_side_format(outlier_bookkeeping, outlier_statement)
    
    print(f"  - Report formatted with {len(combined_report)} rows")
    
    if len(outlier_bookkeeping) > 0 or len(outlier_statement) > 0:
        # Calculate summary statistics
        bookkeeping_total = outlier_bookkeeping['Amount'].sum() if len(outlier_bookkeeping) > 0 else 0
        statement_total = outlier_statement['Amount'].sum() if len(outlier_statement) > 0 else 0
        
        print(f"\n💰 Unmatched amounts:")
        print(f"  - Bookkeeping: £{bookkeeping_total:,.2f}")
        print(f"  - Statement: £{statement_total:,.2f}")
        print(f"  - Difference: £{abs(bookkeeping_total - statement_total):,.2f}")
    
    return combined_report

def create_side_by_side_format(outlier_bookkeeping, outlier_statement):
    """Create side-by-side format grouped by payer"""
    print("🔄 Creating side-by-side format...")
    
    # Get all unique payers from both datasets
    all_payers = set()
    if len(outlier_bookkeeping) > 0:
        all_payers.update(outlier_bookkeeping['Payer'].unique())
    if len(outlier_statement) > 0:
        all_payers.update(outlier_statement['Payer'].unique())
    
    all_payers = sorted(all_payers)
    
    # Create the final report structure
    report_data = []
    
    for payer in all_payers:
        # Get bookkeeping and statement data for this payer
        payer_bookkeeping = outlier_bookkeeping[outlier_bookkeeping['Payer'] == payer] if len(outlier_bookkeeping) > 0 else pd.DataFrame()
        payer_statement = outlier_statement[outlier_statement['Payer'] == payer] if len(outlier_statement) > 0 else pd.DataFrame()
        
        # Get the maximum number of rows for this payer
        max_rows = max(len(payer_bookkeeping), len(payer_statement))
        
        # Create rows for this payer
        for i in range(max_rows):
            row = {'Payer': payer if i == 0 else ''}  # Only show payer name on first row
            
            # Bookkeeping data
            if i < len(payer_bookkeeping):
                bookkeeping_row = payer_bookkeeping.iloc[i]
                row['Bookkeeping Description'] = bookkeeping_row['Description']
                row['Bookkeeping Amount'] = bookkeeping_row['Amount']
            else:
                row['Bookkeeping Description'] = ''
                row['Bookkeeping Amount'] = ''
            
            # Statement data
            if i < len(payer_statement):
                statement_row = payer_statement.iloc[i]
                row['Statement Description'] = statement_row['Description']
                row['Statement Amount'] = statement_row['Amount']
                row['Statement Credit'] = statement_row['Credit']
            else:
                row['Statement Description'] = ''
                row['Statement Amount'] = ''
                row['Statement Credit'] = ''
            
            report_data.append(row)
        
        # Add total row for this payer
        bookkeeping_total = payer_bookkeeping['Amount'].sum() if len(payer_bookkeeping) > 0 else 0
        statement_total = payer_statement['Amount'].sum() if len(payer_statement) > 0 else 0
        statement_credit = payer_statement['Credit'].sum() if len(payer_statement) > 0 else 0
        
        total_row = {
            'Payer': '',
            'Bookkeeping Description': 'Total',
            'Bookkeeping Amount': f'£{bookkeeping_total:.2f}' if bookkeeping_total != 0 else '',
            'Statement Description': 'Total',
            'Statement Amount': f'£{statement_total:.2f}' if statement_total != 0 else '',
            'Statement Credit': f'£{statement_credit:.2f}' if statement_credit != 0 else ''
        }
        report_data.append(total_row)
        
        # Add empty row for spacing (except for last payer)
        if payer != all_payers[-1]:
            empty_row = {
                'Payer': '',
                'Bookkeeping Description': '',
                'Bookkeeping Amount': '',
                'Statement Description': '',
                'Statement Amount': '',
                'Statement Credit': ''
            }
            report_data.append(empty_row)
    
    return pd.DataFrame(report_data)

def save_comprehensive_report(combined_report, bookkeeping_df, statement_df, output_path):
    """Save the comprehensive reconciliation report with multiple sheets"""
    print("\n💾 Saving comprehensive reconciliation report...")
    
    # Ensure output path has .xlsx extension
    if not output_path.lower().endswith('.xlsx'):
        output_path = output_path + '.xlsx'
    
    # Add timestamp to filename if not already present
    if '_RECONCILIATION_' not in output_path:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        directory = os.path.dirname(output_path)
        filename = os.path.basename(output_path)
        name_without_ext = os.path.splitext(filename)[0]
        
        new_filename = f"{name_without_ext}_RECONCILIATION_{timestamp}.xlsx"
        output_path = os.path.join(directory, new_filename)
    
    try:
        print(f"  📝 Preparing to write comprehensive report...")
        print(f"  📄 Output file: {os.path.basename(output_path)}")
        print(f"  📊 Outliers sheet: {len(combined_report)} rows")
        print(f"  📊 Bookkeeping sheet: {len(bookkeeping_df)} rows")
        print(f"  📊 Statement sheet: {len(statement_df)} rows")
        
        # Create Excel writer with formatting
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            # Sheet 1: Outliers (Reconciliation Results)
            combined_report.to_excel(
                writer, 
                sheet_name='Outliers', 
                index=False,
                startrow=0
            )
            
            # Sheet 2: Full Bookkeeping Data
            bookkeeping_df.to_excel(
                writer,
                sheet_name='Bookkeeping',
                index=False,
                startrow=0
            )
            
            # Sheet 3: Full Statement Data
            statement_df.to_excel(
                writer,
                sheet_name='Statement',
                index=False,
                startrow=0
            )
            
            # Get the workbook for formatting
            workbook = writer.book
            
            # Define formats
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#D3D3D3',
                'border': 1
            })
            
            total_format = workbook.add_format({
                'bold': True,
                'bg_color': '#E6E6FA',
                'border': 1
            })
            
            # Format Outliers sheet
            outliers_worksheet = writer.sheets['Outliers']
            
            # Format headers
            for col_num, value in enumerate(combined_report.columns.values):
                outliers_worksheet.write(0, col_num, value, header_format)
            
            # Format total rows
            for row_num in range(1, len(combined_report) + 1):
                if row_num <= len(combined_report) and combined_report.iloc[row_num - 1]['Bookkeeping Description'] == 'Total':
                    for col_num in range(len(combined_report.columns)):
                        cell_value = combined_report.iloc[row_num - 1, col_num]
                        outliers_worksheet.write(row_num, col_num, cell_value, total_format)
            
            # Set column widths for Outliers sheet
            outliers_worksheet.set_column('A:A', 20)  # Payer
            outliers_worksheet.set_column('B:B', 30)  # Bookkeeping Description
            outliers_worksheet.set_column('C:C', 15)  # Bookkeeping Amount
            outliers_worksheet.set_column('D:D', 30)  # Statement Description
            outliers_worksheet.set_column('E:E', 15)  # Statement Amount
            
            # Format Bookkeeping sheet
            bookkeeping_worksheet = writer.sheets['Bookkeeping']
            for col_num, value in enumerate(bookkeeping_df.columns.values):
                bookkeeping_worksheet.write(0, col_num, value, header_format)
            
            # Set column widths for Bookkeeping sheet
            bookkeeping_worksheet.set_column('A:A', 20)  # Payer
            bookkeeping_worksheet.set_column('B:B', 30)  # Description
            bookkeeping_worksheet.set_column('C:C', 15)  # Amount
            
            # Format Statement sheet
            statement_worksheet = writer.sheets['Statement']
            for col_num, value in enumerate(statement_df.columns.values):
                statement_worksheet.write(0, col_num, value, header_format)
            
            # Set column widths for Statement sheet
            statement_worksheet.set_column('A:A', 20)  # Payer
            statement_worksheet.set_column('B:B', 30)  # Description
            statement_worksheet.set_column('C:C', 15)  # Amount
        
        print(f"  ✅ Comprehensive report saved successfully!")
        print(f"  📂 Full path: {output_path}")
        print(f"  📋 Sheets created:")
        print(f"    - Outliers: Unmatched transactions ({len(combined_report)} rows)")
        print(f"    - Bookkeeping: Full bookkeeping data ({len(bookkeeping_df)} rows)")
        print(f"    - Statement: Full statement data ({len(statement_df)} rows)")
        
        return output_path
        
    except Exception as e:
        print(f"  ❌ Error saving report: {str(e)}")
        raise

def main():
    """Main reconciliation process"""
    print("🚀 Starting Spendesk Reconciliation Process")
    print("=" * 60)
    
    try:
        # Step 1: Get file paths from user
        bookkeeping_path, statement_path, output_path = get_file_paths()
        
        # Step 2: Validate files exist
        validate_file_exists(bookkeeping_path, "Bookkeeping")
        validate_file_exists(statement_path, "Statement")
        validate_output_directory(output_path)
        
        # Step 3: Load data from separate workbooks
        bookkeeping_columns = ['Payer', 'Description', 'Signed Total Amount']
        statement_columns = ['Payer', 'Description', 'Debit', 'Credit']
        
        bookkeeping_df = load_workbook_data(bookkeeping_path, 'Bookkeeping', bookkeeping_columns)
        statement_df = load_workbook_data(statement_path, 'Statement', statement_columns)
        
        # Step 4: Prepare comparison data
        comparison_bookkeeping, comparison_statement = prepare_comparison_data(bookkeeping_df, statement_df)
        
        # Step 5: Find matches
        merged = find_matches(comparison_bookkeeping, comparison_statement)
        
        # Step 6: Generate outlier report
        combined_report = generate_outlier_report(bookkeeping_df, statement_df, merged)
        
        # Step 7: Save comprehensive report
        final_location = save_comprehensive_report(combined_report, bookkeeping_df, statement_df, output_path)
        
        print("\n" + "=" * 60)
        print("✅ Reconciliation process completed successfully!")
        print(f"📊 Summary:")
        print(f"  - Total bookkeeping records: {len(bookkeeping_df)}")
        print(f"  - Total statement records: {len(statement_df)}")
        print(f"  - Matched transactions: {len(merged)}")
        print(f"  - Comprehensive report saved with 3 sheets")
        print(f"  - Output file: {final_location}")
        
    except Exception as e:
        print(f"\n❌ Process failed: {str(e)}")
        logging.error(f"Reconciliation failed: {str(e)}")
        raise

if __name__ == "__main__":
    main()