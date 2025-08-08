import pandas as pd
from datetime import datetime
import calendar

def filter_departments(df):
    """Filter out records where department contains 'UPC'"""
    original_count = len(df)
    df_filtered = df[~df['department'].str.contains('UPC', case=False, na=False)]
    filtered_count = original_count - len(df_filtered)
    
    if filtered_count > 0:
        print(f"\nFiltered out {filtered_count} records containing 'UPC' in department")
        print(f"Original record count: {original_count}")
        print(f"Records after filtering: {len(df_filtered)}")
    else:
        print("\nNo records found with 'UPC' in department")
    
    return df_filtered

def validate_dates(df):
    """Validate if dates follow the first/last day of month rule"""
    
    def is_first_day(date_str):
        if pd.isna(date_str):
            return True
        date = pd.to_datetime(date_str)
        return date.day == 1

    def is_last_day(date_str):
        if pd.isna(date_str):
            return True
        date = pd.to_datetime(date_str)
        _, last_day = calendar.monthrange(date.year, date.month)
        return date.day == last_day

    # Convert dates to datetime
    df['startdate'] = pd.to_datetime(df['startdate'], errors='coerce')
    df['enddate'] = pd.to_datetime(df['enddate'], errors='coerce')

    # Find invalid dates
    invalid_start = df[~df['startdate'].apply(lambda x: is_first_day(x) if pd.notna(x) else True)]
    invalid_end = df[~df['enddate'].apply(lambda x: is_last_day(x) if pd.notna(x) else True)]

    if not invalid_start.empty:
        print("\nRows with invalid start dates (not first day of month):")
        print(invalid_start[['id', 'akronim', 'lastname', 'firstname', 'startdate']].to_string())

    if not invalid_end.empty:
        print("\nRows with invalid end dates (not last day of month):")
        print(invalid_end[['id', 'akronim', 'lastname', 'firstname', 'enddate']].to_string())

    return invalid_start.empty and invalid_end.empty

def analyze_akronims(df, label=""):
    """Analyze and print akronim statistics"""
    distinct_akronims = df['akronim'].nunique()
    print(f"\nDistinct akronim count {label}: {distinct_akronims}")
    
    # Get counts for each akronim
    akronim_counts = df['akronim'].value_counts()
    
    # Print akronims with multiple entries
    multiple_entries = akronim_counts[akronim_counts > 1]
    if not multiple_entries.empty:
        print("\nAkronims with multiple entries:")
        for akronim, count in multiple_entries.items():
            print(f"{akronim}: {count} entries")

def is_valid_month_end(date):
    """Check if a date is the last day of its month, handling leap years"""
    if pd.isna(date):
        return True
    date = pd.to_datetime(date)
    _, last_day = calendar.monthrange(date.year, date.month)
    return date.day == last_day

def validate_dates(df):
    """Validate if dates follow the first/last day of month rule"""
    
    def is_first_day(date_str):
        if pd.isna(date_str):
            return True
        date = pd.to_datetime(date_str)
        return date.day == 1

    # Convert dates to datetime
    df['startdate'] = pd.to_datetime(df['startdate'], errors='coerce')
    df['enddate'] = pd.to_datetime(df['enddate'], errors='coerce')

    # Find invalid dates
    invalid_start = df[~df['startdate'].apply(lambda x: is_first_day(x) if pd.notna(x) else True)]
    invalid_end = df[~df['enddate'].apply(is_valid_month_end)]

    if not invalid_start.empty:
        print("\nRows with invalid start dates (not first day of month):")
        print(invalid_start[['id', 'akronim', 'lastname', 'firstname', 'startdate']].to_string())

    if not invalid_end.empty:
        print("\nRows with invalid end dates (not last day of month):")
        print(invalid_end[['id', 'akronim', 'lastname', 'firstname', 'enddate']].to_string())

    return invalid_start.empty and invalid_end.empty

def compress_to_quarters(df):
    """
    Compress consecutive monthly records into quarterly records.
    Handles leap years correctly when determining month end dates.
    """
    # Create a copy to avoid modifying the original
    df = df.copy()
    
    # Ensure dates are datetime
    df['startdate'] = pd.to_datetime(df['startdate'])
    df['enddate'] = pd.to_datetime(df['enddate'])
    
    # Sort by akronim and startdate
    df = df.sort_values(['akronim', 'startdate'])
    
    def get_quarter_dates(date):
        quarter = (date.month - 1) // 3
        start_month = quarter * 3 + 1
        quarter_start = datetime(date.year, start_month, 1)
        
        # Get the correct end date of the last month in quarter
        quarter_end_month = min(start_month + 2, 12)
        _, last_day = calendar.monthrange(date.year, quarter_end_month)
        quarter_end = datetime(date.year, quarter_end_month, last_day)
        
        return quarter_start, quarter_end
    
    compressed_records = []
    current_group = None
    
    for _, row in df.iterrows():
        if current_group is None:
            current_group = row.to_dict()
            continue
            
        # Check if this row should be merged with current group
        curr_end = current_group['enddate']
        next_start = row['startdate']
        
        # Get quarter information for both dates
        curr_quarter_start, curr_quarter_end = get_quarter_dates(curr_end)
        next_quarter_start, next_quarter_end = get_quarter_dates(next_start)
        
        # Check if records are consecutive
        # For February, we need to handle the variable month length
        _, curr_month_days = calendar.monthrange(curr_end.year, curr_end.month)
        consecutive = (next_start - curr_end).days == 1
        same_person = current_group['akronim'] == row['akronim']
        same_quarter = curr_quarter_start == next_quarter_start
        
        if consecutive and same_person and same_quarter:
            # Update end date of current group
            current_group['enddate'] = row['enddate']
        else:
            # Add current group to results and start new group
            compressed_records.append(current_group)
            current_group = row.to_dict()
    
    # Add last group
    if current_group is not None:
        compressed_records.append(current_group)
    
    # Convert back to DataFrame
    compressed_df = pd.DataFrame(compressed_records)
    
    # Ensure same column order as input
    compressed_df = compressed_df[df.columns]
    
    return compressed_df

def test_leap_year_handling():
    """Test function to verify correct handling of leap years"""
    # Create test data including February in leap and non-leap years
    test_data = pd.DataFrame({
        'id': range(6),
        'akronim': ['TEST'] * 6,
        'lastname': ['Test'] * 6,
        'firstname': ['User'] * 6,
        'startdate': [
            '2024-01-01', '2024-02-01', '2024-03-01',  # Leap year
            '2023-01-01', '2023-02-01', '2023-03-01'   # Non-leap year
        ],
        'enddate': [
            '2024-01-31', '2024-02-29', '2024-03-31',  # Note February 29 in leap year
            '2023-01-31', '2023-02-28', '2023-03-31'   # Note February 28 in non-leap year
        ]
    })
    
    print("\nTesting leap year handling...")
    # Validate dates
    if not validate_dates(test_data):
        print("Date validation failed!")
        return
    
    # Compress test data
    compressed = compress_to_quarters(test_data)
    print("\nTest Results:")
    print("2024 Q1 (leap year):", compressed[compressed['startdate'].dt.year == 2024].iloc[0].to_dict())
    print("2023 Q1 (non-leap year):", compressed[compressed['startdate'].dt.year == 2023].iloc[0].to_dict())
    
    return compressed

def standardize_department(df):
    """Standardize department values, changing any containing 'MON' to just 'MON'"""
    # Store original values that will be changed
    mon_related = df[df['department'].str.contains('MON', case=False, na=False)]
    
    if not mon_related.empty:
        print("\nStandardizing department values:")
        for idx, row in mon_related.iterrows():
            print(f"Changing department from '{row['department']}' to 'MON' for akronim: {row['akronim']}")
    
    # Update the values
    df['department'] = df['department'].apply(
        lambda x: 'MON' if pd.notna(x) and 'MON' in str(x).upper() else x
    )
    
    return df

def process_csv(input_file, output_file):
    """Process CSV file to remove duplicates, validate dates, and compress to quarters"""
    try:
        # Read the CSV file
        print(f"Reading file: {input_file}")
        df = pd.read_csv(input_file)
        
        # Store original row count
        original_count = len(df)
        print(f"Original number of rows: {original_count}")

        # Filter out UPC departments
        df_filtered = filter_departments(df)
        # Standardize department values
        df_standarized = standardize_department(df_filtered)
        # Analyze akronims before deduplication
        analyze_akronims(df_standarized, "before deduplication")

        # Remove duplicates considering all columns
        df_no_dupes = df_standarized.drop_duplicates()
        
        # Calculate number of duplicates removed
        dupes_removed = len(df_standarized) - len(df_no_dupes)
        if dupes_removed > 0:
            print(f"\nRemoved {dupes_removed} duplicate rows")
        else:
            print("\nNo duplicate rows found")

        # Validate dates
        print("\nValidating dates...")
        dates_valid = validate_dates(df_no_dupes)

        if not dates_valid:
            print("\nWarning: Invalid dates found. Please fix dates before compression.")
            return

        # Compress to quarters
        print("\nCompressing to quarterly records...")
        df_compressed = compress_to_quarters(df_no_dupes)
        
        # Print compression statistics
        compression_reduction = len(df_no_dupes) - len(df_compressed)
        print(f"\nCompressed {compression_reduction} rows by combining quarterly records")
        
        # Save processed file
        df_compressed.to_csv(output_file, index=False)
        print(f"\nProcessed file saved as: {output_file}")

        # Print summary statistics
        print("\nSummary:")
        print(f"Original rows: {original_count}")
        print(f"Rows with UPC filtered out: {original_count - len(df_filtered)}")
        print(f"Duplicate rows removed: {dupes_removed}")
        print(f"Rows compressed: {compression_reduction}")
        print(f"Final rows: {len(df_compressed)}")
        print(f"Original distinct akronims: {df['akronim'].nunique()}")
        print(f"Final distinct akronims: {df_compressed['akronim'].nunique()}")

    except Exception as e:
        print(f"Error processing file: {str(e)}")

if __name__ == "__main__":
    input_file = "dane.csv"  # Replace with your input file name
    output_file = "output.csv"  # Replace with desired output file name
    process_csv(input_file, output_file)