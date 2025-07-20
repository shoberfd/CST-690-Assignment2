import pandas as pd
import os
import logging
from datetime import datetime

# --- Configuration and Setup ---

def setup_logging():
    log_dir = "logs"
    os.makedirs(log_dir, exist_ok=True) # Ensure logs directory exists
    log_file_path = os.path.join(log_dir, "automation.log")

    logging.basicConfig(
        level=logging.INFO, # Set the minimum level of messages to log
        format='%(asctime)s - %(levelname)s - %(message)s', # Format of log messages
        handlers=[
            logging.FileHandler(log_file_path), # Log to a file
            logging.StreamHandler() # Log to console as well
        ]
    )
    logging.info("Logging setup complete.")

# --- Core Business Logic Functions ---

def load_sales_data(file_path: str) -> pd.DataFrame:
    logging.info(f"Attempting to load sales data from: {file_path}")
    if not os.path.exists(file_path):
        logging.error(f"Error: Sales data file not found at {file_path}")
        raise FileNotFoundError(f"Sales data file not found: {file_path}")
    try:
        df = pd.read_csv(file_path)
        logging.info(f"Successfully loaded {len(df)} records from {file_path}")
        return df
    except pd.errors.EmptyDataError:
        logging.warning(f"The CSV file at {file_path} is empty.")
        return pd.DataFrame() # Return an empty DataFrame if file is empty
    except Exception as e:
        logging.error(f"An unexpected error occurred while reading {file_path}: {e}")
        raise # Re-raise the exception after logging

def clean_and_process_data(df: pd.DataFrame) -> pd.DataFrame:
    logging.info("Starting data cleaning and processing...")

    if df.empty:
        logging.warning("No data to process. Returning empty DataFrame.")
        return df

    # Convert 'Date' column to datetime objects
    if 'Date' in df.columns:
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        # Drop rows where Date conversion failed
        df.dropna(subset=['Date'], inplace=True)
        logging.info("Converted 'Date' column to datetime.")
    else:
        logging.warning("Date column not found. Skipping date conversion.")

    # Ensure numeric columns are of the correct type
    numeric_cols = ['Quantity', 'UnitPrice', 'TotalPrice']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
            # Fill NaN created by coercion with 0, or drop rows if critical
            df[col].fillna(0, inplace=True)
            logging.info(f"Ensured '{col}' column is numeric and handled NaNs.")
        else:
            logging.warning(f"Numeric column '{col}' not found. Skipping type conversion.")

    # Handle missing values in categorical columns (e.g., fill with 'Unknown')
    categorical_cols = ['ProductName', 'Category', 'Region', 'SalespersonID']
    for col in categorical_cols:
        if col in df.columns:
            df[col].fillna('Unknown', inplace=True)
            logging.info(f"Filled missing values in '{col}' with 'Unknown'.")
        else:
            logging.warning(f"Categorical column '{col}' not found. Skipping NaN handling.")

    logging.info("Data cleaning and processing complete.")
    return df

def aggregate_sales_data(df: pd.DataFrame) -> dict:
    logging.info("Aggregating sales data for various reports...")
    reports = {}

    if df.empty:
        logging.warning("No data to aggregate. Returning empty reports.")
        return reports

    # Overall Summary
    total_revenue = df['TotalPrice'].sum()
    total_quantity = df['Quantity'].sum()
    num_transactions = df['TransactionID'].nunique()
    avg_transaction_value = total_revenue / num_transactions if num_transactions > 0 else 0

    summary_data = {
        'Metric': ['Total Revenue', 'Total Quantity Sold', 'Number of Transactions', 'Average Transaction Value'],
        'Value': [total_revenue, total_quantity, num_transactions, avg_transaction_value]
    }
    reports['Summary'] = pd.DataFrame(summary_data)
    logging.info("Generated Sales Summary report.")

    # Sales by Category
    if 'Category' in df.columns:
        sales_by_category = df.groupby('Category')['TotalPrice'].sum().reset_index()
        sales_by_category.rename(columns={'TotalPrice': 'Total Revenue'}, inplace=True)
        reports['Sales by Category'] = sales_by_category.sort_values(by='Total Revenue', ascending=False)
        logging.info("Generated Sales by Category report.")
    else:
        logging.warning("Category column not found. Skipping 'Sales by Category' report.")

    # Sales by Region
    if 'Region' in df.columns:
        sales_by_region = df.groupby('Region')['TotalPrice'].sum().reset_index()
        sales_by_region.rename(columns={'TotalPrice': 'Total Revenue'}, inplace=True)
        reports['Sales by Region'] = sales_by_region.sort_values(by='Total Revenue', ascending=False)
        logging.info("Generated Sales by Region report.")
    else:
        logging.warning("Region column not found. Skipping 'Sales by Region' report.")

    # Top 5 Products by Revenue
    if 'ProductName' in df.columns and 'TotalPrice' in df.columns:
        top_products = df.groupby('ProductName')['TotalPrice'].sum().reset_index()
        top_products.rename(columns={'TotalPrice': 'Total Revenue'}, inplace=True)
        reports['Top 5 Products'] = top_products.sort_values(by='Total Revenue', ascending=False).head(5)
        logging.info("Generated Top 5 Products report.")
    else:
        logging.warning("ProductName or TotalPrice column not found. Skipping 'Top 5 Products' report.")

    logging.info("Sales data aggregation complete.")
    return reports

def save_reports_to_excel(reports: dict, output_dir: str, file_name: str):
    os.makedirs(output_dir, exist_ok=True) # Ensure output directory exists
    output_path = os.path.join(output_dir, file_name)
    logging.info(f"Attempting to save reports to: {output_path}")

    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            if not reports:
                logging.warning("No reports to save. Creating an empty Excel file.")
                # Create a dummy sheet if no reports
                pd.DataFrame({'Message': ['No data available for reports.']}).to_excel(writer, sheet_name='No Data', index=False)
            else:
                for sheet_name, df_report in reports.items():
                    df_report.to_excel(writer, sheet_name=sheet_name, index=False)
                    logging.info(f"Sheet '{sheet_name}' saved to Excel.")
        logging.info(f"Successfully saved all reports to {output_path}")
    except Exception as e:
        logging.error(f"Error saving reports to Excel file {output_path}: {e}")
        raise # Re-raise the exception after logging

# --- Main Automation Workflow ---

def generate_sales_report(sales_data_file: str, output_report_dir: str):
    setup_logging()
    logging.info("Starting Daily Sales Report Generation Automation.")

    try:
        # 1. Load Data
        raw_df = load_sales_data(sales_data_file)
        if raw_df.empty:
            logging.warning("No sales data loaded. Skipping report generation.")
            # Still create an empty report file to indicate process ran but no data
            today_date = datetime.now().strftime("%Y-%m-%d")
            report_file_name = f"Daily_Sales_Report_{today_date}.xlsx"
            save_reports_to_excel({}, output_report_dir, report_file_name)
            logging.info("Daily Sales Report Generation Automation finished (no data).")
            return

        # 2. Clean and Process Data
        processed_df = clean_and_process_data(raw_df.copy()) # Use a copy to avoid modifying original

        # 3. Aggregate Sales Data
        aggregated_reports = aggregate_sales_data(processed_df)

        # 4. Save Report
        today_date = datetime.now().strftime("%Y-%m-%d")
        report_file_name = f"Daily_Sales_Report_{today_date}.xlsx"
        save_reports_to_excel(aggregated_reports, output_report_dir, report_file_name)

        logging.info("Daily Sales Report Generation Automation completed successfully.")

    except FileNotFoundError as fnfe:
        logging.error(f"Automation failed: {fnfe}")
    except pd.errors.EmptyDataError as ede:
        logging.error(f"Automation failed due to empty data file: {ede}")
    except Exception as e:
        logging.error(f"An unhandled error occurred during automation: {e}", exc_info=True)