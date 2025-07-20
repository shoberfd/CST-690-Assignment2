import os
from dotenv import load_dotenv
from report_generator import generate_sales_report, setup_logging
import logging

def main(): # Main function to run the sales report generation process
    # Set up logging to capture all messages
    setup_logging()

    logging.info("Loading environment variables...")
    load_dotenv()     # Load environment variables from .env file
    logging.info("Environment variables loaded successfully.")

    # Get configuration from environment variables
    sales_data_file = os.getenv("SALES_DATA_FILE")
    output_report_dir = os.getenv("OUTPUT_REPORT_DIR")

    if not sales_data_file or not output_report_dir:
        logging.error("Missing required environment variables. "
                      "Please ensure SALES_DATA_FILE and OUTPUT_REPORT_DIR are set in your .env file.")
        return

    logging.info(f"Configuration loaded: Sales Data File='{sales_data_file}', Output Directory='{output_report_dir}'")

    # Run the report generation process
    generate_sales_report(sales_data_file, output_report_dir)

if __name__ == "__main__":
    main()

# CITATIONS FOR MAIN.PY AND REPORT_GENERATOR.PY
# 1. Python Software Foundation. (n.d.). Logging facility for Python. 
#    Python 3.x Documentation. Retrieved from https://docs.python.org/3/library/logging.html
# 2. Python Software Foundation. (n.d.). Python-dotenv: Read key-value pairs from a 
#    .env file and set them as environment variables.
#    Retrieved from https://pypi.org/project/python-dotenv/
# 3. Pandas Development Team. (n.d.). Pandas: Powerful data structures for data analysis, 
#    time series, and statistics. Retrieved from https://pandas.pydata.org/
# 4. OpenPyXL Development Team. (n.d.). OpenPyXL: A library to read/write Excel 2010 
#    xlsx/xlsm/xltx/xltm files.
#    Retrieved from https://openpyxl.readthedocs.io/en/stable/
# 5. Python Software Foundation. (n.d.). Python: A programming language that lets you work
#    quickly and integrate systems more effectively. Retrieved from https://www.python.org/
# 6. Stack Overflow. (n.d.). How to set up logging in Python? Retrieved from
#    https://stackoverflow.com/questions/1579727/how-to-set-up-logging-in-python