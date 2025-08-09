"""
Sales Report Generator
----------------------
Reads a sales_data.csv file and generates:
1. Total sales per product
2. Total sales per salesperson
3. Monthly sales summary
4. Year-to-date total
Outputs:
    - Formatted console tables
    - Excel file with separate sheets for each summary
Handles invalid/missing rows gracefully.
"""


import pandas as pd
from pathlib import Path
from datetime import datetime
import sys

# =====================
# CONFIGURATION
# =====================
INPUT_FILE = "sales_data.csv"
OUTPUT_FILE = "sales_report.xlsx"


# =====================
# UTILITY FUNCTIONS
# =====================

def read_sales_data(file_path: str) -> pd.DataFrame:
    """
    Reads the sales CSV file into a pandas DataFrame.
    Validates columns and handles missing/invalid data.
    """
    required_columns = {"Date", "Product", "Quantity", "Unit Price", "Salesperson"}
    try:
        df = pd.read_csv(file_path)
    except FileNotFoundError:
        sys.exit(f"❌ Error: File '{file_path}' not found.")
    except Exception as e:
        sys.exit(f"❌ Error reading CSV file: {e}")

    # Validate column presence
    if not required_columns.issubset(df.columns):
        missing = required_columns - set(df.columns)
        sys.exit(f"❌ Missing required columns in CSV: {', '.join(missing)}")

    # Handle invalid dates
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    invalid_dates = df["Date"].isna().sum()
    if invalid_dates > 0:
        print(f"⚠ Warning: {invalid_dates} rows have invalid or missing dates and were removed.")
        df = df.dropna(subset=["Date"])

    # Clean Quantity and Unit Price
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df["Unit Price"] = pd.to_numeric(df["Unit Price"], errors="coerce")
    invalid_price_qty = df["Quantity"].isna().sum() + df["Unit Price"].isna().sum()
    if invalid_price_qty > 0:
        print(f"⚠ Warning: {invalid_price_qty} rows had invalid/missing Quantity or Unit Price and were removed.")
        df = df.dropna(subset=["Quantity", "Unit Price"])

    # Calculate total amount for each row
    df["Total Sale"] = df["Quantity"] * df["Unit Price"]

    return df


def generate_product_sales(df: pd.DataFrame) -> pd.DataFrame:
    """
    Returns total sales per product.
    """
    return df.groupby("Product", as_index=False)["Total Sale"].sum().sort_values(by="Total Sale", ascending=False)


def generate_salesperson_sales(df: pd.DataFrame) -> pd.DataFrame:
    """
    Returns total sales per salesperson.
    """
    return df.groupby("Salesperson", as_index=False)["Total Sale"].sum().sort_values(by="Total Sale", ascending=False)


def generate_monthly_sales(df: pd.DataFrame) -> pd.DataFrame:
    """
    Returns monthly sales summary.
    """
    df["YearMonth"] = df["Date"].dt.to_period("M")
    monthly = df.groupby("YearMonth", as_index=False)["Total Sale"].sum().sort_values(by="YearMonth")
    monthly["YearMonth"] = monthly["YearMonth"].astype(str)
    return monthly


def generate_ytd_sales(df: pd.DataFrame) -> pd.DataFrame:
    """
    Returns year-to-date sales total by year.
    """
    df["Year"] = df["Date"].dt.year
    return df.groupby("Year", as_index=False)["Total Sale"].sum().sort_values(by="Year")


def save_to_excel(product_sales, salesperson_sales, monthly_sales, ytd_sales, output_file: str):
    """
    Saves all DataFrames to one Excel file with separate sheets.
    """
    try:
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            product_sales.to_excel(writer, sheet_name="Product Sales", index=False)
            salesperson_sales.to_excel(writer, sheet_name="Salesperson Sales", index=False)
            monthly_sales.to_excel(writer, sheet_name="Monthly Sales", index=False)
            ytd_sales.to_excel(writer, sheet_name="YTD Sales", index=False)
        print(f"✅ Excel report saved as '{output_file}'")
    except Exception as e:
        sys.exit(f"❌ Error writing Excel file: {e}")


def print_console_table(title: str, df: pd.DataFrame):
    """
    Prints a dataframe as a formatted console table.
    """
    print(f"\n=== {title} ===")
    print(df.to_string(index=False, justify="left"))


# =====================
# MAIN SCRIPT
# =====================

def main():
    print("📊 Generating Sales Report...")

    # Step 1: Read Data
    sales_df = read_sales_data(INPUT_FILE)
    if sales_df.empty:
        sys.exit("❌ No valid data to process.")

    # Step 2: Generate Summaries
    product_sales = generate_product_sales(sales_df)
    salesperson_sales = generate_salesperson_sales(sales_df)
    monthly_sales = generate_monthly_sales(sales_df)
    ytd_sales = generate_ytd_sales(sales_df)

    # Step 3: Print results to console
    print_console_table("Total Sales per Product", product_sales)
    print_console_table("Total Sales per Salesperson", salesperson_sales)
    print_console_table("Monthly Sales Summary", monthly_sales)
    print_console_table("Year-to-Date Sales", ytd_sales)

    # Step 4: Save to Excel
    save_to_excel(product_sales, salesperson_sales, monthly_sales, ytd_sales, OUTPUT_FILE)

    print("\n🎯 Sales Report generation completed successfully!")


if _name_ == "_main_":
    main()