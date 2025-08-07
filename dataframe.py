import pandas as pd
import os

def process_excel_sheet(excel_path, sheet_name): 
    try: 
        print(f"Reading sheet {sheet_name} from {os.path.basename(excel_path)}")
        
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=5, index_col=0)
        
        df.dropna(how='all', axis=0, inplace=True)
        df.dropna(how='all', axis=1, inplace=True)
        
        df.index.name = 'Chỉ tiêu'
        df.columns.name = 'Quý'  
        
        print(f"Successfully processed sheet {sheet_name}\n")
        return df

    except ValueError: 
        print(f"Error: Sheet {sheet_name} not found in {excel_path}. Please check the sheet name.")
        return None
    except Exception as e: 
        print(f"An unexpected error occurred while processing sheet {sheet_name}: {e}")
        return None

def analyze_financials(): 
    # 1. Enter valid Excel file path 
    while True: 
        file_path = input("Please enter the full path to your financial Excel file: ")
        # Check if the path exists and if the file has an .xlsx or .xls extension
        if os.path.exists(file_path) and file_path.lower().endswith(('.xlsx', '.xls')):
            print(f"File found: {file_path}!\n")
            return file_path  # Fixed: Added return statement
        else:
            print(f"File path {file_path} is not found! Please check again")

# 4. Display the head of each successfully loaded DataFrame
def get_balance_sheet(financial_dataframes):  # Fixed: Added parameter
    df_bs = financial_dataframes.get('balance_sheet')
    if df_bs is not None: 
        print("--- Balance Sheet (Cân đối kế toán) ---")
        print(df_bs.head(1000))
        print("\n" + "="*70 + "\n")
    else:
        print("Balance sheet data is not available.")

def get_income_statement(financial_dataframes):  # Fixed: Added parameter
    df_is = financial_dataframes.get('income_statement')
    if df_is is not None:
        print("--- Income Statement (Kết quả kinh doanh) ---")
        print(df_is.head(1000))
        print("\n" + "="*70 + "\n")
    else:
        print("Income statement data is not available.")

def get_financial_ratios(financial_dataframes):  # Fixed: Added parameter
    df_fr = financial_dataframes.get('financial_ratios')
    if df_fr is not None:
        print("--- Financial Ratios (Chỉ số tài chính) ---")
        print(df_fr.head(1000))
        print("\n" + "="*70 + "\n")
    else:
        print("Financial ratios data is not available.")

# Main execution
def main():
    # Get file path
    file_path = analyze_financials()
    
    # 2. Define sheet names 
    sheet_mapping = {
        'balance_sheet': 'CDKT',
        'income_statement': 'KQKD',
        'financial_ratios': 'CSTC'
    }

    # 3. Read data from each sheet into a dictionary of DataFrames
    financial_dataframes = {}
    for df_key, sheet_name in sheet_mapping.items(): 
        financial_dataframes[df_key] = process_excel_sheet(file_path, sheet_name)
    
    # 4. User selection for displaying data
    while True:
        input_num = input("Select sheet to get the DataFrame (1: Balance Sheet, 2: Income Statement, 3: Financial Ratios, 0: Exit): ")
        
        if input_num == '1':
            get_balance_sheet(financial_dataframes)
        elif input_num == '2':
            get_income_statement(financial_dataframes)
        elif input_num == '3':
            get_financial_ratios(financial_dataframes)
        elif input_num == '0':
            print("Exiting program.")
            break
        else:
            print("Invalid input. Please enter 1, 2, 3, or 0.")

# Run the main function if script is executed directly
if __name__ == "__main__":
    main()