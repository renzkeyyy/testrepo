import openpyxl
import requests
from datetime import datetime

file_path = "test_data.xlsx"
LOCAL_API_URL = "http://localhost:8080/stocks.json"

def fetch_all_local_prices():
    """Fetches the JSON wall from your local Java server and extracts prices and percent changes."""
    try:
        response = requests.get(LOCAL_API_URL, timeout=5)
        
        if response.status_code == 200:
            data = response.json()
            stocks_list = data.get('stocks', [])
            
            price_map = {}
            for stock in stocks_list:
                symbol = stock.get("symbol")
                price = stock.get("price", {}).get("amount")
                percent_change = stock.get("percentChange") 
                
                if symbol and price is not None:
                    price_map[symbol] = {
                        "price": float(price),
                        "change": float(percent_change) if percent_change is not None else 0.0
                    }
                    
            return price_map
        else:
            print(f"Error: Local server returned status code {response.status_code}")
            return None
            
    except requests.exceptions.ConnectionError:
        print("CRITICAL ERROR: Could not connect to http://localhost:8080.")
        print("Is your Java Phisix server running in the other terminal?")
        return None
    except Exception as e:
        print(f"Unexpected error fetching local data: {e}")
        return None

# --- THE EXECUTION LOGIC ---
try:
    print("Connecting to local Java microservice...")
    live_data = fetch_all_local_prices()
    
    if live_data:
        workbook = openpyxl.load_workbook(file_path)
        raw_sheet = workbook['RawData'] if 'RawData' in workbook.sheetnames else workbook.active
        
        # Ensure History tab exists
        if 'History' not in workbook.sheetnames:
            history_sheet = workbook.create_sheet('History')
            history_sheet.append(["Date", "Company Name", "Ticker Symbol", "Last Traded price", "Percent Change"])
        else:
            history_sheet = workbook['History']
            
        print("Workbook loaded. Updating RawData and checking History...\n")
        today_str = datetime.now().strftime("%Y-%m-%d")

        # Create a memory bank of already logged (Date, Ticker) pairs to prevent same-day duplicates
        logged_entries = set()
        for row in range(2, history_sheet.max_row + 1):
            hist_date = history_sheet.cell(row=row, column=1).value
            hist_ticker = history_sheet.cell(row=row, column=3).value
            if hist_date and hist_ticker:
                logged_entries.add((str(hist_date).strip(), str(hist_ticker).strip()))

        # Process the watchlist
        for row in range(2, raw_sheet.max_row + 1):
            company_name = raw_sheet.cell(row=row, column=2).value
            cell_value = raw_sheet.cell(row=row, column=3).value
            
            if cell_value:
                ticker_str = str(cell_value).strip()
                
                if ticker_str in live_data:
                    stock_info = live_data[ticker_str]
                    price = stock_info["price"]
                    pct_change = stock_info["change"]
                    company_str = str(company_name) if company_name else ""
                    
                    # ACTION A: Always overwrite the daily snapshot in RawData
                    raw_sheet.cell(row=row, column=1, value=today_str)
                    raw_sheet.cell(row=row, column=4, value=price)
                    raw_sheet.cell(row=row, column=5, value=pct_change)
                    
                    # ACTION B: Append to History ONLY IF this ticker hasn't been logged today
                    if (today_str, ticker_str) not in logged_entries:
                        history_sheet.append([today_str, company_str, ticker_str, price, pct_change])
                        print(f"Logged {ticker_str}: {price:.2f} PHP (New Trading Day)")
                    else:
                        print(f"Updated RawData only (History already has {ticker_str} for {today_str})")
                else:
                    print(f"Skipped {ticker_str}: Not found in local database.")

        workbook.save(file_path)
        print("\nWorkbook updated and saved successfully!")

except PermissionError:
    print(f"Warning: '{file_path}' is open in Excel. Close it and try again.")
except Exception as e:
    print(f"An unexpected error occurred in Excel processing: {e}")