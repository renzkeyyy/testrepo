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
                # Exactly matches the JSON camelCase spelling
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
        sheet = workbook['RawData'] if 'RawData' in workbook.sheetnames else workbook.active
            
        print("Workbook loaded. Updating watchlist...")
        today_str = datetime.now().strftime("%Y-%m-%d")

        for row in range(2, sheet.max_row + 1):
            cell_value = sheet.cell(row=row, column=3).value
            
            if cell_value:
                ticker_str = str(cell_value).strip()
                
                if ticker_str in live_data:
                    stock_info = live_data[ticker_str]
                    price = stock_info["price"]
                    pct_change = stock_info["change"]
                    
                    # Write Date (Col A), Price (Col D), and Percent Change (Col E)
                    sheet.cell(row=row, column=1, value=today_str)
                    sheet.cell(row=row, column=4, value=price)
                    sheet.cell(row=row, column=5, value=pct_change)
                    
                    print(f"Updated row {row}: {ticker_str} -> {price:.2f} PHP | {pct_change:.2f}%")
                else:
                    print(f"Skipped {ticker_str}: Not found in local database.")

        workbook.save(file_path)
        print("\nWorkbook updated and saved successfully!")

except PermissionError:
    print(f"Warning: '{file_path}' is open in Excel. Close it and try again.")
except Exception as e:
    print(f"An unexpected error occurred in Excel processing: {e}")