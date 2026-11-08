import openpyxl
import requests
from datetime import datetime

file_path = "test_data.xlsx"
LOCAL_API_URL = "http://localhost:8080/stocks.json"

def fetch_all_local_prices():
    """Fetches the JSON wall from your local Java server and turns it into a fast dictionary."""
    try:
        response = requests.get(LOCAL_API_URL, timeout=5)
        
        if response.status_code == 200:
            data = response.json()
            # Phisix stores the list of companies inside a 'stocks' key
            stocks_list = data.get('stocks', [])
            
            # Create a clean dictionary mapped as { "URC": 120.5, "BDO": 150.0 }
            price_map = {}
            for stock in stocks_list:
                symbol = stock.get("symbol")
                price = stock.get("price", {}).get("amount")
                if symbol and price is not None:
                    price_map[symbol] = float(price)
                    
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

try:
    # 1. Grab all prices from your local server at once
    print("Connecting to local Java microservice...")
    live_prices = fetch_all_local_prices()
    
    if live_prices:
        # 2. Open the Excel file
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook['RawData'] if 'RawData' in workbook.sheetnames else workbook.active
            
        print("Workbook loaded. Updating watchlist...")
        today_str = datetime.now().strftime("%Y-%m-%d")

        # 3. Match the watchlist to the local data
        for row in range(2, sheet.max_row + 1):
            cell_value = sheet.cell(row=row, column=3).value
            
            if cell_value:
                ticker_str = str(cell_value).strip()
                
                # Instantly look up the price in our dictionary
                if ticker_str in live_prices:
                    price = live_prices[ticker_str]
                    sheet.cell(row=row, column=1, value=today_str)
                    sheet.cell(row=row, column=4, value=price)
                    print(f"Updated row {row}: {ticker_str} -> {price:.2f} PHP (via Local API)")
                else:
                    print(f"Skipped {ticker_str}: Not found in local database.")

        workbook.save(file_path)
        print("\nWorkbook updated and saved successfully!")

except PermissionError:
    print(f"Warning: '{file_path}' is open in Excel. Close it and try again.")
except Exception as e:
    print(f"An unexpected error occurred in Excel processing: {e}")