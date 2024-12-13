import requests
import pandas as pd
import time
from openpyxl import Workbook


API_URL = "https://api.coingecko.com/api/v3/coins/markets"
PARAMS = {
    "vs_currency": "usd",
    "order": "market_cap_desc",
    "per_page": 50,
    "page": 1,
    "sparkline": False
}
EXCEL_FILE = "crypto_data.xlsx"


def fetch_crypto_data():
    response = requests.get(API_URL, params=PARAMS)
    if response.status_code == 200:
        return response.json()
    else:
        print(f"Error fetching data: {response.status_code}")
        return []


def analyze_data(data):
    df = pd.DataFrame(data)


    top_5 = df.nlargest(5, "market_cap")[["name", "market_cap"]]

   
    avg_price = df["current_price"].mean()

    highest_change = df.loc[df["price_change_percentage_24h"].idxmax(), ["name", "price_change_percentage_24h"]]
    lowest_change = df.loc[df["price_change_percentage_24h"].idxmin(), ["name", "price_change_percentage_24h"]]

    analysis = {
        "Top 5 Market Cap": top_5,
        "Average Price": avg_price,
        "Highest Change": highest_change,
        "Lowest Change": lowest_change
    }
    return analysis


def write_to_excel(data, analysis):
    df = pd.DataFrame(data)[["name", "symbol", "current_price", "market_cap", "total_volume", "price_change_percentage_24h"]]
    
   
    writer = pd.ExcelWriter(EXCEL_FILE, engine="openpyxl")

   
    df.to_excel(writer, index=False, sheet_name="Live Data")

    
    analysis_df = pd.DataFrame({
        "Metric": ["Top 5 Market Cap", "Average Price", "Highest Change", "Lowest Change"],
        "Value": [
            analysis["Top 5 Market Cap"].to_string(index=False),
            f"${analysis['Average Price']:.2f}",
            f"{analysis['Highest Change']['name']} ({analysis['Highest Change']['price_change_percentage_24h']:.2f}%)",
            f"{analysis['Lowest Change']['name']} ({analysis['Lowest Change']['price_change_percentage_24h']:.2f}%)"
        ]
    })
    analysis_df.to_excel(writer, index=False, sheet_name="Analysis")

    writer.close()


def main():
    while True:
        print("Fetching data...")
        data = fetch_crypto_data()
        if data:
            print("Analyzing data...")
            analysis = analyze_data(data)
            print("Updating Excel...")
            write_to_excel(data, analysis)
            print("Excel updated. Waiting 5 minutes...")
        else:
            print("No data fetched. Retrying in 5 minutes...")

        time.sleep(300)  

if __name__ == "__main__":
    main()


