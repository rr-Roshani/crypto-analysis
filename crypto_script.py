import requests
import pandas as pd
import time
from openpyxl import load_workbook
from datetime import datetime

def fetch_crypto_data():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        "vs_currency": "usd",
        "order": "market_cap_desc",
        "per_page": 50,
        "page": 1,
        "sparkline": False
    }
    try:
        response = requests.get(url, params=params, timeout=10)
        response.raise_for_status()  # Raise error for bad responses (4xx, 5xx)
        return response.json()
    except requests.RequestException as e:
        print(f"⚠️ API Error: {e}")
        return None  # Return None to indicate failure

def save_to_excel(data, filename="crypto_data.xlsx"):
    df = pd.DataFrame(data, columns=["name", "symbol", "current_price", "market_cap", "total_volume", "price_change_percentage_24h"])
    df.rename(columns={
        "name": "Cryptocurrency Name",
        "symbol": "Symbol",
        "current_price": "Current Price (USD)",
        "market_cap": "Market Capitalization",
        "total_volume": "24h Trading Volume",
        "price_change_percentage_24h": "24h Price Change (%)"
    }, inplace=True)
    
    df["Timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # Add timestamp

    try:
        with pd.ExcelWriter(filename, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
            df.to_excel(writer, index=False, sheet_name="Live Data")
    except FileNotFoundError:
        df.to_excel(filename, index=False, sheet_name="Live Data")
    
def analyze_data(data):
    df = pd.DataFrame(data)
    
    if "market_cap" in df.columns:
        top_5 = df.nlargest(5, "market_cap")[["name", "market_cap"]]
    else:
        top_5 = pd.DataFrame()
    
    avg_price = df["current_price"].mean() if "current_price" in df.columns else 0
    
    if "price_change_percentage_24h" in df.columns:
        highest_change = df.nlargest(1, "price_change_percentage_24h")[["name", "price_change_percentage_24h"]]
        lowest_change = df.nsmallest(1, "price_change_percentage_24h")[["name", "price_change_percentage_24h"]]
    else:
        highest_change = lowest_change = pd.DataFrame()

    return top_5, avg_price, highest_change, lowest_change

def main():
    filename = "crypto_data.xlsx"
    while True:
        data = fetch_crypto_data()
        if data:
            save_to_excel(data, filename)
            top_5, avg_price, highest_change, lowest_change = analyze_data(data)
            print("\n🔹 Top 5 Cryptocurrencies by Market Cap:\n", top_5)
            print("\n💰 Average Price of Top 50 Cryptocurrencies:", avg_price)
            print("\n📈 Highest 24h % Change:\n", highest_change)
            print("\n📉 Lowest 24h % Change:\n", lowest_change)
        else:
            print("⚠️ Failed to fetch data, retrying in 5 minutes...")
        time.sleep(300)  # Wait 5 minutes before next update

if __name__ == "__main__":
    main()
