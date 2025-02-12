import requests
import pandas as pd
import openpyxl
import schedule
import time

def fetch_crypto_data():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        "vs_currency": "usd",
        "order": "market_cap_desc",
        "per_page": 50,
        "page": 1,
        "sparkline": False
    }
    response = requests.get(url, params=params)
    data = response.json()
    return data

def analyze_data(data):
    df = pd.DataFrame(data, columns=["name", "symbol", "current_price", "market_cap", "total_volume", "price_change_percentage_24h"])
    top_5 = df.nlargest(5, "market_cap")
    avg_price = df["current_price"].mean()
    highest_change = df.loc[df["price_change_percentage_24h"].idxmax()]
    lowest_change = df.loc[df["price_change_percentage_24h"].idxmin()]
    return df, top_5, avg_price, highest_change, lowest_change

def update_excel():
    data = fetch_crypto_data()
    df, top_5, avg_price, highest_change, lowest_change = analyze_data(data)
    file_path = "crypto_data.xlsx"
    
    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Live Data", index=False)
        top_5.to_excel(writer, sheet_name="Top 5", index=False)
        summary = pd.DataFrame({
            "Metric": ["Average Price", "Highest 24h Change", "Lowest 24h Change"],
            "Value": [avg_price, highest_change["price_change_percentage_24h"], lowest_change["price_change_percentage_24h"]]
        })
        summary.to_excel(writer, sheet_name="Analysis Summary", index=False)
    
    print("Excel updated!")

schedule.every(5).minutes.do(update_excel)

print("Script running... Press Ctrl+C to stop.")
while True:
    schedule.run_pending()
    time.sleep(1)

