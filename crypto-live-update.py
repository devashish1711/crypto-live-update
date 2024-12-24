import requests
import pandas as pd
import time
from openpyxl import load_workbook

# Function to fetch cryptocurrency data
def fetch_crypto_data():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        'vs_currency': 'usd',
        'order': 'market_cap_desc',
        'per_page': 50,
        'page': 1,
    }
    response = requests.get(url, params=params)
    if response.status_code == 200:
        data = response.json()
        cryptos = []
        for crypto in data:
            cryptos.append({
                'Name': crypto['name'],
                'Symbol': crypto['symbol'],
                'Current Price (USD)': crypto['current_price'],
                'Market Capitalization (USD)': crypto['market_cap'],
                '24h Trading Volume (USD)': crypto['total_volume'],
                '24h Price Change (%)': crypto['price_change_percentage_24h'],
            })
        return pd.DataFrame(cryptos)
    else:
        print("Failed to fetch data from API.")
        return None

# Function to save data to an Excel file
def save_to_excel(df, file_name="Top_50_Cryptos_Live.xlsx"):
    try:
        with pd.ExcelWriter(file_name, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, index=False, sheet_name='Live Data')
        print(f"Data updated in Excel file: {file_name}")
    except Exception as e:
        print(f"Error saving data to Excel: {e}")

# Function to analyze cryptocurrency data
def analyze_data(df):
    top_5_by_market_cap = df.nlargest(5, 'Market Capitalization (USD)')
    average_price = df['Current Price (USD)'].mean()
    highest_24h_change = df.loc[df['24h Price Change (%)'].idxmax()]
    lowest_24h_change = df.loc[df['24h Price Change (%)'].idxmin()]

    print("\n--- Analysis ---")
    print("Top 5 Cryptocurrencies by Market Cap:")
    print(top_5_by_market_cap[['Name', 'Market Capitalization (USD)']])

    print(f"\nAverage Price of Top 50 Cryptocurrencies: ${average_price:.2f}")

    print("\nCryptocurrency with the Highest 24h Price Change:")
    print(highest_24h_change[['Name', '24h Price Change (%)']])

    print("\nCryptocurrency with the Lowest 24h Price Change:")
    print(lowest_24h_change[['Name', '24h Price Change (%)']])

# Main function to fetch, save, and analyze data
def main():
    file_name = "Top_50_Cryptos_Live.xlsx"
    while True:
        crypto_data = fetch_crypto_data()
        if crypto_data is not None:
            save_to_excel(crypto_data, file_name)
            analyze_data(crypto_data)
        else:
            print("Retrying in 5 minutes...")
        
        # Wait for 5 minutes before fetching data again
        time.sleep(300)

# Run the script
if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nScript stopped by user.")
