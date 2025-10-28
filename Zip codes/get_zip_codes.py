
#import os
#API_KEY = os.getenv("GEOCODIO_API_KEY")
#if not API_KEY:
 #   raise ValueError("Set GEOCODIO_API_KEY environment variable before running the script.")


from dotenv import load_dotenv
import os

# Load variables from .env
load_dotenv()

#Access your API key 
API_KEY = os.getenv("GEOCODIO_KEY")
BASE_URL = "https://api.geocod.io/v1.7/geocode"


import pandas as pd
import requests
import time

# === CONFIGURATION ===
#INPUT_FILE = "addresses.xlsx"           # Input file name
#OUTPUT_FILE = "addresses_with_zips.xlsx"  # Output file name
INPUT_FILE = "Palm_test.xlsx"
OUTPUT_FILE = "Zips test.xlsx" 
BATCH_SIZE = 10                       # Number of rows per batch

# === FUNCTION TO GET ZIP FOR ONE ADDRESS ===
def get_zip(address):
    """Fetch ZIP code for a single address using Geocod.io API."""
    params = {"q": address, "api_key": API_KEY}
    try:
        response = requests.get(BASE_URL, params=params)
        data = response.json()
        return data["results"][0]["address_components"]["zip"]
    except Exception as e:
        print(f"Error for '{address}': {e}")
        return "N/A"

# === MAIN SCRIPT ===
def main():
    # 1. Load Excel file
    df = pd.read_excel(INPUT_FILE)

    if "Address" not in df.columns:
        raise ValueError("Excel file must have a column named 'Address'")

    zips = []
    total = len(df)

    # 2. Loop through data in batches
    for i, addr in enumerate(df["Address"], start=1):
        zips.append(get_zip(addr))
        time.sleep(0.2)  # Small delay to avoid rate limits

        # Every 100 entries, save progress and take a small break
        if i % BATCH_SIZE == 0:
            df_partial = df.copy()
            df_partial["ZIP"] = zips + [""] * (total - len(zips))
            df_partial.to_excel(OUTPUT_FILE, index=False)
            print(f"✅ Saved progress at row {i}/{total}")
            time.sleep(5)  # Short pause between batches

    # 3. Save the final output
    df["ZIP"] = zips
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"🎉 Finished! ZIP codes saved to '{OUTPUT_FILE}'.")

# === RUN SCRIPT ===
if __name__ == "__main__":
    main()
