import os
import time
import pandas as pd
import requests
from dotenv import load_dotenv

# === Load API Key ===
load_dotenv()
API_KEY = os.getenv("GEOCODIO_KEY")

if not API_KEY:
    raise ValueError("❌ Missing API key. Please add GEOCODIO_KEY to your .env file.")

# === API Configuration ===
BASE_URL = "https://api.geocod.io/v1.7/geocode"

# === File Configuration ===
INPUT_FILE = "Zips-joe-properties_codes.xlsx"
OUTPUT_FILE = "Zips-joe-properties_final.xlsx"
BATCH_SIZE = 100  # Save progress every 100 rows
DEFAULT_CITY_STATE = ", Tampa, FL"  # Optional: to improve accuracy

# === Function to Get ZIP for One Address ===
def get_zip(address):
    """Fetch ZIP code for a single address using Geocod.io API."""
    # Append city/state if it’s missing “FL”
    if "FL" not in address:
        address = address + DEFAULT_CITY_STATE

    params = {"q": address, "api_key": API_KEY}

    try:
        response = requests.get(BASE_URL, params=params)
        response.raise_for_status()  # Raises an error for 4xx/5xx codes
        data = response.json()

        if "results" in data and len(data["results"]) > 0:
            return data["results"][0]["address_components"].get("zip", "N/A")
        else:
            print(f"⚠️ No results for: {address}")
            return "N/A"

    except requests.exceptions.RequestException as e:
        print(f"❌ Request error for '{address}': {e}")
        return "N/A"
    except Exception as e:
        print(f"⚠️ Unexpected error for '{address}': {e}")
        return "N/A"

# === Main Script ===
def main():
    print("🚀 Starting ZIP code extraction...")

    # 1. Load Excel
    df = pd.read_excel(INPUT_FILE)
    if "Address" not in df.columns:
        raise ValueError("❌ Excel file must have a column named 'Address'.")

    zips = []
    total = len(df)

    # 2. Loop through addresses
    for i, addr in enumerate(df["Address"], start=1):
        zip_code = get_zip(addr)
        zips.append(zip_code)
        time.sleep(0.3)  # Avoid API rate limit (Geocodio allows ~1 request/sec)

        # Save every batch
        if i % BATCH_SIZE == 0 or i == total:
            df_partial = df.copy()
            df_partial["ZIP"] = zips + [""] * (total - len(zips))
            df_partial.to_excel(OUTPUT_FILE, index=False)
            print(f"✅ Saved progress at row {i}/{total}")

    print(f"🎉 Done! All ZIP codes saved to {OUTPUT_FILE}")

# === Run ===
if __name__ == "__main__":
    main()
