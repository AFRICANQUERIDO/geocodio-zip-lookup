from dotenv import load_dotenv
import os
import pandas as pd
import requests
import time
import json

# === LOAD API KEY ===
load_dotenv()
API_KEY = os.getenv("GEOCODIO_KEY")
BASE_URL = "https://api.geocod.io/v1.7/geocode"

INPUT_FILE = "Palm_test.xlsx"
OUTPUT_FILE = "Zips test.xlsx"
BATCH_SIZE = 10

def get_zip(address):
    """Fetch ZIP code for a single address using Geocod.io API with extra debug info."""
    params = {"q": address, "api_key": API_KEY}
    try:
        response = requests.get(BASE_URL, params=params, timeout=10)
        print(f"📡 Requesting: {response.url}", flush=True)
        print(f"↩ Status Code: {response.status_code}", flush=True)

        if response.status_code != 200:
            print(f"❌ API returned non-200 for '{address}': {response.text}", flush=True)
            return "N/A"

        data = response.json()

        # Debug: show full JSON if results are missing
        if "results" not in data or len(data["results"]) == 0:
            print(f"❌ No results for '{address}'. Full response:\n{json.dumps(data, indent=2)}", flush=True)
            return "N/A"

        return data["results"][0]["address_components"].get("zip", "N/A")

    except Exception as e:
        print(f"💥 Exception for '{address}': {e}", flush=True)
        return "N/A"

def main():
    print("🚀 Starting script...", flush=True)
    df = pd.read_excel(INPUT_FILE)

    if "Address" not in df.columns:
        raise ValueError("Excel file must have a column named 'Address'")

    zips = []
    total = len(df)

    for i, addr in enumerate(df["Address"], start=1):
        print(f"🔎 Processing {i}/{total}: {addr}", flush=True)
        zips.append(get_zip(addr))
        time.sleep(0.2)  # avoid hitting rate limits

        # Save progress every batch
        if i % BATCH_SIZE == 0:
            df_partial = df.copy()
            df_partial["ZIP"] = zips + [""] * (total - len(zips))
            df_partial.to_excel(OUTPUT_FILE, index=False)
            print(f"✅ Saved progress at row {i}/{total}", flush=True)
            time.sleep(5)

    # Final save
    df["ZIP"] = zips
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"🎉 Finished! ZIP codes saved to '{OUTPUT_FILE}'.")

if __name__ == "__main__":
    main()



