import requests
import pandas as pd
import json
from typing import Dict, Any
from datetime import datetime
from update_excel import update_excel  # Import the function

# Configuration
API_URL = "https://quluipmdicjsolnsopkg.supabase.co/functions/v1/get_events"
API_KEY = "e5ae8df4-ef80-404c-afbc-0a8e37067e44"
AUTH_TOKEN = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InF1bHVpcG1kaWNqc29sbnNvcGtnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3MjQ2OTU4MzcsImV4cCI6MjA0MDI3MTgzN30.aXy2DJAU7pmR-yiniP57d0moGd-REDMlnWi8D4DkQgU"
PAGE_SIZE = 1000

headers = {
    "Authorization": f"Bearer {AUTH_TOKEN}",
    "x-api-key": API_KEY
}

def fetch_page(offset: int = 0, limit: int = PAGE_SIZE) -> Dict[str, Any]:
    """Fetch a single page of events"""
    params = {
        "offset": offset,
        "limit": limit
    }
    response = requests.get(API_URL, headers=headers, params=params)
    response.raise_for_status()
    return response.json()

def fetch_all_events():
    """Fetch all events using pagination and combine into a DataFrame"""
    print("Fetching first page and total count...")
    first_page = fetch_page()
    total_count = first_page['count']
    total_pages = (total_count + PAGE_SIZE - 1) // PAGE_SIZE

    print(f"Total records: {total_count}")
    print(f"Total pages: {total_pages}")

    all_data = first_page['data']

    for page in range(1, total_pages):
        offset = page * PAGE_SIZE
        print(f"Fetching page {page + 1} of {total_pages} (offset: {offset})...")
        try:
            response_data = fetch_page(offset)
            all_data.extend(response_data['data'])
        except Exception as e:
            print(f"Error fetching page {page + 1}: {str(e)}")
            continue

    return all_data

def process_events():
    """Fetch events and update the Excel file on OneDrive"""
    try:
        print("Starting data fetch...")
        all_events = fetch_all_events()

        print("Converting to DataFrame...")
        df = pd.DataFrame(all_events)

        # Convert 'event_extra_data' column to JSON strings if it exists
        if 'event_extra_data' in df.columns:
            df['event_extra_data'] = df['event_extra_data'].apply(json.dumps)

        # Update the Excel file in OneDrive
        print("Updating the Excel file in OneDrive...")
        update_excel(df)

        print(f"Update complete! Saved {len(df)} records to the Excel file.")

    except Exception as e:
        print(f"Error processing events: {str(e)}")

if __name__ == "__main__":
    process_events()
