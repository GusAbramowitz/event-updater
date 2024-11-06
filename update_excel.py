import pandas as pd
import os

# Define the OneDrive path to the specific Excel file
one_drive_path = r"C:\Users\gusab\OneDrive\CompanyData\Constant Update Events Clear30.xlsx"

def update_excel(new_data):
    # Check if the file already exists
    if os.path.exists(one_drive_path):
        # If it exists, read the existing data
        existing_data = pd.read_excel(one_drive_path)
        # Combine existing data with new data, dropping duplicates to avoid repeated records
        combined_data = pd.concat([existing_data, new_data]).drop_duplicates()
    else:
        # If the file does not exist, the new data becomes the combined data
        combined_data = new_data

    # Save the combined data back to the OneDrive file
    combined_data.to_excel(one_drive_path, index=False)
    print(f"Data saved to OneDrive at {one_drive_path}")

# Example usage:
new_data = pd.DataFrame({
    'Event ID': [1, 2, 3],
    'Event Name': ['Event A', 'Event B', 'Event C'],
    'Date': ['2024-11-05', '2024-11-06', '2024-11-07']
})

# Call the function to update the Excel file in OneDrive
update_excel(new_data)

