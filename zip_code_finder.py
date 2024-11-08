import os
import shutil
import csv
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut
import pandas as pd
import tkinter
from tkinter import filedialog

# pip install geopy
# pip install pandas openpyxl

def add_column_from_excel(source_file, column_name="ZIP Code"):
    """
    Reads a column from a source Excel file and adds it to a target Excel file.

    Parameters:
    source_file (str): Path to the source Excel file.
    column_name (str): The name of the column to add from the source file.
    """
    try:
        # copy the source file to output
        # Check if the file exists
        if not os.path.isfile(source_file):
            raise FileNotFoundError(f"The file {source_file} does not exist.")
        
        # Split the path into directory, filename, and extension
        directory, filename = os.path.split(source_file)
        name, ext = os.path.splitext(filename)
        
        # Create new filename with "_with_zips" appended
        new_filename = f"{name}_with_zips.csv"
        new_file_path = os.path.join(directory, new_filename)
        
        # Create the new file (copying the original for use later)
        shutil.copy(source_file, new_file_path)

        # Read the source Excel file
        df = pd.read_excel(new_file_path)
        address_list = []

        for row_index, row in df.iterrows():
            addr_obj = []
            address = row['Address']
            addr_obj.append(address)
            addr_obj.append(get_zip_code(address))
            print(str(addr_obj))
            address_list.append(addr_obj)

        with open(new_file_path, mode='w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(['Address', 'ZIP Code'])  # Write the header
            writer.writerows(address_list)  # Write the data row
        print(f"Column '{column_name}' added successfully to '{new_filename}'.")

    except Exception as e:
        print(f"An error occurred: {e}")


def get_zip_code(address):
    geolocator = Nominatim(user_agent="zip_code_lookup", timeout=5)

    try:
        # Geocode the address
        location = geolocator.geocode(address)

        if location:
            # Reverse geocode to get detailed address information
            location_details = geolocator.reverse((location.latitude, location.longitude), exactly_one=True)
            address_details = location_details.raw['address']

            # Extract the ZIP code
            if 'postcode' in address_details:
                return address_details['postcode']
            else:
                return 'ZIP code not found'
    except GeocoderTimedOut as e:
        print(f"Error: geocode failed on input {address} with message {e.message}")


if __name__ == "__main__":
    tkinter.Tk().withdraw() # prevents an empty tkinter window from appearing
    file_path = filedialog.askopenfile()

    rel_path = os.path.relpath(file_path.name, os.getcwd())
    print(rel_path)
    add_column_from_excel(rel_path)

    # Example address without zip code 
    # address = "2601 N MCMILLAN AVE Oklahoma City"
    # zip_code = get_zip_code(address)
    # print(f"The ZIP code for {address} is {zip_code}")
    
    # # Example address without zip code 
    # address = "2601 N MCMILLAN AVE OK"
    # zip_code = get_zip_code(address)
    # print(f"The ZIP code for {address} is {zip_code}")