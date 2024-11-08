import os
import shutil
import csv
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut
import pandas as pd
import tkinter
from tkinter import filedialog

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
            addr = get_address(address)
            addr_obj.append(addr['house_number'])
            addr_obj.append(addr['street'])
            addr_obj.append(addr['city'])
            addr_obj.append(addr['state'])
            addr_obj.append(addr['zip_code'])
            full_addr = f"{addr['house_number']} {addr['street']} {addr['city']}, {addr['state']} {addr['zip_code']}"
            addr_obj.append(full_addr)
            addr_obj.append(address.split(' ')[0] == str(addr['house_number']))

            print(str(addr_obj))
            address_list.append(addr_obj)

        with open(new_file_path, mode='w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(['Given Address', 'House Number', 'Street', 'City', 'State', 'ZIP Code', 'Full Address', 'Exact Match'])  # Write the header
            writer.writerows(address_list)  # Write the data row
        print(f"File with full addresses created.")

    except Exception as e:
        print(f"An error occurred: {e}")


def get_address(address):
    geolocator = Nominatim(user_agent="address_lookup", timeout=5)

    try:
        # Geocode the address to get latitude and longitude
        location = geolocator.geocode(address)

        if location:
            # Reverse geocode to get detailed address information
            location_details = geolocator.reverse((location.latitude, location.longitude), exactly_one=True)
            address_details = location_details.raw['address']

            # Extract address components
            house_number = address_details.get('house_number', 'House number not found')
            street = address_details.get('road', 'Street not found')
            city = address_details.get('city', address_details.get('town', 'City not found'))
            state = address_details.get('state', 'State not found')
            zip_code = address_details.get('postcode', 'ZIP code not found')

            # Return a dictionary with the full address details
            return {
                'house_number': house_number,
                'street': street,
                'city': city,
                'state': state,
                'zip_code': zip_code
            }
        else:
            return {
                'house_number': 'N/A',
                'street': 'N/A',
                'city': 'N/A',
                'state': 'N/A',
                'zip_code': 'N/A'
            }
    except GeocoderTimedOut as e:
        print(f"Error: geocode failed on input {address} with message {e}")
        return None


if __name__ == "__main__":
    tkinter.Tk().withdraw() # prevents an empty tkinter window from appearing
    file_path = filedialog.askopenfile()

    rel_path = os.path.relpath(file_path.name, os.getcwd())
    print(rel_path)
    add_column_from_excel(rel_path)