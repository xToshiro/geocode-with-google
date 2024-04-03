# Import necessary libraries
import pandas as pd
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager
import time
import json
import logging
from tqdm.auto import tqdm

# Configuration variables for file paths and column names
INPUT_FILE = "input.xlsx"  # Path to the input Excel file
OUTPUT_FILE = "output.xlsx"  # Path to the output Excel file
LOG_FILE = "geocoding.log"  # Path to the log file
CACHE_FILE = "geocoding_cache.json"  # Path to the cache file
ADDRESS_COLUMN_NAMES = ['Rua / Avenida ', 'Número ', 'Bairro ', 'CEP']  # Column names in the input file

# Setup logging to file with INFO level and specific format
logging.basicConfig(filename=LOG_FILE, level=logging.INFO,
                    format='%(asctime)s:%(levelname)s:%(message)s')

def setup_selenium_edge():
    """Sets up Selenium WebDriver for Edge browser with headless option."""
    service = Service(EdgeChromiumDriverManager().install())
    options = webdriver.EdgeOptions()
    options.add_argument("--headless")  # Run browser in headless mode for automation
    driver = webdriver.Edge(service=service, options=options)
    return driver

def check_and_prepare_output_file():
    """Checks if the output file exists, creates it from the input file if not, and adds necessary columns."""
    try:
        df_output = pd.read_excel(OUTPUT_FILE)
        logging.info("Output file found.")
    except FileNotFoundError:
        logging.info("Output file not found. Creating a new file.")
        df_output = pd.read_excel(INPUT_FILE)
        df_output['Latitude'] = None
        df_output['Longitude'] = None
        df_output['Processed'] = False
        df_output.to_excel(OUTPUT_FILE, index=False)
    return df_output

def load_cache():
    """Loads geocoding results from a cache file, creates a new cache if not found or corrupted."""
    try:
        with open(CACHE_FILE, 'r') as file:
            cache = json.load(file)
    except (FileNotFoundError, json.JSONDecodeError):
        logging.info("Cache file not found or corrupted. Creating a new one.")
        cache = {}
    return cache

def update_cache(cache):
    """Updates the cache file with new geocoding results."""
    with open(CACHE_FILE, 'w') as file:
        json.dump(cache, file)

def get_lat_long(driver, address, cache):
    """Retrieves latitude and longitude for an address using cache first, then Google Maps if not cached."""
    if address in cache:  # Check cache first
        logging.info(f"Address '{address}' found in cache.")
        return cache[address]['latitude'], cache[address]['longitude'], True
    try:
        # Perform Google Maps search
        driver.get(f"https://www.google.com/maps/search/{'+'.join(address.split())}")
        time.sleep(10)  # Wait for the page to load
        url = driver.current_url
        coords = url.split('@')[1].split(',')[0:2]
        latitude, longitude = coords[0], coords[1].split('!')[0]
        cache[address] = {'latitude': latitude, 'longitude': longitude}  # Update cache
        update_cache(cache)
        return latitude, longitude, True
    except Exception as e:
        logging.error(f"Error obtaining data for address '{address}': {e}")
        return None, None, False

# Main execution
df_output = check_and_prepare_output_file()
cache = load_cache()
driver = setup_selenium_edge()

# Process addresses not yet geocoded
for index, row in tqdm(df_output.iterrows(), total=df_output.shape[0], desc="Geocoding addresses"):
    if not row['Processed']:  # Skip already processed rows
        address = f"{row['Rua / Avenida ']} {row['Número ']} {row['Bairro ']} {row['CEP']}".strip()
        if address:
            latitude, longitude, success = get_lat_long(driver, address, cache)
            if success:
                df_output.at[index, 'Latitude'] = latitude
                df_output.at[index, 'Longitude'] = longitude
            df_output.at[index, 'Processed'] = success
            df_output.to_excel(OUTPUT_FILE, index=False)  # Save progress

driver.quit()  # Close the browser after completion
logging.info("Geocoding process completed.")

print("The spreadsheet has been updated with latitude and longitude coordinates.")
