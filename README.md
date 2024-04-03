# geocode-with-google

# Automated Geocoding with Selenium and Pandas

This project automates the process of geocoding addresses using Google Maps through the Selenium WebDriver. It reads addresses from an Excel spreadsheet, queries Google Maps to retrieve the latitude and longitude for each address, and updates the spreadsheet with these coordinates. Additionally, the project utilizes a caching mechanism to speed up the process by avoiding repeated queries for the same address, and logs the process for monitoring and debugging purposes.

## Features

- **Selenium WebDriver**: Automates browser interaction to query Google Maps.
- **Pandas**: Manages and updates Excel spreadsheets.
- **Caching**: Reduces the number of queries to Google Maps by caching addresses.
- **Logging**: Provides a log file of the geocoding process for troubleshooting.

## Requirements

To run this project, you will need Python 3.6 or newer. All dependencies are listed in `requirements.txt`, including:

- Selenium
- Pandas
- Webdriver-Manager
- tqdm

## Setup

1. **Clone the repository**:
    ```
    git clone https://your-repository-url
    cd your-repository-directory
    ```

2. **Create a virtual environment** (optional, but recommended):
    ```
    python -m venv venv
    source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
    ```

3. **Install the requirements**:
    ```
    pip install -r requirements.txt
    ```

## Configuration

Before running the script, you must configure the following settings in the script:

- `INPUT_FILE`: Path to the input Excel file containing addresses.
- `OUTPUT_FILE`: Path to the output Excel file where results will be saved.
- `LOG_FILE`: Path to the log file for the geocoding process.
- `CACHE_FILE`: Path to the cache file to store geocoded addresses.
- `ADDRESS_COLUMN_NAMES`: List of column names in the input file that contain the address information.

## Running the Script

To start the geocoding process, execute:

python geocode-with-google-and-selenium.py


Ensure your input Excel file is correctly formatted and located at the path specified in `INPUT_FILE`. The script will update the `OUTPUT_FILE` with latitude, longitude, and a processed flag for each address.

## Contributing

Contributions to improve the script or address any issues are welcome. Please feel free to submit a pull request or open an issue in the repository.

## License

This project is open-source and available under the GNU General Public License (GPL) version 3, dated 29 June 2007. For more details, see the LICENSE file in the repository or visit [https://www.gnu.org/licenses/gpl-3.0.html](https://www.gnu.org/licenses/gpl-3.0.html).
