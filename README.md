# eBay Listing Scraper

A specialized web scraping tool designed to extract product listing data from the eBay website based on part numbers or keywords.

> **DISCLAIMER:** This project is shared for **EDUCATIONAL PURPOSES ONLY**. Please read the [Legal Disclaimer](#legal-disclaimer) section before using this code.

## Overview

This application automates the process of searching for products on eBay's website and extracts detailed information about each listing, including title, price, seller information, and more. The scraper exports all the collected data into a formatted Excel spreadsheet for easy analysis.

## Features

- **Keyword-Based Searching**: Search for products using specific part numbers or keywords from a text file
- **Multithreaded Processing**: Efficiently scrapes data using concurrent processing
- **Comprehensive Data Extraction**: Captures detailed listing information including:
  - Item title
  - Price
  - Seller information
  - Shipping cost
  - Item URL
  - Item images
- **Organized Output**: Exports data to a well-formatted Excel spreadsheet
- **Screenshot Capture**: Takes screenshots of search results for reference
- **Robust Error Handling**: Implements comprehensive exception handling and logging

## How to Use

### Prerequisites

- Windows OS
- Python 3.x (if running from source)
- Chrome browser

### Running the Application

#### From Source Code

1. Run `main.py`

### Setup Instructions

1. When first run, the application will create a `Keywords.txt` file
2. Add your part numbers or keywords to this file, one per line
   - Lines starting with `#` will be ignored (can be used for comments)
   - Example format:

     ```text
     # Enter your keywords below
     90916-03100
     15643-31050
     ```

3. Run the application again to start the scraping process
4. Results will be saved to an Excel file in the output directory

## Output

The scraper generates the following outputs:

1. **Excel Spreadsheet**: Contains all scraped data with the following columns:
   - Keyword
   - Title
   - Price
   - Shipping Cost
   - Seller Information
   - Item ID
   - Link to Item
   - Image Path

2. **Screenshots**: Captures of search result pages are saved in the "eBay Listings Screenshots" folder

3. **Log Files**: Detailed logs are stored in the "logs" folder for troubleshooting

## Project Structure

- `main.py`: Entry point of the application
- `my_libs/`: Contains the core functionality
  - `ebay_listings/`: eBay-specific scraping modules
    - `ebay_listings_data_extraction.py`: Extracts data from eBay pages
    - `ebay_listings_scrape.py`: Main scraping logic
    - `ebay_listings_xlsx_writer.py`: Formats and writes data to Excel
  - `logging_config.py`: Configures application logging
  - `utils.py`: Utility functions used throughout the application
  - `web_driver.py`: Manages Chrome WebDriver
  - `xlsxwriter_formats.py`: Excel formatting helpers

## Troubleshooting

- If the application fails to run, check the log files in the `logs` directory
- Ensure Chrome is installed on your system
- If the Excel file fails to save, make sure it's not already open in another application

## Legal Considerations

This tool is for personal use only. Please respect eBay's terms of service and use the tool responsibly with appropriate rate limiting to avoid overwhelming their servers.

## Legal Disclaimer

### IMPORTANT: READ BEFORE DOWNLOADING, COPYING, INSTALLING, OR USING

This software project is shared for **EDUCATIONAL PURPOSES ONLY** to demonstrate programming techniques for web automation and data extraction. By using, modifying, or distributing this code, you acknowledge and agree to the following:

1. **Terms of Service Compliance**: Most websites, including eBay, have Terms of Service that may prohibit automated data collection. Using this code to scrape websites may violate these terms.

2. **Personal Responsibility**: You are solely responsible for how you use this code. The author(s) of this project cannot be held liable for any misuse or legal consequences resulting from your use of this code.

3. **Rate Limiting**: If you choose to use this code, implement appropriate rate limiting to avoid overloading target websites' servers.

4. **Alternative API Usage**: Where available, consider using official APIs instead of web scraping.

5. **No Warranty**: This software is provided "AS IS" without warranty of any kind, express or implied.

6. **No Legal Advice**: This disclaimer is not legal advice. Consult with a legal professional if you have questions about the legality of web scraping in your jurisdiction.

Before using this code for any purpose, ensure you have the right to collect data from your target website, preferably by obtaining explicit permission.

The author(s) of this project disclaim any responsibility for how this code is used and any consequences that may arise from its use.
