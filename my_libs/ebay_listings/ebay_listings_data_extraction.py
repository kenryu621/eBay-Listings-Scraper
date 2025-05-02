import logging
import os
import time
from typing import Any, Callable, Optional

import my_libs.utils as Utils
import my_libs.web_driver as Driver
import xlsxwriter
from my_libs.ebay_listings.ebay_listings_xlsx_writer import (
    MyEbayListingExcel,
    eBayListingData,
)
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

MAX_SCRAPE_ROW: int = 100


class KeywordScraper:
    """
    A class to scrape eBay product data for a given keyword and store the results in an Excel workbook.

    Attributes:
        driver (webdriver.Chrome): The Selenium WebDriver instance used for web scraping.
        keyword (str): The search keyword used to query eBay.
        workbook (MyEbayListingExcel): The workbook instance where data will be saved.
        total_worksheet (xlsxwriter.worksheet.Worksheet): Worksheet for storing data for all keywords.
        keyword_worksheet (xlsxwriter.worksheet.Worksheet): Worksheet specific to the current keyword.
        output_dir (str): Directory where images and the workbook will be saved.
        image_folder (str): Directory where downloaded images will be stored.
        row_count (int): Counter to keep track of the number of rows processed in the current worksheet.
    """

    def __init__(
        self,
        driver: webdriver.Chrome,
        keyword: str,
        workbook: MyEbayListingExcel,
        output_dir: str,
        listing_images_folder_path: str,
        screenshots_folder_path: str,
    ) -> None:
        self.driver: webdriver.Chrome = driver
        self.keyword: str = keyword
        self.workbook: MyEbayListingExcel = workbook
        self.total_worksheet: xlsxwriter.worksheet.Worksheet = workbook.total_worksheet
        self.keyword_worksheet: xlsxwriter.worksheet.Worksheet = workbook.new_worksheet(
            keyword
        )
        self.output_dir: str = output_dir
        self.listing_images_folder_path: str = listing_images_folder_path
        self.screenshots_folder_path: str = screenshots_folder_path
        self.processed_row_count = 0

    def scrape_keyword_data(self) -> None:
        """
        Scrape product data from eBay search results for the specified keyword.

        This method navigates to the eBay search results page for the keyword and processes data in pages
        until the maximum number of rows is reached or there are no more results.

        Returns:
            None: This method does not return a value. It updates the workbook with scraped data.
        """
        try:
            # Open eBay homepage
            url = Utils.build_ebay_search_url(self.keyword)
            logging.info(f"Navigating to URL: {url}")
            self.driver.get(url)
            Driver.check_ebay_captcha(self.driver, url)

            has_more_rows = True

            screenshot_path = os.path.join(
                self.screenshots_folder_path, f"{self.keyword}.png"
            )

            screenshot_taken = False

            while self.processed_row_count < MAX_SCRAPE_ROW and has_more_rows:
                try:
                    rows = self.fetch_data_rows()
                    if not screenshot_taken:
                        time.sleep(2)
                        screenshot_taken = Utils.take_screenshot(
                            screenshot_path, self.driver
                        )
                    if not rows:
                        logging.info(f"No results found for {self.keyword}")
                        break

                    self.process_rows_data(rows)

                    # Check if we have processed all rows in the current page
                    if len(rows) < MAX_SCRAPE_ROW:
                        logging.info(
                            f"Processed all rows for {self.keyword} on this page."
                        )
                        has_more_rows = False

                except Exception as e:
                    logging.error(
                        f"Error while processing rows for {self.keyword}: {e}"
                    )
                    break
        except Exception as e:
            logging.error(f"Error while scraping data for {self.keyword}: {e}")

    def fetch_data_rows(self) -> list[WebElement]:
        """
        Fetch data rows from the eBay search results page.

        This method waits for the search results to load and retrieves the elements representing each product
        listing.

        Returns:
            list[WebElement]: A list of Selenium WebElement instances representing the product rows.
        """
        try:
            logging.info(f"Fetching data for {self.keyword}...")
            # Wait until search results are loaded
            WebDriverWait(self.driver, 60).until(
                EC.presence_of_element_located((By.CLASS_NAME, "srp-river-results"))
            )

            rows = self.driver.find_elements(
                By.CSS_SELECTOR, "div.srp-river-results>ul>li>div.s-item__wrapper"
            )

            return rows
        except Exception as e:
            logging.error(
                f"An error occurred while fetching data for {self.keyword}: {e}"
            )
            return []

    def process_rows_data(self, rows: list[WebElement]) -> None:
        """
        Process and save data for each row retrieved from the eBay search results.

        This method parses each row to extract relevant product data, downloads images if available,
        and writes the data to the specified worksheets in the workbook.

        Args:
            rows (list[WebElement]): A list of WebElement instances representing the product rows.

        Returns:
            None: This method does not return a value. It updates the workbook with processed data.
        """
        for index, row in enumerate(rows, start=self.processed_row_count + 1):
            if self.processed_row_count >= MAX_SCRAPE_ROW:
                break

            data = self.parse_data_row(row)
            image_url = data.get(eBayListingData.IMAGE_URL)
            image_path = (
                Utils.download_image(
                    image_url,
                    self.listing_images_folder_path,
                    f"eBay Listing {self.keyword} {index}",
                )
                if image_url
                else None
            )
            data[eBayListingData.IMAGE_PATH] = image_path
            self.workbook.write_data_row(self.keyword_worksheet, data)
            self.processed_row_count += 1
            self.workbook.write_data_row(self.total_worksheet, data)

    def parse_data_row(self, row: WebElement) -> dict[eBayListingData, Any]:
        """
        Extract and parse product data from a single eBay search result row.

        This method extracts text and attributes from the row element, performs transformations if needed,
        and returns a dictionary containing the parsed data.

        Args:
            row (WebElement): A Selenium WebElement instance representing a single product row.

        Returns:
            dict[eBayListingDataKey, Any]: A dictionary with keys corresponding to eBayListingDataKey enum and
            values containing the extracted product data.
        """
        data: dict[eBayListingData, Any] = {}

        def safe_extract_text(
            selector, transform_func: Callable[[str], Any] = str
        ) -> Any:
            """
            Safely extract text from an element located by the CSS selector and transform it using a function.

            Args:
                selector: The CSS selector to locate the element.
                transform_func: Function to transform the extracted text.

            Returns:
                Any: Transformed text or None if extraction fails.
            """
            try:
                element = row.find_element(By.CSS_SELECTOR, selector)
                return transform_func(element.text)
            except (NoSuchElementException, ValueError):
                logging.debug("Failed to extract data for '%s'", selector)
                return None

        def safe_extract_attribute(selector, attribute_name) -> Any:
            """
            Safely extract an attribute value from an element located by the CSS selector.

            Args:
                selector: The CSS selector to locate the element.
                attribute_name: The attribute name to extract.

            Returns:
                Any: The attribute value or None if extraction fails.
            """
            try:
                element = row.find_element(By.CSS_SELECTOR, selector)
                return element.get_attribute(attribute_name)
            except NoSuchElementException:
                logging.debug(
                    "Failed to extract attribute '%s' for '%s'",
                    attribute_name,
                    selector,
                )
                return None

        def extract_item_id_from_href(href: Optional[str]) -> Optional[str]:
            if href:
                return href.split("/")[-1]
            return None

        data[eBayListingData.KEYWORD] = self.keyword
        title = safe_extract_text("div.s-item__title span")
        data[eBayListingData.TITLE] = Utils.escape_quotes(title)
        link = safe_extract_attribute("a.s-item__link", "href")
        if link:
            cleaned_url = Utils.ebay_clean_product_url(link)
            data[eBayListingData.TITLE_HREF] = cleaned_url
            data[eBayListingData.ITEM_ID] = extract_item_id_from_href(cleaned_url)
        seller_info = safe_extract_text("span.s-item__seller-info-text")
        if seller_info:
            seller_id = seller_info.split("(")[0].strip()
            data[eBayListingData.SELLER] = seller_id
            if len(seller_id) > 0:
                data[eBayListingData.SELLER_SEARCH_LINK] = (
                    Utils.build_seller_search_url(seller_id)
                )
        data[eBayListingData.IMAGE_URL] = safe_extract_attribute(
            "div.s-item__image-wrapper img", "src"
        )
        data[eBayListingData.PRICE] = safe_extract_text(
            "span.s-item__price",
            lambda text: float(text.replace(",", "").replace("$", "")),
        )
        data[eBayListingData.SHIPPING_COST] = safe_extract_text(
            "span.s-item__shipping"
        ) or safe_extract_text("span.s-item__freeXDays")

        return data


def process_keywords(keywords: list[str], output_dir: str) -> None:
    """
    Scrape product data for a list of keywords and save the results to the specified directory.

    This function initializes the web driver, creates an Excel workbook, and processes each keyword by
    creating a `KeywordScraper` instance and invoking the data scraping methods. It also ensures the
    creation and deletion of image folders.

    Args:
        keywords (list[str]): A list of keywords to search for.
        output_dir (str): The directory where the scraped data and images will be saved.

    Returns:
        None: This function does not return a value. It saves data and images to the specified directory.
    """
    if not keywords:
        logging.warning("No keywords provided. Skipping data fetch.")
        return

    logging.info("Fetching and saving data for keywords: %s", ", ".join(keywords))

    product_images_folder_path = Utils.create_subfolder(
        output_dir, "eBay Listing Images"
    )
    screenshot_folder_path = Utils.create_subfolder(
        output_dir, "eBay Listing Screenshots"
    )
    logging.info("Image folder created or ensured at: %s", product_images_folder_path)
    try:
        driver = Driver.initialize_driver(headless=True)

        workbook = MyEbayListingExcel("eBay Listings", output_dir)

        for keyword in keywords:
            keyword = keyword.strip()
            if keyword:
                scraper = KeywordScraper(
                    driver,
                    keyword,
                    workbook,
                    output_dir,
                    product_images_folder_path,
                    screenshot_folder_path,
                )
                scraper.scrape_keyword_data()
            else:
                logging.warning("Empty search keyword encountered. Skipping...")

        workbook.save_workbook()
        logging.info("All tasks completed successfully")

    except Exception as e:
        logging.error(f"Error encountered in eBay listing scraper: {e}")

    finally:
        if driver:
            Driver.close_driver(driver)
        Utils.delete_folder(product_images_folder_path)
