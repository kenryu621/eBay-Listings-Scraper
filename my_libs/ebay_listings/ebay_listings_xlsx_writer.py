import my_libs.utils as Utils
from my_libs.dependencies import *


class eBayListingData(Enum):
    """
    Enum class representing the various keys used for extracting data from eBay listing rows.

    Members:
        IMAGE_URL
        IMAGE_PATH
        KEYWORD
        TITLE
        TITLE_HREF
        SELLER
        SELLER_SEARCH_LINK
        ITEM_ID
        PRICE
    """

    IMAGE_URL = DataAttr()
    IMAGE_PATH = DataAttr(header="Image", column=0)
    KEYWORD = DataAttr(header="Keyword", column=1)
    TITLE = DataAttr(header="Title", column=2)
    TITLE_HREF = DataAttr()
    SELLER = DataAttr(header="Seller", column=5)
    SELLER_SEARCH_LINK = DataAttr()
    ITEM_ID = DataAttr(header="Item ID", column=6)
    PRICE = DataAttr(header="Listing Price", column=3)
    SHIPPING_COST = DataAttr(header="Shipping Cost", column=4)


class MyEbayListingExcel:
    """
    A class to manage and generate Excel files for eBay listing data.

    This class provides methods to create, save, and manage an Excel workbook with multiple worksheets. Each worksheet contains data related to eBay listings, including item details, images, and seller information.

    Attributes:
        workbook (xlsxwriter.Workbook): The Excel workbook instance.
        total_worksheet (xlsxwriter.worksheet.Worksheet): The worksheet for all keywords.
        formats (dict[FormatType, xlsxwriter.format.Format]): Dictionary of formats used in the workbook.
        row_counts (dict[str, int]): Dictionary tracking row counts for different worksheets.

    Methods:
        __init__(name: str, output_dir: str) -> None:
            Initialize an Excel workbook for storing eBay listing data.

        create_workbook(keyword: str, output_directory: str) -> xlsxwriter.Workbook:
            Create a new Excel workbook with a filename based on the provided keyword.

        save_workbook() -> None:
            Adjust column widths for all sheets and save the workbook.

        new_worksheet(keyword: str) -> xlsxwriter.worksheet.Worksheet:
            Add a new worksheet to the workbook and set up headers.

        add_headers(sheet: xlsxwriter.worksheet.Worksheet) -> None:
            Add headers to a specified worksheet.

        write_data_row(
            sheet: xlsxwriter.worksheet.Worksheet,
            data: dict[eBayListingData, Any]
        ) -> None:
            Add data, images, and hyperlinks to a specified row in the Excel sheet.
    """

    def __init__(self, name: str, output_dir: str) -> None:
        """
        Initialize an Excel workbook for storing eBay listing data.

        Args:
            name (str): The name to use for the Excel file.
            output_dir (str): The directory where the workbook will be saved.

        Attributes:
            workbook (xlsxwriter.Workbook): The Excel workbook instance.
            total_worksheet (xlsxwriter.worksheet.Worksheet): The worksheet for all keywords.
            formats (dict[FormatType, xlsxwriter.format.Format]): Dictionary of formats used in the workbook.
            row_counts (dict[str, int]): Dictionary tracking row counts for different keywords.
        """

        self.workbook: xlsxwriter.Workbook = self.create_workbook(name, output_dir)
        self.formats: dict[FormatType, xlsxwriter.format.Format] = initialize_formats(
            self.workbook
        )
        self.row_counts: dict[str, int] = {}
        self.total_worksheet: xlsxwriter.worksheet.Worksheet = self.new_worksheet(
            "All Keywords"
        )

    def create_workbook(
        self, keyword: str, output_directory: str
    ) -> xlsxwriter.Workbook:
        """
        Create a new Excel workbook with a filename based on the provided keyword.

        Args:
            keyword (str): The keyword used to name the Excel file.
            output_directory (str): The directory where the workbook will be saved.

        Returns:
            xlsxwriter.Workbook: The created Workbook instance.
        """

        output_file = os.path.join(output_directory, f"{keyword}.xlsx")
        workbook = xlsxwriter.Workbook(output_file)
        logging.info(f"Workbook created successfully at {output_file}")
        return workbook

    def save_workbook(self) -> None:
        """
        Adjust column widths for all sheets and save the workbook.

        Notes:
            - Autofits columns in all worksheets.
            - Sets the width for the first column to one-sixth of 100.
            - Saves and closes the workbook, handling any errors related to file permissions.
            - Logs a message indicating success or an error if the workbook cannot be saved.
            - Retries if the file is open elsewhere, prompting the user to close it and retry.
        """
        for sheet in self.workbook.worksheets():
            # Autofit columns
            sheet.autofit()
            # Set column width for the first column
            sheet.set_column(0, 0, 100 / 6)
        while True:
            try:
                self.workbook.close()
                logging.info("Workbook successfully saved.")
                break  # Exit the loop if the workbook is saved successfully
            except OSError as e:
                if e.errno == errno.EACCES:  # Permission denied error
                    logging.error(
                        "PermissionError: Please close the Excel file if it is open and press Enter to retry."
                    )
                    input("Please close the Excel file and press Enter to retry...")
                else:
                    logging.error(
                        "An OSError occurred while saving the workbook: %s", e
                    )
                    input("An unexpected error occurred. Press Enter to retry...")
            except Exception as e:
                logging.error(
                    "An unexpected error occurred while saving the workbook: %s", e
                )
                input("An unexpected error occurred. Press Enter to retry...")

    def new_worksheet(self, keyword: str) -> xlsxwriter.worksheet.Worksheet:
        """
        Add a new worksheet to the workbook and set up headers.

        Args:
            keyword (str): The name of the new worksheet.

        Returns:
            xlsxwriter.worksheet.Worksheet: The newly created worksheet.
        """

        new_worksheet = self.workbook.add_worksheet(keyword)
        self.add_headers(new_worksheet)
        self.row_counts[keyword] = 1
        return new_worksheet

    def add_headers(self, sheet: xlsxwriter.worksheet.Worksheet) -> None:
        """
        Add headers to a specified worksheet.

        Args:
            sheet (xlsxwriter.worksheet.Worksheet): The worksheet to add headers to.
        """
        headers_row = Utils.get_enum_headers_row(eBayListingData)
        sheet.write_row(
            0,
            Utils.get_enum_col(eBayListingData.IMAGE_PATH),
            headers_row,
            cell_format=self.formats[FormatType.HEADER],
        )

    def write_data_row(
        self,
        sheet: xlsxwriter.worksheet.Worksheet,
        data: dict[eBayListingData, Any],
    ) -> None:
        """
        Add data, images, and hyperlinks to a specified row in the Excel sheet.

        Args:
            sheet (xlsxwriter.worksheet.Worksheet): The Excel sheet where data will be written.
            data (dict[eBayListingDataKey, Any]): A dictionary containing the extracted product data.

        Notes:
            - Sets row height to 100.
            - Embeds an image if an image path is provided.
            - Writes title as a hyperlink if a link is available, otherwise as plain text.
            - Writes seller name or a search link if provided.
            - Writes item ID.
            - Updates the row index for the sheet.
        """

        sheet_name = sheet.get_name()

        if not sheet_name or sheet_name not in self.row_counts:
            logging.error(
                f"Sheet name '{sheet_name}' is missing or not found in row counts."
            )
            return

        row_idx = self.row_counts.get(sheet_name, 0)
        sheet.set_row(row_idx, 100)

        # Insert the image if an image path is provided
        image_path = data.get(eBayListingData.IMAGE_PATH)
        if image_path:
            try:
                sheet.embed_image(
                    row_idx, Utils.get_enum_col(eBayListingData.IMAGE_PATH), image_path
                )
            except Exception as e:
                logging.error(f"Error embedding image at row {row_idx}: {e}")

        fields_to_write = [
            (eBayListingData.KEYWORD, None),
            (eBayListingData.TITLE, eBayListingData.TITLE_HREF),
            (eBayListingData.PRICE, None),
            (eBayListingData.SELLER, eBayListingData.SELLER_SEARCH_LINK),
            (eBayListingData.ITEM_ID, None),
            (eBayListingData.SHIPPING_COST, None),
        ]

        for data_key, url_key in fields_to_write:
            check_genuine = data_key == eBayListingData.TITLE
            is_currency = data_key == eBayListingData.PRICE

            Utils.write_data(
                sheet,
                self.formats,
                row_idx,
                Utils.get_enum_col(data_key),
                data,
                data_key,
                url_key=url_key,
                check_genuine=check_genuine,
                is_currency=is_currency,
            )

        # Update row index
        self.row_counts[sheet_name] += 1
