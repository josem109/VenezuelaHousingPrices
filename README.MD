# Venezuela Housing Prices Scraper

This project is designed to scrape housing prices from MercadoLibre Venezuela using Selenium and save the data into an Excel file.

## Prerequisites

- Python 3.x
- Selenium
- openpyxl
- ChromeDriver

## Setup

1. **Clone the repository:**

   ```sh
   git clone https://github.com/yourusername/VenezuelaHousingPrices.git
   cd VenezuelaHousingPrices
   ```

2. **Install the required Python packages:**

   ```sh
   pip install selenium openpyxl
   ```

3. **Download ChromeDriver:**

   - Download the ChromeDriver that matches your version of Chrome from [here](https://sites.google.com/a/chromium.org/chromedriver/downloads).
   - Place the `chromedriver` executable in a directory that is in your system's PATH, or specify the path in the script.

4. **Prepare the Excel file:**
   - Create an Excel file named `master_state.xlsx` with at least two sheets: `Sheet1` and `Details`.
   - In `Sheet1`, column C should contain the URLs to be processed.

## Usage

1. **Run the scraper:**

   ```sh
   python scraping_mercadolibre.py
   ```

2. **The script will:**
   - Open the first URL from `master_state.xlsx` that hasn't been processed.
   - Scrape the housing links from the page.
   - Save the links to the `Details` sheet in `master_state.xlsx`.
   - Navigate through the pages using the "Next" button and repeat the process.

## Notes

- Ensure that the Excel file `master_state.xlsx` is in the same directory as the script.
- Adjust the sleep times (`time.sleep()`) if the pages take longer to load.

## License

This project is licensed under the MIT License.
