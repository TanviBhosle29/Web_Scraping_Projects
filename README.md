# Web_Scraping_Projects
Web Scraping and Data Collection using Python

# Selenium Table Scraper & Excel Export
This project automates the extraction of tabular data from a complex web dashboard using Selenium, and exports the results into a single, well-organized Excel file. It is designed to handle dynamic dropdown date selection and scrape two visually distinct tables for each date.

---
## What Does It Do?
- **Automates Chrome:** Opens a web dashboard, navigates through language and section menus, and switches to the correct iframe.
- **Handles Dropdown Dates:** Iterates over each date option in a dropdown menu.
- **Scrapes Two Tables:** For every selected date, scrapes data from two separate tables (with different columns and structures) embedded in div-based containers.
- **Exports to Excel:** Collects all data and writes it to a single Excel file (`hihello_ALLDATES.xlsx`), with a separate sheet for each date and table (e.g., `2024-01-01_table1`, `2024-01-01_table2`).
- **Debug Friendly:** Prints detailed debug information at each step, so you can see exactly what was scraped and saved.

---
## Quick Start
1. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
2. **Edit the script:**
   - Open `selenium_scraper.py` and update the target URL and XPaths if your web dashboard structure is different.
   - Make sure the dropdown and table XPaths match your web page.
3. **Run the script:**
   ```bash
   python selenium_scraper.py
4. **Find your results:**
   - The Excel file will be saved to your Downloads folder as `hihello_ALLDATES.xlsx`.
   - Each sheet is named after the selected date and table (e.g. `2024-01-01_table1`).

---
## How the Code Works (Human-Friendly)
- **Setup:** Launches Chrome and navigates to your dashboard.
- **Iframe Navigation:** Switches into the correct iframe chain to access the data.
- **Dropdown Handling:** Clicks the dropdown, iterates through each date option, and selects it.
- **Table Scraping:**
  - For each date, finds the containers for Table 1 and Table 2.
  - Parses the text content, splitting into rows and columns based on known structure.
  - Handles cases where data is missing or the table is not present.
- **Excel Export:**
  - Collects all data into a dictionary.
  - At the end, writes all tables to Excel, one sheet per date and table.
  - Sheet names are human-readable and based on the dropdown date.
- **Debugging:**
  - Prints out what it finds at every step, so you can follow along or troubleshoot.

---
## Troubleshooting & Tips
- **ChromeDriver:** Make sure ChromeDriver is installed and in your PATH.
- **XPath Errors:** If you get errors finding elements, double-check your XPaths in the script.
- **Missing Data:** If a table is missing for a date, the script will skip that sheet without error.
- **Sheet Names:** Sheet names are based on the dropdown date. If the dropdown is empty, a timestamp is used instead.
- **Excel Issues:** If you get Excel errors, make sure you have `xlsxwriter` and `openpyxl` installed.
- **Dynamic Pages:** If the page is slow to load, adjust the sleep/wait times in the script.

---
## Dependencies
- selenium
- pandas
- openpyxl
- xlsxwriter

Install all with
```bash
pip install -r requirements.txt
---

## For Non-Experts 
- You do **not** need to know Python to use this! Just install the requirements, edit the script for your website, and run it.
- The script is heavily commented to help you understand or adapt it if needed.
- If you get stuck, check the debug output or ask for help!
