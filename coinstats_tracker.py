# IMPORTS
import csv
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# URLs
url = 'https://coinstats.app/p/76ktZa'

# Selenium opening chrome without tabs
options = webdriver.ChromeOptions()
options.add_argument('--ignore-certificate-errors')
options.add_argument('--incognito')
options.add_argument('--headless')
s = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=s, options=options)

# CREDENTIALS
google_creds_filename = "Portfolio.json"

my_coinstats = {'currencies': [], 'total': 0}


def pull_cs_account_info():
    print("1. Gathering Coinstats portfolio information...")
    driver.get(url)

    # Obtain the number of rows in body
    rows = len(driver.find_elements(
        By.XPATH, "/html/body/div[1]/main/div/div/div[4]/table/tbody/tr"))

    for r in range(1, rows+1):
        values = []
        for p in range(2, 6):
            # obtaining the text from each column of the table
            values.append(driver.find_element(
                By.XPATH, "/html/body/div[1]/main/div/div/div[4]/table/tbody/tr["+str(r)+"]/td["+str(p)+"]").text)

        # Get current amount of currency and your totals
        currency_name = values[0]
        current_quantity = values[1].replace(",", "")
        current_price = values[2][1:].replace(",", "")

        currency_dict = {
            'symbol': currency_name,
            'quantity': current_quantity,
            'current_price': current_price,
            'owned': float(current_quantity) * float(current_price) * 0.89
        }

        if(currency_dict.get('quantity') != '0'):
            my_coinstats['currencies'].append(currency_dict)
            my_coinstats['total'] += currency_dict['owned']

    return my_coinstats


def connect_to_google_ss(google_creds_filename, ss_name):
    print("2. Connecting to Google Sheets...")

    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']

    credentials = (ServiceAccountCredentials
                   .from_json_keyfile_name(google_creds_filename,
                                           scope)
                   )
    gc = gspread.authorize(credentials)
    spreadsheet = gc.open(ss_name)
    return spreadsheet


def generate_portfolio_overview(my_coinstats, spreadsheet):
    # Fill first worksheet
    # open portfolio overview worksheet from file
    wks1 = spreadsheet.get_worksheet(4)

    # ADD PORTFOLIO OVERVIEW DETAILS INTO SPREADSHEET
    currency_count = len(my_coinstats['currencies'])

    currency_cell_list = wks1.range('A2:B' + str(2 + currency_count))
    last_cell = 0

    # Iterate over each currency
    for idx, currency in enumerate(my_coinstats['currencies']):
        cell = 0 + (idx*2)
        # Symbols
        currency_cell_list[cell].value = currency['symbol']
        cell += 1

        # Current Price
        currency_cell_list[cell].value = float(currency['current_price'])
        cell += 1

    print("3. Writing information to sheet 1...")
    # Update spreadsheet with currency overview
    wks1.update_cells(currency_cell_list)


# FULL CODE
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# Getting portfolio info
my_coinstats = pull_cs_account_info()
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# Connect to google spreadsheets and fill info
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
spreadsheet = connect_to_google_ss(google_creds_filename, "Capitale")
# Filling out first sheet, portfolio overview
generate_portfolio_overview(my_coinstats, spreadsheet)

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# Displaying results to user
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# Get spreadsheet url
spreadsheet_url = "https://docs.google.com/spreadsheets/d/%s" % spreadsheet.id
# Let user know process has been completed
print("\n=====Process completed, worksheets filled=====\n"
      "\nTo see results please visit:", spreadsheet_url)
