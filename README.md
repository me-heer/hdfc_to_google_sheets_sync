# HDFC to Google Sheets Sync

Web scrape transactions from HDFC Netbanking, parse them and update Google Sheets.

## Dependencies

- This script uses [Chrome Webdriver](https://chromedriver.chromium.org/downloads) to web scrape transactions. For this, you need to have Chrome installed and a compatible version of Chrome Webdriver placed in the /driver directory. You can use any other Webdriver, though.
- A sample .env file is provided at the root of this repository, to run this script, you need to fill in the variables according to your setup. Make sure your bank username and password are not accidentally pushed to Github. (Note: The TRANSACTIONS_FILE_PATH directory you provide will be used to download the transactions file, and after successful execution, clear the directory. Please make sure the transactions directory is solely used for this script.)

## Usage

```
usage: main.py [-h] [-f FROM_DATE] [-t TO_DATE]

A script to web scrape HDFC Netbanking Transactions and add them to Google Sheets

options:
  -h, --help            show this help message and exit
  -f FROM_DATE, --from-date FROM_DATE
                        transactions will be fetched from this date, format= dd/mm/YYYY
  -t TO_DATE, --to-date TO_DATE
                        transactions will be fetched to this date, format=dd/mm/YYYY
```
