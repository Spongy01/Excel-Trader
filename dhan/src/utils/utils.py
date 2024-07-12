import os
import sys
import pandas as pd
import xlwings as xl


def create_scrip_file(filename):
    # filename = f'../{filename}' # go to parent directory
    scrip_app = xl.apps.add()
    if not os.path.exists(filename):
        print("Scrip File NOT FOUND.\nAdd api-scrip-master.csv to the current directory to proceed")
        sys.exit()
    # Scrip file exists.
    df = pd.read_csv(filename, index_col=False, low_memory=False)

    column_a = "SEM_EXM_EXCH_ID"
    column_b = "SEM_CUSTOM_SYMBOL"
    column_c = "SEM_SMST_SECURITY_ID"
    new_column = "Watchlist Item"

    if new_column in df.columns:
        # watchlist scrip already exists dont calculate again
        print("Scrip File has Watchlist Column.")
        return

    print("Creating Watchlist Symbols...")
    # create watchlist item column
    df[new_column] = df[column_a] + ':' + df[column_b] + "&" + df[column_c].astype(str)
    print("Dumping watchlist symbols in Scrip File...")

    workbook = scrip_app.books.open(fullname=filename)
    excelsheet = workbook.sheets[0]
    excelsheet.range('A1').value = df
    workbook.save()

    workbook.close()
    scrip_app.quit()
    print("Scrip File Updated.")