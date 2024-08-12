import os
import sys
import pandas as pd
import xlwings as xl
from dhan.api.api.dhanhq import dhanhq
import logging
import yaml

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def get_credentials():
    credentials_file = 'credentials.yaml'

    # Check if credentials file exists
    if os.path.exists(credentials_file):
        with open(credentials_file, 'r') as file:
            try:
                credentials = yaml.safe_load(file)
                if 'client_id' in credentials and 'api_key' in credentials:
                    print("Credentials found in file.")
                    return credentials['client_id'], credentials['api_key']
                else:
                    print("Invalid credentials format in file. Asking user for new credentials.")
            except yaml.YAMLError as exc:
                print(f"Error loading YAML file: {exc}")

    # Ask user for credentials
    client_id = input("Enter your Client ID: ")
    api_key = input("Enter your API Key: ")

    # Save credentials to YAML file
    credentials = {
        'client_id': client_id,
        'api_key': api_key
    }

    with open(credentials_file, 'w') as file:
        try:
            yaml.safe_dump(credentials, file)
            print("Credentials saved to file.")
        except yaml.YAMLError as exc:
            print(f"Error saving YAML file: {exc}")

    return client_id, api_key


def create_scrip_file(filename):
    # filename = f'../{filename}' # go to parent directory
    scrip_app = xl.apps.add()
    if not os.path.exists(filename):
        print("Scrip File NOT FOUND.\nAdd api-scrip-master.csv from Dhan website to the current directory to proceed")
        input("Press any key to exit.")
        sys.exit()
    # Scrip file exists.
    df = pd.read_csv(filename, index_col=False, low_memory=False)

    column_a = "SEM_EXM_EXCH_ID"
    column_b = "SEM_CUSTOM_SYMBOL"
    column_c = "SEM_SMST_SECURITY_ID"
    column_d = "SEM_INSTRUMENT_NAME"
    new_column = "Watchlist Item"

    if new_column in df.columns:
        # watchlist scrip already exists dont calculate again
        print("Scrip File has Watchlist Column.")
        return

    print("Creating Watchlist Symbols...")
    # create watchlist item column
    df[new_column] = df[column_a] + "~" + df[column_d] + ':' + df[column_b] + "&" + df[column_c].astype(str)
    print("Dumping watchlist symbols in Scrip File...")

    workbook = scrip_app.books.open(fullname=filename)
    excelsheet = workbook.sheets[0]
    excelsheet.range('A1').value = df
    workbook.save()

    workbook.close()
    scrip_app.quit()
    print("Scrip File Updated.")


def get_segment(segment, trader):
    exch_id, instr_nm = segment.split('~', 1)
    if exch_id == 'NSE':
        if instr_nm == 'OPTIDX':
            return trader.NSE_FNO
        elif instr_nm == 'EQUITY':
            return trader.NSE
        else:
            logger.error("Segment Not in NSE BSE FNO CUR MCX")
            return None
    elif exch_id == 'BSE':
        if instr_nm == 'OPTIDX':
            return trader.BSE_FNO
        elif instr_nm == 'EQUITY':
            return trader.BSE
        else:
            logger.error("Segment Not in NSE BSE FNO CUR MCX")
            return None

    logger.error("Segment Not in NSE BSE FNO CUR MCX")
    return None


def get_buy_sell(buy_sell, trader):
    if buy_sell.lower() in ['+', 'b', 'buy']:
        buy_sell = trader.B
    elif buy_sell.lower() in ['-', 's', 'sell']:
        buy_sell = trader.S
    else:
        logger.error("Buy Sell data incorrect format")
        return None
    return buy_sell


def get_product_type(product_type, trader):

    if product_type is None or product_type == trader.MARGIN:
        product_type = trader.MARGIN
    elif product_type.lower in ['cnc', 'c']:
        product_type = trader.CNC
    elif product_type.lower in ['mis', 'm']:
        product_type = trader.INTRA
    else:
        logger.error("Product Type Not Intra or CNC or Margin")
        return None
    return product_type


def place_trade(instrument, values, trader: dhanhq):
    print("seg")
    segment, rest = instrument.split(':', 1)
    segment = get_segment(segment, trader)
    print("got Seg")
    if segment is None:
        return
    print("seq")
    security = int(rest.split('&')[-1])
    print("got seq")
    product_type = values[0]  # to change this
    product_type = get_product_type(product_type, trader)
    if product_type is None:
        return

    buy_sell = values[1]
    buy_sell = get_buy_sell(buy_sell, trader)
    if buy_sell is None:
        return

    quantity = int(values[2])

    trigger = int(values[3]) if values[3] is not None else None
    limit = int(values[4]) if values[4] is not None else 0

    order_type = trader.MARKET
    if limit != 0:
        order_type = trader.LIMIT

    logger.info(f"Trade Parameters : \n segment:{segment}\nsecurity:{security}\nbuy sell:{buy_sell}"
                f"order type:{order_type}\nproduct type:{product_type}\nquantity:{quantity}\n"
                f"limit price:{limit}\ntrigger price:{trigger}")

    if trigger is not None:
        resp = trader.place_order(security_id=security,
                                  exchange_segment=segment,
                                  transaction_type=buy_sell,
                                  quantity=quantity,
                                  order_type=order_type,
                                  product_type=product_type,  # to change this
                                  price=limit)
    else:
        resp = trader.place_order(security_id=security,
                                  exchange_segment=segment,
                                  transaction_type=buy_sell,
                                  quantity=quantity,
                                  order_type=order_type,
                                  product_type=product_type,  # to change this
                                  price=limit)
    logger.info(f"Trade Place Response : {resp}")
    return resp


def modify_cancel_trade(order_id, quant, values, trader: dhanhq):
    modify_cancel = values[0]
    if modify_cancel.lower() in ['-', 'c', 'cancel']:
        response = trader.cancel_order(order_id=order_id)
        return response

    if modify_cancel.lower() in ['+', 'm', 'modify']:
        quantity = int(values[1]) if values[1] is not None else quant
        trigger = int(values[2]) if values[2] is not None else None
        limit = int(values[3]) if values[3] is not None else 0

        order_type = trader.MARKET
        if limit != 0:
            order_type = trader.LIMIT

        logger.info(f"Trade Modification Parameters : \n Order ID:{int(order_id)}\nOrder_type:{order_type}\n"
                    f"quantity:{quantity}\n"
                    f"Price:{limit}\n")

        response = trader.modify_order(order_id=int(order_id), order_type=order_type, quantity=quantity, price=limit,
                                       validity=trader.DAY)
        return response

    return {'error': 'modify/cancel tab had invalid values'}


def get_positions_list(trader: dhanhq):
    positions_list = trader.get_positions()
    positions_list_status = positions_list['status']

    if positions_list_status != 'success':
        logger.error("Failed to generate Positions " + str(positions_list))
        return []

    positions_list = positions_list['data']

    compact_list = []
    utils_list = []
    for position in positions_list:
        trading_symbol = position['tradingSymbol']
        security_id = position['securityId']
        exchange_segment = position['exchangeSegment']
        product_type = position['productType']
        buyAvg = position['buyAvg']
        sellAvg = position['sellAvg']
        netQty = position['netQty']
        profit = position['unrealizedProfit'] if position['positionType'] != 'CLOSED' else position['realizedProfit']

        compact_list_data = [trading_symbol, buyAvg, sellAvg, netQty, profit]
        compact_list.append(compact_list_data)
        utils_list_data = [security_id, exchange_segment, product_type]
        utils_list.append(utils_list_data)

    return compact_list, utils_list


def get_order_list(trader: dhanhq):
    # logger.info("Getting order list . ")
    order_list = trader.get_order_list()
    order_list_status = order_list['status']

    if order_list_status != 'success':
        logger.error("Failed to generate order list " + str(order_list))
        return []

    # print("Order List : " + str(order_list))
    order_list = order_list['data']
    compact_list = []
    for order in order_list:
        order_id = order['orderId']
        order_trans_type = order['transactionType']
        order_symbol = order['tradingSymbol'] + " ; " + order['securityId']
        order_price = order['price']
        order_quantity = order['quantity']
        order_status = order['orderStatus']

        order_data = [order_id, order_symbol, order_trans_type, order_quantity, order_price, order_status]
        compact_list.append(order_data)

    # print("Compact List : " + str(compact_list))
    return compact_list
