import xlwings as xl
import os
import sys
from ..utils import config
import logging
import time
from dhan.src.excel.renderer import render_single_component
from dhan.src.utils import utils

# Configure logger
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class ExcelHandler:
    def __init__(self, trader):
        self.app = xl.apps.add()
        self.app.visible = True
        self.filename = config.filename
        self.workbook = self.create_excel_app()
        self.range_to_monitor = 'B4:B14'
        self.previous_values = {}
        self.workbookname = self.workbook.fullname
        self.trader = trader

    def create_excel_app(self):
        logger.info("Starting a workbook Object")
        if not os.path.exists(self.filename):
            logger.warn("XL file does not exist, creating one")
            try:
                workbook = self.app.books.add()

                # add sheets if file does not exist
                for sheet in ["Order Sheet", "Market"]:
                    try:
                        workbook.sheets(sheet).clear()
                    except Exception as e:
                        workbook.sheets.add(sheet)

                try:
                    workbook.sheets("Sheet1").delete()
                except Exception as e:
                    pass

                # if workbook is new, render the template inside it
                logger.info("New Workbook - Creating Template for the workbook")
                self.render_excel_template(workbook)

                workbook.save(self.filename)
                workbook.close()
            except Exception as e:
                logger.error(f"Error while creating a workbook : {e}")
                sys.exit()

        workbook = self.app.books.open(self.filename)

        return workbook

    def render_excel_template(self, workbook):
        market_sheet = workbook.sheets("Market")
        order_sheet = workbook.sheets("Order Sheet")
        watchlist_data_points = 0  # get from an external file

        watchlist_start = (2, 2)
        watchlist_end = (2, 7)

        tradestation_start = (watchlist_end[0], watchlist_end[1] + 1)
        tradestation_end = (tradestation_start[0], tradestation_start[1] + 5)

        ordermanagement_start = (2, tradestation_end[1] + 3)
        ordermanagement_end = (ordermanagement_start[0], ordermanagement_start[1] + 9)

        openposition_start = (ordermanagement_start[0], ordermanagement_end[1] + 3)
        openposition_end = (ordermanagement_start[0], openposition_start[1] + 7)

        # render watchlist headers
        render_single_component(market_sheet, watchlist_start, watchlist_end, ["Watchlist"], color=(255, 204, 204),
                                merge=True, align_center=True)
        render_single_component(market_sheet, (watchlist_start[0] + 1, watchlist_start[0]),
                                (watchlist_start[0] + 1, watchlist_start[0] + 5),
                                ["Symbol", "Buy Quntity", "Buy Price", "Sell Price", "Sell Qauntity",
                                 "Last Trade Price"], color=(255, 204, 204))

        # render tradestation headers
        render_single_component(market_sheet, tradestation_start, tradestation_end, ["Trade Station"],
                                color=(204, 255, 204),
                                merge=True, align_center=True)
        render_single_component(market_sheet, (tradestation_start[0] + 1, tradestation_start[1]),
                                (tradestation_start[0] + 1, tradestation_start[1] + 5),
                                ["Order Type \n(MIS/CNC/Normal)",
                                 "Buy/Sell\n(+/-)", "Quantity",
                                 "Trigger \n Price",
                                 "Limit Price", "Confirm?"], color=(204, 255, 204))

        # render ordermanagement headers
        order_sheet = market_sheet
        render_single_component(order_sheet, ordermanagement_start, ordermanagement_end, ["Order management"],
                                color=(204, 255, 204),
                                merge=True, align_center=True)
        render_single_component(order_sheet, (ordermanagement_start[0] + 1, ordermanagement_start[1]),
                                (ordermanagement_start[0] + 1, ordermanagement_start[1] + 9),
                                ["Symbol", "Buy/Sell",
                                 "Quantity", "Order Price",
                                 "Status", "Modify/Cancel",
                                 "Modification Quantity",
                                 "Trigger Price",
                                 "Limit Price", "Confirm?"], (204, 255, 204))

        # render open positions tab header
        render_single_component(order_sheet, openposition_start, openposition_end, ["Open Positions"],
                                color=(204, 255, 255),
                                merge=True, align_center=True)
        render_single_component(order_sheet, (openposition_start[0] + 1, openposition_start[1]),
                                (openposition_start[0] + 1, openposition_start[1] + 7),
                                ["Symbol", "Buy Avg P", "Sell Avg P",
                                 "Net Quantity",
                                 "MTM \n(Profit/Loss)",
                                 "Partially\n(Quantity)", "limit price",
                                 "Square Off"], (204, 255, 255))

    def monitor_changes(self):
        sheet = self.workbook.sheets('Market')
        # order_sheet = self.workbook.sheets('Order Sheet')
        order_sheet = sheet
        # Initialize previous values array if the file alre
        # ady has data
        current_values = sheet.range(self.range_to_monitor).value
        logger.info(f"Cell Data found: {current_values}")
        for cell_index, value in enumerate(current_values):
            cell_address = f'B{cell_index + 4}'
            self.previous_values[cell_address] = value

        while True:
            if self.workbookname not in [i.fullname for i in self.app.books]:
                logger.warning("Workbook is closed.")
                return

            # Trade Station

            # if self.workbook.sheets.active == sheet:
            current_values = sheet.range(self.range_to_monitor).value
            # logger.info(f"Values : {current_values}")
            for cell_index, value in enumerate(current_values):
                cell_address = f'B{cell_index + 4}'
                if value is not None:
                    row_number = cell_index + 4
                    values_h_to_m = sheet.range(f'H{row_number}:M{row_number}').value
                    # logger.info(f"Values in H to M for row {row_number}: {values_h_to_m}")
                    if values_h_to_m[-1] is not None:
                        logger.info("Placing A Trade")
                        response = utils.place_trade(instrument=value, values=values_h_to_m, trader=self.trader)
                        sheet.range(f'H{row_number}:M{row_number}').value = None
                        sheet.range(f'N{row_number}').value = str(response)

                if cell_address in self.previous_values:
                    previous_value = self.previous_values[cell_address]
                    if value != previous_value:
                        print(f"Cell {cell_address} changed: {previous_value} -> {value}")
                        # Log or process the change here

                # Update previous values
                self.previous_values[cell_address] = value

            # Order Table

            order_list = utils.get_order_list(trader=self.trader)
            # logger.info(f"ORDER LIST :  {order_list}")
            order_sheet.range(f'O4').value = order_list
            length = len(order_list)
            order_range = f'O4:O{3 + length}'
            order_id_list = order_sheet.range(order_range).value
            for cell_index, order_id in enumerate(order_id_list):
                row_number = cell_index + 4
                color = (255, 255, 255)
                # print(f"ORDER SHEET VALUE : row{row_number} : Val: {order_sheet.range(f'T{row_number}')} ")
                if order_sheet.range(f'T{row_number}').value == 'TRADED':
                    if order_sheet.range(f'Q{row_number}').value == 'BUY':
                        color = (153, 255, 153)
                    elif order_sheet.range(f'Q{row_number}').value == 'SELL':
                        color = (255, 153, 153)
                    else:
                        color = (255, 102, 178)

                elif order_sheet.range(f'T{row_number}').value == 'REJECTED':
                    color = (224, 224, 224)
                elif order_sheet.range(f'T{row_number}').value == 'PENDING':
                    color = (255, 255, 153)

                order_sheet.range(f'P{row_number}:T{row_number}').color = color

                mod_request = order_sheet.range(f'U{row_number}:Y{row_number}').value
                if mod_request[-1] is not None:
                    logger.info(f"Modifying/Cancelling A Trade : {row_number}")

                    response = utils.modify_cancel_trade(order_id=order_id,
                                                         quant=order_sheet.range(f'R{row_number}').value,
                                                         values=mod_request, trader=self.trader)

                    order_sheet.range(f'U{row_number}:Y{row_number}').value = None
                    order_sheet.range(f'Z{row_number}').value = str(response)

            # Positions Table
            position_list, position_util_list = utils.get_positions_list(trader=self.trader)
            order_sheet.range(f'AB4').value = position_list
            length = len(position_list)
            for cell_index in range(0, length):
                row_number = cell_index + 4
                position_request = order_sheet.range(f'AG{row_number}:AI{row_number}').value

                limit_p = order_sheet.range(f'AG{row_number}').value

                square_off = position_request[2]
                partial_square_off = position_request[1]

                net_q = order_sheet.range(f'AE{row_number}').value
                logger.info(f"Position Request: {row_number} || Net_Q : {net_q}")
                buysell = "b"
                if net_q == 0:
                    pass
                elif net_q > 0:
                    buysell = 's'
                elif net_q < 0:
                    buysell = 'b'

                segment = "NSE~OPTIDX" if position_util_list[cell_index][1] == self.trader.NSE_FNO else "INVALID"
                security_id = str(position_util_list[cell_index][0])
                instrument = segment + ":" + "DUMMY" + "&" + security_id

                if partial_square_off is not None:
                    try:
                        quantity = int(partial_square_off)
                        data = [position_util_list[cell_index][-1], buysell, quantity, None, limit_p, "dummy"]
                        response = utils.place_trade(instrument=instrument, values=data, trader=self.trader)
                        sheet.range(f'AG{row_number}:AI{row_number}').value = None
                        sheet.range(f'AJ{row_number}').value = str(response)
                    except ValueError:
                        sheet.range(f'AG{row_number}:AI{row_number}').value = None
                        sheet.range(f'AJ{row_number}').value = "Partial Quantity was not an Integer"
                    continue

                if square_off is not None:


                    data = [position_util_list[cell_index][-1], buysell, abs(net_q), None, limit_p, "dummy"]

                    response = utils.place_trade(instrument=instrument, values=data, trader=self.trader)
                    sheet.range(f'AG{row_number}:AI{row_number}').value = None
                    sheet.range(f'AJ{row_number}').value = str(response)

            time.sleep(0.1)
