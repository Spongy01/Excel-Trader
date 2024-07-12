import xlwings as xl
import os
import sys
from ..utils import config
import logging
import time
from dhan.src.excel.renderer import render_single_component

# Configure logger
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class ExcelHandler:
    def __init__(self):
        self.app = xl.apps.add()
        self.app.visible = True
        self.filename = config.filename
        self.workbook = self.create_excel_app()
        self.range_to_monitor = 'B4:B14'
        self.previous_values = {}
        self.workbookname = self.workbook.fullname

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

        ordermanagement_start = (2 + watchlist_data_points + 7, 2)
        ordermanagement_end = (ordermanagement_start[0], ordermanagement_start[1] + 9)

        openposition_start = (ordermanagement_start[0], ordermanagement_end[1] + 3)
        openposition_end = (ordermanagement_start[0], openposition_start[1] + 6)

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
        render_single_component(order_sheet, ordermanagement_start, ordermanagement_end, ["Order management"],
                                color=(204, 255, 204),
                                merge=True, align_center=True)
        render_single_component(order_sheet, (ordermanagement_start[0] + 1, ordermanagement_start[1]),
                                (ordermanagement_start[0] + 1, ordermanagement_start[1] + 9),
                                ["Symbol", "Buy/Sell",
                                 "Quantity", "Order Price",
                                 "Status", "Modify/Cancel",
                                 "Modification",
                                 "Trigger Price",
                                 "Limit Price", "Confirm?"], (204, 255, 204))

        # render open positions tab header
        render_single_component(order_sheet, openposition_start, openposition_end, ["Open Positions"],
                                color=(204, 255, 255),
                                merge=True, align_center=True)
        render_single_component(order_sheet, (openposition_start[0] + 1, openposition_start[1]),
                                (openposition_start[0] + 1, openposition_start[1] + 6),
                                ["Symbol", "Buy/Sell",
                                 "Quantity", "Average Price",
                                 "MTM \n(Profit/Loss)",
                                 "Partially\n(Quantity)",
                                 "Square Off All"], (204, 255, 255))

    def monitor_changes(self):
        sheet = self.workbook.sheets('Market')

        # Initialize previous values array if the file already has data
        current_values = sheet.range(self.range_to_monitor).value
        logger.info(f"Cell Data found: {current_values}")
        for cell_index, value in enumerate(current_values):
            cell_address = f'B{cell_index + 3}'
            self.previous_values[cell_address] = value

        while True:
            if self.workbookname not in [i.fullname for i in self.app.books]:
                logger.warning("Workbook is closed.")
                return
            current_values = sheet.range(self.range_to_monitor).value
            # logger.info(f"Values : {current_values}")
            for cell_index, value in enumerate(current_values):
                cell_address = f'B{cell_index + 3}'
                if cell_address in self.previous_values:
                    previous_value = self.previous_values[cell_address]
                    if value != previous_value:
                        print(f"Cell {cell_address} changed: {previous_value} -> {value}")
                        # Log or process the change here

                # Update previous values
                self.previous_values[cell_address] = value
                time.sleep(0.1)
