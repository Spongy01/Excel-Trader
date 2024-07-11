import xlwings as xl
import os
import sys
from ..utils import config
import logging
import time

# Configure logger
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class ExcelHandler:
    def __init__(self):
        self.app = xl.apps.add()
        self.app.visible = True
        self.filename = config.filename
        self.workbook = self.create_excel_app()
        self.range_to_monitor = 'A1:A10'
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
                             
                workbook.save(self.filename)
                workbook.close()
            except Exception as e:
                logger.error(f"Error while creating a workbook : {e}")
                sys.exit()

        
        workbook = self.app.books.open(self.filename)
        
        return workbook


    def monitor_changes(self):
        sheet = self.workbook.sheets('Market')
        while True:
            if(self.workbookname not in [i.fullname for i in self.app.books]):
                logger.warning("Workbook is not open.")
                return
            current_values = sheet.range(self.range_to_monitor).value
            # logger.info(f"Values : {current_values}")
            for cell_index, value in enumerate(current_values):
                cell_address = f'A{cell_index + 1}'
                if cell_address in self.previous_values:
                    previous_value = self.previous_values[cell_address]
                    if value != previous_value:
                        print(f"Cell {cell_address} changed: {previous_value} -> {value}")
                        # Log or process the change here

                # Update previous values
                self.previous_values[cell_address] = value
                time.sleep(0.1)
