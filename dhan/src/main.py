# imports
import logging
import os, sys

sys.path.append('../')
from dhan.src.excel.excel_handler import ExcelHandler

# Configure logger
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

if __name__ == '__main__':
    logger.info("Init Excel Object")
    handler = ExcelHandler()

    handler.monitor_changes()

    logger.info("Exiting..")
