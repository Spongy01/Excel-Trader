# imports
import logging
import sys

import requests
import yaml

sys.path.append('../')
# from dhan.src.excel.excel_handler import ExcelHandler

from dhan.src.excel.excel_handler import ExcelHandler
from dhan.src.utils import utils, config

from dhan.api.api.dhanhq.dhanhq import dhanhq

# Configure logger
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def init_dhan():
    client_id, api_key = utils.get_credentials()
    dhan_object = dhanhq(client_id, api_key)
    return dhan_object


def main():
    logger.info("Initializing Dhan")

    # hereNSE~EQUITY:Biocon&11373

    trader = init_dhan()

    logger.info("Check for Scrip File")
    utils.create_scrip_file(config.scrip)
    logger.info("Init Excel Object")
    handler = ExcelHandler(trader)

    try:
        handler.monitor_changes()
    except Exception as e:
        logger.error(f"Handler Terminiated {e}")
        dumm = input("Enter something to Exit")
    logger.info("Exiting..")


if __name__ == '__main__':
    main()
