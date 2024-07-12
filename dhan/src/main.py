# imports
import logging
import sys

from dhan.src.excel.excel_handler import ExcelHandler
from utils import utils, config

sys.path.append('../')


# Configure logger
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def main():
    logger.info("Check for Scrip File")
    utils.create_scrip_file(config.scrip)
    logger.info("Init Excel Object")
    handler = ExcelHandler()

    handler.monitor_changes()

    logger.info("Exiting..")


if __name__ == '__main__':
    main()
