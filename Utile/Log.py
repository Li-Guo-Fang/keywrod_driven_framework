import logging.config
import logging

logging.config.fileConfig(r"D:\keywrod_driven_framework\Conf\Logger.conf")
logging = logging.getLogger("example01")

def debug(massage):
    logging.debug(massage)

def warning(massage):
    logging.warning(massage)

def info(massage):
    logging.info(massage)

def wrror(massage):
    logging.wrror(massage)
