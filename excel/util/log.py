import logging
import sys
import datetime
reload(sys)
sys.setdefaultencoding('utf8')
def get_logger(file_name):
    logger = logging.getLogger(file_name)
    logger.setLevel(level = logging.INFO)
    log_file = 'sys_%s.log' % datetime.datetime.strftime(datetime.datetime.now(), '%Y-%m-%d')
    handler = logging.FileHandler(log_file)
    handler.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    return logger
