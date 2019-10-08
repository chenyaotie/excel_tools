import logging
def get_logger(file_name):
    logger = logging.getLogger(file_name)
    logger.setLevel(level = logging.INFO)
    handler = logging.FileHandler("log.log")
    handler.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    return logger

