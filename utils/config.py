import os
import logging


def excel_column_to_index(column_letter: str) -> int:
    column_index = 0
    for i, char in enumerate(reversed(column_letter.upper())):
        column_index += (ord(char) - ord('A') + 1) * (26 ** i)
    return column_index - 1


EMPLOYEE_NAME_CELL = (excel_column_to_index("B"), 3 - 2)
EMPLOYEE_NAME_BACKUP_CELL = (excel_column_to_index("C"), 3 - 2)
WEEKLY_PERIOD_CELL_RANGE = (excel_column_to_index("C"), 8 - 2, excel_column_to_index("D"), 8 - 2)
PROJECT_HOUR_CELL_RANGE = (excel_column_to_index("B"), 18 - 2, excel_column_to_index("I"), 18 - 2)
COMMENT_CELL_RANGE = (excel_column_to_index("B"), 23 - 2, excel_column_to_index("I"), 23 - 2)

INPUT_DIRECTORY = 'input'
if not os.path.exists(INPUT_DIRECTORY):
    error_msg = f'Please create an input directory and place the input files there. {INPUT_DIRECTORY}'
    logging.exception(error_msg)
    raise Exception(error_msg)

OUTPUT_DIRECTORY = 'output'
if not os.path.exists(OUTPUT_DIRECTORY):
    os.makedirs(OUTPUT_DIRECTORY)

INFO_LOG_FILENAME = 'info.log'
INFO_LOG_FILEKEY = os.path.join(OUTPUT_DIRECTORY, INFO_LOG_FILENAME)

DEBUG_LOG_FILENAME = 'debug.log'
DEBUG_LOG_FILEKEY = os.path.join(OUTPUT_DIRECTORY, DEBUG_LOG_FILENAME)

OUTPUT_FILE_NAME = 'output.xlsx'
OUTPUT_FILE_KEY = os.path.join(OUTPUT_DIRECTORY, OUTPUT_FILE_NAME)


def set_logger():
    """
    Configure the logger for logging messages.
    """
    # Get the root logger
    logger = logging.getLogger()
    while logger.hasHandlers():
        logger.removeHandler(logger.handlers[0])

    logger.setLevel(logging.DEBUG)

    # Set the log format
    logfmt = "%(asctime)s %(levelname)s: %(message)s"
    datefmt = "%Y-%m-%d %H:%M:%S"
    formatter = logging.Formatter(logfmt, datefmt)

    # Set up console logging (INFO level)
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # Set up debug file logging (DEBUG level)
    debug_file_handler = logging.FileHandler(DEBUG_LOG_FILEKEY, mode='w')
    debug_file_handler.setLevel(logging.DEBUG)
    debug_file_handler.setFormatter(formatter)
    logger.addHandler(debug_file_handler)

    # Set up error file logging (INFO level)
    error_file_handler = logging.FileHandler(INFO_LOG_FILEKEY, mode='w')
    error_file_handler.setLevel(logging.INFO)
    error_file_handler.setFormatter(formatter)
    logger.addHandler(error_file_handler)
