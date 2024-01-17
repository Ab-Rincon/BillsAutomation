import os
import logging


def excel_column_to_index(column_letter: str) -> int:
    column_index = 0
    for i, char in enumerate(reversed(column_letter.upper())):
        column_index += (ord(char) - ord('A') + 1) * (26 ** i)
    return column_index - 1


EMPLOYEE_NAME_CELL = (excel_column_to_index("B"), 3 - 2)
WEEKLY_PERIOD_CELL_RANGE = (excel_column_to_index("C"), 8 - 2, excel_column_to_index("D"), 8 - 2)
PROJECT_HOUR_CELL_RANGE = (excel_column_to_index("B"), 18 - 2, excel_column_to_index("I"), 18 - 2)
COMMENT_CELL_RANGE = (excel_column_to_index("B"), 23 - 2, excel_column_to_index("I"), 23 - 2)

INPUT_DIRECTORY = 'input'
if not os.path.exists(INPUT_DIRECTORY):
    raise Exception(f'Please create an input directory and place the input files there. {INPUT_DIRECTORY}')

OUTPUT_DIRECTORY = 'output'
if not os.path.exists(OUTPUT_DIRECTORY):
    os.makedirs(OUTPUT_DIRECTORY)

LOG_FILE_NAME = 'logfile.log'
LOG_FILE_KEY = os.path.join(OUTPUT_DIRECTORY, LOG_FILE_NAME)

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

    logger.setLevel(logging.INFO)

    # Set the log format
    formatter = logging.Formatter("%(asctime)s %(levelname)s: %(message)s", datefmt="%Y-%m-%d %H:%M:%S")

    # Set up console logging
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    logger.addHandler(console_handler)

    # Set up file logging
    file_handler = logging.FileHandler(LOG_FILE_KEY, mode="w")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
