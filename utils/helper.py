import os
from .config import INPUT_DIRECTORY
import logging


def find_workbook_list() -> list:
    # Get a list of all files in the directory that do not start with "~"
    file_list = [f for f in os.listdir(INPUT_DIRECTORY) if
                 os.path.isfile(os.path.join(INPUT_DIRECTORY, f)) and not f.startswith('~') and f.endswith('.xlsx')]

    if file_list == []:
        error_msg = f'Please create an input directory and place the input files there. {INPUT_DIRECTORY}'
        logging.exception(error_msg)
        raise Exception(error_msg)

    # add the directory to the file name
    file_keys = [os.path.join(INPUT_DIRECTORY, f) for f in file_list]
    return file_keys
