import pandas as pd
import logging
import openpyxl
from .config import (EMPLOYEE_NAME_CELL, WEEKLY_PERIOD_CELL_RANGE, PROJECT_HOUR_CELL_RANGE,
                     COMMENT_CELL_RANGE, EMPLOYEE_NAME_BACKUP_CELL)


def read_excel_data(file_key: str) -> dict[str, str]:
    # Read the Excel file
    logging.info(f"\n\nReading Excel file: {file_key}")
    df_full = pd.read_excel(file_key, engine='openpyxl')
    logging.debug(df_full)

    # Get employee name
    logging.debug(f'Getting employee name from cell {EMPLOYEE_NAME_CELL}')
    employee_name = df_full.iloc[EMPLOYEE_NAME_CELL[1], EMPLOYEE_NAME_CELL[0]]
    logging.debug(f'Retrieved employee name: {employee_name}')
    if pd.isna(employee_name) or employee_name == 'Consultant Name':
        employee_name = ''

    # Get employee name backup
    logging.debug(f'Getting employee name backup from cell {EMPLOYEE_NAME_BACKUP_CELL}')
    employee_name_backup = df_full.iloc[EMPLOYEE_NAME_BACKUP_CELL[1], EMPLOYEE_NAME_BACKUP_CELL[0]]
    logging.debug(f'Retrieved employee name backup: {employee_name_backup}')
    if pd.isna(employee_name_backup) or employee_name_backup == 'Consultant Name':
        employee_name_backup = ''

    employee_name = employee_name + employee_name_backup

    if employee_name is None:
        error_msg = f"Employee Name not found in cell {EMPLOYEE_NAME_CELL}"
        logging.exception(error_msg)
        raise Exception(error_msg)
    logging.info(f"Employee Name: {employee_name}")

    # Get weekly period data
    weekly_period_start = df_full.iloc[WEEKLY_PERIOD_CELL_RANGE[1], WEEKLY_PERIOD_CELL_RANGE[0]].strftime("%m/%d/%Y")
    weekly_period_end = df_full.iloc[WEEKLY_PERIOD_CELL_RANGE[3], WEEKLY_PERIOD_CELL_RANGE[2]].strftime("%m/%d/%Y")
    weekly_period = f'{weekly_period_start} - {weekly_period_end}'
    if weekly_period_start is None or weekly_period_end is None:
        error_msg = f"Weekly Period not found in cells {WEEKLY_PERIOD_CELL_RANGE}"
        logging.exception(error_msg)
        raise Exception(error_msg)
    logging.info(f"Weekly Period: {weekly_period}")

    # Get Project Hour data
    project_hours, project_hours_date = [], []
    for day in range(7):
        project_hours.append(df_full.iloc[PROJECT_HOUR_CELL_RANGE[1], PROJECT_HOUR_CELL_RANGE[0] + day])
        project_hours_date.append(df_full.iloc[PROJECT_HOUR_CELL_RANGE[1] - 6, PROJECT_HOUR_CELL_RANGE[0] + day].strftime("%m/%d/%Y"))
    if len(project_hours) != 7:
        error_msg = f"Project Hours not found in cells {PROJECT_HOUR_CELL_RANGE}"
        logging.exception(error_msg)
        raise Exception(error_msg)
    if len(project_hours_date) != 7:
        error_msg = f"Project Hours Date not found in cells {PROJECT_HOUR_CELL_RANGE}"
        logging.exception(error_msg)
        raise Exception(error_msg)
    logging.info(f"Project Hours: {project_hours}\nProject Hours Date: {project_hours_date}")

    # Get Comment data
    comments = []
    for day in range(7):
        logging.debug(f"Getting comment for day {day}")

        comment_backup = df_full.iloc[COMMENT_CELL_RANGE[1] - 1, COMMENT_CELL_RANGE[0] + day]
        logging.debug(f'The comment_backup is: {comment_backup}')  # Comment backup is row 20 in dataframe

        comment = df_full.iloc[COMMENT_CELL_RANGE[1], COMMENT_CELL_RANGE[0] + day]
        logging.debug(f'The comment is: {comment}')  # comment is row 21 in dataframe
        comment = str(comment).strip() if not pd.isna(comment) else ''

        comment_backup2 = df_full.iloc[COMMENT_CELL_RANGE[1] + 1, COMMENT_CELL_RANGE[0] + day]
        logging.debug(f'The comment_backup2 is: {comment_backup2}')  # Comment backup2 is row 22 in dataframe

        # Check if comment_backup is not NaN and append it to comment
        if not pd.isna(comment_backup):
            comment_backup = f'{comment_backup.strip()}\n'
            comment = f'{comment_backup}{comment}'.strip()

        # Check if comment_backup2 is not NaN and append it to comment
        if not pd.isna(comment_backup2):
            comment_backup2 = f'{comment_backup2.strip()}\n'
            # Append comment_backup2 at the beginning of the comment
            comment = f'{comment}{comment_backup2}'.strip()

        logging.debug(f"Comment for day {day}: {comment}")
        comments.append(comment)

    logging.info(f"Comments: {comments}")

    # Return the data
    return {
        "employee_name": employee_name,
        "weekly_period": weekly_period,
        "project_hours": project_hours,
        "project_hours_date": project_hours_date,
        "comments": comments
    }


def auto_adjust_column_width(file_key: str):
    # Open the Excel file
    logging.debug(f"Auto-adjusting column width for file: {file_key}")
    workbook = openpyxl.load_workbook(file_key)
    sheet = workbook.active  # Assumes you are working with the active sheet

    # Iterate through columns and adjust width
    for column in sheet.columns:
        max_length = 0
        column_letter = openpyxl.utils.get_column_letter(column[0].column)

        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:  # noqa E722
                pass

        adjusted_width = (max_length + 2)  # Add 2 for a little extra padding
        sheet.column_dimensions[column_letter].width = adjusted_width

    # Save the workbook
    workbook.save(file_key)
    logging.debug(f"Column width adjusted for file: {file_key}")
