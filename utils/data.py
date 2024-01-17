import pandas as pd
import logging


def clean_data(excel_data: dict) -> pd.DataFrame:
    logging.info("Cleaning data")
    output_df = pd.DataFrame(columns=['Employee Name', 'Date', 'Hours Worked', 'Comments', 'Weekly Period'], index=range(7))
    output_df['Employee Name'] = excel_data['employee_name']
    output_df['Weekly Period'] = excel_data['weekly_period']

    for day in range(7):
        output_df['Hours Worked'].iloc[day] = excel_data['project_hours'][day]
        output_df['Date'].iloc[day] = excel_data['project_hours_date'][day]
        output_df['Comments'].iloc[day] = excel_data['comments'][day]

    # Replace nan with empty string
    output_df = output_df.fillna('')

    # Strip comments
    output_df['Comments'] = output_df['Comments'].str.strip()

    # Replace \n with '; '
    output_df['Comments'] = output_df['Comments'].str.replace('\n', '; ')
    logging.debug(f'Cleaned data:\n{output_df}')
    return output_df
