import utils.helper as hlp
import utils.excel as exl
import utils.config as cfg
import utils.data as dat
import time
import pandas as pd
import logging


def main():
    cfg.set_logger()

    # Convert the input string to a list
    invoice_file_keys = hlp.find_workbook_list()
    for invoice_file_key in invoice_file_keys:
        excel_data = exl.read_excel_data(invoice_file_key)

        # Clean data
        output_df = dat.clean_data(excel_data)
        if invoice_file_key == invoice_file_keys[0]:
            merged_df = output_df
        
        # Merge data
        merged_df = pd.concat([merged_df, output_df], ignore_index=True)

    logging.debug(f'Merged data:\n{merged_df}')

    # Export dataframe to Excel
    merged_df.to_excel(cfg.OUTPUT_FILE_KEY, index=False)
    
    # Auto-adjust column width
    exl.auto_adjust_column_width(cfg.OUTPUT_FILE_KEY)
    logging.info('Done!')


if __name__ == '__main__':
    main()
    time.sleep(1)  # Gives the end user time to read the message above
