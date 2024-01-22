# BillsAutomation v1.3.0

## Getting Started

Ensure to download the `dist.zip` file from the latest release.

1. **Prepare Invoice Workbook:**
   - Locate the `dist/input` directory on your computer.
   - Place your invoice workbooks (`.xlsx` files) into the `dist/input` directory.
   - Please ensure that all the necessary headers are included in the file.

2. **Run the Application:**
   - Navigate to the `dist` directory.
   - Double-click on the `.exe` file to start the processing of the invoice workbook.

## Post-Processing

1. **Check Output:**
   - After the application has run, open the `dist/output` directory.
   - Retrieve the processed invoice workbook from this directory.

2. **Review Log Files:**
   - Find `logfile.log` in the `dist/output` directory to review any operational logs for errors or confirmation of successful processing.
   - If you experience any issues, the log file may contain details that can help in troubleshooting.

## Notes

- The cell ranges are hard coded which means if the input format is not kept the same, exceptions and/or incorrect data will occur

## Requirements

- It's necessary to have excel installed prior to running the application.

## Developer Section

- Ensure the version is updated in `ReadMe.md` when creating new versions

## Future Development

- Introduce tests
