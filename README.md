# Excel Comparison Application

This application allows you to compare two Excel files and generate a new Excel file based on the comparison.

## How to Use

1. Run the executable file to launch the application.
2. Click the "Select File 1" button to choose the first Excel file for comparison.
3. Click the "Select File 2" button to choose the second Excel file for comparison.
4. Select the comparison mode:
   - Click the "Update Database" button to compare the two Excel files and generate a new file that contains all data from both files.
   - Click the "Generate New Clients" button to compare the two Excel files and generate a new file that only contains new clients (clients that are present in File 2 but not in File 1).
5. If you want the comparison to be case-sensitive, check the "Case-sensitive comparison" checkbox. Otherwise, leave it unchecked.
6. The application will prompt you to save the output file. Choose a location and provide a name for the output file. The application will save the output file in this location.
7. You can view the status of the operation at the bottom of the application window. The application also shows a progress bar to indicate the progress of the operation.
8. To view the log of the operation, click the "Show Log" button.

## Notes

- The application can handle both `.xlsx` and `.csv` file formats.
- The application assumes that the first row of each Excel file contains the column headers.
- The application identifies clients based on the "Client nr" column. Make sure both Excel files have this column.
- The application copies the following columns to the output file: 'Client nr', 'Client', 'Address', 'NIP'. Ensure that these columns exist in both Excel files.

## Troubleshooting

If you encounter any issues while using the application, check the log file (`excel_comparison.log`) for error messages. If the issue persists, contact the application developer.
