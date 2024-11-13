# ExcelSheetProtector

A TypeScript script for automating cell protection in Excel sheets using Office Scripts. This script selectively locks cells with content, iterates through all sheets in the workbook, and applies batch processing for optimal performance.

## Features

- **Automated Protection**: Locks only cells containing content, leaving empty cells editable.
- **Batch Processing**: Processes cells in batches to improve performance and avoid memory issues.
- **Sheet-Level Protection**: Automatically unprotects and re-protects sheets (if not password-protected) to apply changes.

## Usage

This script is intended for use in Office Scripts, which can be run in Excel for the web. To use this script:

1. **Open Excel for the Web** and navigate to the workbook you want to process.
2. **Go to Automate > Code Editor** and paste the code into the editor.
3. **Run the Script** to process all sheets in the workbook.

## Example

To test the script, add some sample data in any worksheet in your workbook and run the script from Excel's **Automate** tab. The script will lock all cells with content in each sheet, but will skip any password-protected sheets.

## Notes

- This script is designed for Office Scripts in **Excel for the Web**.
- This script can also be used in a **Power Automate flow** to automate protection of worksheets.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
