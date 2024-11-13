/**
 * Locks cells with content in the provided Excel sheet.
 * @param worksheet The Excel sheet to process.
 */ 
function main(workbook: ExcelScript.Workbook): void {
  // Get all worksheets in the workbook
  let sheets = workbook.getWorksheets();

  // Iterate through each sheet in the workbook
  for (let i = 0; i < sheets.length; i++) {
    let sheet = sheets[i];

    // Get protection settings for the sheet
    let protection = sheet.getProtection();

    console.log("Processing " + sheet.getName());

    // Check if the sheet is not password protected
    if (!protection.getIsPasswordProtected()) {
      // Unprotect the sheet to make changes
      protection.unprotect();
     
      // Unlock all cells in the sheet
      sheet.getRange().getFormat().getProtection().setLocked(false);

      // Process the sheet
      processSheet(sheet);

      // Protect the sheet again
      protection.protect();
    } else {
      // Skip the sheet if it's password protected
      console.log("-- Password protected. Skip sheet.");
    }
  }

  console.log("Processing done!");
}

/**
 * Processes the provided Excel sheet by extracting the used range,
 * getting values and formulas, and then locking cells with content.
 * @param worksheet The Excel sheet to process.
 */
function processSheet(worksheet: ExcelScript.Worksheet): void {
  // Get the used range of the sheet
  let usedRange: ExcelScript.Range | null = worksheet.getRange().getUsedRange(true);

  // Check if a used range is found
  if (!usedRange) {
    console.log("-- No used range found.");
    return;
  }

  // Get values and formulas from the used range
  let rangeValues: object[] = usedRange.getValues(); 
  let rangeFormulas: object[] = usedRange.getFormulas();
  console.log("-- Processing values and formulas");

  // Lock cells with content in the used range
  lockCellsWithContent(usedRange, rangeValues, rangeFormulas);
}

/**
 * Locks cells in batches based on their content in the given Excel range.
 * @param usedRange The Excel range to process.
 * @param rangeValues The values in the Excel range.
 * @param rangeFormulas The formulas in the Excel range.
 */
function lockCellsWithContent(
  usedRange: ExcelScript.Range,
  rangeValues: string[],
  rangeFormulas: string[]
): void {
  // Define batch processing parameters
  const batchSize: number = 1000;
  const THRESHOLD: number = 10000;
  let bytes: number = 0;

  // Determine the total number of rows in the range
  const rangeLength: number = Math.max(rangeValues.length, rangeFormulas.length);

  // Process cells in batches
  for (let row: number = 0; row < rangeLength; row += batchSize) {
    let batchCellsToLock: ExcelScript.Range[] = [];

    // Collect cells to lock in the current batch
    for (
      let batchRow: number = row;
      batchRow < row + batchSize && batchRow < rangeLength;
      batchRow++
    ) {
      for (let col: number = 0; col < rangeValues[batchRow].length; col++) {
        // Check if the cell has content
        if (
          cellHasContent(
            rangeValues[batchRow][col] as string,
            rangeFormulas[batchRow][col] as string
          )
        ) {
          batchCellsToLock.push(usedRange.getCell(batchRow, col));
        }
      }

      // Update the byte count based on the processed values
      bytes += getByteSize(JSON.stringify(rangeValues[batchRow]));

      // Check if the byte threshold is reached
      if (bytes >= THRESHOLD) {
        console.log("---- Reached threshold - Processed until row " + String(batchRow));
        // Lock cells in the current batch
        lockCells(batchCellsToLock);
        batchCellsToLock = [];
        bytes = 0;
      }
    }

    // Lock cells in the final batch of the range
    lockCells(batchCellsToLock);
    batchCellsToLock = [];
    bytes = 0;
  }
}

/**
 * Checks if a cell has content based on its value and formula.
 * @param value The value of the cell.
 * @param formula The formula of the cell.
 * @returns True if the cell has content, false otherwise.
 */
function cellHasContent(value: string, formula: string): boolean {
  return value !== "" || formula !== "";
}

/**
 * Locks the provided Excel cells by setting their protection to locked.
 * @param cellsToLock The Excel cells to lock.
 */
function lockCells(cellsToLock: ExcelScript.Range[]): void {
  cellsToLock.forEach((cell) => {
    cell.getFormat().getProtection().setLocked(true);
  });
}


/**
 * Get the approximate byte size of a string.
 * @param str The string to calculate the byte size for.
 * @returns The byte size of the string.
 */
function getByteSize(str: string): number {
  return JSON.stringify(str).length
}
