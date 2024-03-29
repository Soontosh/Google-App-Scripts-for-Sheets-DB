function onEdit(e) {
  // If not the correct sheet, return
  if (e.range.getSheet().getName() != "Current Members") {matrix(e); return;}

  // If row edited was 1, return
  if (e.range.getRow() == 1) {return;}

  // If the cell value is not equal to '', return
  if (!(e.range.getValue() == '')) {
    addMembers(e)
    return;
  }

  /* Reorganize the row, move every cell after edited cell one cell to the left */
  const row = e.range.getRow();
  const col = e.range.getColumn();
  const sheet = e.range.getSheet();
  const numCols = sheet.getLastColumn();
  for (let i = col; i < numCols; i++) {
    const nextCell = sheet.getRange(row, i + 1);
    sheet.getRange(row, i).setValue(nextCell.getValue());
    nextCell.clearContent();
  }

  // Get list of all values in the row as teams_names
  const teams_names = sheet.getRange(row, 1, 1, numCols).getValues()[0];

  // Get value of first column in row (name)
  const name = sheet.getRange(row, 1).getValue();

  // Get "Matrix Management" sheet
  const matrixSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Matrix Management');

  /* Get all relevant rows in the sheet */
  // Initialize rowNames as an empty object
  let rowNames = {};

  // Initialize rowNumber to 1
  let rowNumber = 1;

  // Loop through each row in the first column, excluding the first row
  while (true) {
      // Get the cell in the row
      const cell = matrixSheet.getRange(rowNumber + 1, 1);

      // Get the background color of the cell
      const background = cell.getBackground();

      // If the background is blank, break the loop
      if (background !== "#38761d") {
        break;
      }

      // If the background is not blank, set rowNames[rowNumber + 10] to the value of the cell
      rowNames[rowNumber] = cell.getValue();

      // Increment rowNumber
      rowNumber++;
  }

  // Get all rows where row name is not in teams_names
  const invalidRows = Object.keys(rowNames).filter(row => !teams_names.includes(rowNames[row]));

  // Loop through each cell
  for (let row of invalidRows) {
    for (let i = 1; i <= numCols; i++) {
      const cell = matrixSheet.getRange(row, i);

      // For each cell, check if member's name is in cell
      if (cell.getValue().includes(name + ", ")) {
        // If so, remove name from the cell as well as the ', ' after it(if it is there)
        let newValue = cell.getValue().replace(name + ", ", '')
        cell.setValue(newValue);
      } else if (cell.getValue().includes(name)) {
        // If so, remove name from the cell as well as the ', ' after it(if it is there)
        let newValue = cell.getValue().replace(name, '')
        cell.setValue(newValue);
      }
    }
  }
}