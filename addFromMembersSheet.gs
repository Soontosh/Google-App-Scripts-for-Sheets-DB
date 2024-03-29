function addMembers(e) {
  // If not the correct sheet, return
  if (e.range.getSheet().getName() != "Current Members") {return;}

  // If the cell value is equal to '', return
  if ((e.range.getValue() == '')) {return;}

  // Make sure column is below 7, else, return
  if (e.range.getColumn() >= 7) {return;}

  //Get value of edited cell
  const value = e.range.getValue();


  // Get value of first column in row (name)
  const name = e.range.getSheet().getRange(e.range.getRow(), 1).getValue();

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
      rowNames[cell.getValue()] = rowNumber + 1;

      // Increment rowNumber
      rowNumber++;
  }

  // Check if value is in rowNames
  if (!(rowNames[value] != undefined)) {
    // If not, clear value of the cell
    //e.range.setValue("Invalid Value, Please Clear this Cell");  TEMPORARY FIX IMPLEMENTED BY STEVE TO ADD NEW MEMBERS
  } else {
    // Otherwise, get the cell of the first column of the row associated with the value as new_cell
    const new_cell = matrixSheet.getRange(rowNames[value], 2);

    // If name is already in new_cell, return
    if (new_cell.getValue().includes(name)) {
      return;
    }

    // If name is not already in new_cell, add ", " + name to the cell
    if (new_cell.getValue() === '') {
      new_cell.setValue(name);
    } else {
      new_cell.setValue(new_cell.getValue() + ", " + name);
    }
  }
}