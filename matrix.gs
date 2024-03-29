//Get the sheet
const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

// Initialize rowNames as an empty object
let rowNames = {};

// Initialize rowNumber to 1
let rowNumber = 1;

// Loop through each row in the first column, excluding the first row
while (true) {
    // Get the cell in the row
    const cell = sheet.getRange(rowNumber + 1, 1);

    // Get the background color of the cell
    const background = cell.getBackground();

    // If the background is blank, break the loop
    if (background !== "#38761d") {
      break;
    }

    // If the background is not blank, set rowNames[rowNumber + 10] to the value of the cell
    rowNames[rowNumber + 10] = cell.getValue();

    // Increment rowNumber
    rowNumber++;
}

Logger.log(rowNames)
//Min col is 2
const minCol = 2;

// Initialize maxCol to 1
let maxCol = 1;

// Loop through the first cell in every column excluding the first column
while (true) {
    // Get the first cell in the column
    const cell = sheet.getRange(1, maxCol + 1);

    // Get the background color of the cell
    const background = cell.getBackground();

    // If the background is blank, break the loop
    if (background !== "#38761d") {
      break;
    }

    // If the background is not blank, add 1 to maxCol
    maxCol++;
}

function matrix(e) {
  //If not the correct sheet, return
  if (e.range.getSheet().getName() != "Matrix Management") {
    return;
  }

  //Get the range object
  const range = e.range;

  //Get the row number
  const row = range.getRow();

  // Get all the values from rowNames
  const rowNameValues = Object.keys(rowNames);

  // Filter out non-numeric values
  const numericRowNameValues = rowNameValues.filter(value => !isNaN(value));

  // If there are no numeric values, throw an error
  if (numericRowNameValues.length === 0) {
    Logger.log("Not big enough row values");
    return;
  }

  // Find the smallest value
  const smallestRowName = Math.min(...numericRowNameValues);

  //If the row number is too small, return
  if (row < 12) {
    return;
  }

  //Get the column number
  const col = range.getColumn();

  //Set variables
  let rowName = undefined;

  //Value
  const value = range.getValue()

  //If value is blank, do nothing
  if (value == '') {return;}

  //If column is not valid, return
  if (!(minCol <= col && col <= maxCol)) {
    //Notify of error
    range.setValue("Not a valid column " + maxCol + " " + col + " " + minCol);

    //Clear
    range.setValue("")

    //Return 
    return;
  }

  range.setValue("Processing");

  try {
    //Get the value associated with the row from the rowNames dictionary
    rowName = rowNames[row - 1];

    //If rowName is undefined, throw an error
    if(rowName === undefined) {
      throw new Error("No rowName associated with this row");
    }
  } catch(error) {
    //Log the error message
    Logger.log(error.message);
  }

  // Get the associated row
  associatedRow = row - 10;

  // Get the cell at associatedRow, col
  var cell_main = range.getSheet().getRange(associatedRow, col);

  //Get the cell's value
  const cellVal = cell_main.getValue();

  /* Check if cell value is valid */  

  //Get acces to the other sheet
  const teamSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Current Members');

  // Get the number of rows in the other sheet
  const numRows = teamSheet.getLastRow();
  const numCols = teamSheet.getLastColumn();

  // Get the values from the second row to the last row in the first column
  const teamMembers = teamSheet.getRange(2, 1, numRows - 1).getValues();

  // Flatten the array
  const flatTeamMembers = [].concat(...teamMembers);

  // Check if cellVal is in flatTeamMembers
  if (!flatTeamMembers.includes(value)) {
    // If cellVal is not in flatTeamMembers, set the value of cell_main to "Invalid Name"
    range.setValue("Invalid Name");

    // Return
    return;
  }

  //Set real value
  let realValue = undefined;

  if (cellVal != '') {
    realValue = cellVal + ", " + value;
  } else {
    realValue = value;
  }

  if (cellVal.includes(value)) {
    range.setValue("");
    return;
  }

  // Set the value of the cell
  cell_main.setValue(realValue); //Make function to get value

  //Clear value
  range.setValue("")

  if (col == 3) {return;} // End function if inactive column

  // Loop through each row in the first column
  for (let i = 1; i <= numRows; i++) {
    // Get the cell in the row
    cell = teamSheet.getRange(i, 1);

    // If the cell's value matches the value variable
    if (cell.getValue() === value) {
      // Loop through each cell in the row
      for (let j = 1; j <= numCols; j++) {
        // Get the cell
        const cellInRow = teamSheet.getRange(i, j);

        // If the cell is empty
        if (!cellInRow.getValue()) {
          // Set the value of the cell to rowName
          cellInRow.setValue(rowName);

          // Break the loop
          break;
        } else if (cellInRow.getValue() == rowName) {
          break;
        }
      }

      // Break the loop
      break;
    }
  }

}