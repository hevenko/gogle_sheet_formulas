function FINDREFERENCE(cellRef, rangeA1, sheetName) {
    var sheet;
    
    // Check if sheetName is provided; if not, use the active sheet
    if (sheetName && sheetName.trim() !== "") {
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        return "Sheet not found: " + sheetName; // Return an error if the sheet doesn't exist
      }
    } else {
      sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    }
  
    // Get the specified range within the sheet
   // Check if rangeA1 is a string (A1 notation) or a Range object
    var range;
    if (typeof rangeA1 === "string") {
      range = sheet.getRange(rangeA1); // Get the range by A1 notation
    } else {
      return "FINDREFERENCE: range parameter must be string"; // Return error if neither a string nor a Range object
    }
    
    var formulas = range.getFormulas(); // Get all formulas in the range
    var foundCells = [];
    var i = 0;
    var j = 0;
    var wtf = '';
  
    try {
  
      // Loop through the formulas in the range
      for (;i < formulas.length; i++) {
        wtf += '\n';
        for (; j < formulas[i].length; j++) {
          wtf += formulas[i][j] + ' '
          if (
            // formulas[i][j].includes && 
            formulas[i][j].indexOf(cellRef) != -1) {
            foundCells.push(range.getCell(i + 1, j + 1).getA1Notation());
          }
        }
      }
    } catch (e) {
        return "findreference: (i,j):" + i + "," + j + " typeof: " + (formulas[i][j].toString()) + ";formulas.lenght:" + formulas.length + " " + e;
    }
    if (foundCells.length > 0) {
      return foundCells.join(", "); // Return found cells as a comma-separated string
    } else {
      // return "No cells found"; // Return this if no cells contain the reference
      return (i - 1) + "," + (j - 1) + ' cellRef:' + cellRef + ' range:' + rangeA1 + ' sheetName:' + sheetName;
    }
  }
  
  
  function ranger(range) {
    return range[0][2]; 
  }
  
  function tester() {
    FINDREFERENCE('H4','G4:I4', 'Obračun')
  }