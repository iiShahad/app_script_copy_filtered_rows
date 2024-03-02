//This script is used to copy rows with the selected value and paste it in another sheet
function copyRowsWithSelectedValue() {
  var ui = SpreadsheetApp.getUi();
  try {
    //prompt user to enter these source sheet, column letter and the value:
    var sourceSheetName = ui
      .prompt("Enter the source sheet name")
      .getResponseText();
    var columnLetter = ui
      .prompt("Enter the column letter where the value is located")
      .getResponseText();
    var desiredValue = ui.prompt("Enter the selected value").getResponseText();

    //get source sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = ss.getSheetByName(sourceSheetName);
    if (!sourceSheet) {
      ui.alert("Error: Please check the sheet names and try again.");
      return;
    }

    //create the output sheet
    var outputSheet = ss.insertSheet().setName("Output Sheet");

    //get column number
    var columnNumber =
      columnLetter.toUpperCase().charCodeAt(0) - "A".charCodeAt(0) + 1;

    //load sheet data
    var sourceData = sourceSheet.getDataRange().getValues();

    //filter data
    var filteredRows = sourceData.filter((row) => {
      return row[columnNumber - 1] === desiredValue;
    });

    //insert data into output sheet
    if (filteredRows.length > 0) {
      filteredRows.forEach((row) => {
        outputSheet.appendRow(row);
      });
      ui.alert(
        "Rows with the value " +
          desiredValue +
          " have been copied to Output Sheet."
      );
    } else {
      ui.alert("No rows with the value " + desiredValue + " found.");
    }
  } catch (e) {
    ui.alert(e);
  }
}
