//creates the Sort Menu when you open the spreadsheet, and adds the different sort options
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Sort')
  .addItem('Sort By First Name', 'sortFirstName')
  .addItem('Sort By Last Name', 'sortLastName')
  .addItem('Sort By Town', 'sortTownName')
  .addToUi()
}


//Sorts the spreadsheet by alphabetical order based off the data in column A 
function sortFirstName() {
  let spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A:A').activate(); //the column it grabs the data from, change this if names are in a different column
  spreadsheet.getActiveSheet().sort(1, true);
};


function sortTownName() {
  let spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B:B').activate(); //the column it grabs the data from, change this if town names are in a different column
  spreadsheet.getActiveSheet().sort(2, true);
};


function sortLastName() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet();
  let range = sheet.getDataRange(); 
  let values = range.getValues();
  let headers = values.shift(); // Remove the headers


  //Find the last names, and the row they are on
  let lastNamesWithIndices = [];
  for (let i = 0; i < values.length; i++) {
    let value = values[i][0];
    if (value !== "") {
      let words = value.split(" ");
      let varLastName = words[words.length - 1];
      lastNamesWithIndices.push({ index: i, lastName: varLastName });
    }
  }


  // Sort by last name
  lastNamesWithIndices.sort((a, b) => {
    let lastNameA = a.lastName.toUpperCase();
    let lastNameB = b.lastName.toUpperCase();
    if (lastNameA < lastNameB) return -1;
    if (lastNameA > lastNameB) return 1;
    return 0;
  });


  // rearange the rows based off the last name sort
  let sortedValues = [];
  for (let i = 0; i < lastNamesWithIndices.length; i++) {
    let index = lastNamesWithIndices[i].index;
    sortedValues.push(values[index]);
  }


  //put the headers back
  sortedValues.unshift(headers);


  //send to the sheet
  range.setValues(sortedValues);
}
