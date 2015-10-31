function sayHelloBrowser() {
    // Declare a string literal variable.
    var greeting = 'Hello world!';
    // Display a message dialog with the greeting 
    //(visible from the containing spreadsheet).
    Browser.msgBox(greeting);
}

function helloDocument() {
    var greeting = 'Hello world!';
    // Create DocumentApp instance.
    var doc = DocumentApp.create('test_DocumentApp');
    // Write the greeting to a Google document.
    doc.setText(greeting);
    // Close the newly created document
    doc.saveAndClose();
}

function helloLogger() {
    var greeting = 'Hello world!';
    //Write the greeting to a logging window.
    // This is visible from the script editor
    //   window menu "View->Logs...".
    Logger.log(greeting);
}

function helloSpreadsheet() {
    var greeting = 'Hello world!',
        sheet = SpreadsheetApp.getActiveSheet();
    // Post the greeting variable value to cell A1
    // of the active sheet in the containing 
    //  spreadsheet.
    sheet.getRange('A1').setValue(greeting);
    // Using the LanguageApp write the 
    //  greeting to cell:
    // A2 in Spanish, 
    //  cell A3 in German, 
    //  and cell A4 in French.
    sheet.getRange('A2')
        .setValue(LanguageApp.translate(
    greeting, 'en', 'es'));
    sheet.getRange('A3')
        .setValue(LanguageApp.translate(
    greeting, 'en', 'de'));
    sheet.getRange('A4')
        .setValue(LanguageApp.translate(
    greeting, 'en', 'fr'));
}


// Chapter 3

// Cannot be called as a UDF.
function setRangeFontBold(rangeAddress) {
    var sheet = SpreadsheetApp.getActiveSheet();
    sheet.getRange(rangeAddress)
        .setFontWeight('bold');
}
// Call "setRangeFontBold()" from editor.
function call_setCellFontBold() {
    var rangeAddress = Browser.inputBox(

        'Set Range Font Bold',
        'Provide a range address',
    Browser.Buttons.OK_CANCEL);
    if (rangeAddress) {
        setRangeFontBold(rangeAddress);
    }
}
// Given the standard deviation and the mean,
//  return the relative standard deviation.
function RSD(stdev, mean) {
    if (!(typeof stdev === 'number' && typeof mean === 'number')) {
        throw {
            'name': 'TypeError',
                'message':
                'Function "RSD()" requires ' +
                'two numeric arguments'
        };
    }
    return (100 * (stdev / mean)).toFixed(2) * 1;
}
// Given a temperature value in Celsius
//  return the temperature in Fahrenheit.
function celsiusToFahrenheit(celsius) {
    if (typeof celsius !== 'number') {
        throw {
            'name': 'TypeError',
                'message': 'Function requires ' +
                'a single number argument'
        };
    }
    return ((celsius * 9) / 5) + 32;
}
// Given a temperature in Fahrenheit,
// return the temperature in Celsius.
function fahrenheitToCelsius(fahrenheit) {
    if (typeof fahrenheit !== 'number') {
        throw {
            'name': 'TypeError',
                'message': 'Function requires ' +
                ' a single number argument'
        };
    }
    return (fahrenheit - 32) * 5 / 9;
}
// Given the radius, return the
// area.
// Throw an error if the radius is
// negative.
function areaOfCircle(radius) {
    if (typeof radius !== 'number') {
        throw {
            'name': 'TypeError',
                'message': 'Function requires ' +
                'a single numeric argument'
        };
    }
    if (radius < 0) {
        throw {
            'name': 'ValueError',
                'message': 'Radius myst ' +
                ' be non-negative'
        };
    }
    return Math.PI * (radius * radius);
}

function test_intervalInDays() {
    var date1 = new Date(),
        date2 = new Date(1972, 7, 17);
    Logger.log(intervalInDays(date1, date2));
}
// Write String methods to the logger.
function printStringMethods() {
    var strMethods = Object.getOwnPropertyNames(
    String.prototype);
    Logger.log('String has ' + strMethods.length +
        ' properties.');
    Logger.log(strMethods.sort().join('\n'));
}
// Reverse the alphabet.
function test_reverseString() {
    var str = 'abcdefghijklmnopqrstuvwxyz';
    Logger.log(reverseString(str));
}
// Return a string with the characters

// of the input string reversed.
function reverseString(str) {
    var strReversed = '',
        lastCharIndex = str.length - 1,
        i;
    if (typeof str !== 'string') {
        throw {
            'name': 'TypeError',
                'message': 'Function requires a ' +

                ' single string argument.'
        };
    }
    for (i = lastCharIndex; i >= 0; i -= 1) {
        strReversed += str[i];
    }
    return strReversed;
}
// Return a integer between
// 1 and 6 inclusive.
function throwDie() {
    return 1 + Math.floor(Math.random() * 6);
}
// Concatenate cell values from
// an input range.
// Single quotes around concatenated 
// elements are optional.
function concatRng(inputFromRng, concatStr,
addSingleQuotes) {
    var cellValues;
    if (addSingleQuotes) {
        cellValues = inputFromRng.map(

        function (element) {
            return "'" + element + "'";
        });
        return cellValues.join(concatStr);
    }
    return inputFromRng.join(concatStr);
}
// Print stockInfo object property
// names to the logger.
function printFinanceAppKeys() {
    stockSymbol = 'GOOG';
    Logger.log(Object.keys(
    FinanceApp.getStockInfo(
    stockSymbol))
        .sort()
        .join('\n'));
}
// Given a stock symbol, return the
//  stock price (NYSE).
function getStockPrice(stockSymbol) {
    return FinanceApp.getStockInfo(stockSymbol)['price'];
}
// Given a stock symbol, return the
//  full stock name. 
function getStockName(stockSymbol) {
    return FinanceApp.getStockInfo(stockSymbol)['name'];
}

// Chapter 4

// Function to demonstrate the Spreadsheet 
//  object hierarchy.
// All the variables are gathered in a 
//  JavaScript array.
// At each iteration of the for loop the 
//  "toString()" method 
//  is called for each variable and its
//   output is printed to the log.
function showGoogleSpreadsheetHierarchy() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        sh = ss.getActiveSheet(),
        rng = ss.getRange('A1:C10'),
        innerRng = rng.getCell(3, 3),
        innerRngAddress = innerRng.getA1Notation(),
        column = innerRngAddress.slice(0, 1),
        googleObjects = [ss, sh, rng, innerRng,
        innerRngAddress, column],
        i;
    for (i = 0; i < googleObjects.length; i += 1) {
        Logger.log(googleObjects[i].toString());
    }
}
// Print the column letter of the third row and 
//   third column of the range "A1:C10"
//  of the active sheet in the active 
//   spreadsheet.
// This is for demonstration purposes only!
function getColumnLetter() {
    Logger.log(
    SpreadsheetApp.getActiveSpreadsheet()
        .getActiveSheet().getRange('A1:C10')
        .getCell(3, 3).getA1Notation()
        .slice(0, 1));
}
// Extract an array of all the property names 
//  defined for Spreadsheet and write them to
//  column A of the active sheet in the active
//   spreadsheet.
function testSpreadsheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        sh = ss.getActiveSheet(),
        i,
        spreadsheetProperties = [],
        outputRngStart = sh.getRange('A2');
    sh.getRange('A1')
        .setValue('spreadsheet_properties');
    sh.getRange('A1')
        .setFontWeight('bold');
    spreadsheetProperties = Object.keys(ss).sort();
    for (i = 0;
    i < spreadsheetProperties.length;
    i += 1) {
        outputRngStart.offset(i, 0)
            .setValue(spreadsheetProperties[i]);
    }
}
//  Extract, an array of properties from a
//   Sheet object.
// Sort the array alphabetically using the
//  Array sort() method.
// Use the Array join() method to a create
//   a string of all the Sheet properties
//   separated by a new line.
function printSheetProperties() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        sh = ss.getActiveSheet();
    Logger.log(Object.keys(sh)
        .sort().join('\n'));
}
// Call function listSheets() passing it the 
//  Spreadsheet object for the active 
//    spreadsheet.
// The try - catch construct handles the 
//  error thrown by listSheets() if the given
// argument is absent or something
//    other than a Spreadsheet object.
function test_listSheets() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    try {
        listSheets(ss);
    } catch (error) {
        Logger.log(error.message);
    }
}
// Given a Spreadsheet object, 
//  print the names of its sheets
//   to the logger.
// Throw an error if the argument
//   is missing or if it is not
//  of type Spreadsheet.

// Create a Spreadsheet object and call 
//  "sheetExists()" for an array of sheet 
//  names to see if they exist in 
//  the given spreadsheet.
// Print the output to the log.
function test_sheetExists() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        sheetNames = ['Sheet1',
            'sheet1',
            'Sheet2',
            'SheetX'],
        i;
    for (i = 0;
    i < sheetNames.length;
    i += 1) {
        Logger.log('Sheet Name ' + sheetNames[i] +
            ' exists: ' + sheetExists(ss,
        sheetNames[i]));
    }
}
// Given a Spreadsheet object and a sheet name, 
//  check for two arguments of the correct type.
// Return "true" if the given sheet name exists
//  in the given Spreadsheet, 
//  else return "false".
function sheetExists(spreadsheet, sheetName) {
    var sheet;
    if (spreadsheet.toString() !==
        'Spreadsheet') {
        throw {
            'name': 'TypeError',
                'message': 'Function "sheetExists()" ' +
                'first argument for ' +
                '"spreadsheet" is ' +
                'not type "Spreadsheet".'
        };
    }
    if (typeof sheetName !== 'string') {
        throw {
            'name': 'TypeError',
                'message': 'Function "sheetExists()" ' +
                'second argument ' +
                'for "sheetName" ' +
                'is not type string.'
        };
    }
    if (spreadsheet.getSheetByName(sheetName)) {
        return true;
    } else {
        return false;
    }
}
// Copy the first sheet of the active
//  spreadsheet to a newly created 
//  spreadsheet.
function copySheetToSpreadsheet() {
    var ssSource = SpreadsheetApp.getActiveSpreadsheet(),
        ssTarget = SpreadsheetApp.create(
            'CopySheetTest'),
        sourceSpreadsheetName = ssSource.getName(),
        targetSpreadsheetName = ssTarget.getName();
    Logger.log(
        'Copying the first sheet from ' + sourceSpreadsheetName +
        ' to ' + targetSpreadsheetName);
    // [0] extracts the first Sheet object 
    //   from the array created by
    //   method call "getSheets()"
    ssSource.getSheets()[0].copyTo(ssTarget);
}
// Create a Sheet object and pass it 
// as an argument to getSheetSummary().
// Print the return value to the log.
function test_getSheetSummary() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet()
        .getSheets()[0];
    Logger.log(getSheetSummary(sheet));
}
// Given a Sheet object as an argument, 
//  use Sheet methods to extract 
//  information about it.
// Collect this information into an object
// literal and return the object literal.
function getSheetSummary(sheet) {
    var sheetReport = {};
    if (sheet.toString() !== 'Sheet') {
        throw {
            'name': 'TypeError',
                'message': 'Function "getSheetReport()" ' +
                'requires a single ' +
                'argument of type "Sheet".'
        };
    }
    sheetReport['Sheet Name'] = sheet.getName();
    sheetReport['Used Row Count'] = sheet.getLastRow();
    sheetReport['Used Column count'] = sheet.getLastColumn();
    sheetReport['Used Range Address'] =
        'A1:' + sheet.getRange(sheet.getLastRow(),
    sheet.getLastColumn()).getA1Notation();
    return sheetReport;
}

// Chapter 5

// Select a number of cells in a spreadsheet and 
//  then execute the following function.
// The address of the selected range, that is the
//   active range, is written to the log.
function activeRangeFromSpreadsheetApp() {
    var activeRange = SpreadsheetApp.getActiveRange();
    Logger.log(activeRange.getA1Notation());
}
// Get the active cell and print its containing
//  sheet name and address to the log.
// Try re-running after adding a new sheet
//  and selecting a cell at random.
function activeCellFromSheet() {
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet(),
        activeCell = activeSpreadsheet.getActiveCell(),
        activeCellSheetName = activeCell.getSheet().getSheetName(),
        activeCellAddress = activeCell.getA1Notation();
    Logger.log('The active cell is in sheet: ' + activeCellSheetName);
    Logger.log('The active cell address is: ' + activeCellAddress);
}
// Print Range object properties 
// (all are methods) to log.
function printRangeMethods() {
    var rng = SpreadsheetApp.getActiveRange();
    Logger.log(Object.keys(rng)
        .sort().join('\n'));
}
// Creating a Range object using two different 
//  overloaded versions of the Sheet 
//  "getRange()" method.
// "getSheets()[0]" gets the first sheet of the 
//   array of Sheet objects returned by 
//  "getSheets()".
// Both calls to "getRange()" return a Range
// object representing the same range address
//   (A1:B10).
function getRangeObject() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        sh = ss.getSheets()[0],
        rngByAddress = sh.getRange('A1:B10'),
        rngByRowColNums = sh.getRange(1, 1, 10, 2);
    Logger.log(rngByAddress.getA1Notation());
    Logger.log(
    rngByRowColNums.getA1Notation());
}
// Set a number of properties for a range.
// Add a new sheet.
// Set various properties for the cell 
//  A1 of the new sheet.
function setRangeA1Properties() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        newSheet,
        rngA1;
    newSheet = ss.insertSheet();
    rngA1 = newSheet.getRange('A1');
    rngA1.setComment(
        'Hold The date returned by spreadsheet ' + ' function "TODAY()"');
    rngA1.setFormula('=TODAY()');
    rngA1.setBackgroundColor('black');
    rngA1.setFontColor('white');
    rngA1.setFontWeight('bold');
}
// Demonstrate get methods for 'Range' 
//  properties.
// Assumes function "setRangeA1Properties()
//   has been run.
// Prints the properties to the log.
// Demo purposes only!
function printA1PropertiesToLog() {
    var rngA1 = SpreadsheetApp.getActiveSpreadsheet()
        .getSheetByName('RangeTest').getRange('A1');
    Logger.log(rngA1.getComment());
    Logger.log(rngA1.getFormula());
    Logger.log(rngA1.getBackground());
    Logger.log(rngA1.getFontColor());
    Logger.log(rngA1.getFontWeight());
}
// Starting with cell C10 of the active sheet,
// add comments to each of its adjoining cells
//  stating the row and column offsets needed
//  to reference the commented cell 
// from cell C10.
function rangeOffsetDemo() {
    var rng = SpreadsheetApp.getActiveSheet()
        .getRange('C10');
    rng.setBackground('red');
    rng.setValue('Method offset()');
    rng.offset(-1, -1)
        .setComment('Offset -1, -1 from cell ' + rng.getA1Notation());
    rng.offset(-1, 0)
        .setComment('Offset -1, 0 from cell ' + rng.getA1Notation());
    rng.offset(-1, 1)
        .setComment('Offset -1, 1 from cell ' + rng.getA1Notation());
    rng.offset(0, 1)
        .setComment('Offset 0, 1 from cell ' + rng.getA1Notation());
    rng.offset(1, 0)
        .setComment('Offset 1, 0 from cell ' + rng.getA1Notation());
    rng.offset(0, 1)
        .setComment('Offset 0, 1 from cell ' + rng.getA1Notation());
    rng.offset(1, 1)
        .setComment('Offset 1, 1 from cell ' + rng.getA1Notation());
    rng.offset(0, -1)
        .setComment('Offset 0, -1 from cell ' + rng.getA1Notation());
    rng.offset(1, -1)
        .setComment('Offset -1, -1 from cell ' + rng.getA1Notation());
}
// Passing a deliberately "bad" argument to the 
//  Range offset() method.
// The row offset argument is -1 but 
//  there is no row  above row 1
//   (cell A1's row).
// Google Apps Script gives error:
//   "It looks like someone else
// already deleted this cell."
function offsetError() {
    var rng = SpreadsheetApp.getActiveSpreadsheet()
        .getActiveSheet()
        .getRange('A1');
    rng.offset(-1, 0)
        .setValue('bad offset argument.');
}

function offsetError() {
    var rng = SpreadsheetApp.getActiveSpreadsheet()
        .getActiveSheet().getRange('A1');
    Logger.log(rng.offset(-1, 0).getValue());
}
// See the Sheet method getDataRange() in action.
// Print the range address of the used range for
//  a sheet to the log.
function getDataRange() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        sheetName = 'english_premier_league',
        sh = ss.getSheetByName(sheetName),
        dataRange = sh.getDataRange();
    Logger.log(dataRange.getA1Notation());
}
// Read the entire data range of a sheet 
// into a JavaScript array.
// Uses the JavaScript Array.isArray()
//  method twice to verify that method
// getValues()returns an array-of-arrays. 
// Print the number of array elements 
// corresponding to the number of data 
//  range rows.
// Extract and print the first 10
//  elements of the array using the 
//  array slice() method.
function dataRangeToArray() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        sheetName = 'english_premier_league',
        sh = ss.getSheetByName(sheetName),
        dataRange = sh.getDataRange(),
        dataRangeValues = dataRange.getValues();
    Logger.log(Array.isArray(dataRangeValues));
    Logger.log(Array.isArray(dataRangeValues[0]));
    Logger.log(dataRangeValues.length);
    Logger.log(dataRangeValues.slice(0, 10));
}
// Loop over the array returned by 
//  getRange() and a CSV-type output
// to the log using array join() method.
function loopOverArray() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        sheetName = 'english_premier_league',
        sh = ss.getSheetByName(sheetName),
        dataRange = sh.getDataRange(),
        dataRangeValues = dataRange.getValues(),
        i;
    for (i = 0;
    i < dataRangeValues.length;
    i += 1) {
        Logger.log(
        dataRangeValues[i].join(','));
    }
}
// In production code, this function would be 
//   re-factored into smaller functions.
// Read the data range into a JavaScript array.
// Remove and store the header line using the
//    array shift() method.
// Use the array filter() method with an anonymous
//   function as a callback to implement the 
//   filtering logic.
// Determine the element count of the 
//  filter() output array.
// Add a new sheet and store a reference to it.
// Create a Range object from the new
//   Sheet objectusing the getRange() method.
// The four arguments given to getRange() are:
//   (1) first column, (2) first row,
//   (3) last row, and (4) last column.  
// This creates a range corresponding to 
//  range address "A1:C5".
// Write the values of the filtered array to the 
//  newly created range.
function writeFilteredArrayToRange() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        sheetName = 'english_premier_league',
        sh = ss.getSheetByName(sheetName),
        dataRange = sh.getDataRange(),
        dataRangeValues = dataRange.getValues(),
        filteredArray,
        header = dataRangeValues.shift(),
        filteredArray,
        filteredArrayColCount = 3,
        filteredArrayRowCount,
        newSheet,
        outputRange;
    filteredArray = dataRangeValues.filter(

    function (innerArray) {
        if (innerArray[2] >= 40) {
            return innerArray;
        }
    });
    filteredArray.unshift(header);
    filteredArrayRowCount = filteredArray.length;
    newSheet = ss.insertSheet();
    outputRange = newSheet.getRange(1,
    1,
    filteredArrayRowCount,
    filteredArrayColCount);
    outputRange.setValues(filteredArray);
}

function setRangeName() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        sh = ss.getActiveSheet(),
        rng = sh.getRange('A1:B10'),
        rngName = 'MyData';
    ss.setNamedRange(rngName, rng);
}
// Create a range object using the 
//  getDataRange() method.
// Pass the range and a colour name 
//  to function "setAlternateRowsColor()".
// Try changing the 'color' variable to
//   something like:
//   'red', 'green', 'yellow', 'gray', etc.
function call_setAlternateRowsColor() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        sheetName = 'english_premier_league',
        sh = ss.getSheetByName(sheetName),
        dataRange = sh.getDataRange(),
        color = 'grey';
    setAlternateRowsColor(dataRange, color);
}

// Set every second row in a given range to
//   the given colour.
// Check for two arguments:
//   1: Range, 2: string for colour.
// Throw a type error if either argument
//  is missing or of the wrong type.
// Use the Range offset() to loop 
//   over the range rows.
// the for loop counter starts at 0.
// It is tested in each iteration with the
//   modulus operator (%).
// If i is an odd number, the if condition 
// evaluates to true and the colour 
//  change is applied.
// WARNING: If a non-existent colour is given, 
//  then the "color" is set to undefined
// no color). NO error is thrown!
function setAlternateRowsColor(range,
color) {
    if (range.toString() !== 'Range') {
        throw {
            'name': 'TypeError',
                'message':
                'The first argument to ' +
                '"setAlternateRowsColor()"  ' +
                ' must be type Range'
        };
    }
    if (typeof color !== 'string') {
        throw {
            'name': 'TypeError',
                'message':
                'The second argument to ' +
                ' "setAlternateRowsColor()" ' +
                ' must be a string for a color,' +
                '  e.g. "red"'
        };
    }
    var i,
    startCell = range.getCell(1, 1),
        columnCount = range.getLastColumn(),
        lastRow = range.getLastRow();
    for (i = 0; i < lastRow; i += 1) {
        if (i % 2) {
            startCell.offset(i, 0, 1, columnCount)
                .setBackgroundColor(color);
        }
    }
}
// Test function for 
//  "deleteLeadingTrailingSpaces()".
// Creates a Range object and passes
//   it to this function.
function call_deleteLeadingTrailingSpaces() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        sheetName = 'english_premier_league',
        sh = ss.getSheetByName(sheetName),
        dataRange = sh.getDataRange();
    deleteLeadingTrailingSpaces(dataRange);
}

// Process each cell in the given range.
// If the cell is of type text 
//  (typeof === 'string') then
//  remove leading and trailing white space.  
//  Else ignore it.
// Code note: The Range getCell() method
//   takes two 1-based indexes 
//  (row and column).
//   This is in contrast to the offset() method. 
//  rng.getCell(0,0) will throw an error!
function deleteLeadingTrailingSpaces(range) {
    if (range.toString() !== 'Range') {
        throw {
            'name': 'TypeError',
                'message':
                'Argument to ' +
                '"deleteLeadingTrailingSpaces()" ' +
                'must be type Range'
        };
    }
    var i,
    j,
    lastCol = range.getLastColumn(),
        lastRow = range.getLastRow(),
        cell,
        cellValue;
    for (i = 1; i <= lastRow; i += 1) {
        for (j = 1; j <= lastCol; j += 1) {
            cell = range.getCell(i, j);
            cellValue = cell.getValue();
            if (typeof cellValue === 'string') {
                cellValue = cellValue.trim();
                cell.setValue(cellValue);
            }
        }
    }
}
// Create a Sheet object for the active sheet.
// Pass the sheet object to 
//   "getAllDataRangeFormulas()"
// Create an array of the keys in the returned 
//  object in default "sort()".
// Loop over the array of sorted keys and 
//  extract the  values they keys map to.
// Write the output to the log.
function call_getAllDataRangeFormulas() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        sheet = ss.getActiveSheet(),
        sheetFormulas = getAllDataRangeFormulas(sheet),
        formulaLocations = Object.keys(sheetFormulas).sort(),
        formulaCount = formulaLocations.length,
        i;
    for (i = 0; i < formulaCount; i += 1) {
        Logger.log(formulaLocations[i] +
            ' contains ' + sheetFormulas[formulaLocations[i]]);
    }
}
// Take a Sheet object as an argument, 
//  throw an error if the given argument 
//  is of the wrong type.
// Return an object literal where formula 
//  locations map to formulas for all formulas
//   in the input sheet data range.
// Loop through every cell in the data range.
// If a cell has a formula, 
//   store that cells location as 
//  the key and its formula as the value
//   in the object literal.
// Return the populated object literal.
function getAllDataRangeFormulas(sheet) {
    if (sheet.toString() !== 'Sheet') {
        throw {
            'name': 'TypeError',
                'message':
                'Function "getAllDataRangeFormulas()" ' +
                ' expects a single argument of ' +
                ' type Sheet.'
        };
    }
    var dataRange = sheet.getDataRange(),
        i,
        j,
        lastCol = dataRange.getLastColumn(),
        lastRow = dataRange.getLastRow(),
        cell,
        cellFormula,
        formulasLocations = {},
        sheetName = sheet.getSheetName(),
        cellAddress;
    for (i = 1; i <= lastRow; i += 1) {
        for (j = 1; j <= lastCol; j += 1) {
            cell = dataRange.getCell(i, j);
            cellFormula = cell.getFormula();
            if (cellFormula) {
                cellAddress = sheetName + '!' + cell.getA1Notation();
                formulasLocations[cellAddress] = cellFormula;
            }
        }
    }
    return formulasLocations;
}
// Call copyColumns() function passing it:
//  1: The active sheet
//  2: A newly inserted sheet
//  3: An array of column indexes to copy
//     to the new sheet
// The output in the newly inserted sheet 
//  contains the columns for the indexes
//   given in the array in the 
// order specified in the array.
function call_copyColumns() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        inputSheet = ss.getActiveSheet(),
        outputSheet = ss.insertSheet(),
        columnIndexes = [4, 3, 2, 1];
    copyColumns(inputSheet,
    outputSheet,
    columnIndexes);
}

// Given an input sheet, an output sheet,
// and an array:
// Use the numeric column indexes in 
//  the array to copy those columns from 
//  the input sheet to the output sheet.
// The function checks its input arguments
//   and throws an error
// if they are not Sheet, Sheet, Array.
// The array is expected to be an array of
//  integers but it does 
//   not check the array element types
function copyColumns(inputSheet,
outputSheet,
columnIndexes) {

    var dataRangeRowCount = inputSheet.getDataRange()
        .getNumRows(),
        columnsToCopyCount = columnIndexes.length,
        i,
        columnIndexesCount,
        valuesToCopy = [];
    for (i = 0;
    i < columnsToCopyCount;
    i += 1) {
        valuesToCopy = inputSheet.getRange(1,
        columnIndexes[i],
        dataRangeRowCount,
        1).getValues();
        outputSheet.getRange(1,
        i + 1,
        dataRangeRowCount,
        1).setValues(valuesToCopy);
    }
}

// Chapter 6

// Test connection to a  MySQ
//  cloud instance created earlier.
// Check log for output.
function connectMySqlCloud() {
    var connStr =
        'jdbc:google:rdbms://' +
        'elwarbito:chapter6/contacts',
        conn;
    try {
        conn = Jdbc.getCloudSqlConnection(connStr);
        Logger.log('Connection OK!');
    } catch (err) {
        Logger.log(err);
        throw err;
    } finally {
        if (conn) {
            conn.close();
        }
    }
}
// Execute a CREATE TABLE DDL statement for a 
//  database named "contacts".
function createTable() {
    var connStr =
        'jdbc:google:rdbms://' +
        'elwarbito:chapter6/contacts',
        conn,
        stmt,
        ddl;
    ddl = 'CREATE TABLE person(' +
        '  person_id  MEDIUMINT ' +
        'AUTO_INCREMENT' +
        ' NOT NULL PRIMARY KEY,' +
        '  first_name VARCHAR(100) NOT NULL,' +
        '  last_name VARCHAR(100) NOT NULL,' +
        '  date_of_birth DATE,' +
        '  height_cm SMALLINT)';
    try {
        conn = Jdbc.getCloudSqlConnection(connStr);
        stmt = conn.createStatement();
        stmt.execute(ddl);
        Logger.log('Table created!');
    } catch (ex) {
        Logger.log(ex);
        throw (ex);
    } finally {
        Logger.log('Cleaning up.');
        if (stmt) {
            stmt.close();
        }
        if (conn) {
            conn.close();
        }
    }
}
// Add 6 rows to newly created table.
// Data source is a JavaScript array-of-arrays.
// Uses bind parameters.
// Executes an SQL INSERT INTO statement 
//  within a for loop.
function addRowsToTable() {
    var connStr = 'jdbc:google:rdbms://' +
        'elwarbito:chapter6/contacts',
        conn,
        dml,
        prepStmt,
        rows,
        i,
        row,
        firstName,
        lastName,
        dateOfBirth,
        heightcm;
    dml = 'INSERT INTO person(first_name, ' +
        'last_name, ' +
        'date_of_birth, ' +
        'height_cm) ' +
        'VALUES(?, ?, ?, ?)';
    rows = [
        ['Joe', 'Grey', '1970-06-11', 182],
        ['Raj', 'Patel', '1975-03-13', 188],
        ['Amy', 'Lopez', '1972-08-17', 166],
        ['Bill', 'Grimes', '1954-10-20', 181],
        ['Jane', 'Rice', '1961-04-30', 170],
        ['Alex', 'Lee', '1982-08-06', 190]
    ];
    try {
        conn = Jdbc.getCloudSqlConnection(connStr);
        prepStmt = conn.prepareStatement(dml);
        for (i = 0; i < rows.length; i += 1) {
            row = rows[i];
            firstName = row[0];
            lastName = row[1];
            dateOfBirth = row[2];
            heightcm = row[3];
            prepStmt.setString(1, firstName);
            prepStmt.setString(2, lastName);
            prepStmt.setString(3, dateOfBirth);
            prepStmt.setInt(4, heightcm);
            prepStmt.execute();
        }
        Logger.log('Loaded row count: ' + i);
    } catch (ex) {
        Logger.log(ex);
        throw (ex);
    } finally {
        if (prepStmt) {
            prepStmt.close();
        }
        if (conn) {
            conn.close();
        }
    }
}
// Retrieve all rows from a table and 
//  write the to a newly added sheet.
// If re-running this code, ensure 
//  that added sheet is deleted 
//  or re-named.
function writeDatabasRowsToSpreadsheet() {
    var connStr = 'jdbc:google:rdbms://' +
        'elwarbito:chapter6/contacts',
        conn,
        sql = 'SELECT * FROM person',
        stmt,
        rs,
        colCount,
        colVal,
        rowVals = [],
        i,
        ss = SpreadsheetApp.getActiveSpreadsheet(),
        sh = ss.insertSheet();
    sh.setName('QueryResults');
    try {
        conn = Jdbc.getCloudSqlConnection(connStr);
        stmt = conn.createStatement();
        rs = stmt.executeQuery(sql);
        colCount = rs.getMetaData()
            .getColumnCount();
        Logger.log('Col count: ' + colCount);
        while (rs.next()) {
            for (i = 1; i <= colCount; i += 1) {
                colVal = rs.getString(i);
                rowVals.push(colVal);
            }
            sh.appendRow(rowVals);
            rowVals = [];
        }
    } catch (ex) {
        Logger.log(ex);
        throw (ex);
    } finally {
        if (rs) {
            rs.close();
        }
        if (stmt) {
            stmt.close()
        };
        if (conn) {
            conn.close();
        }
    }
}
// Execute the MySQL "SHOW TABLES"
// statement and print the table
//  names in the target database
//  to the log.
// Exception handling is missing!
function showTables() {
    var connStr =
        'jdbc:google:rdbms://' +
        'elwarbito:chapter6/contacts',
        conn,
        stmt,
        rs,
        sql;
    sql = 'SHOW TABLES';
    conn = Jdbc.getCloudSqlConnection(connStr);
    stmt = conn.createStatement();
    rs = stmt.executeQuery(sql);
    while (rs.next()) {
        Logger.log(rs.getString(1));
    }
    rs.close();
    stmt.close();
    conn.close();
}
// Print some metadata about the database created
//  in earlier examples.
// Adds a sheet named "databaseMetadata".
// If re-running, remove sheet added previously.
// Exception handling dropped.
function writeMySQLMetadataToSheet() {
    var connStr = 'jdbc:google:rdbms://' +
        'elwarbito:chapter6/contacts',
        dbMetadata,
        rsTables,
        rsColumns,
        tableNames = [],
        i,
        tableCount,
        ss,
        sh;
    ss = SpreadsheetApp.getActiveSpreadsheet(),
    sh = ss.insertSheet()
    conn = Jdbc.getCloudSqlConnection(connStr);
    dbMetadata = conn.getMetaData();
    sh.setName('DatabaseMetadata');
    sh.appendRow(['Major Version',
    dbMetadata.getDatabaseMajorVersion()]);
    sh.appendRow(['Minor Version',
    dbMetadata.getDatabaseMinorVersion()]);
    sh.appendRow(['Product Name',
    dbMetadata.getDatabaseProductName()]);
    sh.appendRow(['Product Version',
    dbMetadata.getDatabaseProductVersion()]);
    sh.appendRow(['Supports transactions',
    dbMetadata.supportsTransactions()]);
    rsTables = dbMetadata.getTables(
    null, null, null, ['TABLE']);
    while (rsTables.next()) {
        tableNames.push(rsTables.getString(3));
    }
    tableCount = tableNames.length;
    sh.appendRow(
    ['Table Names And Columns Names Are:']);
    for (i = 0; i < tableCount; i += 1) {
        rsColumns = dbMetadata.getColumns(
        null, null, tableNames[i], null);
        while (rsColumns.next()) {
            sh.appendRow(
            [tableNames[i], rsColumns.getString(4)]);
        }
    }
    rsTables.close();
    rsColumns.close();
    conn.close();
}


// Chapter 10
// For this chapter, HTML and client JavaScript will be enclosed in /* */ comment blocks!

// Execute this function and switch to 
//  spreadsheet tab to see the HTML form
//   built in the file index.html.
function demoHtmlServices() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        html = HtmlService.createHtmlOutputFromFile('index');
    ss.show(html);
}
// This function is called by the JavaScript
//  function "formSubmit()" defined in the
//  accompanying HTML file.
function getValuesFromForm(form) {
    var firstName = form.firstName,
        lastName = form.lastName,
        sheet = SpreadsheetApp.getActiveSpreadsheet()
            .getActiveSheet();
    sheet.appendRow([firstName, lastName]);
}
/*
<!--
A very simple data entry form that writes
the text input values back to a spreadsheet.
-->
<div>
<b>Add Row To Spreadsheet</b><br />
<form>
First name: <input id="firstname" 
             name="firstName" type="text" />
<br>
Last name: <input id="lastname" 
            name="lastName" type="text" />
<br>
<input onclick="formSubmit()" 
       type="button" value="Add Row" />
<input onclick="google.script.host.close()" 
       type="button" value="Exit" />
</form>
<script type="text/javascript">
function formSubmit() {
  google.script.run.
    getValuesFromForm(document.forms[0]);
        }
    </script>
</div>
*/
/*
<!-- 
Create a simple user feedback form called "survey.html".
-->
<div>
<h1>Customer Satisfaction</h1>
  <form>
    <fieldset>
    <legend>Enter Customer Details:</legend>
    <p><label>Email: </label>
      <input type="text" name="email" size="30"/>
    </p>
	<p><label>Gender: </label>
         <input type="radio" name="gender"
                value="Male" id="gender"/> Male
         <input type="radio" name="gender"
                value="Female" id="gender"/> Female
    </p>
	<p><label>Country:
  <select name="country" id="country">
  <option value="USA">USA</option>
  <option value="Canada">Canada</option>
  <option value="UK">UK</option>
  <option value="Australia">Australia</option>
  <option value="New Zealand">New Zealand</option>
	</select></label>
    </p>
    </fieldset>
    <fieldset>
    <legend class="mylbl">Lengthy Note</legend>
       <textarea rows="4" cols="58" 
         name="note" id="note">

	</textarea> 
   </fieldset>
   <p>
   <input type="button" 
      value="Send Feedback" onclick="formSubmit()"/>
   <input type="button" 
     value="Cancel" onclick="clear()" />
   </p>
   <p id="message">
   </p>
  </form>
<script type="text/javascript">
function formSubmit() {
  google.script.run.
    getValuesFromForm(document.forms[0]);
  document.forms[0].reset();
  alert('Submitted');
}
function clear() {
  document.forms[0].reset();
}
</script>
</div>
*/
// Required function name for web apps.
function doGet() {
    var html = HtmlService.createHtmlOutputFromFile('survey');
    html.setTitle('Customer Survey');
    html.setHeight(600);
    return html;
}
// Extract values from the web app form
//  and write them to a sheet named
//  "Feedback".
function getValuesFromForm(form) {
    var email = form.email,
        gender = form.gender,
        country = form.country,
        note = form.note,
        ssId =
            '0Amdsdq7IKB9ydExEanFqMm5ocmRSMndyRmFOeTgxckE',
        ss = SpreadsheetApp.openById(ssId),
        shName = 'Feedback',
        sheet = ss.getSheetByName(shName);
    sheet.appendRow([email,
    gender,
    country,
    note]);
}
/*
<div>
<h1>English Premier League</h1>

<? var data = getData(); ?>
<table style="border: 1px solid black;">
  <? for (var i = 0; i < data.length; i++) { ?>
    <tr>
      <? for (var j = 0; 
                  j < data[i].length; j++) { ?>
        <td style="border: 1px solid black;">
        <?= data[i][j] ?></td>
      <? } ?>
    </tr>
  <? } ?>
</table>
</div>
*/
function doGet() {
    var html = HtmlService.createTemplateFromFile('premier_league');
    return html.evaluate();
}

function getData() {
    var ssId =
        '0Amdsdq7IKB9ydE5KYXkyRmJZZ244Qmo0ODVrX0dXekE',
        ss = SpreadsheetApp.openById(ssId),
        rng = ss.getRangeByName('premier_league_table'),
        data = rng.getValues();
    return data;
}


function spreadsheetInstance() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log(ss.getName());
}

function firstSheetInfo() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        sheets = ss.getSheets(),
        // getSheets() returns an array
        // JavaScript arrays are always zero-based
        sh1 = sheets[0];
    Logger.log(sh1.getName());
    Logger.log(sh1.getDataRange().getA1Notation());
}

function printSheetNames() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        sheets = ss.getSheets(),
        i;
    for (i = 0; i < sheets.length; i += 1) {
        Logger.log(sheets[i].getName());
    }
}

function addNewSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        newSheet;
    newSheet = ss.insertSheet();
    newSheet.setName("AddedSheet");
    Browser.msgBox("New Sheet Added!");
}


function removeSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        sheetToRemove = ss.getSheetByName("AddedSheet");
    sheetToRemove.activate();
    ss.deleteActiveSheet();
    Browser.msgBox("SheetDeleted!");
}

function sheetHide() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        sh = ss.getSheetByName('ToHide');
    sh.hideSheet()
}

function listHiddenSheetNames() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        sheets = ss.getSheets();
    sheets.forEach(

    function (sheet) {
        if (sheet.isSheetHidden()) {
            Logger.log(sheet.getName());
        }
    });
}

function sheetsUnhide() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        sheets = ss.getSheets();
    sheets.forEach(

    function (sheet) {
        if (sheet.isSheetHidden()) {
            sheet.showSheet();
        }
    });
}


function getUsedRangeAsArray(sheetName) {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        sh = ss.getSheetByName(sheetName);
    // The getValues() method of the
    //   Range object returns an array of arrays
    return sh.getDataRange().getValues();
}

function test_getUsedRangeAsArray() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        sheetName = 'Sheet1',
        rngValues = getUsedRangeAsArray(sheetName);
    // Print the number of rows in the range
    // The toString() call to suppress the 
    // decimal point so
    //  that, for example, 10.0, is reported as 10
    Logger.log((rngValues.length).toString());
    // Print the number of columns
    // The column count will be the same 
    // for all rows so only need the first row
    Logger.log((rngValues[0].length).toString());
    // Print the value in the first cell
    Logger.log(rngValues[0][0]);
}

function addColorsToRange() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        sheets = ss.getSheets(),
        sh1 = sheets[0],
        addr = 'A4:B10',
        rng;
    // getRange is overloaded. This method can
    //  also accept row and column integers
    rng = sh1.getRange(addr);
    rng.setFontColor('green');
    rng.setBackgroundColor('red');
}

function offsetDemo() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        sh = ss.getSheets()[0],
        cell = sh.getRange('B2');
    cell.setValue('Middle');
    cell.offset(-1, -1).setValue('Top Left');
    cell.offset(0, -1).setValue('Left');
    cell.offset(1, -1).setValue('Bottom Left');
    cell.offset(-1, 0).setValue('Top');
    cell.offset(1, 0).setValue('Bottom');
    cell.offset(-1, 1).setValue('Top Right');
    cell.offset(0, 1).setValue('Right');
    cell.offset(1, 1).setValue('Bottom Right');
}

function offsetOverloadDemo() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        sh = ss.getSheets()[0],
        cell = sh.getRange('A1'),
        offsetRng2 = cell.offset(1, 4, 2),
        offsetRng3 = cell.offset(10, 4, 4, 5);
    Logger.log('Address of offset() overload 2 ' +
        '(rowOffset, columnOffset, numRows) is: ' + offsetRng2.getA1Notation());
    Logger.log('Address of offset() overload 3 ' +
        '(rowOffset, columnOffset, numRows, ' +
        'numColumns) is: ' + offsetRng3.getA1Notation());
}

// Copy rows from one sheet named "Source" to
//  a newly inserted
//   one based on a criterion check of second
//   column.
// Copy the header row to the new sheet.
// If Salary <= 10,000 then copy the entire row
function copyRowsToNewSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        sourceSheet = ss.getSheetByName('Source'),
        newSheetName = 'Target',
        newSheet = ss.insertSheet(newSheetName),
        sourceRng = sourceSheet.getDataRange(),
        sourceRows = sourceRng.getValues(),
        i;
    newSheet.appendRow(sourceRows[0]);
    for (i = 1; i < sourceRows.length; i += 1) {
        if (sourceRows[i][1] <= 10000) {
            newSheet.appendRow(sourceRows[i]);
        }
    }
}

function test_printSheetFormulas() {
    var sheetName = 'Formulas';
    printSheetFormulas(sheetName);
}

function printSheetFormulas(sheetName) {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
        sourceSheet = ss.getSheetByName(sheetName),
        usedRng = sourceSheet.getDataRange(),
        i,
        j,
        cell,
        cellAddr,
        cellFormula;
    for (i = 1; i <= usedRng.getLastRow();
    i += 1) {
        for (j = 1; j <= usedRng.getLastColumn();
        j += 1) {
            cell = usedRng.getCell(i, j);
            cellAddr = cell.getA1Notation();
            cellFormula = cell.getFormula();
            if (cellFormula) {
                Logger.log(cellAddr +
                    ': ' + cellFormula);
            }
        }
    }
}

