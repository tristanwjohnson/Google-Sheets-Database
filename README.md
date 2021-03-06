# Google-Sheets-Database
Code to convert a set of Google spread sheets and sheets into a working database with CRUD functions, race condition handling, and additional functionality. Made Initially for Adibi IP Group.

FUNCTIONS
> accessDatabase(functionName, spreadsheetID, sheetName, parameters) {...}

- @param {string} functionName: the name of the function that we wish to call (uses string matching to identify)
- @param {string} spreadsheetID: the id of the spreadsheet we wish to access
- @param {string} sheetName: the name of the sheet we wish to access
- @param {object} parameters: list of input parameters to be passed into the function
- @return {object} returns the output of the function that is called

This is the wrapper function called from the other files that manages resource locking and calls the helper functions to perform the actual CRUD operations. functionName can be "CREATE", "READ", "UPDATE", "DELETE", "UNDO_DELETE". See below for more info


> create_(sheet, inputData) {...}

- @param {string} sheet: the instance of the sheet we wish to write to
- @param {object} inputData: a list of dictionaries {fieldName: data} representing the data we wish to create new rows with
- @return {object} returns a dictionary of dictionaries {ID: {fieldName: data}} corresponding to the inputted rowIDs

Function to create new rows in the given sheet using inputData
(e.g. accessDatabase("CREATE", spreadsheetID, sheetName, [inputData]))


> read_(sheet, columnName, rowValues) {...}

 - @param {string} sheet: the instance of the sheet we wish to read from
 - @param {string} columnName: the name of the column that we are comparing values in for reading
 - @param {object} rowValues: a list of the values of the rows for the given columnName that we wish to read
 - @return {object} return a dictionary of dictionaries {ID: {fieldName: data}} corresponding to the inputted rowIDs.
                   If a rowID does not exist or is invalid, then the corresponding entry will be null.
                   If the rowIDs input is null itself, this function returns a dictionary corresponding to ALL rows in the sheet.

Function to return the rows specified by the given rowIDs
(e.g. accessDatabase("READ", spreadsheetID, sheetName, [columnName, rowValues]))


> update_(sheet, inputDict) {...}

 - @param {string} sheet: the instance of the sheet we wish to update
 - @param {object} inputDict: a dictionary {fieldName: data} representing the data we wish to create a new row with - to be passed into create_
 - @return {object} returns the output of create_ - dictionary of dictionaries {ID: {fieldName: data}} corresponding to the inputted oldRowID

Function to 'Update' (delete and then create) the row with the oldRowID as its ID
(e.g. accessDatabase("UPDATE", spreadsheetID, sheetName, [inputData]))


> delete_(sheet, columnName, rowValues) {...}

 - @param {string} sheet: the instance of the sheet we wish to delete from
 - @param {string} columnName: the name of the column that we are comparing values in for deletion
 - @param {object} rowValues: the values that we use to compare for deletion
 - @return {object} returns a dictionary of dictionaries {ID: {fieldName: data}} representing the deleted rows by their IDs. Returns null if inputs are invalid - bad column name, null rowValues or empty list, sheet missing a valid column 

Function to 'delete' (set valid=FALSE) all rows with the value of the given columnName equal to rowValue(e.g. accessDatabase("DELETE", spreadsheetID, sheetName, [columnName, rowValues]))


> undoDelete_(sheet, columnName, rowValues) {...}

 - @param {object} sheet: the sheet that this operation is taking place on
 - @param {string} columnName: the field name for the column that we are using to determine which row to revalidate
 - @param {string} {object} rowValues: the values that we use to compare for undoing deletion
 - @return {object} returns a list of dictionaries of the rows that were just undeleted

Function to undo a delete operation (does not handle foreign keys) (e.g. accessDatabase("UNDO_DELETE", spreadsheetID, sheetName, [columnName, rowValues]))


> cleanDatabase() {...}

 - return {integer} the number of rows deleted 
 
Function to remove every false row in the database that is at least a day old
