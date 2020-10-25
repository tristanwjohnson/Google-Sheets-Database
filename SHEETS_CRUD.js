/**
 * This file contains functions that allow the user to read and write from the sheets database.
 * This file also contains the wrapper function for these CRUD operations that manages resource locking.
 *
 * We are assuming that, for every sheet, the unique ID is in column 1, CreatedBy is in column 2, ModifiedBy is in column 3, DateCreated is in column 4, DateModified is in column 5, Valid is in column 6, all other information follows.
 */

/**
 * This is the wrapper function called from the other files that manages resource locking and calls the
 * helper functions to perform the actual CRUD operations.
 *
 * @param {string} functionName: the name of the function that we wish to call (uses string matching to identify)
 * @param {string} spreadsheetID: the id of the spreadsheet we wish to access
 * @param {string} sheetName: the name of the sheet we wish to access
 * @param {object} parameters: list of input parameters to be passed into the function
 * @return {object} returns the output of the function that is called
 */
function accessDatabase(functionName, spreadsheetID, sheetName, parameters) {
    // Variable to store the value to be returned once the code has been unlocked
    var returnValue;
    // Get a script lock, because we're about to modify a shared resource.
    var lock = LockService.getScriptLock();
    try {
      // Wait for up to 30 seconds for other processes to finish.
      lock.waitLock(30000);
    } catch (err) {
      // Error handling if the lock times out.
      Logger.log("The call to " + functionName + " timed out because the resource was in use. This resulted in the following error: " + err);
      return;
    }
  
    // This section of the code is now locked --------------------------------------------
    // Get the instance of the sheet that we wish to access
    Logger.log(spreadsheetID);
    Logger.log(sheetName);
    if (functionName == "CREATE_SHEET") {
      returnValue = createSheet_(spreadsheetID, sheetName, parameters[0]);
    } else if (functionName == "CLEAN_SHEET") {
      returnValue = cleanSheet_();
    } else {
      const sheet = SpreadsheetApp.openById(spreadsheetID).getSheetByName(sheetName);
      // Parse the function name to determine which CRUD function to call
      if (functionName == "CREATE") {
        // parameters[0] = a list of dictionaries representing the new rows we wish to create
        returnValue = create_(sheet, parameters[0]);
      } else if (functionName == "READ") {
        // parameters[0] = the name of the column that we will use to read values
        // parameters[1] = a list of values to be compared with the specified column to determine which rows are read
        returnValue = read_(sheet, parameters[0], parameters[1]);
      } else if (functionName == "UPDATE") {
        // parameters[0] = a dictionary of updated information to be placed in the newly created row (contains the ID of the old row to be deleted)
        returnValue = update_(sheet, parameters[0]);
      } else if (functionName == "DELETE") {
        // parameters[0] = the name of the column that we will use to delete values
        // parameters[1] = a list of values to be compared with the specified column to determine which rows are deleted
        returnValue = delete_(sheet, parameters[0], parameters[1]);
      } else if (functionName == "UNDO_DELETE") {
        // parameters[0] = the name of the column that we will use to find rows to undo delete
        // parameters[1] = a list of values to be compared with the specified column to determine which rows were deleted to undo
        returnValue = undoDelete_(sheet, parameters[0], parameters[1]);
      } else {
        // An invalid function name was inputted
        Logger.log(functionName + " is not a valid function name.");
        return;
      }
    }
    // Release the lock so that other processes can continue.
    lock.releaseLock();
    // This section of the code is now unlocked ------------------------------------------
  
    // return the return value
    return returnValue;
  }
  
  // ------------------------------------------- CREATE ----------------------------------------------------
  
  /**
   * Function to create new rows in the given sheet using inputData
   *
   * @param {string} sheet: the instance of the sheet we wish to write to
   * @param {object} inputData: a list of dictionaries {fieldName: data} representing the data we wish to create new rows with
   * @return {object} returns a dictionary of dictionaries {ID: {fieldName: data}} corresponding to the inputted rowIDs
   */
  function create_(sheet, inputData) {
    Logger.log("Creating a new row for the sheet: " + sheet.getName());
  
    // check if inputs are valid to function
    if (!validateCreateInputs(sheet, inputData)) {
      Logger.log("Inputs to create_ are invalid");
      return;
    }
    // get general information on creation to be stored in the sheet
    const dateTime = getDatetime_();
    const creator = getUserName_();
  
    // get the width of the data range and field values
    var dataRangeWidth = sheet.getDataRange().getWidth();
    var fieldValues = sheet.getRange(1, 1, 1, dataRangeWidth).getValues()[0]; // list of field names
    Logger.log(fieldValues);
    // create a dictionary to store fieldNames and their locations in the sheet (allows us to format the array correctly)
    var fieldLocs = {};
    for (var i = 0; i < fieldValues.length; i++) {
      fieldLocs[fieldValues[i]] = i;
    }
    // instantiate the various variables to be used within the for-loops
    var fieldKeys; // will hold a list of the keys for each entry of input data
    var newRow; // will store the row currently being created
    var fieldLocation; // will store the location of the field (used for inserting data into the new row)
    var uniqueID; // unique id generated for each row
    var listIDs = []; // list of ids to be returned by the function
    var rowDict = {}; // dictionary to hold each row dictionary
    for (var i = 0; i < inputData.length; i++) {
      //define the row dictionary as the inputted row dictionary
      var row = inputData[i];
      //pull the keys
      fieldKeys = Object.keys(inputData[i]);
      newRow = [];
      // iterate through each field in this row and add its data to the newData array
      for (var j = 0; j < fieldKeys.length; j++) {
        // get the index of the array where this specific piece of data should be stored
        fieldLocation = fieldLocs[fieldKeys[j]];
        // if this location is not in fieldLocs (not in the sheet), add it as a new column
        if (fieldLocation == null) {
          // add a column to the sheet and add the new field
          sheet.insertColumnAfter(dataRangeWidth);
          var newFieldRange = sheet.getRange(1, dataRangeWidth + 1);
          newFieldRange.setValue(fieldKeys[j]);
          // update fieldLocs and fieldLocation
          fieldLocs[fieldKeys[j]] = dataRangeWidth;
          fieldLocation = dataRangeWidth;
          // update dataRangeWidth
          dataRangeWidth += 1;
        }
        // add the data in the appropriate location in newRow
        newRow[fieldLocation] = inputData[i][fieldKeys[j]];
      }
  
      // store genral information at start of each row and in row dict
      // if we are creating a new row - not updating one - no ID value has been passed in so set it to the current ID
      if (!newRow[0]) {
        // generate a new id and insert it into the id list
        uniqueID = generateUniqueID_(sheet.getName().slice(0, 1));
        listIDs.push(uniqueID);
        newRow[0] = uniqueID; // ID
        row["ID"] = uniqueID; // ID
      } else {
        uniqueID = newRow[0];
      }
      // if we are creating a new row - not updating one - no createdBy value has been passed in so set it to the current user email 
      if (!newRow[1]) {
        newRow[1] = creator; // DateCreated
        row["CreatedBy"] = creator;
      }
      newRow[2] = creator; //DateModified - updated whenever a row is created - useful for updated rows
      row["ModifiedBy"] = creator;
      // if we are creating a new row - not updating one - no date created value has been passed in so set it to the current date/time 
      if (!newRow[3]) {
        newRow[3] = dateTime; // DateCreated
        row["DateCreated"] = dateTime;
      }
      newRow[4] = dateTime; // DateModified - updated whenever a row is created - useful for updated rows
      row["DateModified"] = dateTime;
      newRow[5] = true; // valid
      row["Valid"] = true;
      // add this row to the end of the sheet
      var newRange = sheet.getRange(sheet.getDataRange().getHeight() + 1, 1, 1, newRow.length);
      newRange.setValues([newRow]);
      rowDict[uniqueID] = row;
    }
    // return the dictionary of rows
    return rowDict;
  }
  
  // -------------------------------------------- READ -----------------------------------------------------
  
  /**
   * Function to return the rows specified by the given rowIDs
   *
   * @param {string} sheet: the instance of the sheet we wish to read from
   * @param {string} columnName: the name of the column that we are comparing values in for reading
   * @param {object} rowValues: a list of the values of the rows for the given columnName that we wish to read
   * @return {object} return a dictionary of dictionaries {ID: {fieldName: data}} corresponding to the inputted rowIDs
   *                  - if a rowID does not exist or is invalid, then the corresponding entry will be null
   *                  - if the rowIDs input is null itself, this function returns a dictionary corresponding to ALL rows in the sheet
   */
  function read_(sheet, columnName, rowValues) {
    Logger.log("Reading from the sheet: " + sheet.getName());
    //get 2D array of data from sheet to be searched
    var data = sheet.getDataRange().getValues();
    //get index of column with field name "name"
    var colIndex = getColIndex_(data, columnName);
    // check to make sure columnName is valid
    if (colIndex == -1) {
      Logger.log(columnName + " does not exist as a column name in " + sheet.getName());
      return;
    }
    // convert set for easier search
    var rowValueSet = new Set(rowValues);
    // itereate through data adding each row as a dictionary to parent dictionary
    var rowDict = {};
    var row;
    for (var i = 1; i < data.length; i++) {
      //check if the value under columnName for this row is in rowValues 
      if (rowValueSet.size == 0 || rowValueSet.has(data[i][colIndex])) {
        // convert row to a dictionary
        row = getRowAsDict(data[0], data[i]);
        // store row dictionary in a large dictionary where keys are IDs and values are the row dicts - to be used dealing FKeys
        if (row["ID"] && row["Valid"]) { // if the row exists, has an ID, and is valid - add to dictionary 
          rowDict[row["ID"]] = row;
        }
  
      }
    }
    return rowDict;
  }
  
  // ------------------------------------------- UPDATE ----------------------------------------------------
  
  /**
   * Function to 'Update' (delete and then create) the row with the oldRowID as its ID
   *
   * @param {string} sheet: the instance of the sheet we wish to update
   * @param {object} inputDict: a dictionary {fieldName: data} representing the data we wish to create a new row with - to be passed into create_
   * @return {object} returns the output of create_ - dictionary of dictionaries {ID: {fieldName: data}} corresponding to the inputted oldRowID
   */
  function update_(sheet, inputDict) {
    Logger.log("Updating: " + sheet.getName());
    // if  you are trying to update a row with no given information retun null
    if (!inputDict) {
      Logger.log("no input data given so could not update");
      return;
    }
    var oldRowID = inputDict["ID"];
    // delete the old row - returns nested dictionary so pull row dictionary out and store in deletedRow
    var deletedRow = delete_(sheet, "ID", [oldRowID])[oldRowID]; // note delete_ takes in a list of ID's hence [oldRowID]
    // if the row you are trying to update doesnt exist, log it, then just create the new row
    if (!deletedRow) {
      Logger.log("No row with ID: " + oldRowID + " exists so could not update");
      //return;
    } else {
      // pass on DateCreated and CreatedBy value from old row to new row
      inputDict["DateCreated"] = deletedRow["DateCreated"];
      inputDict["CreatedBy"] = deletedRow["CreatedBy"];
    }
    // create an updated row with up given information in inputDict
    var newRow = create_(sheet, [inputDict]); // note create_ takes in a list of dictionaries hence [inputDict]
    return newRow;
  }
  
  // ------------------------------------------- DELETE ----------------------------------------------------
  
  /**
   * Function to 'delete' (set valid=FALSE) all rows with the value of the given columnName equal to rowValue
   *
   * @param {string} sheet: the instance of the sheet we wish to delete from
   * @param {string} columnName: the name of the column that we are comparing values in for deletion
   * @param {object} rowValues: the values that we use to compare for deletion
   * @return {object} returns a dictionary of dictionaries {ID: {fieldName: data}} representing the deleted rows by their IDs
   *                  - returns null if inputs are invalid - bad column name, null rowValues or empty list, sheet missing a valid column 
   */
  function delete_(sheet, columnName, rowValues) {
    Logger.log("Deleting a row from the sheet: " + sheet.getName());
    // check validity of inputs
    if (!rowValues || rowValues == []) {
      Logger.log("Gave a null or empty rowValues in a call to delete_")
      return;
    }
    // Pull data from sheet
    var data = sheet.getDataRange().getValues();
    //get index of column with field name "name"
    var colIndex = getColIndex_(data, columnName);
    //get index of col with field name valid
    var validIndex = getColIndex_(data, "Valid");
    // get index of col with field name DateModified
    var modifiedIndex = getColIndex_(data, "DateModified");
    // check to make sure columnName is valid
    if (colIndex == -1) {
      Logger.log(columnName + " does not exist as a column name in " + sheet.getName());
      return;
    }
    // check to make sure valid is in sheet
    if (validIndex == -1) {
      Logger.log("Valid does not exist as a column name in " + sheet.getName());
      return;
    }
    // check to make sure valid is in sheet
    if (modifiedIndex == -1) {
      Logger.log("DateModified does not exist as a column name in " + sheet.getName());
      return;
    }
    // convert set for easier search
    var rowValueSet = new Set(rowValues);
    // itereate through data setting valid flags to false for each row that has rowValue at colIndex
    var rowDict = {};
    var row;
    for (var i = 1; i < data.length; i++) {
      //check if the value under columnName for this row is in rowValues 
      if (rowValueSet.has(data[i][colIndex])) {
        // set valid flag to false in data
        data[i][validIndex] = false;
        // update date modified
        data[i][modifiedIndex] = new Date();
        row = getRowAsDict(data[0], data[i]);
        // store row dictionary in a large dictionary where keys are IDs and values are the row dicts - to be used dealing FKeys
        if (row["ID"]) { // if the row exists and has an ID - row is correctly formatted so add to dictionary 
          rowDict[row["ID"]] = row;
        }
      }
    }
    // put updated data in sheet
    sheet.getDataRange().setValues(data);
    // return dictionary of deleted rows
    return rowDict;
  }
  
  // ----------------------------------------- UNDO DELETE --------------------------------------------------
  
  /**
   * Function to undo a delete operation (does not handle foreign keys)
   *
   * @param {object} sheet: the sheet that this operation is taking place on
   * @param {string} columnName: the field name for the column that we are using to determine which row to revalidate
   * @param {string} {object} rowValues: the values that we use to compare for undoing deletion
   * @return {object} returns a list of dictionaries of the rows that were just undeleted
   */
  function undoDelete_(sheet, columnName, rowValues) {
    Logger.log("Undoing delete from sheet " + sheet.getName());
    // Pull data from sheet
    var data = sheet.getDataRange().getValues();
    //get index of column with field name "name"
    var colIndex = getColIndex_(data, columnName);
    //get index of col with field name valid
    var validIndex = getColIndex_(data, "Valid");
    // check to make sure columnName is valid
    if (colIndex == -1) {
      Logger.log(columnName + " does not exist as a column name in " + sheet.getName());
      return;
    }
    // check to make sure valid is in sheet
    if (validIndex == -1) {
      Logger.log("Valid does not exist as a column name in " + sheet.getName());
      return;
    }
    // convert set for easier search
    var rowValueSet = new Set(rowValues);
    // initialize a list to store the rows that were undeleted
    var listDict = [];
    var row;
    // itereate through data setting valid flags to false for each row that has rowValue at colIndex
    // goes from the bottom up only changing the most recent version of a given entry
    for (var i = data.length - 1; i > -1; i--) {
      if (rowValueSet.has(data[i][colIndex])) {
        // set valid flag to true in data
        data[i][validIndex] = true;
        //listIDs.push(data[i][0]);
        rowValueSet.delete(data[i][colIndex]);
        // get row and store it in listDict
        row = getRowAsDict(data[0], data[i]);
        listDict.push(row);
      }
    }
    // put updated data in sheet
    sheet.getDataRange().setValues(data);
    // return the id list
    // return listIDs;
    return listDict;
  }
  
  // -------------------------------------- HELPER FUNCTIONS ------------------------------------------------
  
  /**
   * Function to generate a random primary key for a new row.
   *
   * @parameter {string} prefix: a character to be appended to the front of the key (for identification purposes)
   * @return {string} returns the key that it generated
   */
  function generateUniqueID_(prefix) {
    // generate two random number in base 32, append them and then take the first 16 bits 
    var randNum = (Math.random().toString(36).substring(2) + Math.random().toString(36).substring(2)).substring(0, 16)
    //append the sheet specific prefix to the beginning for more information
    return prefix + "-" + randNum;
  }
  
  /**
   * Function to get the datetime (default rendering: LA/pacific time, dates are not dependent on timezone though)
   *
   * @return {object} returns the current date as formatted by JavaScript
   */
  function getDatetime_() {
    return new Date();
  }
  
  /**
   * Function to return the email address of the current user
   *
   * @return {string} returns the email of the current user of the Web App
   */
  function getUserEmail_() {
    var email = Session.getActiveUser().getEmail()
    return email;
  }
  
  /**
   * Function to return the user name the current user
   *
   * @return {string} returns the email(can be modified) of the current user of the Web App 
   */
  function getUserName_() {
    return getUserEmail_();
  }
  
  /**
   * Function to check inputs to create_ are valid and fills in sheet if empty
   *
   * @parameter {object} sheet: the sheet we wish to write to
   * @parameter {object} inputData: list dictionaries array holding data we wish to write to the sheet
   * @return {bool} returns true if create has been given valid inputs and false otherwise
   */
  function validateCreateInputs(sheet, inputData) {
    // Check to see if the sheet is formatted correctly
    if (sheet.getRange(1, 1).getValue() != "ID" || sheet.getRange(1, 2).getValue() != "CreatedBy" || sheet.getRange(1, 3).getValue() != "ModifiedBy" || sheet.getRange(1, 4).getValue() != "DateCreated" || sheet.getRange(1, 5).getValue() != "DateModified" || sheet.getRange(1, 6).getValue() != "Valid") {
      // Check to see if whole sheet is null
      if (sheet.getDataRange().getValues().join("") === "") {
        // The sheet is null, so add the common columns and run the rest of the function
        sheet.getRange(1, 1).setValue("ID");
        sheet.getRange(1, 2).setValue("CreatedBy");
        sheet.getRange(1, 3).setValue("ModifiedBy");
        sheet.getRange(1, 3).setValue("DateCreated");
        sheet.getRange(1, 4).setValue("DateModified");
        sheet.getRange(1, 5).setValue("Valid");
      } else {
        // The sheet is formatted incorrectly, so log and throw an error and don't add any new rows
        var errorMsg = "The sheet you are trying to write to is formatted incorrectly. Please ensure that the first column of this sheet is the ID column. Bother Amir if you have any questions.";
        Logger.log(errorMsg);
        throw errorMsg;
        return false;
      }
    }
    if (inputData == null) {
      Logger.log("No data was provided to be inputted into the sheet");
      return false;
    }
    return true;
  }
  
  /**
   * Function to convert a row into a dictionary
   *
   * @parameter {object} keyList: list of keys to be put dictionary
   * @parameter {object} valueList: list of associated values to be put in dictionary
   * @return {object} returns a dictionary
   */
  function getRowAsDict(keyList, valueList) {
    var dictionary = {};
    for (var i = 0; i < keyList.length; i++) {
      if (valueList[i] != "") {
        dictionary[keyList[i]] = valueList[i];
      }
    }
    return dictionary;
  }
  
  /**
   * Helper function that finds the [0-INDEXED] row number containing id as its unique ID (-1 if it does not exist) (returns first instance if multiple)
   *
   * @parameter {object} data: data of the sheet to be searched
   * @parameter {string} rowVal: the value of an entry in the sheet at the given column index
   * @parameter {string} colIndex: the index of the column that we are searching for rowVal
   * @return {integer} index: the row index of the entry with the given ID
   *                          - returns -1 if no row has ID
   */
  function getRowIndex_(data, rowVal, colIndex) {
    // Find the row index of the requested ID
    for (var i = 0; i < data.length; i++) {
      if (data[i][colIndex] == rowVal) {
        return i;
      }
    }
    return -1;
  }
  
  /**
   * Helper function that finds the [0-INDEXED] column number with the name "name" (-1 if it does not exist)
   *
   * @parameter {object} data: data of a sheet to be searched
   * @parameter {string} id: the unique id of an entry in the sheet
   * @return {integer} index: the column index of the entry with the given ID
   *                          - returns -1 if no row has ID
   */
  function getColIndex_(data, name) {
    // Find the row index of the requested id
    for (var i = 0; i < data[0].length; i++) {
      if (data[0][i] == name) {
        return i;
      }
    }
    return -1;
  }
  
  // ------------------------------------------ CREATE SHEET ---------------------------------------------------
  
  /**
   * Function to add a new sheet to a spreadsheet and populate it with column names.
   *
   * @param {string} ssID: the id of the spreadsheet we wish to add the sheet to
   * @param {string} sheetName: the name we would like to give this sheet
   * @param {object} colNames: a list of the column names we wish to populate the sheet with
   * @return {string} returns the sheetName that was passed in
   */
  function createSheet_(ssID, sheetName, colNames) {
    // create a sheet in a given spread sheet with a given name
    var sheet = SpreadsheetApp.openById(ssID).insertSheet(sheetName);
    var headerRow = colNames;
    // add basic database fields
    headerRow.unshift("ID", "CreatedBy", "ModifiedBy", "DateCreated", "DateModified", "Valid");
    var range = sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
    range.setValues([headerRow]);
    //format header
    sheet.setFrozenRows(1);
    range.setFontWeight("bold");
    //be mindful not to have every row added after to keep the formatting!
  
    return sheetName
  }
  
  
  // ------------------------------------------ CLEAN DATABASE ---------------------------------------------------
  
  /**
   * Function to remove every false row in the database that is atleast a day old
   */
  function cleanDatabase() {
    // get every SSID
    const clientSSID = PropertiesService.getScriptProperties().getProperty("clientSpreadsheetID");
    const memberSSID = PropertiesService.getScriptProperties().getProperty("memberSpreadsheetID");
    const matterSSID = PropertiesService.getScriptProperties().getProperty("matterSpreadsheetID");
    const taskSSID = PropertiesService.getScriptProperties().getProperty("taskSpreadsheetID");
    const noteSSID = PropertiesService.getScriptProperties().getProperty("noteSpreadsheetID");
    // get every spreadsheet
    var clientSS = SpreadsheetApp.openById(clientSSID);
    var memberSS = SpreadsheetApp.openById(memberSSID);
    var matterSS = SpreadsheetApp.openById(matterSSID);
    var taskSS = SpreadsheetApp.openById(taskSSID);
    var noteSS = SpreadsheetApp.openById(noteSSID);
    // put every sheet in a list
    var sheetList = clientSS.getSheets();
    sheetList = sheetList.concat(memberSS.getSheets());
    sheetList = sheetList.concat(matterSS.getSheets());
    sheetList = sheetList.concat(taskSS.getSheets());
    sheetList = sheetList.concat(noteSS.getSheets());
    var sheetList = sheetList.concat(taskSS.getSheets());
    // get data range of each sheet
    var sheet, data, row, validIndex, dateIndex, modifiedDate;
    var newData;
    var today = new Date();
    var numRowsDeleted = 0;
    for (var i = 0; i < sheetList.length; i++) {
      newData = [];
      sheet = sheetList[i];
      data = sheet.getDataRange().getValues();
      //get index of col with field name valid
      validIndex = getColIndex_(data, "Valid");
      dateIndex = getColIndex_(data, "DateModified");
      newData.push(data[0]);
      // remove every row that is false and is a day old
      for (var j = 1; j < data.length; j++) {
        row = data[j];
        modifiedDate = new Date(row[dateIndex])
        if (row[validIndex] == true || (today - modifiedDate) < 86400000) {
          newData.push(row);
        }
      }
      // clear the sheet and put updated data back in
      sheet.clear({ contentsOnly: true });
      sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
      numRowsDeleted += (data.length - newData.length);
    }
    return numRowsDeleted;
  }