function GetSheetById(sheetID){
  return SpreadsheetApp.openById(sheetID);;
}
function GetActiveSheet(){
  return SpreadsheetApp.getActiveSpreadsheet();
}
/**
 * 
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} sheet 
 * @param {*} tabName 
 * @param {*} headers 
 */
function GoogleSheet(sheet, tabName, headers) {
  /**
   * Name of the person.
   * @name GoogleSheet#sheet
   * @type {GoogleAppsScript.Spreadsheet.Sheet}
   */
  this.sheet = sheet.getSheetByName(tabName);
  this.headers = headers;
  var headersIsArray = Object.prototype.toString.call(this.headers) == '[object Array]';
  if (headersIsArray) {
    this.initHeaders();
  }
}
GoogleSheet.prototype.getRange = function (range) {
  var range = this.sheet.getRange(range);
  var values = range ? range.getValues() : [];
  return values;
}
GoogleSheet.prototype.setHeaders = function (headers) {
  this.headers = headers;
}

GoogleSheet.prototype.initHeaders = function () {
  var actual_headers = this.getSpreadsheetHeaders();
  var validated_count = 0;
  for (var key in actual_headers) {
    if (actual_headers[key] == this.headers[key]) {
      validated_count++;
    }
  }

  //If no headers are found or not all rows are validated, append a header row
  if (actual_headers.length == 0 || validated_count != this.headers.length) {
    this.sheet.appendRow(this.headers);
  }
  return true;
}

GoogleSheet.prototype.getSpreadsheetHeaders = function () {
  var range = this.sheet.getRange("A1:ZZ1");
  var values = range ? range.getValues()[0] : [];
  for (var key in values) {
    if (values[key] === "") {
      delete values[key];
    }
  }
  return values;

}

GoogleSheet.prototype.appendObject = function (object) {
  var row = [];
  var isArray = Object.prototype.toString.call(this.headers) == '[object Array]';
  if (isArray) {
    for (var index in this.headers) {
      key = this.headers[index];
      row[index] = object[key];
    }
  } else {
    for (var key in object) {
      row.push(object[key]);
    }
  }
  this.sheet.appendRow(row);
}
GoogleSheet.prototype.findOne = function (query) {
  var self = this;
  var search_result = null;
  //Search info is an array of objects of the format {col: i, value: v}, where i is the column to search and v is the value to be found
  var search_info = Object.keys(query).map(
    function(key) {
      var value = query[key];
      var col_idx = self.headers.indexOf(key);
      return {
        col: col_idx,
        value: value
      }
    }
  );

  //Run through all rows in the current sheet and check for the rows that match
  var numrows = this.sheet.getLastRow() - 1; //Subtract 1 to remove the header row
  var numcols = this.sheet.getLastColumn();
  if(numrows < 1) {
    return []; //Do not attempt to query the seelt
  }
  var range = this.sheet.getRange(2, 1, numrows, numcols); //Start at 2 to exclude header row
  var data = range.getValues();

  for(var row_idx in data){
    var row = data[row_idx];
    var matches = 0;
    //check that all the search criteria are matched
    for(var search_info_idx in search_info) {
      var check = search_info[search_info_idx];
      var col = check.col;
      if(row[col] == check.value) {
        matches++;
      }
    }
    //If all criteria are matched, assemble the record as an object and append to search results array
    if(matches === search_info.length) {
      var record = {};
      for (var idx in this.headers) {
        var key = this.headers[idx];
        record[key] = row[idx];
      }
      search_result = record;
      break;
    }
  }
  return search_result;
}

GoogleSheet.prototype.find = function (query) {
  var self = this;
  var search_results = [];
  //Search info is an array of objects of the format {col: i, value: v}, where i is the column to search and v is the value to be found
  var search_info = Object.keys(query).map(
    function(key) {
      var value = query[key];
      var col_idx = self.headers.indexOf(key);
      return {
        col: col_idx,
        value: value
      }
    }
  );

  //Run through all rows in the current sheet and check for the rows that match
  var numrows = this.sheet.getLastRow() - 1; //Subtract 1 to remove the header row
  var numcols = this.sheet.getLastColumn();
  if(numrows < 1) {
    return []; //Do not attempt to query the sheet if there are no rows
  }
  var range = this.sheet.getRange(2, 1, numrows, numcols); //Start at 2 to exclude header row
  var data = range.getValues();

  for(var row_idx in data){
    var row = data[row_idx];
    var matches = 0;
    //check that all the search criteria are matched
    for(var search_info_idx in search_info) {
      var check = search_info[search_info_idx];
      var col = check.col;
      if(row[col] == check.value) {
        matches++;
      }
    }
    //If all criteria are matched, assemble the record as an object and append to search results array
    if(matches === search_info.length) {
      var record = {};
      for (var idx in this.headers) {
        var key = this.headers[idx];
        record[key] = row[idx];
      }
      search_results.push(record);
    }
  }
  return search_results;
}

/**
 * Interates through table records to find all elements that match the given condition
 * @param {Function} isAllowed
 */
GoogleSheet.prototype.filter = function (isAllowed) {
  var search_results = [];

  //Run through all rows in the current sheet and check for the rows that match
  var numrows = this.sheet.getLastRow() - 1; //Subtract 1 to remove the header row
  var numcols = this.sheet.getLastColumn();
  if(numrows < 1) {
    return []; //Do not attempt to query the sheet if there are no rows
  }
  var range = this.sheet.getRange(2, 1, numrows, numcols); //Start at 2 to exclude header row
  var data = range.getValues();

  for (var row_idx in data) {
    var row = data[row_idx];
    //Assemble the row into an object
    var elem = {};
    for (var idx in this.headers) {
      var key = this.headers[idx];
      elem[key] = row[idx];
    }

    if (isAllowed(elem, row_idx, data)) {
      search_results.push(elem);
    }
  }
  return search_results;
}

GoogleSheet.prototype.filterRaw = function (isAllowed) {
  var search_results = [];
  //Run through all rows in the current sheet and check for the rows that match
  var numrows = this.sheet.getLastRow() - 1; //Subtract 1 to remove the header row
  var numcols = this.sheet.getLastColumn();
  if(numrows < 1) {
    return []; //Do not attempt to query the sheet if there are no rows
  }
  var range = this.sheet.getRange(2, 1, numrows, numcols); //Start at 2 to exclude header row
  var data = range.getValues();

  for (var row_idx in data) {
    var row = data[row_idx];
    if (isAllowed(row, row_idx, data)) {
      search_results.push(row);
    }
  }
  return search_results;
}

GoogleSheet.prototype.lookup = function (header, value, property) {
  var numrows = this.sheet.getMaxRows();
  var prop_index = this.headers.indexOf(property);
  if (prop_index == -1) {
    return; //If the header or property is not even found, return undefined
  }
  var propertyrange = this.sheet.getRange(1, prop_index + 1, numrows);
  var properties = propertyrange.getValues();
  var row = this.rowIndexForValue(header, value);
  return properties[row][0];
}
/**
 * Gets the raw data of the sheet at a given row index
 */
GoogleSheet.prototype.rowAtIndex = function (row) {
  var numcols = this.headers.length;
  var range = this.sheet.getRange(row + 1, 1, 1, numcols);
  var values = range.getValues();
  return values[0];
}

//Returns 0 based index
GoogleSheet.prototype.rowIndexForValue = function (header, value) {
  var search_index = this.headers.indexOf(header);
  if (search_index == -1) {
    return -1; //If the header is not even found, return undefined
  }
  var numrows = this.sheet.getMaxRows();
  var searchrange = this.sheet.getRange(1, search_index + 1, numrows);
  var search = searchrange.getValues();
  var found = -1;
  //arr_item is an array with one item in it, representing a single cell row
  for (var key in search) {
    var q = search[key][0];
    if (q == value) {
      found = key;
      break;
    }
  }
  return parseInt(found);
}

GoogleSheet.prototype.getLastRow = function () {
  var numrows = this.sheet.getLastRow();
  return numrows;
}
GoogleSheet.prototype.map = function (callback) {
  var headers = this.headers;
  var numrows = this.sheet.getLastRow();
  for (var i = 1; i < numrows; i++) { // Start at 1 to exclude headers
    var row = this.rowAtIndex(i);
    var object = {};
    //Packge row as data
    for (var j = 0; j < headers.length; j++) {
      var key = headers[j];
      var value = row[j];
      object[key] = value;
    }
    //Callback
    callback(object, i - 1);

  }
}