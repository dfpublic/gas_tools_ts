/**
 * Gets a sheet by the sheet id
 * @param {*} sheetID 
 */
export function GetSheetById(sheetID: string) {
  return SpreadsheetApp.openById(sheetID);;
}

/**
 * Gets the currently active spreadsheet
 */
export function GetActiveSheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}
/**
 * 
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet 
 * @param {string} tabName 
 * @param {Array<string>} headers 
 */
export class GoogleSheet {
  /**
   * Name of the person.
   * @name GoogleSheet#sheet
   * @type {GoogleAppsScript.Spreadsheet.Sheet}
   */
  sheet: GoogleAppsScript.Spreadsheet.Sheet;
  headers: string[];
  constructor(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet, tabName: string, headers: Array<string>) {
    this.sheet = spreadsheet.getSheetByName(tabName);
    this.headers = headers;
    var headersIsArray = Object.prototype.toString.call(this.headers) == '[object Array]';
    if (headersIsArray) {
      this.initHeaders();
    }
  }

  /**
   * Gets the values within a given range
   */
  getRange(range: string) {
    var sheet_range = this.sheet.getRange(range);
    var values = sheet_range ? sheet_range.getValues() : [];
    return values;
  }
  setHeaders(headers: Array<string>) {
    this.headers = headers;
  }

  initHeaders() {
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

  getSpreadsheetHeaders() {
    var max_cols = this.sheet.getMaxColumns();
    var range = this.sheet.getRange(1, 1, 1, max_cols);
    var values = range ? range.getValues()[0] : [];
    for (var key in values) {
      if (values[key] === "") {
        delete values[key];
      }
    }
    return values;

  }

  appendObject(object: any, options: { background?: string } = {}) {
    var row_idx = this.sheet.getLastRow() + 1;
    var row_width = this.headers.length;

    let { background } = options;

    var row: Array<any> = [];
    var isArray = Object.prototype.toString.call(this.headers) == '[object Array]';
    if (isArray) {
      for (var index in this.headers) {
        key = this.headers[index];
        var value = (object[key] === undefined) ? null : object[key];
        row[index] = value;
      }
    } else {
      for (var key in object) {
        row.push(object[key]);
      }
    }
    this.sheet.appendRow(row);
    if (background) {
      this.sheet.getRange(row_idx, 1, 1, row_width).setBackground(background);
    }
  }
  findOne(query: any) {
    var self = this;
    var search_result = null;
    //Search info is an array of objects of the format {col: i, value: v}, where i is the column to search and v is the value to be found
    var search_info = Object.keys(query).map(
      function (key) {
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
    if (numrows < 1) {
      return []; //Do not attempt to query the seelt
    }
    var range = this.sheet.getRange(2, 1, numrows, numcols); //Start at 2 to exclude header row
    var data = range.getValues();

    for (var row_idx in data) {
      var row = data[row_idx];
      var matches = 0;
      //check that all the search criteria are matched
      for (var search_info_idx in search_info) {
        var check = search_info[search_info_idx];
        var col = check.col;
        if (row[col] == check.value) {
          matches++;
        }
      }
      //If all criteria are matched, assemble the record as an object and append to search results array
      if (matches === search_info.length) {
        var record: any = {};
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

  find(query: any) {
    var self = this;
    var search_results = [];
    //Search info is an array of objects of the format {col: i, value: v}, where i is the column to search and v is the value to be found
    var search_info = Object.keys(query).map(
      function (key) {
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
    if (numrows < 1) {
      return []; //Do not attempt to query the sheet if there are no rows
    }
    var range = this.sheet.getRange(2, 1, numrows, numcols); //Start at 2 to exclude header row
    var data = range.getValues();

    for (var row_idx in data) {
      var row = data[row_idx];
      var matches = 0;
      //check that all the search criteria are matched
      for (var search_info_idx in search_info) {
        var check = search_info[search_info_idx];
        var col = check.col;
        if (row[col] == check.value) {
          matches++;
        }
      }
      //If all criteria are matched, assemble the record as an object and append to search results array
      if (matches === search_info.length) {
        var record: any = {};
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
   * @callback FilterCallback
   * @param {*} object
   */
  /**
   * Iterates through table records to find all elements that match the given condition
   * @param {FilterCallback} fn
   */
  filter(fn: Function) {
    var search_results = [];

    //Run through all rows in the current sheet and check for the rows that match
    var numrows = this.sheet.getLastRow() - 1; //Subtract 1 to remove the header row
    var numcols = this.sheet.getLastColumn();
    if (numrows < 1) {
      return []; //Do not attempt to query the sheet if there are no rows
    }
    var range = this.sheet.getRange(2, 1, numrows, numcols); //Start at 2 to exclude header row
    var data = range.getValues();

    for (var row_idx in data) {
      var row = data[row_idx];
      //Assemble the row into an object
      var elem: any = {};
      for (var idx in this.headers) {
        var key = this.headers[idx];
        elem[key] = row[idx];
      }

      if (fn(elem, row_idx, data)) {
        search_results.push(elem);
      }
    }
    return search_results;
  }

  filterRaw(isAllowed: Function) {
    var search_results = [];
    //Run through all rows in the current sheet and check for the rows that match
    var numrows = this.sheet.getLastRow() - 1; //Subtract 1 to remove the header row
    var numcols = this.sheet.getLastColumn();
    if (numrows < 1) {
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

  lookup(header: any, value: any, property: any) {
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
  rowAtIndex(row: number) {
    var numcols = this.headers.length;
    var range = this.sheet.getRange(row + 1, 1, 1, numcols);
    var values = range.getValues();
    return values[0];
  }

  /**
   * Get any metadata for the row to be propagated
   * @param row 
   */
  metaAtIndex(row: number) {
    var numcols = this.headers.length;
    var range = this.sheet.getRange(row + 1, 1, 1, numcols);
    var background = range.getBackground();
    return {
      background
    };

  }

  //Returns 0 based index
  rowIndexForValue(header: string, value: string) {
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
        found = parseInt(key);
        break;
      }
    }
    return found;
  }

  getLastRow() {
    var numrows = this.sheet.getLastRow();
    return numrows;
  }
  map(callback: Function) {
    var headers = this.headers;
    var numrows = this.sheet.getLastRow();
    let results = [];
    for (var i = 1; i < numrows; i++) { // Start at 1 to exclude headers
      var row = this.rowAtIndex(i);
      var ___meta = this.metaAtIndex(i); //Propogate metadata
      var object: any = { ___meta };
      //Package row as data
      for (var j = 0; j < headers.length; j++) {
        var key = headers[j];
        var value = row[j];
        object[key] = value;
      }
      //Callback
      let new_val = callback(object, i - 1);
      results.push(new_val);
    }
    return results;
  }
}