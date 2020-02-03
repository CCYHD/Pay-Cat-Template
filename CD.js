/**
 * A collection of useful methods for Apps Script
 */
var CD = {
  ss: SpreadsheetApp.getActive(),

  /**
  * Sets CD's active spreadsheet to the given ID
  *
  * @param {string} ssId - The ID of the spreadsheet to be used
  * @returns {Spreadsheet} The spreadsheet object
  */
  setSS: function(ssId) {
    this.ss = SpreadsheetApp.openById(ssId);
    return this.ss;
  },
  
  // ---------------------------------- Sheet Input ----------------------------------

  /**
   * Gets the sheet with the given name
   * 
   * @param {string} name - The name of the sheet to get
   * @returns {Sheet} The App Script sheet object
   */
  getSheet: function(name) {
    var sheet = this.ss.getSheetByName(name);
    return sheet;
  },
  
  /**
  * Gets the rows of a sheet as a set of mapped rows
  * 
  * @param {(Sheet|string)} sheet - The sheet object or name of the sheet
  * @returns {array[]} An array of row objects
  */
  getRows: function(sheet) {
    var rows = this.getTable(sheet);
    return this.mapRows(rows);
  },

  /**
  * Gets the rows of a sheet as a 2D array
  * 
  * @param {(Sheet|string)} sheet - The sheet object or name of the sheet
  * @returns {array[]} An array of row arrays
  */
  getTable: function(sheet) {
    if (typeof sheet == "string") {
      sheet = this.getSheet(sheet);
    }
    var rows = sheet.getDataRange().getDisplayValues();
    return rows;
  },
  
  /**
   * Converts an array of rows to array of objects (using header row for properties)
   * 
   * @param {array[]} rows - An array of row arrays
   * @returns {object[]} An array of objects
   */
  mapRows: function(rows) {
    var props = rows.shift();
    var objs = [];
    for (var i in rows) {
      var obj = {};
      for (var j in props) {
        obj[props[j]] = rows[i][j];
      }
      objs.push(obj);
    }
    rows.unshift(props);
    return objs;
  },
  
  /**
   * Combines an array of rows into a mapped obj using column with header "idColName"
   * 
   * @param {array[]} rows - An array of row objects
   * @param {string} idColName - The title from the header row to merge entries based on
   * @param {boolean} combineEntries - Whether to use the last entry found, or generate an array of matching entries
   * @returns {object} An object with a property for each unique entry in the id column, each of which contains an array of mapped rows
   */
  rows2obj: function(rows, idColName, combineEntries) {
    var combinedObj = {};
    for (var i in rows) {
      var obj = rows[i];
      var id = obj[idColName];
      if (combineEntries) {
        if (combinedObj[id] === undefined) {
          combinedObj[id] = [];
        }
        combinedObj[id].push(obj);
      } else {
        combinedObj[id] = obj;
      }
    }
    return combinedObj;
  },
  
  /**
   * Gets all rows from a given sheet and returns a mapped object using a specified column to ID unique entries
   * 
   * @param {string|Sheet} sheet - The sheet object, or name of the sheet
   * @param {string} idColName - The title from the header row to merge entries based on
   * @param {boolean} combineEntries - If true, return an array of all entries matching each id, otherwise return last entries of each id
   * @returns {object} An object with a property for each unique entry in the id column, each of which contains an array of mapped rows
   */
  getObj: function(sheet, idColName, combineEntries) {
    var rows = this.getRows(sheet);
    if (combineEntries) {
      var obj = this.rows2obj(rows, idColName, true);
    } else {
      var obj = this.rows2obj(rows, idColName, false);
    }
    return obj;
  },

  // ---------------------------------- Sheet Printing ----------------------------------

  /**
   * Generates a new sheet with the specified name (or clears it if it already exists)
   * 
   * @param {string} name - The name of the sheet to be created
   * @returns {Object.Sheet} The new sheet
   */
  makeNewSheet: function(name) {
    if (name.length > 100) {
      name = name.slice(0, 99);
    }
    var newSheet = this.ss.getSheetByName(name);
    if (newSheet != null) {
      newSheet.clear();
      return newSheet;
    }
    newSheet = this.ss.insertSheet();
    newSheet.setName(name);
    return newSheet;
  },

  /**
   * Appends a 2D array onto the bottom of sheet
   * 
   * @param {Sheet} sheet - The sheet to append rows onto
   * @param {array[]} rows - An array of row arrays (all of equal length) to append onto sheet
   * @returns {Sheet} The sheet that was used
   */
  appendRows: function(sheet, rows) {
    var lastRow = sheet.getLastRow();
    if (rows.length > 0) {
      var range = sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length);
      range.setValues(rows);
    }
    return sheet;
  },

  /**
   * Remove excess rows and columns from sheet
   * @param {Sheet} sheet - The sheet object to be trimmed
   * @returns {Sheet} The sheet that was used
   */
  trimSheet: function(sheet) {
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    var blankRows = sheet.getMaxRows() - lastRow;
    var blankColumns = sheet.getMaxColumns() - lastColumn;
    if (blankRows > 0) {sheet.deleteRows(lastRow + 1, blankRows)};
    if (blankColumns > 0) {sheet.deleteColumns(lastColumn + 1, blankColumns)};
    return sheet;
  },

  /**
   * Applies horizonal / vertical borders to all cells in a set of rows / columns
   * 
   * @param {Sheet} sheet - The sheet object to add borders to
   * @param {string} rowVsCol - "row" or "col" to determine the direction of the borders
   * @param {string} side - "top", "left", "bottom" or "right" to determine which side to apply the borders to
   * @param {(number|number[])} index - The index (or set of indices) to apply the borders to (starting at 1)
   * @returns {Sheet} The sheet that was used
   */
  insertBorders: function(sheet, rowVsCol, side, index) {
    if (typeof sheet == "string") {
      sheet = this.makeNewSheet(sheet);
    }
    var borderList = {
      "top": [true, null, null, null, null, null],
      "left": [null, true, null, null, null, null],
      "bottom": [null, null, true, null, null, null],
      "right": [null, null, null, true, null, null]
    }
    var b = borderList[side];
    if (typeof index == "object") {
      for (var i in index) {
        if (rowVsCol == "row") {
          var range = sheet.getRange(index[i], 1, 1, sheet.getMaxColumns());
        } else {
          var range = sheet.getRange(1, index[i], sheet.getMaxRows());
        }
        range.setBorder(b[0], b[1], b[2], b[3], b[4], b[5]);
      }
    } else {
      if (rowVsCol == "row") {
        var range = sheet.getRange(index, 1, 1, sheet.getMaxColumns());
      } else {
        var range = sheet.getRange(1, index, sheet.getMaxRows());
      }
      range.setBorder(b[0], b[1], b[2], b[3], b[4], b[5]);
    }
    return sheet;
  },

  /**
   * Makes a sheet more presentable by bolding and adding a border to its top row, resizing its columns and trimming any excess space
   * @param {Sheet} sheet - The sheet to be modified
   * @returns {Sheet} The sheet that was used
   */
  prettyUp: function(sheet) {
    sheet.getRange(1, 1, 1, sheet.getMaxColumns()).setFontWeight("bold");
    this.insertBorders(sheet, "row", "bottom", 1);
    sheet.autoResizeColumns(1, sheet.getMaxColumns());
    this.trimSheet(sheet);
  },
  
  /**
  * Create a new sheet, fill it with a data table, then clean up its formatting
  *
  * @param {string} sheetName - The name of the sheet to be created/cleared out
  * @param {array[]} rows - An array of rows (arrays of equal length) to print to the sheet
  * @returns {Sheet} The generated sheet
  */
  printTable: function(sheetName, table) {
    var sheet = this.makeNewSheet(sheetName);
    this.appendRows(sheet, table);
    this.prettyUp(sheet);
    return sheet;
  },

  /**
  * Create a new sheet, fill it with data from rows, then clean up its formatting
  *
  * @param {string} sheetName - The name of the sheet to be created/cleared out
  * @param {object[]} rows - An array of objects to print to the sheet
  * @returns {Sheet} The generated sheet
  */
  printRows: function(sheet, rows) {
    var table = this.tabulateArray(rows);
    return this.printTable(sheet, table);
  },
  
  // ---------------------------------- Other Output ----------------------------------

  /**
   * Generate a 2D array from  an array of objects (e.g. earnings lines)
   * @param {object[]} array - An array of objects to be tabulated
   * @param {string[]=} headerRow - A row of strings to use for the header row. Each string should match a property name from the object. If left out, the header row is populated with the first objects properties
   * @returns {array[]} An array of rows
   */
  tabulateArray: function(array, headerRow) {
    if (!headerRow) {
      var headerRow = [];
      for (var i in array) {
        for (var property in array[i]) {
          if (headerRow.indexOf(property) == -1) {
            headerRow.push(property);
          }
        }
      }
    }

    var rows = [headerRow];
    for (var i in array) {
      var row = [];
      for (var j in headerRow) {
        var val = array[i][headerRow[j]] === undefined ? "" : array[i][headerRow[j]];
        row.push(val);
      }
      rows.push(row);
    }
    return rows;
  },

  deleteAllSheets: function(ss) {
    ss = ss == undefined ? SpreadsheetApp.getActiveSpreadsheet() : ss;
    var sheets = ss.getSheets();
    for (var i in sheets) {
      if (sheets[i].getName() != "EOF") {
        ss.deleteSheet(sheets[i]);
      }
    }
  },

  /**
   * 
   * @param {Sheet} sheet - The sheet to be used
   * @param {string} range - The range to apply the formatting to
   * @param {string} formula - The formula to be applied
   * @param {string} colour - The colour to apply when the formula is satisfied
   */
  applyConditionalFormatting: function(sheet, range, formula, colour) {
    var range = sheet.getRange(range);
    var rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setBackground(colour)
      .setRanges([range])
      .build();
    var rules = sheet.getConditionalFormatRules();
    rules.push(rule);
    sheet.setConditionalFormatRules(rules);
    return sheet;
  },
  
  /**
   * Add a new column and populate its values
   * @param {Sheet} sheet - The sheet to be used
   * @param {array[]} column - An array of 1 length arrays (representing the column) to be inserted
   * @param {number} index - The column index (starting from 1) for the new column to be inserted after
   */
  addColumn: function(sheet, column, index) {
    sheet.insertColumnAfter(index);
    var range = sheet.getRange(1, +index + 1, column.length, 1);
    range.setValues(column);
  },

  /**
   * Enters a value into a cell using A1 notation
   * @param {Sheet} sheet - The sheet to be editted
   * @param {string} cellRange - The cell index to be addressed (in A1 notation)
   * @param {string} val - The value to be entered
   */
  setCellA1: function(sheet, cellRange, val) {
    var range = sheet.getRange(cellRange);
    range.setValue(val);
  },

  getCellA1: function (sheet, cellRange) {
    var range = sheet.getRange(cellRange);
    var vals = range.getValues();
    return vals[0][0];
  },

  displayDialog: function(text, title, height, width) {
    var htmlOutput = HtmlService.createHtmlOutput(text);
    if (height) {
      htmlOutput.setHeight(height);
    }
    if (width) {
      htmlOutput.setWidth(width);
    }
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, title);
  },

  /**
   * Opens a HTML dialog which diplays an object
   * @param {Object} obj Object to be displayed
   * @param {String} [title] Title of dialog
   */
  displayObject: function(obj, title) {
    title = title ? title : "";
    this.displayDialog("<pre><code>" + JSON.stringify(obj, null, 2) + "</pre></code>", title, 800, 800);
  },
  
  /**
   * Generates a hash map
   * @param {Object[]} array Array of objects to be mapped
   * @param {String} idProperty The property to be used to identify each item
   * @param {String} [returnProperty] The property to map to (if not specified, the full object will be returned)
   */
  hash: function(array, idProperty, returnProperty) {
    var hashMap = {};
    array.forEach(function(item) {
      hashMap[item[idProperty]] = returnProperty ? item[returnProperty] : item;
    });
    return hashMap;
  },

  findIn: function(array, propertyName, value) {
    for (var i in array) {
      if (array[i][propertyName] == value) {
        return array[i];
      }
    }
    return false;
  },
  
  // ---------------------------------- Date Manipulation ----------------------------------

  toAusDate: function(date) {
    return date.getDate() + "/" + (+date.getMonth() + 1) + "/" + date.getYear();
  },


  // ---------------------------------- Drive Stuff ----------------------------------

  getFileId: function(folder, fileName) {
    if (typeof folder == "string") {
      folder = DriveApp.getFolderById(folder);
    }
    var files = folder.getFilesByName(fileName);
    if (files.hasNext()) {
      var file = files.next();
      return file.getId();
    } else {
      return "File not Found";
    }
  }
};


