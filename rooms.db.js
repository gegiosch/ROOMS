var ROOMS_APP = ROOMS_APP || {};

ROOMS_APP.DB = {
  getSheet: function (sheetName) {
    return ROOMS_APP.openSpreadsheet().getSheetByName(sheetName);
  },

  getOrCreateSheet: function (sheetName) {
    return this.getSheet(sheetName) || ROOMS_APP.openSpreadsheet().insertSheet(sheetName);
  },

  getHeaders: function (sheetName) {
    var sheet = this.getSheet(sheetName);
    if (!sheet || sheet.getLastRow() < 1) {
      return [];
    }
    return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  },

  readRows: function (sheetName) {
    var sheet = this.getSheet(sheetName);
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }

    var headers = this.getHeaders(sheetName);
    var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
    var timezone = ROOMS_APP.DEFAULT_TIMEZONE;

    return values.map(function (row) {
      var entry = {};
      headers.forEach(function (header, index) {
        entry[header] = ROOMS_APP.DB.normalizeCellValue_(header, row[index], timezone);
      });
      return entry;
    });
  },

  appendRows: function (sheetName, rows) {
    if (!rows || !rows.length) {
      return;
    }

    var headers = this.getHeaders(sheetName);
    var sheet = this.getSheet(sheetName);
    var values = rows.map(function (row) {
      return headers.map(function (header) {
        return Object.prototype.hasOwnProperty.call(row, header) ? row[header] : '';
      });
    });

    sheet.getRange(sheet.getLastRow() + 1, 1, values.length, headers.length).setValues(values);
  },

  replaceRows: function (sheetName, headers, rows) {
    var sheet = this.getSheet(sheetName);
    if (!sheet) {
      throw new Error('Sheet not found: ' + sheetName);
    }

    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(sheet.getLastColumn(), headers.length)).clearContent();
    }

    if (rows && rows.length) {
      var values = rows.map(function (row) {
        return headers.map(function (header) {
          return Object.prototype.hasOwnProperty.call(row, header) ? row[header] : '';
        });
      });
      sheet.getRange(2, 1, values.length, headers.length).setValues(values);
    }
  },

  upsertByKey: function (sheetName, keyField, rows) {
    if (!rows || !rows.length) {
      return;
    }

    var headers = this.getHeaders(sheetName);
    var existingRows = this.readRows(sheetName);
    var existingMap = {};
    existingRows.forEach(function (row, index) {
      existingMap[row[keyField]] = { row: row, rowNumber: index + 2 };
    });

    var sheet = this.getSheet(sheetName);
    var pendingAppend = [];

    rows.forEach(function (row) {
      var key = row[keyField];
      if (Object.prototype.hasOwnProperty.call(existingMap, key)) {
        var values = headers.map(function (header) {
          return Object.prototype.hasOwnProperty.call(row, header) ? row[header] : '';
        });
        sheet.getRange(existingMap[key].rowNumber, 1, 1, headers.length).setValues([values]);
      } else {
        pendingAppend.push(row);
      }
    });

    this.appendRows(sheetName, pendingAppend);
  },

  normalizeCellValue_: function (header, value, timezone) {
    if (value instanceof Date) {
      if (/ISO$/i.test(header)) {
        return Utilities.formatDate(value, timezone, "yyyy-MM-dd'T'HH:mm:ss");
      }
      if (/Time$/i.test(header)) {
        return Utilities.formatDate(value, timezone, 'HH:mm');
      }
      if (/Date/i.test(header)) {
        return Utilities.formatDate(value, timezone, 'yyyy-MM-dd');
      }
      return Utilities.formatDate(value, timezone, "yyyy-MM-dd'T'HH:mm:ss");
    }

    return value;
  }
};
