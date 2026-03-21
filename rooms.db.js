var ROOMS_APP = ROOMS_APP || {};

ROOMS_APP.DB = {
  STATIC_SHEET_CACHE_TTLS_: (function () {
    var ttls = {};
    ttls[ROOMS_APP.SHEET_NAMES.CONFIG] = 60;
    ttls[ROOMS_APP.SHEET_NAMES.RESOURCES] = 120;
    ttls[ROOMS_APP.SHEET_NAMES.WEEK_SCHEDULE] = 120;
    ttls[ROOMS_APP.SHEET_NAMES.SPECIAL_OPENINGS] = 60;
    ttls[ROOMS_APP.SHEET_NAMES.CLOSURES] = 60;
    ttls[ROOMS_APP.SHEET_NAMES.HOLIDAYS] = 300;
    ttls[ROOMS_APP.SHEET_NAMES.ADMINS] = 60;
    ttls[ROOMS_APP.SHEET_NAMES.REPL_CLASS_OUT] = 60;
    ttls[ROOMS_APP.SHEET_NAMES.REPL_DAY_TEACHERS] = 60;
    ttls[ROOMS_APP.SHEET_NAMES.REPL_ASSIGNMENTS] = 60;
    ttls[ROOMS_APP.SHEET_NAMES.REPL_LONG_ASSIGNMENTS] = 120;
    ttls[ROOMS_APP.SHEET_NAMES.REPORT_RECIPIENTS] = 120;
    return ttls;
  }()),

  beginRequestCache_: function () {
    ROOMS_APP.DB_REQUEST_CACHE_ = {
      rows: {},
      headers: {},
      stats: {
        requestHits: 0,
        scriptHits: 0,
        misses: 0
      }
    };
  },

  endRequestCache_: function () {
    ROOMS_APP.DB_REQUEST_CACHE_ = null;
  },

  getRequestStats_: function () {
    var stats = ROOMS_APP.DB_REQUEST_CACHE_ && ROOMS_APP.DB_REQUEST_CACHE_.stats;
    return {
      requestHits: stats ? stats.requestHits : 0,
      scriptHits: stats ? stats.scriptHits : 0,
      misses: stats ? stats.misses : 0
    };
  },

  getSheet: function (sheetName) {
    return ROOMS_APP.openSpreadsheet().getSheetByName(sheetName);
  },

  getOrCreateSheet: function (sheetName) {
    return this.getSheet(sheetName) || ROOMS_APP.openSpreadsheet().insertSheet(sheetName);
  },

  getHeaders: function (sheetName) {
    var requestCache = ROOMS_APP.DB_REQUEST_CACHE_;
    if (requestCache && Object.prototype.hasOwnProperty.call(requestCache.headers, sheetName)) {
      requestCache.stats.requestHits += 1;
      return requestCache.headers[sheetName].slice();
    }

    var cachedHeaders = this.getCachedHeaders_(sheetName);
    if (cachedHeaders) {
      if (requestCache) {
        requestCache.headers[sheetName] = cachedHeaders.slice();
        requestCache.stats.scriptHits += 1;
      }
      return cachedHeaders.slice();
    }

    var sheet = this.getSheet(sheetName);
    if (!sheet || sheet.getLastRow() < 1) {
      return [];
    }

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (requestCache) {
      requestCache.headers[sheetName] = headers.slice();
      requestCache.stats.misses += 1;
    }
    this.putCachedHeaders_(sheetName, headers);
    return headers.slice();
  },

  readRows: function (sheetName) {
    var requestCache = ROOMS_APP.DB_REQUEST_CACHE_;
    if (requestCache && Object.prototype.hasOwnProperty.call(requestCache.rows, sheetName)) {
      requestCache.stats.requestHits += 1;
      return this.cloneRows_(requestCache.rows[sheetName]);
    }

    var cachedRows = this.getCachedRows_(sheetName);
    if (cachedRows) {
      if (requestCache) {
        requestCache.rows[sheetName] = this.cloneRows_(cachedRows);
        requestCache.stats.scriptHits += 1;
      }
      return this.cloneRows_(cachedRows);
    }

    var sheet = this.getSheet(sheetName);
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }

    var headers = this.getHeaders(sheetName);
    var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
    var timezone = ROOMS_APP.DEFAULT_TIMEZONE;

    var rows = values.map(function (row) {
      var entry = {};
      headers.forEach(function (header, index) {
        entry[header] = ROOMS_APP.DB.normalizeCellValue_(header, row[index], timezone);
      });
      return entry;
    });

    if (requestCache) {
      requestCache.rows[sheetName] = this.cloneRows_(rows);
      requestCache.stats.misses += 1;
    }
    this.putCachedRows_(sheetName, rows);
    return rows;
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
    this.invalidateSheetCache_(sheetName);
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
    this.invalidateSheetCache_(sheetName);
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
    this.invalidateSheetCache_(sheetName);
  },

  cloneRows_: function (rows) {
    return (rows || []).map(function (row) {
      var clone = {};
      Object.keys(row || {}).forEach(function (key) {
        clone[key] = row[key];
      });
      return clone;
    });
  },

  isStaticSheetCacheable_: function (sheetName) {
    return Object.prototype.hasOwnProperty.call(this.STATIC_SHEET_CACHE_TTLS_, sheetName);
  },

  getCacheService_: function () {
    return CacheService.getScriptCache();
  },

  getCacheKey_: function (sheetName, suffix) {
    return 'rooms:db:' + sheetName + ':' + suffix;
  },

  getCachedHeaders_: function (sheetName) {
    if (!this.isStaticSheetCacheable_(sheetName)) {
      return null;
    }
    var cached = this.getCacheService_().get(this.getCacheKey_(sheetName, 'headers'));
    if (!cached) {
      return null;
    }
    return ROOMS_APP.parseJson(cached, null);
  },

  putCachedHeaders_: function (sheetName, headers) {
    if (!this.isStaticSheetCacheable_(sheetName)) {
      return;
    }
    this.getCacheService_().put(
      this.getCacheKey_(sheetName, 'headers'),
      JSON.stringify(headers || []),
      this.STATIC_SHEET_CACHE_TTLS_[sheetName]
    );
  },

  getCachedRows_: function (sheetName) {
    if (!this.isStaticSheetCacheable_(sheetName)) {
      return null;
    }
    var cached = this.getCacheService_().get(this.getCacheKey_(sheetName, 'rows'));
    if (!cached) {
      return null;
    }
    return ROOMS_APP.parseJson(cached, null);
  },

  putCachedRows_: function (sheetName, rows) {
    if (!this.isStaticSheetCacheable_(sheetName)) {
      return;
    }
    this.getCacheService_().put(
      this.getCacheKey_(sheetName, 'rows'),
      JSON.stringify(rows || []),
      this.STATIC_SHEET_CACHE_TTLS_[sheetName]
    );
  },

  invalidateSheetCache_: function (sheetName) {
    var requestCache = ROOMS_APP.DB_REQUEST_CACHE_;
    if (requestCache) {
      delete requestCache.headers[sheetName];
      delete requestCache.rows[sheetName];
    }
    if (!this.isStaticSheetCacheable_(sheetName)) {
      return;
    }
    this.getCacheService_().removeAll([
      this.getCacheKey_(sheetName, 'headers'),
      this.getCacheKey_(sheetName, 'rows')
    ]);
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
