var ROOMS_APP = ROOMS_APP || {};

ROOMS_APP.Timetable = {
  CONFIG_DOCENTI_SHEET_KEY_: 'TIMETABLE_DOCENTI_SHEET',
  CONFIG_LABORATORI_SHEET_KEY_: 'TIMETABLE_LABORATORI_SHEET',
  DEFAULT_DOCENTI_SHEET_: 'ORARIO_DOCENTI',
  DEFAULT_LABORATORI_SHEET_: 'ORARIO_LABORATORI',
  HEADER_SCAN_ROWS_: 12,

  PERIOD_TIME_MAP_: {
    '1': { startTime: '08:00', endTime: '09:00' },
    '2': { startTime: '09:00', endTime: '10:00' },
    '3': { startTime: '10:00', endTime: '11:00' },
    '4': { startTime: '11:00', endTime: '12:00' },
    '5': { startTime: '12:00', endTime: '13:00' },
    '6': { startTime: '13:00', endTime: '14:00' },
    '7': { startTime: '14:00', endTime: '15:00' },
    '8': { startTime: '15:00', endTime: '16:00' }
  },

  WEEKDAY_ALIASES_: {
    MONDAY: 'Monday',
    MON: 'Monday',
    LUNEDI: 'Monday',
    LUNEDÌ: 'Monday',
    LUN: 'Monday',
    TUESDAY: 'Tuesday',
    TUE: 'Tuesday',
    TUES: 'Tuesday',
    MARTEDI: 'Tuesday',
    MARTEDÌ: 'Tuesday',
    MAR: 'Tuesday',
    WEDNESDAY: 'Wednesday',
    WED: 'Wednesday',
    MERCOLEDI: 'Wednesday',
    MERCOLEDÌ: 'Wednesday',
    MER: 'Wednesday',
    THURSDAY: 'Thursday',
    THU: 'Thursday',
    THURS: 'Thursday',
    GIOVEDI: 'Thursday',
    GIOVEDÌ: 'Thursday',
    GIO: 'Thursday',
    FRIDAY: 'Friday',
    FRI: 'Friday',
    VENERDI: 'Friday',
    VENERDÌ: 'Friday',
    VEN: 'Friday',
    SATURDAY: 'Saturday',
    SAT: 'Saturday',
    SABATO: 'Saturday',
    SAB: 'Saturday',
    SUNDAY: 'Sunday',
    SUN: 'Sunday',
    DOMENICA: 'Sunday',
    DOM: 'Sunday'
  },

  WEEKDAY_ORDER_: {
    Monday: 1,
    Tuesday: 2,
    Wednesday: 3,
    Thursday: 4,
    Friday: 5,
    Saturday: 6,
    Sunday: 7
  },

  getPeriodTimeMap: function () {
    return this.PERIOD_TIME_MAP_;
  },

  getPeriodRange: function (period) {
    var key = ROOMS_APP.normalizeString(period);
    return this.PERIOD_TIME_MAP_[key] || null;
  },

  rebuildTimetableOccupancy: function () {
    return this.rebuildTimetableOccupancyFromSheets();
  },

  rebuildTimetableOccupancyFromSheets: function () {
    ROOMS_APP.Schema.ensureAll();
    var currentById = this.readCurrentRowsById_();
    var nowIso = ROOMS_APP.toIsoDateTime(new Date());
    var rows = [];

    rows = rows.concat(this.parseDocentiSheet_(currentById, nowIso));
    rows = rows.concat(this.parseLaboratoriSheet_(currentById, nowIso));

    return this.replaceTimetableRows_(this.mergeUniqueRows_(rows));
  },

  importTimetableClassrooms: function () {
    ROOMS_APP.Schema.ensureAll();
    var currentById = this.readCurrentRowsById_();
    var nowIso = ROOMS_APP.toIsoDateTime(new Date());
    var parsed = this.parseDocentiSheet_(currentById, nowIso);
    var keepSpaces = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.TIMETABLE_OCCUPANCY).filter(function (row) {
      return ROOMS_APP.normalizeString(row.SourceType) === 'TIMETABLE_SPACE';
    });
    return this.replaceTimetableRows_(this.mergeUniqueRows_(keepSpaces.concat(parsed)));
  },

  importTimetableSpaces: function () {
    ROOMS_APP.Schema.ensureAll();
    var currentById = this.readCurrentRowsById_();
    var nowIso = ROOMS_APP.toIsoDateTime(new Date());
    var parsed = this.parseLaboratoriSheet_(currentById, nowIso);
    var keepClassrooms = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.TIMETABLE_OCCUPANCY).filter(function (row) {
      return ROOMS_APP.normalizeString(row.SourceType) === 'TIMETABLE_CLASSROOM';
    });
    return this.replaceTimetableRows_(this.mergeUniqueRows_(keepClassrooms.concat(parsed)));
  },

  parseDocentiSheet_: function (currentById, nowIso) {
    var sheetName = this.getConfiguredSourceSheetName_(
      this.CONFIG_DOCENTI_SHEET_KEY_,
      this.DEFAULT_DOCENTI_SHEET_
    );
    return this.parseMatrixSheet_(
      sheetName,
      'TIMETABLE_CLASSROOM',
      'classroom',
      currentById,
      nowIso,
      this.CONFIG_DOCENTI_SHEET_KEY_
    );
  },

  parseLaboratoriSheet_: function (currentById, nowIso) {
    var sheetName = this.getConfiguredSourceSheetName_(
      this.CONFIG_LABORATORI_SHEET_KEY_,
      this.DEFAULT_LABORATORI_SHEET_
    );
    return this.parseMatrixSheet_(
      sheetName,
      'TIMETABLE_SPACE',
      'space',
      currentById,
      nowIso,
      this.CONFIG_LABORATORI_SHEET_KEY_
    );
  },

  parseMatrixSheet_: function (sheetName, sourceType, sourceKind, currentById, nowIso, configKey) {
    var sheet = ROOMS_APP.DB.getSheet(sheetName);
    if (!sheet) {
      var keyInfo = configKey ? (' (CONFIG key: ' + configKey + ')') : '';
      throw new Error('Foglio sorgente orario non trovato: "' + sheetName + '"' + keyInfo + '.');
    }

    var values = this.readSheetDisplayValues_(sheet);
    if (!values.length) {
      return [];
    }

    var columnMeta = this.detectColumnMeta_(values);
    if (!columnMeta.usableColumns.length) {
      return [];
    }

    var rows = [];
    var rowIndex;
    for (rowIndex = columnMeta.dataStartRow; rowIndex < values.length; rowIndex += 1) {
      var rowLabel = ROOMS_APP.normalizeString(values[rowIndex][0]);
      if (!this.isRowLabelData_(rowLabel, sourceKind)) {
        continue;
      }

      var teacherName = sourceKind === 'classroom' ? rowLabel : '';
      var colIndex;
      for (colIndex = 0; colIndex < columnMeta.usableColumns.length; colIndex += 1) {
        var column = columnMeta.usableColumns[colIndex];
        var meta = columnMeta.columns[column] || {};
        var classCode = this.extractClassCode_(values[rowIndex][column]);
        if (!classCode) {
          continue;
        }

        var resourceId;
        var resourceLabel;
        if (sourceKind === 'classroom') {
          resourceId = classCode;
          resourceLabel = classCode;
        } else {
          resourceId = rowLabel;
          resourceLabel = rowLabel;
        }
        if (!resourceId) {
          continue;
        }

        var occupancyId = this.buildOccupancyId_(sourceType, resourceId, meta.weekday, meta.period, classCode, teacherName);
        var existing = currentById[occupancyId] || {};
        rows.push({
          OccupancyId: occupancyId,
          SourceType: sourceType,
          ResourceId: resourceId,
          ResourceLabel: resourceLabel,
          Weekday: meta.weekday,
          Period: meta.period,
          StartTime: meta.startTime,
          EndTime: meta.endTime,
          ClassCode: classCode,
          TeacherName: teacherName,
          DisplayLabel: classCode,
          IsActive: 'TRUE',
          Notes: '',
          CreatedAtISO: existing.CreatedAtISO || nowIso,
          UpdatedAtISO: nowIso
        });
      }
    }

    return rows;
  },

  detectColumnMeta_: function (values) {
    var rowCount = values.length;
    var columnCount = this.getMaxColumnCount_(values);
    var scanRows = Math.min(this.HEADER_SCAN_ROWS_, rowCount);
    var dayHints = [];
    var periodHints = [];
    var headerBottomRow = -1;
    var rowIndex;
    var colIndex;

    for (rowIndex = 0; rowIndex < scanRows; rowIndex += 1) {
      var dayCount = 0;
      var periodCount = 0;
      for (colIndex = 1; colIndex < columnCount; colIndex += 1) {
        var token = ROOMS_APP.normalizeString(values[rowIndex][colIndex]);
        if (!token) {
          continue;
        }
        var weekday = this.normalizeWeekday_(token);
        if (weekday) {
          dayHints[colIndex] = weekday;
          dayCount += 1;
        }
        var period = this.extractPeriodFromHeaderCell_(token);
        if (period) {
          periodHints[colIndex] = period;
          periodCount += 1;
        }
      }
      if (dayCount > 0 || periodCount >= 3) {
        headerBottomRow = rowIndex;
      }
    }

    var dayByColumn = [];
    var periodByColumn = [];
    var activeWeekday = '';
    var activePeriod = 1;
    for (colIndex = 1; colIndex < columnCount; colIndex += 1) {
      if (dayHints[colIndex]) {
        activeWeekday = dayHints[colIndex];
        activePeriod = 1;
      }
      dayByColumn[colIndex] = activeWeekday;

      if (periodHints[colIndex]) {
        periodByColumn[colIndex] = periodHints[colIndex];
        activePeriod = Number(periodHints[colIndex]) + 1;
      } else if (activeWeekday && activePeriod >= 1 && activePeriod <= 8) {
        periodByColumn[colIndex] = String(activePeriod);
        activePeriod += 1;
      } else {
        periodByColumn[colIndex] = '';
      }
    }

    var columns = [];
    var usableColumns = [];
    for (colIndex = 1; colIndex < columnCount; colIndex += 1) {
      var weekdayValue = dayByColumn[colIndex];
      var periodValue = periodByColumn[colIndex];
      var range = this.getPeriodRange(periodValue);
      columns[colIndex] = {
        weekday: weekdayValue,
        period: periodValue,
        startTime: range ? range.startTime : '',
        endTime: range ? range.endTime : ''
      };
      if (weekdayValue && periodValue && range) {
        usableColumns.push(colIndex);
      }
    }

    return {
      columns: columns,
      usableColumns: usableColumns,
      dataStartRow: Math.max(1, headerBottomRow + 1)
    };
  },

  listOccupanciesForDate: function (resourceId, dateString) {
    ROOMS_APP.Schema.ensureAll();
    var targetDate = ROOMS_APP.toIsoDate(dateString || new Date());
    var targetWeekday = ROOMS_APP.getWeekdayName(targetDate);
    var targetResource = ROOMS_APP.normalizeString(resourceId);
    var hasResourceFilter = Boolean(targetResource);

    return this.listActiveRows_()
      .filter(function (row) {
        var rowWeekday = ROOMS_APP.Timetable.normalizeWeekday_(row.Weekday);
        if (rowWeekday !== targetWeekday) {
          return false;
        }
        if (hasResourceFilter && !ROOMS_APP.Timetable.matchesResourceId_(row.ResourceId, targetResource)) {
          return false;
        }
        return Boolean(row.StartTime && row.EndTime);
      })
      .map(function (row) {
        return ROOMS_APP.Timetable.buildOccurrenceView_(row, targetDate, hasResourceFilter ? targetResource : '');
      });
  },

  listUpcomingOccupanciesForRoom: function (resourceId, fromDate, maxDaysAhead) {
    ROOMS_APP.Schema.ensureAll();
    var startDate = ROOMS_APP.toIsoDate(fromDate || new Date());
    var daysAhead = Math.max(0, Number(maxDaysAhead || ROOMS_APP.getNumberConfig('MAX_DAYS_AHEAD', 30)));
    var upcoming = [];
    var offset;

    for (offset = 0; offset <= daysAhead; offset += 1) {
      var cursor = ROOMS_APP.combineDateTime(startDate, '00:00');
      cursor.setDate(cursor.getDate() + offset);
      var targetDate = ROOMS_APP.toIsoDate(cursor);
      upcoming = upcoming.concat(this.listOccupanciesForDate(resourceId, targetDate));
    }

    return ROOMS_APP.sortBy(upcoming, ['BookingDate', 'StartTime', 'EndTime', 'BookingId']);
  },

  getDisplayLabel: function (occupancy) {
    if (!occupancy) {
      return 'N/D';
    }
    return ROOMS_APP.normalizeString(occupancy.DisplayLabel || occupancy.ClassCode || occupancy.BookerSurname || occupancy.Title || 'N/D');
  },

  isValidClassCode_: function (value) {
    var raw = ROOMS_APP.normalizeString(value).toUpperCase();
    if (!raw) {
      return false;
    }
    if (/[\/,+;|]/.test(raw)) {
      return false;
    }
    var code = raw.replace(/[\s._-]+/g, '');
    if (!code || code === 'P' || code === 'D') {
      return false;
    }
    return /^[1-6][A-Z]{1,4}$/.test(code);
  },

  extractClassCode_: function (value) {
    var raw = ROOMS_APP.normalizeString(value).toUpperCase();
    if (!raw) {
      return '';
    }
    if (this.isValidClassCode_(raw)) {
      return raw.replace(/[\s._-]+/g, '');
    }
    return '';
  },

  listActiveRows_: function () {
    return ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.TIMETABLE_OCCUPANCY)
      .filter(function (row) {
        return ROOMS_APP.asBoolean(row.IsActive);
      });
  },

  buildOccurrenceView_: function (row, dateString, normalizedResourceId) {
    var targetDate = ROOMS_APP.toIsoDate(dateString || new Date());
    var startTime = ROOMS_APP.toTimeString(row.StartTime);
    var endTime = ROOMS_APP.toTimeString(row.EndTime);
    var bookingId = 'TT_' + ROOMS_APP.normalizeString(row.OccupancyId) + '_' + targetDate;
    var displayLabel = this.getDisplayLabel({
      DisplayLabel: row.DisplayLabel,
      ClassCode: row.ClassCode
    });

    return {
      BookingId: bookingId,
      SeriesId: 'TT_' + ROOMS_APP.normalizeString(row.OccupancyId),
      OccupancyId: ROOMS_APP.normalizeString(row.OccupancyId),
      ResourceId: ROOMS_APP.normalizeString(normalizedResourceId || row.ResourceId),
      BookingDate: targetDate,
      StartTime: startTime,
      EndTime: endTime,
      StartISO: ROOMS_APP.toIsoDateTime(ROOMS_APP.combineDateTime(targetDate, startTime)),
      EndISO: ROOMS_APP.toIsoDateTime(ROOMS_APP.combineDateTime(targetDate, endTime)),
      Title: ROOMS_APP.normalizeString(row.ResourceLabel || row.ResourceId),
      BookerEmail: '',
      BookerName: ROOMS_APP.normalizeString(row.TeacherName),
      BookerSurname: ROOMS_APP.normalizeString(row.ClassCode),
      ClassCode: ROOMS_APP.normalizeString(row.ClassCode),
      DisplayLabel: displayLabel,
      Status: 'CONFIRMED',
      SourceKind: 'TIMETABLE',
      SourceType: ROOMS_APP.normalizeString(row.SourceType || 'TIMETABLE_SPACE'),
      CanManage: false,
      IsReadOnly: true,
      Notes: ROOMS_APP.normalizeString(row.Notes),
      CreatedAtISO: ROOMS_APP.normalizeString(row.CreatedAtISO),
      UpdatedAtISO: ROOMS_APP.normalizeString(row.UpdatedAtISO),
      CancelledAtISO: ''
    };
  },

  normalizeWeekday_: function (weekday) {
    var token = ROOMS_APP.normalizeString(weekday).toUpperCase();
    if (!token) {
      return '';
    }

    if (/^\d+$/.test(token)) {
      var asNumber = Number(token);
      var numericMap = {
        1: 'Monday',
        2: 'Tuesday',
        3: 'Wednesday',
        4: 'Thursday',
        5: 'Friday',
        6: 'Saturday',
        7: 'Sunday'
      };
      return numericMap[asNumber] || '';
    }

    return this.WEEKDAY_ALIASES_[token] || '';
  },

  buildOccupancyId_: function (sourceType, resourceId, weekday, period, classCode, teacherName) {
    var key = [
      ROOMS_APP.normalizeString(sourceType),
      ROOMS_APP.normalizeString(resourceId),
      ROOMS_APP.normalizeString(weekday),
      ROOMS_APP.normalizeString(period),
      ROOMS_APP.normalizeString(classCode),
      ROOMS_APP.normalizeString(teacherName)
    ].join('|');
    return 'TT_OCC_' + ROOMS_APP.slugify(key);
  },

  extractPeriodFromHeaderCell_: function (value) {
    var token = ROOMS_APP.normalizeString(value).toUpperCase();
    if (!token || this.isValidClassCode_(token)) {
      return '';
    }
    if (/^[1-8]$/.test(token)) {
      return token;
    }
    var digits = token.replace(/[^0-9]/g, '');
    if (/^[1-8]$/.test(digits)) {
      return digits;
    }
    return '';
  },

  isRowLabelData_: function (label, sourceKind) {
    var token = ROOMS_APP.normalizeString(label);
    if (!token) {
      return false;
    }
    var upper = token.toUpperCase();
    if (this.normalizeWeekday_(upper)) {
      return false;
    }
    if (/^[1-8]$/.test(upper)) {
      return false;
    }
    if (/^(DOCENTE|DOCENTI|ORARIO|GIORNO|DAY|PERIODO|PERIOD|ORE)$/.test(upper)) {
      return false;
    }
    if (sourceKind === 'space' && upper === 'LABORATORI') {
      return false;
    }
    return true;
  },

  getMaxColumnCount_: function (values) {
    var max = 0;
    (values || []).forEach(function (row) {
      max = Math.max(max, (row || []).length);
    });
    return max;
  },

  matchesResourceId_: function (rowResourceId, expectedResourceId) {
    var left = ROOMS_APP.normalizeString(rowResourceId);
    var right = ROOMS_APP.normalizeString(expectedResourceId);
    if (!left || !right) {
      return false;
    }
    if (left.toUpperCase() === right.toUpperCase()) {
      return true;
    }
    return ROOMS_APP.slugify(left) === ROOMS_APP.slugify(right);
  },

  mergeUniqueRows_: function (rows) {
    var byId = {};
    (rows || []).forEach(function (row) {
      var occupancyId = ROOMS_APP.normalizeString(row.OccupancyId);
      if (!occupancyId) {
        return;
      }
      byId[occupancyId] = row;
    });

    var merged = Object.keys(byId).map(function (occupancyId) {
      return byId[occupancyId];
    });

    return merged.sort(function (left, right) {
      var weekdayDelta = ROOMS_APP.Timetable.getWeekdayOrder_(left.Weekday) - ROOMS_APP.Timetable.getWeekdayOrder_(right.Weekday);
      if (weekdayDelta !== 0) {
        return weekdayDelta;
      }

      var resourceDelta = ROOMS_APP.normalizeString(left.ResourceId).localeCompare(ROOMS_APP.normalizeString(right.ResourceId));
      if (resourceDelta !== 0) {
        return resourceDelta;
      }

      var periodDelta = Number(left.Period || 0) - Number(right.Period || 0);
      if (periodDelta !== 0) {
        return periodDelta;
      }

      var classDelta = ROOMS_APP.normalizeString(left.ClassCode).localeCompare(ROOMS_APP.normalizeString(right.ClassCode));
      if (classDelta !== 0) {
        return classDelta;
      }

      return ROOMS_APP.normalizeString(left.TeacherName).localeCompare(ROOMS_APP.normalizeString(right.TeacherName));
    });
  },

  readCurrentRowsById_: function () {
    var map = {};
    ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.TIMETABLE_OCCUPANCY).forEach(function (row) {
      var occupancyId = ROOMS_APP.normalizeString(row.OccupancyId);
      if (occupancyId) {
        map[occupancyId] = row;
      }
    });
    return map;
  },

  replaceTimetableRows_: function (rows) {
    var headers = ROOMS_APP.DB.getHeaders(ROOMS_APP.SHEET_NAMES.TIMETABLE_OCCUPANCY);
    ROOMS_APP.DB.replaceRows(ROOMS_APP.SHEET_NAMES.TIMETABLE_OCCUPANCY, headers, rows || []);
    return {
      ok: true,
      rowCount: (rows || []).length
    };
  },

  readSheetDisplayValues_: function (sheet) {
    var rowCount = sheet.getLastRow();
    var columnCount = sheet.getLastColumn();
    if (rowCount < 1 || columnCount < 1) {
      return [];
    }
    return sheet.getRange(1, 1, rowCount, columnCount).getDisplayValues();
  },

  getWeekdayOrder_: function (weekday) {
    return this.WEEKDAY_ORDER_[ROOMS_APP.normalizeString(weekday)] || 99;
  },

  getConfiguredSourceSheetName_: function (configKey, fallbackName) {
    var configured = ROOMS_APP.normalizeString(ROOMS_APP.getConfigValue(configKey, fallbackName));
    return configured || fallbackName;
  }
};
