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
    var classroomCodeSet = this.getClassroomResourceCodeSet_();
    var rows = [];
    var docentiRows = this.parseDocentiSheet_(currentById, nowIso, classroomCodeSet);
    var laboratoriRows = this.parseLaboratoriSheet_(currentById, nowIso);

    rows = rows.concat(docentiRows);
    rows = rows.concat(laboratoriRows);

    var result = this.replaceTimetableRows_(this.mergeUniqueRows_(rows));
    result.docentiCount = docentiRows.length;
    result.laboratoriCount = laboratoriRows.length;
    result.sample = rows.slice(0, 5);
    Logger.log('Timetable rebuild completed. Docenti: %s, Laboratori: %s, Total: %s', result.docentiCount, result.laboratoriCount, result.rowCount);

    return result;
  },

  importTimetableClassrooms: function () {
    ROOMS_APP.Schema.ensureAll();
    var currentById = this.readCurrentRowsById_();
    var nowIso = ROOMS_APP.toIsoDateTime(new Date());
    var classroomCodeSet = this.getClassroomResourceCodeSet_();
    var parsed = this.parseDocentiSheet_(currentById, nowIso, classroomCodeSet);
    var keepSpaces = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.TIMETABLE_OCCUPANCY).filter(function (row) {
      return ROOMS_APP.normalizeString(row.SourceType) === 'TIMETABLE_SPACE';
    });
    var result = this.replaceTimetableRows_(this.mergeUniqueRows_(keepSpaces.concat(parsed)));
    result.docentiCount = parsed.length;
    result.laboratoriCount = keepSpaces.length;
    return result;
  },

  importTimetableSpaces: function () {
    ROOMS_APP.Schema.ensureAll();
    var currentById = this.readCurrentRowsById_();
    var nowIso = ROOMS_APP.toIsoDateTime(new Date());
    var parsed = this.parseLaboratoriSheet_(currentById, nowIso);
    var keepClassrooms = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.TIMETABLE_OCCUPANCY).filter(function (row) {
      return ROOMS_APP.normalizeString(row.SourceType) === 'TIMETABLE_CLASSROOM';
    });
    var result = this.replaceTimetableRows_(this.mergeUniqueRows_(keepClassrooms.concat(parsed)));
    result.docentiCount = keepClassrooms.length;
    result.laboratoriCount = parsed.length;
    return result;
  },

  parseDocentiSheet_: function (currentById, nowIso, classroomCodeSet) {
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
      this.CONFIG_DOCENTI_SHEET_KEY_,
      {
        requireClassroomResource: true,
        classroomCodeSet: classroomCodeSet || {}
      }
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
      this.CONFIG_LABORATORI_SHEET_KEY_,
      {}
    );
  },

  parseMatrixSheet_: function (sheetName, sourceType, sourceKind, currentById, nowIso, configKey, parseOptions) {
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
        if (sourceKind === 'classroom' && parseOptions && parseOptions.requireClassroomResource && !parseOptions.classroomCodeSet[classCode]) {
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
    var layout = this.detectMatrixLayout_(values, Math.min(this.HEADER_SCAN_ROWS_, rowCount), columnCount);
    if (!layout || layout.periodRow < 0) {
      return {
        columns: [],
        usableColumns: [],
        dataStartRow: rowCount
      };
    }

    var dataStartRow = layout.periodRow + 1;
    var columns = [];
    var usableColumns = [];
    var colIndex;
    for (colIndex = 1; colIndex < columnCount; colIndex += 1) {
      var periodValue = layout.periodByColumn[colIndex] || '';
      var weekdayValue = layout.weekdayByColumn[colIndex] || '';
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
      dataStartRow: dataStartRow,
      periodRow: layout.periodRow,
      weekdayRow: layout.weekdayRow
    };
  },

  detectMatrixLayout_: function (values, scanRows, columnCount) {
    var periodRow = this.detectPeriodHeaderRow_(values, scanRows, columnCount);
    if (periodRow < 0) {
      return null;
    }

    var weekdayMeta = this.detectWeekdayAnchors_(values, periodRow, columnCount);
    var weekdayAnchors = weekdayMeta.anchors || [];
    var weekdayByColumn = [];
    var periodByColumn = [];
    var colIndex;

    for (colIndex = 1; colIndex < columnCount; colIndex += 1) {
      var period = this.extractPeriodFromHeaderCell_(values[periodRow][colIndex]);
      periodByColumn[colIndex] = period;
      weekdayByColumn[colIndex] = period ? this.resolveWeekdayFromAnchors_(colIndex, weekdayAnchors) : '';
    }

    return {
      periodRow: periodRow,
      weekdayRow: weekdayMeta.row,
      periodByColumn: periodByColumn,
      weekdayByColumn: weekdayByColumn
    };
  },

  detectPeriodHeaderRow_: function (values, scanRows, columnCount) {
    var bestRow = -1;
    var bestCount = 0;
    var bestDistinct = 0;
    var rowIndex;
    var colIndex;

    for (rowIndex = 0; rowIndex < scanRows; rowIndex += 1) {
      var count = 0;
      var distinct = {};
      for (colIndex = 1; colIndex < columnCount; colIndex += 1) {
        var period = this.extractPeriodFromHeaderCell_(values[rowIndex][colIndex]);
        if (!period) {
          continue;
        }
        count += 1;
        distinct[period] = true;
      }

      var distinctCount = Object.keys(distinct).length;
      if (count < 4 || distinctCount < 4) {
        continue;
      }

      if (
        count > bestCount ||
        (count === bestCount && distinctCount > bestDistinct) ||
        (count === bestCount && distinctCount === bestDistinct && rowIndex > bestRow)
      ) {
        bestRow = rowIndex;
        bestCount = count;
        bestDistinct = distinctCount;
      }
    }

    return bestRow;
  },

  detectWeekdayAnchors_: function (values, periodRow, columnCount) {
    var best = {
      row: -1,
      anchors: []
    };
    var rowIndex;
    var colIndex;

    for (rowIndex = 0; rowIndex < periodRow; rowIndex += 1) {
      var anchors = [];
      for (colIndex = 1; colIndex < columnCount; colIndex += 1) {
        var weekday = this.normalizeWeekday_(values[rowIndex][colIndex]);
        if (!weekday) {
          continue;
        }
        anchors.push({
          col: colIndex,
          weekday: weekday
        });
      }

      if (anchors.length > best.anchors.length || (anchors.length === best.anchors.length && rowIndex > best.row)) {
        best = {
          row: rowIndex,
          anchors: anchors
        };
      }
    }

    return best;
  },

  resolveWeekdayFromAnchors_: function (columnIndex, anchors) {
    var active = '';
    var index;
    for (index = 0; index < (anchors || []).length; index += 1) {
      if (anchors[index].col > columnIndex) {
        break;
      }
      active = anchors[index].weekday;
    }
    return active;
  },

  listOccupanciesForDate: function (resourceId, dateString) {
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
    if (!code || code === 'P' || code === 'D' || code === 'AI' || code === 'DI') {
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
    if (/[A-Z]/.test(token)) {
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

  getClassroomResourceCodeSet_: function () {
    var set = {};
    ROOMS_APP.Board.listResources_().forEach(function (resource) {
      var byId = ROOMS_APP.normalizeString(resource.ResourceId).toUpperCase();
      var byDisplay = ROOMS_APP.normalizeString(resource.DisplayName).toUpperCase();

      if (ROOMS_APP.Timetable.isValidClassCode_(byId)) {
        set[byId.replace(/[\s._-]+/g, '')] = true;
      }
      if (ROOMS_APP.Timetable.isValidClassCode_(byDisplay)) {
        set[byDisplay.replace(/[\s._-]+/g, '')] = true;
      }
    });
    return set;
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
