var ROOMS_APP = ROOMS_APP || {};

ROOMS_APP.Timetable = {
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
    LUNEDI: 'Monday',
    LUNEDÌ: 'Monday',
    TUESDAY: 'Tuesday',
    MARTEDI: 'Tuesday',
    MARTEDÌ: 'Tuesday',
    WEDNESDAY: 'Wednesday',
    MERCOLEDI: 'Wednesday',
    MERCOLEDÌ: 'Wednesday',
    THURSDAY: 'Thursday',
    GIOVEDI: 'Thursday',
    GIOVEDÌ: 'Thursday',
    FRIDAY: 'Friday',
    VENERDI: 'Friday',
    VENERDÌ: 'Friday',
    SATURDAY: 'Saturday',
    SABATO: 'Saturday',
    SUNDAY: 'Sunday',
    DOMENICA: 'Sunday'
  },

  getPeriodTimeMap: function () {
    return this.PERIOD_TIME_MAP_;
  },

  getPeriodRange: function (period) {
    var key = ROOMS_APP.normalizeString(period);
    return this.PERIOD_TIME_MAP_[key] || null;
  },

  rebuildTimetableOccupancy: function () {
    ROOMS_APP.Schema.ensureAll();
    var normalizedRows = this.buildNormalizedRows_({
      includeClassrooms: true,
      includeSpaces: true
    });
    return this.replaceTimetableRows_(normalizedRows);
  },

  importTimetableClassrooms: function () {
    ROOMS_APP.Schema.ensureAll();
    var current = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.TIMETABLE_OCCUPANCY);
    var keepSpaces = current.filter(function (row) {
      return ROOMS_APP.normalizeString(row.SourceType) === 'TIMETABLE_SPACE';
    });
    var normalizedRows = this.buildNormalizedRows_({
      includeClassrooms: true,
      includeSpaces: false
    });
    return this.replaceTimetableRows_(keepSpaces.concat(normalizedRows));
  },

  importTimetableSpaces: function () {
    ROOMS_APP.Schema.ensureAll();
    var current = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.TIMETABLE_OCCUPANCY);
    var keepClassrooms = current.filter(function (row) {
      return ROOMS_APP.normalizeString(row.SourceType) === 'TIMETABLE_CLASSROOM';
    });
    var normalizedRows = this.buildNormalizedRows_({
      includeClassrooms: false,
      includeSpaces: true
    });
    return this.replaceTimetableRows_(keepClassrooms.concat(normalizedRows));
  },

  listOccupanciesForDate: function (resourceId, dateString) {
    ROOMS_APP.Schema.ensureAll();
    var targetDate = ROOMS_APP.toIsoDate(dateString || new Date());
    var targetWeekday = ROOMS_APP.getWeekdayName(targetDate);
    var targetResource = ROOMS_APP.normalizeString(resourceId).toUpperCase();

    return this.listActiveRows_()
      .filter(function (row) {
        var rowWeekday = ROOMS_APP.Timetable.normalizeWeekday_(row.Weekday);
        var rowResource = ROOMS_APP.normalizeString(row.ResourceId).toUpperCase();
        if (rowWeekday !== targetWeekday) {
          return false;
        }
        if (targetResource && rowResource !== targetResource) {
          return false;
        }
        return Boolean(row.StartTime && row.EndTime);
      })
      .map(function (row) {
        return ROOMS_APP.Timetable.buildOccurrenceView_(row, targetDate);
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

  listActiveRows_: function () {
    return ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.TIMETABLE_OCCUPANCY)
      .filter(function (row) {
        return ROOMS_APP.asBoolean(row.IsActive);
      });
  },

  buildNormalizedRows_: function (options) {
    var settings = options || {};
    var includeClassrooms = settings.includeClassrooms !== false;
    var includeSpaces = settings.includeSpaces !== false;
    var existingById = {};
    var nowIso = ROOMS_APP.toIsoDateTime(new Date());
    var rows = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.TIMETABLE_OCCUPANCY);

    rows.forEach(function (row) {
      var occupancyId = ROOMS_APP.normalizeString(row.OccupancyId);
      if (occupancyId) {
        existingById[occupancyId] = row;
      }
    });

    var normalizedRows = [];
    if (includeClassrooms) {
      normalizedRows = normalizedRows.concat(this.buildClassroomRowsFromRaw_(existingById, nowIso));
    }
    if (includeSpaces) {
      normalizedRows = normalizedRows.concat(this.buildSpaceRowsFromRaw_(existingById, nowIso));
    }

    return ROOMS_APP.sortBy(normalizedRows, ['ResourceId', 'Weekday', 'StartTime', 'Period', 'ClassCode']);
  },

  buildClassroomRowsFromRaw_: function (existingById, nowIso) {
    var rows = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.TIMETABLE_DOCENTI_RAW);
    var normalizedRows = [];

    rows.forEach(function (row) {
      var classCode = ROOMS_APP.normalizeString(row.ClassCode || row.Class || row.Classe);
      if (!classCode) {
        return;
      }

      var weekday = ROOMS_APP.Timetable.normalizeWeekday_(row.Weekday || row.Day || row.Giorno);
      var period = ROOMS_APP.Timetable.normalizePeriod_(row.Period || row.Ora);
      var range = ROOMS_APP.Timetable.getPeriodRange(period);
      if (!weekday || !range) {
        return;
      }

      var teacherName = ROOMS_APP.normalizeString(row.TeacherName || row.Teacher || row.Docente);
      var resourceId = ROOMS_APP.slugify(classCode);
      var occupancyId = ROOMS_APP.Timetable.buildOccupancyId_('TIMETABLE_CLASSROOM', resourceId, weekday, period, classCode, teacherName);
      var existing = existingById[occupancyId] || {};

      normalizedRows.push({
        OccupancyId: occupancyId,
        SourceType: 'TIMETABLE_CLASSROOM',
        ResourceId: resourceId,
        ResourceLabel: classCode,
        Weekday: weekday,
        Period: period,
        StartTime: range.startTime,
        EndTime: range.endTime,
        ClassCode: classCode,
        TeacherName: teacherName,
        DisplayLabel: classCode,
        IsActive: ROOMS_APP.Timetable.normalizeActiveFlag_(row.IsActive),
        Notes: ROOMS_APP.normalizeString(row.Notes),
        CreatedAtISO: existing.CreatedAtISO || nowIso,
        UpdatedAtISO: nowIso
      });
    });

    return normalizedRows;
  },

  buildSpaceRowsFromRaw_: function (existingById, nowIso) {
    var rows = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.TIMETABLE_SPACES_RAW);
    var normalizedRows = [];

    rows.forEach(function (row) {
      var resourceIdRaw = ROOMS_APP.normalizeString(row.ResourceId || row.RoomCode || row.SpaceCode);
      var resourceLabel = ROOMS_APP.normalizeString(row.ResourceLabel || row.SpaceLabel || resourceIdRaw);
      var resourceId = ROOMS_APP.slugify(resourceIdRaw || resourceLabel);
      if (!resourceId) {
        return;
      }

      var classCode = ROOMS_APP.normalizeString(row.ClassCode || row.Class || row.Classe);
      var teacherName = ROOMS_APP.normalizeString(row.TeacherName || row.Teacher || row.Docente);
      var weekday = ROOMS_APP.Timetable.normalizeWeekday_(row.Weekday || row.Day || row.Giorno);
      var period = ROOMS_APP.Timetable.normalizePeriod_(row.Period || row.Ora);
      var range = ROOMS_APP.Timetable.getPeriodRange(period);
      if (!weekday || !range) {
        return;
      }

      var occupancyId = ROOMS_APP.Timetable.buildOccupancyId_('TIMETABLE_SPACE', resourceId, weekday, period, classCode, teacherName);
      var existing = existingById[occupancyId] || {};

      normalizedRows.push({
        OccupancyId: occupancyId,
        SourceType: 'TIMETABLE_SPACE',
        ResourceId: resourceId,
        ResourceLabel: resourceLabel || resourceId,
        Weekday: weekday,
        Period: period,
        StartTime: range.startTime,
        EndTime: range.endTime,
        ClassCode: classCode,
        TeacherName: teacherName,
        DisplayLabel: classCode || resourceLabel || resourceId,
        IsActive: ROOMS_APP.Timetable.normalizeActiveFlag_(row.IsActive),
        Notes: ROOMS_APP.normalizeString(row.Notes),
        CreatedAtISO: existing.CreatedAtISO || nowIso,
        UpdatedAtISO: nowIso
      });
    });

    return normalizedRows;
  },

  replaceTimetableRows_: function (rows) {
    var headers = ROOMS_APP.DB.getHeaders(ROOMS_APP.SHEET_NAMES.TIMETABLE_OCCUPANCY);
    ROOMS_APP.DB.replaceRows(ROOMS_APP.SHEET_NAMES.TIMETABLE_OCCUPANCY, headers, rows);

    return {
      ok: true,
      rowCount: rows.length
    };
  },

  buildOccurrenceView_: function (row, dateString) {
    var targetDate = ROOMS_APP.toIsoDate(dateString || new Date());
    var bookingId = 'TT_' + ROOMS_APP.normalizeString(row.OccupancyId) + '_' + targetDate;
    var displayLabel = this.getDisplayLabel({
      DisplayLabel: row.DisplayLabel,
      ClassCode: row.ClassCode
    });

    return {
      BookingId: bookingId,
      SeriesId: 'TT_' + ROOMS_APP.normalizeString(row.OccupancyId),
      OccupancyId: ROOMS_APP.normalizeString(row.OccupancyId),
      ResourceId: ROOMS_APP.normalizeString(row.ResourceId),
      BookingDate: targetDate,
      StartTime: ROOMS_APP.toTimeString(row.StartTime),
      EndTime: ROOMS_APP.toTimeString(row.EndTime),
      StartISO: ROOMS_APP.toIsoDateTime(ROOMS_APP.combineDateTime(targetDate, ROOMS_APP.toTimeString(row.StartTime))),
      EndISO: ROOMS_APP.toIsoDateTime(ROOMS_APP.combineDateTime(targetDate, ROOMS_APP.toTimeString(row.EndTime))),
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

  normalizePeriod_: function (period) {
    return ROOMS_APP.normalizeString(period);
  },

  normalizeActiveFlag_: function (value) {
    if (value === '' || value == null) {
      return 'TRUE';
    }
    return ROOMS_APP.asBoolean(value) ? 'TRUE' : 'FALSE';
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
  }
};
