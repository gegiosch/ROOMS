var ROOMS_APP = ROOMS_APP || {};

ROOMS_APP.Schema = {
  ADMIN_HEADERS_: [
    'Email',
    'OrgUnitPath',
    'Role',
    'Enabled',
    'CanBook',
    'CanManageReplacement',
    'CanManageAulaMagna',
    'CanUseSimulation',
    'CanAccessAdmin',
    'Notes'
  ],

  HEADER_MIN_WIDTHS_: {
    Email: 180,
    OrgUnitPath: 220,
    Notes: 220,
    Subject: 220,
    Recipients: 260,
    ReferenceDate: 120,
    SentAtISO: 150,
    UpdatedAtISO: 150,
    CreatedAtISO: 150,
    UpdatedBy: 150,
    SentBy: 150,
    ReportType: 120,
    TeacherEmail: 180,
    TeacherName: 180,
    OriginalTeacherEmail: 180,
    OriginalTeacherName: 190,
    ReplacementTeacherEmail: 180,
    ReplacementTeacherSurname: 190,
    ReplacementTeacherName: 180,
    ReplacementTeacherDisplayName: 220,
    StartDate: 120,
    EndDate: 120,
    ReplyTo: 180
  },

  PLAIN_TEXT_SHEETS_: (function () {
    var sheets = {};
    sheets[ROOMS_APP.SHEET_NAMES.CONFIG] = true;
    sheets[ROOMS_APP.SHEET_NAMES.RESOURCES] = true;
    sheets[ROOMS_APP.SHEET_NAMES.HOLIDAYS] = true;
    sheets[ROOMS_APP.SHEET_NAMES.CLOSURES] = true;
    sheets[ROOMS_APP.SHEET_NAMES.WEEK_SCHEDULE] = true;
    sheets[ROOMS_APP.SHEET_NAMES.SPECIAL_OPENINGS] = true;
    sheets[ROOMS_APP.SHEET_NAMES.POLICY_OVERRIDES] = true;
    sheets[ROOMS_APP.SHEET_NAMES.BOOKINGS] = true;
    sheets[ROOMS_APP.SHEET_NAMES.AUDIT] = true;
    sheets[ROOMS_APP.SHEET_NAMES.ADMINS] = true;
    sheets[ROOMS_APP.SHEET_NAMES.TIMETABLE_OCCUPANCY] = true;
    sheets[ROOMS_APP.SHEET_NAMES.TIMETABLE_DOCENTI_RAW] = true;
    sheets[ROOMS_APP.SHEET_NAMES.TIMETABLE_SPACES_RAW] = true;
    sheets[ROOMS_APP.SHEET_NAMES.AULA_MAGNA_EVENTS] = true;
    sheets[ROOMS_APP.SHEET_NAMES.REPL_CLASS_OUT] = true;
    sheets[ROOMS_APP.SHEET_NAMES.REPL_DAY_TEACHERS] = true;
    sheets[ROOMS_APP.SHEET_NAMES.REPL_FIELD_TRIPS] = true;
    sheets[ROOMS_APP.SHEET_NAMES.REPL_FIELD_TRIP_TEACHERS] = true;
    sheets[ROOMS_APP.SHEET_NAMES.REPL_HOURLY_ABSENCES] = true;
    sheets[ROOMS_APP.SHEET_NAMES.REPL_ASSIGNMENTS] = true;
    sheets[ROOMS_APP.SHEET_NAMES.REPL_LONG_ASSIGNMENTS] = true;
    sheets[ROOMS_APP.SHEET_NAMES.REPORT_RECIPIENTS] = true;
    sheets[ROOMS_APP.SHEET_NAMES.REPORT_LOG] = true;
    return sheets;
  }()),

  ensureAll: function () {
    this.ensureConfig();
    this.ensureAdmins();
    this.ensureResources();
    this.ensureBookings_();
    this.ensureAudit_();
    this.ensureTimetableOccupancy();
    this.ensureTimetableDocentiRaw();
    this.ensureTimetableSpacesRaw();
    this.ensureAulaMagnaEvents();
    this.ensureReplacementClassOut();
    this.ensureReplacementDayTeachers();
    this.ensureReplacementFieldTrips();
    this.ensureReplacementFieldTripTeachers();
    this.ensureReplacementHourlyAbsences();
    this.ensureReplacementAssignments();
    this.ensureReplacementLongAssignments();
    this.ensureReportRecipients();
    this.ensureReportLog();
    this.ensureWeekSchedule();
    this.ensureHolidays();
    this.ensureClosures();
    this.ensureSpecialOpenings();
    this.ensurePolicyOverrides();
  },

  ensureConfig: function () {
    var headers = ['Key', 'Value', 'Notes'];
    this.ensureSheetStructure_(ROOMS_APP.SHEET_NAMES.CONFIG, headers);
    this.seedMissingRows_(ROOMS_APP.SHEET_NAMES.CONFIG, 'Key', ROOMS_APP.DEFAULT_CONFIG_ROWS);
  },

  ensureAdmins: function () {
    this.ensureOrderedSheetStructure_(ROOMS_APP.SHEET_NAMES.ADMINS, this.ADMIN_HEADERS_);
    this.seedAdminsFromLegacyConfig_(this.ADMIN_HEADERS_);
  },

  ensureResources: function () {
    var headers = [
      'ResourceId',
      'DisplayName',
      'AreaCode',
      'AreaLabel',
      'FloorCode',
      'FloorLabel',
      'SideCode',
      'SideLabel',
      'LayoutPage',
      'LayoutRow',
      'LayoutCol',
      'LayoutColSpan',
      'LayoutRowSpan',
      'OpenTime',
      'CloseTime',
      'IsBookable',
      'IsActive',
      'SortKey',
      'Notes'
    ];
    this.ensureSheetStructure_(ROOMS_APP.SHEET_NAMES.RESOURCES, headers);
    this.syncResourcesFromCanonical_(headers, this.buildResourceRows_());
  },

  ensureWeekSchedule: function () {
    var headers = ['Weekday', 'IsWorkingDay', 'OpenTime', 'CloseTime'];
    this.ensureSheetStructure_(ROOMS_APP.SHEET_NAMES.WEEK_SCHEDULE, headers);
    this.seedMissingRows_(ROOMS_APP.SHEET_NAMES.WEEK_SCHEDULE, 'Weekday', [
      { Weekday: 'Monday', IsWorkingDay: 'TRUE', OpenTime: '08:00', CloseTime: '18:00' },
      { Weekday: 'Tuesday', IsWorkingDay: 'TRUE', OpenTime: '08:00', CloseTime: '18:00' },
      { Weekday: 'Wednesday', IsWorkingDay: 'TRUE', OpenTime: '08:00', CloseTime: '18:00' },
      { Weekday: 'Thursday', IsWorkingDay: 'TRUE', OpenTime: '08:00', CloseTime: '18:00' },
      { Weekday: 'Friday', IsWorkingDay: 'TRUE', OpenTime: '08:00', CloseTime: '18:00' },
      { Weekday: 'Saturday', IsWorkingDay: 'FALSE', OpenTime: '', CloseTime: '' },
      { Weekday: 'Sunday', IsWorkingDay: 'FALSE', OpenTime: '', CloseTime: '' }
    ]);
  },

  ensureHolidays: function () {
    var headers = ['HolidayDate', 'Label', 'IsBlocked'];
    this.ensureSheetStructure_(ROOMS_APP.SHEET_NAMES.HOLIDAYS, headers);
    this.seedMissingRows_(ROOMS_APP.SHEET_NAMES.HOLIDAYS, 'HolidayDate', [
      { HolidayDate: '2026-01-01', Label: 'Capodanno', IsBlocked: 'TRUE' },
      { HolidayDate: '2026-01-06', Label: 'Epifania', IsBlocked: 'TRUE' },
      { HolidayDate: '2026-04-05', Label: 'Pasqua', IsBlocked: 'TRUE' },
      { HolidayDate: '2026-04-06', Label: "Lunedì dell'Angelo", IsBlocked: 'TRUE' },
      { HolidayDate: '2026-04-25', Label: 'Festa della Liberazione', IsBlocked: 'TRUE' },
      { HolidayDate: '2026-05-01', Label: 'Festa del Lavoro', IsBlocked: 'TRUE' },
      { HolidayDate: '2026-06-02', Label: 'Festa della Repubblica', IsBlocked: 'TRUE' },
      { HolidayDate: '2026-08-15', Label: 'Ferragosto', IsBlocked: 'TRUE' },
      { HolidayDate: '2026-11-01', Label: 'Ognissanti', IsBlocked: 'TRUE' },
      { HolidayDate: '2026-12-08', Label: 'Immacolata Concezione', IsBlocked: 'TRUE' },
      { HolidayDate: '2026-12-25', Label: 'Natale', IsBlocked: 'TRUE' },
      { HolidayDate: '2026-12-26', Label: 'Santo Stefano', IsBlocked: 'TRUE' }
    ]);
  },

  ensureClosures: function () {
    var headers = ['ClosureId', 'StartDate', 'EndDate', 'StartTime', 'EndTime', 'Label', 'IsBlocked', 'Notes'];
    this.ensureSheetStructure_(ROOMS_APP.SHEET_NAMES.CLOSURES, headers);
  },

  ensureSpecialOpenings: function () {
    var headers = ['OpeningId', 'Date', 'OpenTime', 'CloseTime', 'Label', 'IsEnabled', 'Notes'];
    this.ensureSheetStructure_(ROOMS_APP.SHEET_NAMES.SPECIAL_OPENINGS, headers);
  },

  ensureTimetableOccupancy: function () {
    var headers = [
      'OccupancyId',
      'SourceType',
      'ResourceId',
      'ResourceLabel',
      'Weekday',
      'Period',
      'StartTime',
      'EndTime',
      'ClassCode',
      'TeacherName',
      'DisplayLabel',
      'IsActive',
      'Notes',
      'CreatedAtISO',
      'UpdatedAtISO'
    ];
    this.ensureSheetStructure_(ROOMS_APP.SHEET_NAMES.TIMETABLE_OCCUPANCY, headers);
  },

  ensureTimetableDocentiRaw: function () {
    var headers = [
      'RawId',
      'TeacherName',
      'Weekday',
      'Period',
      'ClassCode',
      'IsActive',
      'Notes',
      'UpdatedAtISO'
    ];
    this.ensureSheetStructure_(ROOMS_APP.SHEET_NAMES.TIMETABLE_DOCENTI_RAW, headers);
  },

  ensureTimetableSpacesRaw: function () {
    var headers = [
      'RawId',
      'ResourceId',
      'ResourceLabel',
      'Weekday',
      'Period',
      'ClassCode',
      'TeacherName',
      'IsActive',
      'Notes',
      'UpdatedAtISO'
    ];
    this.ensureSheetStructure_(ROOMS_APP.SHEET_NAMES.TIMETABLE_SPACES_RAW, headers);
  },

  ensureAulaMagnaEvents: function () {
    var headers = [
      'EventId',
      'ResourceId',
      'EventDate',
      'StartTime',
      'EndTime',
      'EventName',
      'IsActive',
      'Notes',
      'CreatedAtISO',
      'UpdatedAtISO'
    ];
    this.ensureSheetStructure_(ROOMS_APP.SHEET_NAMES.AULA_MAGNA_EVENTS, headers);
  },

  ensureReplacementClassOut: function () {
    this.ensureSheetStructure_(ROOMS_APP.SHEET_NAMES.REPL_CLASS_OUT, [
      'Date',
      'ClassCode',
      'IsOut',
      'Notes',
      'UpdatedAtISO',
      'UpdatedBy'
    ]);
  },

  ensureReplacementDayTeachers: function () {
    this.ensureSheetStructure_(ROOMS_APP.SHEET_NAMES.REPL_DAY_TEACHERS, [
      'Date',
      'TeacherEmail',
      'TeacherName',
      'Absent',
      'Accompanist',
      'AccompaniedClasses',
      'Notes',
      'UpdatedAtISO',
      'UpdatedBy'
    ]);
  },

  ensureReplacementFieldTrips: function () {
    this.ensureSheetStructure_(ROOMS_APP.SHEET_NAMES.REPL_FIELD_TRIPS, [
      'TripId',
      'TripType',
      'ClassCode',
      'Title',
      'StartDate',
      'EndDate',
      'StartTime',
      'EndTime',
      'Notes',
      'Enabled',
      'UpdatedAtISO',
      'UpdatedBy'
    ]);
  },

  ensureReplacementFieldTripTeachers: function () {
    this.ensureSheetStructure_(ROOMS_APP.SHEET_NAMES.REPL_FIELD_TRIP_TEACHERS, [
      'TripId',
      'TeacherEmail',
      'TeacherName',
      'Role',
      'Notes',
      'UpdatedAtISO',
      'UpdatedBy'
    ]);
  },

  ensureReplacementHourlyAbsences: function () {
    this.ensureSheetStructure_(ROOMS_APP.SHEET_NAMES.REPL_HOURLY_ABSENCES, [
      'Date',
      'TeacherEmail',
      'TeacherName',
      'Period',
      'Reason',
      'RecoveryRequired',
      'RecoveryStatus',
      'RecoveredOnDate',
      'RecoveredByAssignmentKey',
      'Notes',
      'UpdatedAtISO',
      'UpdatedBy'
    ]);
  },

  ensureReplacementAssignments: function () {
    this.ensureSheetStructure_(ROOMS_APP.SHEET_NAMES.REPL_ASSIGNMENTS, [
      'Date',
      'Period',
      'ClassCode',
      'OriginalTeacherEmail',
      'OriginalTeacherName',
      'OriginalStatus',
      'ClassHandlingType',
      'HandlingType',
      'ReplacementTeacherEmail',
      'ReplacementTeacherName',
      'ReplacementSource',
      'ReplacementStatus',
      'RecoverySourceDate',
      'RecoverySourcePeriod',
      'ShiftOriginPeriod',
      'ShiftTargetPeriod',
      'ShiftTeacherEmail',
      'ShiftTeacherName',
      'Notes',
      'UpdatedAtISO',
      'UpdatedBy'
    ]);
  },

  ensureReplacementLongAssignments: function () {
    this.ensureSheetStructure_(ROOMS_APP.SHEET_NAMES.REPL_LONG_ASSIGNMENTS, [
      'Enabled',
      'OriginalTeacherEmail',
      'OriginalTeacherName',
      'ReplacementTeacherSurname',
      'ReplacementTeacherName',
      'ReplacementTeacherDisplayName',
      'StartDate',
      'EndDate',
      'Reason',
      'Notes',
      'UpdatedAtISO',
      'UpdatedBy'
    ]);
  },

  ensureReportRecipients: function () {
    this.ensureSheetStructure_(ROOMS_APP.SHEET_NAMES.REPORT_RECIPIENTS, [
      'Enabled',
      'ReportType',
      'RecipientType',
      'Email',
      'Notes'
    ]);
  },

  ensureReportLog: function () {
    this.ensureSheetStructure_(ROOMS_APP.SHEET_NAMES.REPORT_LOG, [
      'ReportType',
      'ReferenceDate',
      'SentAtISO',
      'SentBy',
      'Recipients',
      'Subject',
      'Status',
      'Notes'
    ]);
  },

  ensurePolicyOverrides: function () {
    var headers = ['OverrideId', 'ResourceId', 'BookingDate', 'StartTime', 'EndTime', 'RuleKey', 'RuleValue', 'IsEnabled', 'Notes'];
    this.ensureSheetStructure_(ROOMS_APP.SHEET_NAMES.POLICY_OVERRIDES, headers);
  },

  ensureBookings_: function () {
    this.ensureSheetStructure_(ROOMS_APP.SHEET_NAMES.BOOKINGS, [
      'BookingId',
      'SeriesId',
      'ResourceId',
      'BookingDate',
      'StartTime',
      'EndTime',
      'StartISO',
      'EndISO',
      'Title',
      'BookerEmail',
      'BookerName',
      'BookerSurname',
      'Status',
      'CreatedAtISO',
      'UpdatedAtISO',
      'CancelledAtISO',
      'Notes'
    ]);
  },

  ensureAudit_: function () {
    this.ensureSheetStructure_(ROOMS_APP.SHEET_NAMES.AUDIT, [
      'TimestampISO',
      'Action',
      'BookingId',
      'SeriesId',
      'ResourceId',
      'ActorEmail',
      'Result',
      'PayloadJson'
    ]);
  },

  ensureSheetStructure_: function (sheetName, headers) {
    var sheet = ROOMS_APP.DB.getOrCreateSheet(sheetName);
    var hadRows = sheet.getLastRow() > 0;
    if (sheet.getMaxColumns() < headers.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
    }

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    SpreadsheetApp.flush();
    this.setPlainTextSheet_(sheet, sheetName, headers.length);
    this.applyManagedSheetFormatting_(sheet, headers, {
      autoResizeColumns: !hadRows
    });
    ROOMS_APP.DB.invalidateSheetCache_(sheetName);
  },

  ensureOrderedSheetStructure_: function (sheetName, headers) {
    var sheet = ROOMS_APP.DB.getOrCreateSheet(sheetName);
    var hasHeaderRow = sheet.getLastRow() >= 1;
    var existingHeaders = hasHeaderRow && sheet.getLastColumn() > 0
      ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
      : [];
    var currentHeaders = existingHeaders.slice();
    var changed = false;
    var index;

    if (!currentHeaders.length) {
      if (sheet.getMaxColumns() < headers.length) {
        sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
      }
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      SpreadsheetApp.flush();
      currentHeaders = headers.slice();
      changed = true;
    } else {
      for (index = 0; index < headers.length; index += 1) {
        if (currentHeaders[index] === headers[index]) {
          continue;
        }
        if (currentHeaders.indexOf(headers[index]) >= 0) {
          continue;
        }

        sheet.insertColumnBefore(index + 1);
        sheet.getRange(1, index + 1).setValue(headers[index]);
        currentHeaders.splice(index, 0, headers[index]);
        changed = true;
      }
    }

    this.setPlainTextSheet_(sheet, sheetName, currentHeaders.length);
    this.applyManagedSheetFormatting_(sheet, currentHeaders, {
      autoResizeColumns: changed
    });
    if (changed) {
      ROOMS_APP.DB.invalidateSheetCache_(sheetName);
    }
  },

  applyManagedSheetFormatting_: function (sheet, headers, options) {
    var headerCount = headers && headers.length ? headers.length : 0;
    var settings = options || {};
    var index;
    if (!sheet || !headerCount) {
      return;
    }

    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headerCount)
      .setFontWeight('bold')
      .setBackground('#dbeafe')
      .setWrap(true)
      .setVerticalAlignment('middle');
    if (settings.autoResizeColumns !== false) {
      sheet.autoResizeColumns(1, headerCount);
    }

    for (index = 0; index < headerCount; index += 1) {
      this.applyMinimumColumnWidth_(sheet, index + 1, headers[index]);
    }
  },

  applyMinimumColumnWidth_: function (sheet, columnIndex, headerLabel) {
    var normalizedHeader = ROOMS_APP.normalizeString(headerLabel);
    var explicitMinimum = this.HEADER_MIN_WIDTHS_[normalizedHeader] || 0;
    var computedMinimum = this.computeHeaderDrivenMinWidth_(normalizedHeader);
    var minimumWidth = Math.max(explicitMinimum, computedMinimum);
    var currentWidth;
    if (!minimumWidth) {
      return;
    }
    currentWidth = sheet.getColumnWidth(columnIndex);
    if (currentWidth < minimumWidth) {
      sheet.setColumnWidth(columnIndex, minimumWidth);
    }
  },

  computeHeaderDrivenMinWidth_: function (headerLabel) {
    var normalized = ROOMS_APP.normalizeString(headerLabel);
    var computed;
    if (!normalized) {
      return 0;
    }
    computed = (normalized.length * 7) + 34;
    computed = Math.max(computed, 110);
    computed = Math.min(computed, 240);
    return computed;
  },

  seedMissingRows_: function (sheetName, keyField, rows) {
    var existing = ROOMS_APP.DB.readRows(sheetName);
    var existingKeys = {};
    existing.forEach(function (row) {
      existingKeys[row[keyField]] = true;
    });

    var missingRows = rows.filter(function (row) {
      return !existingKeys[row[keyField]];
    });

    ROOMS_APP.DB.appendRows(sheetName, missingRows);
  },

  seedAdminsFromLegacyConfig_: function (headers) {
    var sheetName = ROOMS_APP.SHEET_NAMES.ADMINS;
    var existingRows = ROOMS_APP.DB.readRows(sheetName);
    var existingByEmail = {};

    existingRows.forEach(function (row) {
      var email = ROOMS_APP.normalizeEmail(row.Email);
      if (email && !existingByEmail[email]) {
        existingByEmail[email] = row;
      }
    });

    var configRows = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.CONFIG);
    var adminConfigRow = configRows.filter(function (row) {
      return row.Key === 'ADMIN_GROUP_EMAIL';
    })[0];
    var legacyEmails = String(adminConfigRow && adminConfigRow.Value || '')
      .split(',')
      .map(function (entry) {
        return ROOMS_APP.normalizeEmail(entry);
      })
      .filter(function (entry) {
        return entry && entry.charAt(0) !== '@';
      });

    var toAppend = [];
    legacyEmails.forEach(function (email) {
      if (existingByEmail[email]) {
        return;
      }
      toAppend.push({
        Email: email,
        Role: 'ADMIN',
        Enabled: 'TRUE',
        Notes: 'Migrato da CONFIG.ADMIN_GROUP_EMAIL'
      });
    });

    if (toAppend.length) {
      ROOMS_APP.DB.appendRows(sheetName, toAppend);
    }
  },

  syncResourcesFromCanonical_: function (headers, canonicalRows) {
    var sheetName = ROOMS_APP.SHEET_NAMES.RESOURCES;
    var canonicalOnlyRows = this.normalizeCanonicalResourceRows_(headers, canonicalRows);
    var self = this;
    var existingById = {};
    ROOMS_APP.DB.readRows(sheetName).forEach(function (row) {
      var resourceId = ROOMS_APP.normalizeString(row.ResourceId);
      if (resourceId && !existingById[resourceId]) {
        existingById[resourceId] = row;
      }
    });

    canonicalOnlyRows.forEach(function (row) {
      var existing = existingById[ROOMS_APP.normalizeString(row.ResourceId)] || null;
      if (!existing) {
        row.OpenTime = self.normalizeOptionalTimeValue_(row.OpenTime);
        row.CloseTime = self.normalizeOptionalTimeValue_(row.CloseTime);
        return;
      }
      if (!self.normalizeOptionalTimeValue_(row.OpenTime)) {
        row.OpenTime = self.normalizeOptionalTimeValue_(existing.OpenTime);
      }
      if (!self.normalizeOptionalTimeValue_(row.CloseTime)) {
        row.CloseTime = self.normalizeOptionalTimeValue_(existing.CloseTime);
      }
      row.OpenTime = self.normalizeOptionalTimeValue_(row.OpenTime);
      row.CloseTime = self.normalizeOptionalTimeValue_(row.CloseTime);
    });
    // Canonical inventory is the single source of truth for ROOMS_RESOURCES.
    // Existing rows are fully rewritten, preserving only per-resource open/close overrides.
    ROOMS_APP.DB.replaceRows(sheetName, headers, canonicalOnlyRows);
  },

  normalizeCanonicalResourceRows_: function (headers, canonicalRows) {
    var self = this;
    var seenResourceIds = {};
    var seenSortKeys = {};
    var normalizedRows = [];

    (canonicalRows || []).forEach(function (canonicalRow) {
      var row = {};
      headers.forEach(function (header) {
        row[header] = Object.prototype.hasOwnProperty.call(canonicalRow, header) ? canonicalRow[header] : '';
      });
      row.OpenTime = self.normalizeOptionalTimeValue_(row.OpenTime);
      row.CloseTime = self.normalizeOptionalTimeValue_(row.CloseTime);

      var resourceId = ROOMS_APP.normalizeString(row.ResourceId);
      var sortKey = ROOMS_APP.normalizeString(row.SortKey);
      if (!resourceId) {
        throw new Error('Canonical resource row without ResourceId: ' + JSON.stringify(row));
      }
      if (!sortKey) {
        throw new Error('Canonical resource row without SortKey: ' + resourceId);
      }
      if (seenResourceIds[resourceId]) {
        throw new Error('Duplicate canonical ResourceId: ' + resourceId);
      }
      if (seenSortKeys[sortKey]) {
        throw new Error('Duplicate canonical SortKey: ' + sortKey);
      }

      seenResourceIds[resourceId] = true;
      seenSortKeys[sortKey] = true;
      normalizedRows.push(row);
    });

    return normalizedRows;
  },

  normalizeOptionalTimeValue_: function (value) {
    if (value == null || value === '') {
      return '';
    }
    if (typeof value === 'boolean') {
      return '';
    }
    if (value instanceof Date && !isNaN(value.getTime())) {
      return Utilities.formatDate(value, ROOMS_APP.getTimezone(), 'HH:mm');
    }

    var normalized = ROOMS_APP.normalizeString(value);
    if (!normalized) {
      return '';
    }

    var upper = normalized.toUpperCase();
    if (upper === 'TRUE' || upper === 'FALSE') {
      return '';
    }

    return /^\d{2}:\d{2}$/.test(normalized) ? normalized : '';
  },

  setPlainTextSheet_: function (sheet, sheetName, columnCount) {
    var effectiveColumnCount;
    var effectiveRowCount;
    if (!this.PLAIN_TEXT_SHEETS_[sheetName]) {
      return;
    }

    effectiveColumnCount = Math.max(Number(columnCount) || 0, sheet.getLastColumn(), 1);
    effectiveRowCount = Math.max(sheet.getLastRow(), 1);
    sheet.getRange(1, 1, effectiveRowCount, effectiveColumnCount).setNumberFormat('@');
  },

  buildResourceRows_: function () {
    var specs = [];

    function addGroup(areaCode, areaLabel, floorCode, floorLabel, sideCode, sideLabel, layoutPage, layoutRow, displayNames) {
      displayNames.forEach(function (displayName, index) {
        specs.push({
          ResourceId: ROOMS_APP.slugify(displayName),
          DisplayName: displayName,
          AreaCode: areaCode,
          AreaLabel: areaLabel,
          FloorCode: floorCode,
          FloorLabel: floorLabel,
          SideCode: sideCode,
          SideLabel: sideLabel,
          LayoutPage: String(layoutPage),
          LayoutRow: String(layoutRow),
          LayoutCol: String(index + 1),
          LayoutColSpan: '1',
          LayoutRowSpan: '1',
          OpenTime: '',
          CloseTime: '',
          IsBookable: 'TRUE',
          IsActive: 'TRUE',
          SortKey: areaCode + '_' + ('00' + (index + 1)).slice(-2),
          Notes: ''
        });
      });
    }

    addGroup('GYM', 'Corridoio Palestra', 'F0', 'Piano Terra', 'GYM', 'Corridoio', 2, 1, [
      '4AE',
      'PALESTRA PICCOLA',
      'PALESTRA GRANDE'
    ]);
    addGroup('F1M', 'Piano Primo Mezzato', 'F1', 'Piano Primo', 'MEZZATO', 'Mezzato', 1, 3, [
      'AULA CONSILIARE'
    ]);
    addGroup('F2M', 'Piano Secondo Mezzato', 'F2', 'Piano Secondo', 'MEZZATO', 'Mezzato', 1, 6, [
      '5AE',
      '1ALL',
      '1BL',
      '2AL',
      '2BL',
      'LAB Biologia',
      'LAB Chimica',
      'LAB Fisica'
    ]);
    addGroup('F1L', 'Piano Primo Lato Sinistro', 'F1', 'Piano Primo', 'LEFT', 'Lato Sinistro', 1, 1, [
      '1A',
      '1B',
      '1C',
      '2A',
      '2B',
      '2C'
    ]);
    addGroup('F1R', 'Piano Primo Lato Destro', 'F1', 'Piano Primo', 'RIGHT', 'Lato Destro', 1, 2, [
      '3AT',
      '3AM',
      '4AT',
      '4AM',
      '5AT',
      '5AM'
    ]);
    addGroup('F2L', 'Piano Secondo Lato Sinistro', 'F2', 'Piano Secondo', 'LEFT', 'Lato Sinistro', 1, 4, [
      '1CLS',
      '2CLS',
      '3AL',
      '3BL',
      '4AL',
      '4BLS'
    ]);
    addGroup('F2R', 'Piano Secondo Lato Destro', 'F2', 'Piano Secondo', 'RIGHT', 'Lato Destro', 1, 5, [
      '3CLS',
      '4CLS',
      '4DLS',
      '5AL',
      '5BLS',
      '5CLS'
    ]);
    addGroup('LAB', 'Corridoio Laboratori', 'F1', 'Piano Primo', 'LABS', 'Laboratori', 2, 2, [
      'AULA POLIFUNZIONALE',
      'LAB DISEGNO',
      'LAB Informatica 1',
      'LAB Informatica 2',
      'LAB Informatica 3',
      'LAB Informatica 5',
      'LAB Automazione/Sistemi',
      'LAB Macchine Utensili',
      'LAB Tecnologia',
      'LAB TPS elettronica',
      'LAB TPS elettrotecnica',
      'LAB Mis. Elettroniche',
      'LAB Mis. Elettriche'
    ]);
    addGroup('F0', 'Piano Terra', 'F0', 'Piano Terra', 'CENTER', 'Centrale', 2, 3, [
      '1AC',
      '3AE',
      'CIC',
      'BIBLIOTECA'
    ]);
    addGroup('PA1A', 'Piano Primo - Area Magna', 'F1', 'Piano Primo', 'AULA_MAGNA', 'Aula Magna', 1, 7, [
      'AULA MAGNA'
    ]);

    return specs;
  }
};

function ensureAll() {
  ROOMS_APP.Schema.ensureAll();
}

function ensureConfig() {
  ROOMS_APP.Schema.ensureConfig();
}

function ensureResources() {
  ROOMS_APP.Schema.ensureResources();
}

function ensureAdmins() {
  ROOMS_APP.Schema.ensureAdmins();
}

function ensureWeekSchedule() {
  ROOMS_APP.Schema.ensureWeekSchedule();
}

function ensureHolidays() {
  ROOMS_APP.Schema.ensureHolidays();
}

function ensureClosures() {
  ROOMS_APP.Schema.ensureClosures();
}

function ensureSpecialOpenings() {
  ROOMS_APP.Schema.ensureSpecialOpenings();
}

function ensureAulaMagnaEvents() {
  ROOMS_APP.Schema.ensureAulaMagnaEvents();
}

function ensurePolicyOverrides() {
  ROOMS_APP.Schema.ensurePolicyOverrides();
}
