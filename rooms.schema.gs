var ROOMS_APP = ROOMS_APP || {};

ROOMS_APP.Schema = {
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
    return sheets;
  }()),

  ensureAll: function () {
    this.ensureConfig();
    this.ensureResources();
    this.ensureBookings_();
    this.ensureAudit_();
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
      'IsBookable',
      'IsActive',
      'SortKey',
      'Notes'
    ];
    this.ensureSheetStructure_(ROOMS_APP.SHEET_NAMES.RESOURCES, headers);
    this.seedMissingRows_(ROOMS_APP.SHEET_NAMES.RESOURCES, 'ResourceId', this.buildResourceRows_());
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
    if (sheet.getMaxColumns() < headers.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
    }

    this.setPlainTextSheet_(sheet, sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#dbeafe')
      .setWrap(true);
    sheet.autoResizeColumns(1, headers.length);
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

  setPlainTextSheet_: function (sheet, sheetName) {
    if (!this.PLAIN_TEXT_SHEETS_[sheetName]) {
      return;
    }

    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setNumberFormat('@');
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
      '1CL',
      '2CL',
      '3AL',
      '3BL',
      '4AL',
      '4BL'
    ]);
    addGroup('F2R', 'Piano Secondo Lato Destro', 'F2', 'Piano Secondo', 'RIGHT', 'Lato Destro', 1, 5, [
      '3CL',
      '4CL',
      '4DL',
      '5AL',
      '5BL',
      '5CL'
    ]);
    addGroup('LAB', 'Corridoio Laboratori', 'F1', 'Piano Primo', 'LABS', 'Laboratori', 2, 2, [
      '1AC',
      '3AE',
      'CIC',
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

function ensurePolicyOverrides() {
  ROOMS_APP.Schema.ensurePolicyOverrides();
}
