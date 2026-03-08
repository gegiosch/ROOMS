var ROOMS_APP = ROOMS_APP || {};

ROOMS_APP.SPREADSHEET_ID = '1f7KIwGqTOX57D-xv8Bs6yFVXwL3z71bDabkW4PEZ5pI';
ROOMS_APP.DEFAULT_TIMEZONE = 'Europe/Rome';
ROOMS_APP.CONFIG_CACHE = null;

ROOMS_APP.SHEET_NAMES = {
  CONFIG: 'CONFIG',
  RESOURCES: 'ROOMS_RESOURCES',
  BOOKINGS: 'ROOMS_BOOKINGS',
  AUDIT: 'ROOMS_AUDIT',
  HOLIDAYS: 'ROOMS_HOLIDAYS',
  CLOSURES: 'ROOMS_CLOSURES',
  WEEK_SCHEDULE: 'ROOMS_WEEK_SCHEDULE',
  SPECIAL_OPENINGS: 'ROOMS_SPECIAL_OPENINGS',
  POLICY_OVERRIDES: 'ROOMS_POLICY_OVERRIDES'
};

ROOMS_APP.DEFAULT_CONFIG_ROWS = [
  { Key: 'APP_NAME', Value: 'ROOMS', Notes: 'Standalone room management web app.' },
  { Key: 'SCHOOL_NAME', Value: 'IIS Alessandrini', Notes: 'School name displayed in the UI.' },
  { Key: 'TIMEZONE', Value: 'Europe/Rome', Notes: 'Apps Script and booking timezone.' },
  { Key: 'BOOKING_ENABLED', Value: 'TRUE', Notes: 'Global booking switch.' },
  { Key: 'ALLOW_RECURRING', Value: 'TRUE', Notes: 'Enable weekly recurring reservations.' },
  { Key: 'MAX_DURATION_MIN', Value: '180', Notes: 'Normal user max duration in minutes.' },
  { Key: 'MAX_DAYS_AHEAD', Value: '30', Notes: 'Normal user booking horizon in days.' },
  { Key: 'ADMIN_GROUP_EMAIL', Value: '', Notes: 'Comma-separated admin emails or @domain entries.' },
  { Key: 'RECURRING_MAX_OCCURRENCES', Value: '20', Notes: 'Max generated weekly occurrences.' },
  { Key: 'RECURRING_MAX_WEEKS', Value: '12', Notes: 'Max weekly recurrence span.' },
  { Key: 'SLOT_MINUTES', Value: '30', Notes: 'Slot granularity used by free slot calculations.' },
  { Key: 'OPEN_TIME', Value: '08:00', Notes: 'Fallback opening time.' },
  { Key: 'CLOSE_TIME', Value: '18:00', Notes: 'Fallback closing time.' },
  { Key: 'BOARD_REFRESH_SEC', Value: '60', Notes: 'Refresh interval for the public board.' },
  { Key: 'BOARD_ROTATION_SEC', Value: '15', Notes: 'Page rotation interval for the public board.' },
  { Key: 'BOARD_PAGE_COUNT', Value: '2', Notes: 'Maximum visible board pages.' },
  { Key: 'SHOW_BOOKER_NAME', Value: 'FALSE', Notes: 'Room detail visibility toggle.' }
];

ROOMS_APP.PALETTE = {
  background: '#020617',
  panel: '#0F172A',
  text: '#E5E7EB',
  accent: '#60A5FA',
  free: '#34D399',
  occupied: '#FB7185',
  nextOccupied: '#FBBF24',
  border: '#1E293B'
};

ROOMS_APP.openSpreadsheet = function () {
  return SpreadsheetApp.openById(ROOMS_APP.SPREADSHEET_ID);
};

ROOMS_APP.getTimezone = function () {
  if (ROOMS_APP.CONFIG_CACHE && ROOMS_APP.CONFIG_CACHE.TIMEZONE) {
    return ROOMS_APP.CONFIG_CACHE.TIMEZONE;
  }

  try {
    return ROOMS_APP.getConfigValue('TIMEZONE', ROOMS_APP.DEFAULT_TIMEZONE);
  } catch (error) {
    return ROOMS_APP.DEFAULT_TIMEZONE;
  }
};

ROOMS_APP.getConfigMap = function (forceRefresh) {
  if (!forceRefresh && ROOMS_APP.CONFIG_CACHE) {
    return ROOMS_APP.CONFIG_CACHE;
  }

  ROOMS_APP.Schema.ensureAll();
  var rows = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.CONFIG);
  var config = {};

  rows.forEach(function (row) {
    config[row.Key] = row.Value;
  });

  ROOMS_APP.CONFIG_CACHE = config;
  return config;
};

ROOMS_APP.invalidateConfigCache = function () {
  ROOMS_APP.CONFIG_CACHE = null;
};

ROOMS_APP.getConfigValue = function (key, fallback) {
  var config = ROOMS_APP.getConfigMap();
  return Object.prototype.hasOwnProperty.call(config, key) ? config[key] : fallback;
};

ROOMS_APP.getBooleanConfig = function (key, fallback) {
  return ROOMS_APP.asBoolean(ROOMS_APP.getConfigValue(key, fallback));
};

ROOMS_APP.getNumberConfig = function (key, fallback) {
  return ROOMS_APP.asNumber(ROOMS_APP.getConfigValue(key, fallback), fallback);
};

ROOMS_APP.getCurrentUserEmail = function () {
  return ROOMS_APP.normalizeEmail(Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '');
};

ROOMS_APP.listConfiguredAdmins = function () {
  return String(ROOMS_APP.getConfigValue('ADMIN_GROUP_EMAIL', ''))
    .split(',')
    .map(function (value) {
      return ROOMS_APP.normalizeEmail(value);
    })
    .filter(function (value) {
      return Boolean(value);
    });
};

ROOMS_APP.asBoolean = function (value) {
  if (typeof value === 'boolean') {
    return value;
  }

  return String(value).toUpperCase() === 'TRUE';
};

ROOMS_APP.asNumber = function (value, fallback) {
  var parsed = Number(value);
  return isNaN(parsed) ? Number(fallback || 0) : parsed;
};

ROOMS_APP.normalizeEmail = function (value) {
  return String(value || '').trim().toLowerCase();
};

ROOMS_APP.normalizeString = function (value) {
  return String(value == null ? '' : value).trim();
};

ROOMS_APP.parseJson = function (value, fallback) {
  try {
    return JSON.parse(value);
  } catch (error) {
    return fallback;
  }
};

ROOMS_APP.toIsoDate = function (value) {
  if (!value) {
    return '';
  }

  if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value)) {
    return value;
  }

  var date = value instanceof Date ? value : new Date(value);
  return Utilities.formatDate(date, ROOMS_APP.getTimezone(), 'yyyy-MM-dd');
};

ROOMS_APP.toTimeString = function (value) {
  if (!value) {
    return '';
  }

  if (typeof value === 'string' && /^\d{2}:\d{2}$/.test(value)) {
    return value;
  }

  var date = value instanceof Date ? value : new Date(value);
  return Utilities.formatDate(date, ROOMS_APP.getTimezone(), 'HH:mm');
};

ROOMS_APP.toIsoDateTime = function (value) {
  var date = value instanceof Date ? value : new Date(value);
  return Utilities.formatDate(date, ROOMS_APP.getTimezone(), "yyyy-MM-dd'T'HH:mm:ss");
};

ROOMS_APP.combineDateTime = function (dateString, timeString) {
  return new Date(dateString + 'T' + timeString + ':00');
};

ROOMS_APP.minutesBetween = function (startTime, endTime) {
  var startParts = String(startTime).split(':');
  var endParts = String(endTime).split(':');
  var start = Number(startParts[0]) * 60 + Number(startParts[1]);
  var end = Number(endParts[0]) * 60 + Number(endParts[1]);
  return end - start;
};

ROOMS_APP.addMinutes = function (timeString, minutes) {
  var base = ROOMS_APP.combineDateTime('2000-01-01', timeString);
  var next = new Date(base.getTime() + minutes * 60000);
  return Utilities.formatDate(next, ROOMS_APP.getTimezone(), 'HH:mm');
};

ROOMS_APP.slugify = function (value) {
  return ROOMS_APP.normalizeString(value)
    .replace(/\s+/g, '_')
    .replace(/[^\w/]+/g, '_')
    .replace(/\/+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '')
    .toUpperCase();
};

ROOMS_APP.getWeekdayName = function (dateString) {
  var weekdayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  return weekdayNames[ROOMS_APP.combineDateTime(dateString, '12:00').getDay()];
};

ROOMS_APP.daysBetween = function (fromDateString, toDateString) {
  var start = ROOMS_APP.combineDateTime(fromDateString, '00:00');
  var end = ROOMS_APP.combineDateTime(toDateString, '00:00');
  return Math.round((end.getTime() - start.getTime()) / 86400000);
};

ROOMS_APP.sortBy = function (rows, keys) {
  return rows.slice().sort(function (left, right) {
    for (var index = 0; index < keys.length; index += 1) {
      var key = keys[index];
      var a = left[key];
      var b = right[key];
      if (a < b) {
        return -1;
      }
      if (a > b) {
        return 1;
      }
    }
    return 0;
  });
};
