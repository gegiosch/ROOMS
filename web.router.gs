var ROOMS_APP = ROOMS_APP || {};

function doGet(e) {
  var startedAt = Date.now();
  var params = (e && e.parameter) || {};
  var page = params.page || 'board';

  try {
    var response;
    if (page === 'api') {
      response = jsonResponse_(routeApiRequest_(params));
      logTiming_('doGet', startedAt, {
        page: page,
        mode: 'api'
      });
      return response;
    }

    if (page === 'room') {
      response = renderBoardPage_(params);
      logTiming_('doGet', startedAt, {
        page: page,
        mode: 'render-board'
      });
      return response;
    }

    if (page === 'admin') {
      ROOMS_APP.Auth.requireAdmin();
      response = renderTemplate_('ui.admin', {
        pageTitle: 'Admin',
        initialModelJson: JSON.stringify({
          page: 'admin'
        })
      });
      logTiming_('doGet', startedAt, {
        page: page,
        mode: 'render-admin'
      });
      return response;
    }

    response = renderBoardPage_(params);
    logTiming_('doGet', startedAt, {
      page: page,
      mode: 'render-board'
    });
    return response;
  } catch (error) {
    logTiming_('doGet', startedAt, {
      page: page,
      mode: 'error',
      error: String(error && error.message ? error.message : error)
    });
    throw error;
  }
}

function doPost(e) {
  var startedAt = Date.now();
  var body = ROOMS_APP.parseJson((e && e.postData && e.postData.contents) || '{}', {});
  var action = body.action || body.pageAction || '';
  try {
    var response = jsonResponse_(routeApiRequest_(body));
    logTiming_('doPost', startedAt, {
      action: action
    });
    return response;
  } catch (error) {
    logTiming_('doPost', startedAt, {
      action: action,
      error: String(error && error.message ? error.message : error)
    });
    throw error;
  }
}

function runSetup() {
  var actor = ROOMS_APP.Auth.requireAdmin();
  ROOMS_APP.Schema.ensureAll();
  ROOMS_APP.invalidateConfigCache();
  return {
    ok: true,
    message: 'Setup completato.',
    actorEmail: actor.email,
    executedAtISO: ROOMS_APP.toIsoDateTime(new Date())
  };
}

function getBoardViewModel() {
  var startedAt = Date.now();
  try {
    var model = ROOMS_APP.Board.getBoardViewModel();
    logTiming_('getBoardViewModel', startedAt, {
      pageCount: model && model.pageCount || 0
    });
    return model;
  } catch (error) {
    logTiming_('getBoardViewModel', startedAt, {
      error: String(error && error.message ? error.message : error)
    });
    throw error;
  }
}

function getRoomViewModel(resourceId, dateString) {
  return ROOMS_APP.Booking.getRoomViewModel(ROOMS_APP.normalizeString(resourceId), dateString);
}

function getRoomPanelData(resourceId, dateString) {
  var startedAt = Date.now();
  var normalizedResourceId = ROOMS_APP.normalizeString(resourceId);
  var targetDate = dateString || ROOMS_APP.toIsoDate(new Date());
  try {
    var model = ROOMS_APP.Booking.getRoomViewModel(normalizedResourceId, targetDate);
    logTiming_('getRoomPanelData', startedAt, {
      resourceId: normalizedResourceId,
      date: targetDate,
      ok: model && model.ok ? 'TRUE' : 'FALSE'
    });
    return model;
  } catch (error) {
    logTiming_('getRoomPanelData', startedAt, {
      resourceId: normalizedResourceId,
      date: targetDate,
      error: String(error && error.message ? error.message : error)
    });
    throw error;
  }
}

function createRoomBooking(payload) {
  return ROOMS_APP.Booking.createBooking(payload);
}

function cancelRoomBooking(bookingId, notes) {
  return ROOMS_APP.Booking.cancelBooking(bookingId, notes);
}

function updateRoomBooking(bookingId, payload) {
  return ROOMS_APP.Booking.updateBooking(bookingId, payload || {});
}

function applyRoomPanelChanges(resourceId, dateString, changes) {
  return ROOMS_APP.Booking.applyRoomChanges(
    ROOMS_APP.normalizeString(resourceId),
    dateString || ROOMS_APP.toIsoDate(new Date()),
    changes || {}
  );
}

function previewRecurringRoomBooking(payload) {
  return ROOMS_APP.Recurring.previewWeekly(payload);
}

function commitRecurringRoomBooking(payload) {
  return ROOMS_APP.Recurring.commitWeekly(payload);
}

function importTimetableClassrooms() {
  ROOMS_APP.Auth.requireAdmin();
  return ROOMS_APP.Timetable.importTimetableClassrooms();
}

function importTimetableSpaces() {
  ROOMS_APP.Auth.requireAdmin();
  return ROOMS_APP.Timetable.importTimetableSpaces();
}

function rebuildTimetableOccupancy() {
  ROOMS_APP.Auth.requireAdmin();
  return ROOMS_APP.Timetable.rebuildTimetableOccupancy();
}

function rebuildTimetableOccupancyFromSheets() {
  ROOMS_APP.Auth.requireAdmin();
  return ROOMS_APP.Timetable.rebuildTimetableOccupancyFromSheets();
}

function getAdminBootstrap() {
  var user = ROOMS_APP.Auth.requireAdmin();
  var tableNames = [
    ROOMS_APP.SHEET_NAMES.CONFIG,
    ROOMS_APP.SHEET_NAMES.ADMINS,
    ROOMS_APP.SHEET_NAMES.RESOURCES,
    ROOMS_APP.SHEET_NAMES.WEEK_SCHEDULE,
    ROOMS_APP.SHEET_NAMES.HOLIDAYS,
    ROOMS_APP.SHEET_NAMES.CLOSURES,
    ROOMS_APP.SHEET_NAMES.SPECIAL_OPENINGS,
    ROOMS_APP.SHEET_NAMES.AULA_MAGNA_EVENTS
  ];
  var tables = {};

  tableNames.forEach(function (tableName) {
    tables[tableName] = {
      headers: ROOMS_APP.DB.getHeaders(tableName),
      rows: ROOMS_APP.DB.readRows(tableName)
    };
  });

  tables[ROOMS_APP.SHEET_NAMES.AUDIT] = {
    headers: ROOMS_APP.DB.getHeaders(ROOMS_APP.SHEET_NAMES.AUDIT),
    rows: ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.AUDIT).slice(-200).reverse()
  };

  return {
    user: user,
    tables: tables
  };
}

function adminReplaceTable(tableName, rows) {
  ROOMS_APP.Auth.requireAdmin();
  var allowed = {};
  allowed[ROOMS_APP.SHEET_NAMES.CONFIG] = true;
  allowed[ROOMS_APP.SHEET_NAMES.ADMINS] = true;
  allowed[ROOMS_APP.SHEET_NAMES.RESOURCES] = true;
  allowed[ROOMS_APP.SHEET_NAMES.WEEK_SCHEDULE] = true;
  allowed[ROOMS_APP.SHEET_NAMES.HOLIDAYS] = true;
  allowed[ROOMS_APP.SHEET_NAMES.CLOSURES] = true;
  allowed[ROOMS_APP.SHEET_NAMES.SPECIAL_OPENINGS] = true;
  allowed[ROOMS_APP.SHEET_NAMES.AULA_MAGNA_EVENTS] = true;

  if (!allowed[tableName]) {
    throw new Error('Table is not editable: ' + tableName);
  }

  ROOMS_APP.DB.replaceRows(tableName, ROOMS_APP.DB.getHeaders(tableName), rows || []);
  if (tableName === ROOMS_APP.SHEET_NAMES.CONFIG) {
    ROOMS_APP.invalidateConfigCache();
  }
  ROOMS_APP.Booking.writeAudit_('ADMIN_REPLACE_TABLE', '', '', tableName, ROOMS_APP.getCurrentUserEmail(), 'OK', {
    rowCount: (rows || []).length
  });

  return {
    tableName: tableName,
    rowCount: (rows || []).length
  };
}

function include_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function renderTemplate_(filename, viewModel) {
  var template = HtmlService.createTemplateFromFile(filename);
  Object.keys(viewModel || {}).forEach(function (key) {
    template[key] = viewModel[key];
  });

  return template.evaluate()
    .setTitle(ROOMS_APP.getConfigValue('APP_NAME', 'ROOMS') + ' | ' + (viewModel.pageTitle || ''))
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function routeApiRequest_(payload) {
  var action = payload.action || payload.pageAction;
  if (action === 'board') {
    return ROOMS_APP.Board.getBoardViewModel();
  }
  if (action === 'room') {
    return getRoomPanelData(normalizeRoomIdParam_(payload), payload.date);
  }
  if (action === 'roomPanel') {
    return getRoomPanelData(normalizeRoomIdParam_(payload), payload.date);
  }
  if (action === 'createBooking') {
    return ROOMS_APP.Booking.createBooking(payload);
  }
  if (action === 'updateBooking') {
    return ROOMS_APP.Booking.updateBooking(payload.bookingId, payload);
  }
  if (action === 'cancelBooking') {
    return ROOMS_APP.Booking.cancelBooking(payload.bookingId, payload.notes);
  }
  if (action === 'applyRoomPanelChanges') {
    return ROOMS_APP.Booking.applyRoomChanges(
      normalizeRoomIdParam_(payload),
      payload.date,
      payload.changes || payload
    );
  }
  if (action === 'previewRecurring') {
    return ROOMS_APP.Recurring.previewWeekly(payload);
  }
  if (action === 'commitRecurring') {
    return ROOMS_APP.Recurring.commitWeekly(payload);
  }
  if (action === 'importTimetableClassrooms') {
    ROOMS_APP.Auth.requireAdmin();
    return ROOMS_APP.Timetable.importTimetableClassrooms();
  }
  if (action === 'importTimetableSpaces') {
    ROOMS_APP.Auth.requireAdmin();
    return ROOMS_APP.Timetable.importTimetableSpaces();
  }
  if (action === 'rebuildTimetableOccupancy') {
    ROOMS_APP.Auth.requireAdmin();
    return ROOMS_APP.Timetable.rebuildTimetableOccupancy();
  }
  if (action === 'rebuildTimetableOccupancyFromSheets') {
    ROOMS_APP.Auth.requireAdmin();
    return ROOMS_APP.Timetable.rebuildTimetableOccupancyFromSheets();
  }
  if (action === 'runSetup') {
    return runSetup();
  }

  return {
    ok: true,
    message: 'ROOMS API online.'
  };
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function normalizeRoomIdParam_(payload) {
  return ROOMS_APP.normalizeString((payload && (payload.resourceId || payload.room || payload.roomId)) || '');
}

function renderBoardPage_(params) {
  var bootstrapModel = {
    generatedAtISO: '',
    date: '',
    refreshSec: 60,
    rotationSec: 15,
    pageCount: 1,
    fullscreenCompactEnabled: true,
    appName: 'ROOMS',
    schoolName: '',
    user: {
      email: '',
      isAdmin: false
    },
    branchOrder: [],
    branchLabels: {},
    pages: []
  };
  return renderTemplate_('ui.board', {
    pageTitle: 'Board',
    initialModelJson: JSON.stringify(bootstrapModel),
    initialRoomId: normalizeRoomIdParam_(params)
  });
}

function logTiming_(label, startedAt, payload) {
  var elapsedMs = Date.now() - Number(startedAt || Date.now());
  var details = payload || {};
  var serialized;
  try {
    serialized = JSON.stringify(details);
  } catch (error) {
    serialized = '{}';
  }
  Logger.log('[PERF] %s %sms %s', label, elapsedMs, serialized);
}
