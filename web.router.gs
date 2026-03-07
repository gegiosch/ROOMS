var ROOMS_APP = ROOMS_APP || {};

function doGet(e) {
  ROOMS_APP.Schema.ensureAll();
  var page = (e && e.parameter && e.parameter.page) || 'board';

  if (page === 'api') {
    return jsonResponse_(routeApiRequest_((e && e.parameter) || {}));
  }

  if (page === 'room') {
    return renderTemplate_('ui.room', {
      pageTitle: 'Room',
      initialModelJson: JSON.stringify({
        page: 'room',
        roomId: (e && e.parameter && e.parameter.room) || '',
        date: (e && e.parameter && e.parameter.date) || ROOMS_APP.toIsoDate(new Date())
      })
    });
  }

  if (page === 'admin') {
    return renderTemplate_('ui.admin', {
      pageTitle: 'Admin',
      initialModelJson: JSON.stringify({
        page: 'admin'
      })
    });
  }

  return renderTemplate_('ui.board', {
    pageTitle: 'Board',
    initialModelJson: JSON.stringify(ROOMS_APP.Board.getBoardViewModel())
  });
}

function doPost(e) {
  ROOMS_APP.Schema.ensureAll();
  var body = ROOMS_APP.parseJson((e && e.postData && e.postData.contents) || '{}', {});
  return jsonResponse_(routeApiRequest_(body));
}

function getBoardViewModel() {
  return ROOMS_APP.Board.getBoardViewModel();
}

function getRoomViewModel(resourceId, dateString) {
  return ROOMS_APP.Booking.getRoomViewModel(resourceId, dateString);
}

function createRoomBooking(payload) {
  return ROOMS_APP.Booking.createBooking(payload);
}

function cancelRoomBooking(bookingId, notes) {
  return ROOMS_APP.Booking.cancelBooking(bookingId, notes);
}

function previewRecurringRoomBooking(payload) {
  return ROOMS_APP.Recurring.previewWeekly(payload);
}

function commitRecurringRoomBooking(payload) {
  return ROOMS_APP.Recurring.commitWeekly(payload);
}

function getAdminBootstrap() {
  var user = ROOMS_APP.Auth.requireAdmin();
  var tableNames = [
    ROOMS_APP.SHEET_NAMES.CONFIG,
    ROOMS_APP.SHEET_NAMES.RESOURCES,
    ROOMS_APP.SHEET_NAMES.WEEK_SCHEDULE,
    ROOMS_APP.SHEET_NAMES.HOLIDAYS,
    ROOMS_APP.SHEET_NAMES.CLOSURES,
    ROOMS_APP.SHEET_NAMES.SPECIAL_OPENINGS
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
  allowed[ROOMS_APP.SHEET_NAMES.RESOURCES] = true;
  allowed[ROOMS_APP.SHEET_NAMES.WEEK_SCHEDULE] = true;
  allowed[ROOMS_APP.SHEET_NAMES.HOLIDAYS] = true;
  allowed[ROOMS_APP.SHEET_NAMES.CLOSURES] = true;
  allowed[ROOMS_APP.SHEET_NAMES.SPECIAL_OPENINGS] = true;

  if (!allowed[tableName]) {
    throw new Error('Table is not editable: ' + tableName);
  }

  ROOMS_APP.Schema.ensureAll();
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
    return ROOMS_APP.Booking.getRoomViewModel(payload.resourceId, payload.date);
  }
  if (action === 'createBooking') {
    return ROOMS_APP.Booking.createBooking(payload);
  }
  if (action === 'previewRecurring') {
    return ROOMS_APP.Recurring.previewWeekly(payload);
  }
  if (action === 'commitRecurring') {
    return ROOMS_APP.Recurring.commitWeekly(payload);
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
