var ROOMS_APP = ROOMS_APP || {};

function doGet(e) {
  var params = (e && e.parameter) || {};
  var page = params.page || 'board';

  try {
    ROOMS_APP.Schema.ensureAll();

    if (page === 'api') {
      return jsonResponse_(routeApiRequest_(params));
    }

    if (page === 'room') {
      return renderRoomPage_(params);
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
  } catch (error) {
    if (page === 'room') {
      try {
        return renderRoomPage_(params, error);
      } catch (fatalRoomError) {
        return HtmlService.createHtmlOutput(buildPlainRoomFallbackHtml_(params, fatalRoomError))
          .setTitle('ROOMS | Room');
      }
    }

    throw error;
  }
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
  return ROOMS_APP.Booking.getRoomViewModel(ROOMS_APP.normalizeString(resourceId), dateString);
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
    return ROOMS_APP.Booking.getRoomViewModel(normalizeRoomIdParam_(payload), payload.date);
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

function normalizeRoomIdParam_(payload) {
  return ROOMS_APP.normalizeString((payload && (payload.resourceId || payload.room || payload.roomId)) || '');
}

function renderRoomPage_(params, upstreamError) {
  var normalizedRoomId = normalizeRoomIdParam_(params);
  var roomDate = params.date || ROOMS_APP.toIsoDate(new Date());
  var model;

  if (upstreamError) {
    model = ROOMS_APP.Booking.buildRoomFallbackModel_(normalizedRoomId, roomDate, 'Errore caricamento pagina aula');
    model.debugMessage = String(upstreamError && upstreamError.message ? upstreamError.message : upstreamError);
  } else {
    model = ROOMS_APP.Booking.getRoomViewModel(normalizedRoomId, roomDate);
  }

  var safeModel = model && typeof model === 'object'
    ? model
    : ROOMS_APP.Booking.buildRoomFallbackModel_(normalizedRoomId, roomDate, 'Errore caricamento pagina aula');

  var roomFound = Boolean(safeModel.resource && safeModel.resource.ResourceId);
  var roomMeta = safeModel.resource || {};
  var bookings = Array.isArray(safeModel.bookings) ? safeModel.bookings : [];
  var slots = Array.isArray(safeModel.slots) ? safeModel.slots : [];
  var freeSlots = Array.isArray(safeModel.freeSlots) ? safeModel.freeSlots : [];
  var errorMessage = ROOMS_APP.normalizeString(safeModel.errorMessage);

  try {
    return renderTemplate_('ui.room', {
      pageTitle: 'Room',
      initialModelJson: JSON.stringify(safeModel),
      serverRequestedRoomId: normalizedRoomId || '-',
      serverRoomFound: roomFound,
      serverErrorMessage: errorMessage,
      serverRoomName: roomFound ? roomMeta.DisplayName : 'Aula non trovata',
      serverRoomMetaLabel: roomFound
        ? [roomMeta.AreaLabel || '', roomMeta.FloorLabel || '', roomMeta.SideLabel || ''].filter(function (value) { return Boolean(value); }).join(' / ')
        : 'Dati aula non disponibili',
      serverStatus: safeModel.status || 'UNKNOWN',
      serverBookingsCount: bookings.length,
      serverSlotsCount: slots.length,
      serverFreeSlotsCount: freeSlots.length,
      serverDate: safeModel.date || roomDate,
      serverUserEmail: (safeModel.user && safeModel.user.email) || '',
      serverDebugMessage: ROOMS_APP.normalizeString(safeModel.debugMessage || '')
    });
  } catch (error) {
    return HtmlService.createHtmlOutput(buildPlainRoomFallbackHtml_(params, error))
      .setTitle('ROOMS | Room');
  }
}

function buildPlainRoomFallbackHtml_(params, error) {
  var requested = normalizeRoomIdParam_(params) || '-';
  var message = String(error && error.message ? error.message : error);
  return [
    '<!doctype html>',
    '<html><head><meta charset="utf-8"><title>ROOMS | Room</title>',
    '<style>body{font-family:Arial,sans-serif;padding:16px;background:#020617;color:#e5e7eb}a{color:#60a5fa}</style>',
    '</head><body>',
    '<a href="?page=board">&larr; Torna alla board</a>',
    '<h1>ROOM PAGE</h1>',
    '<p><strong>Aula richiesta:</strong> ' + requested + '</p>',
    '<p><strong>Errore caricamento pagina aula</strong></p>',
    '<pre>' + message.replace(/[<>&]/g, function (char) {
      return { '<': '&lt;', '>': '&gt;', '&': '&amp;' }[char];
    }) + '</pre>',
    '</body></html>'
  ].join('');
}
