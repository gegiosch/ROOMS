var ROOMS_APP = ROOMS_APP || {};

function doGet(e) {
  var startedAt = Date.now();
  var params = (e && e.parameter) || {};
  var page = params.page || 'board';
  var isMonitorMode = isMonitorModeRequest_(params);

  try {
    var response;
    if (page === 'api') {
      response = jsonResponse_(withRuntimeContext_(extractRuntimeContext_(params), function () {
        return routeApiRequest_(params);
      }));
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

    if (page === 'admin' && !isMonitorMode) {
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
    var response = jsonResponse_(withRuntimeContext_(extractRuntimeContext_(body), function () {
      return routeApiRequest_(body);
    }));
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

function getRedirectTargetForHost(host) {
  return ROOMS_APP.buildRedirectTargetForHost(host);
}

function runSetup() {
  var actor = ROOMS_APP.Auth.requireAdmin();
  ROOMS_APP.Schema.ensureAll();
  ROOMS_APP.Auth.bumpUserContextCacheVersion_();
  ROOMS_APP.invalidateConfigCache();
  return {
    ok: true,
    message: 'Setup completato.',
    actorEmail: actor.email,
    executedAtISO: ROOMS_APP.toIsoDateTime(new Date())
  };
}

function getBoardViewModel(requestContext) {
  var startedAt = Date.now();
  return withRuntimeContext_(extractRuntimeContextFromArgs_(arguments), function () {
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
  });
}

function getRoomViewModel(resourceId, dateString, requestContext) {
  return withRuntimeContext_(extractRuntimeContext_(requestContext), function () {
    return ROOMS_APP.Booking.getRoomViewModel(ROOMS_APP.normalizeString(resourceId), dateString);
  });
}

function getRoomPanelData(resourceId, dateString, panelOptions, requestContext) {
  var startedAt = Date.now();
  var normalizedResourceId = ROOMS_APP.normalizeString(resourceId);
  var options = panelOptions && typeof panelOptions === 'object' ? panelOptions : {};
  var runtimeContext = requestContext;
  if (!runtimeContext && options && (
    options.simulatedNow ||
    options.simulatedDateTime ||
    options.__simulatedNow
  )) {
    runtimeContext = options;
  }
  var slotOptions = {
    splitFreeSlotsHalfHour: Boolean(
      options && (
        options.splitFreeSlotsHalfHour === true ||
        options.splitFreeSlotsHalfHour === 'true'
      )
    )
  };
  return withRuntimeContext_(extractRuntimeContext_(runtimeContext), function () {
    var simulation = ROOMS_APP.Auth.getSimulationContext_();
    var targetDate = dateString || (simulation.active ? simulation.dateIso : ROOMS_APP.toIsoDate(ROOMS_APP.Auth.getEffectiveNow()));
    try {
      var model = ROOMS_APP.Booking.getRoomViewModel(normalizedResourceId, targetDate, slotOptions);
      logTiming_('getRoomPanelData', startedAt, {
        resourceId: normalizedResourceId,
        date: targetDate,
        splitFree: slotOptions.splitFreeSlotsHalfHour ? 'TRUE' : 'FALSE',
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
  });
}

function getAulaMagnaEditorModel(resourceId, dateString, requestContext) {
  return withRuntimeContext_(extractRuntimeContext_(requestContext), function () {
    return ROOMS_APP.Booking.getAulaMagnaEditorModel(resourceId, dateString);
  });
}

function applyAulaMagnaEventChanges(resourceId, changes, requestContext) {
  return withRuntimeContext_(extractRuntimeContext_(requestContext), function () {
    return ROOMS_APP.Booking.applyAulaMagnaEventChanges(resourceId, changes || {});
  });
}

function createRoomBooking(payload, requestContext) {
  return withRuntimeContext_(extractRuntimeContext_(requestContext), function () {
    return ROOMS_APP.Booking.createBooking(payload);
  });
}

function cancelRoomBooking(bookingId, notes, requestContext) {
  return withRuntimeContext_(extractRuntimeContext_(requestContext), function () {
    return ROOMS_APP.Booking.cancelBooking(bookingId, notes);
  });
}

function updateRoomBooking(bookingId, payload, requestContext) {
  return withRuntimeContext_(extractRuntimeContext_(requestContext), function () {
    return ROOMS_APP.Booking.updateBooking(bookingId, payload || {});
  });
}

function applyRoomPanelChanges(resourceId, dateString, changes, requestContext) {
  return withRuntimeContext_(extractRuntimeContext_(requestContext), function () {
    var simulation = ROOMS_APP.Auth.getSimulationContext_();
    return ROOMS_APP.Booking.applyRoomChanges(
      ROOMS_APP.normalizeString(resourceId),
      dateString || (simulation.active ? simulation.dateIso : ROOMS_APP.toIsoDate(ROOMS_APP.Auth.getEffectiveNow())),
      changes || {}
    );
  });
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
  ROOMS_APP.Schema.ensureAdmins();
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

  if (tableName === ROOMS_APP.SHEET_NAMES.ADMINS) {
    ROOMS_APP.Schema.ensureAdmins();
  }
  ROOMS_APP.DB.replaceRows(tableName, ROOMS_APP.DB.getHeaders(tableName), rows || []);
  if (tableName === ROOMS_APP.SHEET_NAMES.CONFIG) {
    ROOMS_APP.invalidateConfigCache();
  }
  if (tableName === ROOMS_APP.SHEET_NAMES.ADMINS) {
    ROOMS_APP.Auth.bumpUserContextCacheVersion_();
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
    return getBoardViewModel(payload);
  }
  if (action === 'room') {
    return getRoomPanelData(normalizeRoomIdParam_(payload), payload.date, payload, payload);
  }
  if (action === 'roomPanel') {
    return getRoomPanelData(normalizeRoomIdParam_(payload), payload.date, payload, payload);
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
  if (action === 'getAulaMagnaEditorModel') {
    return ROOMS_APP.Booking.getAulaMagnaEditorModel(
      normalizeRoomIdParam_(payload),
      payload.date
    );
  }
  if (action === 'applyAulaMagnaEventChanges') {
    return ROOMS_APP.Booking.applyAulaMagnaEventChanges(
      normalizeRoomIdParam_(payload),
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
  var isMonitorMode = isMonitorModeRequest_(params);
  var bootstrapModel = {
    generatedAtISO: '',
    date: '',
    refreshSec: 60,
    rotationSec: 15,
    pageCount: 1,
    fullscreenCompactEnabled: true,
    isMonitorMode: isMonitorMode,
    monitorUiScale: Math.max(0.5, ROOMS_APP.getNumberConfig('MONITOR_UI_SCALE', 1) || 1),
    appName: 'ROOMS',
    schoolName: '',
    user: {
      email: '',
      orgUnitPath: '',
      isAdmin: false,
      role: 'USER',
      isSuperAdmin: false,
      canBook: false,
      canManageReplacement: false,
      canManageAulaMagna: false,
      canUseSimulation: false,
      canAccessAdmin: false,
      simulationActive: false,
      simulatedNowISO: ''
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

function extractRuntimeContext_(payload) {
  if (!payload || typeof payload !== 'object') {
    return {};
  }

  return {
    mode: ROOMS_APP.normalizeString(payload.mode || ''),
    isMonitorMode: isMonitorModeRequest_(payload),
    simulatedNow: ROOMS_APP.normalizeString(
      payload.simulatedNow ||
      payload.simulatedDateTime ||
      payload.__simulatedNow ||
      ''
    )
  };
}

function withRuntimeContext_(requestContext, callback) {
  var previousContext = ROOMS_APP.RUNTIME_CONTEXT_ || null;
  var actor = ROOMS_APP.Auth.getUserContext();
  var simulation = ROOMS_APP.Auth.getSimulationContext_(requestContext, actor);
  ROOMS_APP.DB.beginRequestCache_();
  ROOMS_APP.RUNTIME_CONTEXT_ = {
    mode: requestContext && requestContext.mode ? requestContext.mode : '',
    isMonitorMode: Boolean(requestContext && requestContext.isMonitorMode),
    simulatedNow: simulation.active ? simulation.iso : '',
    actorEmail: actor.email,
    role: actor.role || 'USER',
    isSuperAdmin: Boolean(actor.isSuperAdmin),
    canUseSimulation: Boolean(actor.canUseSimulation),
    canAccessAdmin: Boolean(actor.canAccessAdmin)
  };

  try {
    return callback();
  } finally {
    ROOMS_APP.RUNTIME_CONTEXT_ = previousContext;
    ROOMS_APP.DB.endRequestCache_();
  }
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

function isMonitorModeRequest_(payload) {
  return ROOMS_APP.normalizeString(payload && payload.mode).toLowerCase() === 'monitor';
}

function extractRuntimeContextFromArgs_(argsLike) {
  var merged = {};
  Array.prototype.slice.call(argsLike || []).forEach(function (entry) {
    if (!entry || typeof entry !== 'object') {
      return;
    }
    Object.keys(entry).forEach(function (key) {
      merged[key] = entry[key];
    });
  });
  return extractRuntimeContext_(merged);
}
