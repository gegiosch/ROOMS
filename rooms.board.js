var ROOMS_APP = ROOMS_APP || {};

ROOMS_APP.Board = {
  BRANCH_ORDER_: ['PA2', 'PA1', 'PA1_AULA', '2P_DX', '2P_SX', '1P_DX', '1P_SX', 'PT', 'LAB'],
  BRANCH_LABELS_: {
    PA2: 'PA2',
    PA1: 'PA1',
    PA1_AULA: 'AULA MAGNA',
    '2P_DX': 'DX',
    '2P_SX': 'SX',
    '1P_DX': 'DX',
    '1P_SX': 'SX',
    PT: 'PT',
    LAB: 'LABORATORI'
  },
  BRANCH_PAGE_CAPACITY_: 12,
  AULA_MAGNA_RESOURCE_ID_: 'AULA_MAGNA',

  normalizeMonitorUiScale_: function (value) {
    var parsed = Number(value);
    if (!isFinite(parsed) || parsed <= 0) {
      return 1;
    }
    return Math.max(0.5, Math.min(2, parsed));
  },

  getBoardViewModel: function () {
    var startedAt = Date.now();
    var stepStartedAt = startedAt;
    var user = ROOMS_APP.Auth.getUserContext();
    var runtimeContext = ROOMS_APP.RUNTIME_CONTEXT_ || {};
    var isMonitorMode = Boolean(runtimeContext.isMonitorMode);
    var simulation = ROOMS_APP.Auth.getSimulationContext_(null, user);
    var now = simulation.active && simulation.date ? simulation.date : new Date();
    var nowIso = simulation.active ? simulation.iso : ROOMS_APP.toIsoDateTime(now);
    var today = simulation.active ? simulation.dateIso : ROOMS_APP.toIsoDate(now);
    var currentTime = simulation.active ? simulation.time : Utilities.formatDate(now, ROOMS_APP.getTimezone(), 'HH:mm');
    var resources = this.listResourcesForBoard_();
    var activeResourceIds = {};
    resources.forEach(function (resource) {
      activeResourceIds[ROOMS_APP.normalizeString(resource.ResourceId)] = true;
    });
    var resourcesMs = Date.now() - stepStartedAt;

    stepStartedAt = Date.now();
    var bookings = ROOMS_APP.Booking.listBookingsForDate(today).map(function (booking) {
      return ROOMS_APP.Board.enrichUserBookingForBoard_(booking);
    });
    var bookingsMs = Date.now() - stepStartedAt;

    stepStartedAt = Date.now();
    var timetableOccupancies = ROOMS_APP.Timetable.listOccupanciesForDate('', today).filter(function (occupancy) {
      return activeResourceIds[ROOMS_APP.normalizeString(occupancy.ResourceId)] === true;
    });
    timetableOccupancies = this.uniqueByKey_(timetableOccupancies, 'BookingId');
    var timetableMs = Date.now() - stepStartedAt;

    stepStartedAt = Date.now();
    var aulaMagna = this.buildAulaMagnaEventBlock_(resources, now, simulation);
    var eventsMs = Date.now() - stepStartedAt;

    stepStartedAt = Date.now();
    var occupancies = ROOMS_APP.sortBy(
      bookings.concat(timetableOccupancies),
      ['ResourceId', 'StartTime', 'EndTime', 'SourceKind', 'BookingId']
    );
    var byResource = {};
    var branchBuckets = this.createEmptyBranchBuckets_();
    var visibleOccupancyKeys = {};
    var resourceNames = {};

    occupancies.forEach(function (occupancy) {
      byResource[occupancy.ResourceId] = byResource[occupancy.ResourceId] || [];
      byResource[occupancy.ResourceId].push(occupancy);
    });

    Object.keys(byResource).forEach(function (resourceId) {
      byResource[resourceId] = ROOMS_APP.sortBy(byResource[resourceId], ['StartTime', 'EndTime']);
    });

    resources.forEach(function (resource) {
      var roomBookings = byResource[resource.ResourceId] || [];
      resourceNames[resource.ResourceId] = resource.DisplayName || resource.ResourceId;
      var visibleCurrent = roomBookings.filter(function (booking) {
        return booking.StartTime <= currentTime && booking.EndTime > currentTime;
      })[0] || null;
      var visibleNext = roomBookings.filter(function (booking) {
        return booking.StartTime > currentTime;
      })[0] || null;
      var blockingBookings = roomBookings.filter(function (booking) {
        return ROOMS_APP.Timetable.isBlockingOccurrence(booking);
      });
      var current = blockingBookings.filter(function (booking) {
        return booking.StartTime <= currentTime && booking.EndTime > currentTime;
      })[0] || null;
      var next = blockingBookings.filter(function (booking) {
        return booking.StartTime > currentTime;
      })[0] || null;
      if (visibleCurrent) {
        visibleOccupancyKeys[ROOMS_APP.Board.getOccupancySummaryKey_(visibleCurrent)] = true;
      }
      if (visibleNext) {
        visibleOccupancyKeys[ROOMS_APP.Board.getOccupancySummaryKey_(visibleNext)] = true;
      }
      var state = current ? 'OCCUPIED' : (next ? 'NEXT_OCCUPIED' : 'FREE');
      var branchKey = ROOMS_APP.Board.mapResourceBranchKey_(resource);

      branchBuckets[branchKey].push({
        resourceId: resource.ResourceId,
        displayName: resource.DisplayName,
        sortKey: resource.SortKey || '',
        layoutRow: Number(resource.LayoutRow || 0),
        layoutCol: Number(resource.LayoutCol || 0),
        state: state,
        currentLabel: visibleCurrent ? ROOMS_APP.Board.getOccupancyLabel_(visibleCurrent) : 'LIBERA',
        nextLabel: visibleNext ? ROOMS_APP.Board.getNextOccupancyLabel_(visibleNext) : 'LIBERA'
      });
    });

    this.BRANCH_ORDER_.forEach(function (branchKey) {
      branchBuckets[branchKey] = ROOMS_APP.sortBy(branchBuckets[branchKey], ['sortKey', 'layoutRow', 'layoutCol', 'displayName']);
    });

    var configuredPageCount = Math.max(1, ROOMS_APP.getNumberConfig('BOARD_PAGE_COUNT', 1));
    var requiredPageCount = this.getRequiredPageCount_(branchBuckets);
    var pageCount = Math.max(1, Math.min(configuredPageCount, requiredPageCount));
    var pages = this.buildPages_(branchBuckets, pageCount);
    var afternoonBookings = this.buildAfternoonBookingSummary_(bookings, resourceNames, visibleOccupancyKeys, currentTime);
    var composeMs = Date.now() - stepStartedAt;

    var model = {
      generatedAtISO: nowIso,
      date: today,
      refreshSec: ROOMS_APP.getNumberConfig('BOARD_REFRESH_SEC', 60),
      rotationSec: ROOMS_APP.getNumberConfig('BOARD_ROTATION_SEC', 15),
      pageCount: pages.length,
      fullscreenCompactEnabled: ROOMS_APP.getBooleanConfig('BOARD_FULLSCREEN_COMPACT', true),
      isMonitorMode: isMonitorMode,
      monitorUiScale: this.normalizeMonitorUiScale_(ROOMS_APP.getConfigValue('MONITOR_UI_SCALE', '1')),
      palette: ROOMS_APP.PALETTE,
      appName: ROOMS_APP.getConfigValue('APP_NAME', 'ROOMS'),
      schoolName: ROOMS_APP.getConfigValue('SCHOOL_NAME', 'IIS Alessandrini'),
      user: {
        email: user.email,
        orgUnitPath: user.orgUnitPath || '',
        isAdmin: Boolean(user.canAccessAdmin),
        role: user.role || 'USER',
        isSuperAdmin: Boolean(user.isSuperAdmin),
        canBook: Boolean(user.canBook),
        canManageReplacement: Boolean(user.canManageReplacement),
        canManageAulaMagna: Boolean(user.canManageAulaMagna),
        canUseSimulation: Boolean(user.canUseSimulation),
        canAccessAdmin: Boolean(user.canAccessAdmin),
        simulationActive: Boolean(simulation.active),
        simulatedNowISO: simulation.iso || ''
      },
      aulaMagna: aulaMagna,
      afternoonBookings: afternoonBookings,
      branchOrder: this.BRANCH_ORDER_.slice(),
      branchLabels: this.BRANCH_LABELS_,
      pages: pages
    };
    Logger.log(
      '[PERF] Board.getBoardViewModel total=%sms resources=%sms bookings=%sms timetable=%sms events=%sms compose=%sms resourcesCount=%s bookingsCount=%s timetableCount=%s pages=%s',
      Date.now() - startedAt,
      resourcesMs,
      bookingsMs,
      timetableMs,
      eventsMs,
      composeMs,
      resources.length,
      bookings.length,
      timetableOccupancies.length,
      pages.length
    );
    return model;
  },

  buildAulaMagnaEventBlock_: function (resources, now, simulation) {
    var resource = this.findAulaMagnaResource_(resources);
    var simulationActive = Boolean(simulation && simulation.active);
    var nowDate = simulationActive ? simulation.dateIso : ROOMS_APP.toIsoDate(now || new Date());
    var currentTime = simulationActive ? simulation.time : Utilities.formatDate(now || new Date(), ROOMS_APP.getTimezone(), 'HH:mm');
    var horizonDays = Math.max(0, ROOMS_APP.getNumberConfig('AULA_MAGNA_EVENT_DAYS_AHEAD', 14));
    var endDate = this.addDaysToIsoDate_(nowDate, horizonDays);
    var self = this;

    if (!resource) {
      return {
        resourceId: this.AULA_MAGNA_RESOURCE_ID_,
        displayName: 'AULA MAGNA',
        currentEvent: null,
        nextEvents: []
      };
    }

    var events = ROOMS_APP.sortBy(
      ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.AULA_MAGNA_EVENTS).filter(function (row) {
        var eventDate = ROOMS_APP.toIsoDate(row.EventDate);
        var isActive = ROOMS_APP.normalizeString(row.IsActive) === '' ? true : ROOMS_APP.asBoolean(row.IsActive);
        if (!isActive || !eventDate) {
          return false;
        }
        if (eventDate < nowDate || eventDate > endDate) {
          return false;
        }
        return self.matchesAulaMagnaResource_(row.ResourceId, resource.ResourceId);
      }).map(function (row) {
        return {
          EventId: ROOMS_APP.normalizeString(row.EventId),
          EventDate: ROOMS_APP.toIsoDate(row.EventDate),
          StartTime: ROOMS_APP.toTimeString(row.StartTime),
          EndTime: ROOMS_APP.toTimeString(row.EndTime),
          EventName: ROOMS_APP.normalizeString(row.EventName),
          Notes: ROOMS_APP.normalizeString(row.Notes)
        };
      }),
      ['EventDate', 'StartTime', 'EndTime', 'EventName']
    );

    var currentEvent = events.filter(function (eventItem) {
      return eventItem.EventDate === nowDate &&
        eventItem.StartTime <= currentTime &&
        eventItem.EndTime > currentTime;
    })[0] || null;

    var nextEvents = events.filter(function (eventItem) {
      if (!currentEvent) {
        return eventItem.EventDate > nowDate ||
          (eventItem.EventDate === nowDate && eventItem.EndTime > currentTime);
      }
      return !ROOMS_APP.Board.isSameEvent_(eventItem, currentEvent) &&
        (eventItem.EventDate > nowDate ||
          (eventItem.EventDate === nowDate && eventItem.EndTime > currentTime));
    }).slice(0, 10);

    return {
      resourceId: resource.ResourceId,
      displayName: resource.DisplayName || 'AULA MAGNA',
      currentEvent: currentEvent,
      nextEvents: nextEvents
    };
  },

  findAulaMagnaResource_: function (resources) {
    return (resources || []).filter(function (resource) {
      var byId = ROOMS_APP.normalizeString(resource.ResourceId).toUpperCase();
      var byName = ROOMS_APP.slugify(resource.DisplayName || '');
      return byId === ROOMS_APP.Board.AULA_MAGNA_RESOURCE_ID_ || byName === ROOMS_APP.Board.AULA_MAGNA_RESOURCE_ID_;
    })[0] || null;
  },

  matchesAulaMagnaResource_: function (leftResourceId, rightResourceId) {
    var left = ROOMS_APP.normalizeString(leftResourceId);
    var right = ROOMS_APP.normalizeString(rightResourceId);
    if (!left || !right) {
      return false;
    }
    if (left.toUpperCase() === right.toUpperCase()) {
      return true;
    }
    return ROOMS_APP.slugify(left) === ROOMS_APP.slugify(right);
  },

  isSameEvent_: function (left, right) {
    if (!left || !right) {
      return false;
    }
    if (left.EventId && right.EventId) {
      return left.EventId === right.EventId;
    }
    return left.EventDate === right.EventDate &&
      left.StartTime === right.StartTime &&
      left.EndTime === right.EndTime &&
      left.EventName === right.EventName;
  },

  addDaysToIsoDate_: function (isoDate, daysToAdd) {
    var base = ROOMS_APP.combineDateTime(isoDate, '00:00');
    base.setDate(base.getDate() + Number(daysToAdd || 0));
    return ROOMS_APP.toIsoDate(base);
  },

  createEmptyBranchBuckets_: function () {
    var buckets = {};
    this.BRANCH_ORDER_.forEach(function (branchKey) {
      buckets[branchKey] = [];
    });
    return buckets;
  },

  getRequiredPageCount_: function (branchBuckets) {
    var maxBranchLength = 0;
    this.BRANCH_ORDER_.forEach(function (branchKey) {
      maxBranchLength = Math.max(maxBranchLength, branchBuckets[branchKey].length);
    });
    return Math.max(1, Math.ceil(maxBranchLength / this.BRANCH_PAGE_CAPACITY_));
  },

  buildPages_: function (branchBuckets, pageCount) {
    var pages = [];
    var index;
    for (index = 0; index < pageCount; index += 1) {
      var pageBranches = {};
      this.BRANCH_ORDER_.forEach(function (branchKey) {
        var rooms = branchBuckets[branchKey];
        var chunkSize = Math.max(1, Math.ceil(rooms.length / pageCount));
        pageBranches[branchKey] = rooms.slice(index * chunkSize, (index + 1) * chunkSize);
      });
      pages.push({
        pageId: String(index + 1),
        title: 'Pagina ' + String(index + 1),
        branches: pageBranches
      });
    }
    return pages;
  },

  getBookingSurname_: function (booking) {
    var explicitSurname = ROOMS_APP.normalizeString(booking.BookerSurname);
    if (explicitSurname) {
      return explicitSurname;
    }

    var name = ROOMS_APP.normalizeString(booking.BookerName);
    if (name) {
      var parts = name.split(/\s+/);
      return parts[parts.length - 1];
    }

    var email = ROOMS_APP.normalizeString(booking.BookerEmail);
    if (email && email.indexOf('@') > 0) {
      var localPart = email.split('@')[0];
      var emailParts = localPart.split(/[._-]+/);
      return emailParts[emailParts.length - 1].toUpperCase();
    }

    return 'N/D';
  },

  enrichUserBookingForBoard_: function (booking) {
    var enriched = {};
    Object.keys(booking || {}).forEach(function (key) {
      enriched[key] = booking[key];
    });
    enriched.SourceKind = 'USER_BOOKING';
    enriched.SourceType = 'USER_BOOKING';
    return enriched;
  },

  getOccupancySummaryKey_: function (occupancy) {
    return [
      ROOMS_APP.normalizeString(occupancy && occupancy.SourceKind),
      ROOMS_APP.normalizeString(occupancy && occupancy.BookingId),
      ROOMS_APP.normalizeString(occupancy && occupancy.ResourceId),
      ROOMS_APP.toIsoDate(occupancy && occupancy.BookingDate),
      ROOMS_APP.toTimeString(occupancy && occupancy.StartTime),
      ROOMS_APP.toTimeString(occupancy && occupancy.EndTime)
    ].join('|');
  },

  getBookingActivityDescription_: function (booking) {
    return ROOMS_APP.normalizeString(
      (booking && booking.ActivityDescription) ||
      (booking && booking.activityDescription) ||
      (booking && booking.Notes) ||
      (booking && booking.Title) ||
      'Attivita non specificata'
    );
  },

  getBookingDisplayMode_: function (booking) {
    return String(
      booking && (booking.DisplayMode || booking.displayMode) || 'TEACHER'
    ).toUpperCase() === 'ACTIVITY' ? 'ACTIVITY' : 'TEACHER';
  },

  getBookingActorLabel_: function (booking) {
    var surname = ROOMS_APP.normalizeString(booking && booking.BookerSurname);
    var name = ROOMS_APP.normalizeString(booking && booking.BookerName);
    var joined = [surname, name].filter(function (token) {
      return Boolean(token);
    }).join(' ');
    return joined || ROOMS_APP.normalizeString(booking && booking.BookerEmail) || 'N/D';
  },

  buildAfternoonBookingSummary_: function (bookings, resourceNames, visibleOccupancyKeys, currentTime) {
    return ROOMS_APP.sortBy((bookings || []).filter(function (booking) {
      if (!booking || booking.Status === 'CANCELLED') {
        return false;
      }
      if (ROOMS_APP.toTimeString(booking.StartTime) < '14:00') {
        return false;
      }
      if (currentTime && ROOMS_APP.toTimeString(booking.EndTime) <= currentTime) {
        return false;
      }
      return !visibleOccupancyKeys[ROOMS_APP.Board.getOccupancySummaryKey_(booking)];
    }).map(function (booking) {
      return {
        bookingId: ROOMS_APP.normalizeString(booking.BookingId),
        resourceId: ROOMS_APP.normalizeString(booking.ResourceId),
        roomName: resourceNames[booking.ResourceId] || booking.ResourceId || '',
        bookingDate: ROOMS_APP.toIsoDate(booking.BookingDate),
        startTime: ROOMS_APP.toTimeString(booking.StartTime),
        endTime: ROOMS_APP.toTimeString(booking.EndTime),
        activityDescription: ROOMS_APP.Board.getBookingActivityDescription_(booking),
        actorLabel: ROOMS_APP.Board.getBookingActorLabel_(booking)
      };
    }), ['startTime', 'roomName', 'endTime', 'bookingId']);
  },

  getOccupancyLabel_: function (occupancy) {
    if (!occupancy) {
      return 'N/D';
    }
    if (ROOMS_APP.normalizeString(occupancy.SourceKind) === 'TIMETABLE') {
      return ROOMS_APP.normalizeString(
        occupancy.TeacherName ||
        occupancy.BookerName ||
        ROOMS_APP.Timetable.getDisplayLabel(occupancy)
      );
    }
    if (this.getBookingDisplayMode_(occupancy) === 'ACTIVITY') {
      return this.getBookingActivityDescription_(occupancy);
    }
    return ROOMS_APP.normalizeString(occupancy && occupancy.DisplayLabel) || this.getBookingSurname_(occupancy);
  },

  getNextOccupancyLabel_: function (occupancy) {
    if (!occupancy) {
      return 'LIBERA';
    }
    var startTime = ROOMS_APP.toTimeString(occupancy.StartTime || occupancy.startTime || '');
    var label = this.getOccupancyLabel_(occupancy);
    return startTime ? (startTime + ': ' + label) : label;
  },

  uniqueByKey_: function (rows, keyField) {
    var seen = {};
    return (rows || []).filter(function (row) {
      var key = ROOMS_APP.normalizeString(row && row[keyField]);
      if (!key) {
        return false;
      }
      if (seen[key]) {
        return false;
      }
      seen[key] = true;
      return true;
    });
  },

  mapResourceBranchKey_: function (resource) {
    var areaCode = this.normalizeKey_(resource.AreaCode);
    var floorCode = this.normalizeKey_(resource.FloorCode);
    var sideCode = this.normalizeKey_(resource.SideCode);
    var areaLabel = this.normalizeKey_(resource.AreaLabel);
    var floorLabel = this.normalizeKey_(resource.FloorLabel);
    var sideLabel = this.normalizeKey_(resource.SideLabel);
    var resourceId = this.normalizeKey_(resource.ResourceId);

    if (resourceId === 'AULA_MAGNA') {
      return 'PA1_AULA';
    }

    if (areaCode === 'LAB' || areaLabel.indexOf('LABORATORI') >= 0 || sideLabel.indexOf('LABORATORI') >= 0) {
      return 'LAB';
    }

    if (areaCode === 'F2M' || (floorCode === 'F2' && sideCode === 'MEZZATO') || floorLabel.indexOf('SECONDO') >= 0 && floorLabel.indexOf('MEZZATO') >= 0) {
      return 'PA2';
    }

    if (areaCode === 'F1M' || (floorCode === 'F1' && sideCode === 'MEZZATO') || floorLabel.indexOf('PRIMO') >= 0 && floorLabel.indexOf('MEZZATO') >= 0) {
      return 'PA1';
    }

    if (areaCode === 'F2R' || (floorCode === 'F2' && (sideCode === 'RIGHT' || sideLabel.indexOf('DESTRO') >= 0))) {
      return '2P_DX';
    }

    if (areaCode === 'F2L' || (floorCode === 'F2' && (sideCode === 'LEFT' || sideLabel.indexOf('SINISTRO') >= 0))) {
      return '2P_SX';
    }

    if (areaCode === 'F1R' || (floorCode === 'F1' && (sideCode === 'RIGHT' || sideLabel.indexOf('DESTRO') >= 0))) {
      return '1P_DX';
    }

    if (areaCode === 'F1L' || (floorCode === 'F1' && (sideCode === 'LEFT' || sideLabel.indexOf('SINISTRO') >= 0))) {
      return '1P_SX';
    }

    if (areaCode === 'F0' || areaCode === 'GYM' || floorCode === 'F0' || floorLabel.indexOf('TERRA') >= 0) {
      return 'PT';
    }

    return 'PT';
  },

  normalizeKey_: function (value) {
    return ROOMS_APP.normalizeString(value).toUpperCase();
  },

  listResourcesForBoard_: function () {
    return ROOMS_APP.sortBy(
      ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.RESOURCES).filter(function (row) {
        return ROOMS_APP.asBoolean(row.IsActive);
      }),
      ['LayoutPage', 'LayoutRow', 'LayoutCol', 'DisplayName']
    );
  },

  listResources_: function () {
    return this.listResourcesForBoard_();
  }
};
