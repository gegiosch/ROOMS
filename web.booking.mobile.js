var ROOMS_APP = ROOMS_APP || {};

ROOMS_APP.BookingMobile = {
  getBootstrapModel: function () {
    var user = ROOMS_APP.Auth.requireCanBook();
    ROOMS_APP.Auth.assertAllowedDomain(user.email);
    return {
      ok: true,
      user: this.toPublicUser_(user),
      defaultDate: this.getToday_(),
      resources: this.listBookableResources_()
    };
  },

  getAvailabilityModel: function (dateString, resourceId, options) {
    var user = ROOMS_APP.Auth.requireCanBook();
    ROOMS_APP.Auth.assertAllowedDomain(user.email);
    var date = ROOMS_APP.toIsoDate(dateString || this.getToday_());
    var roomId = ROOMS_APP.normalizeString(resourceId);
    var viewOptions = options && typeof options === 'object' ? options : {};
    if (!roomId) {
      throw new Error('Selezionare un\'aula.');
    }
    var model = ROOMS_APP.Booking.getRoomViewModel(roomId, date, {
      splitFreeSlotsHalfHour: Boolean(viewOptions.splitFreeSlotsHalfHour)
    });
    var resource = model && model.resource ? model.resource : null;
    return {
      ok: Boolean(model && model.ok),
      date: model && model.date ? model.date : date,
      resourceId: resource && resource.ResourceId ? resource.ResourceId : roomId,
      roomName: resource ? (resource.DisplayName || resource.ResourceId) : roomId,
      isOpen: Boolean(model && model.isOpen),
      statusSummary: model && model.statusSummary ? model.statusSummary : '',
      freeSlots: model && model.freeSlots ? model.freeSlots : [],
      bookings: model && model.bookings ? model.bookings : [],
      options: {
        splitFreeSlotsHalfHour: Boolean(viewOptions.splitFreeSlotsHalfHour)
      },
      errorMessage: model && model.errorMessage ? model.errorMessage : ''
    };
  },

  listMyBookings: function () {
    var user = ROOMS_APP.Auth.requireCanBook();
    ROOMS_APP.Auth.assertAllowedDomain(user.email);
    var today = this.getToday_();
    var resourceMap = this.buildResourceMap_();
    var normalizedEmail = ROOMS_APP.normalizeEmail(user.email);
    return ROOMS_APP.sortBy(
      ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.BOOKINGS).filter(function (row) {
        return row &&
          row.Status !== 'CANCELLED' &&
          row.BookingDate >= today &&
          ROOMS_APP.normalizeEmail(row.BookerEmail) === normalizedEmail;
      }),
      ['BookingDate', 'StartTime', 'EndTime']
    ).map(function (row) {
      return ROOMS_APP.BookingMobile.toMobileBooking_(row, resourceMap);
    });
  },

  createBookings: function (draftRows) {
    var user = ROOMS_APP.Auth.requireCanBook();
    ROOMS_APP.Auth.assertAllowedDomain(user.email);
    var rows = Array.isArray(draftRows) ? draftRows : [];
    var resourceMap = this.buildResourceMap_();
    var result = {
      savedCount: 0,
      failedCount: 0,
      saved: [],
      failed: []
    };

    rows.forEach(function (row, index) {
      var draftId = ROOMS_APP.normalizeString(row && row.draftId) || ('slot-' + String(index + 1));
      try {
        var payload = ROOMS_APP.BookingMobile.normalizeMobilePayload_(row, user, resourceMap);
        var booking = ROOMS_APP.Booking.createBooking(payload, { actor: user });
        result.savedCount += 1;
        result.saved.push({
          draftId: draftId,
          booking: ROOMS_APP.BookingMobile.toMobileBooking_(booking, resourceMap)
        });
      } catch (error) {
        result.failedCount += 1;
        result.failed.push({
          draftId: draftId,
          message: String(error && error.message ? error.message : error)
        });
      }
    });

    return result;
  },

  updateBooking: function (bookingId, payload) {
    var user = ROOMS_APP.Auth.requireCanBook();
    ROOMS_APP.Auth.assertAllowedDomain(user.email);
    var resourceMap = this.buildResourceMap_();
    var normalized = this.normalizeMobilePayload_(payload || {}, user, resourceMap);
    var updated = ROOMS_APP.Booking.updateBooking(bookingId, normalized);
    return this.toMobileBooking_(updated, resourceMap);
  },

  cancelBooking: function (bookingId, notes) {
    var user = ROOMS_APP.Auth.requireCanBook();
    ROOMS_APP.Auth.assertAllowedDomain(user.email);
    var resourceMap = this.buildResourceMap_();
    var cancelled = ROOMS_APP.Booking.cancelBooking(bookingId, notes || 'Cancellata da mobile.');
    return this.toMobileBooking_(cancelled, resourceMap);
  },

  normalizeMobilePayload_: function (row, user, resourceMap) {
    var resourceId = ROOMS_APP.normalizeString(row && (row.resourceId || row.ResourceId));
    var resource = resourceMap[resourceId] || {};
    var activityDescription = ROOMS_APP.normalizeString(row && (row.activityDescription || row.ActivityDescription));
    if (!activityDescription) {
      throw new Error('Inserire la descrizione attività.');
    }
    return {
      bookingId: ROOMS_APP.normalizeString(row && (row.bookingId || row.BookingId)),
      seriesId: ROOMS_APP.normalizeString(row && (row.seriesId || row.SeriesId)),
      resourceId: resourceId,
      bookingDate: ROOMS_APP.toIsoDate(row && (row.bookingDate || row.date || row.BookingDate)),
      startTime: ROOMS_APP.toTimeString(row && (row.startTime || row.StartTime)),
      endTime: ROOMS_APP.toTimeString(row && (row.endTime || row.EndTime)),
      title: ROOMS_APP.normalizeString(row && (row.title || row.Title)) || resource.DisplayName || resourceId,
      activityDescription: activityDescription,
      displayMode: ROOMS_APP.Booking.normalizeDisplayMode_(row && (row.displayMode || row.DisplayMode)),
      notes: ROOMS_APP.normalizeString(row && (row.notes || row.Notes)),
      bookerName: ROOMS_APP.normalizeString(row && (row.bookerName || row.BookerName)) || user.firstName,
      bookerSurname: ROOMS_APP.normalizeString(row && (row.bookerSurname || row.BookerSurname)) || user.surname
    };
  },

  toMobileBooking_: function (booking, resourceMap) {
    var resource = resourceMap && booking ? resourceMap[booking.ResourceId] : null;
    var displayMode = ROOMS_APP.Booking.normalizeDisplayMode_(booking && booking.DisplayMode);
    return {
      bookingId: ROOMS_APP.normalizeString(booking && booking.BookingId),
      resourceId: ROOMS_APP.normalizeString(booking && booking.ResourceId),
      roomName: resource ? (resource.DisplayName || resource.ResourceId) : ROOMS_APP.normalizeString(booking && booking.ResourceId),
      bookingDate: ROOMS_APP.toIsoDate(booking && booking.BookingDate),
      startTime: ROOMS_APP.toTimeString(booking && booking.StartTime),
      endTime: ROOMS_APP.toTimeString(booking && booking.EndTime),
      title: ROOMS_APP.normalizeString(booking && booking.Title),
      activityDescription: ROOMS_APP.normalizeString(booking && booking.ActivityDescription),
      displayMode: displayMode,
      displayModeLabel: displayMode === 'ACTIVITY' ? 'Attività' : 'Docente',
      displayLabel: ROOMS_APP.Booking.getBookingDisplayLabel_(booking),
      status: ROOMS_APP.normalizeString(booking && booking.Status),
      notes: ROOMS_APP.normalizeString(booking && booking.Notes),
      canManage: true
    };
  },

  toPublicUser_: function (user) {
    return {
      email: user.email || '',
      displayName: user.displayName || [user.firstName || '', user.surname || ''].join(' ').trim(),
      firstName: user.firstName || '',
      surname: user.surname || '',
      canBook: Boolean(user.canBook)
    };
  },

  listBookableResources_: function () {
    return ROOMS_APP.Board.listResources_().filter(function (resource) {
      return resource &&
        ROOMS_APP.asBoolean(resource.IsActive) &&
        ROOMS_APP.asBoolean(resource.IsBookable) &&
        ROOMS_APP.slugify(resource.DisplayName || '') !== 'aula_magna';
    }).map(function (resource) {
      return {
        resourceId: resource.ResourceId,
        displayName: resource.DisplayName || resource.ResourceId
      };
    });
  },

  buildResourceMap_: function () {
    var output = {};
    ROOMS_APP.Board.listResources_().forEach(function (resource) {
      if (resource && resource.ResourceId) {
        output[resource.ResourceId] = resource;
      }
    });
    return output;
  },

  getToday_: function () {
    var user = ROOMS_APP.Auth.getUserContext();
    var simulation = ROOMS_APP.Auth.getSimulationContext_(null, user);
    return simulation.active ? simulation.dateIso : ROOMS_APP.toIsoDate(ROOMS_APP.Auth.getEffectiveNow(null, user));
  }
};

function renderMobileBookingPage_(params) {
  return renderTemplate_('ui.booking.mobile', {
    pageTitle: 'Prenotazione aule',
    initialModelJson: JSON.stringify({
      route: 'booking-mobile',
      generatedAtISO: ROOMS_APP.toIsoDateTime(new Date())
    })
  });
}

function getMobileBookingBootstrap(requestContext) {
  return withRuntimeContext_(extractRuntimeContext_(requestContext), function () {
    return ROOMS_APP.BookingMobile.getBootstrapModel();
  });
}

function getMobileBookingAvailability(dateString, resourceId, options, requestContext) {
  return withRuntimeContext_(extractRuntimeContextFromArgs_(arguments), function () {
    return ROOMS_APP.BookingMobile.getAvailabilityModel(dateString, resourceId, options || {});
  });
}

function listMobileMyBookings(requestContext) {
  return withRuntimeContext_(extractRuntimeContext_(requestContext), function () {
    return ROOMS_APP.BookingMobile.listMyBookings();
  });
}

function saveMobileBookings(rows, requestContext) {
  return withRuntimeContext_(extractRuntimeContextFromArgs_(arguments), function () {
    return ROOMS_APP.BookingMobile.createBookings(rows || []);
  });
}

function updateMobileBooking(bookingId, payload, requestContext) {
  return withRuntimeContext_(extractRuntimeContextFromArgs_(arguments), function () {
    return ROOMS_APP.BookingMobile.updateBooking(bookingId, payload || {});
  });
}

function cancelMobileBooking(bookingId, notes, requestContext) {
  return withRuntimeContext_(extractRuntimeContextFromArgs_(arguments), function () {
    return ROOMS_APP.BookingMobile.cancelBooking(bookingId, notes || '');
  });
}
