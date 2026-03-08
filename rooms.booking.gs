var ROOMS_APP = ROOMS_APP || {};

ROOMS_APP.Booking = {
  listBookingsForDay: function (resourceId, dateString) {
    ROOMS_APP.Schema.ensureAll();
    var rows = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.BOOKINGS).filter(function (row) {
      return row.ResourceId === resourceId &&
        row.BookingDate === dateString &&
        row.Status !== 'CANCELLED';
    });

    return ROOMS_APP.sortBy(rows, ['StartTime', 'EndTime', 'CreatedAtISO']);
  },

  listBookingsForDate: function (dateString) {
    ROOMS_APP.Schema.ensureAll();
    return ROOMS_APP.sortBy(
      ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.BOOKINGS).filter(function (row) {
        return row.BookingDate === dateString && row.Status !== 'CANCELLED';
      }),
      ['ResourceId', 'StartTime', 'EndTime']
    );
  },

  createBooking: function (payload, options) {
    var actor = (options && options.actor) || ROOMS_APP.Auth.getUserContext();
    var validation = ROOMS_APP.Policy.validateBookingRequest(payload, actor);
    if (!validation.ok) {
      this.writeAudit_('CREATE_BOOKING', payload.bookingId || '', payload.seriesId || '', payload.resourceId || '', actor.email, 'ERROR', {
        errors: validation.errors,
        payload: payload
      });
      throw new Error(validation.errors.join(' '));
    }

    var normalized = validation.normalized;
    var bookingId = normalized.bookingId || Utilities.getUuid();
    var seriesId = normalized.seriesId || '';
    var nowIso = ROOMS_APP.toIsoDateTime(new Date());
    var booking = {
      BookingId: bookingId,
      SeriesId: seriesId,
      ResourceId: normalized.resourceId,
      BookingDate: normalized.bookingDate,
      StartTime: normalized.startTime,
      EndTime: normalized.endTime,
      StartISO: ROOMS_APP.toIsoDateTime(ROOMS_APP.combineDateTime(normalized.bookingDate, normalized.startTime)),
      EndISO: ROOMS_APP.toIsoDateTime(ROOMS_APP.combineDateTime(normalized.bookingDate, normalized.endTime)),
      Title: normalized.title || validation.resource.DisplayName,
      BookerEmail: actor.email,
      BookerName: normalized.bookerName,
      BookerSurname: normalized.bookerSurname,
      Status: 'CONFIRMED',
      CreatedAtISO: nowIso,
      UpdatedAtISO: nowIso,
      CancelledAtISO: '',
      Notes: normalized.notes
    };

    ROOMS_APP.DB.appendRows(ROOMS_APP.SHEET_NAMES.BOOKINGS, [booking]);
    this.writeAudit_('CREATE_BOOKING', bookingId, seriesId, booking.ResourceId, actor.email, 'OK', booking);
    return booking;
  },

  cancelBooking: function (bookingId, notes) {
    var actor = ROOMS_APP.Auth.getUserContext();
    var rows = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.BOOKINGS);
    var booking = rows.filter(function (row) {
      return row.BookingId === bookingId;
    })[0];

    if (!booking) {
      throw new Error('Booking not found.');
    }

    if (!actor.isAdmin && booking.BookerEmail !== actor.email) {
      throw new Error('Only the creator or an admin can cancel this booking.');
    }

    booking.Status = 'CANCELLED';
    booking.UpdatedAtISO = ROOMS_APP.toIsoDateTime(new Date());
    booking.CancelledAtISO = booking.UpdatedAtISO;
    booking.Notes = ROOMS_APP.normalizeString(notes || booking.Notes);

    ROOMS_APP.DB.upsertByKey(ROOMS_APP.SHEET_NAMES.BOOKINGS, 'BookingId', [booking]);
    this.writeAudit_('CANCEL_BOOKING', booking.BookingId, booking.SeriesId, booking.ResourceId, actor.email, 'OK', booking);
    return booking;
  },

  buildRoomFallbackModel_: function (requestedResourceId, dateString, errorMessage) {
    return {
      ok: false,
      errorMessage: errorMessage || 'Dati aula non disponibili',
      date: dateString || ROOMS_APP.toIsoDate(new Date()),
      requestedResourceId: ROOMS_APP.normalizeString(requestedResourceId),
      resource: null,
      resources: [],
      bookings: [],
      ownBookings: [],
      freeSlots: [],
      slots: [],
      isOpen: false,
      status: 'UNKNOWN',
      currentBooking: null,
      user: ROOMS_APP.Auth.getUserContext(),
      config: this.getRoomConfig_()
    };
  },

  getRoomViewModel: function (resourceId, dateString) {
    var date = dateString || ROOMS_APP.toIsoDate(new Date());
    var requestedResourceId = ROOMS_APP.normalizeString(resourceId || '');
    var user = ROOMS_APP.Auth.getUserContext();

    try {
      var allResources = ROOMS_APP.Board.listResources_();
      if (!allResources.length) {
        var emptyModel = this.buildRoomFallbackModel_(requestedResourceId, date, 'Dati aula non disponibili');
        emptyModel.user = user;
        return emptyModel;
      }

      var resource = requestedResourceId
        ? this.findResourceByAnyKey_(allResources, requestedResourceId)
        : allResources[0];
      if (requestedResourceId && !resource) {
        var notFoundModel = this.buildRoomFallbackModel_(requestedResourceId, date, 'Aula non trovata');
        notFoundModel.user = user;
        notFoundModel.resources = allResources;
        return notFoundModel;
      }

      var selectedResourceId = resource ? resource.ResourceId : '';
      var bookings = resource ? this.listBookingsForDay(selectedResourceId, date) : [];
      var timeline = resource ? ROOMS_APP.Slots.getDaySlots(selectedResourceId, date) : { bookings: [], freeSlots: [], slots: [], isOpen: false };
      var now = new Date();
      var currentTime = Utilities.formatDate(now, ROOMS_APP.getTimezone(), 'HH:mm');
      var current = bookings.filter(function (booking) {
        return booking.StartTime <= currentTime && booking.EndTime > currentTime;
      })[0] || null;
      var ownBookings = bookings.filter(function (booking) {
        return booking.BookerEmail === user.email;
      });

      return {
        ok: true,
        errorMessage: '',
        date: date,
        requestedResourceId: requestedResourceId,
        resource: resource,
        bookings: bookings,
        ownBookings: ownBookings,
        freeSlots: timeline.freeSlots || [],
        slots: timeline.slots || [],
        isOpen: timeline.isOpen,
        status: current ? 'OCCUPIED' : 'FREE',
        currentBooking: current,
        user: user,
        resources: allResources,
        config: this.getRoomConfig_()
      };
    } catch (error) {
      var fallback = this.buildRoomFallbackModel_(requestedResourceId, date, 'Errore caricamento pagina aula');
      fallback.user = user;
      fallback.debugMessage = String(error && error.message ? error.message : error);
      return fallback;
    }
  },

  findResourceByAnyKey_: function (resources, requestedResourceId) {
    var normalized = ROOMS_APP.normalizeString(requestedResourceId).toUpperCase();
    var requestedSlug = ROOMS_APP.slugify(requestedResourceId);

    return resources.filter(function (resource) {
      var resourceId = ROOMS_APP.normalizeString(resource.ResourceId).toUpperCase();
      var displayName = ROOMS_APP.normalizeString(resource.DisplayName);
      return resourceId === normalized ||
        displayName.toUpperCase() === normalized ||
        ROOMS_APP.slugify(displayName) === requestedSlug;
    })[0] || null;
  },

  getRoomConfig_: function () {
    return {
      bookingEnabled: ROOMS_APP.getBooleanConfig('BOOKING_ENABLED', true),
      allowRecurring: ROOMS_APP.getBooleanConfig('ALLOW_RECURRING', true),
      showBookerName: ROOMS_APP.getBooleanConfig('SHOW_BOOKER_NAME', false)
    };
  },

  writeAudit_: function (action, bookingId, seriesId, resourceId, actorEmail, result, payload) {
    ROOMS_APP.DB.appendRows(ROOMS_APP.SHEET_NAMES.AUDIT, [{
      TimestampISO: ROOMS_APP.toIsoDateTime(new Date()),
      Action: action,
      BookingId: bookingId || '',
      SeriesId: seriesId || '',
      ResourceId: resourceId || '',
      ActorEmail: actorEmail || '',
      Result: result,
      PayloadJson: JSON.stringify(payload || {})
    }]);
  }
};
