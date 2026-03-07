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

  getRoomViewModel: function (resourceId, dateString) {
    var date = dateString || ROOMS_APP.toIsoDate(new Date());
    var allResources = ROOMS_APP.Board.listResources_();
    var selectedResourceId = resourceId || (allResources[0] ? allResources[0].ResourceId : '');
    var resource = ROOMS_APP.Policy.getResource(selectedResourceId);
    var bookings = resource ? this.listBookingsForDay(selectedResourceId, date) : [];
    var timeline = resource ? ROOMS_APP.Slots.getDaySlots(selectedResourceId, date) : { bookings: [], freeSlots: [], isOpen: false };
    var now = new Date();
    var currentTime = Utilities.formatDate(now, ROOMS_APP.getTimezone(), 'HH:mm');
    var current = bookings.filter(function (booking) {
      return booking.StartTime <= currentTime && booking.EndTime > currentTime;
    })[0] || null;

    return {
      date: date,
      resource: resource,
      bookings: bookings,
      freeSlots: timeline.freeSlots,
      slots: timeline.slots || [],
      isOpen: timeline.isOpen,
      status: current ? 'OCCUPIED' : 'FREE',
      currentBooking: current,
      user: ROOMS_APP.Auth.getUserContext(),
      resources: allResources,
      config: {
        bookingEnabled: ROOMS_APP.getBooleanConfig('BOOKING_ENABLED', true),
        allowRecurring: ROOMS_APP.getBooleanConfig('ALLOW_RECURRING', true),
        showBookerName: ROOMS_APP.getBooleanConfig('SHOW_BOOKER_NAME', false)
      }
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
