var ROOMS_APP = ROOMS_APP || {};

ROOMS_APP.Recurring = {
  previewWeekly: function (payload) {
    var actor = ROOMS_APP.Auth.getUserContext();
    if (!ROOMS_APP.getBooleanConfig('ALLOW_RECURRING', true) && !actor.isAdmin) {
      throw new Error('Recurring bookings are disabled.');
    }

    var startDate = ROOMS_APP.toIsoDate(payload.startDate);
    var endDate = ROOMS_APP.toIsoDate(payload.endDate);
    var weekday = ROOMS_APP.normalizeString(payload.weekday) || ROOMS_APP.getWeekdayName(startDate);
    var maxOccurrences = ROOMS_APP.getNumberConfig('RECURRING_MAX_OCCURRENCES', 20);
    var maxWeeks = ROOMS_APP.getNumberConfig('RECURRING_MAX_WEEKS', 12);
    var dates = this.listMatchingDates_(startDate, endDate, weekday, actor.isAdmin ? 366 : maxWeeks * 7);
    var results = [];

    dates.forEach(function (dateString) {
      if (!actor.isAdmin && results.length >= maxOccurrences) {
        return;
      }

      var validation = ROOMS_APP.Policy.validateBookingRequest({
        resourceId: payload.resourceId,
        bookingDate: dateString,
        startTime: payload.startTime,
        endTime: payload.endTime,
        title: payload.title,
        notes: payload.notes,
        bookerName: payload.bookerName,
        bookerSurname: payload.bookerSurname
      }, actor);

      results.push({
        bookingDate: dateString,
        weekday: ROOMS_APP.getWeekdayName(dateString),
        status: validation.ok ? 'VALID' : 'SKIPPED',
        errors: validation.errors
      });
    });

    return {
      resourceId: payload.resourceId,
      weekday: weekday,
      total: results.length,
      validCount: results.filter(function (row) { return row.status === 'VALID'; }).length,
      skippedCount: results.filter(function (row) { return row.status !== 'VALID'; }).length,
      occurrences: results
    };
  },

  commitWeekly: function (payload) {
    var preview = this.previewWeekly(payload);
    var actor = ROOMS_APP.Auth.getUserContext();
    var seriesId = Utilities.getUuid();
    var created = [];

    preview.occurrences.forEach(function (occurrence) {
      if (occurrence.status !== 'VALID') {
        return;
      }

      created.push(ROOMS_APP.Booking.createBooking({
        resourceId: payload.resourceId,
        bookingDate: occurrence.bookingDate,
        startTime: payload.startTime,
        endTime: payload.endTime,
        title: payload.title,
        notes: payload.notes,
        bookerName: payload.bookerName,
        bookerSurname: payload.bookerSurname,
        seriesId: seriesId
      }, { actor: actor }));
    });

    ROOMS_APP.Booking.writeAudit_('CREATE_RECURRING_SERIES', '', seriesId, payload.resourceId, actor.email, 'OK', {
      preview: preview,
      createdCount: created.length
    });

    return {
      seriesId: seriesId,
      createdCount: created.length,
      preview: preview,
      bookings: created
    };
  },

  listMatchingDates_: function (startDate, endDate, weekday, maxDaysSpan) {
    var dates = [];
    var cursor = ROOMS_APP.combineDateTime(startDate, '12:00');
    var end = ROOMS_APP.combineDateTime(endDate, '12:00');
    var deadline = new Date(cursor.getTime() + maxDaysSpan * 86400000);
    var hardEnd = end.getTime() < deadline.getTime() ? end : deadline;

    while (cursor.getTime() <= hardEnd.getTime()) {
      var dateString = Utilities.formatDate(cursor, ROOMS_APP.getTimezone(), 'yyyy-MM-dd');
      if (ROOMS_APP.getWeekdayName(dateString) === weekday) {
        dates.push(dateString);
      }
      cursor = new Date(cursor.getTime() + 86400000);
    }

    return dates;
  }
};
