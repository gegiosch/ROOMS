var ROOMS_APP = ROOMS_APP || {};

ROOMS_APP.Recurring = {
  buildRowsPreview: function (rows, options) {
    var actor = (options && options.actor) || ROOMS_APP.Auth.getUserContext();
    ROOMS_APP.Auth.assertAllowedDomain(actor.email);
    if (!actor.canBook) {
      throw new Error('Booking permission required.');
    }
    if (!ROOMS_APP.getBooleanConfig('ALLOW_RECURRING', true) && !actor.canAccessAdmin) {
      throw new Error('Recurring bookings are disabled.');
    }
    var occurrences = this.expandRows_(rows || [], actor);
    var workingRows = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.BOOKINGS);
    return this.previewOccurrencesAgainstRows_(occurrences, actor, workingRows);
  },

  commitRows: function (rows, options) {
    var settings = options || {};
    var actor = settings.actor || ROOMS_APP.Auth.getUserContext();
    ROOMS_APP.Auth.assertAllowedDomain(actor.email);
    if (!actor.canBook) {
      throw new Error('Booking permission required.');
    }
    if (!ROOMS_APP.getBooleanConfig('ALLOW_RECURRING', true) && !actor.canAccessAdmin) {
      throw new Error('Recurring bookings are disabled.');
    }
    var self = this;
    var occurrences = this.expandRows_(rows || [], actor);
    if (!occurrences.length) {
      return this.emptyRowsResult_();
    }
    return ROOMS_APP.Booking.withBookingLock_(function () {
      var workingRows = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.BOOKINGS);
      var preview = self.previewOccurrencesAgainstRows_(occurrences, actor, workingRows);
      var hasBlocked = preview.conflictCount || preview.invalidCount;
      if (hasBlocked && !settings.allowPartial) {
        return {
          savedCount: 0,
          failedCount: preview.conflictCount + preview.invalidCount,
          saved: [],
          failed: preview.conflicts.concat(preview.invalid),
          preview: preview,
          blockedByConflicts: true
        };
      }
      var nowIso = ROOMS_APP.toIsoDateTime(new Date());
      var rowsToAppend = [];
      var saved = [];
      preview.valid.forEach(function (entry) {
        var bookingId = Utilities.getUuid();
        var seriesId = entry.seriesId || (entry.isRecurring ? entry.seriesKey : '');
        var booking = ROOMS_APP.Booking.buildBookingFromValidation_(bookingId, seriesId, entry.validation, actor, nowIso, null);
        workingRows.push(booking);
        rowsToAppend.push(booking);
        saved.push({
          draftId: entry.draftId,
          occurrenceId: entry.occurrenceId,
          bookingId: bookingId,
          bookingDate: entry.bookingDate,
          startTime: entry.startTime,
          endTime: entry.endTime
        });
      });
      if (rowsToAppend.length) {
        ROOMS_APP.DB.appendRows(ROOMS_APP.SHEET_NAMES.BOOKINGS, rowsToAppend);
      }
      return {
        savedCount: saved.length,
        failedCount: preview.conflictCount + preview.invalidCount,
        saved: saved,
        failed: preview.conflicts.concat(preview.invalid),
        preview: preview,
        blockedByConflicts: false
      };
    });
  },

  emptyRowsResult_: function () {
    return {
      total: 0,
      validCount: 0,
      conflictCount: 0,
      invalidCount: 0,
      valid: [],
      conflicts: [],
      invalid: [],
      occurrences: []
    };
  },

  previewOccurrencesAgainstRows_: function (occurrences, actor, workingRows) {
    var previewRows = workingRows ? workingRows.slice() : [];
    var valid = [];
    var conflicts = [];
    var invalid = [];
    occurrences.forEach(function (occurrence) {
      try {
        var validation = ROOMS_APP.Booking.validateAgainstWorkingRows_(occurrence.payload, actor, previewRows, '', {});
        valid.push(Object.assign({}, occurrence, {
          status: 'VALID',
          validation: validation,
          errors: []
        }));
        previewRows.push(ROOMS_APP.Booking.buildBookingFromValidation_(
          'PREVIEW_' + occurrence.occurrenceId,
          occurrence.seriesKey || '',
          validation,
          actor,
          ROOMS_APP.toIsoDateTime(new Date()),
          null
        ));
      } catch (error) {
        var message = String(error && error.message ? error.message : error);
        var row = Object.assign({}, occurrence, {
          status: message === ROOMS_APP.Booking.SLOT_TAKEN_MESSAGE_ ? 'CONFLICT' : 'INVALID',
          message: message,
          errors: [message]
        });
        if (row.status === 'CONFLICT') {
          conflicts.push(row);
        } else {
          invalid.push(row);
        }
      }
    });
    return {
      total: occurrences.length,
      validCount: valid.length,
      conflictCount: conflicts.length,
      invalidCount: invalid.length,
      valid: valid,
      conflicts: conflicts,
      invalid: invalid,
      occurrences: valid.concat(conflicts).concat(invalid)
    };
  },

  expandRows_: function (rows, actor) {
    var self = this;
    var output = [];
    var maxOccurrences = ROOMS_APP.getNumberConfig('RECURRING_MAX_OCCURRENCES', 20);
    var maxWeeks = ROOMS_APP.getNumberConfig('RECURRING_MAX_WEEKS', 12);
    (Array.isArray(rows) ? rows : []).forEach(function (row, rowIndex) {
      var normalized = self.normalizeRow_(row, actor, rowIndex);
      var dates = [normalized.bookingDate];
      if (normalized.repeatWeekly) {
        if (!normalized.repeatUntil) {
          throw new Error('Indicare la data finale della ricorrenza.');
        }
        if (normalized.repeatUntil < normalized.bookingDate) {
          throw new Error('La data finale della ricorrenza non può precedere la data iniziale.');
        }
        dates = self.listWeeklyDates_(normalized.bookingDate, normalized.repeatUntil, actor.canAccessAdmin ? 366 : maxWeeks * 7);
        if (!actor.canAccessAdmin && dates.length > maxOccurrences) {
          dates = dates.slice(0, maxOccurrences);
        }
      }
      var seriesKey = normalized.repeatWeekly ? Utilities.getUuid() : '';
      dates.forEach(function (dateString, occurrenceIndex) {
        output.push({
          draftId: normalized.draftId,
          occurrenceId: normalized.draftId + ':' + dateString + ':' + normalized.startTime + ':' + normalized.endTime,
          seriesKey: seriesKey,
          isRecurring: Boolean(normalized.repeatWeekly),
          bookingDate: dateString,
          startTime: normalized.startTime,
          endTime: normalized.endTime,
          payload: Object.assign({}, normalized.payload, {
            bookingDate: dateString,
            seriesId: normalized.repeatWeekly ? seriesKey : normalized.payload.seriesId
          }),
          occurrenceIndex: occurrenceIndex
        });
      });
    });
    return output;
  },

  normalizeRow_: function (row, actor, rowIndex) {
    var source = row || {};
    var resourceId = ROOMS_APP.normalizeString(source.resourceId || source.ResourceId);
    var bookingDate = ROOMS_APP.toIsoDate(source.bookingDate || source.date || source.BookingDate);
    var startTime = ROOMS_APP.toTimeString(source.startTime || source.StartTime);
    var endTime = ROOMS_APP.toTimeString(source.endTime || source.EndTime);
    var activityDescription = ROOMS_APP.normalizeString(source.activityDescription || source.ActivityDescription);
    if (!activityDescription) {
      throw new Error('Descrizione attività obbligatoria.');
    }
    return {
      draftId: ROOMS_APP.normalizeString(source.draftId || source.bookingId || source.BookingId) || ('row-' + String((rowIndex || 0) + 1)),
      bookingDate: bookingDate,
      startTime: startTime,
      endTime: endTime,
      repeatWeekly: ROOMS_APP.asBoolean(source.repeatWeekly || source.RepeatWeekly),
      repeatUntil: ROOMS_APP.toIsoDate(source.repeatUntil || source.recurrenceEndDate || source.RepeatUntil || source.RecurringUntil),
      payload: {
        bookingId: ROOMS_APP.normalizeString(source.bookingId || source.BookingId),
        seriesId: ROOMS_APP.normalizeString(source.seriesId || source.SeriesId),
        resourceId: resourceId,
        bookingDate: bookingDate,
        startTime: startTime,
        endTime: endTime,
        title: ROOMS_APP.normalizeString(source.title || source.Title),
        activityDescription: activityDescription,
        displayMode: ROOMS_APP.Booking.normalizeDisplayMode_(source.displayMode || source.DisplayMode),
        notes: ROOMS_APP.normalizeString(source.notes || source.Notes),
        bookerName: ROOMS_APP.normalizeString(source.bookerName || source.BookerName),
        bookerSurname: ROOMS_APP.normalizeString(source.bookerSurname || source.BookerSurname)
      }
    };
  },

  listWeeklyDates_: function (startDate, endDate, maxDaysSpan) {
    var dates = [];
    var cursor = ROOMS_APP.combineDateTime(startDate, '12:00');
    var end = ROOMS_APP.combineDateTime(endDate, '12:00');
    var deadline = new Date(cursor.getTime() + maxDaysSpan * 86400000);
    var hardEnd = end.getTime() < deadline.getTime() ? end : deadline;
    while (cursor.getTime() <= hardEnd.getTime()) {
      dates.push(Utilities.formatDate(cursor, ROOMS_APP.getTimezone(), 'yyyy-MM-dd'));
      cursor = new Date(cursor.getTime() + 7 * 86400000);
    }
    return dates;
  },

  previewWeekly: function (payload) {
    var actor = ROOMS_APP.Auth.getUserContext();
    ROOMS_APP.Auth.assertAllowedDomain(actor.email);
    if (!actor.canBook) {
      throw new Error('Booking permission required.');
    }
    if (!ROOMS_APP.getBooleanConfig('ALLOW_RECURRING', true) && !actor.canAccessAdmin) {
      throw new Error('Recurring bookings are disabled.');
    }

    var startDate = ROOMS_APP.toIsoDate(payload.startDate);
    var endDate = ROOMS_APP.toIsoDate(payload.endDate);
    var weekday = ROOMS_APP.normalizeString(payload.weekday) || ROOMS_APP.getWeekdayName(startDate);
    var maxOccurrences = ROOMS_APP.getNumberConfig('RECURRING_MAX_OCCURRENCES', 20);
    var maxWeeks = ROOMS_APP.getNumberConfig('RECURRING_MAX_WEEKS', 12);
    var dates = this.listMatchingDates_(startDate, endDate, weekday, actor.canAccessAdmin ? 366 : maxWeeks * 7);
    var results = [];

    dates.forEach(function (dateString) {
      if (!actor.canAccessAdmin && results.length >= maxOccurrences) {
        return;
      }

      var validation = ROOMS_APP.Policy.validateBookingRequest({
        resourceId: payload.resourceId,
        bookingDate: dateString,
        startTime: payload.startTime,
        endTime: payload.endTime,
        title: payload.title,
        activityDescription: payload.activityDescription,
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
    ROOMS_APP.Auth.assertAllowedDomain(actor.email);
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
        activityDescription: payload.activityDescription,
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
