var ROOMS_APP = ROOMS_APP || {};

ROOMS_APP.Booking = {
  SLOT_TAKEN_MESSAGE_: 'Lo slot non è più disponibile. Aggiorna la disponibilità dell’aula.',
  AULA_MAGNA_RESOURCE_ID_: 'AULA_MAGNA',

  normalizeDisplayMode_: function (value) {
    return String(value || '').toUpperCase() === 'ACTIVITY' ? 'ACTIVITY' : 'TEACHER';
  },

  normalizeRequesterMode_: function (value) {
    return ROOMS_APP.Policy && typeof ROOMS_APP.Policy.normalizeBookingRequesterMode_ === 'function'
      ? ROOMS_APP.Policy.normalizeBookingRequesterMode_(value)
      : (String(value || '').toUpperCase() === 'MANUAL' || String(value || '').toUpperCase() === 'NONE'
        ? String(value || '').toUpperCase()
        : 'LIST');
  },

  getBookingDisplayLabel_: function (booking) {
    var displayMode = this.normalizeDisplayMode_(booking && (booking.DisplayMode || booking.displayMode));
    var actorLabel = ROOMS_APP.normalizeString([
      ROOMS_APP.normalizeString(booking && (booking.BookerSurname || booking.bookerSurname)),
      ROOMS_APP.normalizeString(booking && (booking.BookerName || booking.bookerName))
    ].filter(function (token) {
      return Boolean(token);
    }).join(' '));
    var activityLabel = ROOMS_APP.normalizeString(
      (booking && (booking.ActivityDescription || booking.activityDescription)) ||
      (booking && (booking.Title || booking.title))
    );
    if (displayMode === 'ACTIVITY' && activityLabel) {
      return activityLabel;
    }
    return actorLabel ||
      ROOMS_APP.normalizeString(booking && (booking.RequesterManualName || booking.requesterManualName)) ||
      activityLabel ||
      ROOMS_APP.normalizeString(booking && (booking.BookerEmail || booking.bookerEmail)) ||
      'N/D';
  },

  isBookingCancelled_: function (booking) {
    return ROOMS_APP.normalizeString(booking && booking.Status).toUpperCase() === 'CANCELLED';
  },

  listBookingsForDay: function (resourceId, dateString) {
    var rows = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.BOOKINGS).filter(function (row) {
      return row.ResourceId === resourceId &&
        row.BookingDate === dateString &&
        !ROOMS_APP.Booking.isBookingCancelled_(row);
    });

    return ROOMS_APP.sortBy(rows, ['StartTime', 'EndTime', 'CreatedAtISO']);
  },

  listBookingsForDate: function (dateString) {
    return ROOMS_APP.sortBy(
      ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.BOOKINGS).filter(function (row) {
        return row.BookingDate === dateString && !ROOMS_APP.Booking.isBookingCancelled_(row);
      }),
      ['ResourceId', 'StartTime', 'EndTime']
    );
  },

  getBookingReportModel: function (startDate, endDate, actor) {
    var reportActor = actor || ROOMS_APP.Auth.requireCanViewReports();
    if (!ROOMS_APP.Auth.canAccessAdminTab(reportActor, 'report')) {
      throw new Error('Report access required.');
    }
    var today = ROOMS_APP.toIsoDate(ROOMS_APP.Auth.getEffectiveNow(null, reportActor));
    var normalizedStart = ROOMS_APP.toIsoDate(startDate || today) || today;
    var normalizedEnd = ROOMS_APP.toIsoDate(endDate || normalizedStart) || normalizedStart;
    var resourceMap = this.buildReportResourceMap_();
    var sections = {
      regular: [],
      aulaMagna: []
    };
    var self = this;
    if (normalizedEnd < normalizedStart) {
      normalizedEnd = normalizedStart;
    }
    ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.BOOKINGS).forEach(function (booking) {
      var bookingDate = ROOMS_APP.toIsoDate(booking.BookingDate);
      var resourceId = ROOMS_APP.normalizeString(booking.ResourceId);
      var resource = resourceMap[resourceId] || resourceMap[resourceId.toUpperCase()] || null;
      var row;
      if (!bookingDate || bookingDate < normalizedStart || bookingDate > normalizedEnd || self.isBookingCancelled_(booking)) {
        return;
      }
      row = self.buildBookingReportRow_(booking, resource);
      if (self.isAulaMagnaReportResource_(booking.ResourceId, resource)) {
        sections.aulaMagna.push(row);
      } else {
        sections.regular.push(row);
      }
    });
    ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.AULA_MAGNA_EVENTS).forEach(function (eventRow) {
      var eventDate = ROOMS_APP.toIsoDate(eventRow.EventDate);
      var isActive = ROOMS_APP.normalizeString(eventRow.IsActive) === '' ? true : ROOMS_APP.asBoolean(eventRow.IsActive);
      if (!isActive || !eventDate || eventDate < normalizedStart || eventDate > normalizedEnd) {
        return;
      }
      sections.aulaMagna.push(self.buildAulaMagnaEventReportRow_(eventRow, resourceMap));
    });
    sections.regular = ROOMS_APP.sortBy(sections.regular, ['date', 'startTime', 'room']);
    sections.aulaMagna = ROOMS_APP.sortBy(sections.aulaMagna, ['date', 'startTime', 'room']);
    return {
      startDate: normalizedStart,
      endDate: normalizedEnd,
      today: today,
      regularBookings: sections.regular,
      aulaMagnaBookings: sections.aulaMagna,
      user: reportActor
    };
  },

  buildReportResourceMap_: function () {
    var map = {};
    ROOMS_APP.Board.listResources_().forEach(function (resource) {
      var resourceId;
      if (!resource || !resource.ResourceId) {
        return;
      }
      resourceId = ROOMS_APP.normalizeString(resource.ResourceId);
      map[resourceId] = resource;
      map[resourceId.toUpperCase()] = resource;
    });
    return map;
  },

  isAulaMagnaReportResource_: function (resourceId, resource) {
    var normalizedId = ROOMS_APP.normalizeString(resourceId).toUpperCase();
    if (normalizedId === this.AULA_MAGNA_RESOURCE_ID_) {
      return true;
    }
    if (resource && this.isAulaMagnaResource_(resource)) {
      return true;
    }
    return ROOMS_APP.slugify(resource && resource.DisplayName || resourceId) === this.AULA_MAGNA_RESOURCE_ID_;
  },

  getBookingReportRequesterLabel_: function (booking) {
    var requesterMode = this.normalizeRequesterMode_(booking && booking.RequesterMode);
    var listedName = ROOMS_APP.normalizeString([
      ROOMS_APP.normalizeString(booking && booking.BookerSurname),
      ROOMS_APP.normalizeString(booking && booking.BookerName)
    ].filter(function (token) {
      return Boolean(token);
    }).join(' '));
    if (requesterMode === 'MANUAL') {
      return ROOMS_APP.normalizeString(booking && booking.RequesterManualName);
    }
    if (requesterMode === 'NONE') {
      return '';
    }
    return listedName || ROOMS_APP.normalizeString(booking && booking.RequesterManualName);
  },

  buildBookingReportRow_: function (booking, resource) {
    return {
      date: ROOMS_APP.toIsoDate(booking.BookingDate),
      startTime: ROOMS_APP.toTimeString(booking.StartTime),
      endTime: ROOMS_APP.toTimeString(booking.EndTime),
      room: ROOMS_APP.normalizeString(resource && resource.DisplayName || booking.ResourceId),
      activity: ROOMS_APP.normalizeString(booking.ActivityDescription || booking.Title),
      requester: this.getBookingReportRequesterLabel_(booking),
      notes: ROOMS_APP.normalizeString(booking.Notes),
      source: 'BOOKING'
    };
  },

  buildAulaMagnaEventReportRow_: function (eventRow, resourceMap) {
    var resourceId = ROOMS_APP.normalizeString(eventRow.ResourceId || this.AULA_MAGNA_RESOURCE_ID_);
    var resource = resourceMap[resourceId] || resourceMap[resourceId.toUpperCase()] || null;
    return {
      date: ROOMS_APP.toIsoDate(eventRow.EventDate),
      startTime: ROOMS_APP.toTimeString(eventRow.StartTime),
      endTime: ROOMS_APP.toTimeString(eventRow.EndTime),
      room: ROOMS_APP.normalizeString(resource && resource.DisplayName || 'Aula Magna'),
      activity: ROOMS_APP.normalizeString(eventRow.EventName),
      requester: '',
      notes: ROOMS_APP.normalizeString(eventRow.Notes),
      source: 'AULA_MAGNA_EVENT'
    };
  },

  listUpcomingBookingsForRoom: function (resourceId, fromDate) {
    var startDate = fromDate || ROOMS_APP.toIsoDate(new Date());
    return ROOMS_APP.sortBy(
      ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.BOOKINGS).filter(function (row) {
        return row.ResourceId === resourceId &&
          row.BookingDate >= startDate &&
          !ROOMS_APP.Booking.isBookingCancelled_(row);
      }),
      ['BookingDate', 'StartTime', 'EndTime']
    );
  },

  createBooking: function (payload, options) {
    payload = payload || {};
    var actor = (options && options.actor) || ROOMS_APP.Auth.getUserContext();
    ROOMS_APP.Auth.assertAllowedDomain(actor.email);
    if (!actor.canBook) {
      throw new Error('Booking permission required.');
    }
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
    var self = this;
    var booking;
    try {
      booking = this.withBookingLock_(function () {
        var lockedValidation = ROOMS_APP.Policy.validateBookingRequest(normalized, actor);
        if (!lockedValidation.ok) {
          if (self.hasOnlyConflictError_(lockedValidation.errors)) {
            throw new Error(self.SLOT_TAKEN_MESSAGE_);
          }
          throw new Error(lockedValidation.errors.join(' '));
        }

        var nowIso = ROOMS_APP.toIsoDateTime(new Date());
        var model = self.buildBookingFromValidation_(bookingId, seriesId, lockedValidation, actor, nowIso, null);
        ROOMS_APP.DB.appendRows(ROOMS_APP.SHEET_NAMES.BOOKINGS, [model]);
        return model;
      });
    } catch (error) {
      this.writeAudit_('CREATE_BOOKING', bookingId, seriesId, normalized.resourceId, actor.email, 'ERROR', {
        errors: [String(error && error.message ? error.message : error)],
        payload: payload
      });
      throw error;
    }

    this.writeAudit_('CREATE_BOOKING', bookingId, seriesId, booking.ResourceId, actor.email, 'OK', booking);
    return booking;
  },

  updateBooking: function (bookingId, payload) {
    payload = payload || {};
    var actor = ROOMS_APP.Auth.getUserContext();
    ROOMS_APP.Auth.assertAllowedDomain(actor.email);
    if (!actor.canBook) {
      throw new Error('Booking permission required.');
    }
    var existing = this.findBookingById_(bookingId);
    if (!existing) {
      throw new Error('Booking not found.');
    }
    if (!ROOMS_APP.Auth.canManageBooking(existing, actor)) {
      throw new Error('Only the creator or an authorized manager can modify this booking.');
    }

    var request = {
      bookingId: existing.BookingId,
      seriesId: existing.SeriesId,
      resourceId: Object.prototype.hasOwnProperty.call(payload, 'resourceId') ? payload.resourceId : existing.ResourceId,
      bookingDate: Object.prototype.hasOwnProperty.call(payload, 'bookingDate') ? payload.bookingDate : existing.BookingDate,
      startTime: Object.prototype.hasOwnProperty.call(payload, 'startTime') ? payload.startTime : existing.StartTime,
      endTime: Object.prototype.hasOwnProperty.call(payload, 'endTime') ? payload.endTime : existing.EndTime,
      title: Object.prototype.hasOwnProperty.call(payload, 'title') ? payload.title : existing.Title,
      activityDescription: Object.prototype.hasOwnProperty.call(payload, 'activityDescription') ? payload.activityDescription : existing.ActivityDescription,
      displayMode: Object.prototype.hasOwnProperty.call(payload, 'displayMode') ? payload.displayMode : existing.DisplayMode,
      requesterMode: Object.prototype.hasOwnProperty.call(payload, 'requesterMode') ? payload.requesterMode : existing.RequesterMode,
      requesterManualName: Object.prototype.hasOwnProperty.call(payload, 'requesterManualName') ? payload.requesterManualName : existing.RequesterManualName,
      notes: Object.prototype.hasOwnProperty.call(payload, 'notes') ? payload.notes : existing.Notes,
      bookerName: Object.prototype.hasOwnProperty.call(payload, 'bookerName') ? payload.bookerName : existing.BookerName,
      bookerSurname: Object.prototype.hasOwnProperty.call(payload, 'bookerSurname') ? payload.bookerSurname : existing.BookerSurname
    };
    var validation = ROOMS_APP.Policy.validateBookingRequest(request, actor);
    if (!validation.ok) {
      this.writeAudit_('UPDATE_BOOKING', bookingId, existing.SeriesId, existing.ResourceId, actor.email, 'ERROR', {
        errors: validation.errors,
        payload: payload
      });
      throw new Error(validation.errors.join(' '));
    }

    var self = this;
    var updated = this.withBookingLock_(function () {
      var lockedExisting = self.findBookingById_(bookingId);
      if (!lockedExisting) {
        throw new Error('Booking not found.');
      }
      if (!ROOMS_APP.Auth.canManageBooking(lockedExisting, actor)) {
        throw new Error('Only the creator or an authorized manager can modify this booking.');
      }

      var lockedValidation = ROOMS_APP.Policy.validateBookingRequest(request, actor);
      if (!lockedValidation.ok) {
        if (self.hasOnlyConflictError_(lockedValidation.errors)) {
          throw new Error(self.SLOT_TAKEN_MESSAGE_);
        }
        throw new Error(lockedValidation.errors.join(' '));
      }

      var nowIso = ROOMS_APP.toIsoDateTime(new Date());
      var next = self.buildBookingFromValidation_(
        lockedExisting.BookingId,
        lockedValidation.normalized.seriesId || lockedExisting.SeriesId || '',
        lockedValidation,
        actor,
        nowIso,
        lockedExisting
      );
      ROOMS_APP.DB.upsertByKey(ROOMS_APP.SHEET_NAMES.BOOKINGS, 'BookingId', [next]);
      return next;
    });

    this.writeAudit_('UPDATE_BOOKING', updated.BookingId, updated.SeriesId, updated.ResourceId, actor.email, 'OK', updated);
    return updated;
  },

  cancelBooking: function (bookingId, notes) {
    var actor = ROOMS_APP.Auth.getUserContext();
    ROOMS_APP.Auth.assertAllowedDomain(actor.email);
    if (!actor.canBook) {
      throw new Error('Booking permission required.');
    }
    var self = this;
    var booking = this.withBookingLock_(function () {
      var existing = self.findBookingById_(bookingId);
      if (!existing) {
        throw new Error('Booking not found.');
      }

      if (!ROOMS_APP.Auth.canManageBooking(existing, actor)) {
        throw new Error('Only the creator or an authorized manager can cancel this booking.');
      }

      if (self.isBookingCancelled_(existing)) {
        return existing;
      }

      existing.Status = 'CANCELLED';
      existing.UpdatedAtISO = ROOMS_APP.toIsoDateTime(new Date());
      existing.CancelledAtISO = existing.UpdatedAtISO;
      existing.Notes = ROOMS_APP.normalizeString(notes || existing.Notes);
      ROOMS_APP.DB.upsertByKey(ROOMS_APP.SHEET_NAMES.BOOKINGS, 'BookingId', [existing]);
      return existing;
    });

    this.writeAudit_('CANCEL_BOOKING', booking.BookingId, booking.SeriesId, booking.ResourceId, actor.email, 'OK', booking);
    return booking;
  },

  applyRoomChanges: function (resourceId, dateString, changes) {
    var actor = ROOMS_APP.Auth.getUserContext();
    ROOMS_APP.Auth.assertAllowedDomain(actor.email);

    var targetResourceId = ROOMS_APP.normalizeString(resourceId || '');
    var targetDate = ROOMS_APP.toIsoDate(dateString || new Date());
    if (!targetResourceId) {
      throw new Error('ResourceId is required.');
    }

    var normalizedChanges = this.normalizeBatchChanges_(changes);
    if ((normalizedChanges.creates.length || normalizedChanges.updates.length || normalizedChanges.deletes.length) && !actor.canBook) {
      throw new Error('Booking permission required.');
    }
    if ((normalizedChanges.timetableDeletes.length || normalizedChanges.timetableUpdates.length) && !actor.canManageReplacement) {
      throw new Error('Replacement management permission required.');
    }
    if (
      !normalizedChanges.creates.length &&
      !normalizedChanges.updates.length &&
      !normalizedChanges.deletes.length &&
      !normalizedChanges.timetableDeletes.length &&
      !normalizedChanges.timetableUpdates.length
    ) {
      return this.getRoomViewModel(targetResourceId, targetDate);
    }
    var self = this;
    var result;
    try {
      result = this.withBookingLock_(function () {
        var workingRows = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.BOOKINGS);
        var overrideRows = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.POLICY_OVERRIDES);
        var nowIso = ROOMS_APP.toIsoDateTime(new Date());
        var createdIds = [];
        var updatedIds = [];
        var deletedIds = [];
        var overridesDirty = false;
        var ignoredTimetableBookingIds = {};

        function disableTimetableBooking_(bookingId, notes) {
          if (!actor.canManageReplacement) {
            throw new Error('Le occupazioni da orario sono gestibili solo da utenti autorizzati.');
          }
          var occurrence = self.resolveTimetableOccurrenceByBookingId_(targetResourceId, targetDate, bookingId);
          if (!occurrence) {
            throw new Error('Occupazione da orario non trovata.');
          }
          var overrideRow = self.buildTimetableDisableOverrideRow_(occurrence, targetResourceId, targetDate, notes, nowIso);
          overrideRows = self.upsertOverrideRow_(overrideRows, overrideRow);
          overridesDirty = true;
          ignoredTimetableBookingIds[ROOMS_APP.normalizeString(bookingId)] = true;
          return occurrence;
        }

        normalizedChanges.timetableDeletes.forEach(function (entry) {
          var bookingId = ROOMS_APP.normalizeString(entry.bookingId);
          if (!bookingId) {
            return;
          }
          disableTimetableBooking_(bookingId, entry.notes);
          deletedIds.push(bookingId);
        });

        normalizedChanges.timetableUpdates.forEach(function (entry) {
          var bookingId = ROOMS_APP.normalizeString(entry.bookingId);
          if (!bookingId) {
            return;
          }
          var occurrence = disableTimetableBooking_(bookingId, '');
          var payload = entry.payload || {};
          var request = {
            bookingId: '',
            seriesId: '',
            resourceId: targetResourceId,
            bookingDate: targetDate,
            startTime: occurrence.StartTime,
            endTime: occurrence.EndTime,
            title: Object.prototype.hasOwnProperty.call(payload, 'title') ? payload.title : occurrence.Title,
            activityDescription: Object.prototype.hasOwnProperty.call(payload, 'activityDescription') ? payload.activityDescription : occurrence.ActivityDescription,
            displayMode: Object.prototype.hasOwnProperty.call(payload, 'displayMode') ? payload.displayMode : occurrence.DisplayMode,
            requesterMode: Object.prototype.hasOwnProperty.call(payload, 'requesterMode') ? payload.requesterMode : occurrence.RequesterMode,
            requesterManualName: Object.prototype.hasOwnProperty.call(payload, 'requesterManualName') ? payload.requesterManualName : occurrence.RequesterManualName,
            notes: Object.prototype.hasOwnProperty.call(payload, 'notes') ? payload.notes : occurrence.Notes,
            bookerName: Object.prototype.hasOwnProperty.call(payload, 'bookerName') ? payload.bookerName : occurrence.BookerName,
            bookerSurname: Object.prototype.hasOwnProperty.call(payload, 'bookerSurname') ? payload.bookerSurname : occurrence.BookerSurname
          };
          var validation = self.validateAgainstWorkingRows_(request, actor, workingRows, '', ignoredTimetableBookingIds);
          var newBookingId = validation.normalized.bookingId || Utilities.getUuid();
          var seriesId = validation.normalized.seriesId || '';
          var booking = self.buildBookingFromValidation_(newBookingId, seriesId, validation, actor, nowIso, null);
          workingRows.push(booking);
          createdIds.push(newBookingId);
          updatedIds.push(bookingId);
        });

        normalizedChanges.deletes.forEach(function (entry) {
          var bookingId = ROOMS_APP.normalizeString(entry.bookingId);
          if (bookingId.indexOf('TT_') === 0) {
            throw new Error('Le occupazioni da orario sono in sola lettura.');
          }
          var existingIndex = self.findBookingIndexById_(workingRows, bookingId);
          if (existingIndex < 0) {
            throw new Error('Booking not found.');
          }
          var existing = workingRows[existingIndex];
          if (existing.ResourceId !== targetResourceId) {
            throw new Error('Cannot modify bookings for a different room.');
          }
          if (!ROOMS_APP.Auth.canManageBooking(existing, actor)) {
            throw new Error('Only the creator or an authorized manager can cancel this booking.');
          }
          if (self.isBookingCancelled_(existing)) {
            return;
          }

          existing.Status = 'CANCELLED';
          existing.UpdatedAtISO = nowIso;
          existing.CancelledAtISO = nowIso;
          if (Object.prototype.hasOwnProperty.call(entry, 'notes')) {
            existing.Notes = ROOMS_APP.normalizeString(entry.notes);
          }
          workingRows[existingIndex] = existing;
          deletedIds.push(bookingId);
        });

        normalizedChanges.updates.forEach(function (entry) {
          var bookingId = ROOMS_APP.normalizeString(entry.bookingId);
          if (bookingId.indexOf('TT_') === 0) {
            throw new Error('Le occupazioni da orario sono in sola lettura.');
          }
          var existingIndex = self.findBookingIndexById_(workingRows, bookingId);
          if (existingIndex < 0) {
            throw new Error('Booking not found.');
          }
          var existing = workingRows[existingIndex];
          if (self.isBookingCancelled_(existing)) {
            throw new Error('Booking not found.');
          }
          if (existing.ResourceId !== targetResourceId) {
            throw new Error('Cannot modify bookings for a different room.');
          }
          if (!ROOMS_APP.Auth.canManageBooking(existing, actor)) {
            throw new Error('Only the creator or an authorized manager can modify this booking.');
          }

          var payload = entry.payload || {};
          var request = {
            bookingId: existing.BookingId,
            seriesId: existing.SeriesId,
            resourceId: Object.prototype.hasOwnProperty.call(payload, 'resourceId') ? payload.resourceId : existing.ResourceId,
            bookingDate: Object.prototype.hasOwnProperty.call(payload, 'bookingDate') ? payload.bookingDate : existing.BookingDate,
            startTime: Object.prototype.hasOwnProperty.call(payload, 'startTime') ? payload.startTime : existing.StartTime,
            endTime: Object.prototype.hasOwnProperty.call(payload, 'endTime') ? payload.endTime : existing.EndTime,
            title: Object.prototype.hasOwnProperty.call(payload, 'title') ? payload.title : existing.Title,
            activityDescription: Object.prototype.hasOwnProperty.call(payload, 'activityDescription') ? payload.activityDescription : existing.ActivityDescription,
            displayMode: Object.prototype.hasOwnProperty.call(payload, 'displayMode') ? payload.displayMode : existing.DisplayMode,
            requesterMode: Object.prototype.hasOwnProperty.call(payload, 'requesterMode') ? payload.requesterMode : existing.RequesterMode,
            requesterManualName: Object.prototype.hasOwnProperty.call(payload, 'requesterManualName') ? payload.requesterManualName : existing.RequesterManualName,
            notes: Object.prototype.hasOwnProperty.call(payload, 'notes') ? payload.notes : existing.Notes,
            bookerName: Object.prototype.hasOwnProperty.call(payload, 'bookerName') ? payload.bookerName : existing.BookerName,
            bookerSurname: Object.prototype.hasOwnProperty.call(payload, 'bookerSurname') ? payload.bookerSurname : existing.BookerSurname
          };

          if (ROOMS_APP.normalizeString(request.resourceId) !== targetResourceId) {
            throw new Error('Cannot move booking to a different room from this panel.');
          }

          var validation = self.validateAgainstWorkingRows_(request, actor, workingRows, existing.BookingId, ignoredTimetableBookingIds);
          var next = self.buildBookingFromValidation_(
            existing.BookingId,
            validation.normalized.seriesId || existing.SeriesId || '',
            validation,
            actor,
            nowIso,
            existing
          );

          workingRows[existingIndex] = next;
          updatedIds.push(existing.BookingId);
        });

        normalizedChanges.creates.forEach(function (entry) {
          var payload = entry.payload || {};
          var createResourceId = ROOMS_APP.normalizeString(payload.resourceId || targetResourceId);
          if (createResourceId !== targetResourceId) {
            throw new Error('Cannot create bookings for a different room from this panel.');
          }

          var request = {
            bookingId: ROOMS_APP.normalizeString(payload.bookingId || ''),
            seriesId: ROOMS_APP.normalizeString(payload.seriesId || ''),
            resourceId: targetResourceId,
            bookingDate: Object.prototype.hasOwnProperty.call(payload, 'bookingDate') ? payload.bookingDate : targetDate,
            startTime: payload.startTime,
            endTime: payload.endTime,
            title: payload.title,
            activityDescription: payload.activityDescription,
            displayMode: payload.displayMode,
            requesterMode: payload.requesterMode,
            requesterManualName: payload.requesterManualName,
            notes: payload.notes,
            bookerName: payload.bookerName,
            bookerSurname: payload.bookerSurname
          };
          var validation = self.validateAgainstWorkingRows_(request, actor, workingRows, '', ignoredTimetableBookingIds);
          var bookingId = validation.normalized.bookingId || Utilities.getUuid();
          var seriesId = validation.normalized.seriesId || '';
          var booking = self.buildBookingFromValidation_(bookingId, seriesId, validation, actor, nowIso, null);
          workingRows.push(booking);
          createdIds.push(bookingId);
        });

        ROOMS_APP.DB.replaceRows(
          ROOMS_APP.SHEET_NAMES.BOOKINGS,
          ROOMS_APP.DB.getHeaders(ROOMS_APP.SHEET_NAMES.BOOKINGS),
          workingRows
        );
        if (overridesDirty) {
          ROOMS_APP.DB.replaceRows(
            ROOMS_APP.SHEET_NAMES.POLICY_OVERRIDES,
            ROOMS_APP.DB.getHeaders(ROOMS_APP.SHEET_NAMES.POLICY_OVERRIDES),
            overrideRows
          );
        }
        return {
          createdIds: createdIds,
          updatedIds: updatedIds,
          deletedIds: deletedIds,
          overridesUpdated: overridesDirty
        };
      });
    } catch (error) {
      this.writeAudit_('APPLY_ROOM_CHANGES', '', '', targetResourceId, actor.email, 'ERROR', {
        resourceId: targetResourceId,
        date: targetDate,
        error: String(error && error.message ? error.message : error),
        changes: normalizedChanges
      });
      throw error;
    }

    this.writeAudit_('APPLY_ROOM_CHANGES', '', '', targetResourceId, actor.email, 'OK', {
      resourceId: targetResourceId,
      date: targetDate,
      createdCount: result.createdIds.length,
      updatedCount: result.updatedIds.length,
      deletedCount: result.deletedIds.length,
      overridesUpdated: Boolean(result.overridesUpdated),
      createdIds: result.createdIds,
      updatedIds: result.updatedIds,
      deletedIds: result.deletedIds
    });

    return this.getRoomViewModel(targetResourceId, targetDate);
  },

  listAdminBookableResources_: function () {
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

  buildAdminRoomBookingBaseModel_: function (dateString, resourceId, actor) {
    var bookingActor = actor || this.requireAdminRoomBookingActor_();
    var requestedDate = ROOMS_APP.normalizeString(dateString || '');
    var dateState = requestedDate
      ? ROOMS_APP.Replacements.resolveSelectedDate_(requestedDate, true)
      : {
        requestedDate: '',
        selectedDate: '',
        isValid: false,
        autoShifted: false,
        message: ''
      };
    var date = dateState.selectedDate || '';
    var resources = this.listAdminBookableResources_();
    var selectedResourceId = ROOMS_APP.normalizeString(resourceId || '');
    return {
      date: date,
      dateState: dateState,
      resourceId: selectedResourceId,
      resources: resources,
      teacherOptions: ROOMS_APP.Replacements.listAdminTeacherDirectory_(),
      user: bookingActor
    };
  },

  getAdminRoomBookingBootstrapModel: function (dateString, resourceId) {
    var actor = this.requireAdminRoomBookingActor_();
    return this.buildAdminRoomBookingBaseModel_(dateString, resourceId, actor);
  },

  getAdminRoomBookingAvailabilityModel: function (dateString, resourceId, options) {
    var actor = this.requireAdminRoomBookingActor_();
    var baseModel = this.buildAdminRoomBookingBaseModel_(dateString, resourceId, actor);
    var selectedResourceId = baseModel.resourceId;
    var availabilityOptions = options && typeof options === 'object' ? options : {};
    var roomModel = selectedResourceId
      ? this.getRoomViewModel(selectedResourceId, baseModel.date, {
        splitFreeSlotsHalfHour: Boolean(availabilityOptions.splitFreeSlotsHalfHour)
      })
      : null;
    return {
      date: baseModel.date,
      dateState: baseModel.dateState,
      resourceId: selectedResourceId,
      freeSlots: roomModel && roomModel.freeSlots ? roomModel.freeSlots : [],
      bookings: roomModel && roomModel.bookings ? roomModel.bookings : [],
      isOpen: roomModel ? Boolean(roomModel.isOpen) : false,
      statusSummary: roomModel ? (roomModel.statusSummary || '') : '',
      options: roomModel && roomModel.options ? roomModel.options : {
        splitFreeSlotsHalfHour: Boolean(availabilityOptions.splitFreeSlotsHalfHour)
      }
    };
  },

  getAdminRoomBookingModel: function (dateString, resourceId, options) {
    var actor = this.requireAdminRoomBookingActor_();
    var baseModel = this.buildAdminRoomBookingBaseModel_(dateString, resourceId, actor);
    var availabilityModel = this.getAdminRoomBookingAvailabilityModel(baseModel.date, baseModel.resourceId, options);
    return {
      date: baseModel.date,
      dateState: baseModel.dateState,
      resourceId: baseModel.resourceId,
      resources: baseModel.resources,
      freeSlots: availabilityModel.freeSlots,
      bookings: availabilityModel.bookings,
      isOpen: availabilityModel.isOpen,
      statusSummary: availabilityModel.statusSummary,
      options: availabilityModel.options,
      teacherOptions: baseModel.teacherOptions,
      user: actor
    };
  },

  saveAdminRoomBookingRegistry: function (draftRows) {
    var actor = this.requireAdminRoomBookingActor_();
    if (!actor.canBook) {
      throw new Error('Booking permission required.');
    }
    var self = this;
    var rows = (Array.isArray(draftRows) ? draftRows : []).map(function (row, index) {
      var requesterMode = self.normalizeRequesterMode_(row && (row.requesterMode || row.RequesterMode));
      var requesterManualName = requesterMode === 'MANUAL'
        ? ROOMS_APP.normalizeString(row && (row.requesterManualName || row.RequesterManualName))
        : '';
      var teacherName = requesterMode === 'LIST'
        ? ROOMS_APP.normalizeString(row && (row.teacherName || row.bookerName || row.BookerName))
        : requesterManualName;
      var parsedName = self.parseAdminBookingTeacherName_(teacherName);
      var displayMode = requesterMode === 'NONE'
        ? 'ACTIVITY'
        : self.normalizeDisplayMode_(row && (row.displayMode || row.DisplayMode));
      return {
        draftId: ROOMS_APP.normalizeString(row && row.draftId) || ('row-' + String(index + 1)),
        resourceId: ROOMS_APP.normalizeString(row && (row.resourceId || row.ResourceId)),
        bookingDate: ROOMS_APP.toIsoDate(row && (row.bookingDate || row.date || row.BookingDate)),
        startTime: ROOMS_APP.toTimeString(row && (row.startTime || row.StartTime)),
        endTime: ROOMS_APP.toTimeString(row && (row.endTime || row.EndTime)),
        title: ROOMS_APP.normalizeString(row && (row.title || row.Title)),
        activityDescription: ROOMS_APP.normalizeString(row && (row.activityDescription || row.ActivityDescription)),
        displayMode: displayMode,
        requesterMode: requesterMode,
        requesterManualName: requesterManualName,
        notes: ROOMS_APP.normalizeString(row && (row.notes || row.Notes)),
        bookerName: requesterMode === 'NONE' ? '' : (ROOMS_APP.normalizeString(row && (row.bookerName || row.BookerName)) || parsedName.firstName),
        bookerSurname: requesterMode === 'NONE' ? '' : (ROOMS_APP.normalizeString(row && (row.bookerSurname || row.BookerSurname)) || parsedName.surname),
        teacherEmail: requesterMode === 'LIST' ? ROOMS_APP.normalizeEmail(row && row.teacherEmail) : ''
      };
    });
    var result = {
      savedCount: 0,
      failedCount: 0,
      saved: [],
      failed: []
    };
    if (!rows.length) {
      return result;
    }

    result = this.withBookingLock_(function () {
      var workingRows = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.BOOKINGS);
      var rowsToAppend = [];
      var nowIso = ROOMS_APP.toIsoDateTime(new Date());
      var timetableCache = {};
      var validationContext = {
        resourceMap: self.buildAdminBookingResourceMap_(),
        openingCache: {},
        closureRows: ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.CLOSURES)
      };
      var saved = [];
      var failed = [];

      rows.forEach(function (row) {
        var validation;
        var bookingId;
        var seriesId;
        var booking;
        try {
          validation = self.validateAdminBookingDraftRow_(row, actor, workingRows, timetableCache, validationContext);
          bookingId = Utilities.getUuid();
          seriesId = '';
          booking = self.buildBookingFromValidation_(bookingId, seriesId, validation, actor, nowIso, null);
          workingRows.push(booking);
          rowsToAppend.push(booking);
          saved.push({
            draftId: row.draftId,
            bookingId: bookingId
          });
        } catch (error) {
          failed.push({
            draftId: row.draftId,
            message: String(error && error.message ? error.message : error)
          });
        }
      });

      if (rowsToAppend.length) {
        ROOMS_APP.DB.appendRows(ROOMS_APP.SHEET_NAMES.BOOKINGS, rowsToAppend);
      }

      return {
        savedCount: saved.length,
        failedCount: failed.length,
        saved: saved,
        failed: failed
      };
    });

    this.writeAudit_('ADMIN_ROOM_BOOKING_BATCH', '', '', '', actor.email, result.failedCount ? 'PARTIAL' : 'OK', {
      requestedCount: rows.length,
      savedCount: result.savedCount,
      failedCount: result.failedCount
    });
    return result;
  },

  getAdminExistingRoomOccupancies: function (dateString, resourceId) {
    var actor = this.requireAdminRoomBookingActor_();
    var requestedDate = ROOMS_APP.normalizeString(dateString || '');
    var targetDate = requestedDate ? ROOMS_APP.toIsoDate(requestedDate) : '';
    var targetResourceId = ROOMS_APP.normalizeString(resourceId || '');
    return {
      date: targetDate,
      dateState: {
        requestedDate: requestedDate,
        selectedDate: targetDate,
        isValid: Boolean(targetDate),
        autoShifted: false,
        message: ''
      },
      resourceId: targetResourceId,
      resources: this.listAdminBookableResources_(),
      rows: targetResourceId && targetDate
        ? this.buildAdminExistingOccupancyRows_(targetDate, targetResourceId, actor)
        : [],
      user: actor
    };
  },

  applyAdminExistingRoomOccupancyChange: function (payload) {
    payload = payload || {};
    var actor = this.requireAdminRoomBookingActor_();
    var action = ROOMS_APP.normalizeString(payload.action).toUpperCase();
    var sourceType = ROOMS_APP.normalizeString(payload.sourceType).toUpperCase();
    var targetResourceId = ROOMS_APP.normalizeString(payload.resourceId);
    var targetDate = ROOMS_APP.toIsoDate(payload.bookingDate || payload.date);
    var bookingId = ROOMS_APP.normalizeString(payload.bookingId);
    var timetableBookingId = ROOMS_APP.normalizeString(payload.timetableBookingId);
    var overrideId = ROOMS_APP.normalizeString(payload.overrideId);
    var notes = ROOMS_APP.normalizeString(payload.notes);
    var replacementPayload = this.normalizeAdminExistingOccupancyPayload_(payload);
    var changes = {};

    if (!targetResourceId || !targetDate) {
      throw new Error('Data e aula sono obbligatorie.');
    }
    if (action !== 'UPDATE' && action !== 'CANCEL') {
      throw new Error('Azione non valida.');
    }

    if (sourceType === 'PRENOTAZIONE') {
      if (!bookingId) {
        throw new Error('Prenotazione non trovata.');
      }
      if (action === 'UPDATE') {
        changes.updates = [{
          bookingId: bookingId,
          payload: replacementPayload
        }];
      } else {
        changes.deletes = [{
          bookingId: bookingId,
          notes: notes
        }];
      }
      this.applyRoomChanges(targetResourceId, targetDate, changes);
      return this.getAdminExistingRoomOccupancies(targetDate, targetResourceId);
    }

    if (sourceType === 'ORARIO') {
      if (!actor.canManageReplacement) {
        throw new Error('Le occupazioni da orario sono gestibili solo da utenti autorizzati.');
      }
      if (!timetableBookingId) {
        throw new Error('Occupazione da orario non trovata.');
      }
      if (action === 'UPDATE') {
        changes.timetableUpdates = [{
          bookingId: timetableBookingId,
          payload: replacementPayload
        }];
      } else {
        changes.timetableDeletes = [{
          bookingId: timetableBookingId,
          notes: notes
        }];
      }
      this.applyRoomChanges(targetResourceId, targetDate, changes);
      return this.getAdminExistingRoomOccupancies(targetDate, targetResourceId);
    }

    if (sourceType === 'OVERRIDE') {
      if (!actor.canManageReplacement) {
        throw new Error('Le eccezioni da orario sono gestibili solo da utenti autorizzati.');
      }
      if (action === 'UPDATE') {
        if (bookingId) {
          this.applyRoomChanges(targetResourceId, targetDate, {
            updates: [{
              bookingId: bookingId,
              payload: replacementPayload
            }]
          });
        } else {
          this.applyRoomChanges(targetResourceId, targetDate, {
            creates: [replacementPayload]
          });
        }
      } else {
        this.cancelAdminExistingOverride_(targetResourceId, targetDate, overrideId, bookingId, notes, actor);
      }
      return this.getAdminExistingRoomOccupancies(targetDate, targetResourceId);
    }

    throw new Error('Tipo riga non valido.');
  },

  requireAdminRoomBookingActor_: function () {
    return ROOMS_APP.Auth.requireAdminTab('bookings');
  },

  buildAdminExistingOccupancyRows_: function (dateString, resourceId, actor) {
    var self = this;
    var roomModel = this.getRoomViewModel(resourceId, dateString);
    var effectiveBookings = roomModel && Array.isArray(roomModel.bookings) ? roomModel.bookings : [];
    var activeOverrides = this.listActiveTimetableOverrideRows_(resourceId, dateString);
    var overrideByInterval = {};
    var matchedOverrideIds = {};
    var rows = [];

    activeOverrides.forEach(function (overrideRow) {
      var key = self.buildAdminExistingIntervalKey_(overrideRow.StartTime, overrideRow.EndTime);
      if (!key) {
        return;
      }
      overrideByInterval[key] = overrideByInterval[key] || [];
      overrideByInterval[key].push(overrideRow);
    });

    effectiveBookings.forEach(function (booking) {
      var isTimetableRow = ROOMS_APP.normalizeString(booking && booking.SourceKind).toUpperCase() === 'TIMETABLE' ||
        ROOMS_APP.normalizeString(booking && booking.BookingId).indexOf('TT_') === 0;
      var intervalKey = self.buildAdminExistingIntervalKey_(booking.StartTime, booking.EndTime);
      var overrideRow = intervalKey && overrideByInterval[intervalKey] && overrideByInterval[intervalKey].length
        ? overrideByInterval[intervalKey][0]
        : null;
      if (isTimetableRow) {
        rows.push(self.buildAdminExistingTimetableRow_(booking, actor));
        return;
      }
      if (overrideRow) {
        matchedOverrideIds[ROOMS_APP.normalizeString(overrideRow.OverrideId)] = true;
      }
      rows.push(self.buildAdminExistingBookingRow_(booking, overrideRow, actor));
    });

    activeOverrides.forEach(function (overrideRow) {
      if (matchedOverrideIds[ROOMS_APP.normalizeString(overrideRow.OverrideId)]) {
        return;
      }
      rows.push(self.buildAdminExistingOverrideRow_(overrideRow, actor));
    });

    return ROOMS_APP.sortBy(rows, ['bookingDate', 'startTime', 'endTime', 'sourceType', 'rowId']);
  },

  buildAdminExistingBookingRow_: function (booking, overrideRow, actor) {
    var sourceType = overrideRow ? 'OVERRIDE' : 'PRENOTAZIONE';
    return {
      rowId: (sourceType === 'OVERRIDE' ? 'OVR_BOOKING_' : 'BOOKING_') + ROOMS_APP.normalizeString(booking.BookingId),
      sourceType: sourceType,
      sourceLabel: sourceType,
      bookingId: ROOMS_APP.normalizeString(booking.BookingId),
      overrideId: overrideRow ? ROOMS_APP.normalizeString(overrideRow.OverrideId) : '',
      timetableBookingId: '',
      resourceId: ROOMS_APP.normalizeString(booking.ResourceId),
      bookingDate: ROOMS_APP.toIsoDate(booking.BookingDate),
      startTime: ROOMS_APP.toTimeString(booking.StartTime),
      endTime: ROOMS_APP.toTimeString(booking.EndTime),
      title: ROOMS_APP.normalizeString(booking.Title),
      activityDescription: ROOMS_APP.normalizeString(booking.ActivityDescription || booking.Title),
      displayMode: this.normalizeDisplayMode_(booking.DisplayMode),
      requesterMode: this.normalizeRequesterMode_(booking.RequesterMode),
      requesterManualName: ROOMS_APP.normalizeString(booking.RequesterManualName),
      bookerName: ROOMS_APP.normalizeString(booking.BookerName),
      bookerSurname: ROOMS_APP.normalizeString(booking.BookerSurname),
      occupantLabel: this.getBookingDisplayLabel_(booking),
      notes: ROOMS_APP.normalizeString(booking.Notes),
      isSuppression: false,
      canModify: ROOMS_APP.Auth.canManageBooking(booking, actor),
      canCancel: sourceType === 'OVERRIDE' ? Boolean(actor && actor.canManageReplacement) : ROOMS_APP.Auth.canManageBooking(booking, actor),
      helpText: sourceType === 'OVERRIDE'
        ? 'Eccezione applicata solo a questa data e fascia oraria.'
        : 'Prenotazione manuale modificabile normalmente.'
    };
  },

  buildAdminExistingTimetableRow_: function (occurrence, actor) {
    return {
      rowId: 'TIMETABLE_' + ROOMS_APP.normalizeString(occurrence.BookingId),
      sourceType: 'ORARIO',
      sourceLabel: 'ORARIO',
      bookingId: '',
      overrideId: '',
      timetableBookingId: ROOMS_APP.normalizeString(occurrence.BookingId),
      resourceId: ROOMS_APP.normalizeString(occurrence.ResourceId),
      bookingDate: ROOMS_APP.toIsoDate(occurrence.BookingDate),
      startTime: ROOMS_APP.toTimeString(occurrence.StartTime),
      endTime: ROOMS_APP.toTimeString(occurrence.EndTime),
      title: ROOMS_APP.normalizeString(occurrence.Title),
      activityDescription: ROOMS_APP.normalizeString(occurrence.ActivityDescription || occurrence.DisplayLabel || occurrence.Title),
      displayMode: this.normalizeDisplayMode_(occurrence.DisplayMode),
      bookerName: ROOMS_APP.normalizeString(occurrence.BookerName || occurrence.TeacherName),
      bookerSurname: ROOMS_APP.normalizeString(occurrence.BookerSurname || occurrence.ClassCode),
      occupantLabel: ROOMS_APP.Timetable.getDisplayLabel(occurrence),
      notes: ROOMS_APP.normalizeString(occurrence.Notes),
      isSuppression: false,
      canModify: Boolean(actor && actor.canManageReplacement),
      canCancel: Boolean(actor && actor.canManageReplacement),
      helpText: 'Occupazione da orario: le modifiche creano una eccezione solo per questa data.'
    };
  },

  buildAdminExistingOverrideRow_: function (overrideRow, actor) {
    var startTime = ROOMS_APP.toTimeString(overrideRow.StartTime);
    var endTime = ROOMS_APP.toTimeString(overrideRow.EndTime);
    return {
      rowId: 'OVERRIDE_' + ROOMS_APP.normalizeString(overrideRow.OverrideId),
      sourceType: 'OVERRIDE',
      sourceLabel: 'OVERRIDE',
      bookingId: '',
      overrideId: ROOMS_APP.normalizeString(overrideRow.OverrideId),
      timetableBookingId: '',
      resourceId: ROOMS_APP.normalizeString(overrideRow.ResourceId),
      bookingDate: ROOMS_APP.toIsoDate(overrideRow.BookingDate),
      startTime: startTime,
      endTime: endTime,
      title: 'Eccezione orario',
      activityDescription: 'Occorrenza orario rimossa per questa data',
      displayMode: 'ACTIVITY',
      bookerName: '',
      bookerSurname: '',
      occupantLabel: 'Orario sospeso',
      notes: ROOMS_APP.normalizeString(overrideRow.Notes),
      isSuppression: true,
      canModify: Boolean(actor && actor.canManageReplacement),
      canCancel: Boolean(actor && actor.canManageReplacement),
      helpText: 'Eccezione attiva: cancellandola viene ripristinata l\'occupazione da orario sottostante.'
    };
  },

  listActiveTimetableOverrideRows_: function (resourceId, dateString) {
    var targetResourceId = ROOMS_APP.normalizeString(resourceId);
    var targetDate = ROOMS_APP.toIsoDate(dateString);
    return ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.POLICY_OVERRIDES).filter(function (row) {
      return ROOMS_APP.normalizeString(row.ResourceId) === targetResourceId &&
        ROOMS_APP.toIsoDate(row.BookingDate) === targetDate &&
        ROOMS_APP.normalizeString(row.RuleKey).toUpperCase() === 'TIMETABLE_DISABLED' &&
        ROOMS_APP.asBoolean(row.IsEnabled);
    });
  },

  buildAdminExistingIntervalKey_: function (startTime, endTime) {
    var start = ROOMS_APP.toTimeString(startTime);
    var end = ROOMS_APP.toTimeString(endTime);
    return start && end ? (start + '|' + end) : '';
  },

  normalizeAdminExistingOccupancyPayload_: function (payload) {
    var startTime = ROOMS_APP.toTimeString(payload.startTime || payload.StartTime);
    var endTime = ROOMS_APP.toTimeString(payload.endTime || payload.EndTime);
    var bookerName = ROOMS_APP.normalizeString(payload.bookerName || payload.BookerName);
    var bookerSurname = ROOMS_APP.normalizeString(payload.bookerSurname || payload.BookerSurname);
    return {
      resourceId: ROOMS_APP.normalizeString(payload.resourceId || payload.ResourceId),
      bookingDate: ROOMS_APP.toIsoDate(payload.bookingDate || payload.date || payload.BookingDate),
      startTime: startTime,
      endTime: endTime,
      title: ROOMS_APP.normalizeString(payload.title || payload.Title),
      activityDescription: ROOMS_APP.normalizeString(payload.activityDescription || payload.ActivityDescription),
      displayMode: this.normalizeDisplayMode_(payload.displayMode || payload.DisplayMode),
      notes: ROOMS_APP.normalizeString(payload.notes || payload.Notes),
      bookerName: bookerName,
      bookerSurname: bookerSurname
    };
  },

  cancelAdminExistingOverride_: function (resourceId, dateString, overrideId, bookingId, notes, actor) {
    var self = this;
    var targetResourceId = ROOMS_APP.normalizeString(resourceId);
    var targetDate = ROOMS_APP.toIsoDate(dateString);
    var targetOverrideId = ROOMS_APP.normalizeString(overrideId);
    var targetBookingId = ROOMS_APP.normalizeString(bookingId);
    if (!targetOverrideId) {
      throw new Error('Eccezione non trovata.');
    }
    return this.withBookingLock_(function () {
      var overrideRows = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.POLICY_OVERRIDES);
      var bookingRows = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.BOOKINGS);
      var overrideIndex = -1;
      var bookingIndex = -1;
      var nowIso = ROOMS_APP.toIsoDateTime(new Date());
      var index;

      for (index = 0; index < overrideRows.length; index += 1) {
        if (ROOMS_APP.normalizeString(overrideRows[index] && overrideRows[index].OverrideId) === targetOverrideId) {
          overrideIndex = index;
          break;
        }
      }
      if (overrideIndex < 0) {
        throw new Error('Eccezione non trovata.');
      }
      if (ROOMS_APP.normalizeString(overrideRows[overrideIndex].ResourceId) !== targetResourceId ||
        ROOMS_APP.toIsoDate(overrideRows[overrideIndex].BookingDate) !== targetDate) {
        throw new Error('Eccezione non coerente con data e aula selezionate.');
      }

      overrideRows[overrideIndex].IsEnabled = 'FALSE';
      if (Object.prototype.hasOwnProperty.call(overrideRows[overrideIndex], 'Notes')) {
        overrideRows[overrideIndex].Notes = notes || overrideRows[overrideIndex].Notes || '';
      }

      if (targetBookingId) {
        bookingIndex = self.findBookingIndexById_(bookingRows, targetBookingId);
        if (bookingIndex < 0) {
          throw new Error('Prenotazione collegata all\'eccezione non trovata.');
        }
        if (!ROOMS_APP.Auth.canManageBooking(bookingRows[bookingIndex], actor)) {
          throw new Error('Only the creator or an authorized manager can cancel this booking.');
        }
        bookingRows[bookingIndex].Status = 'CANCELLED';
        bookingRows[bookingIndex].UpdatedAtISO = nowIso;
        bookingRows[bookingIndex].CancelledAtISO = nowIso;
        if (notes) {
          bookingRows[bookingIndex].Notes = notes;
        }
        ROOMS_APP.DB.replaceRows(
          ROOMS_APP.SHEET_NAMES.BOOKINGS,
          ROOMS_APP.DB.getHeaders(ROOMS_APP.SHEET_NAMES.BOOKINGS),
          bookingRows
        );
      }

      ROOMS_APP.DB.replaceRows(
        ROOMS_APP.SHEET_NAMES.POLICY_OVERRIDES,
        ROOMS_APP.DB.getHeaders(ROOMS_APP.SHEET_NAMES.POLICY_OVERRIDES),
        overrideRows
      );
      self.writeAudit_('ADMIN_ROOM_OVERRIDE_CANCEL', targetBookingId, '', targetResourceId, actor.email, 'OK', {
        overrideId: targetOverrideId,
        bookingDate: targetDate,
        notes: notes
      });
      return true;
    });
  },

  validateAdminBookingDraftRow_: function (row, actor, workingRows, timetableCache, validationContext) {
    var context = validationContext || {};
    var errors = [];
    var resource = context.resourceMap ? context.resourceMap[row.resourceId] : ROOMS_APP.Policy.getResource(row.resourceId);
    var openingKey = [row.resourceId, row.bookingDate].join('|');
    var opening;
    var closure;
    if (!row.resourceId || !row.bookingDate || !row.startTime || !row.endTime) {
      errors.push('Aula, data e orario sono obbligatori.');
    }
    if (row.requesterMode === 'LIST' && !row.teacherEmail) {
      errors.push('Seleziona il docente/richiedente.');
    }
    if (row.requesterMode === 'MANUAL' && !row.requesterManualName) {
      errors.push('Inserire il nominativo manuale.');
    }
    if (!row.activityDescription) {
      errors.push('Descrizione attività obbligatoria.');
    }
    if (!resource || !ROOMS_APP.asBoolean(resource.IsActive) || !ROOMS_APP.asBoolean(resource.IsBookable)) {
      errors.push('Aula non prenotabile.');
    }
    if (resource && ROOMS_APP.slugify(resource.DisplayName || '') === 'aula_magna') {
      errors.push('Aula Magna è gestita tramite eventi dedicati.');
    }
    if (!/^\d{4}-\d{2}-\d{2}$/.test(row.bookingDate || '')) {
      errors.push('Formato data non valido.');
    }
    if (!/^\d{2}:\d{2}$/.test(row.startTime || '') || !/^\d{2}:\d{2}$/.test(row.endTime || '')) {
      errors.push('Formato orario non valido.');
    }
    if (row.startTime >= row.endTime) {
      errors.push('Ora fine precedente o uguale all\'inizio.');
    }
    if (errors.length) {
      throw new Error(errors.join(' '));
    }

    context.openingCache = context.openingCache || {};
    if (!Object.prototype.hasOwnProperty.call(context.openingCache, openingKey)) {
      context.openingCache[openingKey] = ROOMS_APP.Policy.getEffectiveOpeningForResource(row.resourceId, row.bookingDate);
    }
    opening = context.openingCache[openingKey];
    if (!opening.isOpen || row.startTime < opening.openTime || row.endTime > opening.closeTime) {
      throw new Error('Slot non disponibile nella finestra di apertura dell\'aula.');
    }
    closure = this.findAdminBlockingClosure_(context.closureRows, row.bookingDate, row.startTime, row.endTime);
    if (closure) {
      throw new Error('Slot non disponibile per chiusura: ' + (closure.Label || 'chiusura'));
    }
    if (
      this.hasConflictInRows_(workingRows, row.resourceId, row.bookingDate, row.startTime, row.endTime, '') ||
      this.hasAdminTimetableConflict_(row.resourceId, row.bookingDate, row.startTime, row.endTime, timetableCache)
    ) {
      throw new Error('Slot non più disponibile al momento del salvataggio.');
    }

    return {
      ok: true,
      errors: [],
      normalized: {
        resourceId: row.resourceId,
        bookingDate: row.bookingDate,
        startTime: row.startTime,
        endTime: row.endTime,
        title: row.title || (resource.DisplayName || row.resourceId),
        activityDescription: row.activityDescription,
        displayMode: row.requesterMode === 'NONE' ? 'ACTIVITY' : this.normalizeDisplayMode_(row.displayMode),
        requesterMode: row.requesterMode,
        requesterManualName: row.requesterMode === 'MANUAL' ? row.requesterManualName : '',
        notes: row.notes,
        bookerName: row.bookerName,
        bookerSurname: row.bookerSurname,
        seriesId: '',
        bookingId: ''
      },
      resource: resource,
      actor: actor
    };
  },

  hasAdminTimetableConflict_: function (resourceId, bookingDate, startTime, endTime, timetableCache) {
    var key = [resourceId, bookingDate].join('|');
    var cache = timetableCache || {};
    if (!Object.prototype.hasOwnProperty.call(cache, key)) {
      cache[key] = ROOMS_APP.Timetable.listOccupanciesForDate(resourceId, bookingDate).filter(function (occupancy) {
        return ROOMS_APP.Timetable.isBlockingOccurrence(occupancy);
      });
    }
    return (cache[key] || []).some(function (occupancy) {
      return !(endTime <= occupancy.StartTime || startTime >= occupancy.EndTime);
    });
  },

  buildAdminBookingResourceMap_: function () {
    var output = {};
    ROOMS_APP.Board.listResources_().forEach(function (resource) {
      if (resource && resource.ResourceId) {
        output[resource.ResourceId] = resource;
      }
    });
    return output;
  },

  findAdminBlockingClosure_: function (closureRows, dateString, startTime, endTime) {
    return (closureRows || []).filter(function (closure) {
      var closureStart;
      var closureEnd;
      if (!ROOMS_APP.asBoolean(closure && closure.IsBlocked)) {
        return false;
      }
      if (closure.StartDate > dateString || closure.EndDate < dateString) {
        return false;
      }
      closureStart = closure.StartTime || '00:00';
      closureEnd = closure.EndTime || '23:59';
      return !(endTime <= closureStart || startTime >= closureEnd);
    })[0] || null;
  },

  parseAdminBookingTeacherName_: function (value) {
    var tokens = ROOMS_APP.normalizeString(value).split(/\s+/).filter(function (token) {
      return Boolean(token);
    });
    if (!tokens.length) {
      return { firstName: '', surname: '' };
    }
    if (tokens.length === 1) {
      return { firstName: '', surname: tokens[0] };
    }
    return {
      surname: tokens[0],
      firstName: tokens.slice(1).join(' ')
    };
  },

  buildRoomFallbackModel_: function (requestedResourceId, dateString, errorMessage) {
    var roomConfig = this.getRoomConfig_();
    var user = ROOMS_APP.Auth.getUserContext();
    var effectiveNow = ROOMS_APP.Auth.getEffectiveNow(null, user);
    var simulation = ROOMS_APP.Auth.getSimulationContext_(null, user);
    var fallbackDate = dateString || (simulation.active ? simulation.dateIso : ROOMS_APP.toIsoDate(effectiveNow));
    return {
      ok: false,
      errorMessage: errorMessage || 'Dati aula non disponibili',
      date: fallbackDate,
      requestedResourceId: ROOMS_APP.normalizeString(requestedResourceId),
      resource: null,
      resources: [],
      bookings: [],
      userBookingsDay: [],
      timetableBookingsDay: [],
      upcomingBookings: [],
      userUpcomingBookings: [],
      timetableUpcomingBookings: [],
      ownBookings: [],
      upcomingEvents: [],
      isAulaMagna: false,
      freeSlots: [],
      slots: [],
      dailySlots: [],
      isOpen: false,
      openTime: roomConfig.openTime,
      closeTime: roomConfig.closeTime,
      openingWindow: {
        isOpen: false,
        source: 'UNKNOWN',
        openTime: '',
        closeTime: ''
      },
      status: 'UNKNOWN',
      statusSummary: 'Dati aula non disponibili',
      currentBooking: null,
      user: user,
      userDisplayName: user.displayName || [user.firstName || '', user.surname || ''].join(' ').trim(),
      simulation: {
        active: Boolean(simulation.active),
        simulatedNowISO: simulation.iso || ''
      },
      config: roomConfig
    };
  },

  getRoomViewModel: function (resourceId, dateString, options) {
    var startedAt = Date.now();
    var stepStartedAt = startedAt;
    var requestedResourceId = ROOMS_APP.normalizeString(resourceId || '');
    var user = ROOMS_APP.Auth.getUserContext();
    var effectiveNow = ROOMS_APP.Auth.getEffectiveNow(null, user);
    var simulation = ROOMS_APP.Auth.getSimulationContext_(null, user);
    var date = dateString || (simulation.active ? simulation.dateIso : ROOMS_APP.toIsoDate(effectiveNow));
    var viewOptions = options || {};

    try {
      var allResources = ROOMS_APP.Board.listResources_();
      var resourcesMs = Date.now() - stepStartedAt;
      if (!allResources.length) {
        var emptyModel = this.buildRoomFallbackModel_(requestedResourceId, date, 'Dati aula non disponibili');
        emptyModel.user = user;
        Logger.log('[PERF] Booking.getRoomViewModel total=%sms resources=%sms result=empty', Date.now() - startedAt, resourcesMs);
        return emptyModel;
      }

      var resource = requestedResourceId
        ? this.findResourceByAnyKey_(allResources, requestedResourceId)
        : allResources[0];
      if (requestedResourceId && !resource) {
        var notFoundModel = this.buildRoomFallbackModel_(requestedResourceId, date, 'Aula non trovata');
        notFoundModel.user = user;
        notFoundModel.resources = allResources;
        Logger.log('[PERF] Booking.getRoomViewModel total=%sms resources=%sms result=not_found', Date.now() - startedAt, resourcesMs);
        return notFoundModel;
      }

      var selectedResourceId = resource ? resource.ResourceId : '';
      stepStartedAt = Date.now();
      var userBookingsDay = resource ? this.enrichBookingsWithPermissions_(this.listBookingsForDay(selectedResourceId, date), user) : [];
      var bookingsMs = Date.now() - stepStartedAt;

      stepStartedAt = Date.now();
      var timetableBookingsDay = resource ? ROOMS_APP.Timetable.listOccupanciesForDate(selectedResourceId, date) : [];
      var timetableMs = Date.now() - stepStartedAt;

      var bookings = this.mergeRoomOccupancies_(userBookingsDay, timetableBookingsDay);

      stepStartedAt = Date.now();
      var openingWindow = resource ? ROOMS_APP.Policy.getEffectiveOpeningForResource(selectedResourceId, date) : {
        isOpen: false,
        source: 'UNKNOWN',
        openTime: '',
        closeTime: ''
      };
      var timeline = resource ? ROOMS_APP.Slots.getDaySlotsFromOccupancies(
        selectedResourceId,
        date,
        userBookingsDay,
        timetableBookingsDay,
        openingWindow,
        {
          splitFreeSlotsHalfHour: Boolean(viewOptions.splitFreeSlotsHalfHour)
        }
      ) : {
        bookings: [],
        timetableOccupancies: [],
        occupancies: [],
        freeSlots: [],
        slots: [],
        isOpen: false
      };
      var slotsMs = Date.now() - stepStartedAt;

      stepStartedAt = Date.now();
      var upcomingEvents = resource ? this.listUpcomingEventsForResource_(resource, date) : [];
      var isAulaMagna = this.isAulaMagnaResource_(resource);
      var eventsMs = Date.now() - stepStartedAt;

      stepStartedAt = Date.now();
      var currentTime = simulation.active ? simulation.time : Utilities.formatDate(effectiveNow, ROOMS_APP.getTimezone(), 'HH:mm');
      var currentDateIso = simulation.active ? simulation.dateIso : ROOMS_APP.toIsoDate(effectiveNow);
      var isCurrentDate = date === currentDateIso;
      timeline.freeSlots = ROOMS_APP.Slots.filterBookableFreeSlotsForDate_(
        timeline.freeSlots || [],
        date,
        currentDateIso,
        currentTime
      );
      var blockingBookings = bookings.filter(function (booking) {
        return ROOMS_APP.Timetable.isBlockingOccurrence(booking);
      });
      var current = isCurrentDate ? blockingBookings.filter(function (booking) {
        return booking.StartTime <= currentTime && booking.EndTime > currentTime;
      })[0] || null : null;
      var ownBookings = userBookingsDay.filter(function (booking) {
        return ROOMS_APP.normalizeEmail(booking.BookerEmail) === ROOMS_APP.normalizeEmail(user.email);
      });
      var statusSummary = '';
      if (!timeline.isOpen) {
        statusSummary = 'Aula non disponibile nel giorno selezionato';
      } else if (current) {
        statusSummary = 'Occupata ora da ' + (current.DisplayLabel || current.BookerSurname || current.BookerName || current.Title || 'N/D');
      } else if (!isCurrentDate) {
        statusSummary = 'Orario del giorno selezionato';
      } else {
        statusSummary = 'Nessuna occupazione in corso';
      }
      var composeMs = Date.now() - stepStartedAt;

      var model = {
        ok: true,
        errorMessage: '',
        date: date,
        requestedResourceId: requestedResourceId,
        resource: resource,
        bookings: bookings,
        userBookingsDay: userBookingsDay,
        timetableBookingsDay: timetableBookingsDay,
        upcomingBookings: [],
        userUpcomingBookings: [],
        timetableUpcomingBookings: [],
        ownBookings: ownBookings,
        upcomingEvents: upcomingEvents,
        isAulaMagna: isAulaMagna,
        freeSlots: timeline.freeSlots || [],
        slots: timeline.slots || [],
        dailySlots: timeline.slots || [],
        isOpen: timeline.isOpen,
        openTime: timeline.openTime || this.getRoomConfig_().openTime,
        closeTime: timeline.closeTime || this.getRoomConfig_().closeTime,
        openingWindow: openingWindow,
        status: current ? 'OCCUPIED' : 'FREE',
        statusSummary: statusSummary,
        currentBooking: current,
        user: user,
        userDisplayName: user.displayName || [user.firstName || '', user.surname || ''].join(' ').trim(),
        simulation: {
          active: Boolean(simulation.active),
          simulatedNowISO: simulation.iso || ''
        },
        resources: allResources,
        config: this.getRoomConfig_(),
        options: {
          splitFreeSlotsHalfHour: Boolean(viewOptions.splitFreeSlotsHalfHour)
        }
      };
      var dbStats = ROOMS_APP.DB.getRequestStats_();
      Logger.log(
        '[PERF] Booking.getRoomViewModel total=%sms resources=%sms bookings=%sms timetable=%sms slots=%sms events=%sms compose=%sms dailyBookings=%s freeSlots=%s splitFree=%s dbRequestHits=%s dbScriptHits=%s dbMisses=%s',
        Date.now() - startedAt,
        resourcesMs,
        bookingsMs,
        timetableMs,
        slotsMs,
        eventsMs,
        composeMs,
        userBookingsDay.length + timetableBookingsDay.length,
        (timeline.freeSlots || []).length,
        Boolean(viewOptions.splitFreeSlotsHalfHour),
        dbStats.requestHits,
        dbStats.scriptHits,
        dbStats.misses
      );
      return model;
    } catch (error) {
      var fallback = this.buildRoomFallbackModel_(requestedResourceId, date, 'Errore caricamento pagina aula');
      fallback.user = user;
      fallback.debugMessage = String(error && error.message ? error.message : error);
      Logger.log(
        '[PERF] Booking.getRoomViewModel total=%sms result=error error=%s',
        Date.now() - startedAt,
        fallback.debugMessage
      );
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

  isAulaMagnaResource_: function (resource) {
    if (!resource) {
      return false;
    }

    var byId = ROOMS_APP.normalizeString(resource.ResourceId).toUpperCase();
    if (byId === this.AULA_MAGNA_RESOURCE_ID_) {
      return true;
    }

    return ROOMS_APP.slugify(resource.DisplayName || '') === this.AULA_MAGNA_RESOURCE_ID_;
  },

  listUpcomingEventsForResource_: function (resource, fromDate) {
    if (!this.isAulaMagnaResource_(resource)) {
      return [];
    }

    var fromIsoDate = ROOMS_APP.toIsoDate(fromDate || new Date());
    var horizonDays = Math.max(0, ROOMS_APP.getNumberConfig('AULA_MAGNA_EVENT_DAYS_AHEAD', 14));
    var endIsoDate = this.addDaysToIsoDate_(fromIsoDate, horizonDays);
    var expectedResourceId = ROOMS_APP.normalizeString(resource.ResourceId);

    return ROOMS_APP.sortBy(
      ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.AULA_MAGNA_EVENTS).filter(function (row) {
        var eventResourceId = ROOMS_APP.normalizeString(row.ResourceId);
        var eventDate = ROOMS_APP.toIsoDate(row.EventDate);
        var isActive = ROOMS_APP.normalizeString(row.IsActive) === '' ? true : ROOMS_APP.asBoolean(row.IsActive);
        if (!isActive) {
          return false;
        }
        if (!eventDate || eventDate < fromIsoDate || eventDate > endIsoDate) {
          return false;
        }
        return ROOMS_APP.Timetable.matchesResourceId_(eventResourceId, expectedResourceId);
      }).map(function (row) {
        return {
          EventId: ROOMS_APP.normalizeString(row.EventId),
          ResourceId: ROOMS_APP.normalizeString(row.ResourceId),
          EventDate: ROOMS_APP.toIsoDate(row.EventDate),
          StartTime: ROOMS_APP.toTimeString(row.StartTime),
          EndTime: ROOMS_APP.toTimeString(row.EndTime),
          EventName: ROOMS_APP.normalizeString(row.EventName),
          IsActive: ROOMS_APP.normalizeString(row.IsActive) === '' ? 'TRUE' : String(row.IsActive),
          Notes: ROOMS_APP.normalizeString(row.Notes)
        };
      }),
      ['EventDate', 'StartTime', 'EndTime', 'EventName']
    );
  },

  addDaysToIsoDate_: function (isoDate, daysToAdd) {
    var base = ROOMS_APP.combineDateTime(isoDate, '00:00');
    base.setDate(base.getDate() + Number(daysToAdd || 0));
    return ROOMS_APP.toIsoDate(base);
  },

  getRoomConfig_: function () {
    return {
      bookingEnabled: ROOMS_APP.getBooleanConfig('BOOKING_ENABLED', true),
      allowRecurring: ROOMS_APP.getBooleanConfig('ALLOW_RECURRING', true),
      showBookerName: ROOMS_APP.getBooleanConfig('SHOW_BOOKER_NAME', false),
      slotMinutes: 60,
      freeSplitMinutes: 30,
      openTime: ROOMS_APP.getConfigValue('OPEN_TIME', '08:00'),
      closeTime: ROOMS_APP.getConfigValue('CLOSE_TIME', '18:00'),
      eventDaysAhead: ROOMS_APP.getNumberConfig('AULA_MAGNA_EVENT_DAYS_AHEAD', 14)
    };
  },

  getAulaMagnaEditorModel: function (resourceId, dateString) {
    var actor = ROOMS_APP.Auth.requireCanManageAulaMagna();
    var requestedResourceId = ROOMS_APP.normalizeString(resourceId || this.AULA_MAGNA_RESOURCE_ID_);
    var baseModel = this.getRoomViewModel(requestedResourceId, dateString || ROOMS_APP.toIsoDate(new Date()));
    var allRows = ROOMS_APP.sortBy(
      ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.AULA_MAGNA_EVENTS).filter(function (row) {
        var eventResourceId = ROOMS_APP.normalizeString(row.ResourceId);
        var isActive = ROOMS_APP.normalizeString(row.IsActive) === '' ? true : ROOMS_APP.asBoolean(row.IsActive);
        return isActive && ROOMS_APP.Timetable.matchesResourceId_(eventResourceId, requestedResourceId);
      }).map(function (row) {
        return {
          EventId: ROOMS_APP.normalizeString(row.EventId),
          ResourceId: ROOMS_APP.normalizeString(row.ResourceId),
          EventDate: ROOMS_APP.toIsoDate(row.EventDate),
          StartTime: ROOMS_APP.toTimeString(row.StartTime),
          EndTime: ROOMS_APP.toTimeString(row.EndTime),
          EventName: ROOMS_APP.normalizeString(row.EventName),
          Notes: ROOMS_APP.normalizeString(row.Notes),
          IsActive: ROOMS_APP.normalizeString(row.IsActive) === '' ? 'TRUE' : String(row.IsActive),
          CreatedAtISO: ROOMS_APP.normalizeString(row.CreatedAtISO),
          UpdatedAtISO: ROOMS_APP.normalizeString(row.UpdatedAtISO)
        };
      }),
      ['EventDate', 'StartTime', 'EndTime', 'EventName']
    );
    return {
      resourceId: baseModel && baseModel.resource ? baseModel.resource.ResourceId : requestedResourceId,
      displayName: baseModel && baseModel.resource ? baseModel.resource.DisplayName : 'AULA MAGNA',
      isAdmin: Boolean(actor.canAccessAdmin),
      isSuperAdmin: Boolean(actor.isSuperAdmin),
      canManageAulaMagna: Boolean(actor.canManageAulaMagna),
      events: safeCopy_(allRows)
    };

    function safeCopy_(rows) {
      return (rows || []).map(function (row) {
        var cloned = {};
        Object.keys(row || {}).forEach(function (key) {
          cloned[key] = row[key];
        });
        return cloned;
      });
    }
  },

  applyAulaMagnaEventChanges: function (resourceId, changes) {
    var actor = ROOMS_APP.Auth.requireCanManageAulaMagna();
    var nowIso = ROOMS_APP.toIsoDateTime(new Date());
    var targetResourceId = ROOMS_APP.normalizeString(resourceId || this.AULA_MAGNA_RESOURCE_ID_);
    var replaceRows = changes && Array.isArray(changes.replaceRows) ? changes.replaceRows : null;
    var normalizedChanges = this.normalizeAulaMagnaEventChanges_(changes);
    var self = this;

    return this.withBookingLock_(function () {
      var headers = ROOMS_APP.DB.getHeaders(ROOMS_APP.SHEET_NAMES.AULA_MAGNA_EVENTS);
      var rows = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.AULA_MAGNA_EVENTS);
      var created = 0;
      var updated = 0;
      var deleted = 0;

      if (replaceRows) {
        var keepRows = rows.filter(function (row) {
          return !ROOMS_APP.Timetable.matchesResourceId_(row.ResourceId, targetResourceId);
        });
        var rebuiltRows = replaceRows.map(function (payload) {
          return self.validateAulaMagnaEventRow_(
            targetResourceId,
            payload,
            ROOMS_APP.normalizeString(payload.EventId || payload.eventId || Utilities.getUuid()),
            ROOMS_APP.normalizeString(payload.CreatedAtISO || payload.createdAtISO || nowIso),
            nowIso
          );
        });
        ROOMS_APP.DB.replaceRows(ROOMS_APP.SHEET_NAMES.AULA_MAGNA_EVENTS, headers, keepRows.concat(rebuiltRows));
        self.writeAudit_('APPLY_AULA_MAGNA_EVENTS', '', '', targetResourceId, actor.email, 'OK', {
          mode: 'replace',
          rowCount: rebuiltRows.length
        });
        return self.getRoomViewModel(targetResourceId, ROOMS_APP.toIsoDate(new Date()));
      }

      normalizedChanges.deletes.forEach(function (eventId) {
        var index = self.findAulaMagnaEventIndexById_(rows, eventId);
        if (index < 0) {
          return;
        }
        rows.splice(index, 1);
        deleted += 1;
      });

      normalizedChanges.updates.forEach(function (entry) {
        var index = self.findAulaMagnaEventIndexById_(rows, entry.eventId);
        if (index < 0) {
          throw new Error('Evento non trovato: ' + entry.eventId);
        }
        var existing = rows[index];
        var next = self.validateAulaMagnaEventRow_(targetResourceId, entry.payload, existing.EventId, existing.CreatedAtISO || nowIso, nowIso);
        rows[index] = next;
        updated += 1;
      });

      normalizedChanges.creates.forEach(function (payload) {
        rows.push(self.validateAulaMagnaEventRow_(targetResourceId, payload, Utilities.getUuid(), nowIso, nowIso));
        created += 1;
      });

      ROOMS_APP.DB.replaceRows(ROOMS_APP.SHEET_NAMES.AULA_MAGNA_EVENTS, headers, rows);
      self.writeAudit_('APPLY_AULA_MAGNA_EVENTS', '', '', targetResourceId, actor.email, 'OK', {
        created: created,
        updated: updated,
        deleted: deleted
      });
      return self.getRoomViewModel(targetResourceId, ROOMS_APP.toIsoDate(new Date()));
    });
  },

  normalizeAulaMagnaEventChanges_: function (changes) {
    var source = changes || {};
    return {
      creates: (Array.isArray(source.creates) ? source.creates : []).map(function (entry) {
        return entry && typeof entry === 'object' ? entry : {};
      }),
      updates: (Array.isArray(source.updates) ? source.updates : []).map(function (entry) {
        return {
          eventId: ROOMS_APP.normalizeString(entry && entry.eventId),
          payload: entry && entry.payload && typeof entry.payload === 'object' ? entry.payload : {}
        };
      }).filter(function (entry) {
        return Boolean(entry.eventId);
      }),
      deletes: (Array.isArray(source.deletes) ? source.deletes : []).map(function (entry) {
        return ROOMS_APP.normalizeString(entry && entry.eventId ? entry.eventId : entry);
      }).filter(function (eventId) {
        return Boolean(eventId);
      })
    };
  },

  findAulaMagnaEventIndexById_: function (rows, eventId) {
    var normalized = ROOMS_APP.normalizeString(eventId);
    if (!normalized) {
      return -1;
    }
    var index;
    for (index = 0; index < (rows || []).length; index += 1) {
      if (ROOMS_APP.normalizeString(rows[index] && rows[index].EventId) === normalized) {
        return index;
      }
    }
    return -1;
  },

  validateAulaMagnaEventRow_: function (resourceId, payload, eventId, createdAtIso, updatedAtIso) {
    var eventDate = ROOMS_APP.toIsoDate(payload.EventDate || payload.eventDate);
    var startTime = ROOMS_APP.toTimeString(payload.StartTime || payload.startTime);
    var endTime = ROOMS_APP.toTimeString(payload.EndTime || payload.endTime);
    var eventName = ROOMS_APP.normalizeString(payload.EventName || payload.eventName);

    if (!eventDate || !/^\d{4}-\d{2}-\d{2}$/.test(eventDate)) {
      throw new Error('Data evento non valida.');
    }
    if (!startTime || !endTime || startTime >= endTime) {
      throw new Error('Orario evento non valido.');
    }
    if (!eventName) {
      throw new Error('Nome evento obbligatorio.');
    }

    return {
      EventId: ROOMS_APP.normalizeString(eventId),
      ResourceId: ROOMS_APP.normalizeString(resourceId || this.AULA_MAGNA_RESOURCE_ID_),
      EventDate: eventDate,
      StartTime: startTime,
      EndTime: endTime,
      EventName: eventName,
      IsActive: ROOMS_APP.normalizeString(payload.IsActive || payload.isActive) === '' ? 'TRUE' : String(payload.IsActive || payload.isActive),
      Notes: ROOMS_APP.normalizeString(payload.Notes || payload.notes),
      CreatedAtISO: ROOMS_APP.normalizeString(createdAtIso || updatedAtIso),
      UpdatedAtISO: ROOMS_APP.normalizeString(updatedAtIso)
    };
  },

  withBookingLock_: function (callback) {
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      return callback();
    } finally {
      lock.releaseLock();
    }
  },

  findBookingById_: function (bookingId) {
    var id = ROOMS_APP.normalizeString(bookingId);
    if (!id) {
      return null;
    }
    return ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.BOOKINGS).filter(function (row) {
      return row.BookingId === id;
    })[0] || null;
  },

  hasOnlyConflictError_: function (errors) {
    return Array.isArray(errors) && errors.length === 1 && String(errors[0]).indexOf('conflicts with an existing booking') >= 0;
  },

  isConflictError_: function (errorMessage) {
    return String(errorMessage || '').indexOf('conflicts with an existing booking') >= 0;
  },

  validateAgainstWorkingRows_: function (request, actor, workingRows, ignoreBookingId, ignoredTimetableBookingIds) {
    var validation = ROOMS_APP.Policy.validateBookingRequest(request, actor);
    var nonConflictErrors = (validation.errors || []).filter(function (errorMessage) {
      return !ROOMS_APP.Booking.isConflictError_(errorMessage);
    });
    if (nonConflictErrors.length) {
      throw new Error(nonConflictErrors.join(' '));
    }

    var normalized = validation.normalized || {};
    var hasWorkingRowConflict = this.hasConflictInRows_(
      workingRows,
      normalized.resourceId,
      normalized.bookingDate,
      normalized.startTime,
      normalized.endTime,
      ignoreBookingId
    );
    var hasTimetableConflict = this.hasTimetableConflict_(
      normalized.resourceId,
      normalized.bookingDate,
      normalized.startTime,
      normalized.endTime,
      ignoredTimetableBookingIds
    );

    if (hasWorkingRowConflict || hasTimetableConflict) {
      throw new Error(this.SLOT_TAKEN_MESSAGE_);
    }

    return {
      ok: true,
      errors: [],
      normalized: normalized,
      resource: validation.resource,
      actor: actor
    };
  },

  hasConflictInRows_: function (rows, resourceId, bookingDate, startTime, endTime, ignoreBookingId) {
    return (rows || []).some(function (row) {
      if (!row || ROOMS_APP.Booking.isBookingCancelled_(row)) {
        return false;
      }
      if (row.ResourceId !== resourceId || row.BookingDate !== bookingDate) {
        return false;
      }
      if (ignoreBookingId && row.BookingId === ignoreBookingId) {
        return false;
      }
      return !(endTime <= row.StartTime || startTime >= row.EndTime);
    });
  },

  hasTimetableConflict_: function (resourceId, bookingDate, startTime, endTime, ignoredTimetableBookingIds) {
    var ignored = ignoredTimetableBookingIds || {};
    return ROOMS_APP.Timetable.listOccupanciesForDate(resourceId, bookingDate).some(function (occupancy) {
      if (!ROOMS_APP.Timetable.isBlockingOccurrence(occupancy)) {
        return false;
      }
      var bookingId = ROOMS_APP.normalizeString(occupancy.BookingId);
      if (bookingId && ignored[bookingId]) {
        return false;
      }
      return !(endTime <= occupancy.StartTime || startTime >= occupancy.EndTime);
    });
  },

  resolveTimetableOccurrenceByBookingId_: function (resourceId, bookingDate, bookingId) {
    var normalizedBookingId = ROOMS_APP.normalizeString(bookingId);
    if (!normalizedBookingId) {
      return null;
    }
    return ROOMS_APP.Timetable.listOccupanciesForDate(resourceId, bookingDate).filter(function (row) {
      return ROOMS_APP.normalizeString(row.BookingId) === normalizedBookingId;
    })[0] || null;
  },

  buildTimetableDisableRuleValue_: function (occurrence) {
    return [
      ROOMS_APP.normalizeString(occurrence && occurrence.OccupancyId),
      ROOMS_APP.normalizeString(occurrence && occurrence.StartTime),
      ROOMS_APP.normalizeString(occurrence && occurrence.EndTime)
    ].join('|');
  },

  buildTimetableDisableOverrideRow_: function (occurrence, resourceId, bookingDate, notes, nowIso) {
    var normalizedResource = ROOMS_APP.normalizeString(resourceId);
    var normalizedDate = ROOMS_APP.toIsoDate(bookingDate || new Date());
    var ruleValue = this.buildTimetableDisableRuleValue_(occurrence);
    var overrideId = 'OVR_TT_' + ROOMS_APP.slugify([
      normalizedResource,
      normalizedDate,
      ruleValue
    ].join('|'));
    return {
      OverrideId: overrideId,
      ResourceId: normalizedResource,
      BookingDate: normalizedDate,
      StartTime: ROOMS_APP.normalizeString(occurrence && occurrence.StartTime),
      EndTime: ROOMS_APP.normalizeString(occurrence && occurrence.EndTime),
      RuleKey: 'TIMETABLE_DISABLED',
      RuleValue: ruleValue,
      IsEnabled: 'TRUE',
      Notes: ROOMS_APP.normalizeString(notes || ('Auto ' + nowIso))
    };
  },

  upsertOverrideRow_: function (rows, nextRow) {
    var output = (rows || []).slice();
    var targetId = ROOMS_APP.normalizeString(nextRow && nextRow.OverrideId);
    var index;
    for (index = 0; index < output.length; index += 1) {
      if (ROOMS_APP.normalizeString(output[index] && output[index].OverrideId) === targetId) {
        output[index] = nextRow;
        return output;
      }
    }
    output.push(nextRow);
    return output;
  },

  findBookingIndexById_: function (rows, bookingId) {
    var id = ROOMS_APP.normalizeString(bookingId);
    if (!id) {
      return -1;
    }
    var index;
    for (index = 0; index < (rows || []).length; index += 1) {
      if (rows[index] && rows[index].BookingId === id) {
        return index;
      }
    }
    return -1;
  },

  normalizeBatchChanges_: function (changes) {
    var source = changes || {};
    var normalized = {
      creates: [],
      updates: [],
      deletes: [],
      timetableDeletes: [],
      timetableUpdates: []
    };

    (Array.isArray(source.creates) ? source.creates : []).forEach(function (entry) {
      normalized.creates.push({
        payload: entry && typeof entry === 'object' ? entry : {}
      });
    });
    (Array.isArray(source.updates) ? source.updates : []).forEach(function (entry) {
      normalized.updates.push({
        bookingId: ROOMS_APP.normalizeString(entry && entry.bookingId),
        payload: entry && entry.payload && typeof entry.payload === 'object' ? entry.payload : {}
      });
    });
    (Array.isArray(source.deletes) ? source.deletes : []).forEach(function (entry) {
      normalized.deletes.push({
        bookingId: ROOMS_APP.normalizeString(entry && entry.bookingId),
        notes: entry && Object.prototype.hasOwnProperty.call(entry, 'notes') ? entry.notes : ''
      });
    });
    (Array.isArray(source.timetableDeletes) ? source.timetableDeletes : []).forEach(function (entry) {
      normalized.timetableDeletes.push({
        bookingId: ROOMS_APP.normalizeString(entry && entry.bookingId),
        notes: entry && Object.prototype.hasOwnProperty.call(entry, 'notes') ? entry.notes : ''
      });
    });
    (Array.isArray(source.timetableUpdates) ? source.timetableUpdates : []).forEach(function (entry) {
      normalized.timetableUpdates.push({
        bookingId: ROOMS_APP.normalizeString(entry && entry.bookingId),
        payload: entry && entry.payload && typeof entry.payload === 'object' ? entry.payload : {}
      });
    });

    return normalized;
  },

  buildBookingFromValidation_: function (bookingId, seriesId, validation, actor, nowIso, existingBooking) {
    var normalized = validation.normalized;
    var identity = ROOMS_APP.extractIdentityFromEmail(actor.email);
    var current = existingBooking || {};
    var requesterMode = this.normalizeRequesterMode_(normalized.requesterMode || current.RequesterMode);
    var effectiveName = requesterMode === 'NONE' ? '' : (normalized.bookerName || current.BookerName || identity.firstName);
    var effectiveSurname = requesterMode === 'NONE' ? '' : (normalized.bookerSurname || current.BookerSurname || identity.surname);
    var createdAt = existingBooking ? (existingBooking.CreatedAtISO || nowIso) : nowIso;

    return {
      BookingId: bookingId,
      SeriesId: seriesId || '',
      ResourceId: normalized.resourceId,
      BookingDate: normalized.bookingDate,
      StartTime: normalized.startTime,
      EndTime: normalized.endTime,
      StartISO: ROOMS_APP.toIsoDateTime(ROOMS_APP.combineDateTime(normalized.bookingDate, normalized.startTime)),
      EndISO: ROOMS_APP.toIsoDateTime(ROOMS_APP.combineDateTime(normalized.bookingDate, normalized.endTime)),
      Title: normalized.title || validation.resource.DisplayName,
      ActivityDescription: normalized.activityDescription,
      DisplayMode: requesterMode === 'NONE' ? 'ACTIVITY' : this.normalizeDisplayMode_(normalized.displayMode || current.DisplayMode),
      RequesterMode: requesterMode,
      RequesterManualName: requesterMode === 'MANUAL'
        ? ROOMS_APP.normalizeString(normalized.requesterManualName || current.RequesterManualName)
        : '',
      BookerEmail: current.BookerEmail || actor.email,
      BookerName: effectiveName,
      BookerSurname: effectiveSurname,
      Status: 'CONFIRMED',
      CreatedAtISO: createdAt,
      UpdatedAtISO: nowIso,
      CancelledAtISO: '',
      Notes: normalized.notes
    };
  },

  enrichBookingsWithPermissions_: function (bookings, user) {
    return (bookings || []).map(function (booking) {
      var enriched = {};
      Object.keys(booking || {}).forEach(function (key) {
        enriched[key] = booking[key];
      });
      enriched.DisplayMode = ROOMS_APP.Booking.normalizeDisplayMode_(booking && booking.DisplayMode);
      enriched.CanManage = ROOMS_APP.Auth.canManageBooking(booking, user);
      enriched.SourceKind = 'USER_BOOKING';
      enriched.SourceType = 'USER_BOOKING';
      enriched.IsReadOnly = false;
      enriched.DisplayLabel = ROOMS_APP.Booking.getBookingDisplayLabel_(enriched);
      return enriched;
    });
  },

  mergeRoomOccupancies_: function (userBookings, timetableBookings) {
    return ROOMS_APP.sortBy((userBookings || []).concat(timetableBookings || []), [
      'BookingDate',
      'StartTime',
      'EndTime',
      'SourceKind',
      'BookingId'
    ]);
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
