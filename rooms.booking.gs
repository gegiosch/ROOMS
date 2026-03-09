var ROOMS_APP = ROOMS_APP || {};

ROOMS_APP.Booking = {
  SLOT_TAKEN_MESSAGE_: 'Lo slot non è più disponibile. Aggiorna la disponibilità dell’aula.',

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

  listUpcomingBookingsForRoom: function (resourceId, fromDate) {
    var startDate = fromDate || ROOMS_APP.toIsoDate(new Date());
    ROOMS_APP.Schema.ensureAll();
    return ROOMS_APP.sortBy(
      ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.BOOKINGS).filter(function (row) {
        return row.ResourceId === resourceId &&
          row.BookingDate >= startDate &&
          row.Status !== 'CANCELLED';
      }),
      ['BookingDate', 'StartTime', 'EndTime']
    );
  },

  createBooking: function (payload, options) {
    payload = payload || {};
    var actor = (options && options.actor) || ROOMS_APP.Auth.getUserContext();
    ROOMS_APP.Auth.assertAllowedDomain(actor.email);
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
    var existing = this.findBookingById_(bookingId);
    if (!existing) {
      throw new Error('Booking not found.');
    }
    if (!ROOMS_APP.Auth.canManageBooking(existing, actor)) {
      throw new Error('Only the creator or an admin can modify this booking.');
    }

    var request = {
      bookingId: existing.BookingId,
      seriesId: existing.SeriesId,
      resourceId: Object.prototype.hasOwnProperty.call(payload, 'resourceId') ? payload.resourceId : existing.ResourceId,
      bookingDate: Object.prototype.hasOwnProperty.call(payload, 'bookingDate') ? payload.bookingDate : existing.BookingDate,
      startTime: Object.prototype.hasOwnProperty.call(payload, 'startTime') ? payload.startTime : existing.StartTime,
      endTime: Object.prototype.hasOwnProperty.call(payload, 'endTime') ? payload.endTime : existing.EndTime,
      title: Object.prototype.hasOwnProperty.call(payload, 'title') ? payload.title : existing.Title,
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
        throw new Error('Only the creator or an admin can modify this booking.');
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
    var self = this;
    var booking = this.withBookingLock_(function () {
      var existing = self.findBookingById_(bookingId);
      if (!existing) {
        throw new Error('Booking not found.');
      }

      if (!ROOMS_APP.Auth.canManageBooking(existing, actor)) {
        throw new Error('Only the creator or an admin can cancel this booking.');
      }

      if (existing.Status === 'CANCELLED') {
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
    if (!normalizedChanges.creates.length && !normalizedChanges.updates.length && !normalizedChanges.deletes.length) {
      return this.getRoomViewModel(targetResourceId, targetDate);
    }
    var self = this;
    var result;
    try {
      result = this.withBookingLock_(function () {
        var workingRows = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.BOOKINGS);
        var nowIso = ROOMS_APP.toIsoDateTime(new Date());
        var createdIds = [];
        var updatedIds = [];
        var deletedIds = [];

        normalizedChanges.deletes.forEach(function (entry) {
          var bookingId = ROOMS_APP.normalizeString(entry.bookingId);
          var existingIndex = self.findBookingIndexById_(workingRows, bookingId);
          if (existingIndex < 0) {
            throw new Error('Booking not found.');
          }
          var existing = workingRows[existingIndex];
          if (existing.ResourceId !== targetResourceId) {
            throw new Error('Cannot modify bookings for a different room.');
          }
          if (!ROOMS_APP.Auth.canManageBooking(existing, actor)) {
            throw new Error('Only the creator or an admin can cancel this booking.');
          }
          if (existing.Status === 'CANCELLED') {
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
          var existingIndex = self.findBookingIndexById_(workingRows, bookingId);
          if (existingIndex < 0) {
            throw new Error('Booking not found.');
          }
          var existing = workingRows[existingIndex];
          if (existing.Status === 'CANCELLED') {
            throw new Error('Booking not found.');
          }
          if (existing.ResourceId !== targetResourceId) {
            throw new Error('Cannot modify bookings for a different room.');
          }
          if (!ROOMS_APP.Auth.canManageBooking(existing, actor)) {
            throw new Error('Only the creator or an admin can modify this booking.');
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
            notes: Object.prototype.hasOwnProperty.call(payload, 'notes') ? payload.notes : existing.Notes,
            bookerName: Object.prototype.hasOwnProperty.call(payload, 'bookerName') ? payload.bookerName : existing.BookerName,
            bookerSurname: Object.prototype.hasOwnProperty.call(payload, 'bookerSurname') ? payload.bookerSurname : existing.BookerSurname
          };

          if (ROOMS_APP.normalizeString(request.resourceId) !== targetResourceId) {
            throw new Error('Cannot move booking to a different room from this panel.');
          }

          var validation = self.validateAgainstWorkingRows_(request, actor, workingRows, existing.BookingId);
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
            notes: payload.notes,
            bookerName: payload.bookerName,
            bookerSurname: payload.bookerSurname
          };
          var validation = self.validateAgainstWorkingRows_(request, actor, workingRows, '');
          var bookingId = validation.normalized.bookingId || Utilities.getUuid();
          var seriesId = validation.normalized.seriesId || '';
          var booking = self.buildBookingFromValidation_(bookingId, seriesId, validation, actor, nowIso, null);
          workingRows.push(booking);
          createdIds.push(bookingId);
        });

        var headers = ROOMS_APP.DB.getHeaders(ROOMS_APP.SHEET_NAMES.BOOKINGS);
        ROOMS_APP.DB.replaceRows(ROOMS_APP.SHEET_NAMES.BOOKINGS, headers, workingRows);
        return {
          createdIds: createdIds,
          updatedIds: updatedIds,
          deletedIds: deletedIds
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
      createdIds: result.createdIds,
      updatedIds: result.updatedIds,
      deletedIds: result.deletedIds
    });

    return this.getRoomViewModel(targetResourceId, targetDate);
  },

  buildRoomFallbackModel_: function (requestedResourceId, dateString, errorMessage) {
    var roomConfig = this.getRoomConfig_();
    return {
      ok: false,
      errorMessage: errorMessage || 'Dati aula non disponibili',
      date: dateString || ROOMS_APP.toIsoDate(new Date()),
      requestedResourceId: ROOMS_APP.normalizeString(requestedResourceId),
      resource: null,
      resources: [],
      bookings: [],
      upcomingBookings: [],
      ownBookings: [],
      freeSlots: [],
      slots: [],
      isOpen: false,
      openTime: roomConfig.openTime,
      closeTime: roomConfig.closeTime,
      status: 'UNKNOWN',
      currentBooking: null,
      user: ROOMS_APP.Auth.getUserContext(),
      config: roomConfig
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
      var bookings = resource ? this.enrichBookingsWithPermissions_(this.listBookingsForDay(selectedResourceId, date), user) : [];
      var upcomingBookings = resource ? this.enrichBookingsWithPermissions_(this.listUpcomingBookingsForRoom(selectedResourceId, date), user) : [];
      var timeline = resource ? ROOMS_APP.Slots.getDaySlots(selectedResourceId, date) : { bookings: [], freeSlots: [], slots: [], isOpen: false };
      var now = new Date();
      var currentTime = Utilities.formatDate(now, ROOMS_APP.getTimezone(), 'HH:mm');
      var current = bookings.filter(function (booking) {
        return booking.StartTime <= currentTime && booking.EndTime > currentTime;
      })[0] || null;
      var ownBookings = upcomingBookings.filter(function (booking) {
        return ROOMS_APP.normalizeEmail(booking.BookerEmail) === ROOMS_APP.normalizeEmail(user.email);
      });

      return {
        ok: true,
        errorMessage: '',
        date: date,
        requestedResourceId: requestedResourceId,
        resource: resource,
        bookings: bookings,
        upcomingBookings: upcomingBookings,
        ownBookings: ownBookings,
        freeSlots: timeline.freeSlots || [],
        slots: timeline.slots || [],
        isOpen: timeline.isOpen,
        openTime: timeline.openTime || this.getRoomConfig_().openTime,
        closeTime: timeline.closeTime || this.getRoomConfig_().closeTime,
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
      showBookerName: ROOMS_APP.getBooleanConfig('SHOW_BOOKER_NAME', false),
      slotMinutes: ROOMS_APP.getNumberConfig('SLOT_MINUTES', 30),
      openTime: ROOMS_APP.getConfigValue('OPEN_TIME', '08:00'),
      closeTime: ROOMS_APP.getConfigValue('CLOSE_TIME', '18:00')
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

  validateAgainstWorkingRows_: function (request, actor, workingRows, ignoreBookingId) {
    var validation = ROOMS_APP.Policy.validateBookingRequest(request, actor);
    var nonConflictErrors = (validation.errors || []).filter(function (errorMessage) {
      return !ROOMS_APP.Booking.isConflictError_(errorMessage);
    });
    if (nonConflictErrors.length) {
      throw new Error(nonConflictErrors.join(' '));
    }

    var normalized = validation.normalized || {};
    if (this.hasConflictInRows_(workingRows, normalized.resourceId, normalized.bookingDate, normalized.startTime, normalized.endTime, ignoreBookingId)) {
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
      if (!row || row.Status === 'CANCELLED') {
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
      deletes: []
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

    return normalized;
  },

  buildBookingFromValidation_: function (bookingId, seriesId, validation, actor, nowIso, existingBooking) {
    var normalized = validation.normalized;
    var identity = ROOMS_APP.extractIdentityFromEmail(actor.email);
    var current = existingBooking || {};
    var effectiveName = normalized.bookerName || current.BookerName || identity.firstName;
    var effectiveSurname = normalized.bookerSurname || current.BookerSurname || identity.surname;
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
      enriched.CanManage = ROOMS_APP.Auth.canManageBooking(booking, user);
      return enriched;
    });
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
