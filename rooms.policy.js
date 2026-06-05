var ROOMS_APP = ROOMS_APP || {};

ROOMS_APP.Policy = {
  normalizeBookingDisplayMode_: function (value) {
    return String(value || '').toUpperCase() === 'ACTIVITY' ? 'ACTIVITY' : 'TEACHER';
  },

  normalizeBookingRequesterMode_: function (value) {
    var normalized = String(value || '').toUpperCase();
    if (normalized === 'MANUAL' || normalized === 'NONE') {
      return normalized;
    }
    return 'LIST';
  },

  parseRequesterName_: function (value) {
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

  getResource: function (resourceId) {
    var resources = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.RESOURCES);
    return resources.filter(function (row) {
      return row.ResourceId === resourceId;
    })[0] || null;
  },

  getHoliday: function (dateString) {
    return ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.HOLIDAYS).filter(function (row) {
      return row.HolidayDate === dateString && ROOMS_APP.asBoolean(row.IsBlocked);
    })[0] || null;
  },

  getWeekSchedule: function (weekdayName) {
    return ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.WEEK_SCHEDULE).filter(function (row) {
      return row.Weekday === weekdayName;
    })[0] || null;
  },

  getSpecialOpening: function (dateString) {
    return ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.SPECIAL_OPENINGS).filter(function (row) {
      return row.Date === dateString && ROOMS_APP.asBoolean(row.IsEnabled);
    })[0] || null;
  },

  getClosuresForDate: function (dateString) {
    return ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.CLOSURES).filter(function (row) {
      return ROOMS_APP.asBoolean(row.IsBlocked) && row.StartDate <= dateString && row.EndDate >= dateString;
    });
  },

  getLessonOverrideForDate: function (dateString) {
    var targetDate = ROOMS_APP.toIsoDate(dateString || '');
    if (!targetDate || !ROOMS_APP.SHEET_NAMES.LESSON_OVERRIDES) {
      return null;
    }
    return ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.LESSON_OVERRIDES).filter(function (row) {
      return ROOMS_APP.toIsoDate(row.Date) === targetDate && ROOMS_APP.asBoolean(row.IsEnabled);
    })[0] || null;
  },

  getLessonDayPolicy: function (dateString) {
    var targetDate = ROOMS_APP.toIsoDate(dateString || '');
    var closures = this.getClosuresForDate(targetDate);
    var fullDayClosure = closures.filter(function (closure) {
      return ROOMS_APP.Policy.isFullDayClosure_(closure);
    })[0] || null;
    var override = this.getLessonOverrideForDate(targetDate);
    var validFrom = this.normalizeConfigDate_(ROOMS_APP.getConfigValue('LESSONS_VALID_FROM', ''));
    var validTo = this.normalizeConfigDate_(ROOMS_APP.getConfigValue('LESSONS_VALID_TO', ''));
    var insideValidity = (!validFrom || targetDate >= validFrom) && (!validTo || targetDate <= validTo);
    // Calendar precedence: closures, lesson overrides, lesson validity window, then standard timetable.
    if (fullDayClosure) {
      return {
        date: targetDate,
        source: 'CLOSURE',
        lessonsEnabled: false,
        closureRows: closures,
        fullDayClosure: fullDayClosure,
        label: fullDayClosure.Label || 'Chiusura'
      };
    }
    if (override) {
      return {
        date: targetDate,
        source: 'LESSON_OVERRIDE',
        lessonsEnabled: ROOMS_APP.asBoolean(override.LessonsEnabled),
        openTime: this.normalizeOptionalTime_(override.OpenTime),
        closeTime: this.normalizeOptionalTime_(override.CloseTime),
        lastLessonPeriod: this.normalizeOptionalPositiveNumber_(override.LastLessonPeriod),
        roomsBookable: this.normalizeOptionalBoolean_(override.RoomsBookable),
        labsBookable: this.normalizeOptionalBoolean_(override.LabsBookable),
        mode: ROOMS_APP.normalizeString(override.Mode),
        label: ROOMS_APP.normalizeString(override.Label),
        override: override,
        closureRows: closures
      };
    }
    return {
      date: targetDate,
      source: insideValidity ? 'LESSON_VALIDITY' : 'LESSON_VALIDITY_OUTSIDE',
      lessonsEnabled: insideValidity,
      validFrom: validFrom,
      validTo: validTo,
      closureRows: closures
    };
  },

  getDailyOpening: function (dateString) {
    var holiday = this.getHoliday(dateString);
    if (holiday) {
      return {
        isOpen: false,
        source: 'HOLIDAY',
        label: holiday.Label
      };
    }

    var closures = this.getClosuresForDate(dateString);
    var fullDayClosure = closures.filter(function (closure) {
      return ROOMS_APP.Policy.isFullDayClosure_(closure);
    })[0] || null;
    if (fullDayClosure) {
      return {
        isOpen: false,
        source: 'CLOSURE',
        label: ROOMS_APP.normalizeString(fullDayClosure.Label) || 'Chiusura straordinaria'
      };
    }

    var opening = this.getSpecialOpening(dateString);
    if (opening) {
      return {
        isOpen: true,
        openTime: opening.OpenTime || ROOMS_APP.getConfigValue('OPEN_TIME', '08:00'),
        closeTime: opening.CloseTime || ROOMS_APP.getConfigValue('CLOSE_TIME', '18:00'),
        source: 'SPECIAL_OPENING',
        label: opening.Label
      };
    }

    var weekdayName = ROOMS_APP.getWeekdayName(dateString);
    var weekSchedule = this.getWeekSchedule(weekdayName);
    if (!weekSchedule || !ROOMS_APP.asBoolean(weekSchedule.IsWorkingDay)) {
      return {
        isOpen: false,
        source: 'WEEK_SCHEDULE',
        label: weekdayName
      };
    }

    return {
      isOpen: true,
      openTime: weekSchedule.OpenTime || ROOMS_APP.getConfigValue('OPEN_TIME', '08:00'),
      closeTime: weekSchedule.CloseTime || ROOMS_APP.getConfigValue('CLOSE_TIME', '18:00'),
      source: 'WEEK_SCHEDULE',
      label: weekdayName
    };
  },

  getEffectiveOpeningForResource: function (resourceId, dateString) {
    var base = this.getDailyOpening(dateString);
    if (!base || !base.isOpen) {
      return {
        isOpen: false,
        source: base ? base.source : 'UNKNOWN',
        label: base ? base.label : 'N/D',
        openTime: '',
        closeTime: '',
        baseOpenTime: '',
        baseCloseTime: '',
        resourceOpenTime: '',
        resourceCloseTime: ''
      };
    }

    var resource = this.getResource(resourceId) || {};
    var lessonPolicy = this.getLessonDayPolicy(dateString);
    var lessonBookable = this.getLessonOverrideBookabilityForResource_(lessonPolicy, resource);
    var resourceOpen = this.normalizeOptionalTime_(resource.OpenTime);
    var resourceClose = this.normalizeOptionalTime_(resource.CloseTime);
    var finalOpen = base.openTime;
    var finalClose = base.closeTime;

    if (lessonBookable === false) {
      return {
        isOpen: false,
        source: 'LESSON_OVERRIDE',
        label: lessonPolicy.label || 'Giorno non prenotabile',
        openTime: '',
        closeTime: '',
        baseOpenTime: base.openTime,
        baseCloseTime: base.closeTime,
        resourceOpenTime: resourceOpen,
        resourceCloseTime: resourceClose
      };
    }

    if (lessonPolicy && lessonPolicy.source === 'LESSON_OVERRIDE') {
      if (lessonPolicy.openTime && lessonPolicy.openTime > finalOpen) {
        finalOpen = lessonPolicy.openTime;
      }
      if (lessonPolicy.closeTime && lessonPolicy.closeTime < finalClose) {
        finalClose = lessonPolicy.closeTime;
      }
    }

    if (resourceOpen && resourceOpen > finalOpen) {
      finalOpen = resourceOpen;
    }
    if (resourceClose && resourceClose < finalClose) {
      finalClose = resourceClose;
    }

    if (!finalOpen || !finalClose || finalOpen >= finalClose) {
      return {
        isOpen: false,
        source: 'RESOURCE_RESTRICTION',
        label: 'Finestra locale aula non disponibile',
        openTime: '',
        closeTime: '',
        baseOpenTime: base.openTime,
        baseCloseTime: base.closeTime,
        resourceOpenTime: resourceOpen,
        resourceCloseTime: resourceClose
      };
    }

    return {
      isOpen: true,
      source: base.source,
      label: base.label,
      openTime: finalOpen,
      closeTime: finalClose,
      baseOpenTime: base.openTime,
      baseCloseTime: base.closeTime,
      resourceOpenTime: resourceOpen,
      resourceCloseTime: resourceClose
    };
  },

  normalizeOptionalTime_: function (value) {
    var normalized = ROOMS_APP.normalizeString(value);
    if (!normalized) {
      return '';
    }
    return /^\d{2}:\d{2}$/.test(normalized) ? normalized : '';
  },

  normalizeConfigDate_: function (value) {
    var normalized = ROOMS_APP.normalizeString(value);
    return /^\d{4}-\d{2}-\d{2}$/.test(normalized) ? normalized : '';
  },

  normalizeOptionalPositiveNumber_: function (value) {
    var normalized = ROOMS_APP.normalizeString(value);
    var parsed;
    if (!normalized) {
      return 0;
    }
    parsed = Number(normalized);
    return isFinite(parsed) && parsed > 0 ? parsed : 0;
  },

  normalizeOptionalBoolean_: function (value) {
    var normalized = ROOMS_APP.normalizeString(value).toUpperCase();
    if (!normalized) {
      return null;
    }
    return normalized === 'TRUE';
  },

  isFullDayClosure_: function (closure) {
    var start = this.normalizeOptionalTime_(closure && closure.StartTime) || '00:00';
    var end = this.normalizeOptionalTime_(closure && closure.EndTime) || '23:59';
    return start <= '00:00' && end >= '23:59';
  },

  getLessonOverrideBookabilityForResource_: function (lessonPolicy, resource) {
    var kind = this.getResourceCalendarKind_(resource);
    if (!lessonPolicy || lessonPolicy.source !== 'LESSON_OVERRIDE') {
      return null;
    }
    if (kind === 'LAB') {
      return lessonPolicy.labsBookable;
    }
    return lessonPolicy.roomsBookable;
  },

  getResourceCalendarKind_: function (resource) {
    var areaCode = ROOMS_APP.normalizeString(resource && resource.AreaCode).toUpperCase();
    var areaLabel = ROOMS_APP.normalizeString(resource && resource.AreaLabel).toUpperCase();
    var sideLabel = ROOMS_APP.normalizeString(resource && resource.SideLabel).toUpperCase();
    var displayName = ROOMS_APP.normalizeString(resource && resource.DisplayName).toUpperCase();
    if (areaCode === 'LAB' ||
        areaLabel.indexOf('LABORATOR') >= 0 ||
        sideLabel.indexOf('LABORATOR') >= 0 ||
        displayName.indexOf('LAB') >= 0) {
      return 'LAB';
    }
    return 'ROOM';
  },

  findBlockingClosure: function (dateString, startTime, endTime) {
    var closures = this.getClosuresForDate(dateString);

    return closures.filter(function (closure) {
      var closureStart = closure.StartTime || '00:00';
      var closureEnd = closure.EndTime || '23:59';
      return !(endTime <= closureStart || startTime >= closureEnd);
    })[0] || null;
  },

  listClosureOccupanciesForDate: function (resourceId, dateString) {
    var normalizedResourceId = ROOMS_APP.normalizeString(resourceId);
    var targetDate = ROOMS_APP.toIsoDate(dateString || '');
    return this.getClosuresForDate(targetDate).map(function (closure) {
      var startTime = ROOMS_APP.Policy.normalizeOptionalTime_(closure.StartTime) || '00:00';
      var endTime = ROOMS_APP.Policy.normalizeOptionalTime_(closure.EndTime) || '23:59';
      return {
        BookingId: 'CLOSURE-' + targetDate + '-' + startTime + '-' + endTime,
        ResourceId: normalizedResourceId,
        BookingDate: targetDate,
        StartTime: startTime,
        EndTime: endTime,
        Title: closure.Label || 'Chiusura',
        DisplayLabel: closure.Label || 'Chiusura',
        SourceKind: 'CLOSURE',
        SourceType: 'CLOSURE',
        IsNonBlocking: 'FALSE',
        Notes: closure.Notes || ''
      };
    });
  },

  hasConflict: function (resourceId, dateString, startTime, endTime, ignoreBookingId) {
    var bookingConflict = ROOMS_APP.Booking.listBookingsForDay(resourceId, dateString).some(function (booking) {
      if (ignoreBookingId && booking.BookingId === ignoreBookingId) {
        return false;
      }
      return !(endTime <= booking.StartTime || startTime >= booking.EndTime);
    });
    if (bookingConflict) {
      return true;
    }

    return ROOMS_APP.Timetable.listOccupanciesForDate(resourceId, dateString).some(function (occupancy) {
      if (!ROOMS_APP.Timetable.isBlockingOccurrence(occupancy)) {
        return false;
      }
      return !(endTime <= occupancy.StartTime || startTime >= occupancy.EndTime);
    });
  },

  validateBookingRequest: function (request, actor) {
    var user = actor || ROOMS_APP.Auth.getUserContext();
    var errors = [];
    var emailIdentity = ROOMS_APP.extractIdentityFromEmail(user.email);
    var normalized = {
      resourceId: ROOMS_APP.normalizeString(request.resourceId || request.ResourceId),
      bookingDate: ROOMS_APP.toIsoDate(request.bookingDate || request.BookingDate),
      startTime: ROOMS_APP.toTimeString(request.startTime || request.StartTime),
      endTime: ROOMS_APP.toTimeString(request.endTime || request.EndTime),
      title: ROOMS_APP.normalizeString(request.title || request.Title),
      activityDescription: ROOMS_APP.normalizeString(request.activityDescription || request.ActivityDescription),
      displayMode: this.normalizeBookingDisplayMode_(request.displayMode || request.DisplayMode),
      requesterMode: this.normalizeBookingRequesterMode_(request.requesterMode || request.RequesterMode),
      requesterManualName: ROOMS_APP.normalizeString(request.requesterManualName || request.RequesterManualName),
      notes: ROOMS_APP.normalizeString(request.notes || request.Notes),
      bookerName: ROOMS_APP.normalizeString(request.bookerName || request.BookerName),
      bookerSurname: ROOMS_APP.normalizeString(request.bookerSurname || request.BookerSurname),
      seriesId: ROOMS_APP.normalizeString(request.seriesId || request.SeriesId),
      bookingId: ROOMS_APP.normalizeString(request.bookingId || request.BookingId)
    };

    if (normalized.requesterMode === 'NONE') {
      normalized.displayMode = 'ACTIVITY';
      normalized.bookerName = '';
      normalized.bookerSurname = '';
      normalized.requesterManualName = '';
    } else if (normalized.requesterMode === 'MANUAL') {
      normalized.requesterManualName = normalized.requesterManualName || ROOMS_APP.normalizeString([
        normalized.bookerSurname,
        normalized.bookerName
      ].filter(function (token) {
        return Boolean(token);
      }).join(' '));
      if (normalized.requesterManualName && !normalized.bookerName && !normalized.bookerSurname) {
        var manualIdentity = this.parseRequesterName_(normalized.requesterManualName);
        normalized.bookerName = manualIdentity.firstName;
        normalized.bookerSurname = manualIdentity.surname;
      }
    }

    if (normalized.requesterMode !== 'NONE' && !normalized.bookerName && emailIdentity.firstName) {
      normalized.bookerName = emailIdentity.firstName;
    }
    if (normalized.requesterMode !== 'NONE' && !normalized.bookerSurname && emailIdentity.surname) {
      normalized.bookerSurname = emailIdentity.surname;
    }
    if (!user.canAccessAdmin) {
      normalized.bookerName = emailIdentity.firstName || normalized.bookerName;
      normalized.bookerSurname = emailIdentity.surname || normalized.bookerSurname;
    }

    if (!user.email) {
      errors.push('User authentication is required.');
    }

    if (!ROOMS_APP.isEmailInDomain(user.email, ROOMS_APP.getAllowedDomain())) {
      errors.push('Operazione consentita solo con account ' + ROOMS_APP.getAllowedDomain() + '.');
    }

    if (!user.canBook) {
      errors.push('Booking permission required.');
    }

    if (!ROOMS_APP.getBooleanConfig('BOOKING_ENABLED', true) && !user.canAccessAdmin) {
      errors.push('Booking is currently disabled.');
    }

    var resource = this.getResource(normalized.resourceId);
    if (!resource || !ROOMS_APP.asBoolean(resource.IsActive) || !ROOMS_APP.asBoolean(resource.IsBookable)) {
      errors.push('Selected room is not available for booking.');
    }
    if (resource && ROOMS_APP.slugify(resource.DisplayName || '') === 'AULA_MAGNA') {
      errors.push('Aula Magna è gestita tramite eventi dedicati (solo admin).');
    }

    if (!normalized.bookingDate || !normalized.startTime || !normalized.endTime) {
      errors.push('Booking date, start time and end time are required.');
    }

    if (!/^\d{4}-\d{2}-\d{2}$/.test(normalized.bookingDate || '')) {
      errors.push('Booking date format must be YYYY-MM-DD.');
    }

    if (!/^\d{2}:\d{2}$/.test(normalized.startTime || '') || !/^\d{2}:\d{2}$/.test(normalized.endTime || '')) {
      errors.push('Time format must be HH:MM.');
    }

    if (errors.length) {
      return {
        ok: false,
        errors: errors,
        normalized: normalized,
        resource: resource,
        actor: user
      };
    }

    if (normalized.startTime >= normalized.endTime) {
      errors.push('End time must be after start time.');
    }

    var durationMin = ROOMS_APP.minutesBetween(normalized.startTime, normalized.endTime);
    var maxDuration = ROOMS_APP.getNumberConfig('MAX_DURATION_MIN', 180);
    if (!user.canAccessAdmin && durationMin > maxDuration) {
      errors.push('Booking exceeds the maximum allowed duration.');
    }

    var today = ROOMS_APP.toIsoDate(new Date());
    if (!user.canAccessAdmin && ROOMS_APP.daysBetween(today, normalized.bookingDate) > ROOMS_APP.getNumberConfig('MAX_DAYS_AHEAD', 30)) {
      errors.push('Booking exceeds the maximum advance window.');
    }

    if (ROOMS_APP.daysBetween(normalized.bookingDate, today) > 0) {
      errors.push('Booking date is in the past.');
    }

    var dailyOpening = this.getEffectiveOpeningForResource(normalized.resourceId, normalized.bookingDate);
    if (!dailyOpening.isOpen) {
      errors.push('Room is closed on the selected date.');
    } else {
      if (normalized.startTime < dailyOpening.openTime || normalized.endTime > dailyOpening.closeTime) {
        errors.push('Booking must stay inside the opening window.');
      }
    }

    var closure = this.findBlockingClosure(normalized.bookingDate, normalized.startTime, normalized.endTime);
    if (closure) {
      errors.push('Selected slot overlaps a blocked closure: ' + closure.Label);
    }

    if (this.hasConflict(normalized.resourceId, normalized.bookingDate, normalized.startTime, normalized.endTime, normalized.bookingId)) {
      errors.push('Selected slot conflicts with an existing booking.');
    }

    return {
      ok: errors.length === 0,
      errors: errors,
      normalized: normalized,
      resource: resource,
      actor: user
    };
  }
};
