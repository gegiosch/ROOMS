var ROOMS_APP = ROOMS_APP || {};

ROOMS_APP.Slots = {
  BASE_SLOT_MINUTES_: 60,
  HALF_SLOT_MINUTES_: 30,

  getDaySlots: function (resourceId, dateString, options) {
    var opening = ROOMS_APP.Policy.getEffectiveOpeningForResource(resourceId, dateString);
    var bookings = ROOMS_APP.Booking.listBookingsForDay(resourceId, dateString);
    var timetableOccupancies = ROOMS_APP.Timetable.listOccupanciesForDate(resourceId, dateString);
    return this.getDaySlotsFromOccupancies(resourceId, dateString, bookings, timetableOccupancies, opening, options);
  },

  getDaySlotsFromOccupancies: function (resourceId, dateString, bookings, timetableOccupancies, opening, options) {
    var openingWindow = opening || ROOMS_APP.Policy.getEffectiveOpeningForResource(resourceId, dateString);
    var settings = options || {};
    var windowOpenTime = this.normalizeTimeValue_(openingWindow.openTime);
    var windowCloseTime = this.normalizeTimeValue_(openingWindow.closeTime);
    var occupancies = this.sortOccupancies_(
      (bookings || []).map(function (booking) {
        var enriched = {};
        Object.keys(booking || {}).forEach(function (key) {
          enriched[key] = booking[key];
        });
        enriched.SourceKind = 'USER_BOOKING';
        enriched.SourceType = 'USER_BOOKING';
        enriched.DisplayLabel = ROOMS_APP.normalizeString(booking.BookerSurname || booking.BookerName || booking.Title || 'N/D');
        return enriched;
      }).concat(timetableOccupancies || [])
    );

    if (!openingWindow.isOpen || !windowOpenTime || !windowCloseTime || this.toMinutes_(windowOpenTime) >= this.toMinutes_(windowCloseTime)) {
      return {
        isOpen: false,
        openTime: '',
        closeTime: '',
        openingSource: openingWindow.source || 'UNKNOWN',
        bookings: bookings || [],
        timetableOccupancies: timetableOccupancies || [],
        occupancies: occupancies,
        slots: [],
        freeSlots: []
      };
    }

    var baseSlots = this.buildBaseSlots_(dateString, windowOpenTime, windowCloseTime);
    var slotRows = this.applyOccupanciesToSlots_(baseSlots, occupancies);
    if (settings.splitFreeSlotsHalfHour) {
      slotRows = this.splitFreeSlotsIntoHalfHours_(slotRows);
    }
    slotRows = this.filterValidSlots_(slotRows, windowOpenTime, windowCloseTime);

    return {
      isOpen: true,
      openTime: windowOpenTime,
      closeTime: windowCloseTime,
      openingSource: openingWindow.source || '',
      bookings: bookings || [],
      timetableOccupancies: timetableOccupancies || [],
      occupancies: occupancies,
      slots: slotRows,
      freeSlots: this.extractFreeSlots_(slotRows, windowOpenTime, windowCloseTime)
    };
  },

  buildBaseSlots_: function (_dateString, openTime, closeTime) {
    var slots = [];
    var startMinutes = this.toMinutes_(openTime);
    var endMinutes = this.toMinutes_(closeTime);

    if (startMinutes == null || endMinutes == null || startMinutes >= endMinutes) {
      return slots;
    }

    var cursor = startMinutes;
    while (cursor < endMinutes) {
      var next = Math.min(cursor + this.BASE_SLOT_MINUTES_, endMinutes);
      slots.push({
        startTime: this.minutesToTime_(cursor),
        endTime: this.minutesToTime_(next),
        slotMinutes: next - cursor
      });
      cursor = next;
    }

    return slots;
  },

  applyOccupanciesToSlots_: function (slots, occupancies) {
    var self = this;
    var normalizedOccupancies = (occupancies || []).map(function (entry) {
      var startTime = self.normalizeTimeValue_(entry.StartTime);
      var endTime = self.normalizeTimeValue_(entry.EndTime);
      return {
        row: entry,
        startMinutes: self.toMinutes_(startTime),
        endMinutes: self.toMinutes_(endTime)
      };
    }).filter(function (entry) {
      return entry.startMinutes != null && entry.endMinutes != null && entry.endMinutes > entry.startMinutes;
    });

    return (slots || []).map(function (slot) {
      var slotStart = self.toMinutes_(slot.startTime);
      var slotEnd = self.toMinutes_(slot.endTime);
      var occupancyEntry = normalizedOccupancies.filter(function (entry) {
        return !(slotEnd <= entry.startMinutes || slotStart >= entry.endMinutes);
      })[0] || null;
      var occupancy = occupancyEntry ? occupancyEntry.row : null;

      return {
        startTime: slot.startTime,
        endTime: slot.endTime,
        slotMinutes: slot.slotMinutes,
        isOccupied: Boolean(occupancy),
        canBook: !occupancy,
        bookingId: occupancy ? occupancy.BookingId : '',
        sourceKind: occupancy ? (occupancy.SourceKind || 'USER_BOOKING') : '',
        sourceType: occupancy ? (occupancy.SourceType || '') : '',
        displayActor: occupancy ? self.getOccupancyDisplayLabel_(occupancy) : '',
        title: occupancy ? occupancy.Title : '',
        occupancy: occupancy
      };
    });
  },

  splitFreeSlotsIntoHalfHours_: function (slots) {
    var self = this;
    var output = [];

    (slots || []).forEach(function (slot) {
      var slotStart = self.toMinutes_(slot.startTime);
      var slotEnd = self.toMinutes_(slot.endTime);
      if (
        slot.isOccupied ||
        Number(slot.slotMinutes) !== self.BASE_SLOT_MINUTES_ ||
        slotStart == null ||
        slotEnd == null ||
        slotEnd - slotStart !== self.BASE_SLOT_MINUTES_
      ) {
        output.push(slot);
        return;
      }

      var firstEndMinutes = slotStart + self.HALF_SLOT_MINUTES_;
      var firstEnd = self.minutesToTime_(firstEndMinutes);
      output.push({
        startTime: slot.startTime,
        endTime: firstEnd,
        slotMinutes: self.HALF_SLOT_MINUTES_,
        isOccupied: false,
        canBook: true,
        bookingId: '',
        sourceKind: '',
        sourceType: '',
        displayActor: '',
        title: '',
        occupancy: null,
        isSplitFreeSlot: true
      });
      output.push({
        startTime: firstEnd,
        endTime: slot.endTime,
        slotMinutes: self.HALF_SLOT_MINUTES_,
        isOccupied: false,
        canBook: true,
        bookingId: '',
        sourceKind: '',
        sourceType: '',
        displayActor: '',
        title: '',
        occupancy: null,
        isSplitFreeSlot: true
      });
    });

    return output;
  },

  extractFreeSlots_: function (slots, openTime, closeTime) {
    var self = this;
    return (slots || []).filter(function (slot) {
      return !slot.isOccupied && self.isValidSlot_(slot, openTime, closeTime);
    }).map(function (slot) {
      return {
        startTime: slot.startTime,
        endTime: slot.endTime,
        slotMinutes: slot.slotMinutes,
        isOccupied: false,
        canBook: true,
        bookingId: '',
        sourceKind: '',
        sourceType: '',
        displayActor: '',
        title: '',
        occupancy: null,
        isSplitFreeSlot: Boolean(slot.isSplitFreeSlot)
      };
    });
  },

  filterPastFreeSlots_: function (slots, currentTime) {
    var self = this;
    var nowMinutes = this.toMinutes_(currentTime);
    if (nowMinutes == null) {
      return (slots || []).slice();
    }
    return (slots || []).filter(function (slot) {
      return self.toMinutes_(slot && slot.endTime) > nowMinutes;
    });
  },

  filterBookableFreeSlotsForDate_: function (slots, selectedDate, todayDate, currentTimeOfDay) {
    var normalizedSelectedDate = ROOMS_APP.toIsoDate(selectedDate || '');
    var normalizedTodayDate = ROOMS_APP.toIsoDate(todayDate || '');
    if (!normalizedSelectedDate || !normalizedTodayDate) {
      return (slots || []).slice();
    }

    var dateOffset = ROOMS_APP.daysBetween(normalizedTodayDate, normalizedSelectedDate);
    if (dateOffset > 0) {
      return (slots || []).slice();
    }
    if (dateOffset < 0) {
      return (slots || []).slice();
    }

    return this.filterPastFreeSlots_(slots, currentTimeOfDay);
  },

  filterValidSlots_: function (slots, openTime, closeTime) {
    var self = this;
    return (slots || []).filter(function (slot) {
      return self.isValidSlot_(slot, openTime, closeTime);
    });
  },

  isValidSlot_: function (slot, openTime, closeTime) {
    if (!slot) {
      return false;
    }
    var start = this.toMinutes_(slot.startTime);
    var end = this.toMinutes_(slot.endTime);
    var open = this.toMinutes_(openTime);
    var close = this.toMinutes_(closeTime);
    if (start == null || end == null || open == null || close == null) {
      return false;
    }
    if (end <= start) {
      return false;
    }
    if (start < open || end > close) {
      return false;
    }
    return true;
  },

  normalizeTimeValue_: function (value) {
    var normalized = ROOMS_APP.normalizeString(value);
    if (!normalized) {
      return '';
    }
    var match = normalized.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
    if (!match) {
      return '';
    }
    var hours = Number(match[1]);
    var minutes = Number(match[2]);
    if (!isFinite(hours) || !isFinite(minutes)) {
      return '';
    }
    if (minutes < 0 || minutes > 59) {
      return '';
    }
    if (hours < 0 || hours > 24) {
      return '';
    }
    if (hours === 24 && minutes !== 0) {
      return '';
    }
    return (hours < 10 ? '0' : '') + hours + ':' + (minutes < 10 ? '0' : '') + minutes;
  },

  toMinutes_: function (timeString) {
    var normalized = this.normalizeTimeValue_(timeString);
    if (!normalized) {
      return null;
    }
    var parts = normalized.split(':');
    return Number(parts[0]) * 60 + Number(parts[1]);
  },

  minutesToTime_: function (minutes) {
    var total = Number(minutes);
    if (!isFinite(total) || total < 0) {
      return '';
    }
    var bounded = Math.min(total, 24 * 60);
    var hours = Math.floor(bounded / 60);
    var mins = bounded % 60;
    return (hours < 10 ? '0' : '') + hours + ':' + (mins < 10 ? '0' : '') + mins;
  },

  sortOccupancies_: function (rows) {
    return ROOMS_APP.sortBy(rows || [], ['StartTime', 'EndTime', 'SourceKind', 'BookingId']);
  },

  getOccupancyDisplayLabel_: function (occupancy) {
    if (!occupancy) {
      return '';
    }
    if (ROOMS_APP.normalizeString(occupancy.SourceKind) === 'TIMETABLE') {
      return ROOMS_APP.Timetable.getDisplayLabel(occupancy);
    }
    return ROOMS_APP.normalizeString(occupancy.BookerSurname || occupancy.BookerName || occupancy.Title || 'N/D');
  }
};
