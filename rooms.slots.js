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

    if (!openingWindow.isOpen) {
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

    var baseSlots = this.buildBaseSlots_(dateString, openingWindow.openTime, openingWindow.closeTime);
    var slotRows = this.applyOccupanciesToSlots_(baseSlots, occupancies);
    if (settings.splitFreeSlotsHalfHour) {
      slotRows = this.splitFreeSlotsIntoHalfHours_(slotRows);
    }

    return {
      isOpen: true,
      openTime: openingWindow.openTime,
      closeTime: openingWindow.closeTime,
      openingSource: openingWindow.source || '',
      bookings: bookings || [],
      timetableOccupancies: timetableOccupancies || [],
      occupancies: occupancies,
      slots: slotRows,
      freeSlots: this.extractFreeSlots_(slotRows)
    };
  },

  buildBaseSlots_: function (dateString, openTime, closeTime) {
    var slots = [];
    var cursor = ROOMS_APP.combineDateTime(dateString, openTime);
    var end = ROOMS_APP.combineDateTime(dateString, closeTime);

    while (cursor.getTime() < end.getTime()) {
      var slotStart = Utilities.formatDate(cursor, ROOMS_APP.getTimezone(), 'HH:mm');
      var next = new Date(Math.min(cursor.getTime() + this.BASE_SLOT_MINUTES_ * 60000, end.getTime()));
      var slotEnd = Utilities.formatDate(next, ROOMS_APP.getTimezone(), 'HH:mm');
      slots.push({
        startTime: slotStart,
        endTime: slotEnd,
        slotMinutes: ROOMS_APP.minutesBetween(slotStart, slotEnd)
      });
      cursor = next;
    }

    return slots;
  },

  applyOccupanciesToSlots_: function (slots, occupancies) {
    var self = this;
    return (slots || []).map(function (slot) {
      var occupancy = (occupancies || []).filter(function (entry) {
        return !(slot.endTime <= entry.StartTime || slot.startTime >= entry.EndTime);
      })[0] || null;

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
      if (slot.isOccupied || slot.slotMinutes !== self.BASE_SLOT_MINUTES_) {
        output.push(slot);
        return;
      }

      var firstEnd = ROOMS_APP.addMinutes(slot.startTime, self.HALF_SLOT_MINUTES_);
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

  extractFreeSlots_: function (slots) {
    return (slots || []).filter(function (slot) {
      return !slot.isOccupied;
    }).map(function (slot) {
      return {
        startTime: slot.startTime,
        endTime: slot.endTime
      };
    });
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
