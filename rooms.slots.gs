var ROOMS_APP = ROOMS_APP || {};

ROOMS_APP.Slots = {
  getDaySlots: function (resourceId, dateString) {
    var opening = ROOMS_APP.Policy.getDailyOpening(dateString);
    var bookings = ROOMS_APP.Booking.listBookingsForDay(resourceId, dateString);
    var timetableOccupancies = ROOMS_APP.Timetable.listOccupanciesForDate(resourceId, dateString);
    var occupancies = this.sortOccupancies_(
      bookings.map(function (booking) {
        var enriched = {};
        Object.keys(booking || {}).forEach(function (key) {
          enriched[key] = booking[key];
        });
        enriched.SourceKind = 'USER_BOOKING';
        enriched.SourceType = 'USER_BOOKING';
        enriched.DisplayLabel = ROOMS_APP.normalizeString(booking.BookerSurname || booking.BookerName || booking.Title || 'N/D');
        return enriched;
      }).concat(timetableOccupancies)
    );
    var slotMinutes = ROOMS_APP.getNumberConfig('SLOT_MINUTES', 30);

    if (!opening.isOpen) {
      return {
        isOpen: false,
        openTime: '',
        closeTime: '',
        bookings: bookings,
        timetableOccupancies: timetableOccupancies,
        occupancies: occupancies,
        slots: [],
        freeSlots: []
      };
    }

    var slots = [];
    var cursor = ROOMS_APP.combineDateTime(dateString, opening.openTime);
    var end = ROOMS_APP.combineDateTime(dateString, opening.closeTime);

    while (cursor.getTime() < end.getTime()) {
      var slotStart = Utilities.formatDate(cursor, ROOMS_APP.getTimezone(), 'HH:mm');
      var next = new Date(Math.min(cursor.getTime() + slotMinutes * 60000, end.getTime()));
      var slotEnd = Utilities.formatDate(next, ROOMS_APP.getTimezone(), 'HH:mm');
      var occupancy = occupancies.filter(function (entry) {
        return !(slotEnd <= entry.StartTime || slotStart >= entry.EndTime);
      })[0] || null;

      slots.push({
        startTime: slotStart,
        endTime: slotEnd,
        isOccupied: Boolean(occupancy),
        bookingId: occupancy ? occupancy.BookingId : '',
        sourceKind: occupancy ? (occupancy.SourceKind || 'USER_BOOKING') : '',
        sourceType: occupancy ? (occupancy.SourceType || '') : '',
        surname: occupancy ? this.getOccupancyDisplayLabel_(occupancy) : '',
        title: occupancy ? occupancy.Title : ''
      });

      cursor = next;
    }

    return {
      isOpen: true,
      openTime: opening.openTime,
      closeTime: opening.closeTime,
      bookings: bookings,
      timetableOccupancies: timetableOccupancies,
      occupancies: occupancies,
      slots: slots,
      freeSlots: this.buildFreeSlots_(opening, occupancies, slots)
    };
  },

  buildFreeSlots_: function (opening, occupancies, slots) {
    if (!opening || !opening.isOpen) {
      return [];
    }

    if (!occupancies || !occupancies.length) {
      return [{
        startTime: opening.openTime,
        endTime: opening.closeTime
      }];
    }

    return this.mergeFreeSlots_(slots || []);
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
  },

  mergeFreeSlots_: function (slots) {
    var freeSlots = [];

    slots.forEach(function (slot) {
      if (slot.isOccupied) {
        return;
      }

      var current = freeSlots[freeSlots.length - 1];
      if (current && current.endTime === slot.startTime) {
        current.endTime = slot.endTime;
      } else {
        freeSlots.push({
          startTime: slot.startTime,
          endTime: slot.endTime
        });
      }
    });

    return freeSlots;
  }
};
