var ROOMS_APP = ROOMS_APP || {};

ROOMS_APP.Slots = {
  getDaySlots: function (resourceId, dateString) {
    var opening = ROOMS_APP.Policy.getDailyOpening(dateString);
    var bookings = ROOMS_APP.Booking.listBookingsForDay(resourceId, dateString);
    var slotMinutes = ROOMS_APP.getNumberConfig('SLOT_MINUTES', 30);

    if (!opening.isOpen) {
      return {
        isOpen: false,
        openTime: '',
        closeTime: '',
        bookings: bookings,
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
      var booking = bookings.filter(function (entry) {
        return !(slotEnd <= entry.StartTime || slotStart >= entry.EndTime);
      })[0] || null;

      slots.push({
        startTime: slotStart,
        endTime: slotEnd,
        isOccupied: Boolean(booking),
        bookingId: booking ? booking.BookingId : '',
        surname: booking ? booking.BookerSurname : '',
        title: booking ? booking.Title : ''
      });

      cursor = next;
    }

    return {
      isOpen: true,
      openTime: opening.openTime,
      closeTime: opening.closeTime,
      bookings: bookings,
      slots: slots,
      freeSlots: this.buildFreeSlots_(opening, bookings, slots)
    };
  },

  buildFreeSlots_: function (opening, bookings, slots) {
    if (!opening || !opening.isOpen) {
      return [];
    }

    if (!bookings || !bookings.length) {
      return [{
        startTime: opening.openTime,
        endTime: opening.closeTime
      }];
    }

    return this.mergeFreeSlots_(slots || []);
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
