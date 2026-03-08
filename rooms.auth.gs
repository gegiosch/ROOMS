var ROOMS_APP = ROOMS_APP || {};

ROOMS_APP.Auth = {
  getUserContext: function () {
    var email = ROOMS_APP.getCurrentUserEmail();
    var identity = ROOMS_APP.extractIdentityFromEmail(email);

    return {
      email: email,
      isAuthenticated: Boolean(email),
      isAdmin: this.isAdmin(email),
      allowedDomain: ROOMS_APP.getAllowedDomain(),
      isAllowedDomain: ROOMS_APP.isEmailInDomain(email, ROOMS_APP.getAllowedDomain()),
      firstName: identity.firstName,
      surname: identity.surname,
      displayName: identity.displayName
    };
  },

  isAdmin: function (email) {
    var normalized = ROOMS_APP.normalizeEmail(email || ROOMS_APP.getCurrentUserEmail());
    if (!normalized) {
      return false;
    }

    return this.getAdminEntries_().some(function (entry) {
      return entry.Email === normalized;
    });
  },

  requireAdmin: function () {
    var user = this.getUserContext();
    if (!user.isAdmin) {
      throw new Error('Admin access required.');
    }
    return user;
  },

  assertAllowedDomain: function (email) {
    var normalized = ROOMS_APP.normalizeEmail(email || ROOMS_APP.getCurrentUserEmail());
    var allowedDomain = ROOMS_APP.getAllowedDomain();
    if (!normalized || !ROOMS_APP.isEmailInDomain(normalized, allowedDomain)) {
      throw new Error('Operazione consentita solo con account ' + allowedDomain + '.');
    }
    return normalized;
  },

  canManageBooking: function (booking, actor) {
    var user = actor || this.getUserContext();
    return Boolean(
      booking &&
      (user.isAdmin || ROOMS_APP.normalizeEmail(booking.BookerEmail) === ROOMS_APP.normalizeEmail(user.email))
    );
  },

  getAdminEntries_: function () {
    ROOMS_APP.Schema.ensureAdmins();
    return ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.ADMINS)
      .filter(function (row) {
        return ROOMS_APP.normalizeEmail(row.Email) && ROOMS_APP.asBoolean(row.Enabled);
      })
      .map(function (row) {
        return {
          Email: ROOMS_APP.normalizeEmail(row.Email),
          Role: ROOMS_APP.normalizeString(row.Role || 'ADMIN'),
          Enabled: 'TRUE',
          Notes: ROOMS_APP.normalizeString(row.Notes)
        };
      });
  }
};
