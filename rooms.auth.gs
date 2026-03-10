var ROOMS_APP = ROOMS_APP || {};

ROOMS_APP.Auth = {
  SIMULATION_INPUT_PATTERN_: /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2})?$/,

  getUserContext: function () {
    var email = ROOMS_APP.getCurrentUserEmail();
    var identity = ROOMS_APP.extractIdentityFromEmail(email);
    var adminEntry = this.getAdminEntry_(email);
    var role = adminEntry ? ROOMS_APP.normalizeString(adminEntry.Role || 'ADMIN').toUpperCase() : '';

    return {
      email: email,
      isAuthenticated: Boolean(email),
      isAdmin: Boolean(adminEntry),
      role: role || 'USER',
      isSuperAdmin: role === 'SUPERADMIN',
      allowedDomain: ROOMS_APP.getAllowedDomain(),
      isAllowedDomain: ROOMS_APP.isEmailInDomain(email, ROOMS_APP.getAllowedDomain()),
      firstName: identity.firstName,
      surname: identity.surname,
      displayName: identity.displayName
    };
  },

  isAdmin: function (email) {
    return Boolean(this.getAdminEntry_(email || ROOMS_APP.getCurrentUserEmail()));
  },

  isSuperAdmin: function (email) {
    var entry = this.getAdminEntry_(email || ROOMS_APP.getCurrentUserEmail());
    if (!entry) {
      return false;
    }
    return ROOMS_APP.normalizeString(entry.Role || '').toUpperCase() === 'SUPERADMIN';
  },

  requireAdmin: function () {
    var user = this.getUserContext();
    if (!user.isAdmin) {
      throw new Error('Admin access required.');
    }
    return user;
  },

  requireSuperAdmin: function () {
    var user = this.getUserContext();
    if (!user.isSuperAdmin) {
      throw new Error('Superadmin access required.');
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

  getEffectiveNow: function (requestContext, actor) {
    var effectiveActor = actor || this.getUserContext();
    var simulation = this.getSimulationContext_(requestContext, effectiveActor);
    if (!simulation.active || !simulation.date) {
      return new Date();
    }
    return simulation.date;
  },

  getSimulationContext_: function (requestContext, actor) {
    var user = actor || this.getUserContext();
    var candidate = '';
    var source = requestContext || ROOMS_APP.RUNTIME_CONTEXT_ || {};
    if (source && typeof source === 'object') {
      candidate = ROOMS_APP.normalizeString(
        source.simulatedNow ||
        source.simulatedDateTime ||
        source.__simulatedNow ||
        ''
      );
    }

    var parsed = this.parseSimulationDateTime_(candidate);
    return {
      active: Boolean(user && user.isSuperAdmin && parsed),
      iso: parsed ? ROOMS_APP.toIsoDateTime(parsed) : '',
      date: parsed || null
    };
  },

  parseSimulationDateTime_: function (value) {
    var normalized = ROOMS_APP.normalizeString(value);
    if (!normalized || !this.SIMULATION_INPUT_PATTERN_.test(normalized)) {
      return null;
    }

    var raw = normalized.length === 16 ? (normalized + ':00') : normalized;
    var parsed = new Date(raw);
    if (isNaN(parsed.getTime())) {
      return null;
    }
    return parsed;
  },

  getAdminEntry_: function (email) {
    var normalized = ROOMS_APP.normalizeEmail(email || ROOMS_APP.getCurrentUserEmail());
    if (!normalized) {
      return null;
    }

    return this.getAdminEntries_().filter(function (entry) {
      return entry.Email === normalized;
    })[0] || null;
  },

  getAdminEntries_: function () {
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
