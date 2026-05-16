var ROOMS_APP = ROOMS_APP || {};

ROOMS_APP.Auth = {
  SIMULATION_INPUT_PATTERN_: /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2})?$/,
  USER_CONTEXT_CACHE_TTL_: 3600,
  USER_CONTEXT_CACHE_VERSION_KEY_: 'rooms:user-context-cache-version',
  USER_CONTEXT_CACHE_SCHEMA_VERSION_: 'booking-views-v1',
  TAB_PERMISSION_KEYS_: {
    absences: 'canAccessAbsences',
    classEvents: 'canAccessClassEvents',
    longReplacements: 'canAccessLongReplacements',
    dailyReplacements: 'canAccessDailyReplacements',
    bookings: 'canAccessBookings',
    report: 'canAccessReport'
  },
  BOOKING_VIEW_PERMISSION_KEYS_: {
    create: 'canViewBookingCreate',
    mine: 'canViewBookingMine',
    daily: 'canViewDailyBookings',
    future: 'canViewFutureBookings',
    replacementReport: 'canViewReplacementReport',
    ownRecovery: 'canViewOwnRecovery'
  },

  getUserContext: function () {
    var email = ROOMS_APP.getCurrentUserEmail();
    var identity = ROOMS_APP.extractIdentityFromEmail(email);
    var resolved = this.resolveCachedUserContext_(email);
    var permissions = resolved && resolved.permissions ? resolved.permissions : this.buildFallbackPermissions_();
    var role = ROOMS_APP.normalizeString(permissions.role || 'USER').toUpperCase() || 'USER';

    return {
      email: email,
      orgUnitPath: ROOMS_APP.normalizeString(resolved && resolved.orgUnitPath),
      isAuthenticated: Boolean(email),
      isAdmin: Boolean(permissions.canAccessAdmin),
      role: role,
      isSuperAdmin: role === 'SUPERADMIN',
      canBook: Boolean(permissions.canBook),
      canManageReplacement: Boolean(permissions.canManageReplacement),
      canManageAulaMagna: Boolean(permissions.canManageAulaMagna),
      canUseSimulation: Boolean(permissions.canUseSimulation),
      canAccessAdmin: Boolean(permissions.canAccessAdmin),
      canAccessAbsences: Boolean(permissions.canAccessAbsences),
      canAccessClassEvents: Boolean(permissions.canAccessClassEvents),
      canAccessLongReplacements: Boolean(permissions.canAccessLongReplacements),
      canAccessDailyReplacements: Boolean(permissions.canAccessDailyReplacements),
      canAccessBookings: Boolean(permissions.canAccessBookings),
      canAccessReport: Boolean(permissions.canAccessReport),
      canViewReports: Boolean(permissions.canAccessReport),
      canViewBookingCreate: Boolean(permissions.canViewBookingCreate),
      canViewBookingMine: Boolean(permissions.canViewBookingMine),
      canViewDailyBookings: Boolean(permissions.canViewDailyBookings),
      canViewFutureBookings: Boolean(permissions.canViewFutureBookings),
      canViewReplacementReport: Boolean(permissions.canViewReplacementReport),
      canViewOwnRecovery: Boolean(permissions.canViewOwnRecovery),
      allowedDomain: ROOMS_APP.getAllowedDomain(),
      isAllowedDomain: ROOMS_APP.isEmailInDomain(email, ROOMS_APP.getAllowedDomain()),
      firstName: identity.firstName,
      surname: identity.surname,
      displayName: identity.displayName
    };
  },

  isAdmin: function (email) {
    return Boolean(this.getUserContextForEmail_(email || ROOMS_APP.getCurrentUserEmail()).canAccessAdmin);
  },

  isSuperAdmin: function (email) {
    return Boolean(this.getUserContextForEmail_(email || ROOMS_APP.getCurrentUserEmail()).isSuperAdmin);
  },

  requireAdmin: function () {
    var user = this.getUserContext();
    this.assertAllowedDomain(user.email);
    if (!user.canAccessAdmin) {
      throw new Error('Admin access required.');
    }
    return user;
  },

  requireCanBook: function () {
    var user = this.getUserContext();
    if (!user.canBook) {
      throw new Error('Booking permission required.');
    }
    return user;
  },

  requireCanManageReplacement: function () {
    var user = this.getUserContext();
    if (!user.canManageReplacement) {
      throw new Error('Replacement management permission required.');
    }
    return user;
  },

  requireCanManageAulaMagna: function () {
    var user = this.getUserContext();
    if (!user.canManageAulaMagna) {
      throw new Error('Aula Magna management permission required.');
    }
    return user;
  },

  requireCanViewReports: function () {
    return this.requireAdminTab('report');
  },

  canViewReports: function (actor) {
    return this.canAccessAdminTab(actor, 'report');
  },

  requireAdminTab: function (tabKey) {
    var user = this.requireAdmin();
    if (!this.canAccessAdminTab(user, tabKey)) {
      throw new Error('Permissione non sufficiente per questa scheda amministrativa.');
    }
    return user;
  },

  requireAnyAdminTab: function (tabKeys) {
    var user = this.requireAdmin();
    var keys = Array.isArray(tabKeys) ? tabKeys : [tabKeys];
    var index;
    for (index = 0; index < keys.length; index += 1) {
      if (this.canAccessAdminTab(user, keys[index])) {
        return user;
      }
    }
    throw new Error('Permissione non sufficiente per questa scheda amministrativa.');
  },

  canAccessAdminTab: function (actor, tabKey) {
    var user = actor || this.getUserContext();
    var permissionKey = this.TAB_PERMISSION_KEYS_[tabKey] || '';
    return Boolean(
      user &&
      user.isAllowedDomain !== false &&
      user.canAccessAdmin &&
      (
        user.isSuperAdmin ||
        (permissionKey && user[permissionKey])
      )
    );
  },

  requireBookingPageView: function (viewKey) {
    var user = this.getUserContext();
    this.assertAllowedDomain(user.email);
    if (!this.canViewBookingPageSection(user, viewKey)) {
      throw new Error('Booking page view permission required.');
    }
    return user;
  },

  requireAnyBookingPageView: function (viewKeys) {
    var user = this.getUserContext();
    var keys = Array.isArray(viewKeys) ? viewKeys : [viewKeys];
    var index;
    this.assertAllowedDomain(user.email);
    for (index = 0; index < keys.length; index += 1) {
      if (this.canViewBookingPageSection(user, keys[index])) {
        return user;
      }
    }
    throw new Error('Booking page view permission required.');
  },

  canViewBookingPageSection: function (actor, viewKey) {
    var user = actor || this.getUserContext();
    var permissionKey = this.BOOKING_VIEW_PERMISSION_KEYS_[viewKey] || '';
    return Boolean(
      user &&
      user.isAllowedDomain !== false &&
      (
        user.isSuperAdmin ||
        (permissionKey && user[permissionKey])
      )
    );
  },

  isAtaLikeRole_: function (role) {
    var normalized = ROOMS_APP.normalizeString(role).toUpperCase().replace(/[\s-]+/g, '_');
    return normalized === 'ATA' ||
      normalized === 'COLLABORATORE' ||
      normalized === 'COLLABORATORE_SCOLASTICO' ||
      normalized.indexOf('COLLABORATORE') >= 0 ||
      normalized.indexOf('ATA') >= 0;
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
      (
        user.canManageReplacement ||
        (
          user.canBook &&
          ROOMS_APP.normalizeEmail(booking.BookerEmail) === ROOMS_APP.normalizeEmail(user.email)
        )
      )
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
      active: Boolean(user && user.canUseSimulation && parsed),
      iso: parsed ? parsed.iso : '',
      dateIso: parsed ? parsed.dateIso : '',
      time: parsed ? parsed.time : '',
      date: parsed ? parsed.date : null
    };
  },

  parseSimulationDateTime_: function (value) {
    var normalized = ROOMS_APP.normalizeString(value);
    if (!normalized || !this.SIMULATION_INPUT_PATTERN_.test(normalized)) {
      return null;
    }
    var match = normalized.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2}))?$/);
    if (!match) {
      return null;
    }
    var year = Number(match[1]);
    var month = Number(match[2]);
    var day = Number(match[3]);
    var hour = Number(match[4]);
    var minute = Number(match[5]);
    var second = Number(match[6] || '00');
    if (
      !isFinite(year) || !isFinite(month) || !isFinite(day) ||
      !isFinite(hour) || !isFinite(minute) || !isFinite(second) ||
      month < 1 || month > 12 ||
      day < 1 || day > 31 ||
      hour < 0 || hour > 23 ||
      minute < 0 || minute > 59 ||
      second < 0 || second > 59
    ) {
      return null;
    }

    var parsed = new Date(year, month - 1, day, hour, minute, second, 0);
    if (
      isNaN(parsed.getTime()) ||
      parsed.getFullYear() !== year ||
      parsed.getMonth() !== month - 1 ||
      parsed.getDate() !== day ||
      parsed.getHours() !== hour ||
      parsed.getMinutes() !== minute ||
      parsed.getSeconds() !== second
    ) {
      return null;
    }

    var dateIso = [
      String(year),
      ('0' + String(month)).slice(-2),
      ('0' + String(day)).slice(-2)
    ].join('-');
    var time = ('0' + String(hour)).slice(-2) + ':' + ('0' + String(minute)).slice(-2);

    return {
      iso: dateIso + 'T' + time + ':' + ('0' + String(second)).slice(-2),
      dateIso: dateIso,
      time: time,
      date: parsed
    };
  },

  getAdminEntry_: function (email) {
    var normalized = ROOMS_APP.normalizeEmail(email || ROOMS_APP.getCurrentUserEmail());
    return this.resolvePermissionEntry_(normalized, '');
  },

  getAdminEntries_: function () {
    return ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.ADMINS)
      .filter(function (row) {
        return ROOMS_APP.asBoolean(row.Enabled) &&
          (ROOMS_APP.normalizeEmail(row.Email) || ROOMS_APP.Auth.normalizeOrgUnitPath_(row.OrgUnitPath));
      })
      .map(function (row) {
        return {
          Email: ROOMS_APP.normalizeEmail(row.Email),
          OrgUnitPath: ROOMS_APP.Auth.normalizeOrgUnitPath_(row.OrgUnitPath),
          Role: ROOMS_APP.normalizeString(row.Role || 'ADMIN'),
          Enabled: 'TRUE',
          CanBook: row.CanBook,
          CanManageReplacement: row.CanManageReplacement,
          CanManageAulaMagna: row.CanManageAulaMagna,
          CanUseSimulation: row.CanUseSimulation,
          CanAccessAdmin: row.CanAccessAdmin,
          CanAccessAbsences: row.CanAccessAbsences,
          CanAccessClassEvents: row.CanAccessClassEvents,
          CanAccessLongReplacements: row.CanAccessLongReplacements,
          CanAccessDailyReplacements: row.CanAccessDailyReplacements,
          CanAccessBookings: row.CanAccessBookings,
          CanAccessReport: row.CanAccessReport,
          CanViewBookingCreate: row.CanViewBookingCreate,
          CanViewBookingMine: row.CanViewBookingMine,
          CanViewDailyBookings: row.CanViewDailyBookings,
          CanViewFutureBookings: row.CanViewFutureBookings,
          CanViewReplacementReport: row.CanViewReplacementReport,
          CanViewOwnRecovery: row.CanViewOwnRecovery,
          Notes: ROOMS_APP.normalizeString(row.Notes)
        };
      });
  },

  getUserContextForEmail_: function (email) {
    var normalized = ROOMS_APP.normalizeEmail(email);
    if (!normalized || normalized === ROOMS_APP.getCurrentUserEmail()) {
      return this.getUserContext();
    }

    var identity = ROOMS_APP.extractIdentityFromEmail(normalized);
    var resolved = this.resolveCachedUserContext_(normalized);
    var permissions = resolved && resolved.permissions ? resolved.permissions : this.buildFallbackPermissions_();
    var role = ROOMS_APP.normalizeString(permissions.role || 'USER').toUpperCase() || 'USER';
    return {
      email: normalized,
      orgUnitPath: ROOMS_APP.normalizeString(resolved && resolved.orgUnitPath),
      isAuthenticated: Boolean(normalized),
      isAdmin: Boolean(permissions.canAccessAdmin),
      role: role,
      isSuperAdmin: role === 'SUPERADMIN',
      canBook: Boolean(permissions.canBook),
      canManageReplacement: Boolean(permissions.canManageReplacement),
      canManageAulaMagna: Boolean(permissions.canManageAulaMagna),
      canUseSimulation: Boolean(permissions.canUseSimulation),
      canAccessAdmin: Boolean(permissions.canAccessAdmin),
      canAccessAbsences: Boolean(permissions.canAccessAbsences),
      canAccessClassEvents: Boolean(permissions.canAccessClassEvents),
      canAccessLongReplacements: Boolean(permissions.canAccessLongReplacements),
      canAccessDailyReplacements: Boolean(permissions.canAccessDailyReplacements),
      canAccessBookings: Boolean(permissions.canAccessBookings),
      canAccessReport: Boolean(permissions.canAccessReport),
      canViewReports: Boolean(permissions.canAccessReport),
      canViewBookingCreate: Boolean(permissions.canViewBookingCreate),
      canViewBookingMine: Boolean(permissions.canViewBookingMine),
      canViewDailyBookings: Boolean(permissions.canViewDailyBookings),
      canViewFutureBookings: Boolean(permissions.canViewFutureBookings),
      canViewReplacementReport: Boolean(permissions.canViewReplacementReport),
      canViewOwnRecovery: Boolean(permissions.canViewOwnRecovery),
      allowedDomain: ROOMS_APP.getAllowedDomain(),
      isAllowedDomain: ROOMS_APP.isEmailInDomain(normalized, ROOMS_APP.getAllowedDomain()),
      firstName: identity.firstName,
      surname: identity.surname,
      displayName: identity.displayName
    };
  },

  resolveCachedUserContext_: function (email) {
    var normalized = ROOMS_APP.normalizeEmail(email);
    if (!normalized) {
      return {
        email: '',
        orgUnitPath: '',
        permissions: this.buildFallbackPermissions_()
      };
    }

    var cached = this.getCachedUserContext_(normalized);
    if (cached) {
      return cached;
    }

    var orgUnitPath = this.getDirectoryOrgUnitPath_(normalized);
    var permissions = this.resolvePermissions_(normalized, orgUnitPath);
    var payload = {
      email: normalized,
      orgUnitPath: orgUnitPath,
      permissions: permissions
    };
    this.putCachedUserContext_(normalized, payload);
    return payload;
  },

  getCachedUserContext_: function (email) {
    var cached = CacheService.getScriptCache().get(this.getUserContextCacheKey_(email));
    var payload = cached ? ROOMS_APP.parseJson(cached, null) : null;
    if (!payload || !payload.permissions || String(payload.version || '0') !== this.getUserContextCacheVersion_()) {
      return null;
    }
    return {
      email: ROOMS_APP.normalizeEmail(payload.email || email),
      orgUnitPath: this.normalizeOrgUnitPath_(payload.orgUnitPath),
      permissions: this.normalizePermissionPayload_(payload.permissions, email)
    };
  },

  putCachedUserContext_: function (email, payload) {
    CacheService.getScriptCache().put(
      this.getUserContextCacheKey_(email),
      JSON.stringify({
        version: this.getUserContextCacheVersion_(),
        email: ROOMS_APP.normalizeEmail(payload && payload.email),
        orgUnitPath: this.normalizeOrgUnitPath_(payload && payload.orgUnitPath),
        permissions: this.normalizePermissionPayload_(payload && payload.permissions, email)
      }),
      this.USER_CONTEXT_CACHE_TTL_
    );
  },

  getUserContextCacheKey_: function (email) {
    return 'rooms:user:' + this.USER_CONTEXT_CACHE_SCHEMA_VERSION_ + ':' + ROOMS_APP.normalizeEmail(email);
  },

  getUserContextCacheVersion_: function () {
    return String(
      PropertiesService.getScriptProperties().getProperty(this.USER_CONTEXT_CACHE_VERSION_KEY_) || '0'
    );
  },

  bumpUserContextCacheVersion_: function () {
    PropertiesService.getScriptProperties().setProperty(
      this.USER_CONTEXT_CACHE_VERSION_KEY_,
      String(Date.now())
    );
  },

  getDirectoryOrgUnitPath_: function (email) {
    var normalized = ROOMS_APP.normalizeEmail(email);
    if (!normalized) {
      return '';
    }
    if (typeof AdminDirectory === 'undefined' || !AdminDirectory.Users || typeof AdminDirectory.Users.get !== 'function') {
      return '';
    }

    try {
      var user = AdminDirectory.Users.get(normalized);
      return this.normalizeOrgUnitPath_(user && user.orgUnitPath);
    } catch (error) {
      Logger.log(
        '[WARN] Auth.getDirectoryOrgUnitPath_ email=%s error=%s',
        normalized,
        String(error && error.message ? error.message : error)
      );
      return '';
    }
  },

  resolvePermissions_: function (email, orgUnitPath) {
    var entry = this.resolvePermissionEntry_(email, orgUnitPath);
    if (!entry) {
      return this.buildFallbackPermissions_();
    }
    return this.buildPermissionsFromEntry_(entry, email);
  },

  resolvePermissionEntry_: function (email, orgUnitPath) {
    var normalizedEmail = ROOMS_APP.normalizeEmail(email);
    var normalizedOrgUnitPath = this.normalizeOrgUnitPath_(orgUnitPath);
    var entries = this.getAdminEntries_();
    var orgUnitEntry = null;
    var orgUnitEntryLength = -1;
    var index;

    for (index = 0; index < entries.length; index += 1) {
      if (entries[index].Email && entries[index].Email === normalizedEmail) {
        return entries[index];
      }
    }

    for (index = 0; index < entries.length; index += 1) {
      if (!entries[index].OrgUnitPath || !normalizedOrgUnitPath) {
        continue;
      }
      if (this.isOrgUnitPathMatch_(normalizedOrgUnitPath, entries[index].OrgUnitPath) &&
          entries[index].OrgUnitPath.length > orgUnitEntryLength) {
        orgUnitEntry = entries[index];
        orgUnitEntryLength = entries[index].OrgUnitPath.length;
      }
    }

    return orgUnitEntry;
  },

  buildPermissionsFromEntry_: function (entry, email) {
    var role = ROOMS_APP.normalizeString(entry && entry.Role || 'USER').toUpperCase() || 'USER';
    if (role === 'SUPERADMIN') {
      return {
        role: 'SUPERADMIN',
        canBook: this.readPermissionValue_(entry && entry.CanBook, false),
        canManageReplacement: true,
        canManageAulaMagna: true,
        canUseSimulation: true,
        canAccessAdmin: true,
        canAccessAbsences: true,
        canAccessClassEvents: true,
        canAccessLongReplacements: true,
        canAccessDailyReplacements: true,
        canAccessBookings: true,
        canAccessReport: true,
        canViewBookingCreate: true,
        canViewBookingMine: true,
        canViewDailyBookings: true,
        canViewFutureBookings: true,
        canViewReplacementReport: true,
        canViewOwnRecovery: true
      };
    }

    return {
      role: role,
      canBook: this.readPermissionValue_(entry && entry.CanBook, false),
      canManageReplacement: this.readPermissionValue_(entry && entry.CanManageReplacement, false),
      canManageAulaMagna: this.readPermissionValue_(entry && entry.CanManageAulaMagna, false),
      canUseSimulation: this.readPermissionValue_(entry && entry.CanUseSimulation, false),
      canAccessAdmin: this.readPermissionValue_(entry && entry.CanAccessAdmin, false),
      canAccessAbsences: this.readPermissionValue_(entry && entry.CanAccessAbsences, false),
      canAccessClassEvents: this.readPermissionValue_(entry && entry.CanAccessClassEvents, false),
      canAccessLongReplacements: this.readPermissionValue_(entry && entry.CanAccessLongReplacements, false),
      canAccessDailyReplacements: this.readPermissionValue_(entry && entry.CanAccessDailyReplacements, false),
      canAccessBookings: this.readPermissionValue_(entry && entry.CanAccessBookings, false),
      canAccessReport: this.readPermissionValue_(entry && entry.CanAccessReport, false),
      canViewBookingCreate: this.readPermissionValue_(entry && entry.CanViewBookingCreate, false),
      canViewBookingMine: this.readPermissionValue_(entry && entry.CanViewBookingMine, false),
      canViewDailyBookings: this.readPermissionValue_(entry && entry.CanViewDailyBookings, false),
      canViewFutureBookings: this.readPermissionValue_(entry && entry.CanViewFutureBookings, false),
      canViewReplacementReport: this.readPermissionValue_(entry && entry.CanViewReplacementReport, false),
      canViewOwnRecovery: this.readPermissionValue_(entry && entry.CanViewOwnRecovery, false)
    };
  },

  buildFallbackPermissions_: function () {
    return {
      role: 'USER',
      canBook: false,
      canManageReplacement: false,
      canManageAulaMagna: false,
      canUseSimulation: false,
      canAccessAdmin: false,
      canAccessAbsences: false,
      canAccessClassEvents: false,
      canAccessLongReplacements: false,
      canAccessDailyReplacements: false,
      canAccessBookings: false,
      canAccessReport: false,
      canViewBookingCreate: false,
      canViewBookingMine: false,
      canViewDailyBookings: false,
      canViewFutureBookings: false,
      canViewReplacementReport: false,
      canViewOwnRecovery: false
    };
  },

  normalizePermissionPayload_: function (permissions, email) {
    var source = permissions || {};
    var role = ROOMS_APP.normalizeString(source.role || source.Role || 'USER').toUpperCase() || 'USER';
    return {
      role: role,
      canBook: this.readPermissionValue_(source.canBook, false),
      canManageReplacement: role === 'SUPERADMIN' ? true : this.readPermissionValue_(source.canManageReplacement, false),
      canManageAulaMagna: role === 'SUPERADMIN' ? true : this.readPermissionValue_(source.canManageAulaMagna, false),
      canUseSimulation: role === 'SUPERADMIN' ? true : this.readPermissionValue_(source.canUseSimulation, false),
      canAccessAdmin: role === 'SUPERADMIN' ? true : this.readPermissionValue_(source.canAccessAdmin, false),
      canAccessAbsences: role === 'SUPERADMIN' ? true : this.readPermissionValue_(source.canAccessAbsences, false),
      canAccessClassEvents: role === 'SUPERADMIN' ? true : this.readPermissionValue_(source.canAccessClassEvents, false),
      canAccessLongReplacements: role === 'SUPERADMIN' ? true : this.readPermissionValue_(source.canAccessLongReplacements, false),
      canAccessDailyReplacements: role === 'SUPERADMIN' ? true : this.readPermissionValue_(source.canAccessDailyReplacements, false),
      canAccessBookings: role === 'SUPERADMIN' ? true : this.readPermissionValue_(source.canAccessBookings, false),
      canAccessReport: role === 'SUPERADMIN' ? true : this.readPermissionValue_(source.canAccessReport, false),
      canViewBookingCreate: role === 'SUPERADMIN' ? true : this.readPermissionValue_(source.canViewBookingCreate, false),
      canViewBookingMine: role === 'SUPERADMIN' ? true : this.readPermissionValue_(source.canViewBookingMine, false),
      canViewDailyBookings: role === 'SUPERADMIN' ? true : this.readPermissionValue_(source.canViewDailyBookings, false),
      canViewFutureBookings: role === 'SUPERADMIN' ? true : this.readPermissionValue_(source.canViewFutureBookings, false),
      canViewReplacementReport: role === 'SUPERADMIN' ? true : this.readPermissionValue_(source.canViewReplacementReport, false),
      canViewOwnRecovery: role === 'SUPERADMIN' ? true : this.readPermissionValue_(source.canViewOwnRecovery, false)
    };
  },

  readPermissionValue_: function (value, fallback) {
    if (value === '' || value == null) {
      return Boolean(fallback);
    }
    return ROOMS_APP.asBoolean(value);
  },

  normalizeOrgUnitPath_: function (value) {
    var normalized = ROOMS_APP.normalizeString(value);
    if (!normalized) {
      return '';
    }
    if (normalized.charAt(0) !== '/') {
      normalized = '/' + normalized;
    }
    normalized = normalized.replace(/\/+$/, '');
    return normalized || '/';
  },

  isOrgUnitPathMatch_: function (userPath, entryPath) {
    var normalizedUserPath = this.normalizeOrgUnitPath_(userPath);
    var normalizedEntryPath = this.normalizeOrgUnitPath_(entryPath);
    if (!normalizedUserPath || !normalizedEntryPath) {
      return false;
    }
    if (normalizedEntryPath === '/') {
      return true;
    }
    return normalizedUserPath === normalizedEntryPath ||
      normalizedUserPath.indexOf(normalizedEntryPath + '/') === 0;
  }
};
