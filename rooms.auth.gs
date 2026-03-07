var ROOMS_APP = ROOMS_APP || {};

ROOMS_APP.Auth = {
  getUserContext: function () {
    var email = ROOMS_APP.getCurrentUserEmail();
    return {
      email: email,
      isAuthenticated: Boolean(email),
      isAdmin: this.isAdmin(email)
    };
  },

  isAdmin: function (email) {
    var normalized = ROOMS_APP.normalizeEmail(email || ROOMS_APP.getCurrentUserEmail());
    var admins = ROOMS_APP.listConfiguredAdmins();

    return admins.some(function (entry) {
      if (entry.charAt(0) === '@') {
        return normalized.slice(-entry.length) === entry;
      }
      return entry === normalized;
    });
  },

  requireAdmin: function () {
    var user = this.getUserContext();
    if (!user.isAdmin) {
      throw new Error('Admin access required.');
    }
    return user;
  }
};
