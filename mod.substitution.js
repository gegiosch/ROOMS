var Mod = Mod || {};

Mod.Substitution = {
  handlePage: function () {
    return this.handleDaily();
  },

  handleDaily: function () {
    var actor = ROOMS_APP.Auth.getUserContext();
    var bootstrap = {
      generatedAtISO: '',
      date: '',
      refreshSec: 60,
      rotationSec: 15,
      pageCount: 1,
      fullscreenCompactEnabled: true,
      isMonitorMode: false,
      monitorUiScale: Math.max(0.5, ROOMS_APP.getNumberConfig('MONITOR_UI_SCALE', 1) || 1),
      appName: ROOMS_APP.getConfigValue('APP_NAME', 'ROOMS'),
      schoolName: ROOMS_APP.getConfigValue('SCHOOL_NAME', ''),
      user: actor,
      branchOrder: [],
      branchLabels: {},
      pages: []
    };

    return renderTemplate_('ui.substitution.daily', {
      pageTitle: 'Sostituzioni giornaliere',
      initialModelJson: JSON.stringify(bootstrap),
      initialRoomId: ''
    });
  }
};
