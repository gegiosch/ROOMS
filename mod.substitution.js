var Mod = Mod || {};

Mod.Substitution = {
  handlePage: function () {
    var actor = ROOMS_APP.Auth.getUserContext();
    var dailyBootstrap = this.buildDailyBootstrap_(actor);
    var dailyHtml = this.renderTemplateToString_('ui.substitution.daily', {
      pageTitle: 'Sostituzioni giornaliere',
      initialModelJson: JSON.stringify(dailyBootstrap),
      initialRoomId: ''
    });
    var reportsHtml = this.renderTemplateToString_('ui.substitution.reports', {
      pageTitle: 'Report sostituzioni',
      viewerMode: 'list',
      viewerEmbedded: true,
      viewerUser: actor,
      reports: ROOMS_APP.Replacements.listVisibleArchivedReports_(),
      report: null
    });

    return renderTemplate_('ui.substitution', {
      pageTitle: 'Sostituzioni',
      initialModelJson: JSON.stringify({
        title: 'Sostituzioni',
        subtitle: 'Modulo unificato delle sostituzioni.',
        user: actor
      }),
      dailyContentHtml: this.extractFragment_(dailyHtml, 'SUBSTITUTION_DAILY_CONTENT_START', 'SUBSTITUTION_DAILY_CONTENT_END'),
      dailyScriptsHtml: this.extractFragment_(dailyHtml, 'SUBSTITUTION_DAILY_SCRIPTS_START', 'SUBSTITUTION_DAILY_SCRIPTS_END'),
      reportsHtml: this.extractFragment_(reportsHtml, 'SUBSTITUTION_REPORTS_FRAGMENT_START', 'SUBSTITUTION_REPORTS_FRAGMENT_END')
    });
  },

  handleDaily: function () {
    var actor = ROOMS_APP.Auth.getUserContext();
    return renderTemplate_('ui.substitution.daily', {
      pageTitle: 'Sostituzioni giornaliere',
      initialModelJson: JSON.stringify(this.buildDailyBootstrap_(actor)),
      initialRoomId: ''
    });
  },

  buildDailyBootstrap_: function (actor) {
    return {
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
  },

  renderTemplateToString_: function (filename, viewModel) {
    var template = HtmlService.createTemplateFromFile(filename);
    Object.keys(viewModel || {}).forEach(function (key) {
      template[key] = viewModel[key];
    });
    return template.evaluate().getContent();
  },

  extractFragment_: function (html, startMarker, endMarker) {
    var source = String(html || '');
    var startToken = '<!-- ' + startMarker + ' -->';
    var endToken = '<!-- ' + endMarker + ' -->';
    var startIndex = source.indexOf(startToken);
    var endIndex = source.indexOf(endToken);
    if (startIndex < 0 || endIndex < 0 || endIndex < startIndex) {
      return '';
    }
    return source.slice(startIndex + startToken.length, endIndex);
  }
};
