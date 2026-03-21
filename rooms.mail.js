var ROOMS_APP = ROOMS_APP || {};

ROOMS_APP.Mail = {
  getReportSenderConfig_: function () {
    var mode = ROOMS_APP.normalizeString(ROOMS_APP.getConfigValue('REPORT_SENDER_MODE', 'NOREPLY')).toUpperCase();
    if (mode !== 'DEFAULT' && mode !== 'REPLY_TO' && mode !== 'NOREPLY') {
      mode = 'NOREPLY';
    }

    return {
      fromName: ROOMS_APP.normalizeString(
        ROOMS_APP.getConfigValue('REPORT_FROM_NAME', ROOMS_APP.getConfigValue('APP_NAME', 'ROOMS'))
      ) || ROOMS_APP.getConfigValue('APP_NAME', 'ROOMS'),
      replyTo: ROOMS_APP.normalizeString(ROOMS_APP.getConfigValue('REPORT_REPLY_TO', '')),
      mode: mode
    };
  },

  buildMailAppPayload_: function (message, senderConfig) {
    var payload = {
      to: (message.to || []).join(','),
      cc: (message.cc || []).join(','),
      bcc: (message.bcc || []).join(','),
      subject: message.subject,
      body: message.textBody,
      htmlBody: message.htmlBody,
      name: senderConfig.fromName
    };

    if (senderConfig.mode === 'NOREPLY') {
      payload.noReply = true;
    } else if (senderConfig.mode === 'REPLY_TO' && senderConfig.replyTo) {
      payload.replyTo = senderConfig.replyTo;
    }

    return payload;
  },

  sendReportEmail: function (message) {
    var senderConfig = this.getReportSenderConfig_();
    var payload = this.buildMailAppPayload_(message, senderConfig);
    MailApp.sendEmail(payload);
    return {
      senderMode: senderConfig.mode,
      fromName: senderConfig.fromName,
      replyTo: senderConfig.mode === 'REPLY_TO' ? senderConfig.replyTo : '',
      noReply: Boolean(senderConfig.mode === 'NOREPLY')
    };
  }
};
