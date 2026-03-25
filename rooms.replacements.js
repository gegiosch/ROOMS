var ROOMS_APP = ROOMS_APP || {};

ROOMS_APP.Replacements = {
  REPORT_TYPE_: 'REPLACEMENTS',
  MANUAL_CANDIDATE_VALUE_: '__MANUAL__',
  MAX_NEXT_OPEN_DAY_SCAN_: 60,
  HANDLING_TYPES_: {
    SUBSTITUTION: 'SUBSTITUTION',
    RECOVERY: 'RECOVERY',
    SHIFT_WITHIN_CLASS: 'SHIFT_WITHIN_CLASS',
    CO_TEACHING: 'CO_TEACHING'
  },
  CLASS_HANDLING_TYPES_: {
    NONE: '',
    LATE_ENTRY: 'LATE_ENTRY',
    EARLY_EXIT: 'EARLY_EXIT'
  },
  RECOVERY_STATUSES_: {
    PENDING: 'PENDING',
    RECOVERED: 'RECOVERED'
  },
  SHIFT_SOURCE_VALUE_PREFIX_: 'SHIFT|',
  RECOVERY_SOURCE_VALUE_PREFIX_: 'RECOVERY|',

  ensureSchema_: function () {
    ROOMS_APP.Schema.ensureReplacementClassOut();
    ROOMS_APP.Schema.ensureReplacementDayTeachers();
    ROOMS_APP.Schema.ensureReplacementHourlyAbsences();
    ROOMS_APP.Schema.ensureReplacementAssignments();
    ROOMS_APP.Schema.ensureReplacementLongAssignments();
    ROOMS_APP.Schema.ensureReportRecipients();
    ROOMS_APP.Schema.ensureReportLog();
  },

  getModalModel: function (dateString, draft, options) {
    var actor = ROOMS_APP.Auth.requireCanManageReplacement();
    this.ensureSchema_();
    var settings = options && typeof options === 'object' ? options : {};
    var dateState = this.resolveSelectedDate_(
      dateString,
      settings.allowAutoShift !== false
    );
    var targetDate = dateState.selectedDate;
    var context = this.buildDayContext_(targetDate);
    var normalized = this.normalizeDraft_(
      targetDate,
      draft && typeof draft === 'object' ? draft : context.savedDraft,
      context
    );

    return {
      ok: true,
      user: {
        email: actor.email,
        canManageReplacement: Boolean(actor.canManageReplacement)
      },
      date: targetDate,
      dateState: dateState,
      classes: normalized.classes,
      teachers: normalized.teachers,
      hourlyAbsences: normalized.hourlyAbsences,
      pendingRecoveryRows: normalized.pendingRecoveryRows,
      assignments: normalized.assignments,
      summary: this.buildSummaryFromAssignments_(normalized.assignments, normalized.classes, normalized.teachers),
      longAssignments: context.longAssignments,
      longTeacherOptions: context.longTeacherOptions,
      savedAtISO: context.savedAtISO,
      report: context.reportStatus,
      recipientsConfigured: Boolean(context.recipients.to.length)
    };
  },

  getTeacherDetail: function (dateString, draft, teacherEmail) {
    ROOMS_APP.Auth.requireCanManageReplacement();
    this.ensureSchema_();
    var dateState = this.resolveSelectedDate_(dateString, false);
    if (!dateState.isValid) {
      throw new Error(dateState.message || 'Data non valida.');
    }

    var context = this.buildDayContext_(dateState.selectedDate);
    var normalized = this.normalizeDraft_(dateState.selectedDate, draft, context);
    var teacherKey = this.normalizeTeacherEmail_(teacherEmail);
    var teacher = normalized.teacherMap[teacherKey];
    if (!teacher) {
      throw new Error('Docente non trovato.');
    }

    return {
      date: normalized.date,
      teacherEmail: teacher.teacherEmail,
      teacherName: teacher.teacherName,
      absent: Boolean(teacher.absent),
      accompanist: Boolean(teacher.accompanist),
      accompaniedClasses: teacher.accompaniedClasses.slice(),
      rows: this.buildTeacherDetailRows_(normalized, teacher)
    };
  },

  previewReport: function (dateString, draft) {
    ROOMS_APP.Auth.requireCanManageReplacement();
    this.ensureSchema_();
    var dateState = this.resolveSelectedDate_(dateString, false);
    if (!dateState.isValid) {
      throw new Error(dateState.message || 'Data non valida.');
    }

    var context = this.buildDayContext_(dateState.selectedDate);
    var normalized = this.normalizeDraft_(dateState.selectedDate, draft, context);
    return this.buildReportPayload_(normalized, context.recipients, this.validateDraft_(normalized));
  },

  saveDay: function (dateString, draft) {
    var actor = ROOMS_APP.Auth.requireCanManageReplacement();
    this.ensureSchema_();
    var dateState = this.resolveSelectedDate_(dateString, false);
    if (!dateState.isValid) {
      throw new Error(dateState.message || 'Data non valida.');
    }

    var targetDate = dateState.selectedDate;
    var context = this.buildDayContext_(targetDate);
    var normalized = this.normalizeDraft_(targetDate, draft, context);
    var validationErrors = this.validateDraft_(normalized);
    var savedPreviewPayload;
    if (validationErrors.length) {
      throw new Error(validationErrors.join(' '));
    }
    var nowIso = ROOMS_APP.toIsoDateTime(new Date());
    var updatedBy = actor.email;
    savedPreviewPayload = this.buildReportPayload_(normalized, context.recipients, validationErrors);

    var classRows = normalized.classes.filter(function (entry) {
      return Boolean(entry.isOut);
    }).map(function (entry) {
      return {
        Date: targetDate,
        ClassCode: entry.classCode,
        IsOut: 'TRUE',
        Notes: ROOMS_APP.normalizeString(entry.notes),
        UpdatedAtISO: nowIso,
        UpdatedBy: updatedBy
      };
    });

    var teacherRows = normalized.teachers.filter(function (entry) {
      return entry.absent || entry.accompanist || entry.accompaniedClasses.length || ROOMS_APP.normalizeString(entry.notes);
    }).map(function (entry) {
      return {
        Date: targetDate,
        TeacherEmail: entry.teacherEmail,
        TeacherName: entry.teacherName,
        Absent: entry.absent ? 'TRUE' : 'FALSE',
        Accompanist: entry.accompanist ? 'TRUE' : 'FALSE',
        AccompaniedClasses: entry.accompaniedClasses.join('|'),
        Notes: ROOMS_APP.normalizeString(entry.notes),
        UpdatedAtISO: nowIso,
        UpdatedBy: updatedBy
      };
    });

    var hourlyAbsenceRows = normalized.hourlyAbsences.map(function (entry) {
      return {
        Date: targetDate,
        TeacherEmail: entry.teacherEmail,
        TeacherName: entry.teacherName,
        Period: entry.period,
        Reason: entry.reason,
        RecoveryRequired: entry.recoveryRequired ? 'TRUE' : 'FALSE',
        RecoveryStatus: entry.recoveryStatus,
        RecoveredOnDate: entry.recoveredOnDate,
        RecoveredByAssignmentKey: entry.recoveredByAssignmentKey,
        Notes: ROOMS_APP.normalizeString(entry.notes),
        UpdatedAtISO: nowIso,
        UpdatedBy: updatedBy
      };
    });

    var assignmentRows = normalized.assignments.map(function (entry) {
      return {
        Date: targetDate,
        Period: entry.period,
        ClassCode: entry.classCode,
        OriginalTeacherEmail: entry.originalTeacherEmail,
        OriginalTeacherName: entry.originalTeacherName,
        OriginalStatus: entry.originalStatus,
        ClassHandlingType: entry.classHandlingType,
        HandlingType: entry.handlingType,
        ReplacementTeacherEmail: entry.replacementTeacherEmail,
        ReplacementTeacherName: entry.replacementTeacherName,
        ReplacementSource: entry.replacementSource,
        ReplacementStatus: entry.replacementStatus,
        RecoverySourceDate: entry.recoverySourceDate,
        RecoverySourcePeriod: entry.recoverySourcePeriod,
        ShiftOriginPeriod: entry.shiftOriginPeriod,
        ShiftTargetPeriod: entry.shiftTargetPeriod,
        ShiftTeacherEmail: entry.shiftTeacherEmail,
        ShiftTeacherName: entry.shiftTeacherName,
        Notes: ROOMS_APP.normalizeString(entry.notes),
        UpdatedAtISO: nowIso,
        UpdatedBy: updatedBy
      };
    });

    this.replaceDateRows_(ROOMS_APP.SHEET_NAMES.REPL_CLASS_OUT, targetDate, classRows);
    this.replaceDateRows_(ROOMS_APP.SHEET_NAMES.REPL_DAY_TEACHERS, targetDate, teacherRows);
    this.saveHourlyAbsenceRows_(targetDate, normalized, nowIso, updatedBy);
    this.replaceDateRows_(ROOMS_APP.SHEET_NAMES.REPL_ASSIGNMENTS, targetDate, assignmentRows);

    return {
      ok: true,
      savedAtISO: nowIso,
      model: this.getModalModel(targetDate),
      preview: savedPreviewPayload
    };
  },

  sendReport: function (dateString) {
    var actor = ROOMS_APP.Auth.requireCanManageReplacement();
    this.ensureSchema_();
    var targetDate = ROOMS_APP.toIsoDate(dateString || new Date());
    var context = this.buildDayContext_(targetDate);
    var recipients = context.recipients;
    if (!recipients.to.length) {
      throw new Error('Nessun destinatario TO configurato per il report sostituzioni.');
    }

    var normalized = this.normalizeDraft_(targetDate, context.savedDraft, context);
    var validationErrors = this.validateDraft_(normalized);
    if (validationErrors.length) {
      throw new Error(validationErrors.join(' '));
    }
    var payload = this.buildReportPayload_(normalized, recipients, validationErrors);
    var recipientList = []
      .concat(payload.recipients.to)
      .concat(payload.recipients.cc)
      .concat(payload.recipients.bcc)
      .join(', ');
    var nowIso = ROOMS_APP.toIsoDateTime(new Date());
    var senderInfo = null;

    try {
      senderInfo = ROOMS_APP.Mail.sendReportEmail({
        to: payload.recipients.to,
        cc: payload.recipients.cc,
        bcc: payload.recipients.bcc,
        replyTo: payload.recipients.replyTo,
        subject: payload.subject,
        textBody: payload.textBody,
        htmlBody: payload.htmlBody
      });
      this.appendReportLog_({
        ReportType: this.REPORT_TYPE_,
        ReferenceDate: targetDate,
        SentAtISO: nowIso,
        SentBy: actor.email,
        Recipients: recipientList,
        Subject: payload.subject,
        Status: 'SENT',
        Notes: senderInfo
          ? ('mode=' + senderInfo.senderMode +
            '; name=' + senderInfo.fromName +
            (senderInfo.replyTo ? '; replyTo=' + senderInfo.replyTo : '') +
            (senderInfo.noReply ? '; noReply=true' : ''))
          : ''
      });
    } catch (error) {
      this.appendReportLog_({
        ReportType: this.REPORT_TYPE_,
        ReferenceDate: targetDate,
        SentAtISO: nowIso,
        SentBy: actor.email,
        Recipients: recipientList,
        Subject: payload.subject,
        Status: 'ERROR',
        Notes: String(error && error.message ? error.message : error)
      });
      throw error;
    }

    return {
      ok: true,
      sentAtISO: nowIso,
      report: this.getLatestReportStatus_(targetDate)
    };
  },

  saveLongAssignment: function (payload, referenceDate) {
    var actor = ROOMS_APP.Auth.requireCanManageReplacement();
    this.ensureSchema_();
    var normalized = this.normalizeLongAssignmentInput_(payload || {});
    var rows = this.listLongAssignmentRows_();
    var nowIso = ROOMS_APP.toIsoDateTime(new Date());
    var updatedBy = actor.email;
    var replaced = false;

    this.validateLongAssignment_(normalized, rows, normalized.matchKey);

    var nextRows = rows.map(function (row) {
      if (normalized.matchKey && ROOMS_APP.Replacements.buildLongAssignmentMatchKey_(row) === normalized.matchKey) {
        replaced = true;
        return {
          Enabled: normalized.enabled ? 'TRUE' : 'FALSE',
          OriginalTeacherEmail: normalized.originalTeacherEmail,
          OriginalTeacherName: normalized.originalTeacherName,
          ReplacementTeacherSurname: normalized.replacementTeacherSurname,
          ReplacementTeacherName: normalized.replacementTeacherName,
          ReplacementTeacherDisplayName: normalized.replacementTeacherDisplayName,
          StartDate: normalized.startDate,
          EndDate: normalized.endDate,
          Reason: normalized.reason,
          Notes: normalized.notes,
          UpdatedAtISO: nowIso,
          UpdatedBy: updatedBy
        };
      }
      return row;
    });

    if (!replaced) {
      nextRows.push({
        Enabled: normalized.enabled ? 'TRUE' : 'FALSE',
        OriginalTeacherEmail: normalized.originalTeacherEmail,
        OriginalTeacherName: normalized.originalTeacherName,
        ReplacementTeacherSurname: normalized.replacementTeacherSurname,
        ReplacementTeacherName: normalized.replacementTeacherName,
        ReplacementTeacherDisplayName: normalized.replacementTeacherDisplayName,
        StartDate: normalized.startDate,
        EndDate: normalized.endDate,
        Reason: normalized.reason,
        Notes: normalized.notes,
        UpdatedAtISO: nowIso,
        UpdatedBy: updatedBy
      });
    }

    ROOMS_APP.DB.replaceRows(
      ROOMS_APP.SHEET_NAMES.REPL_LONG_ASSIGNMENTS,
      ROOMS_APP.DB.getHeaders(ROOMS_APP.SHEET_NAMES.REPL_LONG_ASSIGNMENTS),
      nextRows
    );

    return {
      ok: true,
      model: this.getModalModel(referenceDate || normalized.startDate, null, {
        allowAutoShift: false
      })
    };
  },

  toggleLongAssignment: function (matchKey, enabled, referenceDate) {
    var actor = ROOMS_APP.Auth.requireCanManageReplacement();
    this.ensureSchema_();
    var rows = this.listLongAssignmentRows_();
    var nowIso = ROOMS_APP.toIsoDateTime(new Date());
    var found = false;
    var targetRow = null;
    var nextRows = rows.map(function (row) {
      if (ROOMS_APP.Replacements.buildLongAssignmentMatchKey_(row) !== ROOMS_APP.normalizeString(matchKey)) {
        return row;
      }
      found = true;
      targetRow = row;
      var next = ROOMS_APP.Replacements.cloneRow_(row);
      next.Enabled = enabled ? 'TRUE' : 'FALSE';
      next.UpdatedAtISO = nowIso;
      next.UpdatedBy = actor.email;
      return next;
    });

    if (!found) {
      throw new Error('Supplenza lunga non trovata.');
    }
    if (enabled && targetRow) {
      this.validateLongAssignment_({
        matchKey: ROOMS_APP.normalizeString(matchKey),
        enabled: true,
        originalTeacherEmail: targetRow.OriginalTeacherEmail,
        originalTeacherName: targetRow.OriginalTeacherName,
        replacementTeacherSurname: targetRow.ReplacementTeacherSurname,
        replacementTeacherName: targetRow.ReplacementTeacherName,
        replacementTeacherDisplayName: targetRow.ReplacementTeacherDisplayName,
        startDate: targetRow.StartDate,
        endDate: targetRow.EndDate,
        reason: targetRow.Reason,
        notes: targetRow.Notes
      }, rows, matchKey);
    }

    ROOMS_APP.DB.replaceRows(
      ROOMS_APP.SHEET_NAMES.REPL_LONG_ASSIGNMENTS,
      ROOMS_APP.DB.getHeaders(ROOMS_APP.SHEET_NAMES.REPL_LONG_ASSIGNMENTS),
      nextRows
    );

    return {
      ok: true,
      model: this.getModalModel(referenceDate || ROOMS_APP.toIsoDate(new Date()), null, {
        allowAutoShift: false
      })
    };
  },

  deleteLongAssignment: function (matchKey, referenceDate) {
    ROOMS_APP.Auth.requireCanManageReplacement();
    this.ensureSchema_();
    var removed = false;
    var nextRows = this.listLongAssignmentRows_().filter(function (row) {
      var keep = ROOMS_APP.Replacements.buildLongAssignmentMatchKey_(row) !== ROOMS_APP.normalizeString(matchKey);
      if (!keep) {
        removed = true;
      }
      return keep;
    });

    if (!removed) {
      throw new Error('Supplenza lunga non trovata.');
    }

    ROOMS_APP.DB.replaceRows(
      ROOMS_APP.SHEET_NAMES.REPL_LONG_ASSIGNMENTS,
      ROOMS_APP.DB.getHeaders(ROOMS_APP.SHEET_NAMES.REPL_LONG_ASSIGNMENTS),
      nextRows
    );

    return {
      ok: true,
      model: this.getModalModel(referenceDate || ROOMS_APP.toIsoDate(new Date()), null, {
        allowAutoShift: false
      })
    };
  },

  applySavedStateToOccurrences_: function (dateString, occurrences) {
    var targetDate = ROOMS_APP.toIsoDate(dateString || new Date());
    var reflectionState = this.getSavedReflectionState_(targetDate);
    if (!reflectionState.hasSavedState) {
      return (occurrences || []).slice();
    }

    return (occurrences || []).map(function (occurrence) {
      return ROOMS_APP.Replacements.applySavedStateToOccurrence_(occurrence, reflectionState);
    });
  },

  applySavedStateToOccurrence_: function (occurrence, normalized) {
    var next = this.cloneRow_(occurrence);
    var teacherEmail = this.normalizeTeacherEmail_(occurrence && occurrence.TeacherEmail);
    var originalTeacherName = ROOMS_APP.normalizeString(occurrence && occurrence.TeacherName);
    var originalTeacherEmail = teacherEmail;
    if (!originalTeacherEmail) {
      originalTeacherEmail = this.buildTeacherSyntheticEmail_(originalTeacherName);
    }
    var longAssignment = normalized.longAssignmentMap && normalized.longAssignmentMap[originalTeacherEmail]
      ? normalized.longAssignmentMap[originalTeacherEmail]
      : null;
    if (longAssignment) {
      teacherEmail = longAssignment.replacementTeacherEmail;
      next.TeacherEmail = teacherEmail;
      next.TeacherName = longAssignment.replacementTeacherDisplayName;
      next.BookerName = longAssignment.replacementTeacherDisplayName;
      next.DisplayLabel = longAssignment.replacementTeacherDisplayName;
    } else if (!teacherEmail) {
      teacherEmail = originalTeacherEmail;
    }
    var period = this.resolvePeriodFromOccurrence_(occurrence);
    var classCode = ROOMS_APP.normalizeString(occurrence && occurrence.ClassCode).toUpperCase();
    var shiftOriginKey = this.buildAssignmentKey_(period, classCode, teacherEmail);
    var shiftOrigin = normalized.shiftOriginMap && normalized.shiftOriginMap[shiftOriginKey]
      ? normalized.shiftOriginMap[shiftOriginKey]
      : null;
    var assignment = normalized.assignmentMap[this.buildAssignmentKey_(period, classCode, teacherEmail)] || null;

    if (shiftOrigin) {
      next.TeacherEmail = teacherEmail;
      next.TeacherName = 'SPOSTATA';
      next.BookerName = 'SPOSTATA';
      next.DisplayLabel = 'SPOSTATA';
      next.ReplacementStatus = this.HANDLING_TYPES_.SHIFT_WITHIN_CLASS;
      next.IsNonBlocking = true;
      return next;
    }

    if (normalized.classOutSet && normalized.classOutSet[classCode]) {
      next.TeacherEmail = teacherEmail;
      next.TeacherName = 'IN USCITA';
      next.BookerName = 'IN USCITA';
      next.DisplayLabel = 'IN USCITA';
      next.ReplacementStatus = 'IN_USCITA';
      next.IsNonBlocking = true;
      return next;
    }

    if (assignment && assignment.replacementStatus === 'IN_USCITA') {
      next.TeacherEmail = teacherEmail;
      next.TeacherName = 'IN USCITA';
      next.BookerName = 'IN USCITA';
      next.DisplayLabel = 'IN USCITA';
      next.ReplacementStatus = 'IN_USCITA';
      next.IsNonBlocking = true;
      return next;
    }

    if (assignment && assignment.replacementStatus === 'ASSIGNED') {
      next.TeacherEmail = assignment.replacementTeacherEmail;
      next.TeacherName = assignment.replacementTeacherName;
      next.BookerName = assignment.replacementTeacherName;
      next.DisplayLabel = assignment.replacementTeacherName;
      next.ReplacementStatus = 'ASSIGNED';
      next.IsNonBlocking = false;
      return next;
    }

    if (assignment && this.normalizeClassHandlingType_(assignment.classHandlingType) === this.CLASS_HANDLING_TYPES_.LATE_ENTRY) {
      next.TeacherEmail = teacherEmail;
      next.TeacherName = 'ENTRATA POSTICIPATA';
      next.BookerName = 'ENTRATA POSTICIPATA';
      next.DisplayLabel = 'ENTRATA POSTICIPATA';
      next.ReplacementStatus = assignment.replacementStatus || this.CLASS_HANDLING_TYPES_.LATE_ENTRY;
      next.IsNonBlocking = true;
      return next;
    }

    if (assignment && this.normalizeClassHandlingType_(assignment.classHandlingType) === this.CLASS_HANDLING_TYPES_.EARLY_EXIT) {
      next.TeacherEmail = teacherEmail;
      next.TeacherName = 'USCITA ANTICIPATA';
      next.BookerName = 'USCITA ANTICIPATA';
      next.DisplayLabel = 'USCITA ANTICIPATA';
      next.ReplacementStatus = assignment.replacementStatus || this.CLASS_HANDLING_TYPES_.EARLY_EXIT;
      next.IsNonBlocking = true;
      return next;
    }

    if (assignment && this.normalizeHandlingType_(assignment.handlingType) === this.HANDLING_TYPES_.SHIFT_WITHIN_CLASS && assignment.shiftTeacherName) {
      next.TeacherEmail = assignment.shiftTeacherEmail;
      next.TeacherName = assignment.shiftTeacherName;
      next.BookerName = assignment.shiftTeacherName;
      next.DisplayLabel = assignment.shiftTeacherName;
      next.ReplacementStatus = this.HANDLING_TYPES_.SHIFT_WITHIN_CLASS;
      next.IsNonBlocking = false;
      return next;
    }

    if (assignment && assignment.handlingType === this.HANDLING_TYPES_.CO_TEACHING) {
      next.TeacherEmail = teacherEmail;
      next.TeacherName = 'COMPRESENZA';
      next.BookerName = 'COMPRESENZA';
      next.DisplayLabel = 'COMPRESENZA';
      next.ReplacementStatus = assignment.replacementStatus || 'CO_TEACHING';
      next.IsNonBlocking = false;
      return next;
    }

    if (assignment && assignment.replacementStatus === 'TO_ASSIGN') {
      next.TeacherEmail = teacherEmail;
      next.TeacherName = 'DA SOSTITUIRE';
      next.BookerName = 'DA SOSTITUIRE';
      next.DisplayLabel = 'DA SOSTITUIRE';
      next.ReplacementStatus = 'TO_ASSIGN';
      next.IsNonBlocking = false;
      return next;
    }

    next.IsNonBlocking = false;
    return next;
  },

  resolveSelectedDate_: function (dateString, allowAutoShift) {
    var requestedDate = ROOMS_APP.toIsoDate(dateString || new Date());
    var opening = ROOMS_APP.Policy.getDailyOpening(requestedDate);
    if (opening && opening.isOpen) {
      return {
        selectedDate: requestedDate,
        requestedDate: requestedDate,
        isValid: true,
        autoShifted: false,
        message: ''
      };
    }

    if (allowAutoShift) {
      var nextOpenDate = this.findNextOpenDate_(requestedDate);
      if (nextOpenDate) {
        return {
          selectedDate: nextOpenDate,
          requestedDate: requestedDate,
          isValid: true,
          autoShifted: true,
          message: 'La data selezionata era chiusa. Proposta la prima data utile: ' + nextOpenDate + '.'
        };
      }
    }

    return {
      selectedDate: requestedDate,
      requestedDate: requestedDate,
      isValid: false,
      autoShifted: false,
      message: 'La data selezionata ricade in una chiusura o giornata non lavorativa.'
    };
  },

  findNextOpenDate_: function (fromDate) {
    var cursor = ROOMS_APP.combineDateTime(fromDate, '12:00');
    var index;
    for (index = 0; index < this.MAX_NEXT_OPEN_DAY_SCAN_; index += 1) {
      var isoDate = ROOMS_APP.toIsoDate(cursor);
      if (ROOMS_APP.Policy.getDailyOpening(isoDate).isOpen) {
        return isoDate;
      }
      cursor.setDate(cursor.getDate() + 1);
    }
    return '';
  },

  buildDayContext_: function (dateString) {
    var targetDate = ROOMS_APP.toIsoDate(dateString || new Date());
    var activeLongAssignments = this.getActiveLongAssignmentsForDate_(targetDate);
    var activeLongAssignmentMap = this.getActiveLongAssignmentsForDate_(targetDate, true);
    var baseTeachers = this.buildTeacherDayTeachers_(targetDate);
    var savedClassOutRows = this.listRowsForDate_(ROOMS_APP.SHEET_NAMES.REPL_CLASS_OUT, targetDate);
    var savedTeacherRows = this.listRowsForDate_(ROOMS_APP.SHEET_NAMES.REPL_DAY_TEACHERS, targetDate);
    var savedHourlyAbsenceRows = this.listRowsForDate_(ROOMS_APP.SHEET_NAMES.REPL_HOURLY_ABSENCES, targetDate);
    var savedAssignmentRows = this.listRowsForDate_(ROOMS_APP.SHEET_NAMES.REPL_ASSIGNMENTS, targetDate);
    var pendingRecoveryRows = this.listPendingRecoveryRows_(targetDate);
    var periodMap = ROOMS_APP.Timetable.getPeriodTimeMap();
    var teacherMap = {};
    var classSet = {};
    var savedTeacherFallbackMap = {};

    baseTeachers.forEach(function (teacher) {
      teacherMap[teacher.teacherEmail] = teacher;
      Object.keys(teacher.periods || {}).forEach(function (period) {
        var slot = teacher.periods[period];
        if (slot && slot.type === 'CLASS' && slot.classCode) {
          classSet[slot.classCode] = true;
        }
      });
    });

    savedTeacherRows.forEach(function (row) {
      var teacherEmail = ROOMS_APP.Replacements.normalizeTeacherEmail_(row.TeacherEmail);
      if (!teacherEmail) {
        teacherEmail = ROOMS_APP.Replacements.buildTeacherSyntheticEmail_(row.TeacherName);
      }
      if (activeLongAssignmentMap[teacherEmail]) {
        teacherEmail = activeLongAssignmentMap[teacherEmail].replacementTeacherEmail;
      }
      if (!teacherMap[teacherEmail]) {
        teacherMap[teacherEmail] = {
          teacherEmail: teacherEmail,
          teacherName: activeLongAssignmentMap[ROOMS_APP.Replacements.normalizeTeacherEmail_(row.TeacherEmail) || ROOMS_APP.Replacements.buildTeacherSyntheticEmail_(row.TeacherName)]
            ? activeLongAssignmentMap[ROOMS_APP.Replacements.normalizeTeacherEmail_(row.TeacherEmail) || ROOMS_APP.Replacements.buildTeacherSyntheticEmail_(row.TeacherName)].replacementTeacherDisplayName
            : ROOMS_APP.normalizeString(row.TeacherName),
          periods: {}
        };
      }
    });

    savedAssignmentRows.forEach(function (row) {
      var period = ROOMS_APP.normalizeString(row.Period);
      var classCode = ROOMS_APP.normalizeString(row.ClassCode).toUpperCase();
      var originalTeacherEmail = ROOMS_APP.Replacements.normalizeTeacherEmail_(row.OriginalTeacherEmail) || ROOMS_APP.Replacements.buildTeacherSyntheticEmail_(row.OriginalTeacherName);
      var activeLongAssignment = activeLongAssignmentMap[originalTeacherEmail] || null;
      var teacherEmail = activeLongAssignment ? activeLongAssignment.replacementTeacherEmail : originalTeacherEmail;
      var teacherName = activeLongAssignment ? activeLongAssignment.replacementTeacherDisplayName : ROOMS_APP.normalizeString(row.OriginalTeacherName);
      var slot;
      if (!teacherEmail) {
        return;
      }
      if (!teacherMap[teacherEmail]) {
        teacherMap[teacherEmail] = {
          teacherEmail: teacherEmail,
          teacherName: teacherName || teacherEmail,
          periods: {}
        };
      }
      if (period && classCode && !teacherMap[teacherEmail].periods[period]) {
        slot = periodMap[period] || {};
        teacherMap[teacherEmail].periods[period] = {
          period: String(period || ''),
          startTime: slot.startTime || '',
          endTime: slot.endTime || '',
          rawValue: classCode,
          type: 'CLASS',
          classCode: classCode,
          label: classCode
        };
      }
      if (classCode) {
        classSet[classCode] = true;
      }
      if (!savedTeacherFallbackMap[teacherEmail]) {
        savedTeacherFallbackMap[teacherEmail] = {
          TeacherEmail: teacherEmail,
          TeacherName: teacherName,
          Absent: false,
          Accompanist: false,
          AccompaniedClasses: [],
          Notes: ''
        };
      }
      if (ROOMS_APP.normalizeString(row.OriginalStatus) === 'ACCOMPANIST') {
        savedTeacherFallbackMap[teacherEmail].Absent = true;
        savedTeacherFallbackMap[teacherEmail].Accompanist = true;
      } else if (ROOMS_APP.normalizeString(row.OriginalStatus) === 'ABSENT') {
        savedTeacherFallbackMap[teacherEmail].Absent = true;
      }
    });

    savedHourlyAbsenceRows.forEach(function (row) {
      var teacherEmail = ROOMS_APP.Replacements.normalizeTeacherEmail_(row.TeacherEmail);
      var teacherName = ROOMS_APP.normalizeString(row.TeacherName);
      var period = ROOMS_APP.normalizeString(row.Period);
      var slot;
      if (!teacherEmail) {
        teacherEmail = ROOMS_APP.Replacements.buildTeacherSyntheticEmail_(teacherName);
      }
      if (!teacherEmail) {
        return;
      }
      if (!teacherMap[teacherEmail]) {
        teacherMap[teacherEmail] = {
          teacherEmail: teacherEmail,
          teacherName: teacherName || teacherEmail,
          periods: {}
        };
      }
      slot = periodMap[period] || {};
      if (period && !teacherMap[teacherEmail].periods[period]) {
        teacherMap[teacherEmail].periods[period] = {
          period: String(period || ''),
          startTime: slot.startTime || '',
          endTime: slot.endTime || '',
          rawValue: '',
          type: 'FREE',
          classCode: '',
          label: ''
        };
      }
    });

    var classOutMap = {};
    savedClassOutRows.forEach(function (row) {
      if (ROOMS_APP.asBoolean(row.IsOut)) {
        classOutMap[ROOMS_APP.normalizeString(row.ClassCode).toUpperCase()] = {
          notes: ROOMS_APP.normalizeString(row.Notes)
        };
      }
    });

    var teachers = Object.keys(teacherMap).map(function (teacherEmail) {
      var teacher = teacherMap[teacherEmail];
      var saved = ROOMS_APP.Replacements.findSavedTeacherRow_(savedTeacherRows, teacherEmail, teacher.teacherName, activeLongAssignmentMap) || savedTeacherFallbackMap[teacherEmail] || {};
      var accompaniedClasses = ROOMS_APP.Replacements.parsePipeList_(saved.AccompaniedClasses);
      accompaniedClasses.forEach(function (classCode) {
        classSet[classCode] = true;
      });
      return {
        teacherEmail: teacher.teacherEmail,
        teacherName: teacher.teacherName,
        periods: teacher.periods || {},
        absent: ROOMS_APP.asBoolean(saved.Absent),
        accompanist: ROOMS_APP.asBoolean(saved.Accompanist),
        accompaniedClasses: accompaniedClasses,
        notes: ROOMS_APP.normalizeString(saved.Notes)
      };
    }).sort(function (left, right) {
      return left.teacherName.localeCompare(right.teacherName);
    });

    var classes = Object.keys(classSet).sort().map(function (classCode) {
      return {
        classCode: classCode,
        isOut: Boolean(classOutMap[classCode]),
        notes: classOutMap[classCode] ? classOutMap[classCode].notes : ''
      };
    });

    var savedDraft = {
      classOut: classes.map(function (entry) {
        return {
          classCode: entry.classCode,
          isOut: Boolean(entry.isOut),
          notes: entry.notes || ''
        };
      }),
      teachers: teachers.map(function (entry) {
        return {
          teacherEmail: entry.teacherEmail,
          teacherName: entry.teacherName,
          absent: Boolean(entry.absent),
          accompanist: Boolean(entry.accompanist),
          accompaniedClasses: entry.accompaniedClasses.slice(),
          notes: entry.notes || ''
        };
      }),
      hourlyAbsences: savedHourlyAbsenceRows.map(function (row) {
        return ROOMS_APP.Replacements.readHourlyAbsenceRow_(row);
      }),
      assignments: savedAssignmentRows.map(function (row) {
        return ROOMS_APP.Replacements.readAssignmentRow_(row, activeLongAssignmentMap);
      })
    };

    return {
      date: targetDate,
      classes: classes,
      teachers: teachers,
      teacherMap: this.indexByTeacherEmail_(teachers),
      savedDraft: savedDraft,
      savedAtISO: this.computeSavedAtISO_(savedClassOutRows, savedTeacherRows, savedHourlyAbsenceRows, savedAssignmentRows),
      reportStatus: this.getLatestReportStatus_(targetDate),
      recipients: this.getRecipients_(),
      longAssignments: this.buildLongAssignmentsList_(),
      longTeacherOptions: this.listTimetableTeacherDirectory_(),
      activeLongAssignments: activeLongAssignments,
      activeLongAssignmentMap: activeLongAssignmentMap,
      pendingRecoveryRows: pendingRecoveryRows
    };
  },

  normalizeDraft_: function (dateString, draft, context) {
    var targetDate = ROOMS_APP.toIsoDate(dateString || new Date());
    var source = draft || {};
    var baseContext = context || this.buildDayContext_(targetDate);
    var classBaseMap = {};
    var teacherBaseMap = {};
    var hourlySource = Array.isArray(source.hourlyAbsences)
      ? source.hourlyAbsences
      : ((baseContext.savedDraft && baseContext.savedDraft.hourlyAbsences) || []);

    (baseContext.classes || []).forEach(function (entry) {
      classBaseMap[ROOMS_APP.normalizeString(entry.classCode).toUpperCase()] = entry;
    });
    (baseContext.teachers || []).forEach(function (entry) {
      teacherBaseMap[ROOMS_APP.Replacements.normalizeTeacherEmail_(entry.teacherEmail)] = entry;
    });

    var classOutMap = {};
    (Array.isArray(source.classOut) ? source.classOut : []).forEach(function (entry) {
      var classCode = ROOMS_APP.normalizeString(entry && entry.classCode ? entry.classCode : entry && entry.ClassCode).toUpperCase();
      if (!classCode) {
        return;
      }
      classOutMap[classCode] = {
        classCode: classCode,
        isOut: ROOMS_APP.asBoolean(entry && Object.prototype.hasOwnProperty.call(entry, 'isOut') ? entry.isOut : entry && entry.IsOut),
        notes: ROOMS_APP.normalizeString(entry && (entry.notes || entry.Notes))
      };
    });

    Object.keys(classBaseMap).forEach(function (classCode) {
      if (!classOutMap[classCode]) {
        classOutMap[classCode] = {
          classCode: classCode,
          isOut: Boolean(classBaseMap[classCode].isOut),
          notes: ROOMS_APP.normalizeString(classBaseMap[classCode].notes)
        };
      }
    });

    var normalizedClasses = Object.keys(classOutMap).sort().map(function (classCode) {
      return classOutMap[classCode];
    });

    var normalizedTeachers = {};
    Object.keys(teacherBaseMap).forEach(function (teacherEmail) {
      var base = teacherBaseMap[teacherEmail];
      normalizedTeachers[teacherEmail] = {
        teacherEmail: teacherEmail,
        teacherName: base.teacherName,
        periods: base.periods || {},
        absent: Boolean(base.absent),
        accompanist: Boolean(base.accompanist),
        accompaniedClasses: (base.accompaniedClasses || []).slice(),
        notes: ROOMS_APP.normalizeString(base.notes)
      };
    });

    (Array.isArray(source.teachers) ? source.teachers : []).forEach(function (entry) {
      var teacherEmail = ROOMS_APP.Replacements.normalizeTeacherEmail_(entry && (entry.teacherEmail || entry.TeacherEmail));
      var teacherName = ROOMS_APP.normalizeString(entry && (entry.teacherName || entry.TeacherName));
      if (!teacherEmail) {
        if (!teacherName) {
          return;
        }
        teacherEmail = ROOMS_APP.Replacements.buildTeacherSyntheticEmail_(teacherName);
      }
      if (!normalizedTeachers[teacherEmail]) {
        normalizedTeachers[teacherEmail] = {
          teacherEmail: teacherEmail,
          teacherName: teacherName,
          periods: {},
          absent: false,
          accompanist: false,
          accompaniedClasses: [],
          notes: ''
        };
      }
      normalizedTeachers[teacherEmail].teacherName = teacherName || normalizedTeachers[teacherEmail].teacherName;
      normalizedTeachers[teacherEmail].absent = ROOMS_APP.asBoolean(entry && Object.prototype.hasOwnProperty.call(entry, 'absent') ? entry.absent : entry && entry.Absent);
      normalizedTeachers[teacherEmail].accompanist = ROOMS_APP.asBoolean(entry && Object.prototype.hasOwnProperty.call(entry, 'accompanist') ? entry.accompanist : entry && entry.Accompanist);
      normalizedTeachers[teacherEmail].accompaniedClasses = ROOMS_APP.Replacements.parsePipeList_(
        entry && Object.prototype.hasOwnProperty.call(entry, 'accompaniedClasses') ? entry.accompaniedClasses : entry && entry.AccompaniedClasses
      );
      normalizedTeachers[teacherEmail].notes = ROOMS_APP.normalizeString(entry && (entry.notes || entry.Notes));
    });

    var validOutClasses = {};
    normalizedClasses.forEach(function (entry) {
      if (entry.isOut) {
        validOutClasses[entry.classCode] = true;
      }
    });

    var teacherList = Object.keys(normalizedTeachers).map(function (teacherEmail) {
      var teacher = normalizedTeachers[teacherEmail];
      if (teacher.accompanist) {
        teacher.absent = true;
      }
      if (!teacher.absent) {
        teacher.accompanist = false;
        teacher.accompaniedClasses = [];
      }
      teacher.accompaniedClasses = (teacher.accompaniedClasses || []).filter(function (classCode) {
        return Boolean(validOutClasses[ROOMS_APP.normalizeString(classCode).toUpperCase()]);
      }).map(function (classCode) {
        return ROOMS_APP.normalizeString(classCode).toUpperCase();
      });
      return teacher;
    }).sort(function (left, right) {
      return left.teacherName.localeCompare(right.teacherName);
    });

    var teacherMap = this.indexByTeacherEmail_(teacherList);
    var hourlyAbsenceMap = {};
    (hourlySource || []).forEach(function (entry) {
      var normalizedEntry = ROOMS_APP.Replacements.normalizeHourlyAbsenceEntry_(entry, targetDate);
      if (!normalizedEntry.teacherEmail || !normalizedEntry.period) {
        return;
      }
      hourlyAbsenceMap[ROOMS_APP.Replacements.buildHourlyAbsenceKey_(normalizedEntry.teacherEmail, normalizedEntry.period)] = normalizedEntry;
    });
    var hourlyAbsences = Object.keys(hourlyAbsenceMap).map(function (key) {
      return hourlyAbsenceMap[key];
    }).sort(function (left, right) {
      var periodDelta = Number(left.period || 0) - Number(right.period || 0);
      if (periodDelta !== 0) {
        return periodDelta;
      }
      return left.teacherName.localeCompare(right.teacherName);
    });
    var draftAssignmentMap = {};
    (Array.isArray(source.assignments) ? source.assignments : []).forEach(function (entry) {
      var normalizedEntry = ROOMS_APP.Replacements.readAssignmentRow_(entry);
      if (!normalizedEntry.period || !normalizedEntry.classCode || !normalizedEntry.originalTeacherEmail) {
        return;
      }
      draftAssignmentMap[ROOMS_APP.Replacements.buildAssignmentKey_(
        normalizedEntry.period,
        normalizedEntry.classCode,
        normalizedEntry.originalTeacherEmail
      )] = normalizedEntry;
    });

    var normalizedAssignments = this.buildEffectiveAssignments_(
      targetDate,
      normalizedClasses,
      teacherList,
      teacherMap,
      hourlyAbsences,
      draftAssignmentMap
    );
    var pendingRecoveryRows = this.buildPendingRecoveryRows_(
      targetDate,
      hourlyAbsences,
      baseContext.pendingRecoveryRows || [],
      normalizedAssignments
    );

    return {
      date: targetDate,
      classes: normalizedClasses,
      classOutSet: validOutClasses,
      classOutSetKeys: Object.keys(validOutClasses),
      teachers: teacherList,
      teacherMap: teacherMap,
      hourlyAbsences: hourlyAbsences,
      hourlyAbsenceMap: hourlyAbsenceMap,
      assignments: normalizedAssignments,
      assignmentMap: this.indexAssignmentsByKey_(normalizedAssignments),
      pendingRecoveryRows: pendingRecoveryRows,
      activeLongAssignments: baseContext.activeLongAssignments || []
    };
  },

  buildEffectiveAssignments_: function (dateString, classes, teachers, teacherMap, hourlyAbsences, draftAssignmentMap) {
    var classOutSet = {};
    var affectedMap = {};
    classes.forEach(function (entry) {
      if (entry.isOut) {
        classOutSet[entry.classCode] = true;
      }
    });

    teachers.forEach(function (teacher) {
      if (!teacher.absent) {
        return;
      }
      Object.keys(teacher.periods || {}).forEach(function (period) {
        var slot = teacher.periods[period];
        var key;
        if (!slot || slot.type !== 'CLASS' || !slot.classCode) {
          return;
        }
        key = ROOMS_APP.Replacements.buildAssignmentKey_(period, slot.classCode, teacher.teacherEmail);
        affectedMap[key] = {
          teacher: teacher,
          period: period,
          slot: slot,
          originalStatus: teacher.accompanist ? 'ACCOMPANIST' : 'ABSENT',
          hourlyAbsence: null
        };
      });
    });

    (hourlyAbsences || []).forEach(function (entry) {
      var teacher = teacherMap[ROOMS_APP.Replacements.normalizeTeacherEmail_(entry.teacherEmail)] || null;
      var slot = teacher && teacher.periods ? teacher.periods[entry.period] : null;
      var key;
      if (!teacher || teacher.absent || !slot || slot.type !== 'CLASS' || !slot.classCode) {
        return;
      }
      key = ROOMS_APP.Replacements.buildAssignmentKey_(entry.period, slot.classCode, teacher.teacherEmail);
      if (!affectedMap[key]) {
        affectedMap[key] = {
          teacher: teacher,
          period: entry.period,
          slot: slot,
          originalStatus: 'HOURLY_ABSENCE',
          hourlyAbsence: entry
        };
      } else {
        affectedMap[key].hourlyAbsence = entry;
      }
    });

    var assignments = Object.keys(affectedMap).map(function (assignmentKey) {
      var meta = affectedMap[assignmentKey];
      var teacher = meta.teacher;
      var slot = meta.slot || {};
      var draftEntry = draftAssignmentMap[assignmentKey] || {};
      var legacyState = ROOMS_APP.Replacements.normalizeLegacyAssignmentState_(draftEntry);
      var classHandlingType = ROOMS_APP.Replacements.normalizeClassHandlingType_(legacyState.classHandlingType);
      var handlingType = ROOMS_APP.Replacements.normalizeHandlingType_(legacyState.handlingType);
      var isCoveredByOuting = Boolean(
        teacher.accompanist &&
        classOutSet[slot.classCode] &&
        teacher.accompaniedClasses.indexOf(slot.classCode) >= 0
      );
      var replacementTeacherEmail = ROOMS_APP.Replacements.normalizeTeacherEmail_(legacyState.replacementTeacherEmail);
      var replacementTeacherName = ROOMS_APP.normalizeString(legacyState.replacementTeacherName);
      var shiftTeacherEmail = ROOMS_APP.Replacements.normalizeTeacherEmail_(legacyState.shiftTeacherEmail);
      var shiftTeacherName = ROOMS_APP.normalizeString(legacyState.shiftTeacherName);
      var next = {
        date: dateString,
        period: String(meta.period || ''),
        classCode: slot.classCode,
        originalTeacherEmail: teacher.teacherEmail,
        originalTeacherName: teacher.teacherName,
        originalStatus: meta.originalStatus,
        classHandlingType: classHandlingType,
        handlingType: handlingType,
        replacementTeacherEmail: replacementTeacherEmail,
        replacementTeacherName: replacementTeacherName,
        replacementSource: ROOMS_APP.normalizeString(legacyState.replacementSource),
        replacementStatus: '',
        recoverySourceDate: ROOMS_APP.toIsoDate(legacyState.recoverySourceDate),
        recoverySourcePeriod: ROOMS_APP.normalizeString(legacyState.recoverySourcePeriod),
        shiftOriginPeriod: ROOMS_APP.normalizeString(legacyState.shiftOriginPeriod),
        shiftTargetPeriod: ROOMS_APP.normalizeString(legacyState.shiftTargetPeriod || meta.period),
        shiftTeacherEmail: shiftTeacherEmail,
        shiftTeacherName: shiftTeacherName,
        notes: ROOMS_APP.normalizeString(legacyState.notes || (meta.hourlyAbsence && meta.hourlyAbsence.notes)),
        startTime: slot.startTime,
        endTime: slot.endTime
      };

      if (isCoveredByOuting) {
        next.handlingType = ROOMS_APP.Replacements.HANDLING_TYPES_.SUBSTITUTION;
        next.classHandlingType = '';
        next.replacementTeacherEmail = '';
        next.replacementTeacherName = '';
        next.replacementSource = '';
        next.replacementStatus = 'IN_USCITA';
        return next;
      }

      if (classHandlingType) {
        next.replacementStatus = classHandlingType;
        return next;
      }

      if (handlingType === ROOMS_APP.Replacements.HANDLING_TYPES_.RECOVERY) {
        next.replacementStatus = replacementTeacherEmail && next.recoverySourceDate && next.recoverySourcePeriod
          ? 'ASSIGNED'
          : 'TO_ASSIGN';
        return next;
      }

      if (handlingType === ROOMS_APP.Replacements.HANDLING_TYPES_.SHIFT_WITHIN_CLASS) {
        next.replacementStatus = shiftTeacherEmail && next.shiftOriginPeriod
          ? ROOMS_APP.Replacements.HANDLING_TYPES_.SHIFT_WITHIN_CLASS
          : 'TO_ASSIGN';
        if (!next.replacementTeacherEmail) {
          next.replacementTeacherEmail = shiftTeacherEmail;
        }
        if (!next.replacementTeacherName) {
          next.replacementTeacherName = shiftTeacherName;
        }
        return next;
      }

      if (handlingType === ROOMS_APP.Replacements.HANDLING_TYPES_.CO_TEACHING) {
        next.replacementStatus = ROOMS_APP.Replacements.HANDLING_TYPES_.CO_TEACHING;
        return next;
      }

      next.handlingType = ROOMS_APP.Replacements.HANDLING_TYPES_.SUBSTITUTION;
      next.replacementStatus = replacementTeacherEmail || replacementTeacherName ? 'ASSIGNED' : 'TO_ASSIGN';
      if (next.replacementStatus === 'ASSIGNED' && !next.replacementSource) {
        next.replacementSource = 'MANUAL';
      }
      return next;
    });

    return assignments.sort(function (left, right) {
      var periodDelta = Number(left.period || 0) - Number(right.period || 0);
      if (periodDelta !== 0) {
        return periodDelta;
      }
      var teacherDelta = left.originalTeacherName.localeCompare(right.originalTeacherName);
      if (teacherDelta !== 0) {
        return teacherDelta;
      }
      return left.classCode.localeCompare(right.classCode);
    });
  },

  buildTeacherDetailRows_: function (normalized, teacher) {
    var self = this;
    var assignmentsByKey = normalized.assignmentMap;
    var assignedByPeriod = this.buildAssignedTeacherSetByPeriod_(normalized.assignments);
    var rows = [];
    var periodMap = ROOMS_APP.Timetable.getPeriodTimeMap();
    var periodKeys = Object.keys(periodMap).sort(function (left, right) {
      return Number(left) - Number(right);
    });

    periodKeys.forEach(function (period) {
      var slot = teacher.periods[period] || {
        type: 'FREE',
        classCode: '',
        label: '',
        startTime: periodMap[period].startTime,
        endTime: periodMap[period].endTime
      };
      var hourlyAbsence = normalized.hourlyAbsenceMap[self.buildHourlyAbsenceKey_(teacher.teacherEmail, period)] || null;
      var assignmentKey = ROOMS_APP.Replacements.buildAssignmentKey_(period, slot.classCode, teacher.teacherEmail);
      var assignment = assignmentsByKey[assignmentKey] || null;
      var requiresReplacement = Boolean(slot.type === 'CLASS' && (teacher.absent || hourlyAbsence));
      var row = {
        period: period,
        startTime: slot.startTime,
        endTime: slot.endTime,
        slotType: slot.type,
        classCode: slot.classCode,
        label: slot.label,
        canMarkHourlyAbsence: Boolean(slot.type === 'CLASS' && !teacher.absent),
        hourlyAbsence: hourlyAbsence,
        requiresReplacement: requiresReplacement,
        classHandlingType: assignment ? self.normalizeClassHandlingType_(assignment.classHandlingType) : self.CLASS_HANDLING_TYPES_.NONE,
        handlingType: assignment ? self.normalizeHandlingType_(assignment.handlingType) : self.HANDLING_TYPES_.SUBSTITUTION,
        status: assignment ? assignment.replacementStatus : 'NONE',
        assignment: assignment,
        candidates: [],
        manualCandidates: []
      };

      if (!requiresReplacement) {
        if (slot.type === 'CLASS') {
          row.status = 'NO_ACTION';
        } else {
          row.status = 'NONE';
        }
        rows.push(row);
        return;
      }

      if (assignment && assignment.replacementStatus !== 'TO_ASSIGN') {
        row.status = assignment.replacementStatus;
      }
      if (assignment && !row.classHandlingType && assignment.replacementStatus !== 'IN_USCITA') {
        var candidateInfo = ROOMS_APP.Replacements.buildCandidateLists_(normalized, teacher, period, assignment, assignedByPeriod);
        row.candidates = candidateInfo.candidates;
        row.manualCandidates = candidateInfo.manualCandidates;
      }
      rows.push(row);
    });

    return rows;
  },

  buildCandidateLists_: function (normalized, targetTeacher, period, currentAssignment, assignedByPeriod) {
    var handlingType = this.normalizeHandlingType_(currentAssignment && currentAssignment.handlingType);
    if (handlingType === this.HANDLING_TYPES_.RECOVERY) {
      return this.buildRecoveryCandidateLists_(normalized, targetTeacher, period, currentAssignment, assignedByPeriod);
    }
    if (handlingType === this.HANDLING_TYPES_.SHIFT_WITHIN_CLASS) {
      return this.buildShiftCandidateLists_(normalized, targetTeacher, period, currentAssignment, assignedByPeriod);
    }
    return this.buildSubstitutionCandidateLists_(normalized, targetTeacher, period, currentAssignment, assignedByPeriod);
  },

  buildSubstitutionCandidateLists_: function (normalized, targetTeacher, period, currentAssignment, assignedByPeriod) {
    var usedTeachers = this.cloneMap_(assignedByPeriod[period] || {});
    if (currentAssignment && currentAssignment.replacementTeacherEmail) {
      delete usedTeachers[this.normalizeTeacherEmail_(currentAssignment.replacementTeacherEmail)];
    }

    var classOutCandidates = [];
    var pCandidates = [];
    var dCandidates = [];
    var manualCandidates = [];
    var seen = {};

    normalized.teachers.forEach(function (teacher) {
      var teacherEmail = ROOMS_APP.Replacements.normalizeTeacherEmail_(teacher.teacherEmail);
      var slot = teacher.periods[period] || { type: 'FREE', classCode: '', label: '' };
      if (!teacherEmail || teacherEmail === targetTeacher.teacherEmail) {
        return;
      }
      if (teacher.absent) {
        return;
      }
      if (usedTeachers[teacherEmail]) {
        return;
      }
      if (ROOMS_APP.Replacements.hasHourlyAbsenceAtPeriod_(normalized, teacherEmail, period)) {
        return;
      }
      if (slot.type === 'CLASS' && teacher.accompanist && teacher.accompaniedClasses.indexOf(slot.classCode) >= 0 && normalized.classOutSet[slot.classCode]) {
        return;
      }

      if (slot.type === 'CLASS' && normalized.classOutSet[slot.classCode]) {
        classOutCandidates.push(ROOMS_APP.Replacements.buildCandidate_(teacher, 'CLASS_OUT', 'Classe in uscita'));
        seen[teacherEmail] = true;
        return;
      }
      if (slot.type === 'P') {
        pCandidates.push(ROOMS_APP.Replacements.buildCandidate_(teacher, 'P', 'P'));
        seen[teacherEmail] = true;
        return;
      }
      if (slot.type === 'D') {
        dCandidates.push(ROOMS_APP.Replacements.buildCandidate_(teacher, 'D', 'D'));
        seen[teacherEmail] = true;
        return;
      }
      if (slot.type === 'FREE' || slot.type === 'OTHER') {
        manualCandidates.push(ROOMS_APP.Replacements.buildCandidate_(teacher, 'MANUAL', 'Altro docente'));
      }
    });

    var candidates = []
      .concat(classOutCandidates)
      .concat(pCandidates)
      .concat(dCandidates);

    if (manualCandidates.length) {
      candidates.push({
        value: this.MANUAL_CANDIDATE_VALUE_,
        teacherEmail: '',
        teacherName: '',
        source: 'MANUAL',
        label: 'Altro docente...'
      });
    }

    if (currentAssignment && currentAssignment.replacementStatus === 'ASSIGNED') {
      var currentEmail = this.normalizeTeacherEmail_(currentAssignment.replacementTeacherEmail);
      var alreadyPresent = candidates.some(function (entry) {
        return entry.teacherEmail && entry.teacherEmail === currentEmail;
      });
      var currentManualPresent = manualCandidates.some(function (entry) {
        return entry.teacherEmail === currentEmail;
      });
      if (!alreadyPresent && currentAssignment.replacementSource !== 'MANUAL') {
        candidates.unshift({
          value: currentEmail + '|' + currentAssignment.replacementSource,
          teacherEmail: currentEmail,
          teacherName: currentAssignment.replacementTeacherName,
          source: currentAssignment.replacementSource,
          label: currentAssignment.replacementTeacherName + ' (' + currentAssignment.replacementSource + ')'
        });
      }
      if (!currentManualPresent && currentAssignment.replacementSource === 'MANUAL' && currentEmail) {
        manualCandidates.unshift({
          value: currentEmail + '|MANUAL',
          teacherEmail: currentEmail,
          teacherName: currentAssignment.replacementTeacherName,
          source: 'MANUAL',
          label: currentAssignment.replacementTeacherName + ' (Altro docente)'
        });
      }
    }

    return {
      candidates: candidates,
      manualCandidates: manualCandidates
    };
  },

  buildRecoveryCandidateLists_: function (normalized, targetTeacher, period, currentAssignment, assignedByPeriod) {
    var candidates = [];
    var self = this;
    var currentKey = currentAssignment
      ? this.buildHourlyAbsenceDateKey_(
          currentAssignment.recoverySourceDate,
          currentAssignment.replacementTeacherEmail,
          currentAssignment.recoverySourcePeriod
        )
      : '';

    (normalized.pendingRecoveryRows || []).forEach(function (entry) {
      var teacher = normalized.teacherMap[self.normalizeTeacherEmail_(entry.teacherEmail)] || null;
      var key = self.buildHourlyAbsenceDateKey_(entry.date, entry.teacherEmail, entry.period);
      if (!teacher) {
        return;
      }
      if (key !== currentKey && !self.isTeacherEligibleForRecovery_(normalized, teacher, period, assignedByPeriod, currentAssignment)) {
        return;
      }
      candidates.push({
        value: self.RECOVERY_SOURCE_VALUE_PREFIX_ + [
          entry.date,
          entry.period,
          self.normalizeTeacherEmail_(entry.teacherEmail)
        ].join('|'),
        teacherEmail: self.normalizeTeacherEmail_(entry.teacherEmail),
        teacherName: entry.teacherName,
        source: self.HANDLING_TYPES_.RECOVERY,
        recoverySourceDate: entry.date,
        recoverySourcePeriod: entry.period,
        label: entry.teacherName + ' (Recupero ' + self.formatShortDate_(entry.date) + ' ' + entry.period + 'ª ora)'
      });
    });

    if (currentAssignment && currentAssignment.replacementTeacherEmail && currentAssignment.recoverySourceDate && currentAssignment.recoverySourcePeriod && !candidates.length) {
      candidates.push({
        value: this.RECOVERY_SOURCE_VALUE_PREFIX_ + [
          currentAssignment.recoverySourceDate,
          currentAssignment.recoverySourcePeriod,
          this.normalizeTeacherEmail_(currentAssignment.replacementTeacherEmail)
        ].join('|'),
        teacherEmail: this.normalizeTeacherEmail_(currentAssignment.replacementTeacherEmail),
        teacherName: currentAssignment.replacementTeacherName,
        source: this.HANDLING_TYPES_.RECOVERY,
        recoverySourceDate: currentAssignment.recoverySourceDate,
        recoverySourcePeriod: currentAssignment.recoverySourcePeriod,
        label: currentAssignment.replacementTeacherName + ' (Recupero ' + this.formatShortDate_(currentAssignment.recoverySourceDate) + ' ' + currentAssignment.recoverySourcePeriod + 'ª ora)'
      });
    }

    return {
      candidates: candidates,
      manualCandidates: []
    };
  },

  buildShiftCandidateLists_: function (normalized, targetTeacher, period, currentAssignment, assignedByPeriod) {
    var candidates = [];
    var targetClassCode = currentAssignment && currentAssignment.classCode
      ? currentAssignment.classCode
      : '';
    var self = this;
    var currentShiftEmail = currentAssignment ? this.normalizeTeacherEmail_(currentAssignment.shiftTeacherEmail) : '';
    var currentShiftOrigin = currentAssignment ? ROOMS_APP.normalizeString(currentAssignment.shiftOriginPeriod) : '';

    normalized.teachers.forEach(function (teacher) {
      var teacherEmail = self.normalizeTeacherEmail_(teacher.teacherEmail);
      if (!teacherEmail || teacher.absent) {
        return;
      }
      Object.keys(teacher.periods || {}).forEach(function (originPeriod) {
        var originSlot = teacher.periods[originPeriod];
        if (!originSlot || originSlot.type !== 'CLASS' || originSlot.classCode !== targetClassCode) {
          return;
        }
        if (originPeriod === String(period || '')) {
          return;
        }
        if (teacherEmail === currentShiftEmail && originPeriod === currentShiftOrigin) {
          candidates.unshift({
            value: self.SHIFT_SOURCE_VALUE_PREFIX_ + [teacherEmail, originPeriod].join('|'),
            teacherEmail: teacherEmail,
            teacherName: teacher.teacherName,
            source: self.HANDLING_TYPES_.SHIFT_WITHIN_CLASS,
            shiftOriginPeriod: originPeriod,
            shiftTargetPeriod: String(period || ''),
            label: teacher.teacherName + ' (' + self.getShiftNoteLabel_(originPeriod, period) + ')'
          });
          return;
        }
        if (!self.isTeacherEligibleForShift_(normalized, teacher, originPeriod, period, assignedByPeriod, currentAssignment)) {
          return;
        }
        candidates.push({
          value: self.SHIFT_SOURCE_VALUE_PREFIX_ + [teacherEmail, originPeriod].join('|'),
          teacherEmail: teacherEmail,
          teacherName: teacher.teacherName,
          source: self.HANDLING_TYPES_.SHIFT_WITHIN_CLASS,
          shiftOriginPeriod: originPeriod,
          shiftTargetPeriod: String(period || ''),
          label: teacher.teacherName + ' (' + self.getShiftNoteLabel_(originPeriod, period) + ')'
        });
      });
    });

    return {
      candidates: candidates,
      manualCandidates: []
    };
  },

  buildCandidate_: function (teacher, source, reasonLabel) {
    var teacherEmail = this.normalizeTeacherEmail_(teacher.teacherEmail);
    var teacherName = ROOMS_APP.normalizeString(teacher.teacherName);
    return {
      value: teacherEmail + '|' + source,
      teacherEmail: teacherEmail,
      teacherName: teacherName,
      source: source,
      label: teacherName + ' (' + reasonLabel + ')'
    };
  },

  buildAssignedTeacherSetByPeriod_: function (assignments) {
    var byPeriod = {};
    (assignments || []).forEach(function (assignment) {
      var status = ROOMS_APP.normalizeString(assignment.replacementStatus);
      var teacherEmail = ROOMS_APP.Replacements.normalizeTeacherEmail_(assignment.replacementTeacherEmail);
      var shiftTeacherEmail = ROOMS_APP.Replacements.normalizeTeacherEmail_(assignment.shiftTeacherEmail);
      var period;
      if (status === ROOMS_APP.Replacements.HANDLING_TYPES_.SHIFT_WITHIN_CLASS && shiftTeacherEmail) {
        period = ROOMS_APP.normalizeString(assignment.period);
        if (!period) {
          return;
        }
        byPeriod[period] = byPeriod[period] || {};
        byPeriod[period][shiftTeacherEmail] = true;
        return;
      }
      if (status !== 'ASSIGNED') {
        return;
      }
      period = ROOMS_APP.normalizeString(assignment.period);
      if (!period || !teacherEmail) {
        return;
      }
      byPeriod[period] = byPeriod[period] || {};
      byPeriod[period][teacherEmail] = true;
    });
    return byPeriod;
  },

  buildReportPayload_: function (normalized, recipients, validationErrors) {
    var summary = this.buildSummaryFromAssignments_(normalized.assignments, normalized.classes, normalized.teachers);
    var formattedDateLabel = ROOMS_APP.formatItalianExtendedDate(normalized.date);
    var subject = 'Sostituzioni docenti ' + (formattedDateLabel || normalized.date);
    var reportModel = this.buildReportViewModel_(normalized, summary, validationErrors || [], recipients);
    var textBody = this.buildReportTextBody_(reportModel);
    var htmlBody = this.renderReportTemplate_(reportModel);

    return {
      date: normalized.date,
      subject: subject,
      recipients: recipients,
      summary: summary,
      validationErrors: reportModel.validationErrors,
      textBody: textBody,
      htmlBody: htmlBody,
      reportModel: reportModel,
      lines: reportModel.tableRows.map(function (entry) {
        return entry.absentTeacher + ' -> ' + (entry.designatedTeacher || entry.noteLabel || '-') + ' (' + entry.periodSummary + ')';
      })
    };
  },

  buildReportViewModel_: function (normalized, summary, validationErrors, recipients) {
    var formattedDateLabel = ROOMS_APP.formatItalianExtendedDate(normalized.date);
    var recipientConfig = recipients || {};
    var periodColumns = Object.keys(ROOMS_APP.Timetable.getPeriodTimeMap()).sort(function (left, right) {
      return Number(left) - Number(right);
    });
    var classOutList = normalized.classes.filter(function (entry) {
      return entry.isOut;
    }).map(function (entry) {
      return entry.classCode;
    });
    var absentTeachers = normalized.teachers.filter(function (entry) {
      return entry.absent && !entry.accompanist;
    }).map(function (entry) {
      return entry.teacherName;
    });
    var accompanists = normalized.teachers.filter(function (entry) {
      return entry.accompanist;
    }).map(function (entry) {
      return {
        teacherName: entry.teacherName,
        classes: entry.accompaniedClasses.slice()
      };
    });
    var activeLongAssignments = (normalized.activeLongAssignments || []).map(function (entry) {
      return {
        originalTeacherName: entry.originalTeacherName,
        replacementTeacherDisplayName: entry.replacementTeacherDisplayName,
        dateRange: entry.startDate + ' - ' + entry.endDate
      };
    });
    var mainRows = [];
    var teacherGroups = {};
    var teacherOrder = [];

    normalized.assignments.forEach(function (entry) {
      var teacherKey = ROOMS_APP.Replacements.normalizeTeacherEmail_(entry.originalTeacherEmail) || entry.originalTeacherName;
      var descriptor = ROOMS_APP.Replacements.describeReportAssignment_(entry);
      var group;
      var row;
      if (!teacherGroups[teacherKey]) {
        teacherGroups[teacherKey] = [];
        teacherOrder.push(teacherKey);
      }
      group = teacherGroups[teacherKey];
      row = group.filter(function (candidate) {
        return candidate.groupKey === descriptor.groupKey;
      })[0] || null;
      if (!row) {
        row = {
          groupKey: descriptor.groupKey,
          absentTeacher: entry.originalTeacherName,
          designatedTeacher: descriptor.designatedTeacher,
          noteLabel: descriptor.noteLabel,
          noteOnly: descriptor.noteOnly,
          paymentLabel: '',
          status: entry.replacementStatus,
          periodCells: {},
          periodSummary: ''
        };
        group.push(row);
      }
      Object.keys(descriptor.periodCells || {}).forEach(function (periodKey) {
        row.periodCells[periodKey] = descriptor.periodCells[periodKey];
      });
    });

    teacherOrder.forEach(function (teacherKey) {
      var rows = teacherGroups[teacherKey] || [];
      rows.forEach(function (row, index) {
        row.showAbsentTeacher = index === 0;
        row.absentTeacherRowSpan = index === 0 ? rows.length : 0;
        row.periodSummary = periodColumns.map(function (period) {
          return row.periodCells[period] ? (period + 'a: ' + row.periodCells[period]) : '';
        }).filter(function (token) {
          return Boolean(token);
        }).join(', ');
        mainRows.push(row);
      });
    });

    return {
      title: 'SOSTITUZIONI DOCENTI',
      date: normalized.date,
      dateLabel: formattedDateLabel || normalized.date,
      replyToConfigured: Boolean(recipientConfig.replyTo),
      replyToEmail: recipientConfig.replyTo || '',
      validationErrors: (validationErrors || []).slice(),
      summary: summary,
      mainRows: mainRows,
      tableRows: mainRows,
      periodColumns: periodColumns,
      toAssignRows: normalized.assignments.filter(function (entry) {
        return entry.replacementStatus === 'TO_ASSIGN';
      }).map(function (entry) {
        return {
          absentTeacher: entry.originalTeacherName,
          periodClass: entry.period + 'a ora - ' + entry.classCode
        };
      }),
      classOutList: classOutList,
      absentTeachers: absentTeachers,
      accompanists: accompanists,
      activeLongAssignments: activeLongAssignments
    };
  },

  buildReportTextBody_: function (reportModel) {
    var lines = [
      reportModel.title + ' - ' + reportModel.dateLabel,
      ''
    ];
    if (reportModel.validationErrors.length) {
      lines.push('VALIDAZIONI:');
      reportModel.validationErrors.forEach(function (entry) {
        lines.push('- ' + entry);
      });
      lines.push('');
    }
    lines.push('SOSTITUZIONI:');
    if (!reportModel.tableRows.length) {
      lines.push('Nessuna sostituzione richiesta.');
    } else {
      reportModel.tableRows.forEach(function (entry) {
        lines.push(
          '- ' + entry.absentTeacher +
          ' -> ' + (entry.designatedTeacher || '-') +
          ' | ' + (entry.noteLabel || '-') +
          ' | ' + (entry.periodSummary || '-')
        );
      });
    }
    lines.push('');
    lines.push('ASSENTI: ' + (reportModel.absentTeachers.length ? reportModel.absentTeachers.join(', ') : 'Nessuno'));
    lines.push('DOCENTI ACCOMPAGNATORI: ' + (reportModel.accompanists.length ? reportModel.accompanists.map(function (entry) {
      return entry.teacherName + ' [' + (entry.classes.length ? entry.classes.join(', ') : '-') + ']';
    }).join('; ') : 'Nessuno'));
    lines.push('CLASSI IN USCITA: ' + (reportModel.classOutList.length ? reportModel.classOutList.join(', ') : 'Nessuna'));
    if (reportModel.activeLongAssignments.length) {
      lines.push('SUPPLENZE LUNGHE ATTIVE: ' + reportModel.activeLongAssignments.map(function (entry) {
        return entry.originalTeacherName + ' -> ' + entry.replacementTeacherDisplayName + ' (' + entry.dateRange + ')';
      }).join('; '));
    }
    return lines.join('\n');
  },

  describeReportAssignment_: function (entry) {
    var handlingType = this.normalizeHandlingType_(entry.handlingType);
    var classHandlingType = this.normalizeClassHandlingType_(entry.classHandlingType);
    var sourceLabel = this.getReplacementSourceLabel_(entry.replacementSource);
    var noteParts = [];
    var periodCells = {};
    if (entry.replacementStatus === 'IN_USCITA') {
      return {
        groupKey: 'OUTING',
        designatedTeacher: '',
        noteLabel: 'Classe in uscita',
        noteOnly: true,
        periodCells: periodCells
      };
    }
    if (classHandlingType === this.CLASS_HANDLING_TYPES_.LATE_ENTRY) {
      periodCells[entry.period] = 'Entrata';
      noteParts.push('Entrata posticipata');
      if (ROOMS_APP.normalizeString(entry.notes)) {
        noteParts.push(ROOMS_APP.normalizeString(entry.notes));
      }
      return {
        groupKey: 'CLASS|' + classHandlingType + '|' + noteParts.join(' | '),
        designatedTeacher: '',
        noteLabel: noteParts.join(' | '),
        noteOnly: true,
        periodCells: periodCells
      };
    }
    if (classHandlingType === this.CLASS_HANDLING_TYPES_.EARLY_EXIT) {
      periodCells[entry.period] = 'Uscita';
      noteParts.push('Uscita anticipata');
      if (ROOMS_APP.normalizeString(entry.notes)) {
        noteParts.push(ROOMS_APP.normalizeString(entry.notes));
      }
      return {
        groupKey: 'CLASS|' + classHandlingType + '|' + noteParts.join(' | '),
        designatedTeacher: '',
        noteLabel: noteParts.join(' | '),
        noteOnly: true,
        periodCells: periodCells
      };
    }
    if (handlingType === this.HANDLING_TYPES_.CO_TEACHING) {
      periodCells[entry.period] = entry.classCode;
      noteParts.push(this.getHandlingTypeLabel_(handlingType));
      if (ROOMS_APP.normalizeString(entry.notes)) {
        noteParts.push(ROOMS_APP.normalizeString(entry.notes));
      }
      return {
        groupKey: handlingType + '|' + noteParts.join(' | '),
        designatedTeacher: '',
        noteLabel: noteParts.join(' | '),
        noteOnly: true,
        periodCells: periodCells
      };
    }
    if (handlingType === this.HANDLING_TYPES_.SHIFT_WITHIN_CLASS ||
        entry.replacementStatus === this.HANDLING_TYPES_.SHIFT_WITHIN_CLASS) {
      periodCells[entry.period] = entry.classCode;
      noteParts.push(this.getShiftNoteLabel_(entry.shiftOriginPeriod, entry.shiftTargetPeriod || entry.period));
      if (ROOMS_APP.normalizeString(entry.notes)) {
        noteParts.push(ROOMS_APP.normalizeString(entry.notes));
      }
      return {
        groupKey: 'SHIFT|' + ROOMS_APP.normalizeString(entry.shiftTeacherName || entry.replacementTeacherName) + '|' + noteParts.join(' | '),
        designatedTeacher: ROOMS_APP.normalizeString(entry.shiftTeacherName || entry.replacementTeacherName),
        noteLabel: noteParts.join(' | '),
        noteOnly: false,
        periodCells: periodCells
      };
    }
    if (entry.replacementStatus === 'ASSIGNED') {
      periodCells[entry.period] = entry.classCode;
      if (handlingType === this.HANDLING_TYPES_.RECOVERY && entry.recoverySourceDate) {
        noteParts.push('Recupero del ' + this.formatShortDate_(entry.recoverySourceDate));
      } else if (sourceLabel) {
        noteParts.push(sourceLabel);
      }
      if (ROOMS_APP.normalizeString(entry.notes)) {
        noteParts.push(ROOMS_APP.normalizeString(entry.notes));
      }
      return {
        groupKey: handlingType + '|ASSIGNED|' + ROOMS_APP.normalizeString(entry.replacementTeacherName) + '|' + noteParts.join(' | '),
        designatedTeacher: ROOMS_APP.normalizeString(entry.replacementTeacherName),
        noteLabel: noteParts.join(' | '),
        noteOnly: false,
        periodCells: periodCells
      };
    }
    periodCells[entry.period] = entry.classCode;
    return {
      groupKey: 'TO_ASSIGN|' + ROOMS_APP.normalizeString(entry.notes),
      designatedTeacher: 'DA SOSTITUIRE',
      noteLabel: ROOMS_APP.normalizeString(entry.notes),
      noteOnly: false,
      periodCells: periodCells
    };
  },

  getHandlingTypeLabel_: function (handlingType) {
    var normalized = this.normalizeHandlingType_(handlingType);
    if (normalized === this.HANDLING_TYPES_.RECOVERY) {
      return 'Recupero';
    }
    if (normalized === this.HANDLING_TYPES_.SHIFT_WITHIN_CLASS) {
      return 'Spostamento stessa classe';
    }
    if (normalized === this.HANDLING_TYPES_.CO_TEACHING) {
      return 'Compresenza';
    }
    return 'Sostituzione';
  },

  getReplacementSourceLabel_: function (source) {
    var normalized = ROOMS_APP.normalizeString(source).toUpperCase();
    if (normalized === 'CLASS_OUT') {
      return 'Classe in uscita';
    }
    if (normalized === 'MANUAL') {
      return 'Altro docente';
    }
    if (normalized === 'P' || normalized === 'D') {
      return normalized;
    }
    return normalized;
  },

  renderReportTemplate_: function (reportModel) {
    var template = HtmlService.createTemplateFromFile('ui.report.replacements');
    template.report = reportModel;
    return template.evaluate().getContent();
  },

  normalizeHandlingType_: function (value) {
    var normalized = ROOMS_APP.normalizeString(value).toUpperCase();
    if (!normalized) {
      return this.HANDLING_TYPES_.SUBSTITUTION;
    }
    if (normalized === this.HANDLING_TYPES_.RECOVERY ||
        normalized === this.HANDLING_TYPES_.SHIFT_WITHIN_CLASS ||
        normalized === this.HANDLING_TYPES_.CO_TEACHING) {
      return normalized;
    }
    return this.HANDLING_TYPES_.SUBSTITUTION;
  },

  normalizeClassHandlingType_: function (value) {
    var normalized = ROOMS_APP.normalizeString(value).toUpperCase();
    if (normalized === this.CLASS_HANDLING_TYPES_.LATE_ENTRY ||
        normalized === this.CLASS_HANDLING_TYPES_.EARLY_EXIT) {
      return normalized;
    }
    return this.CLASS_HANDLING_TYPES_.NONE;
  },

  validateDraft_: function (normalized) {
    var errors = [];
    var accompanimentByClass = {};
    var assignedByPeriod = this.buildAssignedTeacherSetByPeriod_(normalized.assignments);
    var recoveryLinkMap = {};
    var self = this;

    (normalized && normalized.teachers || []).forEach(function (teacher) {
      if (!teacher || !teacher.accompanist) {
        return;
      }
      (teacher.accompaniedClasses || []).forEach(function (classCode) {
        var normalizedClassCode = ROOMS_APP.normalizeString(classCode).toUpperCase();
        if (!normalizedClassCode) {
          return;
        }
        accompanimentByClass[normalizedClassCode] = accompanimentByClass[normalizedClassCode] || [];
        accompanimentByClass[normalizedClassCode].push(
          ROOMS_APP.normalizeString(teacher.teacherName) || ROOMS_APP.normalizeString(teacher.teacherEmail)
        );
      });
    });

    (normalized && normalized.classes || []).forEach(function (entry) {
      if (!entry || !entry.isOut) {
        return;
      }
      var classCode = ROOMS_APP.normalizeString(entry.classCode).toUpperCase();
      if (!classCode) {
        return;
      }
      if (!accompanimentByClass[classCode] || !accompanimentByClass[classCode].length) {
        errors.push('La classe ' + classCode + ' è in uscita ma non ha alcun docente accompagnatore associato.');
      }
    });

    (normalized && normalized.hourlyAbsences || []).forEach(function (entry) {
      var teacher = normalized.teacherMap[self.normalizeTeacherEmail_(entry.teacherEmail)] || null;
      var slot = teacher && teacher.periods ? teacher.periods[entry.period] : null;
      if (!teacher || !slot || slot.type !== 'CLASS' || !slot.classCode) {
        errors.push('Assenza oraria non valida per ' + (entry.teacherName || entry.teacherEmail) + ' alla ' + entry.period + 'ª ora.');
      }
    });

    (normalized && normalized.assignments || []).forEach(function (entry) {
      var teacher = normalized.teacherMap[self.normalizeTeacherEmail_(entry.originalTeacherEmail)] || null;
      var candidateInfo;
      var matchFound;
      if (!teacher || self.normalizeClassHandlingType_(entry.classHandlingType) || entry.replacementStatus === 'IN_USCITA' || entry.replacementStatus === 'TO_ASSIGN') {
        return;
      }
      candidateInfo = self.buildCandidateLists_(normalized, teacher, entry.period, entry, assignedByPeriod);
      if (self.normalizeHandlingType_(entry.handlingType) === self.HANDLING_TYPES_.SHIFT_WITHIN_CLASS ||
          entry.replacementStatus === self.HANDLING_TYPES_.SHIFT_WITHIN_CLASS) {
        matchFound = candidateInfo.candidates.some(function (candidate) {
          return candidate.teacherEmail === self.normalizeTeacherEmail_(entry.shiftTeacherEmail) &&
            ROOMS_APP.normalizeString(candidate.shiftOriginPeriod) === ROOMS_APP.normalizeString(entry.shiftOriginPeriod);
        });
        if (!matchFound) {
          errors.push('Spostamento non valido per ' + entry.classCode + ' alla ' + entry.period + 'ª ora.');
        }
        return;
      }
      if (self.normalizeHandlingType_(entry.handlingType) === self.HANDLING_TYPES_.RECOVERY) {
        var recoveryKey = self.buildHourlyAbsenceDateKey_(
          entry.recoverySourceDate,
          entry.replacementTeacherEmail,
          entry.recoverySourcePeriod
        );
        if (entry.recoverySourceDate === normalized.date) {
          errors.push('Il recupero per ' + entry.classCode + ' alla ' + entry.period + 'ª ora deve riferirsi a un altro giorno.');
          return;
        }
        if (recoveryKey && recoveryLinkMap[recoveryKey]) {
          errors.push('Il recupero del ' + self.formatShortDate_(entry.recoverySourceDate) + ' ' + entry.recoverySourcePeriod + 'ª ora è già collegato a un\'altra assegnazione.');
          return;
        }
        matchFound = candidateInfo.candidates.some(function (candidate) {
          return candidate.teacherEmail === self.normalizeTeacherEmail_(entry.replacementTeacherEmail) &&
            ROOMS_APP.toIsoDate(candidate.recoverySourceDate) === ROOMS_APP.toIsoDate(entry.recoverySourceDate) &&
            ROOMS_APP.normalizeString(candidate.recoverySourcePeriod) === ROOMS_APP.normalizeString(entry.recoverySourcePeriod);
        });
        if (!matchFound) {
          errors.push('Recupero non valido per ' + entry.classCode + ' alla ' + entry.period + 'ª ora.');
          return;
        }
        recoveryLinkMap[recoveryKey] = true;
        return;
      }
      if (entry.replacementStatus === 'ASSIGNED') {
        matchFound = candidateInfo.candidates.some(function (candidate) {
          return candidate.teacherEmail === self.normalizeTeacherEmail_(entry.replacementTeacherEmail);
        }) || candidateInfo.manualCandidates.some(function (candidate) {
          return candidate.teacherEmail === self.normalizeTeacherEmail_(entry.replacementTeacherEmail);
        });
        if (!matchFound) {
          errors.push('Sostituzione non valida per ' + entry.classCode + ' alla ' + entry.period + 'ª ora.');
        }
      }
    });

    return errors;
  },

  buildSummaryFromAssignments_: function (assignments, classes, teachers) {
    var summary = {
      classOutCount: 0,
      absentCount: 0,
      accompanistCount: 0,
      hourlyAbsenceCount: 0,
      assignedCount: 0,
      recoveryCount: 0,
      shiftCount: 0,
      toAssignCount: 0,
      inUscitaCount: 0
    };

    (classes || []).forEach(function (entry) {
      if (entry.isOut) {
        summary.classOutCount += 1;
      }
    });
    (teachers || []).forEach(function (entry) {
      if (entry.absent) {
        summary.absentCount += 1;
      }
      if (entry.accompanist) {
        summary.accompanistCount += 1;
      }
    });
    (assignments || []).forEach(function (entry) {
      if (entry.originalStatus === 'HOURLY_ABSENCE') {
        summary.hourlyAbsenceCount += 1;
      }
      if (entry.replacementStatus === 'ASSIGNED') {
        summary.assignedCount += 1;
        if (ROOMS_APP.Replacements.normalizeHandlingType_(entry.handlingType) === ROOMS_APP.Replacements.HANDLING_TYPES_.RECOVERY) {
          summary.recoveryCount += 1;
        }
      } else if (entry.replacementStatus === ROOMS_APP.Replacements.HANDLING_TYPES_.SHIFT_WITHIN_CLASS) {
        summary.assignedCount += 1;
        summary.shiftCount += 1;
      } else if (entry.replacementStatus === 'TO_ASSIGN') {
        summary.toAssignCount += 1;
      } else if (entry.replacementStatus === 'IN_USCITA') {
        summary.inUscitaCount += 1;
      }
    });
    return summary;
  },

  getLatestReportStatus_: function (dateString) {
    var targetDate = ROOMS_APP.toIsoDate(dateString || new Date());
    var rows = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.REPORT_LOG).filter(function (row) {
      return ROOMS_APP.normalizeString(row.ReportType).toUpperCase() === ROOMS_APP.Replacements.REPORT_TYPE_ &&
        ROOMS_APP.toIsoDate(row.ReferenceDate) === targetDate;
    });
    rows = ROOMS_APP.sortBy(rows, ['SentAtISO']);
    var latest = rows.length ? rows[rows.length - 1] : null;
    return {
      sentAtISO: latest ? ROOMS_APP.normalizeString(latest.SentAtISO) : '',
      sentBy: latest ? ROOMS_APP.normalizeString(latest.SentBy) : '',
      status: latest ? ROOMS_APP.normalizeString(latest.Status) : '',
      subject: latest ? ROOMS_APP.normalizeString(latest.Subject) : '',
      recipients: latest ? ROOMS_APP.normalizeString(latest.Recipients) : '',
      notes: latest ? ROOMS_APP.normalizeString(latest.Notes) : ''
    };
  },

  getSavedReflectionState_: function (dateString) {
    var targetDate = ROOMS_APP.toIsoDate(dateString || new Date());
    var classOutSet = {};
    var assignmentMap = {};
    var shiftOriginMap = {};
    var activeLongAssignmentMap = this.getActiveLongAssignmentsForDate_(targetDate, true);

    this.listRowsForDate_(ROOMS_APP.SHEET_NAMES.REPL_CLASS_OUT, targetDate).forEach(function (row) {
      if (!ROOMS_APP.asBoolean(row.IsOut)) {
        return;
      }
      classOutSet[ROOMS_APP.normalizeString(row.ClassCode).toUpperCase()] = true;
    });

    this.listRowsForDate_(ROOMS_APP.SHEET_NAMES.REPL_ASSIGNMENTS, targetDate).forEach(function (row) {
      var assignment = ROOMS_APP.Replacements.readAssignmentRow_(row, activeLongAssignmentMap);
      if (!assignment.period || !assignment.classCode || !assignment.originalTeacherEmail) {
        return;
      }
      assignmentMap[ROOMS_APP.Replacements.buildAssignmentKey_(
        assignment.period,
        assignment.classCode,
        assignment.originalTeacherEmail
      )] = assignment;
      if (ROOMS_APP.Replacements.normalizeHandlingType_(assignment.handlingType) === ROOMS_APP.Replacements.HANDLING_TYPES_.SHIFT_WITHIN_CLASS &&
          assignment.shiftTeacherEmail &&
          assignment.shiftOriginPeriod) {
        shiftOriginMap[ROOMS_APP.Replacements.buildAssignmentKey_(
          assignment.shiftOriginPeriod,
          assignment.classCode,
          assignment.shiftTeacherEmail
        )] = assignment;
      }
    });

    return {
      hasSavedState: Boolean(
        Object.keys(classOutSet).length ||
        Object.keys(assignmentMap).length ||
        Object.keys(shiftOriginMap).length ||
        Object.keys(activeLongAssignmentMap).length
      ),
      classOutSet: classOutSet,
      classOutSetKeys: Object.keys(classOutSet),
      assignmentMap: assignmentMap,
      shiftOriginMap: shiftOriginMap,
      longAssignmentMap: activeLongAssignmentMap
    };
  },

  getRecipients_: function () {
    var rows = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.REPORT_RECIPIENTS).filter(function (row) {
      return ROOMS_APP.asBoolean(row.Enabled) &&
        ROOMS_APP.normalizeString(row.ReportType).toUpperCase() === ROOMS_APP.Replacements.REPORT_TYPE_;
    });
    var output = {
      to: [],
      cc: [],
      bcc: [],
      replyTo: ''
    };

    rows.forEach(function (row) {
      var email = ROOMS_APP.normalizeEmail(row.Email);
      var type = ROOMS_APP.normalizeString(row.RecipientType).toUpperCase();
      if (!email || !ROOMS_APP.Replacements.isValidRecipientEmail_(email)) {
        return;
      }
      if (type === 'CC') {
        output.cc.push(email);
      } else if (type === 'BCC') {
        output.bcc.push(email);
      } else if (type === 'REPLY') {
        if (!output.replyTo) {
          output.replyTo = email;
        }
      } else {
        output.to.push(email);
      }
    });
    return output;
  },

  appendReportLog_: function (row) {
    ROOMS_APP.DB.appendRows(ROOMS_APP.SHEET_NAMES.REPORT_LOG, [row]);
  },

  replaceDateRows_: function (sheetName, targetDate, replacementRows) {
    var rows = ROOMS_APP.DB.readRows(sheetName).filter(function (row) {
      return ROOMS_APP.toIsoDate(row.Date || row.ReferenceDate) !== targetDate;
    });
    var headers = ROOMS_APP.DB.getHeaders(sheetName);
    ROOMS_APP.DB.replaceRows(sheetName, headers, rows.concat(replacementRows || []));
  },

  listRowsForDate_: function (sheetName, targetDate) {
    return ROOMS_APP.DB.readRows(sheetName).filter(function (row) {
      return ROOMS_APP.toIsoDate(row.Date || row.ReferenceDate) === targetDate;
    });
  },

  saveHourlyAbsenceRows_: function (targetDate, normalized, nowIso, updatedBy) {
    var sheetName = ROOMS_APP.SHEET_NAMES.REPL_HOURLY_ABSENCES;
    var headers = ROOMS_APP.DB.getHeaders(sheetName);
    var allRows = ROOMS_APP.DB.readRows(sheetName).map(function (row) {
      return ROOMS_APP.Replacements.readHourlyAbsenceRow_(row);
    });
    var existingByKey = {};
    var assignmentKeysForDate = {};
    var recoveredByKey = {};
    var nextRows;

    allRows.forEach(function (row) {
      existingByKey[ROOMS_APP.Replacements.buildHourlyAbsenceDateKey_(row.date, row.teacherEmail, row.period)] = row;
    });

    (normalized.assignments || []).forEach(function (entry) {
      var assignmentLinkKey = ROOMS_APP.Replacements.buildAssignmentLinkKey_(
        targetDate,
        entry.period,
        entry.classCode,
        entry.originalTeacherEmail
      );
      assignmentKeysForDate[assignmentLinkKey] = true;
      if (ROOMS_APP.Replacements.normalizeHandlingType_(entry.handlingType) !== ROOMS_APP.Replacements.HANDLING_TYPES_.RECOVERY ||
          entry.replacementStatus !== 'ASSIGNED' ||
          !entry.recoverySourceDate ||
          !entry.recoverySourcePeriod ||
          !entry.replacementTeacherEmail) {
        return;
      }
      recoveredByKey[ROOMS_APP.Replacements.buildHourlyAbsenceDateKey_(
        entry.recoverySourceDate,
        entry.replacementTeacherEmail,
        entry.recoverySourcePeriod
      )] = assignmentLinkKey;
    });

    nextRows = allRows.filter(function (row) {
      return row.date !== targetDate;
    }).map(function (row) {
      var key = ROOMS_APP.Replacements.buildHourlyAbsenceDateKey_(row.date, row.teacherEmail, row.period);
      var next = ROOMS_APP.Replacements.cloneRow_(row);
      if (assignmentKeysForDate[next.recoveredByAssignmentKey] && !recoveredByKey[key]) {
        next.recoveryStatus = next.recoveryRequired ? ROOMS_APP.Replacements.RECOVERY_STATUSES_.PENDING : '';
        next.recoveredOnDate = '';
        next.recoveredByAssignmentKey = '';
      }
      if (recoveredByKey[key]) {
        next.recoveryStatus = ROOMS_APP.Replacements.RECOVERY_STATUSES_.RECOVERED;
        next.recoveredOnDate = targetDate;
        next.recoveredByAssignmentKey = recoveredByKey[key];
      }
      return next;
    });

    (normalized.hourlyAbsences || []).forEach(function (entry) {
      var key = ROOMS_APP.Replacements.buildHourlyAbsenceDateKey_(targetDate, entry.teacherEmail, entry.period);
      var existing = existingByKey[key] || {};
      var next = {
        date: targetDate,
        teacherEmail: entry.teacherEmail,
        teacherName: entry.teacherName,
        period: entry.period,
        reason: entry.reason,
        recoveryRequired: Boolean(entry.recoveryRequired),
        recoveryStatus: '',
        recoveredOnDate: '',
        recoveredByAssignmentKey: '',
        notes: ROOMS_APP.normalizeString(entry.notes)
      };
      if (recoveredByKey[key]) {
        next.recoveryStatus = ROOMS_APP.Replacements.RECOVERY_STATUSES_.RECOVERED;
        next.recoveredOnDate = targetDate;
        next.recoveredByAssignmentKey = recoveredByKey[key];
      } else if (next.recoveryRequired) {
        next.recoveryStatus = existing.recoveryStatus === ROOMS_APP.Replacements.RECOVERY_STATUSES_.RECOVERED
          ? ROOMS_APP.Replacements.RECOVERY_STATUSES_.RECOVERED
          : ROOMS_APP.Replacements.RECOVERY_STATUSES_.PENDING;
        next.recoveredOnDate = existing.recoveryStatus === ROOMS_APP.Replacements.RECOVERY_STATUSES_.RECOVERED
          ? existing.recoveredOnDate
          : '';
        next.recoveredByAssignmentKey = existing.recoveryStatus === ROOMS_APP.Replacements.RECOVERY_STATUSES_.RECOVERED
          ? existing.recoveredByAssignmentKey
          : '';
      }
      nextRows.push(next);
    });

    ROOMS_APP.DB.replaceRows(sheetName, headers, nextRows.map(function (row) {
      return {
        Date: row.date,
        TeacherEmail: row.teacherEmail,
        TeacherName: row.teacherName,
        Period: row.period,
        Reason: row.reason,
        RecoveryRequired: row.recoveryRequired ? 'TRUE' : 'FALSE',
        RecoveryStatus: row.recoveryStatus,
        RecoveredOnDate: row.recoveredOnDate,
        RecoveredByAssignmentKey: row.recoveredByAssignmentKey,
        Notes: row.notes,
        UpdatedAtISO: nowIso || row.updatedAtISO || '',
        UpdatedBy: updatedBy || row.updatedBy || ''
      };
    }));
  },

  listPendingRecoveryRows_: function (excludeDate) {
    var excluded = ROOMS_APP.toIsoDate(excludeDate || '');
    return ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.REPL_HOURLY_ABSENCES).map(function (row) {
      return ROOMS_APP.Replacements.readHourlyAbsenceRow_(row);
    }).filter(function (row) {
      if (!row.recoveryRequired) {
        return false;
      }
      if (row.recoveryStatus === ROOMS_APP.Replacements.RECOVERY_STATUSES_.RECOVERED) {
        return false;
      }
      if (excluded && row.date === excluded) {
        return false;
      }
      return Boolean(row.date && row.teacherEmail && row.period);
    }).sort(function (left, right) {
      if (left.date !== right.date) {
        return left.date.localeCompare(right.date);
      }
      return Number(left.period || 0) - Number(right.period || 0);
    });
  },

  buildPendingRecoveryRows_: function (targetDate, hourlyAbsences, persistedRows, assignments) {
    var map = {};
    (persistedRows || []).forEach(function (entry) {
      map[ROOMS_APP.Replacements.buildHourlyAbsenceDateKey_(entry.date, entry.teacherEmail, entry.period)] = entry;
    });
    (assignments || []).forEach(function (entry) {
      if (ROOMS_APP.Replacements.normalizeHandlingType_(entry.handlingType) !== ROOMS_APP.Replacements.HANDLING_TYPES_.RECOVERY ||
          entry.replacementStatus !== 'ASSIGNED') {
        return;
      }
      delete map[ROOMS_APP.Replacements.buildHourlyAbsenceDateKey_(
        entry.recoverySourceDate,
        entry.replacementTeacherEmail,
        entry.recoverySourcePeriod
      )];
    });
    return Object.keys(map).map(function (key) {
      return map[key];
    }).sort(function (left, right) {
      if (left.date !== right.date) {
        return left.date.localeCompare(right.date);
      }
      return Number(left.period || 0) - Number(right.period || 0);
    });
  },

  readHourlyAbsenceRow_: function (row) {
    return {
      date: ROOMS_APP.toIsoDate(row && row.Date),
      teacherEmail: this.normalizeTeacherEmail_(row && row.TeacherEmail) || this.buildTeacherSyntheticEmail_(row && row.TeacherName),
      teacherName: ROOMS_APP.normalizeString(row && row.TeacherName),
      period: ROOMS_APP.normalizeString(row && row.Period),
      reason: ROOMS_APP.normalizeString(row && row.Reason),
      recoveryRequired: ROOMS_APP.asBoolean(row && row.RecoveryRequired),
      recoveryStatus: ROOMS_APP.normalizeString(row && row.RecoveryStatus),
      recoveredOnDate: ROOMS_APP.toIsoDate(row && row.RecoveredOnDate),
      recoveredByAssignmentKey: ROOMS_APP.normalizeString(row && row.RecoveredByAssignmentKey),
      notes: ROOMS_APP.normalizeString(row && row.Notes),
      updatedAtISO: ROOMS_APP.normalizeString(row && row.UpdatedAtISO),
      updatedBy: this.normalizeTeacherEmail_(row && row.UpdatedBy)
    };
  },

  normalizeHourlyAbsenceEntry_: function (entry, targetDate) {
    var teacherEmail = this.normalizeTeacherEmail_(entry && (entry.teacherEmail || entry.TeacherEmail));
    var teacherName = ROOMS_APP.normalizeString(entry && (entry.teacherName || entry.TeacherName));
    if (!teacherEmail && teacherName) {
      teacherEmail = this.buildTeacherSyntheticEmail_(teacherName);
    }
    return {
      date: ROOMS_APP.toIsoDate(entry && (entry.date || entry.Date || targetDate)),
      teacherEmail: teacherEmail,
      teacherName: teacherName,
      period: ROOMS_APP.normalizeString(entry && (entry.period || entry.Period)),
      reason: ROOMS_APP.normalizeString(entry && (entry.reason || entry.Reason)),
      recoveryRequired: ROOMS_APP.asBoolean(entry && Object.prototype.hasOwnProperty.call(entry, 'recoveryRequired') ? entry.recoveryRequired : entry && entry.RecoveryRequired),
      recoveryStatus: ROOMS_APP.normalizeString(entry && (entry.recoveryStatus || entry.RecoveryStatus)) || this.RECOVERY_STATUSES_.PENDING,
      recoveredOnDate: ROOMS_APP.toIsoDate(entry && (entry.recoveredOnDate || entry.RecoveredOnDate)),
      recoveredByAssignmentKey: ROOMS_APP.normalizeString(entry && (entry.recoveredByAssignmentKey || entry.RecoveredByAssignmentKey)),
      notes: ROOMS_APP.normalizeString(entry && (entry.notes || entry.Notes))
    };
  },

  readAssignmentRow_: function (row, activeLongAssignmentMap) {
    var originalTeacherEmail = this.normalizeTeacherEmail_(row && (row.OriginalTeacherEmail || row.originalTeacherEmail));
    var originalTeacherName = ROOMS_APP.normalizeString(row && (row.OriginalTeacherName || row.originalTeacherName));
    var activeLongAssignment;
    var legacyState;
    if (!originalTeacherEmail && originalTeacherName) {
      originalTeacherEmail = this.buildTeacherSyntheticEmail_(originalTeacherName);
    }
    activeLongAssignment = activeLongAssignmentMap && activeLongAssignmentMap[originalTeacherEmail]
      ? activeLongAssignmentMap[originalTeacherEmail]
      : null;
    if (activeLongAssignment) {
      originalTeacherEmail = activeLongAssignment.replacementTeacherEmail;
      originalTeacherName = activeLongAssignment.replacementTeacherDisplayName;
    }
    legacyState = this.normalizeLegacyAssignmentState_(row);
    return {
      period: ROOMS_APP.normalizeString(row && (row.Period || row.period)),
      classCode: ROOMS_APP.normalizeString(row && (row.ClassCode || row.classCode)).toUpperCase(),
      originalTeacherEmail: originalTeacherEmail,
      originalTeacherName: originalTeacherName,
      originalStatus: ROOMS_APP.normalizeString(row && (row.OriginalStatus || row.originalStatus)),
      classHandlingType: legacyState.classHandlingType,
      handlingType: legacyState.handlingType,
      replacementTeacherEmail: this.normalizeTeacherEmail_(row && (row.ReplacementTeacherEmail || row.replacementTeacherEmail)),
      replacementTeacherName: ROOMS_APP.normalizeString(row && (row.ReplacementTeacherName || row.replacementTeacherName)),
      replacementSource: ROOMS_APP.normalizeString(row && (row.ReplacementSource || row.replacementSource)),
      replacementStatus: ROOMS_APP.normalizeString(row && (row.ReplacementStatus || row.replacementStatus)),
      recoverySourceDate: ROOMS_APP.toIsoDate(row && (row.RecoverySourceDate || row.recoverySourceDate)),
      recoverySourcePeriod: ROOMS_APP.normalizeString(row && (row.RecoverySourcePeriod || row.recoverySourcePeriod)),
      shiftOriginPeriod: ROOMS_APP.normalizeString(row && (row.ShiftOriginPeriod || row.shiftOriginPeriod)),
      shiftTargetPeriod: ROOMS_APP.normalizeString(row && (row.ShiftTargetPeriod || row.shiftTargetPeriod)),
      shiftTeacherEmail: this.normalizeTeacherEmail_(row && (row.ShiftTeacherEmail || row.shiftTeacherEmail)),
      shiftTeacherName: ROOMS_APP.normalizeString(row && (row.ShiftTeacherName || row.shiftTeacherName)),
      notes: ROOMS_APP.normalizeString(row && (row.Notes || row.notes))
    };
  },

  normalizeLegacyAssignmentState_: function (entry) {
    var rawHandling = ROOMS_APP.normalizeString(entry && (entry.HandlingType || entry.handlingType)).toUpperCase();
    var rawClassHandling = ROOMS_APP.normalizeString(entry && (entry.ClassHandlingType || entry.classHandlingType)).toUpperCase();
    if (!rawClassHandling && (rawHandling === this.CLASS_HANDLING_TYPES_.LATE_ENTRY || rawHandling === this.CLASS_HANDLING_TYPES_.EARLY_EXIT)) {
      rawClassHandling = rawHandling;
      rawHandling = this.HANDLING_TYPES_.SUBSTITUTION;
    }
    return {
      classHandlingType: this.normalizeClassHandlingType_(rawClassHandling),
      handlingType: this.normalizeHandlingType_(rawHandling),
      replacementTeacherEmail: entry && (entry.ReplacementTeacherEmail || entry.replacementTeacherEmail),
      replacementTeacherName: entry && (entry.ReplacementTeacherName || entry.replacementTeacherName),
      replacementSource: entry && (entry.ReplacementSource || entry.replacementSource),
      notes: entry && (entry.Notes || entry.notes),
      recoverySourceDate: entry && (entry.RecoverySourceDate || entry.recoverySourceDate),
      recoverySourcePeriod: entry && (entry.RecoverySourcePeriod || entry.recoverySourcePeriod),
      shiftOriginPeriod: entry && (entry.ShiftOriginPeriod || entry.shiftOriginPeriod),
      shiftTargetPeriod: entry && (entry.ShiftTargetPeriod || entry.shiftTargetPeriod),
      shiftTeacherEmail: entry && (entry.ShiftTeacherEmail || entry.shiftTeacherEmail),
      shiftTeacherName: entry && (entry.ShiftTeacherName || entry.shiftTeacherName)
    };
  },

  buildHourlyAbsenceKey_: function (teacherEmail, period) {
    return [
      this.normalizeTeacherEmail_(teacherEmail),
      ROOMS_APP.normalizeString(period)
    ].join('|');
  },

  buildHourlyAbsenceDateKey_: function (dateString, teacherEmail, period) {
    return [
      ROOMS_APP.toIsoDate(dateString),
      this.normalizeTeacherEmail_(teacherEmail),
      ROOMS_APP.normalizeString(period)
    ].join('|');
  },

  buildAssignmentLinkKey_: function (dateString, period, classCode, teacherEmail) {
    return [
      ROOMS_APP.toIsoDate(dateString),
      ROOMS_APP.normalizeString(period),
      ROOMS_APP.normalizeString(classCode).toUpperCase(),
      this.normalizeTeacherEmail_(teacherEmail)
    ].join('|');
  },

  hasHourlyAbsenceAtPeriod_: function (normalized, teacherEmail, period) {
    return Boolean((normalized.hourlyAbsenceMap || {})[this.buildHourlyAbsenceKey_(teacherEmail, period)]);
  },

  isTeacherEligibleForRecovery_: function (normalized, teacher, period, assignedByPeriod, currentAssignment) {
    var teacherEmail = this.normalizeTeacherEmail_(teacher && teacher.teacherEmail);
    var usedTeachers = this.cloneMap_(assignedByPeriod[period] || {});
    var slot = teacher && teacher.periods ? teacher.periods[period] : null;
    if (!teacherEmail || !teacher || teacher.absent || teacher.accompanist) {
      return false;
    }
    if (currentAssignment && this.normalizeTeacherEmail_(currentAssignment.replacementTeacherEmail) === teacherEmail) {
      delete usedTeachers[teacherEmail];
    }
    if (usedTeachers[teacherEmail] || this.hasHourlyAbsenceAtPeriod_(normalized, teacherEmail, period)) {
      return false;
    }
    return !slot || slot.type === 'FREE';
  },

  isTeacherEligibleForShift_: function (normalized, teacher, originPeriod, targetPeriod, assignedByPeriod, currentAssignment) {
    var teacherEmail = this.normalizeTeacherEmail_(teacher && teacher.teacherEmail);
    var usedTeachers = this.cloneMap_(assignedByPeriod[targetPeriod] || {});
    var targetSlot = teacher && teacher.periods ? teacher.periods[targetPeriod] : null;
    var self = this;
    if (!teacherEmail || !teacher || teacher.absent || teacher.accompanist) {
      return false;
    }
    if (currentAssignment && this.normalizeTeacherEmail_(currentAssignment.shiftTeacherEmail) === teacherEmail) {
      delete usedTeachers[teacherEmail];
    }
    if (usedTeachers[teacherEmail] || this.hasHourlyAbsenceAtPeriod_(normalized, teacherEmail, targetPeriod)) {
      return false;
    }
    if (targetSlot && targetSlot.type !== 'FREE') {
      return false;
    }
    return !(normalized.assignments || []).some(function (assignment) {
      return assignment !== currentAssignment &&
        self.normalizeHandlingType_(assignment.handlingType) === self.HANDLING_TYPES_.SHIFT_WITHIN_CLASS &&
        self.normalizeTeacherEmail_(assignment.shiftTeacherEmail) === teacherEmail &&
        ROOMS_APP.normalizeString(assignment.shiftOriginPeriod) === ROOMS_APP.normalizeString(originPeriod);
    });
  },

  getShiftNoteLabel_: function (originPeriod, targetPeriod) {
    var origin = Number(originPeriod || 0);
    var target = Number(targetPeriod || 0);
    if (origin && target && target < origin) {
      return 'Anticipa alla ' + targetPeriod + 'ª ora';
    }
    return 'Posticipa alla ' + targetPeriod + 'ª ora';
  },

  formatShortDate_: function (dateString) {
    var normalized = ROOMS_APP.toIsoDate(dateString);
    if (!normalized || normalized.length !== 10) {
      return normalized;
    }
    return normalized.slice(8, 10) + '/' + normalized.slice(5, 7);
  },

  listLongAssignmentRows_: function () {
    return ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.REPL_LONG_ASSIGNMENTS).map(function (row) {
      return {
        Enabled: ROOMS_APP.asBoolean(row.Enabled) ? 'TRUE' : 'FALSE',
        OriginalTeacherEmail: ROOMS_APP.Replacements.normalizeTeacherEmail_(row.OriginalTeacherEmail) || ROOMS_APP.Replacements.buildTeacherSyntheticEmail_(row.OriginalTeacherName),
        OriginalTeacherName: ROOMS_APP.normalizeString(row.OriginalTeacherName),
        ReplacementTeacherSurname: ROOMS_APP.normalizeString(row.ReplacementTeacherSurname),
        ReplacementTeacherName: ROOMS_APP.normalizeString(row.ReplacementTeacherName),
        ReplacementTeacherDisplayName: ROOMS_APP.Replacements.buildReplacementTeacherDisplayName_(
          row.ReplacementTeacherSurname,
          row.ReplacementTeacherName,
          row.ReplacementTeacherDisplayName
        ),
        StartDate: ROOMS_APP.toIsoDate(row.StartDate),
        EndDate: ROOMS_APP.toIsoDate(row.EndDate),
        Reason: ROOMS_APP.normalizeString(row.Reason),
        Notes: ROOMS_APP.normalizeString(row.Notes),
        UpdatedAtISO: ROOMS_APP.normalizeString(row.UpdatedAtISO),
        UpdatedBy: ROOMS_APP.normalizeEmail(row.UpdatedBy)
      };
    }).filter(function (row) {
      return Boolean(row.OriginalTeacherEmail && row.OriginalTeacherName && row.ReplacementTeacherDisplayName && row.StartDate && row.EndDate);
    });
  },

  buildLongAssignmentsList_: function () {
    var today = ROOMS_APP.toIsoDate(ROOMS_APP.Auth.getEffectiveNow());
    return this.listLongAssignmentRows_().map(function (row) {
      var enabled = ROOMS_APP.asBoolean(row.Enabled);
      var status = 'DISABLED';
      if (enabled) {
        if (today < row.StartDate) {
          status = 'FUTURE';
        } else if (today > row.EndDate) {
          status = 'EXPIRED';
        } else {
          status = 'ACTIVE';
        }
      }
      return {
        matchKey: ROOMS_APP.Replacements.buildLongAssignmentMatchKey_(row),
        enabled: enabled,
        originalTeacherEmail: row.OriginalTeacherEmail,
        originalTeacherName: row.OriginalTeacherName,
        replacementTeacherSurname: row.ReplacementTeacherSurname,
        replacementTeacherName: row.ReplacementTeacherName,
        replacementTeacherDisplayName: row.ReplacementTeacherDisplayName,
        startDate: row.StartDate,
        endDate: row.EndDate,
        reason: row.Reason,
        notes: row.Notes,
        status: status,
        updatedAtISO: row.UpdatedAtISO,
        updatedBy: row.UpdatedBy
      };
    }).sort(function (left, right) {
      var teacherDelta = left.originalTeacherName.localeCompare(right.originalTeacherName);
      if (teacherDelta !== 0) {
        return teacherDelta;
      }
      return left.startDate.localeCompare(right.startDate);
    });
  },

  getActiveLongAssignmentsForDate_: function (dateString, asMap) {
    var targetDate = ROOMS_APP.toIsoDate(dateString || new Date());
    var activeRows = this.listLongAssignmentRows_().filter(function (row) {
      return ROOMS_APP.asBoolean(row.Enabled) && row.StartDate <= targetDate && row.EndDate >= targetDate;
    }).map(function (row) {
      return {
        originalTeacherEmail: row.OriginalTeacherEmail,
        originalTeacherName: row.OriginalTeacherName,
        replacementTeacherSurname: row.ReplacementTeacherSurname,
        replacementTeacherName: row.ReplacementTeacherName,
        replacementTeacherDisplayName: row.ReplacementTeacherDisplayName,
        replacementTeacherEmail: ROOMS_APP.Replacements.buildTeacherSyntheticEmail_(row.ReplacementTeacherDisplayName),
        startDate: row.StartDate,
        endDate: row.EndDate,
        reason: row.Reason,
        notes: row.Notes
      };
    });
    if (!asMap) {
      return activeRows;
    }
    return activeRows.reduce(function (map, row) {
      map[row.originalTeacherEmail] = row;
      return map;
    }, {});
  },

  listTimetableTeacherDirectory_: function () {
    var sheetName = ROOMS_APP.Timetable.getConfiguredSourceSheetName_(
      ROOMS_APP.Timetable.CONFIG_DOCENTI_SHEET_KEY_,
      ROOMS_APP.Timetable.DEFAULT_DOCENTI_SHEET_
    );
    var sheet = ROOMS_APP.DB.getSheet(sheetName);
    if (!sheet) {
      return [];
    }

    var values = ROOMS_APP.Timetable.readSheetDisplayValues_(sheet);
    var columnMeta = ROOMS_APP.Timetable.detectColumnMeta_(values);
    var rows = [];
    var seen = {};
    var rowIndex;

    for (rowIndex = columnMeta.dataStartRow; rowIndex < values.length; rowIndex += 1) {
      var rowLabel = ROOMS_APP.normalizeString(values[rowIndex] && values[rowIndex][0]);
      if (!ROOMS_APP.Timetable.isRowLabelData_(rowLabel, 'classroom')) {
        continue;
      }
      var teacherEmail = this.buildTeacherSyntheticEmail_(rowLabel);
      if (!teacherEmail || seen[teacherEmail]) {
        continue;
      }
      seen[teacherEmail] = true;
      rows.push({
        teacherEmail: teacherEmail,
        teacherName: rowLabel
      });
    }

    return rows.sort(function (left, right) {
      return left.teacherName.localeCompare(right.teacherName);
    });
  },

  buildTeacherDayTeachers_: function (dateString) {
    var targetDate = ROOMS_APP.toIsoDate(dateString || new Date());
    var weekday = ROOMS_APP.getWeekdayName(targetDate);
    var activeLongAssignments = this.getActiveLongAssignmentsForDate_(targetDate, true);
    var sheetName = ROOMS_APP.Timetable.getConfiguredSourceSheetName_(
      ROOMS_APP.Timetable.CONFIG_DOCENTI_SHEET_KEY_,
      ROOMS_APP.Timetable.DEFAULT_DOCENTI_SHEET_
    );
    var sheet = ROOMS_APP.DB.getSheet(sheetName);
    if (!sheet) {
      return [];
    }

    var values = ROOMS_APP.Timetable.readSheetDisplayValues_(sheet);
    var columnMeta = ROOMS_APP.Timetable.detectColumnMeta_(values);
    var teacherMap = {};
    var rowIndex;

    for (rowIndex = columnMeta.dataStartRow; rowIndex < values.length; rowIndex += 1) {
      var rowLabel = ROOMS_APP.normalizeString(values[rowIndex] && values[rowIndex][0]);
      if (!ROOMS_APP.Timetable.isRowLabelData_(rowLabel, 'classroom')) {
        continue;
      }

      var originalTeacherEmail = this.buildTeacherSyntheticEmail_(rowLabel);
      var longAssignment = activeLongAssignments[originalTeacherEmail] || null;
      var teacherEmail = longAssignment ? longAssignment.replacementTeacherEmail : originalTeacherEmail;
      var teacherName = longAssignment ? longAssignment.replacementTeacherDisplayName : rowLabel;
      var teacher = teacherMap[teacherEmail] || {
        teacherEmail: teacherEmail,
        teacherName: teacherName,
        periods: {},
        originalTeacherEmails: {}
      };
      if (originalTeacherEmail) {
        teacher.originalTeacherEmails[originalTeacherEmail] = rowLabel;
      }

      columnMeta.usableColumns.forEach(function (column) {
        var meta = columnMeta.columns[column] || {};
        if (meta.weekday !== weekday) {
          return;
        }
        var rawValue = ROOMS_APP.normalizeString(values[rowIndex][column]);
        teacher.periods[String(meta.period)] = ROOMS_APP.Replacements.classifyTeacherSlotValue_(rawValue, meta.period, meta.startTime, meta.endTime);
      });

      teacherMap[teacherEmail] = teacher;
    }

    return Object.keys(teacherMap).map(function (teacherEmail) {
      return teacherMap[teacherEmail];
    }).sort(function (left, right) {
      return left.teacherName.localeCompare(right.teacherName);
    });
  },

  classifyTeacherSlotValue_: function (rawValue, period, startTime, endTime) {
    var normalized = ROOMS_APP.normalizeString(rawValue).toUpperCase();
    var classCode = ROOMS_APP.Timetable.extractClassCode_(normalized);
    var type = 'FREE';
    var label = '';

    if (classCode) {
      type = 'CLASS';
      label = classCode;
    } else if (normalized === 'P') {
      type = 'P';
      label = 'P';
    } else if (normalized === 'D') {
      type = 'D';
      label = 'D';
    } else if (normalized) {
      type = 'OTHER';
      label = normalized;
    }

    return {
      period: String(period || ''),
      startTime: startTime || '',
      endTime: endTime || '',
      rawValue: normalized,
      type: type,
      classCode: classCode,
      label: label
    };
  },

  resolvePeriodFromOccurrence_: function (occurrence) {
    var startTime = ROOMS_APP.toTimeString(occurrence && occurrence.StartTime);
    var periodMap = ROOMS_APP.Timetable.getPeriodTimeMap();
    var keys = Object.keys(periodMap);
    var index;
    for (index = 0; index < keys.length; index += 1) {
      if (periodMap[keys[index]].startTime === startTime) {
        return keys[index];
      }
    }
    return '';
  },

  findSavedTeacherRow_: function (rows, teacherEmail, teacherName, activeLongAssignments) {
    var normalizedEmail = this.normalizeTeacherEmail_(teacherEmail);
    var normalizedName = ROOMS_APP.normalizeString(teacherName);
    return (rows || []).filter(function (row) {
      var rowEmail = ROOMS_APP.Replacements.normalizeTeacherEmail_(row.TeacherEmail);
      if (!rowEmail) {
        rowEmail = ROOMS_APP.Replacements.buildTeacherSyntheticEmail_(row.TeacherName);
      }
      var activeLong = activeLongAssignments && activeLongAssignments[rowEmail] ? activeLongAssignments[rowEmail] : null;
      if (activeLong) {
        rowEmail = activeLong.replacementTeacherEmail;
      }
      if (rowEmail && rowEmail === normalizedEmail) {
        return true;
      }
      return activeLong
        ? activeLong.replacementTeacherDisplayName === normalizedName
        : ROOMS_APP.normalizeString(row.TeacherName) === normalizedName;
    })[0] || null;
  },

  computeSavedAtISO_: function () {
    var latest = '';
    Array.prototype.slice.call(arguments).forEach(function (rows) {
      (rows || []).forEach(function (row) {
        var value = ROOMS_APP.normalizeString(row.UpdatedAtISO);
        if (value && (!latest || value > latest)) {
          latest = value;
        }
      });
    });
    return latest;
  },

  indexByTeacherEmail_: function (teachers) {
    var map = {};
    (teachers || []).forEach(function (entry) {
      map[ROOMS_APP.Replacements.normalizeTeacherEmail_(entry.teacherEmail)] = entry;
    });
    return map;
  },

  indexAssignmentsByKey_: function (assignments) {
    var map = {};
    (assignments || []).forEach(function (entry) {
      map[ROOMS_APP.Replacements.buildAssignmentKey_(entry.period, entry.classCode, entry.originalTeacherEmail)] = entry;
    });
    return map;
  },

  buildAssignmentKey_: function (period, classCode, teacherEmail) {
    return [
      ROOMS_APP.normalizeString(period),
      ROOMS_APP.normalizeString(classCode).toUpperCase(),
      this.normalizeTeacherEmail_(teacherEmail)
    ].join('|');
  },

  normalizeLongAssignmentInput_: function (payload) {
    var originalTeacherName = ROOMS_APP.normalizeString(payload && payload.originalTeacherName);
    var originalTeacherEmail = this.normalizeTeacherEmail_(payload && payload.originalTeacherEmail);
    if (!originalTeacherEmail && originalTeacherName) {
      originalTeacherEmail = this.buildTeacherSyntheticEmail_(originalTeacherName);
    }
    return {
      matchKey: ROOMS_APP.normalizeString(payload && payload.matchKey),
      enabled: payload && Object.prototype.hasOwnProperty.call(payload, 'enabled') ? ROOMS_APP.asBoolean(payload.enabled) : true,
      originalTeacherEmail: originalTeacherEmail,
      originalTeacherName: originalTeacherName,
      replacementTeacherSurname: ROOMS_APP.normalizeString(payload && payload.replacementTeacherSurname),
      replacementTeacherName: ROOMS_APP.normalizeString(payload && payload.replacementTeacherName),
      replacementTeacherDisplayName: this.buildReplacementTeacherDisplayName_(
        payload && payload.replacementTeacherSurname,
        payload && payload.replacementTeacherName,
        payload && payload.replacementTeacherDisplayName
      ),
      startDate: ROOMS_APP.toIsoDate(payload && payload.startDate),
      endDate: ROOMS_APP.toIsoDate(payload && payload.endDate),
      reason: ROOMS_APP.normalizeString(payload && payload.reason),
      notes: ROOMS_APP.normalizeString(payload && payload.notes)
    };
  },

  validateLongAssignment_: function (candidate, rows, skipMatchKey) {
    if (!candidate.originalTeacherEmail || !candidate.originalTeacherName) {
      throw new Error('Selezionare il docente originale.');
    }
    if (!candidate.replacementTeacherSurname || !candidate.replacementTeacherName) {
      throw new Error('Inserire cognome e nome del supplente.');
    }
    if (!candidate.startDate || !candidate.endDate) {
      throw new Error('Inserire un intervallo date valido.');
    }
    if (candidate.startDate > candidate.endDate) {
      throw new Error('La data iniziale deve essere precedente o uguale alla data finale.');
    }

    var overlap = (rows || []).filter(function (row) {
      if (!ROOMS_APP.asBoolean(row.Enabled) || !candidate.enabled) {
        return false;
      }
      if (ROOMS_APP.Replacements.buildLongAssignmentMatchKey_(row) === ROOMS_APP.normalizeString(skipMatchKey)) {
        return false;
      }
      if (ROOMS_APP.Replacements.normalizeTeacherEmail_(row.OriginalTeacherEmail) !== candidate.originalTeacherEmail) {
        return false;
      }
      return !(candidate.endDate < row.StartDate || candidate.startDate > row.EndDate);
    })[0] || null;

    if (overlap) {
      throw new Error('Intervallo sovrapposto per il docente originale selezionato.');
    }
  },

  buildReplacementTeacherDisplayName_: function (surname, name, fallbackDisplayName) {
    var normalizedSurname = ROOMS_APP.normalizeString(surname).toUpperCase();
    var normalizedName = ROOMS_APP.normalizeString(name).toUpperCase();
    var display = [normalizedSurname, normalizedName].filter(function (token) {
      return Boolean(token);
    }).join(' ');
    return display || ROOMS_APP.normalizeString(fallbackDisplayName).toUpperCase();
  },

  buildLongAssignmentMatchKey_: function (row) {
    return [
      this.normalizeTeacherEmail_(row && row.OriginalTeacherEmail),
      ROOMS_APP.normalizeString(row && row.ReplacementTeacherDisplayName),
      ROOMS_APP.toIsoDate(row && row.StartDate),
      ROOMS_APP.toIsoDate(row && row.EndDate),
      ROOMS_APP.normalizeString(row && row.UpdatedAtISO)
    ].join('|');
  },

  parsePipeList_: function (value) {
    if (Array.isArray(value)) {
      return value.map(function (entry) {
        return ROOMS_APP.normalizeString(entry).toUpperCase();
      }).filter(function (entry) {
        return Boolean(entry);
      });
    }
    return ROOMS_APP.normalizeString(value)
      .split('|')
      .map(function (entry) {
        return ROOMS_APP.normalizeString(entry).toUpperCase();
      })
      .filter(function (entry) {
        return Boolean(entry);
      });
  },

  buildTeacherSyntheticEmail_: function (teacherName) {
    var base = ROOMS_APP.slugify(teacherName || '');
    return base ? (base + '@rooms.local') : '';
  },

  normalizeTeacherEmail_: function (value) {
    return ROOMS_APP.normalizeEmail(value);
  },

  isValidRecipientEmail_: function (value) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(ROOMS_APP.normalizeEmail(value));
  },

  cloneMap_: function (value) {
    var output = {};
    Object.keys(value || {}).forEach(function (key) {
      output[key] = value[key];
    });
    return output;
  },

  cloneRow_: function (row) {
    var output = {};
    Object.keys(row || {}).forEach(function (key) {
      output[key] = row[key];
    });
    return output;
  },

  escapeHtml_: function (value) {
    return String(value || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }
};
