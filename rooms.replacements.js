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
  ABSENCE_MODES_: {
    DAY: 'DAY',
    PLANNED: 'PLANNED'
  },
  ABSENCE_TYPES_: {
    DAILY: 'DAILY',
    HOURLY_PERMISSION: 'HOURLY_PERMISSION',
    SERVICE_OUT_OF_CLASS: 'SERVICE_OUT_OF_CLASS'
  },
  SERVICE_OUT_OF_CLASS_: {
    GENERAL: 'SERVIZIO_FUORI_AULA',
    ASSEMBLY: 'ASSEMBLEA',
    ALTERNATIVA: 'ALTERNATIVA',
    ACCOMPANIMENT: 'ACCOMPAGNAMENTO'
  },
  ABSENCE_WORKLOAD_DERIVED_VERSION_: '2026-05-10.1',
  MIN_CLASS_PRESENCE_PERIODS_: 4,
  TRIP_TYPES_: {
    DAILY: 'DAILY',
    MULTI_DAY: 'MULTI_DAY',
    HOURLY: 'HOURLY'
  },
  TRIP_ROLE_: 'ACCOMPANIST',
  SHIFT_SOURCE_VALUE_PREFIX_: 'SHIFT|',
  RECOVERY_SOURCE_VALUE_PREFIX_: 'RECOVERY|',

  ensureSchema_: function () {
    ROOMS_APP.Schema.ensureReplacementClassOut();
    ROOMS_APP.Schema.ensureReplacementDayTeachers();
    ROOMS_APP.Schema.ensureReplacementFieldTrips();
    ROOMS_APP.Schema.ensureReplacementFieldTripTeachers();
    ROOMS_APP.Schema.ensureReplacementAbsences();
    ROOMS_APP.Schema.ensureReplacementHourlyAbsences();
    ROOMS_APP.Schema.ensureReplacementAssignments();
    ROOMS_APP.Schema.ensureReplacementLongAssignments();
    ROOMS_APP.Schema.ensureReplacementDailyCache();
    ROOMS_APP.Schema.ensureReportRecipients();
    ROOMS_APP.Schema.ensureReportLog();
    ROOMS_APP.Schema.ensureReportArchive();
    ROOMS_APP.Schema.ensureReportArchiveHistory();
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
    var context = this.buildDayContext_(targetDate, settings);
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
      trips: context.trips,
      tripClassOptions: context.tripClassOptions,
      tripTeacherOptionsByClass: context.tripTeacherOptionsByClass,
      tripDayState: context.tripDayState,
      absenceRegistry: context.absenceRegistryState || {},
      editorDataLoaded: Boolean(context.editorDataLoaded),
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

    var context = this.buildDayContext_(dateState.selectedDate, {
      includeEditorData: false
    });
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
      serviceOutOfClass: Boolean(teacher.serviceOutOfClass),
      rows: this.buildTeacherDetailRows_(normalized, teacher)
    };
  },

  getAbsenceRegistryModel: function (referenceDate) {
    var actor = ROOMS_APP.Auth.requireCanManageReplacement();
    var dateState;
    var teacherOptions;
    var rows;
    this.ensureSchema_();
    dateState = this.resolveSelectedDate_(referenceDate || ROOMS_APP.toIsoDate(ROOMS_APP.Auth.getEffectiveNow()), true);
    teacherOptions = this.listAdminTeacherDirectory_();
    rows = this.listAbsenceRows_();
    return {
      ok: true,
      user: {
        email: actor.email,
        canManageReplacement: Boolean(actor.canManageReplacement)
      },
      operationalDate: dateState.selectedDate,
      dateState: dateState,
      teacherOptions: teacherOptions,
      dayRows: rows.filter(function (row) {
        return row.absenceMode === ROOMS_APP.Replacements.ABSENCE_MODES_.DAY;
      }),
      plannedRows: rows.filter(function (row) {
        return row.absenceMode === ROOMS_APP.Replacements.ABSENCE_MODES_.PLANNED;
      })
    };
  },

  getAbsenceTeacherPeriods: function (teacherEmail, dateString) {
    ROOMS_APP.Auth.requireCanManageReplacement();
    this.ensureSchema_();
    var dateState = this.resolveSelectedDate_(dateString, false);
    var teacherKey = this.normalizeTeacherEmail_(teacherEmail);
    if (!dateState.isValid) {
      return {
        ok: true,
        teacherEmail: teacherKey,
        date: dateState.selectedDate,
        dateState: dateState,
        periods: []
      };
    }
    return {
      ok: true,
      teacherEmail: teacherKey,
      date: dateState.selectedDate,
      dateState: dateState,
      periods: this.getTeacherServicePeriodsForDate_(teacherKey, dateState.selectedDate)
    };
  },

  saveAbsenceRegistry: function (rows, referenceDate) {
    var actor = ROOMS_APP.Auth.requireCanManageReplacement();
    var nowIso = ROOMS_APP.toIsoDateTime(new Date());
    var dateState;
    var persistedRows;
    var persistedById = {};
    var candidateRows = Array.isArray(rows) ? rows : [];
    var normalizedCandidates = [];
    var normalizedRows = [];
    var affectedDates;
    var seen = {};
    this.ensureSchema_();
    dateState = this.resolveSelectedDate_(referenceDate || ROOMS_APP.toIsoDate(ROOMS_APP.Auth.getEffectiveNow()), true);
    persistedRows = this.listAbsenceRows_();
    persistedRows.forEach(function (row) {
      persistedById[row.absenceId] = row;
    });

    candidateRows.forEach(function (entry) {
      var candidate = ROOMS_APP.Replacements.normalizeAbsencePayload_(entry || {});
      var existing = persistedById[candidate.absenceId] || null;
      if (candidate.absenceId.indexOf('draft-') === 0) {
        candidate.absenceId = ROOMS_APP.Replacements.buildAbsenceId_();
      }
      if (candidate.absenceMode === ROOMS_APP.Replacements.ABSENCE_MODES_.DAY) {
        candidate.startDate = dateState.selectedDate;
        candidate.endDate = '';
        candidate.isMultiDay = false;
      }
      if (candidate.absenceType !== ROOMS_APP.Replacements.ABSENCE_TYPES_.DAILY &&
          !ROOMS_APP.Replacements.isDailyServiceOutOfClassAbsence_(candidate)) {
        candidate.endDate = '';
        candidate.isMultiDay = false;
      }
      ROOMS_APP.Replacements.validateAbsenceCandidate_(candidate);
      if (seen[candidate.absenceId]) {
        return;
      }
      seen[candidate.absenceId] = true;
      normalizedCandidates.push(candidate);
      normalizedRows.push(ROOMS_APP.Replacements.buildAbsenceSheetRow_(candidate, actor, nowIso, existing));
    });

    ROOMS_APP.DB.replaceRows(
      ROOMS_APP.SHEET_NAMES.REPL_ABSENCES,
      ROOMS_APP.DB.getHeaders(ROOMS_APP.SHEET_NAMES.REPL_ABSENCES),
      normalizedRows
    );
    affectedDates = this.collectAbsenceCacheDates_(persistedRows.concat(normalizedCandidates), dateState.selectedDate);
    this.syncOperationalHourlyRowsForSourceCandidates_(normalizedCandidates, affectedDates, actor, nowIso, true);
    this.prepareDailyContextCacheAfterSourceChange_(
      dateState.selectedDate,
      affectedDates,
      'ABSENCE_REGISTRY_SAVE'
    );
    return this.getAbsenceRegistryModel(dateState.selectedDate);
  },

  saveAbsence: function (payload) {
    var actor = ROOMS_APP.Auth.requireCanManageReplacement();
    var nowIso = ROOMS_APP.toIsoDateTime(new Date());
    var candidate;
    var rows;
    var replaced = false;
    var previousRow = null;
    var nextRows;
    this.ensureSchema_();
    candidate = this.normalizeAbsencePayload_(payload || {});
    this.validateAbsenceCandidate_(candidate);
    rows = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.REPL_ABSENCES);
    nextRows = rows.map(function (row) {
      if (ROOMS_APP.normalizeString(row.AbsenceId) !== candidate.absenceId) {
        return row;
      }
      replaced = true;
      previousRow = ROOMS_APP.Replacements.readAbsenceRow_(row);
      return ROOMS_APP.Replacements.buildAbsenceSheetRow_(candidate, actor, nowIso, previousRow || row);
    });
    if (!replaced) {
      nextRows.push(this.buildAbsenceSheetRow_(candidate, actor, nowIso, null));
    }
    ROOMS_APP.DB.replaceRows(
      ROOMS_APP.SHEET_NAMES.REPL_ABSENCES,
      ROOMS_APP.DB.getHeaders(ROOMS_APP.SHEET_NAMES.REPL_ABSENCES),
      nextRows
    );
    this.syncOperationalHourlyRowsForSourceCandidates_([candidate], this.collectAbsenceCacheDates_(previousRow ? [previousRow, candidate] : [candidate], candidate.startDate), actor, nowIso, false);
    this.prepareDailyContextCacheAfterSourceChange_(
      candidate.startDate,
      this.collectAbsenceCacheDates_(previousRow ? [previousRow, candidate] : [candidate], candidate.startDate),
      'ABSENCE_SAVE'
    );
    return this.getAbsenceRegistryModel(candidate.startDate);
  },

  buildAbsenceSheetRow_: function (candidate, actor, nowIso, existingRow) {
    var existing = existingRow || {};
    var derived = this.buildAbsenceEffectiveWorkload_(candidate, nowIso);
    return {
      AbsenceId: candidate.absenceId,
      TeacherEmail: candidate.teacherEmail,
      TeacherName: candidate.teacherName,
      AbsenceMode: candidate.absenceMode,
      AbsenceType: candidate.absenceType,
      ServiceOutOfClassScope: candidate.serviceOutOfClassScope,
      StartDate: candidate.startDate,
      EndDate: candidate.endDate,
      HourlyPeriodsJson: JSON.stringify(candidate.hourlyPeriods),
      MatchedPeriodsJson: JSON.stringify(derived.matchedPeriods),
      EffectiveWorkloadJson: JSON.stringify(derived.effectiveWorkload),
      HasActionableWorkload: derived.hasActionableWorkload ? 'TRUE' : 'FALSE',
      ActionableHoursCount: String(derived.actionableHoursCount || 0),
      DerivedAtISO: derived.derivedAtISO,
      DerivedVersion: derived.derivedVersion,
      RecoveryRequired: candidate.recoveryRequired ? 'TRUE' : 'FALSE',
      Notes: candidate.notes,
      Status: 'ACTIVE',
      Enabled: 'TRUE',
      CreatedAtISO: existing.createdAtISO || existing.CreatedAtISO || nowIso,
      CreatedBy: existing.createdBy || existing.CreatedBy || actor.email,
      UpdatedAtISO: nowIso,
      UpdatedBy: actor.email
    };
  },

  isHourlyAbsenceCandidate_: function (candidate) {
    return Boolean(candidate && (
      candidate.absenceType === this.ABSENCE_TYPES_.HOURLY_PERMISSION ||
      this.isHourlyServiceOutOfClassAbsence_(candidate)
    ));
  },

  getTeacherDayMapForWorkloadDerivation_: function (dateString) {
    var targetDate = ROOMS_APP.toIsoDate(dateString);
    var bucket = this.getReplacementRequestCacheBucket_();
    var cacheKey = 'absenceWorkloadTeachers:' + targetDate;
    if (bucket && bucket[cacheKey]) {
      return bucket[cacheKey];
    }
    var teacherMap = this.indexByTeacherEmail_(this.buildTeacherDayTeachers_(targetDate));
    if (bucket) {
      bucket[cacheKey] = teacherMap;
    }
    return teacherMap;
  },

  buildAbsenceEffectiveWorkload_: function (candidate, nowIso) {
    var selectedMap = {};
    var selectedPeriods = [];
    var matchedPeriods = [];
    var items = [];
    var targetDate = ROOMS_APP.toIsoDate(candidate && candidate.startDate);
    var teacherEmail = this.normalizeTeacherEmail_(candidate && candidate.teacherEmail);
    var teacherName = ROOMS_APP.normalizeString(candidate && candidate.teacherName);
    var effectiveTeacherEmail = teacherEmail;
    var effectiveTeacherName = teacherName;
    var teacherMap;
    var teacher;
    var identity;
    var derivedAtIso = nowIso || ROOMS_APP.toIsoDateTime(new Date());

    (candidate && candidate.hourlyPeriods || []).forEach(function (period) {
      var normalizedPeriod = ROOMS_APP.normalizeString(period);
      if (normalizedPeriod) {
        selectedMap[normalizedPeriod] = true;
      }
    });
    selectedPeriods = Object.keys(selectedMap).sort(function (left, right) {
      return Number(left) - Number(right);
    });

    if (this.isHourlyAbsenceCandidate_(candidate) && targetDate && teacherEmail) {
      teacherMap = this.getTeacherDayMapForWorkloadDerivation_(targetDate);
      teacher = teacherMap[teacherEmail] || null;
      if (!teacher && teacherName) {
        teacher = teacherMap[this.buildTeacherSyntheticEmail_(teacherName)] || null;
      }
      if (!teacher) {
        identity = this.resolveAbsenceTeacherIdentityForDate_(
          teacherEmail,
          teacherName,
          this.getActiveLongAssignmentsForDate_(targetDate, true)
        );
        teacher = identity && identity.teacherEmail ? teacherMap[identity.teacherEmail] : null;
      }
      if (teacher) {
        effectiveTeacherEmail = teacher.teacherEmail || effectiveTeacherEmail;
        effectiveTeacherName = teacher.teacherName || effectiveTeacherName;
      }
      selectedPeriods.forEach(function (period) {
        var slot = teacher && teacher.periods ? teacher.periods[period] : null;
        var classCode = ROOMS_APP.Replacements.getTeacherSlotAssignmentClassCode_(slot);
        var label;
        if (!teacher || !slot || !ROOMS_APP.Replacements.isCoverableTeacherSlot_(teacher, slot) || !classCode) {
          return;
        }
        label = classCode === ROOMS_APP.Replacements.SERVICE_OUT_OF_CLASS_.ALTERNATIVA
          ? 'Alternativa'
          : (slot.label || classCode);
        matchedPeriods.push(period);
        items.push({
          period: period,
          classCode: classCode,
          label: label,
          slotType: slot.type || '',
          rawValue: slot.rawValue || '',
          serviceType: slot.serviceType || '',
          serviceSubtype: slot.serviceSubtype || '',
          startTime: slot.startTime || '',
          endTime: slot.endTime || '',
          actionable: true
        });
      });
    }

    return {
      matchedPeriods: matchedPeriods,
      hasActionableWorkload: Boolean(items.length),
      actionableHoursCount: items.length,
      derivedAtISO: derivedAtIso,
      derivedVersion: this.ABSENCE_WORKLOAD_DERIVED_VERSION_,
      effectiveWorkload: {
        version: this.ABSENCE_WORKLOAD_DERIVED_VERSION_,
        date: targetDate,
        teacherEmail: effectiveTeacherEmail,
        teacherName: effectiveTeacherName,
        absenceType: candidate && candidate.absenceType || '',
        serviceOutOfClassScope: candidate && candidate.serviceOutOfClassScope || '',
        selectedPeriods: selectedPeriods,
        matchedPeriods: matchedPeriods,
        items: items
      }
    };
  },

  syncOperationalHourlyRowsForSourceCandidates_: function (candidates, affectedDates, actor, nowIso, pruneMissingForDates) {
    var sheetName = ROOMS_APP.SHEET_NAMES.REPL_HOURLY_ABSENCES;
    var headers = ROOMS_APP.DB.getHeaders(sheetName);
    var dateMap = {};
    var activeSourceIds = {};
    var replaceSourceIds = {};
    var existingBySourceKey = {};
    var nextRows;
    var self = this;

    (affectedDates || []).forEach(function (dateValue) {
      var dateKey = ROOMS_APP.toIsoDate(dateValue);
      if (dateKey) {
        dateMap[dateKey] = true;
      }
    });
    (candidates || []).forEach(function (candidate) {
      if (candidate && candidate.absenceId) {
        activeSourceIds[candidate.absenceId] = true;
        replaceSourceIds[candidate.absenceId] = true;
      }
      if (candidate && candidate.startDate) {
        dateMap[ROOMS_APP.toIsoDate(candidate.startDate)] = true;
      }
    });

    nextRows = ROOMS_APP.DB.readRows(sheetName).map(function (row) {
      return ROOMS_APP.Replacements.readHourlyAbsenceRow_(row);
    }).filter(function (row) {
      var sourceId = ROOMS_APP.normalizeString(row.sourceAbsenceId);
      var inAffectedDate = Boolean(dateMap[row.date]);
      if (sourceId) {
        existingBySourceKey[[sourceId, row.period].join('|')] = row;
      }
      if (sourceId && replaceSourceIds[sourceId]) {
        return false;
      }
      if (pruneMissingForDates && inAffectedDate && row.sourceType === 'REPL_ABSENCES' && sourceId && !activeSourceIds[sourceId]) {
        return false;
      }
      return true;
    });

    (candidates || []).forEach(function (candidate) {
      self.buildOperationalHourlyRowsForAbsenceCandidate_(candidate, actor, nowIso, existingBySourceKey).forEach(function (row) {
        nextRows.push(row);
      });
    });

    ROOMS_APP.DB.replaceRows(sheetName, headers, nextRows.map(function (row) {
      return ROOMS_APP.Replacements.serializeHourlyAbsenceRow_(row, row.updatedAtISO || nowIso, row.updatedBy || actor && actor.email);
    }));
  },

  buildOperationalHourlyRowsForAbsenceCandidate_: function (candidate, actor, nowIso, existingBySourceKey) {
    var derived;
    var isServiceOut;
    if (!this.isHourlyAbsenceCandidate_(candidate)) {
      return [];
    }
    derived = this.buildAbsenceEffectiveWorkload_(candidate, nowIso);
    isServiceOut = this.isHourlyServiceOutOfClassAbsence_(candidate);
    return ((derived.effectiveWorkload && derived.effectiveWorkload.items) || []).filter(function (item) {
      return Boolean(item && item.actionable && item.period && item.classCode);
    }).map(function (item) {
      var existing = existingBySourceKey && existingBySourceKey[[candidate.absenceId, item.period].join('|')] || {};
      var recoveryRequired = !isServiceOut && Boolean(candidate.recoveryRequired);
      var recovered = recoveryRequired && existing.recoveryStatus === ROOMS_APP.Replacements.RECOVERY_STATUSES_.RECOVERED;
      return {
        date: candidate.startDate,
        teacherEmail: derived.effectiveWorkload.teacherEmail || candidate.teacherEmail,
        teacherName: derived.effectiveWorkload.teacherName || candidate.teacherName,
        period: item.period,
        reason: isServiceOut ? ROOMS_APP.Replacements.SERVICE_OUT_OF_CLASS_.GENERAL : '',
        classCode: item.classCode,
        label: item.label,
        slotType: item.slotType,
        startTime: item.startTime,
        endTime: item.endTime,
        derivedActionable: true,
        sourceAbsenceId: candidate.absenceId,
        sourceType: 'REPL_ABSENCES',
        recoveryRequired: recoveryRequired,
        recoveryStatus: recovered ? ROOMS_APP.Replacements.RECOVERY_STATUSES_.RECOVERED : (recoveryRequired ? ROOMS_APP.Replacements.RECOVERY_STATUSES_.PENDING : ''),
        recoveredOnDate: recovered ? existing.recoveredOnDate : '',
        recoveredByAssignmentKey: recovered ? existing.recoveredByAssignmentKey : '',
        notes: candidate.notes,
        updatedAtISO: nowIso || '',
        updatedBy: actor && actor.email || ''
      };
    });
  },

  removeOperationalHourlyRowsForSourceAbsences_: function (sourceRows) {
    var sheetName = ROOMS_APP.SHEET_NAMES.REPL_HOURLY_ABSENCES;
    var headers = ROOMS_APP.DB.getHeaders(sheetName);
    var sourceIds = {};
    var fallbackKeys = {};
    (sourceRows || []).forEach(function (row) {
      var sourceId = ROOMS_APP.normalizeString(row && row.absenceId);
      var teacherEmail = ROOMS_APP.Replacements.normalizeTeacherEmail_(row && row.teacherEmail);
      var dateKey = ROOMS_APP.toIsoDate(row && row.startDate);
      if (sourceId) {
        sourceIds[sourceId] = true;
      }
      (row && row.hourlyPeriods || []).forEach(function (period) {
        fallbackKeys[ROOMS_APP.Replacements.buildHourlyAbsenceDateKey_(dateKey, teacherEmail, period)] = true;
      });
    });
    ROOMS_APP.DB.replaceRows(sheetName, headers, ROOMS_APP.DB.readRows(sheetName).map(function (row) {
      return ROOMS_APP.Replacements.readHourlyAbsenceRow_(row);
    }).filter(function (row) {
      if (row.sourceAbsenceId && sourceIds[row.sourceAbsenceId]) {
        return false;
      }
      if (!row.sourceAbsenceId && fallbackKeys[ROOMS_APP.Replacements.buildHourlyAbsenceDateKey_(row.date, row.teacherEmail, row.period)]) {
        return false;
      }
      return true;
    }).map(function (row) {
      return ROOMS_APP.Replacements.serializeHourlyAbsenceRow_(row, row.updatedAtISO, row.updatedBy);
    }));
  },

  serializeHourlyAbsenceRow_: function (row, nowIso, updatedBy) {
    return {
      Date: row.date,
      TeacherEmail: row.teacherEmail,
      TeacherName: row.teacherName,
      Period: row.period,
      Reason: ROOMS_APP.normalizeString(row.reason),
      ClassCode: ROOMS_APP.normalizeString(row.classCode).toUpperCase(),
      Label: ROOMS_APP.normalizeString(row.label),
      SlotType: ROOMS_APP.normalizeString(row.slotType),
      StartTime: ROOMS_APP.normalizeString(row.startTime),
      EndTime: ROOMS_APP.normalizeString(row.endTime),
      DerivedActionable: row.derivedActionable ? 'TRUE' : '',
      SourceAbsenceId: ROOMS_APP.normalizeString(row.sourceAbsenceId),
      SourceType: ROOMS_APP.normalizeString(row.sourceType),
      RecoveryRequired: row.recoveryRequired ? 'TRUE' : 'FALSE',
      RecoveryStatus: row.recoveryStatus,
      RecoveredOnDate: row.recoveredOnDate,
      RecoveredByAssignmentKey: row.recoveredByAssignmentKey,
      Notes: row.notes,
      UpdatedAtISO: nowIso || row.updatedAtISO || '',
      UpdatedBy: updatedBy || row.updatedBy || ''
    };
  },

  deleteAbsence: function (absenceId) {
    ROOMS_APP.Auth.requireCanManageReplacement();
    this.ensureSchema_();
    var normalizedAbsenceId = ROOMS_APP.normalizeString(absenceId);
    var removed = false;
    var removedRows = [];
    var nextRows = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.REPL_ABSENCES).filter(function (row) {
      var keep = ROOMS_APP.normalizeString(row.AbsenceId) !== normalizedAbsenceId;
      if (!keep) {
        removed = true;
        removedRows.push(ROOMS_APP.Replacements.readAbsenceRow_(row));
      }
      return keep;
    });
    if (!removed) {
      throw new Error('Assenza non trovata.');
    }
    ROOMS_APP.DB.replaceRows(
      ROOMS_APP.SHEET_NAMES.REPL_ABSENCES,
      ROOMS_APP.DB.getHeaders(ROOMS_APP.SHEET_NAMES.REPL_ABSENCES),
      nextRows
    );
    this.removeOperationalHourlyRowsForSourceAbsences_(removedRows);
    this.prepareDailyContextCacheAfterSourceChange_(
      removedRows[0] && removedRows[0].startDate,
      this.collectAbsenceCacheDates_(removedRows, removedRows[0] && removedRows[0].startDate),
      'ABSENCE_DELETE'
    );
    return this.getAbsenceRegistryModel('');
  },

  deleteAbsencesForTeacherDate: function (teacherEmail, dateString) {
    ROOMS_APP.Auth.requireCanManageReplacement();
    this.ensureSchema_();
    var targetDate = ROOMS_APP.toIsoDate(dateString || '');
    var targetEmail = this.normalizeTeacherEmail_(teacherEmail);
    var activeLongAssignmentMap;
    var removedRows = [];
    var nextRows;
    if (!targetDate || !targetEmail) {
      throw new Error('Docente o data non validi.');
    }
    activeLongAssignmentMap = this.getActiveLongAssignmentsForDate_(targetDate, true);
    nextRows = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.REPL_ABSENCES).filter(function (row) {
      var entry = ROOMS_APP.Replacements.readAbsenceRow_(row);
      var identity = ROOMS_APP.Replacements.resolveAbsenceTeacherIdentityForDate_(
        entry.teacherEmail,
        entry.teacherName,
        activeLongAssignmentMap
      );
      var endDate;
      var matchesDate;
      if (!entry || !entry.absenceId || !entry.enabled || entry.status === 'DELETED') {
        return true;
      }
      if (entry.absenceType === ROOMS_APP.Replacements.ABSENCE_TYPES_.HOURLY_PERMISSION ||
          ROOMS_APP.Replacements.isHourlyServiceOutOfClassAbsence_(entry)) {
        matchesDate = entry.startDate === targetDate;
      } else {
        endDate = entry.endDate || entry.startDate;
        matchesDate = Boolean(entry.startDate <= targetDate && endDate >= targetDate);
      }
      if (matchesDate && ROOMS_APP.Replacements.normalizeTeacherEmail_(identity.teacherEmail) === targetEmail) {
        removedRows.push(entry);
        return false;
      }
      return true;
    });
    if (!removedRows.length) {
      throw new Error('Nessuna assenza/evento sorgente trovata per il docente selezionato.');
    }
    ROOMS_APP.DB.replaceRows(
      ROOMS_APP.SHEET_NAMES.REPL_ABSENCES,
      ROOMS_APP.DB.getHeaders(ROOMS_APP.SHEET_NAMES.REPL_ABSENCES),
      nextRows
    );
    this.removeOperationalHourlyRowsForSourceAbsences_(removedRows);
    this.prepareDailyContextCacheAfterSourceChange_(
      targetDate,
      this.collectAbsenceCacheDates_(removedRows, targetDate),
      'ABSENCE_DELETE_FROM_REPLACEMENTS'
    );
    return {
      ok: true,
      removedCount: removedRows.length,
      model: this.getModalModel(targetDate, null, {
        includeEditorData: false,
        useDailyCache: false
      })
    };
  },

  previewReport: function (dateString, draft) {
    ROOMS_APP.Auth.requireCanManageReplacement();
    this.ensureSchema_();
    var dateState = this.resolveSelectedDate_(dateString, false);
    if (!dateState.isValid) {
      throw new Error(dateState.message || 'Data non valida.');
    }

    var context = this.buildDayContext_(dateState.selectedDate, {
      includeEditorData: false
    });
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
    var context = this.buildDayContext_(targetDate, {
      includeEditorData: false
    });
    var normalized = this.normalizeDraft_(targetDate, draft, context);
    var validationErrors = this.validateDraft_(normalized);
    var savedPreviewPayload;
    if (validationErrors.length) {
      throw new Error(validationErrors.join(' '));
    }
    var nowIso = ROOMS_APP.toIsoDateTime(new Date());
    var updatedBy = actor.email;
    savedPreviewPayload = this.buildReportPayload_(normalized, context.recipients, validationErrors);

    var teacherRows = normalized.teachers.filter(function (entry) {
      return entry.absent || ROOMS_APP.normalizeString(entry.notes);
    }).map(function (entry) {
      return {
        Date: targetDate,
        TeacherEmail: entry.teacherEmail,
        TeacherName: entry.teacherName,
        Absent: entry.absent ? 'TRUE' : 'FALSE',
        Accompanist: 'FALSE',
        AccompaniedClasses: '',
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
        Reason: ROOMS_APP.normalizeString(entry.reason),
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

    this.replaceDateRows_(ROOMS_APP.SHEET_NAMES.REPL_CLASS_OUT, targetDate, []);
    this.replaceDateRows_(ROOMS_APP.SHEET_NAMES.REPL_DAY_TEACHERS, targetDate, teacherRows);
    this.saveHourlyAbsenceRows_(targetDate, normalized, nowIso, updatedBy);
    this.replaceDateRows_(ROOMS_APP.SHEET_NAMES.REPL_ASSIGNMENTS, targetDate, assignmentRows);
    this.prepareDailyContextCacheAfterSourceChange_(targetDate, [targetDate], 'REPLACEMENT_DAY_SAVE');

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
    var context = this.buildDayContext_(targetDate, {
      includeEditorData: false
    });
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
      this.archivePublishedReportSnapshot_(this.REPORT_TYPE_, targetDate, payload, {
        actorEmail: actor.email,
        status: 'PUBLISHED',
        visibleToTeachers: true,
        notes: senderInfo
          ? ('mode=' + senderInfo.senderMode +
            '; name=' + senderInfo.fromName +
            (senderInfo.replyTo ? '; replyTo=' + senderInfo.replyTo : '') +
            (senderInfo.noReply ? '; noReply=true' : ''))
          : ''
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
    var previousRow = null;

    this.validateLongAssignment_(normalized, rows, normalized.matchKey);

    var nextRows = rows.map(function (row) {
      if (normalized.matchKey && ROOMS_APP.Replacements.buildLongAssignmentMatchKey_(row) === normalized.matchKey) {
        replaced = true;
        previousRow = row;
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
    this.prepareDailyContextCacheAfterSourceChange_(
      referenceDate || normalized.startDate,
      this.collectDateRangeCacheDates_(normalized.startDate, normalized.endDate, referenceDate || normalized.startDate).concat(
        previousRow
          ? this.collectDateRangeCacheDates_(previousRow.StartDate, previousRow.EndDate, referenceDate || normalized.startDate)
          : []
      ),
      'LONG_ASSIGNMENT_SAVE'
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
    this.prepareDailyContextCacheAfterSourceChange_(
      referenceDate || ROOMS_APP.toIsoDate(new Date()),
      this.collectDateRangeCacheDates_(targetRow.StartDate, targetRow.EndDate, referenceDate || ROOMS_APP.toIsoDate(new Date())),
      'LONG_ASSIGNMENT_TOGGLE'
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
    var removedRow = null;
    var nextRows = this.listLongAssignmentRows_().filter(function (row) {
      var keep = ROOMS_APP.Replacements.buildLongAssignmentMatchKey_(row) !== ROOMS_APP.normalizeString(matchKey);
      if (!keep) {
        removed = true;
        removedRow = row;
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
    this.prepareDailyContextCacheAfterSourceChange_(
      referenceDate || ROOMS_APP.toIsoDate(new Date()),
      this.collectDateRangeCacheDates_(
        removedRow && removedRow.StartDate,
        removedRow && removedRow.EndDate,
        referenceDate || ROOMS_APP.toIsoDate(new Date())
      ),
      'LONG_ASSIGNMENT_DELETE'
    );

    return {
      ok: true,
      model: this.getModalModel(referenceDate || ROOMS_APP.toIsoDate(new Date()), null, {
        allowAutoShift: false
      })
    };
  },

  saveEducationalTrip: function (payload, referenceDate) {
    var normalized = this.normalizeEducationalTripPayload_(payload || {});
    var replaced = false;
    var trips = this.listEducationalTrips_().map(function (trip) {
      if (ROOMS_APP.normalizeString(trip.tripId) !== ROOMS_APP.normalizeString(normalized.tripId)) {
        return trip;
      }
      replaced = true;
      return Object.assign({}, normalized, {
        tripId: normalized.tripId || trip.tripId
      });
    });
    if (!replaced) {
      trips.push(normalized);
    }
    return this.saveReplacementFieldTripRegistry(trips, referenceDate || normalized.startDate);
  },

  toggleEducationalTrip: function (tripId, enabled, referenceDate) {
    var found = false;
    var trips = this.listEducationalTrips_().map(function (trip) {
      if (ROOMS_APP.normalizeString(trip.tripId) !== ROOMS_APP.normalizeString(tripId)) {
        return trip;
      }
      found = true;
      return Object.assign({}, trip, {
        enabled: Boolean(enabled)
      });
    });
    if (!found) {
      throw new Error('Uscita didattica non trovata.');
    }
    return this.saveReplacementFieldTripRegistry(trips, referenceDate);
  },

  deleteEducationalTrip: function (tripId, referenceDate) {
    var normalizedTripId = ROOMS_APP.normalizeString(tripId);
    var removed = false;
    var trips = this.listEducationalTrips_().filter(function (trip) {
      var keep = ROOMS_APP.normalizeString(trip.tripId) !== normalizedTripId;
      if (!keep) {
        removed = true;
      }
      return keep;
    });
    if (!removed) {
      throw new Error('Uscita didattica non trovata.');
    }
    return this.saveReplacementFieldTripRegistry(trips, referenceDate);
  },

  saveReplacementFieldTripRegistry: function (entries, referenceDate) {
    var actor = ROOMS_APP.Auth.requireCanManageReplacement();
    var self = this;
    var nowIso = ROOMS_APP.toIsoDateTime(new Date());
    var normalizedTrips = [];
    var seenTripIds = {};
    var persistenceRows;
    var previousTrips;
    this.ensureSchema_();
    previousTrips = this.listEducationalTrips_();

    normalizedTrips = (entries || []).map(function (entry) {
      var normalized = self.normalizeEducationalTripPayload_(entry || {});
      var tripId = normalized.tripId || self.buildTripId_();
      if (seenTripIds[tripId]) {
        throw new Error('Sono presenti uscite didattiche duplicate nella bozza locale.');
      }
      seenTripIds[tripId] = true;
      normalized.tripId = tripId;
      return normalized;
    });

    this.validateEducationalTripRegistry_(normalizedTrips);
    persistenceRows = this.buildEducationalTripRegistryRows_(normalizedTrips, actor.email, nowIso);

    ROOMS_APP.DB.replaceRows(
      ROOMS_APP.SHEET_NAMES.REPL_FIELD_TRIPS,
      ROOMS_APP.DB.getHeaders(ROOMS_APP.SHEET_NAMES.REPL_FIELD_TRIPS),
      persistenceRows.tripRows
    );
    ROOMS_APP.DB.replaceRows(
      ROOMS_APP.SHEET_NAMES.REPL_FIELD_TRIP_TEACHERS,
      ROOMS_APP.DB.getHeaders(ROOMS_APP.SHEET_NAMES.REPL_FIELD_TRIP_TEACHERS),
      persistenceRows.tripTeacherRows
    );
    this.prepareDailyContextCacheAfterSourceChange_(
      referenceDate || (normalizedTrips[0] && normalizedTrips[0].startDate) || ROOMS_APP.toIsoDate(new Date()),
      this.collectTripCacheDates_(previousTrips.concat(normalizedTrips), referenceDate),
      'FIELD_TRIP_REGISTRY_SAVE'
    );
    return {
      ok: true,
      model: this.getModalModel(referenceDate || (normalizedTrips[0] && normalizedTrips[0].startDate) || ROOMS_APP.toIsoDate(new Date()), null, {
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

    if (this.isClassOutAtPeriodInState_(normalized, classCode, period)) {
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

  buildDayContext_: function (dateString, options) {
    var targetDate = ROOMS_APP.toIsoDate(dateString || new Date());
    var settings = options && typeof options === 'object' ? options : {};
    var includeEditorData = settings.includeEditorData !== false;
    var cachedContext = !includeEditorData && settings.useDailyCache !== false
      ? this.readDailyContextCache_(targetDate)
      : null;
    if (cachedContext) {
      return cachedContext;
    }
    var activeLongAssignments = this.getActiveLongAssignmentsForDate_(targetDate);
    var activeLongAssignmentMap = this.getActiveLongAssignmentsForDate_(targetDate, true);
    var baseTeachers = this.buildTeacherDayTeachers_(targetDate);
    var allTrips = this.listEducationalTrips_();
    var tripDayState = this.buildTripDayState_(targetDate, allTrips);
    var tripTeacherRegistry = includeEditorData
      ? this.getTripTeacherOptionsRegistry_()
      : { classOptions: [], byClass: {} };
    var activeAbsenceRows = this.listActiveAbsenceRowsForDate_(targetDate);
    var absenceRegistryState = this.buildAbsenceRegistryDayState_(targetDate, activeAbsenceRows, activeLongAssignmentMap);
    var savedClassOutRows = this.listRowsForDate_(ROOMS_APP.SHEET_NAMES.REPL_CLASS_OUT, targetDate);
    var savedTeacherRows = this.mergeTeacherRowsWithRegistry_(
      this.listRowsForDate_(ROOMS_APP.SHEET_NAMES.REPL_DAY_TEACHERS, targetDate),
      absenceRegistryState,
      activeLongAssignmentMap
    );
    var savedHourlyAbsenceRows = this.mergeHourlyAbsenceRowsWithRegistry_(
      this.listRowsForDate_(ROOMS_APP.SHEET_NAMES.REPL_HOURLY_ABSENCES, targetDate),
      absenceRegistryState,
      activeLongAssignmentMap
    );
    var savedAssignmentRows = this.listRowsForDate_(ROOMS_APP.SHEET_NAMES.REPL_ASSIGNMENTS, targetDate);
    var pendingRecoveryRows = this.listPendingRecoveryRows_(targetDate);
    var periodMap = ROOMS_APP.Timetable.getPeriodTimeMap();
    var legacyTripState = this.buildLegacyOutingState_(targetDate, savedClassOutRows, savedTeacherRows);
    var combinedTripState = this.mergeTripDayStates_(tripDayState, legacyTripState);
    var teacherMap = {};
    var teacherEmailByName = {};
    var classSet = {};
    var savedTeacherFallbackMap = {};

    baseTeachers.forEach(function (teacher) {
      teacherMap[teacher.teacherEmail] = teacher;
      if (teacher.teacherName) {
        teacherEmailByName[ROOMS_APP.slugify(teacher.teacherName)] = teacher.teacherEmail;
      }
      Object.keys(teacher.periods || {}).forEach(function (period) {
        var slot = teacher.periods[period];
        var assignmentClassCode = ROOMS_APP.Replacements.getTeacherSlotAssignmentClassCode_(slot);
        if (slot && ROOMS_APP.Replacements.isCoverableTeacherSlot_(teacher, slot) && assignmentClassCode) {
          classSet[assignmentClassCode] = true;
        }
      });
    });

    Object.keys(combinedTripState.classOutSet || {}).forEach(function (classCode) {
      classSet[classCode] = true;
    });

    savedTeacherRows.forEach(function (row) {
      var teacherEmail = ROOMS_APP.Replacements.normalizeTeacherEmail_(row.TeacherEmail);
      var rowTeacherName = ROOMS_APP.normalizeString(row.TeacherName);
      if (!teacherEmail) {
        teacherEmail = ROOMS_APP.Replacements.buildTeacherSyntheticEmail_(row.TeacherName);
      }
      if (activeLongAssignmentMap[teacherEmail]) {
        teacherEmail = activeLongAssignmentMap[teacherEmail].replacementTeacherEmail;
      }
      if (!teacherMap[teacherEmail] && rowTeacherName && teacherEmailByName[ROOMS_APP.slugify(rowTeacherName)]) {
        teacherEmail = teacherEmailByName[ROOMS_APP.slugify(rowTeacherName)];
      }
      if (!teacherMap[teacherEmail]) {
        teacherMap[teacherEmail] = {
          teacherEmail: teacherEmail,
          teacherName: activeLongAssignmentMap[ROOMS_APP.Replacements.normalizeTeacherEmail_(row.TeacherEmail) || ROOMS_APP.Replacements.buildTeacherSyntheticEmail_(row.TeacherName)]
            ? activeLongAssignmentMap[ROOMS_APP.Replacements.normalizeTeacherEmail_(row.TeacherEmail) || ROOMS_APP.Replacements.buildTeacherSyntheticEmail_(row.TeacherName)].replacementTeacherDisplayName
            : ROOMS_APP.normalizeString(row.TeacherName),
          periods: {},
          supportTeacher: false
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
          periods: {},
          supportTeacher: false
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
          label: classCode,
          coverageRequired: true
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
      var classCode = ROOMS_APP.normalizeString(row.ClassCode).toUpperCase();
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
          periods: {},
          supportTeacher: false
        };
      }
      slot = periodMap[period] || {};
      if (classCode) {
        classSet[classCode] = true;
      }
      if (period && !teacherMap[teacherEmail].periods[period]) {
        teacherMap[teacherEmail].periods[period] = {
          period: String(period || ''),
          startTime: slot.startTime || '',
          endTime: slot.endTime || '',
          rawValue: classCode || '',
          type: ROOMS_APP.asBoolean(row.DerivedActionable) && classCode
            ? (classCode === ROOMS_APP.Replacements.SERVICE_OUT_OF_CLASS_.ALTERNATIVA ? 'SERVICE' : 'CLASS')
            : 'FREE',
          classCode: ROOMS_APP.asBoolean(row.DerivedActionable) ? classCode : '',
          label: ROOMS_APP.asBoolean(row.DerivedActionable)
            ? (ROOMS_APP.normalizeString(row.Label) || (classCode === ROOMS_APP.Replacements.SERVICE_OUT_OF_CLASS_.ALTERNATIVA ? 'Alternativa' : classCode))
            : '',
          coverageRequired: Boolean(ROOMS_APP.asBoolean(row.DerivedActionable) && classCode)
        };
      }
    });

    var teachers = Object.keys(teacherMap).map(function (teacherEmail) {
      var teacher = teacherMap[teacherEmail];
      var saved = ROOMS_APP.Replacements.findSavedTeacherRow_(savedTeacherRows, teacherEmail, teacher.teacherName, activeLongAssignmentMap) || savedTeacherFallbackMap[teacherEmail] || {};
      var accompaniedClasses = ROOMS_APP.Replacements.uniqueStrings_(
        ROOMS_APP.Replacements.parsePipeList_(saved.AccompaniedClasses).concat(
          combinedTripState.teacherTripClassesByTeacher[teacherEmail] || []
        )
      );
      accompaniedClasses.forEach(function (classCode) {
        classSet[classCode] = true;
      });
      return {
        teacherEmail: teacher.teacherEmail,
        teacherName: teacher.teacherName,
        periods: teacher.periods || {},
        absent: ROOMS_APP.asBoolean(saved.Absent) && !ROOMS_APP.asBoolean(saved.Accompanist),
        accompanist: Boolean(accompaniedClasses.length),
        accompaniedClasses: accompaniedClasses,
        supportTeacher: Boolean(teacher.supportTeacher),
        serviceOutOfClass: Boolean(absenceRegistryState.serviceTeacherMap && absenceRegistryState.serviceTeacherMap[teacher.teacherEmail]),
        notes: ROOMS_APP.normalizeString(saved.Notes)
      };
    }).sort(function (left, right) {
      return left.teacherName.localeCompare(right.teacherName);
    });

    var classes = Object.keys(classSet).sort().map(function (classCode) {
      return {
        classCode: classCode,
        isOut: Boolean(combinedTripState.classOutSet[classCode]),
        notes: ''
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
          supportTeacher: Boolean(entry.supportTeacher),
          serviceOutOfClass: Boolean(entry.serviceOutOfClass),
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

    var context = {
      date: targetDate,
      classes: classes,
      teachers: teachers,
      teacherMap: this.indexByTeacherEmail_(teachers),
      tripDayState: combinedTripState,
      absenceRegistryState: absenceRegistryState,
      savedDraft: savedDraft,
      savedAtISO: this.computeSavedAtISO_(
        savedClassOutRows,
        savedTeacherRows,
        savedHourlyAbsenceRows,
        savedAssignmentRows,
        absenceRegistryState.sourceRows,
        combinedTripState.sourceRows,
        combinedTripState.sourceTeacherRows
      ),
      reportStatus: this.getLatestReportStatus_(targetDate),
      recipients: this.getRecipients_(),
      longAssignments: includeEditorData ? this.buildLongAssignmentsList_() : [],
      longTeacherOptions: includeEditorData ? this.listAdminTeacherDirectory_() : [],
      trips: includeEditorData ? allTrips : [],
      tripClassOptions: tripTeacherRegistry.classOptions,
      tripTeacherOptionsByClass: tripTeacherRegistry.byClass,
      activeLongAssignments: activeLongAssignments,
      activeLongAssignmentMap: activeLongAssignmentMap,
      pendingRecoveryRows: pendingRecoveryRows,
      editorDataLoaded: includeEditorData
    };
    if (!includeEditorData && settings.writeDailyCache !== false && settings.useDailyCache !== false) {
      this.writeDailyContextCache_(targetDate, context, 'REBUILT_ON_DEMAND');
    }
    return context;
  },

  normalizeDraft_: function (dateString, draft, context) {
    var targetDate = ROOMS_APP.toIsoDate(dateString || new Date());
    var source = draft || {};
    var baseContext = context || this.buildDayContext_(targetDate);
    var absenceRegistryState = baseContext.absenceRegistryState || {};
    var classBaseMap = {};
    var teacherBaseMap = {};
    var hourlySource = absenceRegistryState.hasEntries
      ? (((baseContext.savedDraft && baseContext.savedDraft.hourlyAbsences) || []).concat(
        (Array.isArray(source.hourlyAbsences) ? source.hourlyAbsences : []).filter(function (entry) {
          var normalizedEntry = ROOMS_APP.Replacements.normalizeHourlyAbsenceEntry_(entry, targetDate);
          var hourlyKey = ROOMS_APP.Replacements.buildHourlyAbsenceKey_(normalizedEntry.teacherEmail, normalizedEntry.period);
          if (absenceRegistryState.controlledTeacherMap && absenceRegistryState.controlledTeacherMap[normalizedEntry.teacherEmail]) {
            return false;
          }
          if (absenceRegistryState.controlledHourlyMap && absenceRegistryState.controlledHourlyMap[hourlyKey]) {
            return false;
          }
          return true;
        })
      ))
      : (Array.isArray(source.hourlyAbsences)
      ? source.hourlyAbsences
      : ((baseContext.savedDraft && baseContext.savedDraft.hourlyAbsences) || []));

    (baseContext.classes || []).forEach(function (entry) {
      classBaseMap[ROOMS_APP.normalizeString(entry.classCode).toUpperCase()] = entry;
    });
    (baseContext.teachers || []).forEach(function (entry) {
      teacherBaseMap[ROOMS_APP.Replacements.normalizeTeacherEmail_(entry.teacherEmail)] = entry;
    });

    var classOutMap = {};
    Object.keys(classBaseMap).forEach(function (classCode) {
      classOutMap[classCode] = {
        classCode: classCode,
        isOut: Boolean(classBaseMap[classCode].isOut),
        notes: ROOMS_APP.normalizeString(classBaseMap[classCode].notes)
      };
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
        supportTeacher: Boolean(base.supportTeacher),
        serviceOutOfClass: Boolean(base.serviceOutOfClass),
        notes: ROOMS_APP.normalizeString(base.notes)
      };
    });

    (Array.isArray(source.teachers) ? source.teachers : []).forEach(function (entry) {
      var teacherEmail = ROOMS_APP.Replacements.normalizeTeacherEmail_(entry && (entry.teacherEmail || entry.TeacherEmail));
      var teacherName = ROOMS_APP.normalizeString(entry && (entry.teacherName || entry.TeacherName));
      var isRegistryControlled;
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
          supportTeacher: false,
          serviceOutOfClass: false,
          notes: ''
        };
      }
      isRegistryControlled = Boolean(absenceRegistryState.controlledTeacherMap && absenceRegistryState.controlledTeacherMap[teacherEmail]);
      normalizedTeachers[teacherEmail].teacherName = teacherName || normalizedTeachers[teacherEmail].teacherName;
      normalizedTeachers[teacherEmail].supportTeacher = Boolean(normalizedTeachers[teacherEmail].supportTeacher ||
        ROOMS_APP.asBoolean(entry && (entry.supportTeacher || entry.SupportTeacher)));
      normalizedTeachers[teacherEmail].serviceOutOfClass = Boolean(normalizedTeachers[teacherEmail].serviceOutOfClass ||
        ROOMS_APP.asBoolean(entry && (entry.serviceOutOfClass || entry.ServiceOutOfClass)));
      if (!isRegistryControlled) {
        normalizedTeachers[teacherEmail].absent = ROOMS_APP.asBoolean(entry && Object.prototype.hasOwnProperty.call(entry, 'absent') ? entry.absent : entry && entry.Absent);
        normalizedTeachers[teacherEmail].notes = ROOMS_APP.normalizeString(entry && (entry.notes || entry.Notes));
      }
    });

    var validOutClasses = {};
    normalizedClasses.forEach(function (entry) {
      if (entry.isOut) {
        validOutClasses[entry.classCode] = true;
      }
    });

    var teacherList = Object.keys(normalizedTeachers).map(function (teacherEmail) {
      var teacher = normalizedTeachers[teacherEmail];
      teacher.accompaniedClasses = (teacher.accompaniedClasses || []).filter(function (classCode) {
        return Boolean(validOutClasses[ROOMS_APP.normalizeString(classCode).toUpperCase()]);
      }).map(function (classCode) {
        return ROOMS_APP.normalizeString(classCode).toUpperCase();
      });
      teacher.accompanist = Boolean(teacher.accompaniedClasses.length);
      return teacher;
    }).sort(function (left, right) {
      return left.teacherName.localeCompare(right.teacherName);
    });

    var teacherMap = this.indexByTeacherEmail_(teacherList);
    var hourlyAbsenceMap = this.buildValidHourlyAbsenceMap_(hourlySource || [], targetDate, teacherMap);
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
      draftAssignmentMap,
      baseContext.tripDayState || {}
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
      activeLongAssignments: baseContext.activeLongAssignments || [],
      tripDayState: baseContext.tripDayState || {}
    };
  },

  buildEffectiveAssignments_: function (dateString, classes, teachers, teacherMap, hourlyAbsences, draftAssignmentMap, tripDayState) {
    var classOutSet = {};
    var affectedMap = {};
    var effectiveTripState = tripDayState || {};
    classes.forEach(function (entry) {
      if (entry.isOut) {
        classOutSet[entry.classCode] = true;
      }
    });

    teachers.forEach(function (teacher) {
      if (!teacher.absent && !teacher.accompanist && !teacher.serviceOutOfClass) {
        return;
      }
      Object.keys(teacher.periods || {}).forEach(function (period) {
        var slot = teacher.periods[period];
        var assignmentClassCode = ROOMS_APP.Replacements.getTeacherSlotAssignmentClassCode_(slot);
        var key;
        if (!slot || !ROOMS_APP.Replacements.isCoverableTeacherSlot_(teacher, slot) || !assignmentClassCode) {
          return;
        }
        if (teacher.accompanist && !teacher.absent &&
          !ROOMS_APP.Replacements.isTeacherAccompanyingAtPeriodInState_(effectiveTripState, teacher.teacherEmail, period)) {
          return;
        }
        key = ROOMS_APP.Replacements.buildAssignmentKey_(period, assignmentClassCode, teacher.teacherEmail);
        affectedMap[key] = {
          teacher: teacher,
          period: period,
          slot: Object.assign({}, slot, { classCode: assignmentClassCode }),
          originalStatus: teacher.serviceOutOfClass && !teacher.absent ? 'SERVICE_OUT_OF_CLASS' : (teacher.absent ? 'ABSENT' : 'ACCOMPANIST'),
          hourlyAbsence: null
        };
      });
    });

    (hourlyAbsences || []).forEach(function (entry) {
      var teacher = teacherMap[ROOMS_APP.Replacements.normalizeTeacherEmail_(entry.teacherEmail)] || null;
      var slot = teacher && teacher.periods ? teacher.periods[entry.period] : null;
      var assignmentClassCode = ROOMS_APP.Replacements.getTeacherSlotAssignmentClassCode_(slot) || ROOMS_APP.normalizeString(entry.classCode).toUpperCase();
      var key;
      if (!teacher || teacher.absent || !assignmentClassCode) {
        return;
      }
      if (!slot || !ROOMS_APP.Replacements.isCoverableTeacherSlot_(teacher, slot)) {
        if (!entry.derivedActionable) {
          return;
        }
        slot = {
          period: String(entry.period || ''),
          startTime: entry.startTime || '',
          endTime: entry.endTime || '',
          rawValue: assignmentClassCode,
          type: assignmentClassCode === ROOMS_APP.Replacements.SERVICE_OUT_OF_CLASS_.ALTERNATIVA ? 'SERVICE' : 'CLASS',
          classCode: assignmentClassCode,
          label: entry.label || (assignmentClassCode === ROOMS_APP.Replacements.SERVICE_OUT_OF_CLASS_.ALTERNATIVA ? 'Alternativa' : assignmentClassCode),
          coverageRequired: true
        };
      }
      if (!ROOMS_APP.Replacements.isCoverableTeacherSlot_(teacher, slot)) {
        return;
      }
      key = ROOMS_APP.Replacements.buildAssignmentKey_(entry.period, assignmentClassCode, teacher.teacherEmail);
      if (!affectedMap[key]) {
        affectedMap[key] = {
          teacher: teacher,
          period: entry.period,
          slot: Object.assign({}, slot, { classCode: assignmentClassCode }),
          originalStatus: ROOMS_APP.Replacements.isServiceOutOfClassHourlyAbsenceEntry_(entry) ? 'SERVICE_OUT_OF_CLASS' : 'HOURLY_ABSENCE',
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
        ROOMS_APP.Replacements.isTeacherAccompanyingAtPeriodInState_(effectiveTripState, teacher.teacherEmail, meta.period) &&
        ROOMS_APP.Replacements.isClassOutAtPeriodInState_(effectiveTripState, slot.classCode, meta.period)
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
        endTime: slot.endTime,
        actualEntryPeriod: '',
        actualExitPeriod: ''
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

    var normalizedAssignments = assignments.sort(function (left, right) {
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
    var normalizedForFlow = {
      assignments: normalizedAssignments,
      teachers: teachers,
      tripDayState: effectiveTripState
    };
    normalizedAssignments.forEach(function (entry) {
      var classHandlingType = ROOMS_APP.Replacements.normalizeClassHandlingType_(entry.classHandlingType);
      var flow;
      if (!classHandlingType) {
        return;
      }
      flow = ROOMS_APP.Replacements.buildClassFlowForValidation_(normalizedForFlow, entry.classCode);
      if (classHandlingType === ROOMS_APP.Replacements.CLASS_HANDLING_TYPES_.LATE_ENTRY) {
        entry.actualEntryPeriod = ROOMS_APP.Replacements.findNextClassFlowPeriod_(flow, entry.period, function (flowEntry) {
          return flowEntry.state === 'VALID';
        });
        return;
      }
      if (classHandlingType === ROOMS_APP.Replacements.CLASS_HANDLING_TYPES_.EARLY_EXIT) {
        entry.actualExitPeriod = ROOMS_APP.Replacements.findPreviousClassFlowPeriod_(flow, entry.period, function (flowEntry) {
          return flowEntry.state === 'VALID';
        });
      }
    });

    return normalizedAssignments;
  },

  buildTeacherDetailRows_: function (normalized, teacher) {
    var self = this;
    var assignmentsByKey = normalized.assignmentMap;
    var assignmentsByTeacherPeriod = {};
    var assignedByPeriod = this.buildAssignedTeacherSetByPeriod_(normalized.assignments);
    var rows = [];
    var periodMap = ROOMS_APP.Timetable.getPeriodTimeMap();
    var periodKeys = Object.keys(periodMap).sort(function (left, right) {
      return Number(left) - Number(right);
    });

    (normalized.assignments || []).forEach(function (assignment) {
      var teacherEmail = self.normalizeTeacherEmail_(assignment && assignment.originalTeacherEmail);
      var period = ROOMS_APP.normalizeString(assignment && assignment.period);
      if (teacherEmail && period) {
        assignmentsByTeacherPeriod[self.buildHourlyAbsenceKey_(teacherEmail, period)] = assignment;
      }
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
      var fallbackAssignment = assignmentsByTeacherPeriod[self.buildHourlyAbsenceKey_(teacher.teacherEmail, period)] || null;
      var assignmentClassCode = ROOMS_APP.Replacements.getTeacherSlotAssignmentClassCode_(slot) ||
        ROOMS_APP.normalizeString(fallbackAssignment && fallbackAssignment.classCode).toUpperCase();
      var assignmentKey = ROOMS_APP.Replacements.buildAssignmentKey_(period, assignmentClassCode, teacher.teacherEmail);
      var assignment = assignmentsByKey[assignmentKey] || fallbackAssignment || null;
      var isAccompanying = ROOMS_APP.Replacements.isTeacherAccompanyingAtPeriod_(
        normalized,
        teacher.teacherEmail,
        period
      );
      if (assignmentClassCode === ROOMS_APP.Replacements.SERVICE_OUT_OF_CLASS_.ALTERNATIVA &&
          (slot.type === 'FREE' || slot.type === 'OTHER')) {
        slot = Object.assign({}, slot, {
          type: 'SERVICE',
          classCode: assignmentClassCode,
          label: 'Alternativa',
          coverageRequired: true
        });
      }
      var isCoverableSlot = ROOMS_APP.Replacements.isCoverableTeacherSlot_(teacher, slot) || Boolean(assignment);
      var requiresReplacement = Boolean((isCoverableSlot && (teacher.absent || hourlyAbsence || isAccompanying)) || assignment);
      var row = {
        period: period,
        startTime: slot.startTime,
        endTime: slot.endTime,
        slotType: slot.type,
        classCode: assignmentClassCode,
        label: slot.label,
        canMarkHourlyAbsence: Boolean(isCoverableSlot && !teacher.absent && !isAccompanying),
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
        if (isCoverableSlot) {
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
      if (ROOMS_APP.Replacements.isTeacherAccompanyingAtPeriod_(normalized, teacherEmail, period)) {
        return;
      }

      if (slot.type === 'CLASS' && ROOMS_APP.Replacements.isClassOutAtPeriod_(normalized, slot.classCode, period)) {
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
      var actualEntryPeriod = this.getActualClassTransitionPeriod_(entry, this.CLASS_HANDLING_TYPES_.LATE_ENTRY);
      if (actualEntryPeriod) {
        periodCells[actualEntryPeriod] = 'Entrata';
      }
      noteParts.push(entry.classCode + ' entrata posticipata');
      if (ROOMS_APP.normalizeString(entry.notes)) {
        noteParts.push(ROOMS_APP.normalizeString(entry.notes));
      }
      return {
        groupKey: 'CLASS|' + classHandlingType + '|' + entry.classCode + '|' + actualEntryPeriod + '|' + noteParts.join(' | '),
        designatedTeacher: '',
        noteLabel: noteParts.join(' | '),
        noteOnly: true,
        periodCells: periodCells
      };
    }
    if (classHandlingType === this.CLASS_HANDLING_TYPES_.EARLY_EXIT) {
      var actualExitPeriod = this.getActualClassTransitionPeriod_(entry, this.CLASS_HANDLING_TYPES_.EARLY_EXIT);
      if (actualExitPeriod) {
        periodCells[actualExitPeriod] = 'Uscita';
      }
      noteParts.push(entry.classCode + ' uscita anticipata');
      if (ROOMS_APP.normalizeString(entry.notes)) {
        noteParts.push(ROOMS_APP.normalizeString(entry.notes));
      }
      return {
        groupKey: 'CLASS|' + classHandlingType + '|' + entry.classCode + '|' + actualExitPeriod + '|' + noteParts.join(' | '),
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
      noteParts = this.buildPublicReportNoteParts_(entry, sourceLabel);
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

  buildPublicReportNoteParts_: function (entry, sourceLabel) {
    var handlingType = this.normalizeHandlingType_(entry && entry.handlingType);
    var noteParts = [];
    if (handlingType === this.HANDLING_TYPES_.RECOVERY && entry && entry.recoverySourceDate) {
      noteParts.push('Recupero del ' + this.formatShortDate_(entry.recoverySourceDate));
    } else if (sourceLabel) {
      noteParts.push(sourceLabel);
    }
    if (ROOMS_APP.normalizeString(entry && entry.notes)) {
      noteParts.push(ROOMS_APP.normalizeString(entry.notes));
    }
    return noteParts;
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
    var assignedByPeriod = this.buildAssignedTeacherSetByPeriod_(normalized.assignments);
    var recoveryLinkMap = {};
    var self = this;

    (normalized && normalized.hourlyAbsences || []).forEach(function (entry) {
      var teacher = normalized.teacherMap[self.normalizeTeacherEmail_(entry.teacherEmail)] || null;
      var slot = teacher && teacher.periods ? teacher.periods[entry.period] : null;
      if (!teacher || !slot || !self.isCoverableTeacherSlot_(teacher, slot) || !self.getTeacherSlotAssignmentClassCode_(slot)) {
        errors.push('Assenza oraria non valida per ' + (entry.teacherName || entry.teacherEmail) + ' alla ' + entry.period + 'ª ora.');
      }
    });

    (normalized && normalized.assignments || []).forEach(function (entry) {
      var teacher = normalized.teacherMap[self.normalizeTeacherEmail_(entry.originalTeacherEmail)] || null;
      var candidateInfo;
      var matchFound;
      if (!teacher) {
        return;
      }
      if (self.normalizeClassHandlingType_(entry.classHandlingType)) {
        var classHandlingError = self.validateClassHandlingAssignment_(normalized, entry);
        if (classHandlingError) {
          errors.push(classHandlingError);
        }
        return;
      }
      if (entry.replacementStatus === 'IN_USCITA' || entry.replacementStatus === 'TO_ASSIGN') {
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

  validateClassHandlingAssignment_: function (normalized, entry) {
    var classHandlingType = this.normalizeClassHandlingType_(entry && entry.classHandlingType);
    var classCode = ROOMS_APP.normalizeString(entry && entry.classCode).toUpperCase();
    var sourcePeriod = ROOMS_APP.normalizeString(entry && entry.period);
    var flow;
    var transitionPeriod;
    var invalidState;
    if (!classHandlingType || !classCode || !sourcePeriod) {
      return '';
    }
    flow = this.buildClassFlowForValidation_(normalized, classCode);
    if (!flow.hasScheduledPeriods) {
      return 'Gestione classe non valida per ' + classCode + ': nessuna lezione programmata trovata.';
    }
    if (classHandlingType === this.CLASS_HANDLING_TYPES_.LATE_ENTRY) {
      transitionPeriod = this.findNextClassFlowPeriod_(flow, sourcePeriod, function (entryFlow) {
        return entryFlow.state === 'VALID';
      });
      if (!transitionPeriod) {
        return 'Entrata posticipata non valida per ' + classCode + ': nessuna ora utile di ingresso trovata dopo la ' + sourcePeriod + 'ª ora.';
      }
      invalidState = flow.periods.filter(function (entryFlow) {
        return Number(entryFlow.period || 0) < Number(transitionPeriod || 0) &&
          entryFlow.scheduled &&
          entryFlow.state !== 'LATE_ENTRY_BLOCK' &&
          entryFlow.state !== 'CLASS_OUT';
      })[0] || null;
      if (invalidState) {
        return 'Entrata posticipata non valida per ' + classCode + ': la classe ha gia una lezione valida o incoerente prima della ' + transitionPeriod + 'ª ora.';
      }
      if (this.countClassPresencePeriods_(flow, transitionPeriod, '') < this.MIN_CLASS_PRESENCE_PERIODS_) {
        return 'Entrata posticipata non valida per ' + classCode + ': la classe deve mantenere almeno 4 ore di presenza.';
      }
      return '';
    }
    transitionPeriod = this.findPreviousClassFlowPeriod_(flow, sourcePeriod, function (entryFlow) {
      return entryFlow.state === 'VALID';
    });
    if (!transitionPeriod) {
      return 'Uscita anticipata non valida per ' + classCode + ': la classe non ha ancora svolto una lezione valida prima della ' + sourcePeriod + 'ª ora.';
    }
    invalidState = flow.periods.filter(function (entryFlow) {
      return Number(entryFlow.period || 0) > Number(transitionPeriod || 0) &&
        entryFlow.scheduled &&
        entryFlow.state !== 'EARLY_EXIT_BLOCK' &&
        entryFlow.state !== 'CLASS_OUT';
    })[0] || null;
    if (invalidState) {
      return 'Uscita anticipata non valida per ' + classCode + ': restano lezioni successive non coerentemente chiuse dopo la ' + transitionPeriod + 'ª ora.';
    }
    if (this.countClassPresencePeriods_(flow, '', transitionPeriod) < this.MIN_CLASS_PRESENCE_PERIODS_) {
      return 'Uscita anticipata non valida per ' + classCode + ': la classe deve mantenere almeno 4 ore di presenza.';
    }
    return '';
  },

  buildClassFlowForValidation_: function (normalized, classCode) {
    var targetClassCode = ROOMS_APP.normalizeString(classCode).toUpperCase();
    var periodMap = ROOMS_APP.Timetable.getPeriodTimeMap();
    var periods = Object.keys(periodMap).sort(function (left, right) {
      return Number(left) - Number(right);
    });
    var assignmentsByClassPeriod = {};
    var scheduledByPeriod = {};
    (normalized.assignments || []).forEach(function (entry) {
      var key = [
        ROOMS_APP.normalizeString(entry.classCode).toUpperCase(),
        ROOMS_APP.normalizeString(entry.period)
      ].join('|');
      assignmentsByClassPeriod[key] = assignmentsByClassPeriod[key] || [];
      assignmentsByClassPeriod[key].push(entry);
    });
    (normalized.teachers || []).forEach(function (teacher) {
      Object.keys(teacher.periods || {}).forEach(function (period) {
        var slot = teacher.periods[period];
        if (!slot || !ROOMS_APP.Replacements.isCoverableTeacherSlot_(teacher, slot) || ROOMS_APP.normalizeString(slot.classCode).toUpperCase() !== targetClassCode) {
          return;
        }
        scheduledByPeriod[ROOMS_APP.normalizeString(period)] = true;
      });
    });
    return {
      classCode: targetClassCode,
      hasScheduledPeriods: Object.keys(scheduledByPeriod).length > 0,
      periods: periods.map(function (period) {
        var key = [targetClassCode, ROOMS_APP.normalizeString(period)].join('|');
        var periodAssignments = assignmentsByClassPeriod[key] || [];
        var classHandlingTypes = periodAssignments.map(function (assignment) {
          return ROOMS_APP.Replacements.normalizeClassHandlingType_(assignment.classHandlingType);
        }).filter(function (value) {
          return Boolean(value);
        });
        var hasClassHandling = Boolean(classHandlingTypes.length);
        var hasToAssign = periodAssignments.some(function (assignment) {
          return ROOMS_APP.normalizeString(assignment.replacementStatus) === 'TO_ASSIGN';
        });
        var hasClassOut = ROOMS_APP.Replacements.isClassOutAtPeriod_(normalized, targetClassCode, period) ||
          periodAssignments.some(function (assignment) {
            return ROOMS_APP.normalizeString(assignment.replacementStatus) === 'IN_USCITA';
          });
        var state = 'NO_CLASS';
        if (scheduledByPeriod[period]) {
          state = 'VALID';
          if (hasClassOut) {
            state = 'CLASS_OUT';
          } else if (classHandlingTypes.indexOf(ROOMS_APP.Replacements.CLASS_HANDLING_TYPES_.LATE_ENTRY) >= 0) {
            state = 'LATE_ENTRY_BLOCK';
          } else if (classHandlingTypes.indexOf(ROOMS_APP.Replacements.CLASS_HANDLING_TYPES_.EARLY_EXIT) >= 0) {
            state = 'EARLY_EXIT_BLOCK';
          } else if (hasToAssign) {
            state = 'UNCOVERED';
          }
        }
        return {
          period: period,
          scheduled: Boolean(scheduledByPeriod[period]),
          state: state
        };
      })
    };
  },

  findNextClassFlowPeriod_: function (flow, sourcePeriod, predicate) {
    var source = Number(sourcePeriod || 0);
    var match = (flow && flow.periods || []).filter(function (entry) {
      return Number(entry.period || 0) > source && predicate(entry);
    })[0] || null;
    return match ? match.period : '';
  },

  findPreviousClassFlowPeriod_: function (flow, sourcePeriod, predicate) {
    var source = Number(sourcePeriod || 0);
    var matches = (flow && flow.periods || []).filter(function (entry) {
      return Number(entry.period || 0) < source && predicate(entry);
    });
    return matches.length ? matches[matches.length - 1].period : '';
  },

  countClassPresencePeriods_: function (flow, startPeriod, endPeriod) {
    var start = Number(startPeriod || 0);
    var end = Number(endPeriod || 0);
    return (flow && flow.periods || []).filter(function (entry) {
      var period = Number(entry.period || 0);
      if (!entry.scheduled) {
        return false;
      }
      if (start && period < start) {
        return false;
      }
      if (end && period > end) {
        return false;
      }
      return entry.state === 'VALID' || entry.state === 'UNCOVERED';
    }).length;
  },

  getActualClassTransitionPeriod_: function (entry, classHandlingType) {
    var normalizedHandling = this.normalizeClassHandlingType_(classHandlingType || (entry && entry.classHandlingType));
    var actualPeriod = ROOMS_APP.normalizeString(entry && (normalizedHandling === this.CLASS_HANDLING_TYPES_.LATE_ENTRY
      ? entry.actualEntryPeriod
      : entry.actualExitPeriod));
    if (actualPeriod) {
      return actualPeriod;
    }
    return ROOMS_APP.normalizeString(entry && entry.period);
  },

  buildSummaryFromAssignments_: function (assignments, classes, teachers) {
    var summary = {
      classOutCount: 0,
      absentCount: 0,
      accompanistCount: 0,
      serviceOutOfClassCount: 0,
      hourlyAbsenceCount: 0,
      managedTeacherCount: 0,
      managedPeriodCount: 0,
      coverageNeedCount: 0,
      assignedCount: 0,
      recoveryCount: 0,
      shiftCount: 0,
      toAssignCount: 0,
      inUscitaCount: 0
    };
    var managedTeacherMap = {};

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
      if (entry.originalTeacherEmail) {
        managedTeacherMap[ROOMS_APP.Replacements.normalizeTeacherEmail_(entry.originalTeacherEmail)] = true;
      }
      summary.managedPeriodCount += 1;
      if (entry.replacementStatus !== 'IN_USCITA') {
        summary.coverageNeedCount += 1;
      }
      if (entry.originalStatus === 'HOURLY_ABSENCE') {
        summary.hourlyAbsenceCount += 1;
      }
      if (entry.originalStatus === 'SERVICE_OUT_OF_CLASS') {
        summary.serviceOutOfClassCount += 1;
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
    summary.managedTeacherCount = Object.keys(managedTeacherMap).length;
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

  buildReportArchiveKey_: function (reportType, referenceDate) {
    return ROOMS_APP.normalizeString(reportType).toUpperCase() + '|' + ROOMS_APP.toIsoDate(referenceDate);
  },

  archivePublishedReportSnapshot_: function (reportType, referenceDate, payload, options) {
    var archiveSheetName = ROOMS_APP.SHEET_NAMES.REPORT_ARCHIVE;
    var historySheetName = ROOMS_APP.SHEET_NAMES.REPORT_ARCHIVE_HISTORY;
    var targetDate = ROOMS_APP.toIsoDate(referenceDate);
    var normalizedType = ROOMS_APP.normalizeString(reportType).toUpperCase();
    var reportKey = this.buildReportArchiveKey_(normalizedType, targetDate);
    var actorEmail = ROOMS_APP.normalizeEmail(options && options.actorEmail);
    var status = ROOMS_APP.normalizeString(options && options.status) || 'PUBLISHED';
    var visibleToTeachers = options && Object.prototype.hasOwnProperty.call(options, 'visibleToTeachers')
      ? Boolean(options.visibleToTeachers)
      : true;
    var notes = ROOMS_APP.normalizeString(options && options.notes);
    var nowIso = ROOMS_APP.toIsoDateTime(new Date());
    var archiveHeaders = ROOMS_APP.DB.getHeaders(archiveSheetName);
    var currentRows = ROOMS_APP.DB.readRows(archiveSheetName);
    var matchingRows = currentRows.filter(function (row) {
      return ROOMS_APP.normalizeString(row.ReportKey) === reportKey;
    });
    var historyRows = matchingRows.map(function (row, index) {
      var version = Math.max(1, ROOMS_APP.asNumber(row.Version, index + 1));
      return {
        HistoryId: reportKey + '|V' + version + '|' + nowIso + '|' + index,
        ReportKey: reportKey,
        ReportType: normalizedType,
        ReferenceDate: targetDate,
        Subject: ROOMS_APP.normalizeString(row.Subject),
        HtmlSnapshot: String(row.HtmlSnapshot || ''),
        PdfFileId: ROOMS_APP.normalizeString(row.PdfFileId),
        Version: version,
        ArchivedAtISO: nowIso,
        ArchivedBy: actorEmail,
        Status: ROOMS_APP.normalizeString(row.Status) || 'ARCHIVED',
        Notes: ROOMS_APP.normalizeString(row.Notes)
      };
    });
    var nextVersion = 1;
    var createdAtISO = nowIso;
    var createdBy = actorEmail;
    var nextCurrentRows;
    var currentRow;

    if (matchingRows.length) {
      nextVersion = matchingRows.reduce(function (maxVersion, row) {
        return Math.max(maxVersion, ROOMS_APP.asNumber(row.Version, 0));
      }, 0) + 1;
      createdAtISO = ROOMS_APP.normalizeString(matchingRows[0].CreatedAtISO) || nowIso;
      createdBy = ROOMS_APP.normalizeEmail(matchingRows[0].CreatedBy) || actorEmail;
    }

    currentRow = {
      ReportKey: reportKey,
      ReportType: normalizedType,
      ReferenceDate: targetDate,
      Subject: ROOMS_APP.normalizeString(payload && payload.subject),
      HtmlSnapshot: String(payload && payload.htmlBody || ''),
      PdfFileId: '',
      Version: nextVersion,
      VisibleToTeachers: visibleToTeachers ? 'TRUE' : 'FALSE',
      CreatedAtISO: createdAtISO,
      CreatedBy: createdBy,
      UpdatedAtISO: nowIso,
      UpdatedBy: actorEmail,
      Status: status,
      Notes: notes
    };

    nextCurrentRows = currentRows.filter(function (row) {
      return ROOMS_APP.normalizeString(row.ReportKey) !== reportKey;
    }).concat([currentRow]);

    if (historyRows.length) {
      ROOMS_APP.DB.appendRows(historySheetName, historyRows);
    }
    ROOMS_APP.DB.replaceRows(archiveSheetName, archiveHeaders, nextCurrentRows);

    return {
      reportKey: reportKey,
      version: nextVersion,
      updatedAtISO: nowIso
    };
  },

  listVisibleArchivedReports_: function () {
    if (!ROOMS_APP.DB.getSheet(ROOMS_APP.SHEET_NAMES.REPORT_ARCHIVE)) {
      return [];
    }
    return ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.REPORT_ARCHIVE)
      .filter(function (row) {
        return ROOMS_APP.asBoolean(row.VisibleToTeachers);
      })
      .map(function (row) {
        var referenceDate = ROOMS_APP.toIsoDate(row.ReferenceDate);
        var version = Math.max(1, ROOMS_APP.asNumber(row.Version, 1));
        return {
          reportKey: ROOMS_APP.normalizeString(row.ReportKey),
          reportType: ROOMS_APP.normalizeString(row.ReportType).toUpperCase(),
          referenceDate: referenceDate,
          referenceDateLabel: ROOMS_APP.formatItalianExtendedDate(referenceDate) || referenceDate,
          subject: ROOMS_APP.normalizeString(row.Subject),
          version: version,
          status: version > 1 ? 'Aggiornato' : 'Pubblicato',
          updatedAtISO: ROOMS_APP.normalizeString(row.UpdatedAtISO),
          updatedBy: ROOMS_APP.normalizeEmail(row.UpdatedBy)
        };
      })
      .sort(function (left, right) {
        if (left.referenceDate > right.referenceDate) {
          return -1;
        }
        if (left.referenceDate < right.referenceDate) {
          return 1;
        }
        if (left.updatedAtISO > right.updatedAtISO) {
          return -1;
        }
        if (left.updatedAtISO < right.updatedAtISO) {
          return 1;
        }
        return 0;
      });
  },

  getVisibleArchivedReportByKey_: function (reportKey) {
    var targetKey = ROOMS_APP.normalizeString(reportKey);
    if (!targetKey) {
      return null;
    }
    if (!ROOMS_APP.DB.getSheet(ROOMS_APP.SHEET_NAMES.REPORT_ARCHIVE)) {
      return null;
    }
    var row = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.REPORT_ARCHIVE).filter(function (entry) {
      return ROOMS_APP.normalizeString(entry.ReportKey) === targetKey &&
        ROOMS_APP.asBoolean(entry.VisibleToTeachers);
    })[0] || null;
    var referenceDate;
    var version;
    if (!row) {
      return null;
    }
    referenceDate = ROOMS_APP.toIsoDate(row.ReferenceDate);
    version = Math.max(1, ROOMS_APP.asNumber(row.Version, 1));
    return {
      reportKey: targetKey,
      reportType: ROOMS_APP.normalizeString(row.ReportType).toUpperCase(),
      referenceDate: referenceDate,
      referenceDateLabel: ROOMS_APP.formatItalianExtendedDate(referenceDate) || referenceDate,
      subject: ROOMS_APP.normalizeString(row.Subject),
      htmlSnapshot: String(row.HtmlSnapshot || ''),
      pdfFileId: ROOMS_APP.normalizeString(row.PdfFileId),
      version: version,
      status: version > 1 ? 'Aggiornato' : 'Pubblicato',
      updatedAtISO: ROOMS_APP.normalizeString(row.UpdatedAtISO),
      updatedBy: ROOMS_APP.normalizeEmail(row.UpdatedBy)
    };
  },

  getSavedReflectionState_: function (dateString) {
    var targetDate = ROOMS_APP.toIsoDate(dateString || new Date());
    var tripState = this.mergeTripDayStates_(
      this.buildTripDayState_(targetDate, this.listEducationalTrips_()),
      this.buildLegacyOutingState_(targetDate, this.listRowsForDate_(ROOMS_APP.SHEET_NAMES.REPL_CLASS_OUT, targetDate), [])
    );
    var assignmentMap = {};
    var shiftOriginMap = {};
    var activeLongAssignmentMap = this.getActiveLongAssignmentsForDate_(targetDate, true);

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
        Object.keys(tripState.classOutSet || {}).length ||
        Object.keys(assignmentMap).length ||
        Object.keys(shiftOriginMap).length ||
        Object.keys(activeLongAssignmentMap).length
      ),
      classOutSet: tripState.classOutSet || {},
      classOutSetKeys: Object.keys(tripState.classOutSet || {}),
      classOutPeriodsByClass: tripState.classOutPeriodsByClass || {},
      teacherTripPeriodsByTeacher: tripState.teacherTripPeriodsByTeacher || {},
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
    var persistedAssignmentKeysForDate = {};
    var recoveredByKey = {};
    var nextRows;

    allRows.forEach(function (row) {
      existingByKey[ROOMS_APP.Replacements.buildHourlyAbsenceDateKey_(row.date, row.teacherEmail, row.period)] = row;
    });

    this.listRowsForDate_(ROOMS_APP.SHEET_NAMES.REPL_ASSIGNMENTS, targetDate).forEach(function (row) {
      persistedAssignmentKeysForDate[ROOMS_APP.Replacements.buildAssignmentLinkKey_(
        targetDate,
        row && row.Period,
        row && row.ClassCode,
        row && row.OriginalTeacherEmail
      )] = true;
    });

    (normalized.assignments || []).forEach(function (entry) {
      var assignmentLinkKey = ROOMS_APP.Replacements.buildAssignmentLinkKey_(
        targetDate,
        entry.period,
        entry.classCode,
        entry.originalTeacherEmail
      );
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
      if (persistedAssignmentKeysForDate[next.recoveredByAssignmentKey] && !recoveredByKey[key]) {
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
        reason: ROOMS_APP.normalizeString(entry.reason),
        classCode: ROOMS_APP.normalizeString(entry.classCode).toUpperCase(),
        label: ROOMS_APP.normalizeString(entry.label),
        slotType: ROOMS_APP.normalizeString(entry.slotType),
        startTime: ROOMS_APP.normalizeString(entry.startTime),
        endTime: ROOMS_APP.normalizeString(entry.endTime),
        derivedActionable: Boolean(entry.derivedActionable),
        sourceAbsenceId: ROOMS_APP.normalizeString(entry.sourceAbsenceId),
        sourceType: ROOMS_APP.normalizeString(entry.sourceType),
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
      return ROOMS_APP.Replacements.serializeHourlyAbsenceRow_(row, nowIso, updatedBy);
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
      classCode: ROOMS_APP.normalizeString(row && row.ClassCode).toUpperCase(),
      label: ROOMS_APP.normalizeString(row && row.Label),
      slotType: ROOMS_APP.normalizeString(row && row.SlotType),
      startTime: ROOMS_APP.normalizeString(row && row.StartTime),
      endTime: ROOMS_APP.normalizeString(row && row.EndTime),
      derivedActionable: ROOMS_APP.asBoolean(row && row.DerivedActionable),
      sourceAbsenceId: ROOMS_APP.normalizeString(row && row.SourceAbsenceId),
      sourceType: ROOMS_APP.normalizeString(row && row.SourceType),
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
      classCode: ROOMS_APP.normalizeString(entry && (entry.classCode || entry.ClassCode)).toUpperCase(),
      label: ROOMS_APP.normalizeString(entry && (entry.label || entry.Label)),
      slotType: ROOMS_APP.normalizeString(entry && (entry.slotType || entry.SlotType)),
      startTime: ROOMS_APP.normalizeString(entry && (entry.startTime || entry.StartTime)),
      endTime: ROOMS_APP.normalizeString(entry && (entry.endTime || entry.EndTime)),
      derivedActionable: ROOMS_APP.asBoolean(entry && (entry.derivedActionable || entry.DerivedActionable)),
      sourceAbsenceId: ROOMS_APP.normalizeString(entry && (entry.sourceAbsenceId || entry.SourceAbsenceId)),
      sourceType: ROOMS_APP.normalizeString(entry && (entry.sourceType || entry.SourceType)),
      recoveryRequired: ROOMS_APP.asBoolean(entry && Object.prototype.hasOwnProperty.call(entry, 'recoveryRequired') ? entry.recoveryRequired : entry && entry.RecoveryRequired),
      recoveryStatus: ROOMS_APP.normalizeString(entry && (entry.recoveryStatus || entry.RecoveryStatus)) || this.RECOVERY_STATUSES_.PENDING,
      recoveredOnDate: ROOMS_APP.toIsoDate(entry && (entry.recoveredOnDate || entry.RecoveredOnDate)),
      recoveredByAssignmentKey: ROOMS_APP.normalizeString(entry && (entry.recoveredByAssignmentKey || entry.RecoveredByAssignmentKey)),
      notes: ROOMS_APP.normalizeString(entry && (entry.notes || entry.Notes))
    };
  },

  isServiceOutOfClassHourlyAbsenceEntry_: function (entry) {
    return ROOMS_APP.normalizeString(entry && (entry.reason || entry.Reason)).toUpperCase() === this.SERVICE_OUT_OF_CLASS_.GENERAL;
  },

  buildValidHourlyAbsenceMap_: function (entries, targetDate, teacherMap) {
    var map = {};
    (entries || []).forEach(function (entry) {
      var normalizedEntry = ROOMS_APP.Replacements.normalizeHourlyAbsenceEntry_(entry, targetDate);
      var teacher = teacherMap && teacherMap[normalizedEntry.teacherEmail] ? teacherMap[normalizedEntry.teacherEmail] : null;
      var slot = teacher && teacher.periods ? teacher.periods[normalizedEntry.period] : null;
      var assignmentClassCode = ROOMS_APP.Replacements.getTeacherSlotAssignmentClassCode_(slot) || normalizedEntry.classCode;
      if (!normalizedEntry.teacherEmail || !normalizedEntry.period) {
        return;
      }
      if (!teacher || teacher.absent || !assignmentClassCode) {
        return;
      }
      if (!normalizedEntry.derivedActionable &&
          (!slot || !ROOMS_APP.Replacements.isCoverableTeacherSlot_(teacher, slot))) {
        return;
      }
      if (!normalizedEntry.teacherName) {
        normalizedEntry.teacherName = teacher.teacherName;
      }
      normalizedEntry.classCode = assignmentClassCode;
      map[ROOMS_APP.Replacements.buildHourlyAbsenceKey_(normalizedEntry.teacherEmail, normalizedEntry.period)] = normalizedEntry;
    });
    return map;
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
    if (!teacherEmail || !teacher || teacher.absent || this.isTeacherAccompanyingAtPeriod_(normalized, teacherEmail, period)) {
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
    if (!teacherEmail || !teacher || teacher.absent || this.isTeacherAccompanyingAtPeriod_(normalized, teacherEmail, targetPeriod)) {
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
    return this.listAdminTeacherDirectory_();
  },

  listAdminTeacherDirectory_: function () {
    var teacherMap = {};
    this.getDocentiTimetableDerivedData_().teacherDirectory.concat(this.listSupportTeacherDirectory_()).forEach(function (entry) {
      var teacherEmail = ROOMS_APP.Replacements.normalizeTeacherEmail_(entry && entry.teacherEmail);
      var teacherName = ROOMS_APP.normalizeString(entry && entry.teacherName);
      if (!teacherEmail) {
        return;
      }
      teacherMap[teacherEmail] = {
        teacherEmail: teacherEmail,
        teacherName: teacherName || teacherEmail
      };
    });
    return Object.keys(teacherMap).map(function (teacherEmail) {
      return teacherMap[teacherEmail];
    }).sort(function (left, right) {
      return left.teacherName.localeCompare(right.teacherName);
    }).map(function (entry) {
      return {
        teacherEmail: entry.teacherEmail,
        teacherName: entry.teacherName
      };
    });
  },

  listSupportTeacherDirectory_: function () {
    var supportMap = {};
    this.listSupportTeacherEntries_().forEach(function (entry) {
      if (!entry.teacherEmail) {
        return;
      }
      supportMap[entry.teacherEmail] = {
        teacherEmail: entry.teacherEmail,
        teacherName: entry.teacherName || entry.teacherEmail
      };
    });
    return Object.keys(supportMap).map(function (teacherEmail) {
      return supportMap[teacherEmail];
    }).sort(function (left, right) {
      return left.teacherName.localeCompare(right.teacherName);
    });
  },

  buildTeacherDayTeachers_: function (dateString) {
    var targetDate = ROOMS_APP.toIsoDate(dateString || new Date());
    var weekday = ROOMS_APP.getWeekdayName(targetDate);
    var activeLongAssignments = this.getActiveLongAssignmentsForDate_(targetDate, true);
    var snapshot = this.getDocentiTimetableSnapshot_();
    var opening = ROOMS_APP.Policy.getDailyOpening(targetDate);
    var teacherMap = {};
    var rowIndex;

    if (snapshot.sheet) {
      for (rowIndex = snapshot.columnMeta.dataStartRow; rowIndex < snapshot.values.length; rowIndex += 1) {
        var rowLabel = ROOMS_APP.normalizeString(snapshot.values[rowIndex] && snapshot.values[rowIndex][0]);
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
          originalTeacherEmails: {},
          supportTeacher: false
        };
        if (originalTeacherEmail) {
          teacher.originalTeacherEmails[originalTeacherEmail] = rowLabel;
        }

        snapshot.columnMeta.usableColumns.forEach(function (column) {
          var meta = snapshot.columnMeta.columns[column] || {};
          if (meta.weekday !== weekday) {
            return;
          }
          if (!ROOMS_APP.Replacements.isPeriodActiveForOpening_(opening, meta.startTime, meta.endTime)) {
            return;
          }
          var rawValue = ROOMS_APP.normalizeString(snapshot.values[rowIndex][column]);
          teacher.periods[String(meta.period)] = ROOMS_APP.Replacements.classifyTeacherSlotValue_(rawValue, meta.period, meta.startTime, meta.endTime);
        });

        teacherMap[teacherEmail] = teacher;
      }
    }

    this.buildSupportTeacherDayTeachers_(targetDate).forEach(function (supportTeacher) {
      var teacher = teacherMap[supportTeacher.teacherEmail] || {
        teacherEmail: supportTeacher.teacherEmail,
        teacherName: supportTeacher.teacherName,
        periods: {},
        originalTeacherEmails: {},
        supportTeacher: true
      };
      teacher.teacherName = teacher.teacherName || supportTeacher.teacherName;
      teacher.supportTeacher = true;
      Object.keys(supportTeacher.periods || {}).forEach(function (period) {
        var existingSlot = teacher.periods[period] || null;
        if (!existingSlot || existingSlot.type === 'FREE') {
          teacher.periods[period] = supportTeacher.periods[period];
        }
      });
      teacherMap[supportTeacher.teacherEmail] = teacher;
    });

    return Object.keys(teacherMap).map(function (teacherEmail) {
      return teacherMap[teacherEmail];
    }).sort(function (left, right) {
      return left.teacherName.localeCompare(right.teacherName);
    });
  },

  buildSupportTeacherDayTeachers_: function (dateString) {
    var targetDate = ROOMS_APP.toIsoDate(dateString || new Date());
    var weekday = ROOMS_APP.getWeekdayName(targetDate);
    var periodMap = ROOMS_APP.Timetable.getPeriodTimeMap();
    var opening = ROOMS_APP.Policy.getDailyOpening(targetDate);
    var teacherMap = {};
    var self = this;
    ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.TIMETABLE_DOCENTI_SOSTEGNO).forEach(function (row) {
      var supportEntry = self.normalizeSupportTeacherRow_(row);
      var teacher;
      if (!supportEntry || !supportEntry.teacherEmail) {
        return;
      }
      teacher = teacherMap[supportEntry.teacherEmail] || {
        teacherEmail: supportEntry.teacherEmail,
        teacherName: supportEntry.teacherName,
        periods: {},
        supportTeacher: true
      };
      Object.keys(periodMap).forEach(function (period) {
        var meta = periodMap[period] || {};
        var header = self.getSupportTeacherPeriodHeader_(weekday, meta.startTime);
        var rawValue = header ? ROOMS_APP.normalizeString(row && row[header]) : '';
        if (!self.isPeriodActiveForOpening_(opening, meta.startTime, meta.endTime)) {
          return;
        }
        if (!rawValue) {
          return;
        }
        teacher.periods[String(period)] = self.classifyTeacherSlotValue_(rawValue, period, meta.startTime, meta.endTime, {
          supportTeacher: true
        });
      });
      teacherMap[supportEntry.teacherEmail] = teacher;
    });
    return Object.keys(teacherMap).map(function (teacherEmail) {
      return teacherMap[teacherEmail];
    }).sort(function (left, right) {
      return left.teacherName.localeCompare(right.teacherName);
    });
  },

  getSupportTeacherPeriodHeader_: function (weekday, startTime) {
    var prefixMap = {
      Monday: 'LUN',
      Tuesday: 'MAR',
      Wednesday: 'MER',
      Thursday: 'GIO',
      Friday: 'VEN'
    };
    var prefix = prefixMap[weekday] || '';
    var hour = Number(String(startTime || '').slice(0, 2));
    if (!prefix || !hour) {
      return '';
    }
    return prefix + '_' + String(hour);
  },

  isPeriodActiveForOpening_: function (opening, startTime, endTime) {
    var dailyOpening = opening || {};
    if (dailyOpening.isOpen === false) {
      return false;
    }
    var openTime = ROOMS_APP.normalizeString(dailyOpening.openTime);
    var closeTime = ROOMS_APP.normalizeString(dailyOpening.closeTime);
    var start = ROOMS_APP.normalizeString(startTime);
    var end = ROOMS_APP.normalizeString(endTime);
    if (openTime && end && end <= openTime) {
      return false;
    }
    if (closeTime && start && start >= closeTime) {
      return false;
    }
    return true;
  },

  classifyTeacherSlotValue_: function (rawValue, period, startTime, endTime, options) {
    var normalized = ROOMS_APP.normalizeString(rawValue).toUpperCase();
    var settings = options && typeof options === 'object' ? options : {};
    var serviceSubtype = this.normalizeServiceOutOfClassSubtype_(normalized);
    var classCode = ROOMS_APP.Timetable.extractClassCode_(normalized);
    var type = 'FREE';
    var label = '';
    var coverageRequired = false;

    if (serviceSubtype) {
      type = 'SERVICE';
      label = this.getServiceOutOfClassLabel_(serviceSubtype, classCode);
      if (!classCode && serviceSubtype === this.SERVICE_OUT_OF_CLASS_.ALTERNATIVA) {
        classCode = this.SERVICE_OUT_OF_CLASS_.ALTERNATIVA;
      }
      coverageRequired = serviceSubtype === this.SERVICE_OUT_OF_CLASS_.ALTERNATIVA || Boolean(classCode);
    } else if (classCode) {
      if (settings.supportTeacher) {
        type = 'SUPPORT';
        label = classCode;
      } else {
        type = 'CLASS';
        label = classCode;
        coverageRequired = true;
      }
    } else if (normalized === 'P') {
      type = 'P';
      label = 'P';
    } else if (normalized === 'D') {
      type = 'D';
      label = 'D';
    } else if (settings.supportTeacher && normalized) {
      type = 'SUPPORT';
      label = normalized;
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
      label: label,
      supportSlot: Boolean(settings.supportTeacher),
      serviceType: serviceSubtype ? this.SERVICE_OUT_OF_CLASS_.GENERAL : '',
      serviceSubtype: serviceSubtype,
      coverageRequired: coverageRequired
    };
  },

  normalizeServiceOutOfClassSubtype_: function (value) {
    var normalized = ROOMS_APP.normalizeString(value).toUpperCase();
    if (!normalized) {
      return '';
    }
    if (normalized === 'AL' || /^AL[\s._-]/.test(normalized) || /ALTERNATIV/.test(normalized)) {
      return this.SERVICE_OUT_OF_CLASS_.ALTERNATIVA;
    }
    if (/ASSEMBLEA|ASSEMBL/.test(normalized)) {
      return this.SERVICE_OUT_OF_CLASS_.ASSEMBLY;
    }
    if (/ACCOMPAGN|USCITA|USCITE/.test(normalized)) {
      return this.SERVICE_OUT_OF_CLASS_.ACCOMPANIMENT;
    }
    if (/SERVIZIO\s+FUORI\s+AULA|FUORI\s+AULA/.test(normalized)) {
      return this.SERVICE_OUT_OF_CLASS_.GENERAL;
    }
    return '';
  },

  getServiceOutOfClassLabel_: function (serviceSubtype, classCode) {
    var subtype = ROOMS_APP.normalizeString(serviceSubtype);
    var label = subtype === this.SERVICE_OUT_OF_CLASS_.ALTERNATIVA
      ? 'Alternativa'
      : (subtype === this.SERVICE_OUT_OF_CLASS_.ASSEMBLY
        ? 'Servizio fuori aula - Assemblea'
        : (subtype === this.SERVICE_OUT_OF_CLASS_.ACCOMPANIMENT
          ? 'Servizio fuori aula - Accompagnamento'
          : 'Servizio fuori aula'));
    return classCode ? label + ' ' + classCode : label;
  },

  isCoverableTeacherSlot_: function (teacher, slot) {
    if (!slot) {
      return false;
    }
    if (slot.coverageRequired) {
      return true;
    }
    if (this.getTeacherSlotAssignmentClassCode_(slot) === this.SERVICE_OUT_OF_CLASS_.ALTERNATIVA) {
      return true;
    }
    return slot.type === 'CLASS';
  },

  getTeacherSlotAssignmentClassCode_: function (slot) {
    var classCode = ROOMS_APP.normalizeString(slot && slot.classCode).toUpperCase();
    var rawValue = ROOMS_APP.normalizeString(slot && slot.rawValue).toUpperCase();
    var serviceSubtype = ROOMS_APP.normalizeString(slot && slot.serviceSubtype).toUpperCase();
    if (classCode) {
      return classCode;
    }
    if (serviceSubtype === this.SERVICE_OUT_OF_CLASS_.ALTERNATIVA || rawValue === 'AL') {
      return this.SERVICE_OUT_OF_CLASS_.ALTERNATIVA;
    }
    return '';
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

  prepareDailyContextCacheAfterSourceChange_: function (primaryDate, affectedDates, notes) {
    var targetDate = ROOMS_APP.toIsoDate(primaryDate);
    var dateMap = {};
    var dateList;
    (affectedDates || []).forEach(function (dateValue) {
      var dateKey = ROOMS_APP.toIsoDate(dateValue);
      if (dateKey) {
        dateMap[dateKey] = true;
      }
    });
    if (targetDate) {
      dateMap[targetDate] = true;
    }
    dateList = Object.keys(dateMap).sort();
    if (!dateList.length) {
      return null;
    }
    try {
      dateList.forEach(function (dateKey) {
        ROOMS_APP.Replacements.invalidateDailyContextCache_(dateKey, notes || 'SOURCE_CHANGED');
      });
      if (targetDate) {
        return this.refreshDailyContextCache_(targetDate, notes || 'SOURCE_CHANGED');
      }
    } catch (error) {
      if (targetDate) {
        this.invalidateDailyContextCache_(targetDate, (notes || 'SOURCE_CHANGED') + ': ' + (error && error.message ? error.message : error));
      }
    }
    return null;
  },

  collectDateRangeCacheDates_: function (startDate, endDate, fallbackDate) {
    var start = ROOMS_APP.toIsoDate(startDate || fallbackDate);
    var end = ROOMS_APP.toIsoDate(endDate || start);
    var dates = {};
    var cursor;
    var guard = 0;
    if (!start) {
      return [];
    }
    if (!end || end < start) {
      end = start;
    }
    cursor = new Date(start + 'T00:00:00Z');
    while (ROOMS_APP.toIsoDate(cursor) <= end && guard < 370) {
      dates[ROOMS_APP.toIsoDate(cursor)] = true;
      cursor.setUTCDate(cursor.getUTCDate() + 1);
      guard += 1;
    }
    return Object.keys(dates).sort();
  },

  collectAbsenceCacheDates_: function (rows, fallbackDate) {
    var dateMap = {};
    var self = this;
    (rows || []).forEach(function (row) {
      self.collectDateRangeCacheDates_(
        row && (row.startDate || row.StartDate),
        row && (row.endDate || row.EndDate),
        fallbackDate
      ).forEach(function (dateKey) {
        dateMap[dateKey] = true;
      });
    });
    if (fallbackDate) {
      dateMap[ROOMS_APP.toIsoDate(fallbackDate)] = true;
    }
    return Object.keys(dateMap).filter(Boolean).sort();
  },

  collectTripCacheDates_: function (trips, fallbackDate) {
    var dateMap = {};
    var self = this;
    (trips || []).forEach(function (trip) {
      self.collectDateRangeCacheDates_(trip.startDate, trip.endDate, fallbackDate).forEach(function (dateKey) {
        dateMap[dateKey] = true;
      });
    });
    if (fallbackDate) {
      dateMap[ROOMS_APP.toIsoDate(fallbackDate)] = true;
    }
    return Object.keys(dateMap).filter(Boolean).sort();
  },

  getReplacementRequestCacheBucket_: function () {
    if (!ROOMS_APP.DB_REQUEST_CACHE_) {
      return null;
    }
    ROOMS_APP.DB_REQUEST_CACHE_.replacements = ROOMS_APP.DB_REQUEST_CACHE_.replacements || {};
    return ROOMS_APP.DB_REQUEST_CACHE_.replacements;
  },

  DAILY_CONTEXT_CACHE_KEY_: 'DAY_CONTEXT_V2',
  DAILY_CONTEXT_CACHE_CHUNK_SIZE_: 40000,

  readDailyContextCache_: function (dateString) {
    var targetDate = ROOMS_APP.toIsoDate(dateString);
    var cacheKey = this.DAILY_CONTEXT_CACHE_KEY_;
    var rows;
    var chunks;
    var payload;
    if (!targetDate) {
      return null;
    }
    rows = ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.REPL_DAILY_CACHE).filter(function (row) {
      return ROOMS_APP.toIsoDate(row.Date) === targetDate &&
        ROOMS_APP.normalizeString(row.CacheKey) === cacheKey &&
        ROOMS_APP.normalizeString(row.Status).toUpperCase() === 'READY';
    });
    if (!rows.length) {
      return null;
    }
    chunks = rows.sort(function (left, right) {
      return Number(left.ChunkIndex || 0) - Number(right.ChunkIndex || 0);
    });
    if (Number(chunks[0].ChunkCount || 0) !== chunks.length) {
      return null;
    }
    payload = chunks.map(function (row) {
      return String(row.PayloadJson || '');
    }).join('');
    try {
      return JSON.parse(payload);
    } catch (error) {
      return null;
    }
  },

  writeDailyContextCache_: function (dateString, context, notes) {
    var targetDate = ROOMS_APP.toIsoDate(dateString);
    var payload = JSON.stringify(context || {});
    var chunkSize = this.DAILY_CONTEXT_CACHE_CHUNK_SIZE_;
    var chunkCount = Math.max(1, Math.ceil(payload.length / chunkSize));
    var nowIso = ROOMS_APP.toIsoDateTime(new Date());
    var actorEmail = '';
    var sourceHash;
    var rows = [];
    var index;
    try {
      actorEmail = ROOMS_APP.Auth.getEffectiveUser().email || '';
    } catch (error) {
      actorEmail = '';
    }
    if (!targetDate) {
      return;
    }
    sourceHash = this.computeDailyContextSourceHash_(context);
    for (index = 0; index < chunkCount; index += 1) {
      rows.push({
        Date: targetDate,
        CacheKey: this.DAILY_CONTEXT_CACHE_KEY_,
        ChunkIndex: String(index),
        ChunkCount: String(chunkCount),
        PayloadJson: payload.slice(index * chunkSize, (index + 1) * chunkSize),
        Status: 'READY',
        SourceHash: sourceHash,
        UpdatedAtISO: nowIso,
        UpdatedBy: actorEmail,
        Notes: ROOMS_APP.normalizeString(notes)
      });
    }
    this.replaceDailyContextCacheRows_(targetDate, rows);
  },

  invalidateDailyContextCache_: function (dateString, notes) {
    var targetDate = ROOMS_APP.toIsoDate(dateString);
    var nowIso = ROOMS_APP.toIsoDateTime(new Date());
    if (!targetDate) {
      return;
    }
    this.replaceDailyContextCacheRows_(targetDate, [{
      Date: targetDate,
      CacheKey: this.DAILY_CONTEXT_CACHE_KEY_,
      ChunkIndex: '0',
      ChunkCount: '1',
      PayloadJson: '',
      Status: 'DIRTY',
      SourceHash: '',
      UpdatedAtISO: nowIso,
      UpdatedBy: '',
      Notes: ROOMS_APP.normalizeString(notes)
    }]);
  },

  refreshDailyContextCache_: function (dateString, notes) {
    var targetDate = ROOMS_APP.toIsoDate(dateString);
    var context;
    if (!targetDate) {
      return null;
    }
    context = this.buildDayContext_(targetDate, {
      includeEditorData: false,
      useDailyCache: false,
      writeDailyCache: false
    });
    this.writeDailyContextCache_(targetDate, context, notes || 'REFRESHED');
    return context;
  },

  replaceDailyContextCacheRows_: function (targetDate, replacementRows) {
    var sheetName = ROOMS_APP.SHEET_NAMES.REPL_DAILY_CACHE;
    var headers = ROOMS_APP.DB.getHeaders(sheetName);
    var cacheKey = this.DAILY_CONTEXT_CACHE_KEY_;
    var keepRows = ROOMS_APP.DB.readRows(sheetName).filter(function (row) {
      return !(ROOMS_APP.toIsoDate(row.Date) === targetDate &&
        ROOMS_APP.normalizeString(row.CacheKey) === cacheKey);
    });
    ROOMS_APP.DB.replaceRows(sheetName, headers, keepRows.concat(replacementRows || []));
  },

  computeDailyContextSourceHash_: function (context) {
    var source = {
      savedAtISO: context && context.savedAtISO || '',
      reportStatus: context && context.reportStatus || {},
      absenceRegistry: context && context.absenceRegistryState && context.absenceRegistryState.sourceRows || [],
      trips: (context && context.tripDayState && context.tripDayState.sourceRows || []).concat(
        context && context.tripDayState && context.tripDayState.sourceTeacherRows || []
      )
    };
    return Utilities.base64EncodeWebSafe(Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      JSON.stringify(source)
    ));
  },

  getDocentiTimetableSnapshot_: function () {
    var sheetName = ROOMS_APP.Timetable.getConfiguredSourceSheetName_(
      ROOMS_APP.Timetable.CONFIG_DOCENTI_SHEET_KEY_,
      ROOMS_APP.Timetable.DEFAULT_DOCENTI_SHEET_
    );
    var bucket = this.getReplacementRequestCacheBucket_();
    var cacheKey = 'docentiSnapshot:' + sheetName;
    var sheet;
    var snapshot;
    if (bucket && bucket[cacheKey]) {
      return bucket[cacheKey];
    }
    sheet = ROOMS_APP.DB.getSheet(sheetName);
    snapshot = {
      sheetName: sheetName,
      sheet: sheet,
      values: [],
      columnMeta: {
        dataStartRow: 0,
        usableColumns: [],
        columns: {}
      }
    };
    if (sheet) {
      snapshot.values = ROOMS_APP.Timetable.readSheetDisplayValues_(sheet);
      snapshot.columnMeta = ROOMS_APP.Timetable.detectColumnMeta_(snapshot.values);
    }
    if (bucket) {
      bucket[cacheKey] = snapshot;
    }
    return snapshot;
  },

  getDocentiTimetableDerivedData_: function () {
    var snapshot = this.getDocentiTimetableSnapshot_();
    var bucket = this.getReplacementRequestCacheBucket_();
    var cacheKey = 'docentiDerived:' + snapshot.sheetName;
    var teacherDirectoryMap = {};
    var tripByClass = {};
    var teacherDirectory;
    var derived;
    var self = this;
    var rowIndex;
    if (bucket && bucket[cacheKey]) {
      return bucket[cacheKey];
    }
    for (rowIndex = snapshot.columnMeta.dataStartRow; rowIndex < snapshot.values.length; rowIndex += 1) {
      var rowLabel = ROOMS_APP.normalizeString(snapshot.values[rowIndex] && snapshot.values[rowIndex][0]);
      var teacherEmail;
      if (!ROOMS_APP.Timetable.isRowLabelData_(rowLabel, 'classroom')) {
        continue;
      }
      teacherEmail = this.buildTeacherSyntheticEmail_(rowLabel);
      if (!teacherEmail) {
        continue;
      }
      teacherDirectoryMap[teacherEmail] = {
        teacherEmail: teacherEmail,
        teacherName: rowLabel
      };
      snapshot.columnMeta.usableColumns.forEach(function (column) {
        var classCode = ROOMS_APP.Timetable.extractClassCode_(ROOMS_APP.normalizeString(snapshot.values[rowIndex][column]).toUpperCase());
        if (!classCode) {
          return;
        }
        self.addTripTeacherOption_(tripByClass, classCode, {
          teacherEmail: teacherEmail,
          teacherName: rowLabel
        });
      });
    }
    teacherDirectory = Object.keys(teacherDirectoryMap).map(function (teacherEmail) {
      return teacherDirectoryMap[teacherEmail];
    }).sort(function (left, right) {
      return left.teacherName.localeCompare(right.teacherName);
    });
    Object.keys(tripByClass).forEach(function (classCode) {
      tripByClass[classCode] = tripByClass[classCode].sort(function (left, right) {
        return left.teacherName.localeCompare(right.teacherName);
      });
    });
    derived = {
      teacherDirectory: teacherDirectory,
      tripRegistry: {
        byClass: tripByClass,
        classOptions: Object.keys(tripByClass).sort()
      }
    };
    if (bucket) {
      bucket[cacheKey] = derived;
    }
    return derived;
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

  buildTripId_: function () {
    return 'trip-' + String(new Date().getTime()) + '-' + String(Math.floor(Math.random() * 100000));
  },

  buildAbsenceId_: function () {
    return Utilities.getUuid();
  },

  normalizeAbsenceMode_: function (value) {
    var normalized = ROOMS_APP.normalizeString(value).toUpperCase();
    if (normalized === this.ABSENCE_MODES_.PLANNED) {
      return this.ABSENCE_MODES_.PLANNED;
    }
    return this.ABSENCE_MODES_.DAY;
  },

  normalizeAbsenceType_: function (value) {
    var normalized = ROOMS_APP.normalizeString(value).toUpperCase();
    if (normalized === 'PERMESSO_ORARIO' || normalized === this.ABSENCE_TYPES_.HOURLY_PERMISSION) {
      return this.ABSENCE_TYPES_.HOURLY_PERMISSION;
    }
    if (normalized === 'SERVIZIO_FUORI_AULA' ||
        normalized === 'SERVICE_OUT_OF_CLASS' ||
        normalized === this.ABSENCE_TYPES_.SERVICE_OUT_OF_CLASS) {
      return this.ABSENCE_TYPES_.SERVICE_OUT_OF_CLASS;
    }
    return this.ABSENCE_TYPES_.DAILY;
  },

  normalizeServiceOutOfClassScope_: function (value) {
    var normalized = ROOMS_APP.normalizeString(value).toUpperCase();
    if (normalized === 'ORARIA' || normalized === 'HOURLY') {
      return 'HOURLY';
    }
    return 'DAILY';
  },

  isDailyServiceOutOfClassAbsence_: function (entry) {
    return Boolean(entry && entry.absenceType === this.ABSENCE_TYPES_.SERVICE_OUT_OF_CLASS &&
      this.normalizeServiceOutOfClassScope_(entry.serviceOutOfClassScope) === 'DAILY');
  },

  isHourlyServiceOutOfClassAbsence_: function (entry) {
    return Boolean(entry && entry.absenceType === this.ABSENCE_TYPES_.SERVICE_OUT_OF_CLASS &&
      this.normalizeServiceOutOfClassScope_(entry.serviceOutOfClassScope) === 'HOURLY');
  },

  normalizeAbsencePayload_: function (payload) {
    var teacherEmail = this.normalizeTeacherEmail_(payload && (payload.teacherEmail || payload.TeacherEmail));
    var teacherName = ROOMS_APP.normalizeString(payload && (payload.teacherName || payload.TeacherName));
    var absenceType = this.normalizeAbsenceType_(payload && (payload.absenceType || payload.AbsenceType));
    var rawServiceOutOfClassScope = ROOMS_APP.normalizeString(payload && (
      payload.serviceOutOfClassScope ||
      payload.ServiceOutOfClassScope ||
      payload.serviceScope ||
      payload.ServiceScope
    ));
    var serviceOutOfClassScope = absenceType === this.ABSENCE_TYPES_.SERVICE_OUT_OF_CLASS && rawServiceOutOfClassScope
      ? this.normalizeServiceOutOfClassScope_(rawServiceOutOfClassScope)
      : (absenceType === this.ABSENCE_TYPES_.SERVICE_OUT_OF_CLASS ? 'DAILY' : '');
    var isDailyServiceOutOfClass;
    var isMultiDay;
    var hourlyPeriods = {};
    (Array.isArray(payload && (payload.hourlyPeriods || payload.HourlyPeriods)) ? (payload.hourlyPeriods || payload.HourlyPeriods) : []).forEach(function (entry) {
      var normalizedPeriod = ROOMS_APP.normalizeString(entry);
      if (normalizedPeriod) {
        hourlyPeriods[normalizedPeriod] = true;
      }
    });
    if (absenceType === this.ABSENCE_TYPES_.SERVICE_OUT_OF_CLASS && !rawServiceOutOfClassScope && Object.keys(hourlyPeriods).length) {
      serviceOutOfClassScope = 'HOURLY';
    }
    isDailyServiceOutOfClass = absenceType === this.ABSENCE_TYPES_.SERVICE_OUT_OF_CLASS && serviceOutOfClassScope === 'DAILY';
    isMultiDay = (absenceType === this.ABSENCE_TYPES_.DAILY || isDailyServiceOutOfClass) &&
      ROOMS_APP.asBoolean(payload && (payload.isMultiDay || payload.IsMultiDay));
    if (!teacherEmail && teacherName) {
      teacherEmail = this.buildTeacherSyntheticEmail_(teacherName);
    }
    if (isDailyServiceOutOfClass) {
      hourlyPeriods = {};
    }
    return {
      absenceId: ROOMS_APP.normalizeString(payload && (payload.absenceId || payload.AbsenceId)) || this.buildAbsenceId_(),
      teacherEmail: teacherEmail,
      teacherName: teacherName,
      absenceMode: this.normalizeAbsenceMode_(payload && (payload.absenceMode || payload.AbsenceMode)),
      absenceType: absenceType,
      serviceOutOfClassScope: serviceOutOfClassScope,
      startDate: ROOMS_APP.toIsoDate(payload && (payload.startDate || payload.StartDate || payload.date || payload.Date)),
      endDate: (absenceType === this.ABSENCE_TYPES_.DAILY || isDailyServiceOutOfClass) && isMultiDay
        ? ROOMS_APP.toIsoDate(payload && (payload.endDate || payload.EndDate))
        : '',
      isMultiDay: isMultiDay,
      hourlyPeriods: Object.keys(hourlyPeriods).sort(function (left, right) {
        return Number(left) - Number(right);
      }),
      recoveryRequired: absenceType === this.ABSENCE_TYPES_.HOURLY_PERMISSION && ROOMS_APP.asBoolean(payload && (payload.recoveryRequired || payload.RecoveryRequired)),
      notes: ROOMS_APP.normalizeString(payload && (payload.notes || payload.Notes))
    };
  },

  validateAbsenceCandidate_: function (candidate) {
    var periodMap;
    if (!candidate.teacherEmail || !candidate.teacherName) {
      throw new Error('Selezionare un docente valido.');
    }
    if (!candidate.startDate) {
      throw new Error('Inserire una data valida.');
    }
    if (candidate.absenceType === this.ABSENCE_TYPES_.DAILY || this.isDailyServiceOutOfClassAbsence_(candidate)) {
      if (candidate.isMultiDay) {
        if (!candidate.endDate) {
          throw new Error('Inserire la data finale dell\'assenza.');
        }
        if (candidate.endDate < candidate.startDate) {
          throw new Error('La data finale deve essere successiva o uguale alla data iniziale.');
        }
      }
      return;
    }
    if (!candidate.hourlyPeriods.length) {
      throw new Error(candidate.absenceType === this.ABSENCE_TYPES_.SERVICE_OUT_OF_CLASS
        ? 'Selezionare almeno un\'ora di servizio fuori aula.'
        : 'Selezionare almeno un\'ora di permesso.');
    }
    periodMap = ROOMS_APP.Timetable.getPeriodTimeMap();
    candidate.hourlyPeriods.forEach(function (period) {
      if (!periodMap[ROOMS_APP.normalizeString(period)]) {
        throw new Error('Le ore selezionate non sono valide.');
      }
    });
  },

  readAbsenceRow_: function (row) {
    var teacherName = ROOMS_APP.normalizeString(row && row.TeacherName);
    var absenceType = this.normalizeAbsenceType_(row && row.AbsenceType);
    var rawServiceOutOfClassScope = ROOMS_APP.normalizeString(row && row.ServiceOutOfClassScope);
    var serviceOutOfClassScope;
    var hourlyPeriods = ROOMS_APP.parseJson(row && row.HourlyPeriodsJson, []);
    var matchedPeriods = ROOMS_APP.parseJson(row && row.MatchedPeriodsJson, []);
    var effectiveWorkload = ROOMS_APP.parseJson(row && row.EffectiveWorkloadJson, null);
    var derivedVersion = ROOMS_APP.normalizeString(row && row.DerivedVersion);
    var derivedAtISO = ROOMS_APP.normalizeString(row && row.DerivedAtISO);
    var hasDerivedWorkload = Boolean(derivedVersion || derivedAtISO || row && row.EffectiveWorkloadJson || row && row.MatchedPeriodsJson);
    if (!Array.isArray(hourlyPeriods)) {
      hourlyPeriods = [];
    }
    hourlyPeriods = hourlyPeriods.map(function (entry) {
      return ROOMS_APP.normalizeString(entry);
    }).filter(function (entry) {
      return Boolean(entry);
    }).sort(function (left, right) {
      return Number(left) - Number(right);
    });
    if (!Array.isArray(matchedPeriods)) {
      matchedPeriods = [];
    }
    matchedPeriods = matchedPeriods.map(function (entry) {
      return ROOMS_APP.normalizeString(entry);
    }).filter(function (entry) {
      return Boolean(entry);
    }).sort(function (left, right) {
      return Number(left) - Number(right);
    });
    if (!effectiveWorkload || typeof effectiveWorkload !== 'object') {
      effectiveWorkload = null;
    }
    serviceOutOfClassScope = absenceType === this.ABSENCE_TYPES_.SERVICE_OUT_OF_CLASS
      ? (rawServiceOutOfClassScope ? this.normalizeServiceOutOfClassScope_(rawServiceOutOfClassScope) : (hourlyPeriods.length ? 'HOURLY' : 'DAILY'))
      : '';
    return {
      absenceId: ROOMS_APP.normalizeString(row && row.AbsenceId),
      teacherEmail: this.normalizeTeacherEmail_(row && row.TeacherEmail) || this.buildTeacherSyntheticEmail_(teacherName),
      teacherName: teacherName,
      absenceMode: this.normalizeAbsenceMode_(row && row.AbsenceMode),
      absenceType: absenceType,
      serviceOutOfClassScope: serviceOutOfClassScope,
      startDate: ROOMS_APP.toIsoDate(row && row.StartDate),
      endDate: ROOMS_APP.toIsoDate(row && row.EndDate),
      hourlyPeriods: hourlyPeriods,
      matchedPeriods: matchedPeriods,
      effectiveWorkload: effectiveWorkload,
      hasDerivedWorkload: hasDerivedWorkload,
      hasActionableWorkload: hasDerivedWorkload ? ROOMS_APP.asBoolean(row && row.HasActionableWorkload) : Boolean(matchedPeriods.length),
      actionableHoursCount: Number(row && row.ActionableHoursCount || 0),
      derivedAtISO: derivedAtISO,
      derivedVersion: derivedVersion,
      recoveryRequired: ROOMS_APP.asBoolean(row && row.RecoveryRequired),
      notes: ROOMS_APP.normalizeString(row && row.Notes),
      status: ROOMS_APP.normalizeString(row && row.Status).toUpperCase() || 'ACTIVE',
      enabled: !Object.prototype.hasOwnProperty.call(row || {}, 'Enabled') || ROOMS_APP.asBoolean(row && row.Enabled),
      createdAtISO: ROOMS_APP.normalizeString(row && row.CreatedAtISO),
      createdBy: this.normalizeTeacherEmail_(row && row.CreatedBy),
      updatedAtISO: ROOMS_APP.normalizeString(row && row.UpdatedAtISO),
      updatedBy: this.normalizeTeacherEmail_(row && row.UpdatedBy)
    };
  },

  listAbsenceRows_: function () {
    return ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.REPL_ABSENCES).map(function (row) {
      return ROOMS_APP.Replacements.readAbsenceRow_(row);
    }).filter(function (row) {
      return Boolean(row.absenceId && row.teacherEmail && row.startDate && row.enabled && row.status !== 'DELETED');
    }).sort(function (left, right) {
      if (left.startDate !== right.startDate) {
        return left.startDate.localeCompare(right.startDate);
      }
      if (left.teacherName !== right.teacherName) {
        return left.teacherName.localeCompare(right.teacherName);
      }
      return left.absenceId.localeCompare(right.absenceId);
    });
  },

  listActiveAbsenceRowsForDate_: function (targetDate) {
    var self = this;
    var dateKey = ROOMS_APP.toIsoDate(targetDate);
    return this.listAbsenceRows_().filter(function (row) {
      var endDate;
      if (!row || !row.startDate) {
        return false;
      }
      if (row.absenceType === self.ABSENCE_TYPES_.HOURLY_PERMISSION || self.isHourlyServiceOutOfClassAbsence_(row)) {
        return row.startDate === dateKey;
      }
      endDate = row.endDate || row.startDate;
      return Boolean(row.startDate <= dateKey && endDate >= dateKey);
    });
  },

  resolveAbsenceTeacherIdentityForDate_: function (teacherEmail, teacherName, activeLongAssignmentMap) {
    var normalizedEmail = this.normalizeTeacherEmail_(teacherEmail);
    var normalizedName = ROOMS_APP.normalizeString(teacherName);
    var longAssignment = normalizedEmail && activeLongAssignmentMap && activeLongAssignmentMap[normalizedEmail]
      ? activeLongAssignmentMap[normalizedEmail]
      : null;
    if (longAssignment) {
      return {
        teacherEmail: longAssignment.replacementTeacherEmail,
        teacherName: longAssignment.replacementTeacherDisplayName || normalizedName || longAssignment.replacementTeacherEmail
      };
    }
    return {
      teacherEmail: normalizedEmail || this.buildTeacherSyntheticEmail_(normalizedName),
      teacherName: normalizedName || normalizedEmail
    };
  },

  buildAbsenceRegistryDayState_: function (targetDate, absenceRows, activeLongAssignmentMap) {
    var self = this;
    var teacherRowsByKey = {};
    var hourlyRowsByKey = {};
    var controlledTeacherMap = {};
    var controlledHourlyMap = {};
    var serviceTeacherMap = {};
    var serviceHourlyMap = {};
    var sourceRows = [];

    (absenceRows || []).forEach(function (row) {
      var identity = self.resolveAbsenceTeacherIdentityForDate_(row.teacherEmail, row.teacherName, activeLongAssignmentMap);
      var teacherEmail = identity.teacherEmail;
      var teacherName = identity.teacherName;
      if (!teacherEmail) {
        return;
      }
      sourceRows.push({
        UpdatedAtISO: row.updatedAtISO || row.createdAtISO || '',
        AbsenceId: row.absenceId,
        TeacherEmail: teacherEmail,
        TeacherName: teacherName,
        AbsenceType: row.absenceType,
        ServiceOutOfClassScope: row.serviceOutOfClassScope,
        StartDate: row.startDate,
        EndDate: row.endDate,
        HourlyPeriods: (row.hourlyPeriods || []).slice(),
        MatchedPeriods: (row.matchedPeriods || []).slice(),
        HasActionableWorkload: Boolean(row.hasActionableWorkload),
        ActionableHoursCount: Number(row.actionableHoursCount || 0),
        DerivedAtISO: row.derivedAtISO || '',
        DerivedVersion: row.derivedVersion || ''
      });
      if (row.absenceType === self.ABSENCE_TYPES_.DAILY || self.isDailyServiceOutOfClassAbsence_(row)) {
        controlledTeacherMap[teacherEmail] = true;
        if (self.isDailyServiceOutOfClassAbsence_(row)) {
          serviceTeacherMap[teacherEmail] = true;
        }
        teacherRowsByKey[teacherEmail] = {
          Date: targetDate,
          TeacherEmail: teacherEmail,
          TeacherName: teacherName,
          Absent: self.isDailyServiceOutOfClassAbsence_(row) ? 'FALSE' : 'TRUE',
          Accompanist: 'FALSE',
          AccompaniedClasses: '',
          Notes: ROOMS_APP.normalizeString(row.notes),
          UpdatedAtISO: row.updatedAtISO || row.createdAtISO || '',
          UpdatedBy: row.updatedBy || row.createdBy || ''
        };
        return;
      }
      if (row.hasDerivedWorkload) {
        return;
      }
      (row.hourlyPeriods || []).forEach(function (period) {
        var controlledKey = self.buildHourlyAbsenceKey_(teacherEmail, period);
        controlledHourlyMap[controlledKey] = true;
      });
      (row.hasDerivedWorkload
        ? ((row.effectiveWorkload && Array.isArray(row.effectiveWorkload.items) ? row.effectiveWorkload.items : []).filter(function (item) {
          return Boolean(item && item.actionable && ROOMS_APP.normalizeString(item.period));
        }))
        : (row.hourlyPeriods || []).map(function (period) {
          return { period: period };
        })
      ).forEach(function (item) {
        var period = ROOMS_APP.normalizeString(item && item.period);
        var key;
        if (!period) {
          return;
        }
        key = self.buildHourlyAbsenceKey_(teacherEmail, period);
        controlledHourlyMap[key] = true;
        if (self.isHourlyServiceOutOfClassAbsence_(row)) {
          serviceHourlyMap[key] = true;
        }
        hourlyRowsByKey[key] = {
          Date: targetDate,
          TeacherEmail: teacherEmail,
          TeacherName: teacherName,
          Period: period,
          Reason: self.isHourlyServiceOutOfClassAbsence_(row) ? self.SERVICE_OUT_OF_CLASS_.GENERAL : '',
          ClassCode: ROOMS_APP.normalizeString(item && item.classCode).toUpperCase(),
          Label: ROOMS_APP.normalizeString(item && item.label),
          SlotType: ROOMS_APP.normalizeString(item && item.slotType),
          StartTime: ROOMS_APP.normalizeString(item && item.startTime),
          EndTime: ROOMS_APP.normalizeString(item && item.endTime),
          DerivedActionable: row.hasDerivedWorkload ? 'TRUE' : '',
          RecoveryRequired: row.absenceType === self.ABSENCE_TYPES_.HOURLY_PERMISSION && row.recoveryRequired ? 'TRUE' : 'FALSE',
          RecoveryStatus: '',
          RecoveredOnDate: '',
          RecoveredByAssignmentKey: '',
          Notes: ROOMS_APP.normalizeString(row.notes),
          UpdatedAtISO: row.updatedAtISO || row.createdAtISO || '',
          UpdatedBy: row.updatedBy || row.createdBy || ''
        };
      });
    });

    return {
      hasEntries: Boolean(sourceRows.length),
      sourceRows: sourceRows,
      teacherRows: Object.keys(teacherRowsByKey).map(function (key) {
        return teacherRowsByKey[key];
      }),
      hourlyRows: Object.keys(hourlyRowsByKey).map(function (key) {
        return hourlyRowsByKey[key];
      }),
      controlledTeacherMap: controlledTeacherMap,
      controlledHourlyMap: controlledHourlyMap,
      serviceTeacherMap: serviceTeacherMap,
      serviceHourlyMap: serviceHourlyMap,
      dailyTeacherCount: Object.keys(teacherRowsByKey).length,
      hourlyEntryCount: Object.keys(hourlyRowsByKey).length,
      sourceLabel: sourceRows.length
        ? 'Assenze ed eventi caricati da REPL_ABSENCES'
        : 'Assenze giornaliere senza registro salvato'
    };
  },

  buildControlledTeacherRowKey_: function (row, activeLongAssignmentMap) {
    var identity = this.resolveAbsenceTeacherIdentityForDate_(
      row && row.TeacherEmail,
      row && row.TeacherName,
      activeLongAssignmentMap
    );
    return identity.teacherEmail;
  },

  mergeTeacherRowsWithRegistry_: function (legacyRows, registryState, activeLongAssignmentMap) {
    var registryRows = registryState && registryState.teacherRows ? registryState.teacherRows : [];
    var controlledTeacherMap = registryState && registryState.controlledTeacherMap ? registryState.controlledTeacherMap : {};
    return (legacyRows || []).filter(function (row) {
      return !controlledTeacherMap[ROOMS_APP.Replacements.buildControlledTeacherRowKey_(row, activeLongAssignmentMap)];
    }).concat(registryRows);
  },

  mergeHourlyAbsenceRowsWithRegistry_: function (legacyRows, registryState, activeLongAssignmentMap) {
    var registryRows = registryState && registryState.hourlyRows ? registryState.hourlyRows : [];
    var controlledTeacherMap = registryState && registryState.controlledTeacherMap ? registryState.controlledTeacherMap : {};
    var controlledHourlyMap = registryState && registryState.controlledHourlyMap ? registryState.controlledHourlyMap : {};
    var self = this;
    var acceptedKeyMap = {};
    var acceptedRows = (legacyRows || []).filter(function (row) {
      var identity = self.resolveAbsenceTeacherIdentityForDate_(row && row.TeacherEmail, row && row.TeacherName, activeLongAssignmentMap);
      var key = self.buildHourlyAbsenceKey_(identity.teacherEmail, row && row.Period);
      if (ROOMS_APP.normalizeString(row && row.SourceAbsenceId)) {
        acceptedKeyMap[key] = true;
        return true;
      }
      if (controlledTeacherMap[identity.teacherEmail]) {
        return false;
      }
      if (controlledHourlyMap[key]) {
        return false;
      }
      acceptedKeyMap[key] = true;
      return true;
    });
    return acceptedRows.concat((registryRows || []).filter(function (row) {
      var identity = self.resolveAbsenceTeacherIdentityForDate_(row && row.TeacherEmail, row && row.TeacherName, activeLongAssignmentMap);
      return !acceptedKeyMap[self.buildHourlyAbsenceKey_(identity.teacherEmail, row && row.Period)];
    }));
  },

  getTeacherServicePeriodsForDate_: function (teacherEmail, dateString) {
    var teacherMap = this.indexByTeacherEmail_(this.buildTeacherDayTeachers_(dateString));
    var teacher = teacherMap[this.normalizeTeacherEmail_(teacherEmail)];
    if (!teacher) {
      return [];
    }
    return Object.keys(teacher.periods || {}).sort(function (left, right) {
      return Number(left) - Number(right);
    }).map(function (period) {
      var slot = teacher.periods[period] || {};
      var assignmentClassCode = ROOMS_APP.Replacements.getTeacherSlotAssignmentClassCode_(slot);
      if (!ROOMS_APP.Replacements.isCoverableTeacherSlot_(teacher, slot)) {
        return null;
      }
      return {
        period: String(period || ''),
        periodLabel: String(period || '') + 'a ora',
        classCode: assignmentClassCode,
        label: String(period || '') + 'a ora' + (assignmentClassCode ? (' · ' + (slot.label || assignmentClassCode)) : ''),
        startTime: slot.startTime || '',
        endTime: slot.endTime || ''
      };
    }).filter(function (entry) {
      return Boolean(entry);
    });
  },

  normalizeTripType_: function (value) {
    var normalized = ROOMS_APP.normalizeString(value).toUpperCase();
    if (normalized === this.TRIP_TYPES_.MULTI_DAY || normalized === this.TRIP_TYPES_.HOURLY) {
      return normalized;
    }
    return this.TRIP_TYPES_.DAILY;
  },

  normalizeEducationalTripPayload_: function (payload) {
    var tripType = this.normalizeTripType_(payload && payload.tripType);
    var startDate = ROOMS_APP.toIsoDate(payload && (payload.startDate || payload.date));
    var endDate = ROOMS_APP.toIsoDate(payload && payload.endDate);
    var teacherMap = {};
    (Array.isArray(payload && payload.teachers) ? payload.teachers : []).forEach(function (entry) {
      var teacherEmail = ROOMS_APP.Replacements.normalizeTeacherEmail_(entry && entry.teacherEmail);
      var teacherName = ROOMS_APP.normalizeString(entry && entry.teacherName);
      if (!teacherEmail && teacherName) {
        teacherEmail = ROOMS_APP.Replacements.buildTeacherSyntheticEmail_(teacherName);
      }
      if (!teacherEmail) {
        return;
      }
      teacherMap[teacherEmail] = {
        teacherEmail: teacherEmail,
        teacherName: teacherName || teacherEmail,
        role: ROOMS_APP.normalizeString(entry && entry.role).toUpperCase() || ROOMS_APP.Replacements.TRIP_ROLE_,
        notes: ROOMS_APP.normalizeString(entry && entry.notes)
      };
    });
    if (tripType === this.TRIP_TYPES_.MULTI_DAY && !endDate) {
      endDate = startDate;
    }
    if (tripType !== this.TRIP_TYPES_.MULTI_DAY) {
      endDate = startDate;
    }
    return {
      tripId: ROOMS_APP.normalizeString(payload && payload.tripId),
      tripType: tripType,
      classCode: ROOMS_APP.normalizeString(payload && payload.classCode).toUpperCase(),
      title: ROOMS_APP.normalizeString(payload && payload.title),
      startDate: startDate,
      endDate: endDate,
      startTime: tripType === this.TRIP_TYPES_.HOURLY ? ROOMS_APP.toTimeString(payload && payload.startTime) : '',
      endTime: tripType === this.TRIP_TYPES_.HOURLY ? ROOMS_APP.toTimeString(payload && payload.endTime) : '',
      notes: ROOMS_APP.normalizeString(payload && payload.notes),
      enabled: payload && Object.prototype.hasOwnProperty.call(payload, 'enabled') ? ROOMS_APP.asBoolean(payload.enabled) : true,
      teachers: Object.keys(teacherMap).map(function (key) {
        return teacherMap[key];
      })
    };
  },

  validateEducationalTrip_: function (candidate) {
    if (!candidate.classCode) {
      throw new Error('Selezionare una classe.');
    }
    if (!candidate.startDate) {
      throw new Error('Inserire una data valida per l\'uscita didattica.');
    }
    if (candidate.tripType === this.TRIP_TYPES_.MULTI_DAY && (!candidate.endDate || candidate.endDate < candidate.startDate)) {
      throw new Error('Inserire un intervallo date valido per il viaggio di istruzione.');
    }
    if (candidate.tripType === this.TRIP_TYPES_.HOURLY) {
      if (!candidate.startTime || !candidate.endTime) {
        throw new Error('Inserire orario di inizio e fine per l\'uscita oraria.');
      }
      if (candidate.startTime >= candidate.endTime) {
        throw new Error('L\'orario finale deve essere successivo a quello iniziale.');
      }
    }
  },

  iterateEducationalTripDates_: function (trip, callback) {
    var startDate = new Date(String(trip && trip.startDate || '') + 'T12:00:00');
    var endDate = new Date(String((trip && trip.endDate) || (trip && trip.startDate) || '') + 'T12:00:00');
    var cursor;
    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime()) || typeof callback !== 'function') {
      return;
    }
    cursor = new Date(startDate.getTime());
    while (cursor.getTime() <= endDate.getTime()) {
      callback(ROOMS_APP.toIsoDate(cursor));
      cursor.setDate(cursor.getDate() + 1);
    }
  },

  validateEducationalTripRegistry_: function (trips) {
    var assignedTeachersByDate = {};
    var self = this;
    (trips || []).forEach(function (trip) {
      self.validateEducationalTrip_(trip);
    });
    (trips || []).forEach(function (trip) {
      if (!trip || !trip.enabled) {
        return;
      }
      self.iterateEducationalTripDates_(trip, function (dateKey) {
        assignedTeachersByDate[dateKey] = assignedTeachersByDate[dateKey] || {};
        (trip.teachers || []).forEach(function (teacher) {
          var teacherEmail = self.normalizeTeacherEmail_(teacher && teacher.teacherEmail);
          var teacherName = ROOMS_APP.normalizeString(teacher && teacher.teacherName) || teacherEmail;
          if (!teacherEmail) {
            return;
          }
          if (assignedTeachersByDate[dateKey][teacherEmail] && assignedTeachersByDate[dateKey][teacherEmail] !== trip.tripId) {
            throw new Error('Il docente ' + teacherName + ' risulta gia assegnato a un\'altra uscita il ' + dateKey + '.');
          }
          assignedTeachersByDate[dateKey][teacherEmail] = trip.tripId;
        });
      });
    });
  },

  buildEducationalTripRegistryRows_: function (trips, updatedBy, nowIso) {
    var self = this;
    var tripRows = [];
    var tripTeacherRows = [];
    (trips || []).forEach(function (trip) {
      tripRows.push({
        TripId: trip.tripId,
        TripType: trip.tripType,
        ClassCode: trip.classCode,
        Title: trip.title,
        StartDate: trip.startDate,
        EndDate: trip.endDate,
        StartTime: trip.startTime,
        EndTime: trip.endTime,
        Notes: trip.notes,
        Enabled: trip.enabled ? 'TRUE' : 'FALSE',
        UpdatedAtISO: nowIso,
        UpdatedBy: updatedBy
      });
      (trip.teachers || []).forEach(function (teacher) {
        tripTeacherRows.push({
          TripId: trip.tripId,
          TeacherEmail: self.normalizeTeacherEmail_(teacher && teacher.teacherEmail),
          TeacherName: ROOMS_APP.normalizeString(teacher && teacher.teacherName),
          Role: ROOMS_APP.normalizeString(teacher && teacher.role).toUpperCase() || self.TRIP_ROLE_,
          Notes: ROOMS_APP.normalizeString(teacher && teacher.notes),
          UpdatedAtISO: nowIso,
          UpdatedBy: updatedBy
        });
      });
    });
    return {
      tripRows: tripRows,
      tripTeacherRows: tripTeacherRows
    };
  },

  readTripRow_: function (row) {
    return {
      tripId: ROOMS_APP.normalizeString(row && row.TripId),
      tripType: this.normalizeTripType_(row && row.TripType),
      classCode: ROOMS_APP.normalizeString(row && row.ClassCode).toUpperCase(),
      title: ROOMS_APP.normalizeString(row && row.Title),
      startDate: ROOMS_APP.toIsoDate(row && row.StartDate),
      endDate: ROOMS_APP.toIsoDate((row && row.EndDate) || (row && row.StartDate)),
      startTime: ROOMS_APP.toTimeString(row && row.StartTime),
      endTime: ROOMS_APP.toTimeString(row && row.EndTime),
      notes: ROOMS_APP.normalizeString(row && row.Notes),
      enabled: ROOMS_APP.asBoolean(row && row.Enabled),
      updatedAtISO: ROOMS_APP.normalizeString(row && row.UpdatedAtISO),
      updatedBy: this.normalizeTeacherEmail_(row && row.UpdatedBy)
    };
  },

  readTripTeacherRow_: function (row) {
    var teacherName = ROOMS_APP.normalizeString(row && row.TeacherName);
    return {
      tripId: ROOMS_APP.normalizeString(row && row.TripId),
      teacherEmail: this.normalizeTeacherEmail_(row && row.TeacherEmail) || this.buildTeacherSyntheticEmail_(teacherName),
      teacherName: teacherName,
      role: ROOMS_APP.normalizeString(row && row.Role).toUpperCase() || this.TRIP_ROLE_,
      notes: ROOMS_APP.normalizeString(row && row.Notes),
      updatedAtISO: ROOMS_APP.normalizeString(row && row.UpdatedAtISO),
      updatedBy: this.normalizeTeacherEmail_(row && row.UpdatedBy)
    };
  },

  listEducationalTrips_: function () {
    var teacherRowsByTripId = {};
    ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.REPL_FIELD_TRIP_TEACHERS).forEach(function (row) {
      var teacherRow = ROOMS_APP.Replacements.readTripTeacherRow_(row);
      if (!teacherRow.tripId || !teacherRow.teacherEmail) {
        return;
      }
      teacherRowsByTripId[teacherRow.tripId] = teacherRowsByTripId[teacherRow.tripId] || [];
      teacherRowsByTripId[teacherRow.tripId].push(teacherRow);
    });
    return ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.REPL_FIELD_TRIPS).map(function (row) {
      var trip = ROOMS_APP.Replacements.readTripRow_(row);
      trip.teachers = (teacherRowsByTripId[trip.tripId] || []).slice().sort(function (left, right) {
        return left.teacherName.localeCompare(right.teacherName);
      });
      return trip;
    }).filter(function (trip) {
      return Boolean(trip.tripId && trip.classCode && trip.startDate);
    }).sort(function (left, right) {
      if (left.startDate !== right.startDate) {
        return left.startDate.localeCompare(right.startDate);
      }
      if (left.classCode !== right.classCode) {
        return left.classCode.localeCompare(right.classCode);
      }
      return left.tripId.localeCompare(right.tripId);
    });
  },

  listTimetableClassOptions_: function () {
    return this.getTripTeacherOptionsRegistry_().classOptions.slice();
  },

  listTimetableTeachersByClass_: function () {
    var registry = this.getTripTeacherOptionsRegistry_();
    var cloned = {};
    Object.keys(registry.byClass || {}).forEach(function (classCode) {
      cloned[classCode] = (registry.byClass[classCode] || []).map(function (entry) {
        return {
          teacherEmail: entry.teacherEmail,
          teacherName: entry.teacherName
        };
      });
    });
    return cloned;
  },

  getTripTeacherOptionsRegistry_: function () {
    var baseRegistry = this.getDocentiTimetableDerivedData_().tripRegistry;
    var output = {};
    var self = this;
    Object.keys(baseRegistry.byClass || {}).forEach(function (classCode) {
      output[classCode] = (baseRegistry.byClass[classCode] || []).map(function (entry) {
        return {
          teacherEmail: entry.teacherEmail,
          teacherName: entry.teacherName
        };
      });
    });

    this.listSupportTeacherEntries_().forEach(function (entry) {
      if (!entry.teacherEmail) {
        return;
      }
      entry.classCodes.forEach(function (classCode) {
        self.addTripTeacherOption_(output, classCode, {
          teacherEmail: entry.teacherEmail,
          teacherName: entry.teacherName
        });
      });
    });

    Object.keys(output).forEach(function (classCode) {
      output[classCode] = output[classCode].sort(function (left, right) {
        return left.teacherName.localeCompare(right.teacherName);
      });
    });
    return {
      byClass: output,
      classOptions: Object.keys(output).sort()
    };
  },

  listSupportTeacherEntries_: function () {
    var self = this;
    return ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.TIMETABLE_DOCENTI_SOSTEGNO).map(function (row) {
      return self.normalizeSupportTeacherRow_(row);
    }).filter(function (entry) {
      return Boolean(entry && entry.teacherEmail);
    });
  },

  normalizeSupportTeacherRow_: function (row) {
    var surname = ROOMS_APP.normalizeString(row && (row.Cognome || row.cognome || row.Surname || row.surname));
    var firstName = ROOMS_APP.normalizeString(row && (row.Nome || row.nome || row.Name || row.name));
    var teacherName = ROOMS_APP.normalizeString(
      (surname || firstName)
        ? (surname + ' ' + firstName)
        : (row && (row.TeacherName || row.teacherName || row.Docente || row.docente))
    ).replace(/\s+/g, ' ').trim();
    var teacherEmail = this.normalizeTeacherEmail_(row && (row.TeacherEmail || row.teacherEmail || row.Email || row.email)) ||
      this.buildTeacherSyntheticEmail_(teacherName);
    if (!teacherEmail) {
      return null;
    }
    return {
      teacherEmail: teacherEmail,
      teacherName: teacherName || teacherEmail,
      classCodes: this.parseSupportTeacherClassCodes_(row && (row.classi || row.Classi || row.Classes || row.classes || row.ClassCodes || row.classCodes))
    };
  },

  addTripTeacherOption_: function (output, classCode, teacher) {
    var normalizedClassCode = ROOMS_APP.normalizeString(classCode).toUpperCase();
    var teacherEmail = this.normalizeTeacherEmail_(teacher && teacher.teacherEmail);
    if (!normalizedClassCode || !teacherEmail) {
      return;
    }
    output[normalizedClassCode] = output[normalizedClassCode] || [];
    if (output[normalizedClassCode].some(function (entry) {
      return entry.teacherEmail === teacherEmail;
    })) {
      return;
    }
    output[normalizedClassCode].push({
      teacherEmail: teacherEmail,
      teacherName: ROOMS_APP.normalizeString(teacher && teacher.teacherName) || teacherEmail
    });
  },

  parseSupportTeacherClassCodes_: function (value) {
    var seen = {};
    return ROOMS_APP.normalizeString(value).toUpperCase().split(/[^A-Z0-9]+/).map(function (token) {
      return ROOMS_APP.Timetable.extractClassCode_(token);
    }).filter(function (classCode) {
      if (!classCode || seen[classCode]) {
        return false;
      }
      seen[classCode] = true;
      return true;
    });
  },

  getAllPeriodKeys_: function () {
    return Object.keys(ROOMS_APP.Timetable.getPeriodTimeMap()).sort(function (left, right) {
      return Number(left) - Number(right);
    });
  },

  getTripPeriodsForDate_: function (trip, targetDate) {
    var periodMap = ROOMS_APP.Timetable.getPeriodTimeMap();
    var allPeriods = this.getAllPeriodKeys_();
    if (!trip || !trip.enabled || !targetDate) {
      return [];
    }
    if (trip.tripType !== this.TRIP_TYPES_.HOURLY) {
      return allPeriods;
    }
    return allPeriods.filter(function (period) {
      var meta = periodMap[period] || {};
      return Boolean(meta.startTime && meta.endTime && trip.startTime < meta.endTime && trip.endTime > meta.startTime);
    });
  },

  isTripActiveOnDate_: function (trip, targetDate) {
    if (!trip || !trip.enabled || !targetDate) {
      return false;
    }
    if (trip.tripType === this.TRIP_TYPES_.MULTI_DAY) {
      return Boolean(trip.startDate && trip.endDate && trip.startDate <= targetDate && trip.endDate >= targetDate);
    }
    return trip.startDate === targetDate;
  },

  buildTripDayState_: function (targetDate, trips) {
    var state = {
      classOutSet: {},
      classOutPeriodsByClass: {},
      teacherTripPeriodsByTeacher: {},
      teacherTripClassesByTeacher: {},
      sourceRows: [],
      sourceTeacherRows: [],
      activeTrips: []
    };
    var self = this;
    (trips || []).forEach(function (trip) {
      var periods;
      if (!self.isTripActiveOnDate_(trip, targetDate)) {
        return;
      }
      periods = self.getTripPeriodsForDate_(trip, targetDate);
      if (!trip.classCode || !periods.length) {
        return;
      }
      state.activeTrips.push(trip);
      state.sourceRows.push({
        UpdatedAtISO: trip.updatedAtISO || '',
        TripId: trip.tripId
      });
      state.classOutSet[trip.classCode] = true;
      state.classOutPeriodsByClass[trip.classCode] = state.classOutPeriodsByClass[trip.classCode] || {};
      periods.forEach(function (period) {
        state.classOutPeriodsByClass[trip.classCode][period] = true;
      });
      (trip.teachers || []).forEach(function (teacher) {
        var teacherEmail = self.normalizeTeacherEmail_(teacher.teacherEmail);
        if (!teacherEmail) {
          return;
        }
        state.sourceTeacherRows.push({
          UpdatedAtISO: teacher.updatedAtISO || trip.updatedAtISO || '',
          TripId: trip.tripId,
          TeacherEmail: teacherEmail
        });
        state.teacherTripPeriodsByTeacher[teacherEmail] = state.teacherTripPeriodsByTeacher[teacherEmail] || {};
        periods.forEach(function (period) {
          state.teacherTripPeriodsByTeacher[teacherEmail][period] = true;
        });
        state.teacherTripClassesByTeacher[teacherEmail] = self.uniqueStrings_(
          (state.teacherTripClassesByTeacher[teacherEmail] || []).concat([trip.classCode])
        );
      });
    });
    return state;
  },

  buildLegacyOutingState_: function (targetDate, classRows, teacherRows) {
    var state = {
      classOutSet: {},
      classOutPeriodsByClass: {},
      teacherTripPeriodsByTeacher: {},
      teacherTripClassesByTeacher: {},
      sourceRows: (classRows || []).slice(),
      sourceTeacherRows: (teacherRows || []).slice(),
      activeTrips: []
    };
    var allPeriods = this.getAllPeriodKeys_();
    (classRows || []).forEach(function (row) {
      var classCode = ROOMS_APP.normalizeString(row && row.ClassCode).toUpperCase();
      if (!ROOMS_APP.asBoolean(row && row.IsOut) || !classCode) {
        return;
      }
      state.classOutSet[classCode] = true;
      state.classOutPeriodsByClass[classCode] = state.classOutPeriodsByClass[classCode] || {};
      allPeriods.forEach(function (period) {
        state.classOutPeriodsByClass[classCode][period] = true;
      });
    });
    (teacherRows || []).forEach(function (row) {
      var teacherEmail = ROOMS_APP.Replacements.normalizeTeacherEmail_(row && row.TeacherEmail) || ROOMS_APP.Replacements.buildTeacherSyntheticEmail_(row && row.TeacherName);
      var classes = ROOMS_APP.Replacements.parsePipeList_(row && row.AccompaniedClasses);
      if (!ROOMS_APP.asBoolean(row && row.Accompanist) || !teacherEmail || !classes.length) {
        return;
      }
      state.teacherTripPeriodsByTeacher[teacherEmail] = state.teacherTripPeriodsByTeacher[teacherEmail] || {};
      allPeriods.forEach(function (period) {
        state.teacherTripPeriodsByTeacher[teacherEmail][period] = true;
      });
      state.teacherTripClassesByTeacher[teacherEmail] = ROOMS_APP.Replacements.uniqueStrings_(classes);
    });
    return state;
  },

  mergeTripDayStates_: function (left, right) {
    var merged = {
      classOutSet: {},
      classOutPeriodsByClass: {},
      teacherTripPeriodsByTeacher: {},
      teacherTripClassesByTeacher: {},
      sourceRows: [],
      sourceTeacherRows: [],
      activeTrips: []
    };
    var self = this;
    [left || {}, right || {}].forEach(function (state) {
      Object.keys(state.classOutSet || {}).forEach(function (classCode) {
        merged.classOutSet[classCode] = true;
      });
      Object.keys(state.classOutPeriodsByClass || {}).forEach(function (classCode) {
        merged.classOutPeriodsByClass[classCode] = merged.classOutPeriodsByClass[classCode] || {};
        Object.keys(state.classOutPeriodsByClass[classCode] || {}).forEach(function (period) {
          merged.classOutPeriodsByClass[classCode][period] = true;
        });
      });
      Object.keys(state.teacherTripPeriodsByTeacher || {}).forEach(function (teacherEmail) {
        merged.teacherTripPeriodsByTeacher[teacherEmail] = merged.teacherTripPeriodsByTeacher[teacherEmail] || {};
        Object.keys(state.teacherTripPeriodsByTeacher[teacherEmail] || {}).forEach(function (period) {
          merged.teacherTripPeriodsByTeacher[teacherEmail][period] = true;
        });
      });
      Object.keys(state.teacherTripClassesByTeacher || {}).forEach(function (teacherEmail) {
        merged.teacherTripClassesByTeacher[teacherEmail] = self.uniqueStrings_(
          (merged.teacherTripClassesByTeacher[teacherEmail] || []).concat(state.teacherTripClassesByTeacher[teacherEmail] || [])
        );
      });
      merged.sourceRows = merged.sourceRows.concat(state.sourceRows || []);
      merged.sourceTeacherRows = merged.sourceTeacherRows.concat(state.sourceTeacherRows || []);
      merged.activeTrips = merged.activeTrips.concat(state.activeTrips || []);
    });
    return merged;
  },

  isClassOutAtPeriodInState_: function (state, classCode, period) {
    var classPeriods = state && state.classOutPeriodsByClass
      ? state.classOutPeriodsByClass[ROOMS_APP.normalizeString(classCode).toUpperCase()]
      : null;
    return Boolean(classPeriods && classPeriods[ROOMS_APP.normalizeString(period)]);
  },

  isTeacherAccompanyingAtPeriodInState_: function (state, teacherEmail, period) {
    var periods = state && state.teacherTripPeriodsByTeacher
      ? state.teacherTripPeriodsByTeacher[this.normalizeTeacherEmail_(teacherEmail)]
      : null;
    return Boolean(periods && periods[ROOMS_APP.normalizeString(period)]);
  },

  isClassOutAtPeriod_: function (normalized, classCode, period) {
    return this.isClassOutAtPeriodInState_(normalized && normalized.tripDayState, classCode, period);
  },

  isTeacherAccompanyingAtPeriod_: function (normalized, teacherEmail, period) {
    return this.isTeacherAccompanyingAtPeriodInState_(normalized && normalized.tripDayState, teacherEmail, period);
  },

  uniqueStrings_: function (values) {
    var seen = {};
    var output = [];
    (values || []).forEach(function (entry) {
      var normalized = ROOMS_APP.normalizeString(entry).toUpperCase();
      if (!normalized || seen[normalized]) {
        return;
      }
      seen[normalized] = true;
      output.push(normalized);
    });
    return output;
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
