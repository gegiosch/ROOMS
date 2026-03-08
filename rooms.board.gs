var ROOMS_APP = ROOMS_APP || {};

ROOMS_APP.Board = {
  BRANCH_ORDER_: ['PA2', 'PA1', '2P_DX', '2P_SX', '1P_DX', '1P_SX', 'PT', 'LAB'],
  BRANCH_LABELS_: {
    PA2: 'PA2',
    PA1: 'PA1',
    '2P_DX': 'DX',
    '2P_SX': 'SX',
    '1P_DX': 'DX',
    '1P_SX': 'SX',
    PT: 'PT',
    LAB: 'LABORATORI'
  },
  BRANCH_PAGE_CAPACITY_: 12,

  getBoardViewModel: function () {
    ROOMS_APP.Schema.ensureAll();
    var now = new Date();
    var nowIso = ROOMS_APP.toIsoDateTime(now);
    var today = ROOMS_APP.toIsoDate(now);
    var currentTime = Utilities.formatDate(now, ROOMS_APP.getTimezone(), 'HH:mm');
    var resources = this.listResourcesForBoard_();
    var bookings = ROOMS_APP.Booking.listBookingsForDate(today);
    var byResource = {};
    var branchBuckets = this.createEmptyBranchBuckets_();

    bookings.forEach(function (booking) {
      byResource[booking.ResourceId] = byResource[booking.ResourceId] || [];
      byResource[booking.ResourceId].push(booking);
    });

    Object.keys(byResource).forEach(function (resourceId) {
      byResource[resourceId] = ROOMS_APP.sortBy(byResource[resourceId], ['StartTime', 'EndTime']);
    });

    resources.forEach(function (resource) {
      var roomBookings = byResource[resource.ResourceId] || [];
      var current = roomBookings.filter(function (booking) {
        return booking.StartTime <= currentTime && booking.EndTime > currentTime;
      })[0] || null;
      var next = roomBookings.filter(function (booking) {
        return booking.StartTime > currentTime;
      })[0] || null;
      var state = current ? 'OCCUPIED' : (next ? 'NEXT_OCCUPIED' : 'FREE');
      var branchKey = ROOMS_APP.Board.mapResourceBranchKey_(resource);

      branchBuckets[branchKey].push({
        resourceId: resource.ResourceId,
        displayName: resource.DisplayName,
        sortKey: resource.SortKey || '',
        layoutRow: Number(resource.LayoutRow || 0),
        layoutCol: Number(resource.LayoutCol || 0),
        state: state,
        currentLabel: current ? ROOMS_APP.Board.getBookingSurname_(current) : 'LIBERA',
        nextLabel: next ? ROOMS_APP.Board.getBookingSurname_(next) : 'LIBERA'
      });
    });

    this.BRANCH_ORDER_.forEach(function (branchKey) {
      branchBuckets[branchKey] = ROOMS_APP.sortBy(branchBuckets[branchKey], ['sortKey', 'layoutRow', 'layoutCol', 'displayName']);
    });

    var configuredPageCount = Math.max(1, ROOMS_APP.getNumberConfig('BOARD_PAGE_COUNT', 1));
    var requiredPageCount = this.getRequiredPageCount_(branchBuckets);
    var pageCount = Math.max(1, Math.min(configuredPageCount, requiredPageCount));
    var pages = this.buildPages_(branchBuckets, pageCount);

    return {
      generatedAtISO: nowIso,
      date: today,
      refreshSec: ROOMS_APP.getNumberConfig('BOARD_REFRESH_SEC', 60),
      rotationSec: ROOMS_APP.getNumberConfig('BOARD_ROTATION_SEC', 15),
      pageCount: pages.length,
      palette: ROOMS_APP.PALETTE,
      appName: ROOMS_APP.getConfigValue('APP_NAME', 'ROOMS'),
      schoolName: ROOMS_APP.getConfigValue('SCHOOL_NAME', 'IIS Alessandrini'),
      branchOrder: this.BRANCH_ORDER_.slice(),
      branchLabels: this.BRANCH_LABELS_,
      pages: pages
    };
  },

  createEmptyBranchBuckets_: function () {
    var buckets = {};
    this.BRANCH_ORDER_.forEach(function (branchKey) {
      buckets[branchKey] = [];
    });
    return buckets;
  },

  getRequiredPageCount_: function (branchBuckets) {
    var maxBranchLength = 0;
    this.BRANCH_ORDER_.forEach(function (branchKey) {
      maxBranchLength = Math.max(maxBranchLength, branchBuckets[branchKey].length);
    });
    return Math.max(1, Math.ceil(maxBranchLength / this.BRANCH_PAGE_CAPACITY_));
  },

  buildPages_: function (branchBuckets, pageCount) {
    var pages = [];
    var index;
    for (index = 0; index < pageCount; index += 1) {
      var pageBranches = {};
      this.BRANCH_ORDER_.forEach(function (branchKey) {
        var rooms = branchBuckets[branchKey];
        var chunkSize = Math.max(1, Math.ceil(rooms.length / pageCount));
        pageBranches[branchKey] = rooms.slice(index * chunkSize, (index + 1) * chunkSize);
      });
      pages.push({
        pageId: String(index + 1),
        title: 'Pagina ' + String(index + 1),
        branches: pageBranches
      });
    }
    return pages;
  },

  getBookingSurname_: function (booking) {
    var explicitSurname = ROOMS_APP.normalizeString(booking.BookerSurname);
    if (explicitSurname) {
      return explicitSurname;
    }

    var name = ROOMS_APP.normalizeString(booking.BookerName);
    if (name) {
      var parts = name.split(/\s+/);
      return parts[parts.length - 1];
    }

    var email = ROOMS_APP.normalizeString(booking.BookerEmail);
    if (email && email.indexOf('@') > 0) {
      var localPart = email.split('@')[0];
      var emailParts = localPart.split(/[._-]+/);
      return emailParts[emailParts.length - 1].toUpperCase();
    }

    return 'N/D';
  },

  mapResourceBranchKey_: function (resource) {
    var areaCode = this.normalizeKey_(resource.AreaCode);
    var floorCode = this.normalizeKey_(resource.FloorCode);
    var sideCode = this.normalizeKey_(resource.SideCode);
    var areaLabel = this.normalizeKey_(resource.AreaLabel);
    var floorLabel = this.normalizeKey_(resource.FloorLabel);
    var sideLabel = this.normalizeKey_(resource.SideLabel);

    if (areaCode === 'LAB' || areaLabel.indexOf('LABORATORI') >= 0 || sideLabel.indexOf('LABORATORI') >= 0) {
      return 'LAB';
    }

    if (areaCode === 'F2M' || (floorCode === 'F2' && sideCode === 'MEZZATO') || floorLabel.indexOf('SECONDO') >= 0 && floorLabel.indexOf('MEZZATO') >= 0) {
      return 'PA2';
    }

    if (areaCode === 'F1M' || (floorCode === 'F1' && sideCode === 'MEZZATO') || floorLabel.indexOf('PRIMO') >= 0 && floorLabel.indexOf('MEZZATO') >= 0) {
      return 'PA1';
    }

    if (areaCode === 'F2R' || (floorCode === 'F2' && (sideCode === 'RIGHT' || sideLabel.indexOf('DESTRO') >= 0))) {
      return '2P_DX';
    }

    if (areaCode === 'F2L' || (floorCode === 'F2' && (sideCode === 'LEFT' || sideLabel.indexOf('SINISTRO') >= 0))) {
      return '2P_SX';
    }

    if (areaCode === 'F1R' || (floorCode === 'F1' && (sideCode === 'RIGHT' || sideLabel.indexOf('DESTRO') >= 0))) {
      return '1P_DX';
    }

    if (areaCode === 'F1L' || (floorCode === 'F1' && (sideCode === 'LEFT' || sideLabel.indexOf('SINISTRO') >= 0))) {
      return '1P_SX';
    }

    if (areaCode === 'F0' || areaCode === 'GYM' || floorCode === 'F0' || floorLabel.indexOf('TERRA') >= 0) {
      return 'PT';
    }

    return 'PT';
  },

  normalizeKey_: function (value) {
    return ROOMS_APP.normalizeString(value).toUpperCase();
  },

  listResourcesForBoard_: function () {
    return ROOMS_APP.sortBy(
      ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.RESOURCES).filter(function (row) {
        return ROOMS_APP.asBoolean(row.IsActive);
      }),
      ['LayoutPage', 'LayoutRow', 'LayoutCol', 'DisplayName']
    );
  },

  listResources_: function () {
    return this.listResourcesForBoard_();
  }
};
