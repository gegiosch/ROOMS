var ROOMS_APP = ROOMS_APP || {};

ROOMS_APP.Board = {
  getBoardViewModel: function () {
    ROOMS_APP.Schema.ensureAll();
    var now = new Date();
    var nowIso = ROOMS_APP.toIsoDateTime(now);
    var today = ROOMS_APP.toIsoDate(now);
    var currentTime = Utilities.formatDate(now, ROOMS_APP.getTimezone(), 'HH:mm');
    var resources = this.listResources_();
    var bookings = ROOMS_APP.Booking.listBookingsForDate(today);
    var byResource = {};

    bookings.forEach(function (booking) {
      byResource[booking.ResourceId] = byResource[booking.ResourceId] || [];
      byResource[booking.ResourceId].push(booking);
    });

    var pagesMap = {};
    resources.forEach(function (resource) {
      var roomBookings = ROOMS_APP.sortBy(byResource[resource.ResourceId] || [], ['StartTime']);
      var current = roomBookings.filter(function (booking) {
        return booking.StartTime <= currentTime && booking.EndTime > currentTime;
      })[0] || null;
      var next = roomBookings.filter(function (booking) {
        return booking.StartTime > currentTime;
      })[0] || null;
      var state = current ? 'OCCUPIED' : (next ? 'NEXT_OCCUPIED' : 'FREE');
      var pageId = String(resource.LayoutPage || '1');

      pagesMap[pageId] = pagesMap[pageId] || [];
      pagesMap[pageId].push({
        resourceId: resource.ResourceId,
        displayName: resource.DisplayName,
        areaLabel: resource.AreaLabel,
        floorLabel: resource.FloorLabel,
        sideLabel: resource.SideLabel,
        layoutRow: Number(resource.LayoutRow || 1),
        layoutCol: Number(resource.LayoutCol || 1),
        layoutColSpan: Number(resource.LayoutColSpan || 1),
        layoutRowSpan: Number(resource.LayoutRowSpan || 1),
        state: state,
        currentLabel: current ? (current.BookerSurname || current.Title || 'OCCUPATA') : 'LIBERA',
        nextLabel: next ? (next.BookerSurname || next.Title || 'PRENOTATA') : 'LIBERA'
      });
    });

    var maxPages = ROOMS_APP.getNumberConfig('BOARD_PAGE_COUNT', 2);
    var pageIds = Object.keys(pagesMap).sort().slice(0, maxPages);

    return {
      generatedAtISO: nowIso,
      date: today,
      refreshSec: ROOMS_APP.getNumberConfig('BOARD_REFRESH_SEC', 60),
      rotationSec: ROOMS_APP.getNumberConfig('BOARD_ROTATION_SEC', 15),
      pageCount: pageIds.length,
      palette: ROOMS_APP.PALETTE,
      appName: ROOMS_APP.getConfigValue('APP_NAME', 'ROOMS'),
      schoolName: ROOMS_APP.getConfigValue('SCHOOL_NAME', 'IIS Alessandrini'),
      pages: pageIds.map(function (pageId) {
        var tiles = ROOMS_APP.sortBy(pagesMap[pageId], ['layoutRow', 'layoutCol', 'displayName']);
        var rows = {};
        tiles.forEach(function (tile) {
          rows[tile.layoutRow] = true;
        });

        return {
          pageId: pageId,
          title: 'Pagina ' + pageId,
          rowIds: Object.keys(rows).map(function (rowId) { return Number(rowId); }).sort(),
          tiles: tiles
        };
      })
    };
  },

  listResources_: function () {
    return ROOMS_APP.sortBy(
      ROOMS_APP.DB.readRows(ROOMS_APP.SHEET_NAMES.RESOURCES).filter(function (row) {
        return ROOMS_APP.asBoolean(row.IsActive);
      }),
      ['LayoutPage', 'LayoutRow', 'LayoutCol', 'DisplayName']
    );
  }
};
