// ═══════════════════════════════════════════════════════
//  NCM B2B TRACKER — BACKEND v10 (fast live board, no shipments)
//  This version REMOVES the whole Send/Receive shipment
//  system and the "SHIPMENT GPS" sheet. The app is now:
//    • Driver check-in / check-out with round detection + controlled edits
//    • Live "All Vans" board
//    • Branch announcements (post/edit, one per branch per day)
//    • Branch contact persons (tappable to call)
//    • Issue reporting
//  Storage is split for speed:
//    • VAN MOVEMENTS = append/history and edits
//    • VAN LIVE = one current row per van for live reads
//  Old "SHIPMENT GPS" sheet is not used.
//  Deploy → Manage deployments → Edit → New version → Deploy
// ═══════════════════════════════════════════════════════

const CONFIG = {
  TIMEZONE: "Asia/Kathmandu"
};

const PASSCODES = {
  "1111": { branch: "TINKUNE",      role: "branch" },
  "2222": { branch: "CHABAHIL",     role: "branch" },
  "3333": { branch: "NAYA BUSPARK", role: "branch" },
  "4444": { branch: "KALANKI",      role: "branch" },
  "5555": { branch: "SATDOBATO",    role: "branch" },
  "6666": { branch: "NEWROAD",      role: "branch" },
  "7777": { role: "driver" },
  "9999": { role: "admin" }
};

const DESTINATIONS = {
  "TINKUNE":      ["NAYA THIMI","SURYABINAYAK","LUBHU","CHABAHIL","NAYA BUSPARK","KALANKI","SATDOBATO","NEWROAD"],
  "CHABAHIL":     ["KAPAN","BUDHANILKANTHA","SANKHU","SUNDARIJAL","NAYA BUSPARK","KALANKI","SATDOBATO","TINKUNE","NEWROAD"],
  "NAYA BUSPARK": ["NEWROAD","KALANKI","SATDOBATO","TINKUNE","CHABAHIL","SWOYAMBHU","BASUNDHARA"],
  "BASUNDHARA":   ["NAYA BUSPARK","CHABAHIL","SWOYAMBHU","TINKUNE","KALANKI","SATDOBATO","NEWROAD"],
  "SWOYAMBHU":    ["KALANKI","NAYA BUSPARK","BASUNDHARA","CHABAHIL","TINKUNE","SATDOBATO","NEWROAD"],
  "KALANKI":      ["THANKOT","SATDOBATO","TINKUNE","CHABAHIL","NAYA BUSPARK","NEWROAD"],
  "SATDOBATO":    ["TINKUNE","CHABAHIL","NAYA BUSPARK","KALANKI","NEWROAD","CHAPAGAU","GODAWARI"],
  "NEWROAD":      ["CHABAHIL","NAYA BUSPARK","KALANKI","SATDOBATO","TINKUNE"]
};

const MAIN_BRANCHES = Object.keys(DESTINATIONS);

// The 6 core stops a full round should touch, besides Tinkune (start/end).
const ROUND_CORE_BRANCHES = ["CHABAHIL","BASUNDHARA","NAYA BUSPARK","SWOYAMBHU","KALANKI","SATDOBATO"];

const MIN_TRAVEL_SECONDS = 120;
const EXPECTED_LEG_MINUTES = 10;

/* ═══ SMART POLLING CACHE ═══ */
const DRIVER_STATE_TTL = 5;   // seconds; elapsed time is calculated on read
const VAN_BOARD_TTL = 8;      // seconds; VAN LIVE contains one row per van
const INFO_TTL = 60;          // seconds

function getCached_(key) {
  try { return CacheService.getScriptCache().get(key); } catch (e) { return null; }
}
function setCached_(key, value, ttl) {
  try { CacheService.getScriptCache().put(key, value, ttl); } catch (e) {}
}
function removeCached_(key) {
  try { CacheService.getScriptCache().remove(key); } catch (e) {}
}
function invalidateDriverCaches(vanNo) {
  var key = String(vanNo || '').trim();
  removeCached_('drv_state_' + key);
  removeCached_('van_board_raw_' + today());
}

function today() {
  return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd");
}

function normalizeDateStr(value) {
  if (value instanceof Date) return Utilities.formatDate(value, CONFIG.TIMEZONE, "yyyy-MM-dd");
  var str = String(value || "").trim();
  var match = str.match(/^(\d{4}-\d{2}-\d{2})/);
  return match ? match[1] : str;
}

function secondsBetween(a, b) {
  var first = a instanceof Date ? a : new Date(a);
  var second = b instanceof Date ? b : new Date(b);
  if (isNaN(first.getTime()) || isNaN(second.getTime())) return 0;
  return Math.max(0, Math.floor((second.getTime() - first.getTime()) / 1000));
}

function timeLabel(seconds) {
  var total = Math.max(0, Math.floor(Number(seconds) || 0));
  var hours = Math.floor(total / 3600);
  var minutes = Math.floor((total % 3600) / 60);
  if (hours) return hours + " hr" + (minutes ? " " + minutes + " min" : "");
  return minutes + " min";
}

/* ═══════════════════════════════════════════════════════
   DRIVER SIDE — free-form check-in/out + round detection
   ONE sheet only (VAN MOVEMENTS) for speed.
═══════════════════════════════════════════════════════ */

const MOVE_COL = {
  DATE:1, VAN:2, BRANCH:3, ARRIVAL:4, DEPARTURE:5,
  HOLD_SECONDS:6, HOLD_TIME:7, NEXT_BRANCH:8, TRAVEL_SECONDS:9, TRAVEL_TIME:10,
  STATUS:11, ROUND:12, MISSED:13, UPDATED:14
};

// VAN MOVEMENTS is the append/history sheet. VAN LIVE is the tiny current-state
// index: one row per van, used by every read-only board request.
const LIVE_COL = {
  DATE:1, VAN:2, BRANCH:3, ARRIVAL:4, DEPARTURE:5,
  HOLD_SECONDS:6, HOLD_TIME:7, NEXT_BRANCH:8, TRAVEL_SECONDS:9, TRAVEL_TIME:10,
  STATUS:11, ROUND:12, MISSED:13, UPDATED:14, SOURCE_ROW:15
};

function getMovementSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VAN MOVEMENTS");
  if (!sheet) {
    sheet = ss.insertSheet("VAN MOVEMENTS");
    sheet.appendRow([
      "Date","Van No","Branch","Check In","Check Out","Hold Seconds","Hold Time",
      "Next Branch","Travel Seconds","Travel Time","Status","Round","Missed Branches","Last Updated"
    ]);
    sheet.setFrozenRows(1);
    sheet.getRange(2, MOVE_COL.VAN, sheet.getMaxRows() - 1, 1).setNumberFormat('@');
    sheet.getRange(2, MOVE_COL.ARRIVAL, sheet.getMaxRows() - 1, 2).setNumberFormat("yyyy-mm-dd hh:mm:ss");
    sheet.getRange(2, MOVE_COL.UPDATED, sheet.getMaxRows() - 1, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
  } else if (sheet.getLastColumn() < MOVE_COL.UPDATED) {
    var headers = ["Date","Van No","Branch","Check In","Check Out","Hold Seconds","Hold Time",
      "Next Branch","Travel Seconds","Travel Time","Status","Round","Missed Branches","Last Updated"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

function getLiveSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VAN LIVE");
  if (!sheet) {
    sheet = ss.insertSheet("VAN LIVE");
    sheet.appendRow([
      "Date","Van No","Branch","Check In","Check Out","Hold Seconds","Hold Time",
      "Next Branch","Travel Seconds","Travel Time","Status","Round","Missed Branches",
      "Last Updated","Source Movement Row"
    ]);
    sheet.setFrozenRows(1);
    sheet.getRange(2, LIVE_COL.VAN, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat('@');
    sheet.getRange(2, LIVE_COL.ARRIVAL, Math.max(1, sheet.getMaxRows() - 1), 2).setNumberFormat("yyyy-mm-dd hh:mm:ss");
    sheet.getRange(2, LIVE_COL.UPDATED, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
  } else if (sheet.getLastColumn() < LIVE_COL.SOURCE_ROW) {
    sheet.getRange(1, 1, 1, LIVE_COL.SOURCE_ROW).setValues([[
      "Date","Van No","Branch","Check In","Check Out","Hold Seconds","Hold Time",
      "Next Branch","Travel Seconds","Travel Time","Status","Round","Missed Branches",
      "Last Updated","Source Movement Row"
    ]]);
  }
  return sheet;
}

function liveRowNumber_(vanNo) {
  var key = String(vanNo || "").trim();
  if (!key) return 0;
  var cacheKey = 'live_row_' + key;
  var cached = getCached_(cacheKey);
  if (cached && Number(cached) > 1) {
    var cachedRow = Number(cached);
    var cachedSheet = getLiveSheet();
    if (cachedRow <= cachedSheet.getLastRow() &&
        String(cachedSheet.getRange(cachedRow, LIVE_COL.VAN).getDisplayValue()).trim() === key) {
      return cachedRow;
    }
  }
  var sheet = getLiveSheet();
  var last = sheet.getLastRow();
  if (last < 2) return 0;
  var values = sheet.getRange(2, LIVE_COL.VAN, last - 1, 1).getDisplayValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === key) {
      setCached_(cacheKey, String(i + 2), 300);
      return i + 2;
    }
  }
  return 0;
}

function movementValuesFromLive_(row) {
  var values = row.slice(0, MOVE_COL.UPDATED);
  return values;
}

function syncLiveRow_(vanNo, movementValues, sourceRow) {
  var sheet = getLiveSheet();
  var key = String(vanNo || movementValues[MOVE_COL.VAN - 1] || "").trim();
  var liveValues = movementValuesFromLive_(movementValues);
  liveValues.push(Number(sourceRow) || 0);
  var row = liveRowNumber_(key);
  if (!row) {
    row = Math.max(2, sheet.getLastRow() + 1);
    sheet.getRange(row, 1, 1, LIVE_COL.SOURCE_ROW).setValues([liveValues]);
  } else {
    sheet.getRange(row, 1, 1, LIVE_COL.SOURCE_ROW).setValues([liveValues]);
  }
  setCached_('live_row_' + key, String(row), 300);
  return row;
}

function liveRowValues_(vanNo) {
  var row = liveRowNumber_(vanNo);
  if (!row) return null;
  return getLiveSheet().getRange(row, 1, 1, LIVE_COL.SOURCE_ROW).getValues()[0];
}

function ensureLiveSheetToday_() {
  var dateStr = today();
  var marker = 'live_ready_' + dateStr;
  if (getCached_(marker)) return;
  var migrationLock = LockService.getScriptLock();
  if (!migrationLock.tryLock(1000)) return;
  try {
    if (getCached_(marker)) return;
  var liveSheet = getLiveSheet();
  var hasToday = false;
  var liveLast = liveSheet.getLastRow();
  if (liveLast >= 2) {
    var liveDates = liveSheet.getRange(2, LIVE_COL.DATE, liveLast - 1, 1).getValues();
    for (var i = 0; i < liveDates.length; i++) {
      if (normalizeDateStr(liveDates[i][0]) === dateStr) {
        hasToday = true;
        break;
      }
    }
  }
  if (!hasToday) {
    var values = getMovementSheet().getDataRange().getValues();
    var latestByVan = {};
    for (var j = 1; j < values.length; j++) {
      var row = values[j];
      var van = String(row[MOVE_COL.VAN - 1] || '').trim();
      if (van && normalizeDateStr(row[MOVE_COL.DATE - 1]) === dateStr) {
        latestByVan[van] = { values: row, rowIndex: j + 1 };
      }
    }
    Object.keys(latestByVan).forEach(function(vanNo) {
      var item = latestByVan[vanNo];
      syncLiveRow_(vanNo, item.values, item.rowIndex);
    });
  }
  setCached_(marker, '1', 300);
  } finally {
    migrationLock.releaseLock();
  }
}

function getIssueSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ISSUES");
  if (!sheet) {
    sheet = ss.insertSheet("ISSUES");
    sheet.appendRow(["Date","Time","Reported By","Role","Branch","Van No","Issue","Status"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function movementRowsForToday(vanNo) {
  var values = getMovementSheet().getDataRange().getValues();
  var target = String(vanNo || "").trim();
  var dateStr = today();
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    if (normalizeDateStr(values[i][MOVE_COL.DATE - 1]) === dateStr &&
        String(values[i][MOVE_COL.VAN - 1]).trim() === target) {
      rows.push({ rowIndex: i + 1, values: values[i] });
    }
  }
  return rows;
}

function computeRoundInfo(priorRows, arrivingBranch) {
  var roundNumber = 1;
  var visitedThisRound = {};
  var closedRound = null;
  for (var i = 0; i < priorRows.length; i++) {
    var b = String(priorRows[i].values[MOVE_COL.BRANCH - 1] || "").toUpperCase().trim();
    if (b === "TINKUNE") {
      if (i > 0) roundNumber++;
      visitedThisRound = {};
    } else if (ROUND_CORE_BRANCHES.indexOf(b) !== -1) {
      visitedThisRound[b] = true;
    }
  }
  if (arrivingBranch === "TINKUNE" && priorRows.length > 0) {
    var missed = ROUND_CORE_BRANCHES.filter(function(b) { return !visitedThisRound[b]; });
    closedRound = { round: roundNumber, missed: missed };
    roundNumber++;
  }
  return { roundNumber: roundNumber, closedRound: closedRound };
}

function getDriverState(vanNo) {
  var vanKey = String(vanNo || "").trim();
  var cacheKey = 'drv_state_' + vanKey;
  var cached = getCached_(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }
  var state;
  var live = liveRowValues_(vanKey);
  if (live && normalizeDateStr(live[LIVE_COL.DATE - 1]) === today()) {
    state = {
      vanNo: String(live[LIVE_COL.VAN - 1]).trim(),
      started: true,
      branch: live[LIVE_COL.BRANCH - 1],
      nextBranch: live[LIVE_COL.NEXT_BRANCH - 1] || null,
      arrivalTime: live[LIVE_COL.ARRIVAL - 1] || null,
      departureTime: live[LIVE_COL.DEPARTURE - 1] || null,
      holdSeconds: Number(live[LIVE_COL.HOLD_SECONDS - 1]) || 0,
      travelSeconds: Number(live[LIVE_COL.TRAVEL_SECONDS - 1]) || 0,
      status: String(live[LIVE_COL.STATUS - 1] || "AT_STATION"),
      round: Number(live[LIVE_COL.ROUND - 1]) || 1,
      lastMissed: live[LIVE_COL.MISSED - 1] || "",
      rowIndex: Number(live[LIVE_COL.SOURCE_ROW - 1]) || 0
    };
  } else {
    // One-time compatibility fallback for movement history created before
    // VAN LIVE existed. The next write automatically seeds the live row.
    var rows = movementRowsForToday(vanNo);
    if (!rows.length) {
      state = { vanNo: vanKey, started: false, branch: null, nextBranch: null, status: "NOT_STARTED", round: 0 };
    } else {
      var item = rows[rows.length - 1];
      var v = item.values;
      state = {
        vanNo: String(v[MOVE_COL.VAN - 1]).trim(),
        started: true,
        branch: v[MOVE_COL.BRANCH - 1],
        nextBranch: v[MOVE_COL.NEXT_BRANCH - 1] || null,
        arrivalTime: v[MOVE_COL.ARRIVAL - 1] || null,
        departureTime: v[MOVE_COL.DEPARTURE - 1] || null,
        holdSeconds: Number(v[MOVE_COL.HOLD_SECONDS - 1]) || 0,
        travelSeconds: Number(v[MOVE_COL.TRAVEL_SECONDS - 1]) || 0,
        status: String(v[MOVE_COL.STATUS - 1] || "AT_STATION"),
        round: Number(v[MOVE_COL.ROUND - 1]) || 1,
        lastMissed: v[MOVE_COL.MISSED - 1] || "",
        rowIndex: item.rowIndex
      };
      syncLiveRow_(vanKey, v, item.rowIndex);
    }
  }
  setCached_(cacheKey, JSON.stringify(state), DRIVER_STATE_TTL);
  return state;
}

function driverStateResponse(state) {
  var now = new Date();
  var elapsed = 0;
  if (state.status === "MOVING") elapsed = secondsBetween(state.departureTime, now);
  if (state.status === "AT_STATION") elapsed = secondsBetween(state.arrivalTime, now);
  return {
    vanNo: state.vanNo, started: state.started, branch: state.branch,
    nextBranch: state.nextBranch, status: state.status, elapsedSeconds: elapsed,
    holdSeconds: state.holdSeconds || 0, travelSeconds: state.travelSeconds || 0,
    round: state.round || 1, lastMissed: state.lastMissed || "",
    arrivalTime: state.arrivalTime || null, departureTime: state.departureTime || null
  };
}

function handleDriverCheckIn(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var vanNo = String(data.vanNo || "").trim();
    var branch = String(data.branch || "").toUpperCase().trim();
    if (!vanNo) return { success:false, error:"Enter your van number first" };
    if (!branch) return { success:false, error:"Select a branch" };

    var sheet = getMovementSheet();
    var state = getDriverState(vanNo);
    var now = new Date();
    var priorRows = movementRowsForToday(vanNo);

    if (!state.started) {
      var info0 = computeRoundInfo([], branch);
      sheet.appendRow([today(), vanNo, branch, now, "", 0, "", "", 0, "", "AT_STATION", info0.roundNumber, "", now]);
      var firstRow = sheet.getLastRow();
      syncLiveRow_(vanNo, sheet.getRange(firstRow, 1, 1, MOVE_COL.UPDATED).getValues()[0], firstRow);
      invalidateDriverCaches(vanNo);
      return { success:true, message:"Checked in at " + branch, state:driverStateResponse(getDriverState(vanNo)) };
    }

    if (state.status === "AT_STATION") {
      if (String(state.branch || "").toUpperCase().trim() === branch) {
        return { success:true, message:"Check-in already saved at " + state.branch,
          state:driverStateResponse(state) };
      }
      return { success:false, error:"Already checked in at " + state.branch + ". Check out before moving." };
    }
    if (state.status !== "MOVING") {
      return { success:false, error:"Already checked in at " + state.branch + ". Check out before moving." };
    }
    if (branch !== String(state.nextBranch || "").toUpperCase().trim()) {
      return { success:false, error:"Expected check-in at " + state.nextBranch };
    }
    var travelSeconds = secondsBetween(state.departureTime, now);
    if (travelSeconds < MIN_TRAVEL_SECONDS) {
      return { success:false, error:"Check-in opens after 2 minutes of travel" };
    }

    var values = sheet.getRange(state.rowIndex, 1, 1, MOVE_COL.UPDATED).getValues()[0];
    values[MOVE_COL.TRAVEL_SECONDS - 1] = travelSeconds;
    values[MOVE_COL.TRAVEL_TIME - 1] = timeLabel(travelSeconds);
    values[MOVE_COL.STATUS - 1] = "COMPLETE";
    values[MOVE_COL.UPDATED - 1] = now;
    sheet.getRange(state.rowIndex, 1, 1, values.length).setValues([values]);

    var info = computeRoundInfo(priorRows, branch);
    var missedStr = info.closedRound ? (info.closedRound.missed.length ? info.closedRound.missed.join(", ") : "None") : "";
    sheet.appendRow([today(), vanNo, branch, now, "", 0, "", "", 0, "", "AT_STATION", info.roundNumber, missedStr, now]);
    var newRow = sheet.getLastRow();
    syncLiveRow_(vanNo, sheet.getRange(newRow, 1, 1, MOVE_COL.UPDATED).getValues()[0], newRow);
    invalidateDriverCaches(vanNo);

    var msg = "Checked in at " + branch;
    if (info.closedRound) {
      msg = "Round " + info.closedRound.round + " complete at " + branch +
        (info.closedRound.missed.length ? " — missed: " + info.closedRound.missed.join(", ") : " — all branches visited");
    }
    return { success:true, message: msg, state:driverStateResponse(getDriverState(vanNo)) };
  } finally {
    lock.releaseLock();
  }
}

function handleDriverCheckOut(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var vanNo = String(data.vanNo || "").trim();
    var nextBranch = String(data.nextBranch || "").toUpperCase().trim();
    if (!vanNo) return { success:false, error:"Enter your van number first" };
    if (!nextBranch) return { success:false, error:"Select where you're heading" };

    var state = getDriverState(vanNo);
    if (!state.started) return { success:false, error:"Check in first" };
    if (state.status === "MOVING") {
      if (String(state.nextBranch || "").toUpperCase().trim() === nextBranch) {
        return { success:true, message:"Check-out already saved — heading to " + nextBranch,
          state:driverStateResponse(state) };
      }
      return { success:false, error:"Already heading to " + state.nextBranch };
    }
    if (state.status !== "AT_STATION") return { success:false, error:"Van is not at a station" };
    var currentBranch = String(state.branch || "").toUpperCase().trim();
    var allowedDestinations = DESTINATIONS[currentBranch] || [];
    if (allowedDestinations.indexOf(nextBranch) === -1) {
      return { success:false, error:"Choose a branch from the list" };
    }

    var now = new Date();
    var sheet = getMovementSheet();
    var values = sheet.getRange(state.rowIndex, 1, 1, MOVE_COL.UPDATED).getValues()[0];
    var holdSeconds = secondsBetween(state.arrivalTime, now);
    values[MOVE_COL.DEPARTURE - 1] = now;
    values[MOVE_COL.HOLD_SECONDS - 1] = holdSeconds;
    values[MOVE_COL.HOLD_TIME - 1] = timeLabel(holdSeconds);
    values[MOVE_COL.NEXT_BRANCH - 1] = nextBranch;
    values[MOVE_COL.STATUS - 1] = "MOVING";
    values[MOVE_COL.UPDATED - 1] = now;
    sheet.getRange(state.rowIndex, 1, 1, values.length).setValues([values]);
    syncLiveRow_(vanNo, values, state.rowIndex);
    invalidateDriverCaches(vanNo);

    return { success:true, message:"Checked out — heading to " + nextBranch, state:driverStateResponse(getDriverState(vanNo)) };
  } finally {
    lock.releaseLock();
  }
}

function parseEditDate_(value) {
  var d = value instanceof Date ? value : new Date(String(value || ""));
  return isNaN(d.getTime()) ? null : d;
}

function handleEditDriverMovement(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var vanNo = String(data.vanNo || "").trim();
    var editType = String(data.editType || "").trim();
    var state = getDriverState(vanNo);
    if (!vanNo || !state.started || !state.rowIndex) {
      return { success:false, error:"No editable movement found for this van" };
    }
    if (editType !== "checkIn" && editType !== "checkOut") {
      return { success:false, error:"Choose check-in or check-out" };
    }
    var editedTime = parseEditDate_(data.timestamp);
    if (!editedTime) return { success:false, error:"Invalid movement time" };
    var sheet = getMovementSheet();
    if (state.rowIndex < 2 || state.rowIndex > sheet.getLastRow()) {
      return { success:false, error:"Movement row is no longer available" };
    }
    var values = sheet.getRange(state.rowIndex, 1, 1, MOVE_COL.UPDATED).getValues()[0];
    if (String(values[MOVE_COL.VAN - 1]).trim() !== vanNo) {
      return { success:false, error:"Van movement changed. Refresh and try again." };
    }

    if (editType === "checkIn") {
      var branch = String(data.branch || "").toUpperCase().trim();
      if (!DESTINATIONS[branch]) return { success:false, error:"Choose a valid check-in branch" };
      var departure = values[MOVE_COL.DEPARTURE - 1];
      if (departure instanceof Date && editedTime.getTime() > departure.getTime()) {
        return { success:false, error:"Check-in must be before check-out" };
      }
      if (String(values[MOVE_COL.STATUS - 1] || "") === "MOVING") {
        var next = String(values[MOVE_COL.NEXT_BRANCH - 1] || "").toUpperCase().trim();
        if ((DESTINATIONS[branch] || []).indexOf(next) === -1) {
          return { success:false, error:"That branch cannot use the current next destination" };
        }
      }
      values[MOVE_COL.BRANCH - 1] = branch;
      values[MOVE_COL.ARRIVAL - 1] = editedTime;
      if (departure instanceof Date) {
        var travel = secondsBetween(editedTime, departure);
        values[MOVE_COL.TRAVEL_SECONDS - 1] = travel;
        values[MOVE_COL.TRAVEL_TIME - 1] = timeLabel(travel);
      }
    } else {
      if (String(values[MOVE_COL.STATUS - 1] || "") !== "MOVING") {
        return { success:false, error:"Check out is available after a van starts moving" };
      }
      var nextBranch = String(data.nextBranch || "").toUpperCase().trim();
      var currentBranch = String(values[MOVE_COL.BRANCH - 1] || "").toUpperCase().trim();
      if ((DESTINATIONS[currentBranch] || []).indexOf(nextBranch) === -1) {
        return { success:false, error:"Choose a valid next branch" };
      }
      var arrival = values[MOVE_COL.ARRIVAL - 1];
      if (arrival instanceof Date && editedTime.getTime() < arrival.getTime()) {
        return { success:false, error:"Check-out must be after check-in" };
      }
      var hold = secondsBetween(arrival, editedTime);
      values[MOVE_COL.DEPARTURE - 1] = editedTime;
      values[MOVE_COL.HOLD_SECONDS - 1] = hold;
      values[MOVE_COL.HOLD_TIME - 1] = timeLabel(hold);
      values[MOVE_COL.NEXT_BRANCH - 1] = nextBranch;
    }
    values[MOVE_COL.UPDATED - 1] = new Date();
    sheet.getRange(state.rowIndex, 1, 1, values.length).setValues([values]);
    syncLiveRow_(vanNo, values, state.rowIndex);
    invalidateDriverCaches(vanNo);
    return {
      success:true,
      message: editType === "checkIn" ? "Check-in updated" : "Check-out updated",
      state: driverStateResponse(getDriverState(vanNo))
    };
  } finally {
    lock.releaseLock();
  }
}

function getVanBoard() {
  ensureLiveSheetToday_();
  var dateStr = today();
  var cacheKey = 'van_board_raw_' + dateStr;
  var rows = null;
  var cached = getCached_(cacheKey);
  if (cached) {
    try { rows = JSON.parse(cached); } catch (e) { rows = null; }
  }
  if (!rows) {
    rows = [];
    var sheet = getLiveSheet();
    var last = sheet.getLastRow();
    var values = last >= 2 ? sheet.getRange(2, 1, last - 1, LIVE_COL.SOURCE_ROW).getValues() : [];
    for (var i = 0; i < values.length; i++) {
      var v = values[i];
      var van = String(v[LIVE_COL.VAN - 1] || "").trim();
      if (!van) continue;
      if (normalizeDateStr(v[LIVE_COL.DATE - 1]) !== dateStr) continue;
      rows.push({
        vanNo: van,
        status: String(v[LIVE_COL.STATUS - 1] || ""),
        branch: v[LIVE_COL.BRANCH - 1],
        nextBranch: v[LIVE_COL.NEXT_BRANCH - 1],
        round: Number(v[LIVE_COL.ROUND - 1]) || 1,
        arrivalMs: v[LIVE_COL.ARRIVAL - 1] instanceof Date ? v[LIVE_COL.ARRIVAL - 1].getTime() : 0,
        departureMs: v[LIVE_COL.DEPARTURE - 1] instanceof Date ? v[LIVE_COL.DEPARTURE - 1].getTime() : 0
      });
    }
    setCached_(cacheKey, JSON.stringify(rows), VAN_BOARD_TTL);
  }
  var nowMs = Date.now();
  return rows.map(function(item) {
    var out = { vanNo: item.vanNo, round: item.round };
    if (item.status === "MOVING") {
      var elapsed = Math.max(0, Math.floor((nowMs - item.departureMs) / 1000));
      out.status = "moving";
      out.fromBranch = item.branch;
      out.toBranch = item.nextBranch;
      out.elapsedSeconds = elapsed;
      out.isLate = Math.floor(elapsed / 60) >= EXPECTED_LEG_MINUTES;
    } else {
      out.status = "at_branch";
      out.branch = item.branch;
      out.elapsedSeconds = Math.max(0, Math.floor((nowMs - item.arrivalMs) / 1000));
    }
    return out;
  }).sort(function(a, b) {
    if (a.status === 'moving' && b.status !== 'moving') return -1;
    if (a.status !== 'moving' && b.status === 'moving') return 1;
    return 0;
  });
}

function getLiveVans(vanList) {
  ensureLiveSheetToday_();
  var requested = String(vanList || "").split(",").map(function(v) {
    return v.trim();
  }).filter(Boolean).slice(0, 5);
  if (!requested.length) return getVanBoard();
  var dateStr = today();
  var rows = [];
  requested.forEach(function(vanNo) {
    var v = liveRowValues_(vanNo);
    if (!v || normalizeDateStr(v[LIVE_COL.DATE - 1]) !== dateStr) return;
    rows.push({
      vanNo: String(v[LIVE_COL.VAN - 1]).trim(),
      status: String(v[LIVE_COL.STATUS - 1] || ""),
      branch: v[LIVE_COL.BRANCH - 1],
      nextBranch: v[LIVE_COL.NEXT_BRANCH - 1],
      round: Number(v[LIVE_COL.ROUND - 1]) || 1,
      arrivalMs: v[LIVE_COL.ARRIVAL - 1] instanceof Date ? v[LIVE_COL.ARRIVAL - 1].getTime() : 0,
      departureMs: v[LIVE_COL.DEPARTURE - 1] instanceof Date ? v[LIVE_COL.DEPARTURE - 1].getTime() : 0
    });
  });
  var nowMs = Date.now();
  return rows.map(function(item) {
    var out = { vanNo: item.vanNo, round: item.round };
    if (item.status === "MOVING") {
      var elapsed = Math.max(0, Math.floor((nowMs - item.departureMs) / 1000));
      out.status = "moving";
      out.fromBranch = item.branch;
      out.toBranch = item.nextBranch;
      out.elapsedSeconds = elapsed;
      out.isLate = Math.floor(elapsed / 60) >= EXPECTED_LEG_MINUTES;
    } else {
      out.status = "at_branch";
      out.branch = item.branch;
      out.elapsedSeconds = Math.max(0, Math.floor((nowMs - item.arrivalMs) / 1000));
    }
    return out;
  });
}

function submitIssue(data) {
  var message = String(data.issue || "").trim();
  if (!message) return { success:false, error:"Write the issue before submitting" };
  getIssueSheet().appendRow([
    today(), Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "HH:mm:ss"),
    data.reportedBy || "", data.role || "", data.branch || "",
    data.vanNo || "", message, "OPEN"
  ]);
  return { success:true, message:"Issue sent to admin" };
}

function getIssues(includeClosed) {
  var values = getIssueSheet().getDataRange().getValues();
  var result = [];
  for (var i = 1; i < values.length; i++) {
    if (!includeClosed && String(values[i][7]).toUpperCase() === "CLOSED") continue;
    result.push({
      row: i + 1, date: values[i][0], time: values[i][1],
      reportedBy: values[i][2], role: values[i][3], branch: values[i][4],
      vanNo: values[i][5], issue: values[i][6], status: values[i][7]
    });
  }
  return result.reverse();
}

function closeIssue(row) {
  var sheet = getIssueSheet();
  if (!row || row < 2 || row > sheet.getLastRow()) return { success:false, error:"Invalid issue" };
  sheet.getRange(Number(row), 8).setValue("CLOSED");
  return { success:true, message:"Issue closed" };
}

/* ═══════════════════════════════════════════════════════
   ANNOUNCEMENTS & BRANCH CONTACTS — admin/branch panels
═══════════════════════════════════════════════════════ */

function getAnnouncementSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ANNOUNCEMENTS");
  if (!sheet) {
    sheet = ss.insertSheet("ANNOUNCEMENTS");
    sheet.appendRow(["Date","Branch","Message","Time"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getBranchContactSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("BRANCH CONTACTS");
  if (!sheet) {
    sheet = ss.insertSheet("BRANCH CONTACTS");
    sheet.appendRow(["Branch","Contact Person","Phone","Updated"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function handlePutAnnouncement(data) {
  var branch = String(data.branch || "").toUpperCase().trim();
  var message = String(data.message || "").trim();
  if (!branch || !DESTINATIONS[branch]) return { success:false, error:"Invalid branch" };
  if (!message) return { success:false, error:"Write the announcement first" };
  if (message.length > 300) return { success:false, error:"Keep it under 300 characters" };
  var sheet = getAnnouncementSheet();
  var values = sheet.getDataRange().getValues();
  var dateStr = today();
  var timeStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "HH:mm");
  for (var i = 1; i < values.length; i++) {
    if (normalizeDateStr(values[i][0]) === dateStr &&
        String(values[i][1]).toUpperCase().trim() === branch) {
      sheet.getRange(i + 1, 3, 1, 2).setValues([[message, timeStr]]);
      removeCached_('announcements_' + dateStr);
      return { success:true, message:"Announcement updated for " + branch };
    }
  }
  sheet.appendRow([dateStr, branch, message, timeStr]);
  removeCached_('announcements_' + dateStr);
  return { success:true, message:"Announcement posted for " + branch };
}

function getAnnouncements() {
  var dateStr = today();
  var key = 'announcements_' + dateStr;
  var cached = getCached_(key);
  if (cached) { try { return JSON.parse(cached); } catch (e) {} }
  var values = getAnnouncementSheet().getDataRange().getValues();
  var result = [];
  for (var i = 1; i < values.length; i++) {
    if (normalizeDateStr(values[i][0]) === dateStr) {
      result.push({ branch: values[i][1], message: values[i][2], time: values[i][3] });
    }
  }
  setCached_(key, JSON.stringify(result), INFO_TTL);
  return result;
}

function handleSaveBranchContact(data) {
  var branch = String(data.branch || "").toUpperCase().trim();
  var name = String(data.name || "").trim();
  var phone = String(data.phone || "").trim();
  if (!branch || !DESTINATIONS[branch]) return { success:false, error:"Invalid branch" };
  if (!name) return { success:false, error:"Enter the contact person's name" };
  if (!phone) return { success:false, error:"Enter a phone number" };
  if (phone.length > 20) return { success:false, error:"Phone number looks too long" };
  var sheet = getBranchContactSheet();
  var values = sheet.getDataRange().getValues();
  var nowStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm");
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).toUpperCase().trim() === branch) {
      sheet.getRange(i + 1, 2, 1, 3).setValues([[name, phone, nowStr]]);
      removeCached_('branch_contacts');
      return { success:true, message:"Contact updated for " + branch };
    }
  }
  sheet.appendRow([branch, name, phone, nowStr]);
  removeCached_('branch_contacts');
  return { success:true, message:"Contact saved for " + branch };
}

function getBranchContacts() {
  var key = 'branch_contacts';
  var cached = getCached_(key);
  if (cached) { try { return JSON.parse(cached); } catch (e) {} }
  var values = getBranchContactSheet().getDataRange().getValues();
  var result = [];
  for (var i = 1; i < values.length; i++) {
    result.push({ branch: values[i][0], name: values[i][1], phone: values[i][2], updated: values[i][3] });
  }
  setCached_(key, JSON.stringify(result), INFO_TTL);
  return result;
}

/* ═══════════════════════════════════════════════════════
   WEB APP ENTRY POINTS
═══════════════════════════════════════════════════════ */

function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  if (!action) return jsonResponse({ success: false, error: "No action" });

  if (action === "login") {
    var pc = e.parameter.passcode;
    var cfg = PASSCODES[pc];
    if (!cfg) return jsonResponse({ success: false, error: "Invalid passcode" });
    var res = { success: true, role: cfg.role };
    if (cfg.role === "branch") { res.branch = cfg.branch; res.destinations = DESTINATIONS[cfg.branch] || []; }
    return jsonResponse(res);
  }

  if (action === "getDriverState") {
    return jsonResponse({ success:true, state:driverStateResponse(getDriverState(e.parameter.vanNo)) });
  }
  if (action === "getDestinations") {
    var requestedBranch = String(e.parameter.branch || "").toUpperCase().trim();
    return jsonResponse({ success:true, destinations: DESTINATIONS[requestedBranch] || [], mainBranches: MAIN_BRANCHES });
  }
  if (action === "getVanBoard") {
    return jsonResponse({ success:true, data:getVanBoard(), date:today() });
  }
  if (action === "getLiveVans") {
    return jsonResponse({ success:true, data:getLiveVans(e.parameter.vans), date:today() });
  }
  if (action === "getIssues") {
    return jsonResponse({ success:true, data:getIssues(false) });
  }
  if (action === "getAnnouncements") {
    return jsonResponse({ success:true, data:getAnnouncements(), date:today() });
  }
  if (action === "getBranchContacts") {
    return jsonResponse({ success:true, data:getBranchContacts() });
  }

  return jsonResponse({ success: false, error: "Unknown action" });
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return jsonResponse({ success: false, error: "Bad JSON" }); }

  switch (body.action) {
    case 'driverCheckIn':     return jsonResponse(handleDriverCheckIn(body));
    case 'driverCheckOut':    return jsonResponse(handleDriverCheckOut(body));
    case 'editDriverMovement':return jsonResponse(handleEditDriverMovement(body));
    case 'submitIssue':       return jsonResponse(submitIssue(body));
    case 'putAnnouncement':   return jsonResponse(handlePutAnnouncement(body));
    case 'saveBranchContact': return jsonResponse(handleSaveBranchContact(body));
    case 'closeIssue':        return jsonResponse(closeIssue(body.row));
    default:                  return jsonResponse({ success: false, error: 'Unknown action' });
  }
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
