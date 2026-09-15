// ═══════════════════════════════════════════════════════
//  NCM B2B TRACKER — BACKEND v8 (smart polling + admin panels)
//  New in this version:
//   • Round detection for drivers: a round closes whenever the van
//     checks back into TINKUNE. The branches visited since the last
//     TINKUNE check-in are compared against the 6 core branches
//     (Chabahil, Basundhara, Naya Buspark, Swoyambhu, Kalanki,
//     Satdobato) — skipping one is fine, it's just reported as missed,
//     not blocked.
//   • Readable Hold Time / Travel Time text columns alongside the raw
//     seconds, so the sheet itself is easy to read.
//   • getVanBoard(): a fast, company-wide "where is every van right now"
//     list — now additionally CACHED for 30s so 50 phones polling no
//     longer means 50 sheet reads per minute.
//   • Script Cache on driver state + van board: polling no longer
//     re-reads sheets on every request; writes invalidate the cache.
//  Branch Send/Receive system is unchanged.
//  Open from: Extensions → Apps Script inside your sheet
//  After pasting: Deploy → Manage deployments → Edit → New version → Deploy
// ═══════════════════════════════════════════════════════

const CONFIG = {
  SHEET_NAME: "SHIPMENT GPS",
  TIMEZONE: "Asia/Kathmandu",
  TOTAL_COLS: 9
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

const SUB_BRANCHES = {
  "KAPAN":"CHABAHIL","BUDHANILKANTHA":"CHABAHIL","SANKHU":"CHABAHIL","SUNDARIJAL":"CHABAHIL",
  "NAYA THIMI":"TINKUNE","SURYABINAYAK":"TINKUNE","LUBHU":"TINKUNE",
  "GODAWARI":"SATDOBATO","CHAPAGAU":"SATDOBATO",
  "SWOYAMBHU":"NAYA BUSPARK","BASUNDHARA":"NAYA BUSPARK",
  "THANKOT":"KALANKI"
};

const MIN_TRAVEL_SECONDS = 120;
const EXPECTED_LEG_MINUTES = 10;

/* ═══ SMART POLLING CACHE ═══
   Every phone polls every few seconds, and each poll used to re-read the
   whole sheet. Apps Script has strict daily quotas, so the app got slower
   as more people used it. This cache lets ONE sheet read serve EVERYONE
   for a short window. All writes invalidate the cache, so data stays fresh. */
const DRIVER_STATE_TTL = 60;  // seconds
const VAN_BOARD_TTL = 30;     // seconds

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
  removeCached_('drv_state_' + String(vanNo || '').trim());
  removeCached_('van_board_raw_' + today());
}

const COL = { DATE:1, ORIGIN:2, DEST:3, SEND_TIME:4, ID:5, VAN:6, STATUS:7, RECV_BY:8, RECV_TIME:9 };

function resolveMainBranch(name) {
  if (!name) return name;
  var upper = name.toString().toUpperCase().trim();
  return SUB_BRANCHES[upper] || upper;
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
   BRANCH SIDE — shipment Send/Receive (unchanged behavior)
═══════════════════════════════════════════════════════ */

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.appendRow(["Date","Origin","Destination","Send Time","Shipment ID","Van No","Status","Received By","Received Time"]);
    sheet.setFrozenRows(1);
  }
  var lastCol = sheet.getLastColumn();
  if (lastCol < CONFIG.TOTAL_COLS) sheet.insertColumnsAfter(lastCol, CONFIG.TOTAL_COLS - lastCol);
  return sheet;
}

function getAllValues(sheet) {
  var lr = sheet.getLastRow();
  if (lr < 1) lr = 1;
  return sheet.getRange(1, 1, lr, CONFIG.TOTAL_COLS).getValues();
}

function setAllValues(sheet, values) {
  var lr = values.length;
  if (lr < 1) lr = 1;
  sheet.getRange(1, 1, lr, CONFIG.TOTAL_COLS).setValues(values);
}

function parseDateTime(dateVal, timeVal) {
  var d;
  try {
    if (dateVal instanceof Date) {
      d = new Date(dateVal.getFullYear(), dateVal.getMonth(), dateVal.getDate());
    } else {
      var ds = String(dateVal).trim();
      var m = ds.match(/(\d{4})-(\d{2})-(\d{2})/);
      d = m ? new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3])) : new Date(ds);
    }
  } catch(e) { d = new Date(); }
  if (!d || isNaN(d.getTime())) d = new Date();

  var h = 0, mn = 0, s = 0;
  try {
    if (timeVal instanceof Date) {
      h = timeVal.getHours(); mn = timeVal.getMinutes(); s = timeVal.getSeconds();
    } else {
      var p = String(timeVal || "00:00:00").match(/(\d{1,2}):(\d{2}):(\d{2})/);
      if (p) { h = parseInt(p[1]); mn = parseInt(p[2]); s = parseInt(p[3]); }
    }
  } catch(e) {}
  d.setHours(h, mn, s, 0);
  return d;
}

function getIncomingVans(branch) {
  try {
    var sheet = getSheet();
    var values = getAllValues(sheet);
    var now = new Date();
    var vans = {};
    for (var i = 1; i < values.length; i++) {
      var dest = values[i][COL.DEST - 1];
      var van = String(values[i][COL.VAN - 1] || "").trim();
      var status = String(values[i][COL.STATUS - 1] || "").trim();
      var origin = values[i][COL.ORIGIN - 1];
      var dateStr = values[i][COL.DATE - 1];
      var sendTime = values[i][COL.SEND_TIME - 1];
      if (van && resolveMainBranch(dest) === branch && status === "In Transit") {
        var key = van + "|" + origin;
        if (!vans[key]) vans[key] = { vanNo: van, origin: origin, dateStr: dateStr, sendTime: sendTime, count: 0 };
        vans[key].count++;
      }
    }
    var result = [];
    for (var key in vans) {
      var v = vans[key];
      var sentAt = parseDateTime(v.dateStr, v.sendTime);
      var ms = now.getTime() - sentAt.getTime();
      if (isNaN(ms) || ms < 0) ms = 0;
      var min = Math.floor(ms / 60000);
      var sec = Math.floor((ms % 60000) / 1000);
      result.push({ vanNo: v.vanNo, origin: v.origin, count: v.count, elapsedMinutes: min, elapsedSeconds: sec, isLate: min >= 20 });
    }
    result.sort(function(a, b) { return (a.elapsedMinutes * 60 + a.elapsedSeconds) - (b.elapsedMinutes * 60 + b.elapsedSeconds); });
    return result;
  } catch(e) { return []; }
}

function handleSendBatch(data) {
  try {
    var sheet = getSheet();
    var values = getAllValues(sheet);
    var now = new Date();
    var tz = CONFIG.TIMEZONE;
    var dateStr = Utilities.formatDate(now, tz, "yyyy-MM-dd");
    var timeStr = Utilities.formatDate(now, tz, "HH:mm:ss");

    var inTransitIds = {};
    for (var i = 1; i < values.length; i++) {
      var sid = String(values[i][COL.ID - 1] || "").trim();
      var status = String(values[i][COL.STATUS - 1] || "").trim();
      if (sid && status === "In Transit") inTransitIds[sid] = values[i][COL.ORIGIN - 1];
    }

    var newRows = [], results = [];
    var vanNo = String(data.vanNo || "").trim();
    var origin = String(data.origin || "").toUpperCase().trim();
    var dest = String(data.destination || "").toUpperCase().trim();
    var shipments = data.shipments || [];
    var driver = getDriverState(vanNo);
    if (!vanNo || !driver.started || driver.status !== "AT_STATION" ||
        resolveMainBranch(driver.branch) !== resolveMainBranch(origin)) {
      return { success:false, error:"Select a van whose driver is checked in at " + origin };
    }

    for (var j = 0; j < shipments.length; j++) {
      var sid2 = String(shipments[j].shipmentId || "").trim();
      if (!sid2) continue;
      if (inTransitIds[sid2]) {
        results.push({ shipmentId: sid2, status: "duplicate", error: sid2 + " already In Transit from " + inTransitIds[sid2] });
      } else {
        newRows.push([dateStr, origin, dest, timeStr, sid2, vanNo, "In Transit", "", ""]);
        results.push({ shipmentId: sid2, status: "sent" });
        inTransitIds[sid2] = origin;
      }
    }

    if (newRows.length > 0) {
      var startRow = sheet.getLastRow() + 1;
      var needed = startRow + newRows.length - 1;
      if (sheet.getMaxRows() < needed) sheet.insertRowsAfter(sheet.getMaxRows(), needed - sheet.getMaxRows() + 5);
      sheet.getRange(startRow, 1, newRows.length, CONFIG.TOTAL_COLS).setValues(newRows);
    }
    return { success: true, sent: newRows.length, results: results };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function handleReceiveBatch(data) {
  try {
    var sheet = getSheet();
    var values = getAllValues(sheet);
    var timeStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "HH:mm:ss");
    var targetIds = {};
    (data.shipmentIds || []).forEach(function(id) { targetIds[String(id).trim()] = true; });
    var targetVan = String(data.vanNo).trim();
    var branch = data.branch;
    var receivedCount = 0, foundIds = {}, notFoundIds = [], modified = false;

    for (var i = 1; i < values.length; i++) {
      var sid = String(values[i][COL.ID - 1] || "").trim();
      var dest = values[i][COL.DEST - 1];
      var van = String(values[i][COL.VAN - 1] || "").trim();
      var status = String(values[i][COL.STATUS - 1] || "").trim();
      if (targetIds[sid] && resolveMainBranch(dest) === branch && van === targetVan && status === "In Transit" && !foundIds[sid]) {
        values[i][COL.STATUS - 1] = "Received";
        values[i][COL.RECV_BY - 1] = branch;
        values[i][COL.RECV_TIME - 1] = timeStr;
        receivedCount++; foundIds[sid] = true; modified = true;
      }
    }
    if (modified) setAllValues(sheet, values);
    (data.shipmentIds || []).forEach(function(id) {
      var idStr = String(id).trim();
      if (!foundIds[idStr]) notFoundIds.push(idStr);
    });
    return {
      success: true,
      message: receivedCount + " received" + (notFoundIds.length ? ", " + notFoundIds.length + " not found" : ""),
      receivedCount: receivedCount, notFoundIds: notFoundIds
    };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function handleReceiveOne(data) {
  try {
    var sheet = getSheet();
    var values = getAllValues(sheet);
    var timeStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "HH:mm:ss");
    for (var i = 1; i < values.length; i++) {
      var sid = String(values[i][COL.ID - 1] || "").trim();
      var dest = values[i][COL.DEST - 1];
      var van = String(values[i][COL.VAN - 1] || "").trim();
      var status = String(values[i][COL.STATUS - 1] || "").trim();
      if (sid === String(data.shipmentId || "").trim() && resolveMainBranch(dest) === data.branch &&
          van === String(data.vanNo).trim() && status === "In Transit") {
        values[i][COL.STATUS - 1] = "Received";
        values[i][COL.RECV_BY - 1] = data.branch;
        values[i][COL.RECV_TIME - 1] = timeStr;
        setAllValues(sheet, values);
        return { success: true, message: data.shipmentId + " received" };
      }
    }
    return { success: false, error: "Not found or already received" };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function handleReceiveAll(data) {
  try {
    var sheet = getSheet();
    var values = getAllValues(sheet);
    var count = 0, modified = false;
    var timeStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "HH:mm:ss");
    for (var i = 1; i < values.length; i++) {
      var dest = values[i][COL.DEST - 1];
      var van = String(values[i][COL.VAN - 1] || "").trim();
      var status = String(values[i][COL.STATUS - 1] || "").trim();
      if (resolveMainBranch(dest) === data.branch && van === String(data.vanNo).trim() && status === "In Transit") {
        values[i][COL.STATUS - 1] = "Received";
        values[i][COL.RECV_BY - 1] = data.branch;
        values[i][COL.RECV_TIME - 1] = timeStr;
        count++; modified = true;
      }
    }
    if (modified) setAllValues(sheet, values);
    return { success: true, message: count + " shipment(s) received", count: count };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function getPendingByVan(branch, vanNo) {
  var sheet = getSheet();
  var values = getAllValues(sheet);
  var result = [];
  var targetVan = String(vanNo).trim();
  for (var i = 1; i < values.length; i++) {
    var rowDest = values[i][COL.DEST - 1];
    var rowVan = String(values[i][COL.VAN - 1] || "").trim();
    var rowStatus = String(values[i][COL.STATUS - 1] || "").trim();
    if (resolveMainBranch(rowDest) === branch && rowVan === targetVan && rowStatus === "In Transit") {
      result.push({ shipmentId: values[i][COL.ID - 1], origin: values[i][COL.ORIGIN - 1], destination: rowDest });
    }
  }
  return result;
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
    // Upgrade an older-schema sheet in place, just in case.
    var headers = ["Date","Van No","Branch","Check In","Check Out","Hold Seconds","Hold Time",
      "Next Branch","Travel Seconds","Travel Time","Status","Round","Missed Branches","Last Updated"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  // NOTE: number formats are only set when the sheet is first created.
  // Setting formats on every call made even simple READS slow.
  return sheet;
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

// Walks today's rows for a van and figures out: what round are they on now,
// and — for the round that just closed (if any) — which core branches got
// visited vs missed. Skipping a branch is fine, it's just reported.
function computeRoundInfo(priorRows, arrivingBranch) {
  var roundNumber = 1;
  var visitedThisRound = {};
  var closedRound = null; // {round, missed[]} — only set if this check-in closes a round

  for (var i = 0; i < priorRows.length; i++) {
    var b = String(priorRows[i].values[MOVE_COL.BRANCH - 1] || "").toUpperCase().trim();
    if (b === "TINKUNE") {
      // A prior Tinkune row: if it's not the very first row, it closed a round.
      if (i > 0) roundNumber++;
      visitedThisRound = {};
    } else if (ROUND_CORE_BRANCHES.indexOf(b) !== -1) {
      visitedThisRound[b] = true;
    }
  }

  if (arrivingBranch === "TINKUNE" && priorRows.length > 0) {
    var missed = ROUND_CORE_BRANCHES.filter(function(b) { return !visitedThisRound[b]; });
    closedRound = { round: roundNumber, missed: missed };
    roundNumber++; // the round now starting
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
  var rows = movementRowsForToday(vanNo);
  var state;
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
    round: state.round || 1, lastMissed: state.lastMissed || ""
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
      invalidateDriverCaches(vanNo);
      return { success:true, message:"Checked in at " + branch, state:driverStateResponse(getDriverState(vanNo)) };
    }

    if (state.status === "AT_STATION") {
      // A retry can arrive after Google Sheets saved the original request
      // but the phone lost the response. Treat the same check-in as success
      // instead of rejecting it or creating another movement row.
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
      // Idempotent retry: the phone may have timed out after this checkout
      // was already written to the sheet.
      if (String(state.nextBranch || "").toUpperCase().trim() === nextBranch) {
        return { success:true, message:"Check-out already saved — heading to " + nextBranch,
          state:driverStateResponse(state) };
      }
      return { success:false, error:"Already heading to " + state.nextBranch };
    }
    if (state.status !== "AT_STATION") return { success:false, error:"Van is not at a station" };
    var currentBranch = String(state.branch || "").toUpperCase().trim();
    var allowedDestinations = DESTINATIONS[currentBranch] || DESTINATIONS[resolveMainBranch(currentBranch)] || [];
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
    invalidateDriverCaches(vanNo);

    return { success:true, message:"Checked out — heading to " + nextBranch, state:driverStateResponse(getDriverState(vanNo)) };
  } finally {
    lock.releaseLock();
  }
}

// FAST: single read of VAN MOVEMENTS, grouped by van in memory — used for
// the company-wide monitor board. No per-van re-reads, no second sheet.
function getVanBoard() {
  var dateStr = today();
  var cacheKey = 'van_board_raw_' + dateStr;
  var rows = null;
  var cached = getCached_(cacheKey);
  if (cached) {
    try { rows = JSON.parse(cached); } catch (e) { rows = null; }
  }
  if (!rows) {
    // ONE sheet read now serves every user for VAN_BOARD_TTL seconds.
    rows = [];
    var values = getMovementSheet().getDataRange().getValues();
    var latestByVan = {};
    for (var i = 1; i < values.length; i++) {
      var van = String(values[i][MOVE_COL.VAN - 1] || "").trim();
      if (!van) continue;
      if (normalizeDateStr(values[i][MOVE_COL.DATE - 1]) !== dateStr) continue;
      latestByVan[van] = values[i];
    }
    for (var vanNo in latestByVan) {
      var v = latestByVan[vanNo];
      var arr = v[MOVE_COL.ARRIVAL - 1];
      var dep = v[MOVE_COL.DEPARTURE - 1];
      rows.push({
        vanNo: vanNo,
        status: String(v[MOVE_COL.STATUS - 1] || ""),
        branch: v[MOVE_COL.BRANCH - 1],
        nextBranch: v[MOVE_COL.NEXT_BRANCH - 1],
        round: Number(v[MOVE_COL.ROUND - 1]) || 1,
        arrivalMs: arr instanceof Date ? arr.getTime() : 0,
        departureMs: dep instanceof Date ? dep.getTime() : 0
      });
    }
    setCached_(cacheKey, JSON.stringify(rows), VAN_BOARD_TTL);
  }
  // Elapsed time is computed HERE at response time, so cached data never
  // makes the on-screen timers freeze — they stay accurate to the second.
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
   Both are CACHED like the van board (one read serves everyone)
   and every write invalidates the cache instantly.
═══════════════════════════════════════════════════════ */

const INFO_TTL = 60; // seconds

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

// One announcement per branch per day — posting again EDITS it.
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

// One contact row per branch — saving again EDITS it.
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

  if (action === "getPendingByVan") {
    return jsonResponse({ success: true, data: getPendingByVan(e.parameter.branch, e.parameter.vanNo) });
  }
  if (action === "getIncomingVans") {
    return jsonResponse({ success: true, data: getIncomingVans(e.parameter.branch) });
  }
  if (action === "getDriverState") {
    return jsonResponse({ success:true, state:driverStateResponse(getDriverState(e.parameter.vanNo)) });
  }
  if (action === "getDestinations") {
    var requestedBranch = String(e.parameter.branch || "").toUpperCase().trim();
    // Driver-only stops must keep their own route options. Only shipment
    // sub-branches should be resolved back to their main branch.
    var branch = DESTINATIONS[requestedBranch] ? requestedBranch : resolveMainBranch(requestedBranch);
    return jsonResponse({ success:true, destinations: DESTINATIONS[branch] || [], mainBranches: MAIN_BRANCHES });
  }
  if (action === "getVanBoard") {
    return jsonResponse({ success:true, data:getVanBoard(), date:today() });
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
    case 'sendBatch':       return jsonResponse(handleSendBatch(body));
    case 'receiveBatch':    return jsonResponse(handleReceiveBatch(body));
    case 'receiveOne':      return jsonResponse(handleReceiveOne(body));
    case 'receiveAll':      return jsonResponse(handleReceiveAll(body));
    case 'driverCheckIn':   return jsonResponse(handleDriverCheckIn(body));
    case 'driverCheckOut':  return jsonResponse(handleDriverCheckOut(body));
    case 'submitIssue':     return jsonResponse(submitIssue(body));
    case 'putAnnouncement':   return jsonResponse(handlePutAnnouncement(body));
    case 'saveBranchContact': return jsonResponse(handleSaveBranchContact(body));
    case 'closeIssue':      return jsonResponse(closeIssue(body.row));
    default:                return jsonResponse({ success: false, error: 'Unknown action' });
  }
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
