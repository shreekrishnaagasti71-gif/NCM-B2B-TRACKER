/*
 * NCM BRANCH OPERATIONS — Google Apps Script backend
 *
 * Paste this file into Apps Script as Code.js.
 * Deploy as a Web app, execute as the owner, and allow the users who need it.
 *
 * Sheets created automatically:
 *   ANNOUNCEMENTS, BRANCH_CONTACTS, VAN_CURRENT_STATE, VAN_MOVEMENT_HISTORY
 *
 * The seven branches intentionally share the same routing rule.
 * Change PASSCODES before production, or store a JSON object in the
 * Script Property NCM_PASSCODES.
 */

const CONFIG = {
  TIMEZONE: "Asia/Kathmandu",
  SESSION_SECONDS: 21600,
  ANNOUNCEMENT_HOURS: 24,
  BRANCHES: ["NAYA BUSPARK","BASUNDHARA","CHABAHIL","TINKUNE","SATDOBATO","KALANKI","SWOYAMBHU"]
};

const DEFAULT_PASSCODES = {
  "1111": { branch:"TINKUNE", role:"branch" },
  "2222": { branch:"CHABAHIL", role:"branch" },
  "3333": { branch:"NAYA BUSPARK", role:"branch" },
  "4444": { branch:"KALANKI", role:"branch" },
  "5555": { branch:"SATDOBATO", role:"branch" },
  "6666": { branch:"BASUNDHARA", role:"branch" },
  "8888": { branch:"SWOYAMBHU", role:"branch" },
  "7777": { role:"driver" },
  "9999": { role:"admin" }
};

const SHEETS = {
  ANNOUNCEMENTS: ["ID","Branch","Message","Created By","Created At","Expires At"],
  CONTACTS: ["Branch","Name","Phone","Alternative Phone","Note","Updated By","Updated At"],
  STATE: ["Van No","Current Branch","Next Branch","Status","Round","Arrival At","Departure At","Updated At","Version"],
  HISTORY: ["ID","Van No","From Branch","To Branch","Check In","Check Out","Hold Seconds","Travel Seconds","Status","Round","Created At"]
};

function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function now_() { return new Date(); }
function iso_(value) { return value instanceof Date && !isNaN(value) ? value.toISOString() : (value ? new Date(value).toISOString() : null); }
function today_() { return Utilities.formatDate(now_(), CONFIG.TIMEZONE, "yyyy-MM-dd"); }
function seconds_(start, end) {
  var a = start instanceof Date ? start : new Date(start), b = end instanceof Date ? end : new Date(end);
  if (isNaN(a) || isNaN(b)) return 0;
  return Math.max(0, Math.floor((b.getTime() - a.getTime()) / 1000));
}
function durationLabel_(seconds) {
  seconds = Math.max(0, Number(seconds) || 0);
  var h = Math.floor(seconds / 3600), m = Math.floor((seconds % 3600) / 60);
  return h ? h + " hr " + m + " min" : m + " min";
}
function cleanBranch_(value) {
  var v = String(value || "").toUpperCase().trim().replace(/\s+/g, " ");
  if (v === "CHABHIL") v = "CHABAHIL";
  return v;
}
function validBranch_(branch) { return CONFIG.BRANCHES.indexOf(cleanBranch_(branch)) !== -1; }
function cleanText_(value, max) { return String(value == null ? "" : value).trim().slice(0, max || 500); }

function passcodes_() {
  var raw = PropertiesService.getScriptProperties().getProperty("NCM_PASSCODES");
  if (!raw) return DEFAULT_PASSCODES;
  try {
    var parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : DEFAULT_PASSCODES;
  } catch (e) { return DEFAULT_PASSCODES; }
}

function getSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastColumn() < headers.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}
function rows_(sheet) {
  var last = sheet.getLastRow();
  return last < 2 ? [] : sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
}
function sessionKey_(token) { return "ncm_session_" + token; }
function issueSession_(config) {
  var token = Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, "");
  CacheService.getScriptCache().put(sessionKey_(token), JSON.stringify(config), CONFIG.SESSION_SECONDS);
  return token;
}
function auth_(token) {
  if (!token) return null;
  var raw = CacheService.getScriptCache().get(sessionKey_(String(token)));
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}
function requireAuth_(token, roles) {
  var user = auth_(token);
  if (!user) throw new Error("Session expired. Please login again.");
  if (roles && roles.indexOf(user.role) === -1) throw new Error("You do not have permission for this action.");
  return user;
}

function doGet(e) {
  var p = e && e.parameter || {}, action = p.action;
  try {
    if (action === "getAnnouncements") return json_({success:true, data:getAnnouncements_()});
    if (action === "getContacts") { requireAuth_(p.token, ["branch","driver","admin"]); return json_({success:true, data:getContacts_()}); }
    if (action === "getDriverState") {
      requireAuth_(p.token, ["driver","admin"]);
      return json_({success:true, state:driverStateResponse_(getState_(cleanText_(p.vanNo, 30)))});
    }
    return json_({success:false, error:"Unknown action"});
  } catch (err) { return json_({success:false, error:err.message || String(err)}); }
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents || "{}"); } catch (err) { return json_({success:false,error:"Bad JSON"}); }
  try {
    switch (body.action) {
      case "login": return json_(login_(body.passcode));
      case "logout": CacheService.getScriptCache().remove(sessionKey_(body.token)); return json_({success:true});
      case "createAnnouncement": return json_(createAnnouncement_(body));
      case "saveContact": return json_(saveContact_(body));
      case "driverCheckIn": return json_(driverCheckIn_(body));
      case "driverCheckOut": return json_(driverCheckOut_(body));
      default: return json_({success:false,error:"Unknown action"});
    }
  } catch (err) { return json_({success:false,error:err.message || String(err)}); }
}

function login_(passcode) {
  var config = passcodes_()[String(passcode || "").trim()];
  if (!config) return {success:false,error:"Invalid passcode"};
  var session = {role:String(config.role), branch:config.branch || null, issuedAt:now_().toISOString()};
  session.token = issueSession_(session);
  return {success:true, session:session};
}

function getAnnouncements_() {
  var sheet = getSheet_("ANNOUNCEMENTS", SHEETS.ANNOUNCEMENTS), values = rows_(sheet), now = now_(), result = [];
  for (var i = 0; i < values.length; i++) {
    var expires = values[i][5] instanceof Date ? values[i][5] : new Date(values[i][5]);
    if (isNaN(expires) || expires <= now) continue;
    result.push({id:String(values[i][0]), branch:values[i][1], message:values[i][2], createdBy:values[i][3], createdAt:iso_(values[i][4]), expiresAt:iso_(expires)});
  }
  return result.reverse();
}

function createAnnouncement_(data) {
  var user = requireAuth_(data.token, ["branch","admin"]), branch = cleanBranch_(data.branch), message = cleanText_(data.message, 500);
  if (!validBranch_(branch)) return {success:false,error:"Choose one of the seven valid branches."};
  if (user.role === "branch" && user.branch !== branch) return {success:false,error:"A branch user can only post for their own branch."};
  if (!message) return {success:false,error:"Write an announcement first."};
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var created = now_(), expires = new Date(created.getTime() + CONFIG.ANNOUNCEMENT_HOURS * 3600000);
    getSheet_("ANNOUNCEMENTS", SHEETS.ANNOUNCEMENTS).appendRow([
      Utilities.getUuid(), branch, message, user.branch || "ADMIN", created, expires
    ]);
    return {success:true,message:"Announcement published for 24 hours",expiresAt:expires.toISOString()};
  } finally { lock.releaseLock(); }
}

function getContacts_() {
  var sheet = getSheet_("BRANCH_CONTACTS", SHEETS.CONTACTS), values = rows_(sheet), byBranch = {};
  for (var i = 0; i < values.length; i++) {
    var branch = cleanBranch_(values[i][0]);
    if (!validBranch_(branch)) continue;
    byBranch[branch] = {branch:branch,name:values[i][1] || "",phone:values[i][2] || "",altPhone:values[i][3] || "",note:values[i][4] || "",updatedBy:values[i][5] || "",updatedAt:iso_(values[i][6])};
  }
  return CONFIG.BRANCHES.map(function(branch) { return byBranch[branch] || {branch:branch,name:"",phone:"",altPhone:"",note:""}; });
}

function saveContact_(data) {
  var user = requireAuth_(data.token, ["branch","admin"]), branch = cleanBranch_(data.branch);
  if (!validBranch_(branch)) return {success:false,error:"Choose a valid branch."};
  if (user.role === "branch" && user.branch !== branch) return {success:false,error:"You can only edit your own branch contact."};
  var name = cleanText_(data.name,80), phone = cleanText_(data.phone,30), alt = cleanText_(data.altPhone,30), note = cleanText_(data.note,240);
  if (!name || !phone) return {success:false,error:"Name and phone are required."};
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var sheet = getSheet_("BRANCH_CONTACTS", SHEETS.CONTACTS), values = rows_(sheet), row = -1;
    for (var i=0; i<values.length; i++) if (cleanBranch_(values[i][0]) === branch) row = i + 2;
    var record = [branch,name,phone,alt,note,user.branch || "ADMIN",now_()];
    if (row === -1) sheet.appendRow(record); else sheet.getRange(row,1,1,record.length).setValues([record]);
    return {success:true,message:"Contact saved / corrected"};
  } finally { lock.releaseLock(); }
}

function stateHeader_() { return getSheet_("VAN_CURRENT_STATE", SHEETS.STATE); }
function getState_(vanNo) {
  vanNo = cleanText_(vanNo,30); if (!vanNo) return null;
  var values = rows_(stateHeader_());
  for (var i=0; i<values.length; i++) if (String(values[i][0]).trim() === vanNo) {
    return {row:i+2,vanNo:vanNo,currentBranch:cleanBranch_(values[i][1]),nextBranch:cleanBranch_(values[i][2]),status:String(values[i][3] || "AT_STATION"),round:Number(values[i][4]) || 1,arrivalAt:values[i][5],departureAt:values[i][6],updatedAt:values[i][7],version:Number(values[i][8]) || 0};
  }
  return null;
}
function writeState_(state) {
  var sheet = stateHeader_(), record = [state.vanNo,state.currentBranch,state.nextBranch || "",state.status,state.round,state.arrivalAt || "",state.departureAt || "",now_(),(state.version || 0) + 1];
  if (state.row) sheet.getRange(state.row,1,1,record.length).setValues([record]); else sheet.appendRow(record);
  state.version = record[8]; state.updatedAt = record[7]; return state;
}
function driverStateResponse_(state) {
  if (!state) return {vanNo:"",started:false,status:"NOT_STARTED",branch:null,nextBranch:null,round:0,elapsedSeconds:0};
  var elapsed = state.status === "MOVING" ? seconds_(state.departureAt,now_()) : seconds_(state.arrivalAt,now_());
  return {vanNo:state.vanNo,started:true,status:state.status,branch:state.currentBranch,nextBranch:state.nextBranch || null,round:state.round,elapsedSeconds:elapsed,arrivalAt:iso_(state.arrivalAt),departureAt:iso_(state.departureAt),version:state.version};
}
function validDestination_(current, next) { return validBranch_(next) && cleanBranch_(current) !== cleanBranch_(next); }
function historySheet_() { return getSheet_("VAN_MOVEMENT_HISTORY", SHEETS.HISTORY); }
function appendHistory_(record) { historySheet_().appendRow(record); }

function driverCheckIn_(data) {
  var user = requireAuth_(data.token, ["driver","admin"]), van = cleanText_(data.vanNo,30), branch = cleanBranch_(data.branch);
  if (!van || !validBranch_(branch)) return {success:false,error:"Enter a van number and choose a valid branch."};
  var lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    var state = getState_(van), now = now_();
    if (!state) {
      state = {vanNo:van,currentBranch:branch,nextBranch:"",status:"AT_STATION",round:1,arrivalAt:now,departureAt:"",version:0};
      writeState_(state); appendHistory_([Utilities.getUuid(),van,"",branch,now,"",0,0,"AT_STATION",1,now]);
      return {success:true,message:"Checked in at " + branch,state:driverStateResponse_(state)};
    }
    if (state.status === "AT_STATION") {
      if (state.currentBranch === branch) return {success:true,message:"Check-in already saved at " + branch,state:driverStateResponse_(state)};
      return {success:false,error:"Already checked in at " + state.currentBranch + ". Check out first."};
    }
    if (state.nextBranch !== branch) return {success:false,error:"This van is heading to " + state.nextBranch + "."};
    var previousBranch = state.currentBranch;
    var travel = seconds_(state.departureAt,now), round = previousBranch !== "TINKUNE" && branch === "TINKUNE" ? state.round + 1 : state.round;
    state.currentBranch = branch; state.nextBranch = ""; state.status = "AT_STATION"; state.arrivalAt = now; state.departureAt = ""; state.round = round;
    writeState_(state); appendHistory_([Utilities.getUuid(),van,previousBranch,branch,now,"",0,travel,"AT_STATION",round,now]);
    return {success:true,message:"Checked in at " + branch,state:driverStateResponse_(state)};
  } finally { lock.releaseLock(); }
}

function driverCheckOut_(data) {
  requireAuth_(data.token, ["driver","admin"]);
  var van = cleanText_(data.vanNo,30), next = cleanBranch_(data.nextBranch);
  if (!van || !validBranch_(next)) return {success:false,error:"Type one of the seven valid branch names."};
  var lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    var state = getState_(van), now = now_();
    if (!state) return {success:false,error:"Check in first."};
    if (state.status === "MOVING") {
      if (state.nextBranch === next) return {success:true,message:"Check-out already saved — heading to " + next,state:driverStateResponse_(state)};
      return {success:false,error:"This van is already heading to " + state.nextBranch + "."};
    }
    if (!validDestination_(state.currentBranch,next)) return {success:false,error:"Choose a different valid branch."};
    var hold = seconds_(state.arrivalAt,now);
    state.nextBranch = next; state.status = "MOVING"; state.departureAt = now; writeState_(state);
    appendHistory_([Utilities.getUuid(),van,state.currentBranch,next,state.arrivalAt,now,hold,0,"MOVING",state.round,now]);
    return {success:true,message:"Checked out — heading to " + next,state:driverStateResponse_(state)};
  } finally { lock.releaseLock(); }
}
