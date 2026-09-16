/**
 * Kampkalenderen — backend for the handball availability tool.
 * Deploy this as a Google Apps Script Web App bound to a Google Sheet.
 * See handball-backend/README.md for the one-time deploy steps.
 *
 * Writes travel as GET requests with the data in the query string, not
 * POST. Apps Script Web Apps answer a cross-origin request with a redirect,
 * and the redirect step turns a POST into a GET and drops its body before
 * it ever reaches doPost — so a POST-based write silently vanishes. GET
 * requests don't have that problem, so every write goes through doGet too.
 */

var SHEETS = {
  matches: { name: "Matches", cols: ["id","date","time","opponent","location","minPlayers","dateChangedAt","previousDate","createdAt"] },
  players: { name: "Players", cols: ["id","name","position","detail","updatedAt"] },
  availability: { name: "Availability", cols: ["id","matchId","playerId","playerName","position","status","updatedAt"] }
};

function getSheet_(key){
  var def = SHEETS[key];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(def.name);
  if(!sheet){
    sheet = ss.insertSheet(def.name);
    sheet.appendRow(def.cols);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readAll_(key){
  var def = SHEETS[key];
  var sheet = getSheet_(key);
  var range = sheet.getDataRange().getValues();
  var rows = [];
  for(var i = 1; i < range.length; i++){
    var row = range[i];
    if(!row[0]) continue;
    var obj = {};
    for(var c = 0; c < def.cols.length; c++) obj[def.cols[c]] = row[c] === undefined ? "" : row[c];
    rows.push(obj);
  }
  return rows;
}

function findRowIndex_(sheet, def, id){
  var range = sheet.getDataRange().getValues();
  for(var i = 1; i < range.length; i++){
    if(String(range[i][0]) === String(id)) return i + 1; // 1-based sheet row
  }
  return -1;
}

function upsertRow_(key, id, data){
  var def = SHEETS[key];
  var sheet = getSheet_(key);
  var rowIdx = findRowIndex_(sheet, def, id);
  var existing = {};
  if(rowIdx > 0){
    var vals = sheet.getRange(rowIdx, 1, 1, def.cols.length).getValues()[0];
    for(var c = 0; c < def.cols.length; c++) existing[def.cols[c]] = vals[c];
  }
  var merged = Object.assign({}, existing, data, { id: id });
  var out = def.cols.map(function(c){ return merged[c] === undefined ? "" : merged[c]; });
  if(rowIdx > 0){
    sheet.getRange(rowIdx, 1, 1, def.cols.length).setValues([out]);
  } else {
    sheet.appendRow(out);
  }
}

function deleteRow_(key, id){
  var def = SHEETS[key];
  var sheet = getSheet_(key);
  var rowIdx = findRowIndex_(sheet, def, id);
  if(rowIdx > 0) sheet.deleteRow(rowIdx);
}

function newId_(prefix){
  return prefix + "_" + Utilities.getUuid().replace(/-/g, "").slice(0, 12);
}

function checkAdmin_(req){
  var key = PropertiesService.getScriptProperties().getProperty("ADMIN_KEY");
  return !!key && req.adminKey === key;
}

function jsonOut_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** req = { action, payload, adminKey } with payload already a plain object. */
function handleAction_(req){
  var action = req.action;
  var payload = req.payload || {};

  if(action === "saveProfile"){
    if(!payload.id || !payload.name || !payload.position) return jsonOut_({ ok:false, error:"invalid_argument" });
    upsertRow_("players", payload.id, {
      name: String(payload.name).slice(0,80),
      position: String(payload.position),
      detail: String(payload.detail||"").slice(0,120),
      updatedAt: new Date().toISOString()
    });
    return jsonOut_({ ok:true });
  }

  if(action === "deletePlayer"){
    if(!payload.id) return jsonOut_({ ok:false, error:"invalid_argument" });
    deleteRow_("players", payload.id);
    return jsonOut_({ ok:true });
  }

  if(action === "setAvailability"){
    if(!payload.matchId || !payload.playerId) return jsonOut_({ ok:false, error:"invalid_argument" });
    var id = payload.matchId + "__" + payload.playerId;
    if(payload.status === null || payload.status === undefined){
      deleteRow_("availability", id);
    } else {
      upsertRow_("availability", id, {
        matchId: payload.matchId,
        playerId: payload.playerId,
        playerName: String(payload.playerName||"").slice(0,80),
        position: String(payload.position||""),
        status: String(payload.status),
        updatedAt: new Date().toISOString()
      });
    }
    return jsonOut_({ ok:true });
  }

  if(action === "addMatch"){
    if(!checkAdmin_(req)) return jsonOut_({ ok:false, error:"not_admin" });
    if(!payload.date || !payload.opponent) return jsonOut_({ ok:false, error:"invalid_argument" });
    var mid = newId_("m");
    upsertRow_("matches", mid, {
      date: payload.date, time: payload.time||"", opponent: String(payload.opponent).slice(0,120),
      location: String(payload.location||"").slice(0,120), minPlayers: payload.minPlayers||12,
      dateChangedAt: "", previousDate: "", createdAt: new Date().toISOString()
    });
    return jsonOut_({ ok:true, id: mid });
  }

  if(action === "updateMatch"){
    if(!checkAdmin_(req)) return jsonOut_({ ok:false, error:"not_admin" });
    if(!payload.id || !payload.date || !payload.opponent) return jsonOut_({ ok:false, error:"invalid_argument" });
    var sheet = getSheet_("matches");
    var idx = findRowIndex_(sheet, SHEETS.matches, payload.id);
    if(idx < 0) return jsonOut_({ ok:false, error:"not_found" });
    var existingVals = sheet.getRange(idx, 1, 1, SHEETS.matches.cols.length).getValues()[0];
    var existingDate = existingVals[1];
    var patch = {
      date: payload.date, time: payload.time||"", opponent: String(payload.opponent).slice(0,120),
      location: String(payload.location||"").slice(0,120), minPlayers: payload.minPlayers||12
    };
    if(existingDate && String(existingDate) !== String(payload.date)){
      patch.dateChangedAt = new Date().toISOString();
      patch.previousDate = existingDate;
    }
    upsertRow_("matches", payload.id, patch);
    return jsonOut_({ ok:true });
  }

  if(action === "deleteMatch"){
    if(!checkAdmin_(req)) return jsonOut_({ ok:false, error:"not_admin" });
    if(!payload.id) return jsonOut_({ ok:false, error:"invalid_argument" });
    deleteRow_("matches", payload.id);
    return jsonOut_({ ok:true });
  }

  return jsonOut_({ ok:false, error:"unknown_action" });
}

function doGet(e){
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    var params = (e && e.parameter) || {};
    if(params.action){
      var payload = {};
      try{ payload = params.payload ? JSON.parse(params.payload) : {}; }
      catch(err){ return jsonOut_({ ok:false, error:"bad_request" }); }
      return handleAction_({ action: params.action, payload: payload, adminKey: params.adminKey || "" });
    }
    return jsonOut_({
      ok: true,
      matches: readAll_("matches"),
      players: readAll_("players"),
      availability: readAll_("availability")
    });
  } finally {
    lock.releaseLock();
  }
}

// Kept for completeness; the page itself calls doGet with an action
// parameter (see the file header for why POST isn't used for writes).
function doPost(e){
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    var body = {};
    try{ body = JSON.parse(e.postData.contents); }catch(err){ return jsonOut_({ ok:false, error:"bad_request" }); }
    return handleAction_(body);
  } finally {
    lock.releaseLock();
  }
}

/** Run this once from the Apps Script editor (Run > setAdminKey) to set your
 * holdleder passphrase. Change YOUR_SECRET_HERE first. */
function setAdminKey(){
  PropertiesService.getScriptProperties().setProperty("ADMIN_KEY", "YOUR_SECRET_HERE");
}
