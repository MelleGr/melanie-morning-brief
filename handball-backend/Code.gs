/**
 * Kampkalenderen — backend for the handball availability tool.
 * Deploy this as a Google Apps Script Web App bound to a Google Sheet.
 * See handball-backend/README.md for the one-time deploy steps.
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

function checkAdmin_(body){
  var key = PropertiesService.getScriptProperties().getProperty("ADMIN_KEY");
  return !!key && body.adminKey === key;
}

function jsonOut_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e){
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
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

// Body is sent as text/plain (JSON string) to avoid a CORS preflight that
// Apps Script web apps cannot answer. Parse it manually.
function doPost(e){
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    var body = {};
    try{ body = JSON.parse(e.postData.contents); }catch(err){ return jsonOut_({ ok:false, error:"bad_request" }); }
    var action = body.action;

    if(action === "saveProfile"){
      var p = body.payload || {};
      if(!p.id || !p.name || !p.position) return jsonOut_({ ok:false, error:"invalid_argument" });
      upsertRow_("players", p.id, {
        name: String(p.name).slice(0,80),
        position: String(p.position),
        detail: String(p.detail||"").slice(0,120),
        updatedAt: new Date().toISOString()
      });
      return jsonOut_({ ok:true });
    }

    if(action === "deletePlayer"){
      if(!body.payload || !body.payload.id) return jsonOut_({ ok:false, error:"invalid_argument" });
      deleteRow_("players", body.payload.id);
      return jsonOut_({ ok:true });
    }

    if(action === "setAvailability"){
      var a = body.payload || {};
      if(!a.matchId || !a.playerId) return jsonOut_({ ok:false, error:"invalid_argument" });
      var id = a.matchId + "__" + a.playerId;
      if(a.status === null){
        deleteRow_("availability", id);
      } else {
        upsertRow_("availability", id, {
          matchId: a.matchId,
          playerId: a.playerId,
          playerName: String(a.playerName||"").slice(0,80),
          position: String(a.position||""),
          status: String(a.status),
          updatedAt: new Date().toISOString()
        });
      }
      return jsonOut_({ ok:true });
    }

    if(action === "addMatch"){
      if(!checkAdmin_(body)) return jsonOut_({ ok:false, error:"not_admin" });
      var m = body.payload || {};
      if(!m.date || !m.opponent) return jsonOut_({ ok:false, error:"invalid_argument" });
      var mid = newId_("m");
      upsertRow_("matches", mid, {
        date: m.date, time: m.time||"", opponent: String(m.opponent).slice(0,120),
        location: String(m.location||"").slice(0,120), minPlayers: m.minPlayers||12,
        dateChangedAt: "", previousDate: "", createdAt: new Date().toISOString()
      });
      return jsonOut_({ ok:true, id: mid });
    }

    if(action === "updateMatch"){
      if(!checkAdmin_(body)) return jsonOut_({ ok:false, error:"not_admin" });
      var um = body.payload || {};
      if(!um.id || !um.date || !um.opponent) return jsonOut_({ ok:false, error:"invalid_argument" });
      var sheet = getSheet_("matches");
      var idx = findRowIndex_(sheet, SHEETS.matches, um.id);
      if(idx < 0) return jsonOut_({ ok:false, error:"not_found" });
      var existingVals = sheet.getRange(idx, 1, 1, SHEETS.matches.cols.length).getValues()[0];
      var existingDate = existingVals[1];
      var patch = {
        date: um.date, time: um.time||"", opponent: String(um.opponent).slice(0,120),
        location: String(um.location||"").slice(0,120), minPlayers: um.minPlayers||12
      };
      if(existingDate && String(existingDate) !== String(um.date)){
        patch.dateChangedAt = new Date().toISOString();
        patch.previousDate = existingDate;
      }
      upsertRow_("matches", um.id, patch);
      return jsonOut_({ ok:true });
    }

    if(action === "deleteMatch"){
      if(!checkAdmin_(body)) return jsonOut_({ ok:false, error:"not_admin" });
      if(!body.payload || !body.payload.id) return jsonOut_({ ok:false, error:"invalid_argument" });
      deleteRow_("matches", body.payload.id);
      return jsonOut_({ ok:true });
    }

    return jsonOut_({ ok:false, error:"unknown_action" });
  } finally {
    lock.releaseLock();
  }
}

/** Run this once from the Apps Script editor (Run > setAdminKey) to set your
 * holdleder passphrase. Change YOUR_SECRET_HERE first. */
function setAdminKey(){
  PropertiesService.getScriptProperties().setProperty("ADMIN_KEY", "YOUR_SECRET_HERE");
}
