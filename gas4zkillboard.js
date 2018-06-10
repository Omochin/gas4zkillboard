function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu(
    "zKillboard",
    [
      {
        name: "シンプル",
        functionName: "simpleKillboard"
      },
      {
        name: "詳細",
        functionName: "detailedKillboard"
      }
    ]
  );
}


function simpleKillboard() {
  killboard(
    inputAPI(),
    [
      "Region",
      "System",
      "Ship",
      "Character",
      "CorporationTicker",
      "AllianceTicker"
    ]
  );
}


function detailedKillboard() {
  killboard(
    inputAPI(),
    [
      "KillID",
      "Date",
      "Security",
      "Region",
      "System",
      "Ship",
      "Damage",
      "Value",
      "Points",
      "Involved",
      "Character",
      "Corporation",
      "Alliance"
    ]
  );
}


function parseJSON(url) {
  var response = UrlFetchApp.fetch(url);
  return JSON.parse(response.getContentText());
}

var typeIDs = parseJSON("https://raw.githubusercontent.com/Omochin/gas4zkillboard/master/type_ids.json");
var regionIDs = parseJSON("https://raw.githubusercontent.com/Omochin/gas4zkillboard/master/region_ids.json");
var solarSystemIDs = parseJSON("https://raw.githubusercontent.com/Omochin/gas4zkillboard/master/solar_system_ids.json");

var charactersSheet = undefined;
var corporationsSheet = undefined;
var alliancesSheet = undefined;

var characterIDs = undefined;
var corporationIDs = undefined;
var allianceIDs = undefined;

var elementGetterFunctions = {};

elementGetterFunctions["KillID"] = function(killmail) {
  var killID = killmail["killmail_id"];
  return '=HYPERLINK("https://zkillboard.com/kill/' + killID + '/","' + killID + '")';
}

elementGetterFunctions["Date"] = function(killmail) {
  var date = new Date(killmail["killmail_time"]);
  return date;
}

elementGetterFunctions["Security"] = function(killmail) {
  var solar_system_id = killmail["solar_system_id"];
  return solarSystemIDs[solar_system_id][1];
}

elementGetterFunctions["Region"] = function(killmail) {
  var solar_system_id = killmail["solar_system_id"];
  var region_id = solarSystemIDs[solar_system_id][2];
  return regionIDs[region_id];
}

elementGetterFunctions["System"] = function(killmail) {
  var solar_system_id = killmail["solar_system_id"];
  return solarSystemIDs[solar_system_id][0];
}

elementGetterFunctions["Ship"] = function(killmail) {
  var type_id = killmail["victim"]["ship_type_id"];
  return typeIDs[type_id];
}

elementGetterFunctions["Damage"] = function(killmail) {
  return killmail["victim"]["damage_taken"];
}

elementGetterFunctions["Value"] = function(killmail) {
  var value = parseFloat(killmail["zkb"]["totalValue"]).toFixed(2);
  return String(value).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,');
}

elementGetterFunctions["Points"] = function(killmail) {
  return killmail["zkb"]["points"];
}

elementGetterFunctions["Involved"] = function(killmail) {
  return killmail["attackers"].length;
}

function getLink(targetID, targetType) {
  var url = "https://zkillboard.com/" + targetType + "/" + targetID + "/";
  return '=HYPERLINK("' + url + '","' + targetID + '")';
}
function getElement(targetID, targetType, targetIDs, targetSheet, targetKeys) {
  var element = undefined;

  if (targetID in targetIDs) {
    element = targetIDs[targetID];
  } else {
    var targetRow = targetSheet.getLastRow() + 1;
    var target = parseJSON("https://esi.evetech.net/" + targetType + "/" + targetID + "/");

    element = [];
    targetSheet.getRange(targetRow, 1).setValue(targetID);
    for (var i = 0; i < targetKeys.length; i++) {
      var value = target[targetKeys[i]];
      targetSheet.getRange(targetRow, i + 2).setValue(value);
      element.push(value);
    }
    targetIDs[targetID] = element;
  }
  
  return element;
}
function getCharacter(killmail) {
  return getElement(
    killmail['victim']['character_id'],
    "v4/characters",
    characterIDs,
    charactersSheet,
    ["name"]
  );
}
function getCorporation(killmail) {
  return getElement(
    killmail['victim']['corporation_id'],
    "v4/corporations",
    corporationIDs,
    corporationsSheet,
    ["name", "ticker"]
  );
}
function getAlliance(killmail) {
  return getElement(
    killmail['victim']['alliance_id'],
    "v3/alliances",
    allianceIDs,
    alliancesSheet,
    ["name", "ticker"]
  )
}

elementGetterFunctions["CharacterID"] = function(killmail) {
  return getLink(killmail['victim']['character_id'], "character");
}
elementGetterFunctions["Character"] = function(killmail) {
  if ('character_id' in killmail['victim']) {
    return getCharacter(killmail)[0];
  } else {
    return "";
  }
}

elementGetterFunctions["CorporationID"] = function(killmail) {
  return getLink(killmail['victim']['corporation_id'], "corporation");
}
elementGetterFunctions["CorporationTicker"] = function(killmail) {
  return getCorporation(killmail)[1];
}
elementGetterFunctions["Corporation"] = function(killmail) {
  return getCorporation(killmail)[0];
}

elementGetterFunctions["AllianceID"] = function(killmail) {
  if ('alliance_id' in killmail['victim']) {
    return getLink(killmail['victim']['alliance_id'], "alliance");
  } else {
    return "";
  }
}
elementGetterFunctions["AllianceTicker"] = function(killmail) {
  if ('alliance_id' in killmail['victim']) {
    return getAlliance(killmail)[1];
  } else {
    return "";
  }
}
elementGetterFunctions["Alliance"] = function(killmail) {
  if ('alliance_id' in killmail['victim']) {
    return getAlliance(killmail)[0];
  } else {
    return "";
  }
}


function inputAPI() {
  var api = {}
  api["url"] = Browser.inputBox("URLを入力してください");

  var items = api["url"].split("/");
  var page = 1;
  for (var i = 0; i < items.length - 1; i++) {
    if (items[i] == "page") {
      value = items[i + 1];
      if (isNaN(value) || value < 1) {
        page = 1;
      } else {
        page = value;
      }
      break;
    }
  }
  api["page"] = parseInt(page);

  var limit = Browser.inputBox("取得するページ数を入力してください（デフォルト値は1）");
  if (isNaN(limit) || limit < 1) {
    limit = 1;
  }
  api["limit"] = parseInt(limit);
  
  return api;
}

function getSheet(name, spreadsheet) {
  var sheet = undefined;
  var sheets = spreadsheet.getSheets();

  for (var i in sheets) {
    if (sheets[i].getName() == name) {
      sheet = sheets[i];
      break;
    }
  }

  if (sheet) {
    return sheet;
  } else {
    return spreadsheet.insertSheet(name);
  }
}

function getIDs(sheet) {
  var targetIDs = {};
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  
  for (var row = 1; row <= lastRow; row++) {
    var targetID = sheet.getRange(row, 1).getValue();
    targetIDs[targetID] = [];
    for (var column = 2; column <= lastColumn; column++) {
      var value = sheet.getRange(row, column).getValue();
      targetIDs[targetID].push(value);
    }
  }
  return targetIDs;
}

function killboard(api, items) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();  

  charactersSheet = getSheet("Characters", spreadsheet);
  corporationsSheet = getSheet("Corporations", spreadsheet);
  alliancesSheet = getSheet("Alliances", spreadsheet);

  if(!characterIDs) { characterIDs = getIDs(charactersSheet); }
  if(!corporationIDs) { corporationIDs = getIDs(corporationsSheet); }
  if(!allianceIDs) { allianceIDs = getIDs(alliancesSheet); }

  sheet.activate();
  sheet.clear();
  for (var i = 0; i < items.length; i++) {
    sheet.getRange(1, 1 + i).setValue(items[i]);
  }
  
  var split = api["url"].split("/");
  split.splice(3, 0, "api");
  var url = split.join("/");
  var row = 2;
  for (var limit = 0; limit < api["limit"]; limit++) {
    var killmails = parseJSON(url + "page/" + (api["page"] + limit) + "/");
    for (var i = 0; i < killmails.length; i++) {
      for (var column = 0; column < items.length; column++) {
        try {
          value = elementGetterFunctions[items[column]](killmails[i]);
        } catch(error) {
          value = "Unknown";
        }
        sheet.getRange(row, column + 1).setValue(value);
      }
      row++;
    }
  }
}