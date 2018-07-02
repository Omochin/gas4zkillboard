function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu(
    "zKillboard",
    [
      {
        name: "Simple",
        functionName: "simpleKillboard"
      },
      {
        name: "Detailed",
        functionName: "detailedKillboard"
      },
      {
        name: "Retry",
        functionName: "retry"
      }
    ]
  );
}


function simpleKillboard() {
  killboard(
    inputAPI(),
    [
      "Date",
      "Kill or Loss",
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
      "Kill or Loss",
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


function retry() {
  if (ScriptApp.getProjectTriggers().length > 0) {
    Browser.msgBox("A trigger has already been registered.");
    return;
  }

  var lock = LockService.getDocumentLock();
  if (lock.tryLock(100)) {
    lock.releaseLock();
    ScriptApp.newTrigger("run").timeBased().everyMinutes(1).create();
  } else {
    Browser.msgBox("The script is already running.");
  }
}


var MAX_PER_REQUEST = 200;
var MODIFIERS_NAMES = [
  "character",
  "corporation",
  "alliance",
  "group",
  "ship",
  "region",
  "constellation",
  "system",
  "location"
];
var KILL_OR_LOSS_NAMES = [
  "character",
  "corporation",
  "alliance"
];

var CHARACTERS_SHEET_NAME = "Characters";
var CORPORATIONS_SHEET_NAME = "Corporations";
var ALLIANCES_SHEET_NAME = "Alliances";
var PROPERTIES_SHEET_NAME = "Properties";
var SHEET_NAMES = [
  CHARACTERS_SHEET_NAME,
  CORPORATIONS_SHEET_NAME,
  ALLIANCES_SHEET_NAME,  
  PROPERTIES_SHEET_NAME
];

var typeIDs;
var regionIDs;
var solarSystemIDs;

var charactersSheet;
var corporationsSheet;
var alliancesSheet;
var propertiesSheet;

var characterIDs;
var corporationIDs;
var allianceIDs;

var properties = {};

var elementGetterFunctions;


function initialize(spreadsheet) {
  regionIDs = parseJSON("https://gist.githubusercontent.com/Omochin/875f277325658541e5e4532afc3c9acd/raw/43abc1e441e855588372551782d2480f0fd2dd23/region_ids.json");
  solarSystemIDs = parseJSON("https://gist.githubusercontent.com/Omochin/6eb5ae7902f196cfdd3cf43a6d6600e7/raw/3ac6f591d07cae6c8d14c9d71a476b0e832da3fb/solar_system_ids.json");
  typeIDs = parseJSON("https://gist.githubusercontent.com/Omochin/1b21545f4fa2d1f2de4bdcff5e21a5f9/raw/07c264eb86b38c3f90eeb99a4faa54085115180b/type_ids.json");

  charactersSheet = getSheet(CHARACTERS_SHEET_NAME, spreadsheet);
  corporationsSheet = getSheet(CORPORATIONS_SHEET_NAME, spreadsheet);
  alliancesSheet = getSheet(ALLIANCES_SHEET_NAME, spreadsheet);
  propertiesSheet = getSheet(PROPERTIES_SHEET_NAME, spreadsheet);

  characterIDs = getIDs(charactersSheet);
  corporationIDs = getIDs(corporationsSheet);
  allianceIDs = getIDs(alliancesSheet);

  elementGetterFunctions = {};

  elementGetterFunctions["KillID"] = function(killmail) {
    return getLink(killmail["killmail_id"], "kill");
  }

  elementGetterFunctions["Kill or Loss"] = function(killmail, row, column, spreadsheet) {
    if (!getProperty("killOrLoss")) {
      return "";
    }

    var range = spreadsheet.getActiveSheet().getRange(row, column);
    range.setFontColor("white");
    for (var i = 0; i < KILL_OR_LOSS_NAMES.length; i++) {
      var kindName = getProperty(KILL_OR_LOSS_NAMES[i]);
      var kindIDName = KILL_OR_LOSS_NAMES[i] + '_id';
      if (kindIDName in killmail['victim']) {
        if (killmail['victim'][kindIDName] == kindName) {
          range.setBackground("darkred");
          return "LOSS";
        }
      }
    }

    range.setBackground("darkgreen");
    return "KILL";
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

  elementGetterFunctions["Date"] = function(killmail) {
    var date = new Date(killmail["killmail_time"]);
    return Utilities.formatDate(date, "UTC", "yyyy/MM/dd HH:mm");
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
    return getNumberText(killmail["victim"]["damage_taken"]);
  }

  elementGetterFunctions["Value"] = function(killmail) {
    var number = parseFloat(killmail["zkb"]["totalValue"]).toFixed(2);
    return getNumberText(number);
  }

  elementGetterFunctions["Points"] = function(killmail) {
    return killmail["zkb"]["points"];
  }

  elementGetterFunctions["Involved"] = function(killmail) {
    return killmail["attackers"].length;
  }
}

function inputAPI() {
  var api = {}
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(100)) {
    Browser.msgBox("The script is already running.");
    return api;
  }
  lock.releaseLock();

  var url = Browser.inputBox("Please input URL of zKillboard");

  api["killOrLoss"] = false;
  var splitURL = url.split("/");
  for (var i = 3; i < splitURL.length - 1; i++) {
    if (KILL_OR_LOSS_NAMES.indexOf(splitURL[i]) >= 0) {
      api[splitURL[i]] = splitURL[i + 1];
      api["killOrLoss"] = true;
    }

    if (MODIFIERS_NAMES.indexOf(splitURL[i]) >= 0) {
      splitURL[i] = splitURL[i] + "ID";
    }
  }
  splitURL.splice(3, 0, "api");
  api["url"] = splitURL.join("/");

  var maxLimit = Browser.inputBox("Please input the number of pages to retrieve(One page has 200 killmails)");
  if (!isFinite(maxLimit) || maxLimit < 1) {
    maxLimit = 1;
  }
  api["maxLimit"] = parseInt(maxLimit);

  return api;
}

function parseJSON(url) {
  var response = UrlFetchApp.fetch(url);
  return JSON.parse(response.getContentText());
}

function getNumberText(number) {
  return String(number).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,');
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

function getSheet(name, spreadsheet, retain) {
  var sheet = undefined;
  var sheets = spreadsheet.getSheets();

  for (var i in sheets) {
    if (sheets[i].getName() == name) {
      if (!retain && sheets[i].getLastRow() >= 1000) {
        spreadsheet.deleteSheet(sheets[i]);
      } else {
        sheet = sheets[i];
      }
      break;
    }
  }

  if (sheet) {
    return sheet;
  } else {
    return spreadsheet.insertSheet(name);
  }
}

function deleteSheet(name, spreadsheet) {
  var sheets = spreadsheet.getSheets();

  for (var i in sheets) {
    if (sheets[i].getName() == name) {
      spreadsheet.deleteSheet(sheets[i]);
      break;
    }
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

function getProperty(key) {
  if (key in properties) {
    return properties[key];
  }

  var lastRow = propertiesSheet.getLastRow();
  for (var row = 1; row <= lastRow; row++) {
    if (propertiesSheet.getRange(row, 1).getValue() == key) {
      var value = propertiesSheet.getRange(row, 2).getValue();;
      properties[key] = value;
      return value;
    }
  }

  return undefined;
}

function setProperty(key, value) {
  properties[key] = value;

  var lastRow = propertiesSheet.getLastRow();
  for (var row = 1; row <= lastRow; row++) {
    if (propertiesSheet.getRange(row, 1).getValue() == key) {
      propertiesSheet.getRange(row, 2).setValue(value);
      return;
    }
  }

  propertiesSheet.getRange(lastRow + 1, 1).setValue(key);
  propertiesSheet.getRange(lastRow + 1, 2).setValue(value);
}

function deleteTriggers() {
  var triggerKeys = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggerKeys.length; i++) {
    ScriptApp.deleteTrigger(triggerKeys[i]);
  }
}

function run() {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(100)) {
    Browser.msgBox("The script is already running.");
    return;
  }

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  initialize(spreadsheet);
  deleteTriggers();

  var sheet = getSheet(getProperty("activeSheet"), spreadsheet, true);
  sheet.activate();

  var startTime = new Date();
  var url = getProperty("url");
  var elementNames = getProperty("elementNames").split(":");
  var maxLimit = parseInt(getProperty("maxLimit"));
  var startRow = parseInt(getProperty("startRow"));
  var startLimit = parseInt((startRow - 1) / MAX_PER_REQUEST) + 1;
  var killmaiilsIndex = (startRow - 1) % MAX_PER_REQUEST;
  var row = startRow;

  for (var limit = startLimit; limit <= maxLimit; limit++) {
    var killmails = parseJSON(url + "page/" + limit + "/");

    for (; killmaiilsIndex < killmails.length; killmaiilsIndex++) {
      for (var column = 0; column < elementNames.length; column++) {
        try {
          value = elementGetterFunctions[elementNames[column]](
            killmails[killmaiilsIndex],
            row + 1,
            column + 1,
            spreadsheet
          );
        } catch(error) {
          value = "Unknown";
        }
        sheet.getRange(row + 1, column + 1).setValue(value);
      }

      setProperty("startRow", row + 1);

      var diff = parseInt((new Date() - startTime) / (1000 * 60));
      if (diff >= 5) {
        ScriptApp.newTrigger("run").timeBased().everyMinutes(1).create();
        lock.releaseLock();
        return;
      }
      row++;
    }

    killmaiilsIndex = 0;
  }

  lock.releaseLock();
  deleteTriggers();
}

function killboard(api, elementNames) {
  if (!api["url"] || api["url"] == "cancel") {
    return;
  }

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();

  sheet.clear();
  deleteTriggers();

  for (var i = 0; i < SHEET_NAMES.length; i++) {
    deleteSheet(SHEET_NAMES[i], spreadsheet);
    getSheet(SHEET_NAMES[i], spreadsheet);
  }
  propertiesSheet = getSheet(PROPERTIES_SHEET_NAME, spreadsheet);

  for (var i = 0; i < elementNames.length; i++) {
    sheet.getRange(1, 1 + i).setValue(elementNames[i]);
  }

  setProperty("startRow", 1);
  setProperty("activeSheet", sheet.getName());
  setProperty("elementNames", elementNames.join(":"));
  
  for (var key in api) {
    setProperty(key, api[key]);
  }

  sheet.activate();

  ScriptApp.newTrigger("run").timeBased().everyMinutes(1).create();
}
