function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('CS Internal Search') 
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getStaffContext() {
  var email = Session.getActiveUser().getEmail().toLowerCase();
  var level = 1; 
  var isBlockedGoogleSheet = false; 
  if (email.startsWith("wmcs")) isBlockedGoogleSheet = true;
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Staff_Database"); 
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      for (var i = 0; i < data.length; i++) {
        if (String(data[i][0]).toLowerCase().trim() === email) {
          var pos = String(data[i][2]).toLowerCase();         
          if (pos.includes("manager") || pos.includes("head")) level = 3;
          else if (pos.includes("sup") || pos.includes("senior")) level = 2;         
          if (pos.includes("psk")) isBlockedGoogleSheet = true;
          break; 
        }
      }
    }
  } catch (e) {}
  return { level: level, isBlocked: isBlockedGoogleSheet };
}

function getUserConfig() { return getStaffContext(); }

function getAllDataForClient() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var staff = getStaffContext();
  var userLevel = staff.level;
  var sheetNames = ["Tools", "Announcements", "Google sheet", "Senior_Auto_Sync", "Knowledge Base", "Channel"]; 
  var allData = [];
  var seenUrls = {};
  sheetNames.forEach(function(name) {
    if (name === "Senior_Auto_Sync" && staff.isBlocked) return;
    var sheet = ss.getSheetByName(name);
    if (sheet) {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) { 
        var data = sheet.getRange(2, 1, lastRow - 1, 6).getDisplayValues();         
        for (var i = 0; i < data.length; i++) {
          var row = data[i];
          if (row[0] === "" && row[1] === "") continue; 
          var rowRole = String(row[5]).toLowerCase().trim();
          var allowed = false;         
          if (rowRole === "") allowed = true;
          else if (rowRole === "senior" && userLevel >= 2) allowed = true;
          else if (rowRole === "sup" && userLevel >= 3) allowed = true;
          if (allowed) {
            var url = row[1];
            if (!url || !seenUrls[url] || name === "Announcements" || name === "Knowledge Base" || name === "Channel" || name === "Tools") { 
              var cat = "";
              var displayTitle = row[0];
              if (name === "Tools") cat = "Tools";
              else if (name === "Announcements") cat = "Announcements";
              else if (name === "Google sheet") cat = "Google sheet"; 
              else if (name === "Knowledge Base") cat = "Knowledge Base";
              else if (name === "Channel") cat = "Channel"; 
              else if (name === "Senior_Auto_Sync") {
                 if (rowRole === "senior") {
                    cat = "Senior"; 
                    displayTitle += " (Senior)";
                 } else {
                    cat = "Drive File"; 
                 }
              }
              allData.push({
                title: displayTitle,        
                url: row[1],          
                desc: row[2],         
                date: row[4],         
                category: cat, 
                searchText: (row[0] + " " + row[2] + " " + row[3]).toLowerCase()
              });
              if (url) seenUrls[url] = true;
            }
          }
        }
      }
    }
  }); 
  allData.sort(function(a, b) {
    var dateA = a.date ? new Date(a.date).getTime() : 0;
    var dateB = b.date ? new Date(b.date).getTime() : 0;
    return dateB - dateA;
  }); 
  return allData;
}

function getActiveCoinStatus() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Transaction Status"); 
    if (!sheet) return "ERROR_NO_SHEET";
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    var startRow = Math.max(2, lastRow - 300); 
    var numRows = lastRow - startRow + 1;
    var data = sheet.getRange(startRow, 1, numRows, 8).getValues();   
    var activeIssues = [];
    for (var i = 0; i < data.length; i++) {
      var status = String(data[i][0]).toLowerCase().trim(); 
      var msgRed = data[i][1];   
      var msgGreen = data[i][2]; 
      var link = data[i][7];     
      if (status === "ongoing" || status === "") {
        if (msgRed !== "" || msgGreen !== "") {
            var type = (msgRed !== "") ? "RED" : "GREEN";
            var message = (msgRed !== "") ? msgRed : msgGreen; 
            activeIssues.push({
              type: type,
              message: message,
              link: link
            });
        }
      }
    }
    return activeIssues.reverse();
  } catch(e) { return []; }
}

function syncDriveToSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Senior_Auto_Sync");
  if (!sheet) {
    sheet = ss.insertSheet("Senior_Auto_Sync");
    sheet.appendRow(["Title", "URL", "Desc", "Keyword", "Date", "Role"]);
  }
  var existingData = {};
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var oldValues = sheet.getRange(2, 1, lastRow - 1, 6).getDisplayValues();
    for (var i = 0; i < oldValues.length; i++) {
      var url = oldValues[i][1];
      existingData[url] = { date: oldValues[i][4], keyword: oldValues[i][3], row: oldValues[i] };
    }
  }  
  
  // SANITIZED FOLDER IDs
  var folderConfigs = [
    { id: "YOUR_DRIVE_FOLDER_ID_1", role: "" },
    { id: "YOUR_DRIVE_FOLDER_ID_2", role: "Senior" }
  ];  
 
  var allData = [];
  folderConfigs.forEach(function(config) {
    try {
      if (!config.id || config.id.startsWith("YOUR_")) return; 
      var rootFolder = DriveApp.getFolderById(config.id);
      processFolder(rootFolder, config.role, allData, existingData);
    } catch (e) {
      Logger.log("Error root folder " + config.id + ": " + e);
    }
  }); 
  
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow()-1, 6).clearContent();
  }
  if (allData.length > 0) {
    allData.sort(function(a, b) { return b[4].localeCompare(a[4]); });
    sheet.getRange(2, 1, allData.length, 6).setValues(allData);
  }
}

function processFolder(folder, role, allData, existingData) {
  if (!folder) return;
  try {
    var files = folder.getFiles(); 
    while (files.hasNext()) {
      var file = files.next();
      var fileUrl = file.getUrl();
      var lastUpdated = file.getLastUpdated();
      var dateStr = Utilities.formatDate(lastUpdated, Session.getScriptTimeZone(), "yyyy-MM-dd");    
      
      if (existingData[fileUrl] && existingData[fileUrl].date === dateStr) {
         var oldRow = existingData[fileUrl].row;
         oldRow[5] = role; 
         oldRow[2] = "📂 " + folder.getName(); 
         allData.push(oldRow);   
      } else {
         allData.push([file.getName(), fileUrl, "📂 " + folder.getName(), file.getName(), dateStr, role]);
      }
    }
    var subfolders = folder.getFolders();
    while (subfolders.hasNext()) { processFolder(subfolders.next(), role, allData, existingData); }
  } catch (e) {}
}

function maskDataFromWeb(input) {
  if (!input) return "";
  var items = input.split(/[\n,]+/).filter(item => item.trim() !== "");
  var results = items.map(function(item) {
    var trimmed = item.trim();
    if (trimmed.includes("@") && trimmed.includes(".")) {
      var parts = trimmed.toLowerCase().split("@");
      var nameRaw = parts[0]; var domainRaw = parts[1];
      var lastDotIndex = domainRaw.lastIndexOf(".");
      var domainName = domainRaw.substring(0, lastDotIndex);
      var domainExt = domainRaw.substring(lastDotIndex + 1);
      var maskedName = nameRaw.charAt(0) + "*".repeat(Math.max(1, nameRaw.length - 1));
      var maskedDomain = domainName.charAt(0) + "*".repeat(Math.max(1, domainName.length - 1));
      return maskedName + "@" + maskedDomain + "." + domainExt;
    } else if (/^\d{8,12}$/.test(trimmed)) {
      var len = trimmed.length;
      var mask1 = trimmed.substring(0, 3) + "x".repeat(len - 5) + trimmed.slice(-2);
      var mask2 = "x".repeat(len - 2) + trimmed.slice(-2);
      var mask3 = trimmed.substring(0, 2) + "x".repeat(len - 4) + trimmed.slice(-2);
      return mask1 + " / " + mask2 + " / " + mask3;
    }
    return trimmed;
  });
  return results.join("\n");
}
