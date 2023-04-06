function getAllColumnValues() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Manager Sheets")
  let columnNumber = 2;
  let columnValues = sheet.getRange(1, columnNumber, sheet.getLastRow(), 1).getValues();
  let urlArray = []
  for (let i = 1; i < columnValues.length; i++) {
    urlArray.push(columnValues[i][0]);
  }
  console.log(urlArray)
  return urlArray
}

function getValuesFromOtherSpreadsheets() {
  let spreadsheetUrls = getAllColumnValues()
  let allSheetData = [];
  let sheetNameArray = gettingSettingTabData()
  for (let i = 0; i < sheetNameArray.length * 1; i++) {
    let sheetName = sheetNameArray[i][0];
    let setSheetDate = sheetNameArray[i][1]
    console.log(sheetName)
    for (let i = 0; i < spreadsheetUrls.length * 1; i++) {
      let otherSpreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrls[i]);
      let sheet = otherSpreadsheet.getSheetByName(sheetName);
      let dataRange = sheet.getDataRange();
      let values = dataRange.getValues();
      for (let j = 0; j < values.length * 1; j++) {
        values[j].unshift(mainFileName(spreadsheetUrls[i]));
      }
      for (let i = 1; i < values.length; i++) {
        allSheetData = checkSpreadsheetApp(allSheetData, values, i, sheetName, setSheetDate)
      }
    }
  }
  return allSheetData
}
function gettingSettingTabData() {
  let mainSheetUrl = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1yq8Z6sEbdKXLMqhI50IyZkszeV5XirDQb6YGfjVJVro/edit#gid=242822988");
  let sheetSettingTab = mainSheetUrl.getSheetByName("Setting Tab");
  let dataSettingTab = sheetSettingTab.getRange(2, 1, sheetSettingTab.getLastRow() - 1, sheetSettingTab.getLastColumn()).getValues();
  return dataSettingTab
}
function checkSpreadsheetApp(allSheetData, values, i, sheetName, setSheetDate) {
  if (values[i][5] !== '') {
    values[i].unshift(sheetName)
    console.log("sheetName")
    if (sheetName === "2023Q01") {
      values[i].unshift(setSheetDate)
    }
    else if (sheetName === "2023Q02") {
      values[i].unshift(setSheetDate)
    }
    else if (sheetName === "2023Q03") {
      values[i].unshift(setSheetDate)
    }
    else if (sheetName === "2023Q04") {
      values[i].unshift(setSheetDate)
    }
    allSheetData.unshift(values[i])
  }
  return allSheetData
}

function setSheetValues() {
  let allSheetData = sortValuesAscending(getValuesFromOtherSpreadsheets())
  let values = addSerialNumber(allSheetData)
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Score Details");
  let range = sheet.getRange(2, 1, values.length, values[0].length);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
  range.setValues(values);
}


function sortValuesAscending(values) {
  values.sort();
  return values;
}


function addSerialNumber(values) {
  return values.map(function (row, index) {
    return [index + 1].concat(row);
  });
}

function mainFileName(url) {
  let splitUrl = url.split('/');
  let spreadsheetId = splitUrl[5];
  let file = DriveApp.getFileById(spreadsheetId);
  let fileName = file.getName();
  return fileName
}




