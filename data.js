if (typeof require !== "undefined") {
  UnitTestingApp = require("./UnitTestingApp.min.js");
  SheetUrlHelper = require("./SheetUrlHelper.js");
  GettingEmailHelper = require("./EmployeeDataHelper.js");

}

function getValuesFromOtherSpreadsheets() {
  const urlData = new SheetUrlHelper();
  let spreadsheetUrls = urlData.getAllColumnValues();
  let allSheetData = [];
  let sheetNameArray = gettingSettingTabData();
  let emailArray = [];

  for (let sheetNameArrayItem of sheetNameArray) {
    let sheetName = sheetNameArrayItem[0];
    let setSheetDate = sheetNameArrayItem[1];
    for (let spreadsheetUrl of spreadsheetUrls) {
      let otherSpreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
      let urlSheetName = otherSpreadsheet.getName();
      let sheet = otherSpreadsheet.getSheetByName(sheetName);
      let dataRange = sheet.getDataRange();
      let values = dataRange.getValues();

      for (let valuesItem of values) {
        valuesItem.unshift(mainFileName(spreadsheetUrl));
      }

      for (let i = 1; i < values.length; i++) {
        allSheetData = checkSpreadsheetApp(
          allSheetData,
          values,
          i,
          sheetName,
          setSheetDate,
          sheetNameArray
        );

        emailArray.push(forGettingEmail(values, i, sheetName, sheetNameArray, urlSheetName));
      }
    }
  }

  emailArray = emailArray.filter(function (email) {
    return email !== undefined;
  });

  let uniqueEmails = new Set(emailArray);
  emailArray = Array.from(uniqueEmails);

  for (let email of emailArray) {
    console.log(email)
    // MailApp.sendEmail(email, gettingEmailTemplateData()[0][1], gettingEmailTemplateData()[1][1]);
  }

  return allSheetData;
}

function gettingSettingTabData() {
  let mainSheetUrl = SpreadsheetApp.getActiveSpreadsheet();
  let sheetSettingTab = mainSheetUrl.getSheetByName("Setting Tab");
  let dataSettingTab = sheetSettingTab
    .getRange(
      2,
      1,
      sheetSettingTab.getLastRow() - 1,
      sheetSettingTab.getLastColumn()
    )
    .getValues();
  return dataSettingTab;
}
function gettingEmployeeData() {
  let mainSheetUrl = SpreadsheetApp.getActiveSpreadsheet();
  let sheetEmployeeDetails = mainSheetUrl.getSheetByName("Employee Details");
  let dataEmployeeDetails = sheetEmployeeDetails
    .getRange(
      2,
      1,
      sheetEmployeeDetails.getLastRow() - 1,
      sheetEmployeeDetails.getLastColumn()
    )
    .getValues();
  return dataEmployeeDetails;
}
function gettingEmailTemplateData() {
  let mainSheetUrl = SpreadsheetApp.getActiveSpreadsheet();
  let sheetEmailTemplates = mainSheetUrl.getSheetByName("Email Templates");
  let dataEmailTemplates = sheetEmailTemplates
    .getRange(
      1,
      1,
      sheetEmailTemplates.getLastRow(),
      sheetEmailTemplates.getLastColumn()
    )
    .getValues();
  return dataEmailTemplates;
}

function checkSpreadsheetApp(
  allSheetData,
  values,
  i,
  sheetName,
  setSheetDate,
  sheetNameArray
) {

  if (values[i][5] !== "") {
    values[i].unshift(sheetName);
    let currentDate = new Date();
    if (
      sheetName === sheetNameArray[0][0] &&
      currentDate > sheetNameArray[0][1] ||
      sheetName === sheetNameArray[1][0] &&
      currentDate > sheetNameArray[1][1] ||
      sheetName === sheetNameArray[2][0] &&
      currentDate > sheetNameArray[2][1] ||
      sheetName === sheetNameArray[3][0] &&
      currentDate > sheetNameArray[3][1]
    ) {
      values[i].unshift(setSheetDate);
    }
    allSheetData.unshift(values[i]);
  }
  return allSheetData;
}

function forGettingEmail(values, i, sheetName, sheetNameArray, urlSheetName) {
  let currentDate = new Date();
  let employeeData = gettingEmployeeData();
  let email;
  const emailHelper = new GettingEmailHelper();
  if (
    (values[i][5] === "" &&
      sheetName === sheetNameArray[0][0] &&
      currentDate > sheetNameArray[0][1]) ||
    (values[i][5] === "" &&
      sheetName === sheetNameArray[1][0] &&
      currentDate > sheetNameArray[1][1]) ||
    (values[i][5] === "" &&
      sheetName === sheetNameArray[2][0] &&
      currentDate > sheetNameArray[2][1]) ||
    (values[i][5] === "" &&
      sheetName === sheetNameArray[3][0] &&
      currentDate > sheetNameArray[3][1])
  ) {
    email = emailHelper.employeeDataUserSheetEmailFunction(
      employeeData,
      urlSheetName,
      email
    );
    return email
  }

}

function compareAndUpdateData() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Score Details");
  let data = sheet.getDataRange().getValues();
  let oldData = data.slice(1).map(function (row) {
    return row.slice(1);
  });
  const lastRow = sheet.getLastRow();

  const newData = sortValuesAscending(getValuesFromOtherSpreadsheets());
  const newDataFiltered = newData.filter(newRow => {
    return oldData.findIndex(oldRow => {
      return oldRow[3] === newRow[3] && oldRow[1] === newRow[1];
    }) === -1;
  });
  for (let [index, data] of newDataFiltered.entries()) {
    newDataFiltered[index].unshift(index + lastRow);
  }

  if (newDataFiltered.length > 0) {
    sheet.getRange(lastRow + 1, 1, newDataFiltered.length, 10).setValues(newDataFiltered);
  }
}

function sortValuesAscending(values) {
  values.sort();
  return values;
}

function mainFileName(url) {
  let splitUrl = url.split("/");
  let spreadsheetId = splitUrl[5];
  let file = DriveApp.getFileById(spreadsheetId);
  let fileName = file.getName();
  return fileName;
}
