/******************************************************
 * Settings.gs
 * WSBCO Medicare CRM
 * Version 2.0
 *
 * Reads business settings from Settings sheet
 ******************************************************/


/******************************************************
 * GET SETTING VALUE
 ******************************************************/

function getSetting(settingName) {


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Settings");



  if (!sheet) {

    throw new Error(
      "Settings sheet not found."
    );

  }



  const data =
    sheet
      .getDataRange()
      .getValues();



  for (
    let i = 1;
    i < data.length;
    i++
  ) {


    if (

      String(data[i][0])
      .trim()
      .toLowerCase()

      ===

      String(settingName)
      .trim()
      .toLowerCase()

    ) {


      return data[i][1];

    }

  }



  return "";

}





/******************************************************
 * SETTING VALUE
 ******************************************************/

function setSetting(settingName, value) {


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Settings");



  if (!sheet) {

    throw new Error(
      "Settings sheet not found."
    );

  }



  const data =
    sheet
      .getDataRange()
      .getValues();



  for (
    let i = 1;
    i < data.length;
    i++
  ) {


    if (

      String(data[i][0])
      .trim()
      .toLowerCase()

      ===

      String(settingName)
      .trim()
      .toLowerCase()

    ) {


      sheet
        .getRange(i + 1, 2)
        .setValue(value);


      return true;

    }

  }



  // Add new setting if not found

  sheet.appendRow([

    settingName,

    value

  ]);



  return true;

}





/******************************************************
 * CREATE DEFAULT SETTINGS
 ******************************************************/

function createDefaultSettings() {


  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();



  let sheet =
    ss.getSheetByName(
      "Settings"
    );



  if (!sheet) {


    sheet =
      ss.insertSheet(
        "Settings"
      );


  }



  sheet.clear();



  sheet
    .getRange("A1:B14")
    .setValues([


      [
        "Setting",
        "Value"
      ],


      [
        "Company Name",
        "WSBCO Medicare"
      ],


      [
        "Agent Name",
        "William Burton"
      ],


      [
        "Agent Phone",
        "541-305-6079"
      ],


      [
        "Agent Email",
        "williamsburton@gmail.com"
      ],


      [
        "Default State",
        "OR"
      ],


      [
        "Default Lead Status",
        "Lead"
      ],


      [
        "Default Lead Source",
        "Manual Entry"
      ],


      [
        "Time Zone",
        "America/Los_Angeles"
      ],


      [
        "Return Address",
        ""
      ],


      [
        "Twilio Phone",
        ""
      ],


      [
        "Twilio SID",
        ""
      ],


      [
        "Twilio Auth Token",
        ""
      ],


      [
        'Version',
        "2.0"
      ]

    ]);



  sheet
    .getRange("A1:B1")
    .setFontWeight("bold");


  sheet
    .autoResizeColumns(
      1,
      2
    );


}
