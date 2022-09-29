function dedupe() {

    var newUsers = [];
    var activeUsers = [];
    var inactiveUsers = [];
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var values = SpreadsheetApp.getActiveSheet().getDataRange().getValues()
    var notEmail = [];
  
    const validateEmail = (email) => {
      return String(email)
        .trim()
        .toLowerCase()
        .match(
          /^(([^<>()[\]\\.,;:\s@"]+(\.[^<>()[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/
        );
    };
  
    function fixEmail (email) {
       return String(email)
        .trim()
        .toLowerCase();
  
    }
  
    function extractEmails(text) {
      return text.match(/([a-zA-Z0-9._+-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)/gi);
  
    }
  
    function alertMessage() {
      var result = SpreadsheetApp.getUi().alert("Done!");
  
    }
  
  
    for (n = 0; n < values.length; ++n) {
      var first = values[n][0];
  
      if (first != "") {
        if (validateEmail(first)) {
          newUsers.push(fixEmail(first));
        }
        else {
          if (validateEmail(extractEmails(first))) {
            newUsers.push(fixEmail(extractEmails(first)));
          }
          else {
            notEmail.push(first);
          }
  
        }
  
      }
  
    }
  
  
  
    for (n = 0; n < values.length; ++n) {
      var first = values[n][1];
  
      if (first != "") {
        if (validateEmail(first)) {
          activeUsers.push(fixEmail(first));
        }
        else {
          if (validateEmail(extractEmails(first))) {
            activeUsers.push(fixEmail(extractEmails(first)));
          }
          else {
            notEmail.push(first);
          }
  
        }
      }
  
    }
  
      for (n = 0; n < values.length; ++n) {
      var first = values[n][2];
  
      if (first != "") {
        if (validateEmail(first)) {
          inactiveUsers.push(fixEmail(first));
        }
        else {
          if (validateEmail(extractEmails(first))) {
            inactiveUsers.push(fixEmail(extractEmails(first)));
          }
          else {
            notEmail.push(first);
          }
  
        }
  
      }
  
    }
  
  
  
  
    var duplicatesActive = [];
    var duplicatesInactive = [];
  
  
  
    for (v = 0; v < (newUsers.length); ++v) {
  
      for (b = 0; b < (activeUsers.length); ++b) {
  
        if (newUsers[v] == activeUsers[b]) {
          duplicatesActive.push(newUsers[v]);
        }
      }
    }
  
  
    for (v = 0; v < (newUsers.length); ++v) {
  
      for (b = 0; b < (inactiveUsers.length); ++b) {
  
        if (newUsers[v] == inactiveUsers[b]) {
          duplicatesInactive.push(newUsers[v]);
        }
      }
    }
  
  
  
    var duplicates = duplicatesActive.concat(duplicatesInactive);
  
    Logger.log(duplicatesActive.length + " Existing Active " + duplicatesActive);
    Logger.log(duplicatesInactive.length+ " Existing inactive " + duplicatesInactive);
    Logger.log(duplicates.length + " Existing Total " + duplicates);
  
  
    var onlyNew = newUsers.filter((item) => !duplicates.includes(item));
    var revokeActive = activeUsers.filter((item) => !newUsers.includes(item));
  
    
    Logger.log(onlyNew.length + " New Users " + onlyNew);
    Logger.log(revokeActive.length + " Revoke Users " + revokeActive);
    Logger.log(notEmail);
  
  
  function alertMessageTitle() {
    SpreadsheetApp.getUi().alert("Results", duplicatesActive.length + " Existing Active " + "\r\n" + duplicatesActive + "\r\n" + "\r\n" +
  duplicatesInactive.length+ " Existing inactive " + "\r\n" + duplicatesInactive + "\r\n" + "\r\n" + onlyNew.length + " New Users " + "\r\n" + onlyNew + "\r\n" + "\r\n" + revokeActive.length + " Revoke Users " + "\r\n" + revokeActive + "\r\n" + "\r\n" + " Discarded Input Below " + notEmail, SpreadsheetApp.getUi().ButtonSet.OK);
  
  } 
  
  alertMessageTitle();
  
  /** 
  function writeArrayToColumn() {
  //var sheet = SpreadsheetApp.getActiveSheet()
   duplicates.map(function (el) {
          return [el];
      });
  var range = sheet.getRange(2, 7, 5, 1);
  range.setValues(duplicates);
  }
  
  writeArrayToColumn();
  
  
    function printToSheet() {
      var range = sheet.getRange("K1:K10");
      range.setValues(duplicates);
  
    //sheet.getRange(1,7,duplicates.length).setValues(duplicates);
  
  }
  */
  
  
  
  }