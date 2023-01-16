function dedupe() {

    var newUsers = [];
    var activeUsers = [];
    var inactiveUsers = [];
    var notEmail = [];
    var duplicateUsers = [];
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var values = SpreadsheetApp.getActiveSheet().getDataRange().getValues()
  
  
    //regex to find valid entries (emails)
    const validateEmail = (email) => {
      return String(email)
        .trim()
        .toLowerCase()
        .match(
          /^(([^<>()[\]\\.,;:\s@"]+(\.[^<>()[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/
        );
    };
  
    //removing spaces around emails and dropping all characters to lower case
    function fixEmail(email) {
      return String(email)
        .trim()
        .toLowerCase();
  
    }
  
    //regex to extract emails from spaces which include more text
    function extractEmails(text) {
      return String(text).match(/([a-zA-Z0-9._+-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)/gi);
  
    }
  
    //duplicate email mapping
    function dupe(a) {
      var seen = {};
      return a.filter(function (item) {
        return seen.hasOwnProperty(item) ? true : (seen[item] = false);
      });
    }
  
      //unique email mapping
    function uniq(a) {
      var seen = {};
      return a.filter(function (item) {
        return seen.hasOwnProperty(item) ? false : (seen[item] = true);
      });
    }
  
  
  //nested if statements for iterating through first column (new emails submitted)
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
  
  
  //iterating through second column (existing active emails)
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
  
    
    //iterating through third column, inactive emails
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
  
    //spliting duplicates and unique users from first column (new submitted users)
    duplicateUsers = dupe(newUsers);
    newUsers = uniq(newUsers);
  
  
  
  
    var duplicatesActive = [];
    var duplicatesInactive = [];
  
  
  //comparing formatted version of new users to active users
    for (v = 0; v < (newUsers.length); ++v) {
  
      for (b = 0; b < (activeUsers.length); ++b) {
  
        if (newUsers[v] == activeUsers[b]) {
          duplicatesActive.push(newUsers[v]);
        }
      }
    }
  
  //comparing formatted version of new users to inactive users
    for (v = 0; v < (newUsers.length); ++v) {
  
      for (b = 0; b < (inactiveUsers.length); ++b) {
  
        if (newUsers[v] == inactiveUsers[b]) {
          duplicatesInactive.push(newUsers[v]);
        }
      }
    }
  
  
    //compiling all extracted info into arrays to be printed cleanly 
    var duplicates = duplicatesActive.concat(duplicatesInactive);
  
    var onlyNew = newUsers.filter((item) => !duplicates.includes(item));
    var revokeActive = activeUsers.filter((item) => !newUsers.includes(item));
  

    //alert message popup in google sheets to display info
    function alertMessageTitle() {
      SpreadsheetApp.getUi().alert("Results", duplicatesActive.length + " Existing Active " + "\r\n" + duplicatesActive + "\r\n" + "\r\n" +
        duplicatesInactive.length + " Existing inactive " + "\r\n" + duplicatesInactive + "\r\n" + "\r\n" + onlyNew.length + " New Users " + "\r\n" + onlyNew + "\r\n" + "\r\n" + revokeActive.length + " Revoke Users " + "\r\n" + revokeActive + "\r\n" + "\r\n" + duplicateUsers.length + " Duplicate Users Submitted " + "\r\n" + duplicateUsers + "\r\n" + "\r\n" + " Discarded Input Below " + "\r\n" + notEmail, SpreadsheetApp.getUi().ButtonSet.OK);
  
    }
  
    alertMessageTitle();
  
  
  
  
  }
