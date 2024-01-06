function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('getData')
    .addItem('getData', 'getData')
    .addToUi();
}

function getData() {
  // resetAllScriptProperties();
  // setLastTimeStampManually();
  var currentTimestamp = Utilities.formatDate(new Date(), "IST", "yyyy-MM-dd HH:mm:ss");

  var emailsData = searchAndExtractDataSinceLastTimestamp(currentTimestamp);

  console.log('currentTimeStamp: ' + currentTimestamp);

  if (emailsData.length > 0) {
    // console.log(emailsData);
    copyDataToSheet(emailsData);
  } else {
    console.log('No mails found');
  }
  updateLastTimeStamp(currentTimestamp);
}

function resetAllScriptProperties() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var allProperties = scriptProperties.getProperties();

  for (var property in allProperties) {
    scriptProperties.deleteProperty(property);
  }

  Logger.log('All script properties have been reset.');
}

function setLastTimeStampManually() {
  var cDate = new Date();
  cDate.setTime(cDate.getTime() - 90 * 60 * 1000);
  updateLastTimeStamp(cDate);
}

function getMonthName(month) {
  var monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  return monthNames[month - 1];
}

function formatAmount(amount) {
  var amt = Number(amount).toFixed(2);
  var tot = '₹' + amt.toLocaleString('en-IN', { maximumFractionDigits: 2 });
  return tot;

}

function convertToEpoch(istTimestamp) {
  var istDate = Utilities.formatDate(new Date(istTimestamp), "IST", "yyyy/MM/dd HH:mm:ss");
  return new Date(istDate).getTime() / 1000;
}

function formatDate(dateString) {
  var dateTimeParts = dateString.split(" ");
  var datePart = dateTimeParts[0];

  var dateComponents = datePart.split("-");
  var day = parseInt(dateComponents[0], 10);
  var month = parseInt(dateComponents[1], 10);
  var year = parseInt(dateComponents[2], 10);

  var formattedDate = day + ' ' + getMonthName(month) + ' ' + year;

  return formattedDate;
}

function formatDate2(dateString) {
  var dateComponents = dateString.split("-");
  var day = parseInt(dateComponents[0], 10);
  var month = parseInt(dateComponents[1], 10);
  var year = parseInt(dateComponents[2], 10);

  var formattedDate = day + ' ' + getMonthName(month) + ' 20' + year;

  return formattedDate;
}

function extractDataFromEmail(emailText) {

  // console.log(emailText);

  emailText = emailText.replace(/\n/g, ' ');
  emailText = emailText.replace(/\s+/g, ' ');

  // console.log(emailText);

  var creditCardPattern = /Thank you for using your HDFC Bank Credit Card ending 1947 for Rs (\d+\.\d+) at/;
  var upiPattern = /Rs\.(\d+\.\d+) has been debited from account \*\*1947 to VPA ([^\s]+) on (\d{2}-\d{2}-\d{2})\./;

  var creditCardMatch = emailText.match(creditCardPattern);
  var upiMatch = emailText.match(upiPattern);

  // console.log('creditCardMatch: ' + creditCardMatch + ' | upiMatch: ' + upiMatch);

  if (creditCardMatch) {
    return {
      amount: formatAmount(creditCardMatch[1]),
      merchant: emailText.match(/at ([^.]+) on/)[1].trim(),
      date: formatDate(emailText.match(/on (\d{2}-\d{2}-\d{4} \d{2}:\d{2}:\d{2})\./)[1])
    };
  } else if (upiMatch) {
    return {
      amount: formatAmount(upiMatch[1]),
      merchant: upiMatch[2],
      date: formatDate2(upiMatch[3])
    };
  }

  return {
    amount: "",
    merchant: "",
    date: ""
  };
}


function searchAndExtractDataSinceLastTimestamp(currentTimestamp) {
  var lastTimeStamp = getLastTimeStamp();
  var threads = GmailApp.search('from:alerts@hdfcbank.net after:' + convertToEpoch(lastTimeStamp) + ' before:' + convertToEpoch(currentTimestamp));
  var data = [];

  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();

    for (var j = 0; j < messages.length; j++) {
      var emailTimestamp = messages[j].getDate();

      // console.log(lastTimeStamp + ' | ' + emailTimestamp + ' | ' + currentTimestamp);

      if (convertToEpoch(emailTimestamp) >= convertToEpoch(lastTimeStamp) && convertToEpoch(emailTimestamp) <= convertToEpoch(currentTimestamp)) {
        var emailText = messages[j].getPlainBody();
        var extractedData = extractDataFromEmail(emailText);
        data.push(extractedData);
      }
    }
  }

  return data;
}



function copyDataToSheet(data) {
  data = data.filter(function (entry) {
    return entry.date !== "";
  });
  data.sort(function (a, b) {
    return new Date(a.date) - new Date(b.date);
  });

  // console.log(data);

  for (var i = 0; i < data.length; i++) {
    var datePart = data[i].date.split(' ')[1];

    if (datePart !== "") {
        var sheetName = datePart;
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

        if (!sheet) {
            sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
        }
        var rowData = [data[i].date, data[i].amount, data[i].merchant];
        console.log(rowData);
        sheet.appendRow(rowData);
    }
  }
}

function getLastTimeStamp() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var lastTimeStamp = scriptProperties.getProperty('lastTimeStamp');

  if (!lastTimeStamp) {
    var oneDayAgo = new Date();
    oneDayAgo.setUTCDate(oneDayAgo.getUTCDate() - 1);
    var lT = Utilities.formatDate(oneDayAgo, "IST", "yyyy-MM-dd HH:mm:ss");
    updateLastTimeStamp(lT);
    lastTimeStamp = scriptProperties.getProperty('lastTimeStamp');
  }
  console.log('lastTimeStamp: ' + Utilities.formatDate(new Date(lastTimeStamp), "IST", "yyyy-MM-dd HH:mm:ss"));
  return lastTimeStamp;
}

function updateLastTimeStamp(currentTimestamp) {
  console.log('updating lastTimeStamp: ' + Utilities.formatDate(new Date(currentTimestamp), "IST", "yyyy-MM-dd HH:mm:ss"));
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('lastTimeStamp', currentTimestamp);
}
