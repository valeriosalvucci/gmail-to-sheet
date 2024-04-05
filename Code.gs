function main(){
  processSetapay();
  processRakutenMobile();
  processRakutenCard();
}

/// Utils ///
function getEmailsMatchingLabel(label) {
  return GmailApp.search(`label:"auto/${label}" label:"auto/to-process"`);
}

function extractValueFromBody(body, regexPattern) {
  const regex = new RegExp(regexPattern, "g");
  const match = regex.exec(body);
  return match ? match[1] : "";
}

function appendDataToSheet(extractedData, sheetName, flag) {
  // TODO make this function more general and uncouple it
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (flag === 'SIMPLE') {
    if (extractedData.length > 0) {
      extractedData.forEach(row => {
        sheet.appendRow(row);
      });
      Logger.log("Found data. Writing to " + sheetName);
    } else {
      Logger.log("No data found in extractedData. Skipping writing to " + sheetName);
    }
  } else if (flag === 'SETAPAY') {
    let colName = "A";
    let lastRow = sheet.getLastRow();
    let range = sheet.getRange(colName + lastRow);
    let startRange;
    if (range.getValue() !== "") {
      startRange = lastRow + 1;
    } else {
      startRange = range.getNextDataCell(SpreadsheetApp.Direction.UP).getRow() + 1;
    }

    if (extractedData.length > 0) {
      const endRange = startRange + extractedData.length - 1;
      const rangeAddr = 'A' + startRange + ':F' + endRange;
      const dataRange = sheet.getRange(rangeAddr);
      dataRange.setValues(extractedData);
      Logger.log("Found data. Writing to " + sheetName);

    } else {
      Logger.log("No data found in extractedData. Skipping writing to " + sheetName);
    }
  }
}

function removeLabels(toRemoveLabel) {
  toRemoveLabel.forEach(message => {
    const thread = message.getThread();
    thread.removeLabel(GmailApp.getUserLabelByName("auto/to-process"));
  });
}


////  Rakuten card ////

function processRakutenCard(){
  const threads = getEmailsMatchingLabel("r-card");
  const { extractedData, toRemoveLabel } = processEmailsRakutenCard(threads);
  appendDataToSheet(extractedData, "r-card", "SIMPLE");
  removeLabels(toRemoveLabel);
}

function extractUsageBlocks(body) {
  const regex = /■利用日:\s*(.*?)\n■利用先:\s*(.*?)\n■利用者:\s*(.*?)\n■支払方法:\s*(.*?)\n■利用金額:\s*(.*?)\n■支払月:\s*(.*?)\n/gs; //■カード利用獲得ポイント:\s*(.*?)\n■ポイント獲得予定月:\s*(.*?)\n/gs;
  let match;
  const usageBlocks = [];
  while ((match = regex.exec(body)) !== null) {
    const usageInfo = {
      利用日: match[1].trim(),
      利用先: match[2].trim(),
      利用者: match[3].trim(),
      支払方法: match[4].trim(),
      利用金額: match[5].trim(),
      支払月: match[6].trim(),
    };
    usageBlocks.push(usageInfo);
  }
  const arrayofArrays = usageBlocks.map(block => {
      return [
        block['利用日'],
        block['利用先'],
        block['利用者'],
        parseInt(block['支払方法']),
        parseInt(block['利用金額'].split(' ')[0].replace(/,/g, '')),
        block['支払月'],
      ];
    });
  return arrayofArrays;
}


function processEmailsRakutenCard(threads) {
  const extractedData = [];
  const toRemoveLabel = [];

  for (const thread of threads) {
    const messages = thread.getMessages();

    for (const message of messages) {
      const labels = thread.getLabels();
      if (labels.some(label => label.getName() === "auto/to-process")) {
        const body = message.getPlainBody();

        let arrayofArrays = extractUsageBlocks(body)

        // TODO assert the correct number of itemes are extracted. 
        // To be notified of errors, especially if rakuten update their email format

        Logger.log(arrayofArrays)
        arrayofArrays.forEach(array => {
          extractedData.push(array);
        });
        toRemoveLabel.push(message);
      }
    }
  }
  extractedData.sort((a, b) => new Date(a[0]) - new Date(b[0]));
  return { extractedData, toRemoveLabel };
}

//// RakutenMobile ////
function processRakutenMobile(){
  const threads = getEmailsMatchingLabel("r-mobile");
  const { extractedData, toRemoveLabel } = processEmailsRakutenMobile(threads);
  appendDataToSheet(extractedData, "r-mobile", "SIMPLE");
  removeLabels(toRemoveLabel);
}


function processEmailsRakutenMobile(threads) {
  const extractedData = [];
  const toRemoveLabel = [];

  for (const thread of threads) {
    const messages = thread.getMessages();

    for (const message of messages) {
      const labels = thread.getLabels();
      if (labels.some(label => label.getName() === "auto/to-process")) {
        const body = message.getPlainBody();

        var regex = /(\d{4})年(\d{2})月/;
        var match = body.match(regex);
        var year = 0
        var month = 0
        if (match && match.length > 2) {
            year = parseInt(match[1]);
            month = parseInt(match[2]);
        } else {
            console.log("Year and month not found.");
        }

        var regex = /\[([\d,]+)円\]/;
        var match = body.match(regex);
        var amount = 0
        if (match && match.length > 1) {
            amount = parseInt(match[1].replace(/,/g, ''));
        } else {
            console.log("Amount not found.");
        }

        extractedData.push([year, month, amount]);
        toRemoveLabel.push(message);
      }
    }
  }

  extractedData.sort(function(a, b) {
    if (a[0] !== b[0]) {
      return a[0] - b[0];
    }
    return a[1] - b[1];
  });
  return { extractedData, toRemoveLabel };
}

//// Setapay ////
function processSetapay() {
  const threads = getEmailsMatchingLabel("setapay");
  const { extractedData, toRemoveLabel } = processEmailsSetapay(threads);
  appendDataToSheet(extractedData, "setapay", "SETAPAY");
  removeLabels(toRemoveLabel);
}

function processEmailsSetapay(threads) {
  const extractedData = [];
  const toRemoveLabel = [];

  for (const thread of threads) {
    const messages = thread.getMessages();

    for (const message of messages) {
      const labels = thread.getLabels();
      if (labels.some(label => label.getName() === "auto/to-process")) {
        const body = message.getPlainBody();

        const accountNumber = extractValueFromBody(body, /<span>▼アカウントナンバー<\/span><br>\s*<span>(\d{4}-\d{4}-\d{4}-\d{4})<\/span><br>/);
        const dateAndTime = extractValueFromBody(body, /(\d{4}\/\d{2}\/\d{2} \d{2}:\d{2}:\d{2})/);
        const setagayaCoin = extractValueFromBody(body, /せたがやコイン\s*：\s*([\d,]+)/).replace(/,/g, '');
        let amount = 0;
        let paymentRecipient = "Unknown";
        let transactionType = "Unknown";
        let sign = 1;

        if (body.includes("チャージが完了しましたので、お知らせいたします")) {
          transactionType = "charge";
          const chargeAmountRegex = /チャージ金額<\/span><br>\s*<span>([\d,]+)<\/span>/;
          amount = extractValueFromBody(body, chargeAmountRegex).replace(/,/g, '');
          paymentRecipient = extractValueFromBody(body, /▼チャージ方法[\s\S]*?<span>(.*?)<\/span>/);
        } else {
          transactionType = "payment";
          sign = -1;
          paymentRecipient = extractValueFromBody(body, /▼お支払い・送金先<\/span><br>\s*<span>(.*?)<\/span>/);
          const amountRegex = /ご利用金額<\/span><br>\s*<span>([\d,]+)<\/span>/;
          amount = extractValueFromBody(body, amountRegex).replace(/,/g, '');
        }

        extractedData.push([dateAndTime, transactionType, paymentRecipient, sign * parseInt(amount), sign * parseInt(setagayaCoin), accountNumber]);
        toRemoveLabel.push(message);

      }
    }
  }

  extractedData.sort((a, b) => new Date(a[0]) - new Date(b[0]));

  return { extractedData, toRemoveLabel };
}
