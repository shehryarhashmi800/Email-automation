// ============================================================
// HAWKTOK EMAIL AUTOMATION SCRIPT
// Version: 4.0 — UID Based Tracking
// Alert Email: shehryarhashmi800@gmail.com
// ============================================================

var CONFIG = {
  FROM_EMAIL: "YOUR_EMAIL_HERE",
  FROM_NAME: "Bruce",
  ALERT_EMAIL: "YOUR_EMAIL_HERE",
  SHEET_ID: "PASTE_YOUR_SHEET_ID_HERE",
  SHEET_NAME: "Outreach",
  DOC_ID: "PASTE_YOUR_DOC_ID_HERE",
  TIMEZONE: "Asia/Karachi",
  EMAIL_INTERVAL_SECONDS: 70,
  MAX_RETRIES: 3,
  TRACKING_URL: "PASTE_YOUR_WEB_APP_URL_HERE"
};

var EMAIL_START_COL = 6;
var COLS_PER_EMAIL  = 7; // UID, Date, Status, Sent At, Open Count, First Open, Last Open

var OFF = {
  UID:         0,
  DATE:        1,
  STATUS:      2,
  SENT_AT:     3,
  OPEN_COUNT:  4,
  FIRST_OPEN:  5,
  LAST_OPEN:   6
};

// ============================================================
// AUTO-REPLY DETECTION DATABASE
// ============================================================

var AUTO_REPLY_SENDER_PATTERNS = [
  "noreply", "no-reply", "donotreply", "do-not-reply",
  "mailer-daemon", "postmaster", "autoresponder",
  "auto-confirm", "notifications@", "support-noreply", "no_reply"
];

var AUTO_REPLY_SUBJECT_KEYWORDS = [
  "out of office", "out of the office", "away from office",
  "away from my desk", "away from the office", "on leave",
  "on vacation", "on holiday", "annual leave", "maternity leave",
  "paternity leave", "sick leave", "currently unavailable",
  "currently out", "not in the office", "back on", "returning on",
  "return on", "back in office", "limited access to email",
  "limited email access", "auto reply", "auto-reply",
  "automatic reply", "automated reply", "automatic response",
  "automated response", "auto response", "we received your",
  "we got your", "thank you for contacting",
  "thank you for reaching out", "thank you for your email",
  "thank you for your message", "thanks for contacting",
  "your request has been received", "your email has been received",
  "your message has been received", "ticket number", "ticket id",
  "case number", "case id", "reference number", "support ticket",
  "how was your experience", "rate your experience",
  "how would you rate", "leave a review", "satisfaction survey",
  "feedback survey", "rate the support", "rate our service",
  "delivery failed", "delivery failure", "undeliverable",
  "mail delivery failure", "returned mail", "could not be delivered",
  "does not exist", "no such user", "invalid address",
  "mailbox full", "mailbox not found"
];

var AUTO_REPLY_BODY_KEYWORDS = [
  "out of office", "out of the office", "away from office",
  "away from my desk", "on leave until", "on vacation until",
  "on holiday until", "back on", "returning on",
  "will respond when i return", "limited access to email",
  "will be back", "we received your email",
  "we received your message", "we got your message",
  "thank you for contacting", "thank you for reaching out",
  "thank you for your email", "thank you for your message",
  "thanks for contacting", "thanks for reaching out",
  "your request has been received", "your email has been received",
  "your message has been received", "we will get back to you",
  "we will be in touch", "someone will be in touch",
  "our team will contact you", "we aim to respond",
  "we typically respond within", "response time", "ticket number",
  "ticket id", "case number", "case id", "reference number",
  "support ticket", "your ticket", "help desk", "helpdesk",
  "support request", "been assigned to", "assigned to our team",
  "routed to", "forwarded to", "how was your experience",
  "rate your experience", "how would you rate", "leave a review",
  "satisfaction survey", "feedback survey", "click either link",
  "nps score", "how did we do", "tell us how we did",
  "rate the support", "rate our service",
  "do not reply to this email", "do not reply to this message",
  "this is an automated", "this is an automatic",
  "automated response", "automatic response", "automated message",
  "automatic message", "this email was sent automatically",
  "you are receiving this email because",
  "mailer-daemon", "delivery failed", "delivery failure",
  "undeliverable", "mail delivery failure", "could not be delivered",
  "does not exist", "no such user", "invalid address",
  "mailbox full", "mailbox not found", "account does not exist"
];

// ============================================================
// MAIN FUNCTION
// ============================================================

function sendScheduledEmails() {
  var sheet, data, headers, numEmails, docContent, repliedCol;

  try {
    sheet   = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
    data    = sheet.getDataRange().getValues();
    headers = data[0];
  } catch (e) {
    sendAlert("Sheet Not Accessible", "Script cannot open your Google Sheet.\n\nError: " + e.toString() + "\n\nAction: Check Sheet ID and sharing settings.");
    return;
  }

  numEmails  = countEmailColumns(headers);
  repliedCol = EMAIL_START_COL + (numEmails * COLS_PER_EMAIL);

  try {
    docContent = getDocContent(CONFIG.DOC_ID);
  } catch (e) {
    sendAlert("Doc Not Accessible", "Script cannot open your Google Doc.\n\nError: " + e.toString() + "\n\nAction: Check Doc ID and sharing settings.");
    return;
  }

  if (!canSendNow()) {
    Logger.log("Waiting for send interval...");
    return;
  }

  for (var i = 1; i < data.length; i++) {
    var row = data[i];

    if (!row[1] || !row[3]) continue;

    var brandName = row[1].toString().trim();
    var toEmail   = row[3].toString().trim();
    var sendTime  = formatTimeFromSheet(row[4]);
    var replied   = row[repliedCol - 1].toString().trim().toLowerCase();

    if (replied === "yes") {
      Logger.log("Skipping " + toEmail + " — replied.");
      continue;
    }

    for (var e = 0; e < numEmails; e++) {
      var emailNum     = e + 1;
      var base         = EMAIL_START_COL - 1 + (e * COLS_PER_EMAIL);
      var uidIdx       = base + OFF.UID;
      var dateIdx      = base + OFF.DATE;
      var statusIdx    = base + OFF.STATUS;
      var sentAtIdx    = base + OFF.SENT_AT;
      var openCntIdx   = base + OFF.OPEN_COUNT;
      var firstOpenIdx = base + OFF.FIRST_OPEN;
      var lastOpenIdx  = base + OFF.LAST_OPEN;

      var scheduledDate = row[dateIdx];
      var status        = row[statusIdx].toString().trim().toLowerCase();

      if (!scheduledDate || scheduledDate === "-" || scheduledDate === "") continue;
      if (status === "sent" || status === "skipped" || status === "abandoned") continue;

      if (!isValidDate(scheduledDate)) {
        sheet.getRange(i + 1, statusIdx + 1).setValue("Invalid Date");
        Logger.log("Invalid date in row " + (i + 1));
        continue;
      }

      if (!isValidTime(sendTime)) {
        Logger.log("Invalid time in row " + (i + 1) + ": [" + sendTime + "]");
        continue;
      }

      if (!isEmailDue(scheduledDate, sendTime)) continue;

      var retryKey   = "RETRY_" + i + "_" + emailNum;
      var retryCount = parseInt(PropertiesService.getScriptProperties().getProperty(retryKey) || "0");

      if (retryCount >= CONFIG.MAX_RETRIES) {
        sheet.getRange(i + 1, statusIdx + 1).setValue("Abandoned");
        Logger.log("Abandoned: " + toEmail + " Email " + emailNum);
        continue;
      }

      var emailContent = getEmailFromDoc(docContent, brandName, emailNum);

      if (!emailContent) {
        sheet.getRange(i + 1, statusIdx + 1).setValue("Template Missing");
        Logger.log("No template: [" + brandName + " - EMAIL " + emailNum + "]");
        continue;
      }

      // Generate unique UID for this email
      var uid         = generateUID();
      var trackedBody = addTrackingPixel(emailContent.body, uid);

      try {
        GmailApp.sendEmail(toEmail, emailContent.subject, "", {
          from:     CONFIG.FROM_EMAIL,
          name:     CONFIG.FROM_NAME,
          htmlBody: trackedBody
        });

        PropertiesService.getScriptProperties().setProperty("LAST_SENT", new Date().getTime().toString());
        PropertiesService.getScriptProperties().deleteProperty(retryKey);

        var now = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "M/d/yyyy h:mm a");

        // Store UID and update status
        sheet.getRange(i + 1, uidIdx       + 1).setValue(uid);
        sheet.getRange(i + 1, statusIdx    + 1).setValue("Sent");
        sheet.getRange(i + 1, sentAtIdx    + 1).setValue(now);
        sheet.getRange(i + 1, openCntIdx   + 1).setValue(0);
        sheet.getRange(i + 1, firstOpenIdx + 1).setValue("-");
        sheet.getRange(i + 1, lastOpenIdx  + 1).setValue("-");

        Logger.log("Sent Email " + emailNum + " to " + toEmail + " UID: " + uid);
        return;

      } catch (sendError) {
        PropertiesService.getScriptProperties().setProperty(retryKey, (retryCount + 1).toString());
        var errorMsg = sendError.toString();

        if (errorMsg.toLowerCase().indexOf("quota") !== -1 || errorMsg.toLowerCase().indexOf("limit") !== -1) {
          sheet.getRange(i + 1, statusIdx + 1).setValue("Quota Exceeded");
          sendAlert("Gmail Quota Exceeded", "Daily Gmail sending limit reached.\n\nLast attempted: " + toEmail);
          return;
        }

        sheet.getRange(i + 1, statusIdx + 1).setValue("Failed (Retry " + (retryCount + 1) + "/" + CONFIG.MAX_RETRIES + ")");
        Logger.log("Send failed for " + toEmail + ": " + errorMsg);
        return;
      }
    }
  }
}

// ============================================================
// REPLY CHECKER
// ============================================================

function checkReplies() {
  var sheet, data, headers, numEmails, repliedCol;

  try {
    sheet      = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
    data       = sheet.getDataRange().getValues();
    headers    = data[0];
    numEmails  = countEmailColumns(headers);
    repliedCol = EMAIL_START_COL + (numEmails * COLS_PER_EMAIL);
  } catch (e) {
    Logger.log("Reply check error: " + e.toString());
    return;
  }

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[3]) continue;

    var toEmail = row[3].toString().trim();
    var replied = row[repliedCol - 1].toString().trim().toLowerCase();

    if (replied === "yes") continue;

    try {
      var threads = GmailApp.search("from:" + toEmail, 0, 10);

      for (var t = 0; t < threads.length; t++) {
        var messages = threads[t].getMessages();

        for (var m = 0; m < messages.length; m++) {
          var message = messages[m];

          if (message.getFrom().indexOf(CONFIG.FROM_EMAIL) !== -1) continue;
          if (isAutoReply(message)) {
            Logger.log("Auto-reply detected from: " + toEmail + " — ignoring.");
            continue;
          }

          sheet.getRange(i + 1, repliedCol).setValue("Yes");

          for (var e = 0; e < numEmails; e++) {
            var statusIdx = EMAIL_START_COL - 1 + (e * COLS_PER_EMAIL) + OFF.STATUS;
            var status    = row[statusIdx].toString().trim().toLowerCase();
            if (status === "pending" || status === "" || status === "-") {
              sheet.getRange(i + 1, statusIdx + 1).setValue("Skipped");
            }
          }

          Logger.log("Real reply detected from: " + toEmail);
          break;
        }
      }
    } catch (e) {
      Logger.log("Reply check error for " + toEmail + ": " + e.toString());
    }
  }
}

// ============================================================
// OPEN TRACKER — UID BASED
// ============================================================

function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.track) {
      logEmailOpen(e.parameter.track);
    }
  } catch (err) {
    Logger.log("doGet error: " + err.toString());
  }
  return HtmlService.createHtmlOutput(
    '<img src="data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7" width="1" height="1"/>'
  );
}

function logEmailOpen(uid) {
  try {
    var sheet   = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
    var data    = sheet.getDataRange().getValues();
    var headers = data[0];
    var numEmails = countEmailColumns(headers);

    // Search for matching UID in sheet
    for (var i = 1; i < data.length; i++) {
      var row = data[i];

      for (var e = 0; e < numEmails; e++) {
        var base   = EMAIL_START_COL - 1 + (e * COLS_PER_EMAIL);
        var uidIdx = base + OFF.UID;

        if (row[uidIdx].toString().trim() === uid) {
          // Found matching UID
          var openCntIdx   = base + OFF.OPEN_COUNT;
          var firstOpenIdx = base + OFF.FIRST_OPEN;
          var lastOpenIdx  = base + OFF.LAST_OPEN;

          var now          = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "M/d/yyyy h:mm a");
          var currentCount = parseInt(sheet.getRange(i + 1, openCntIdx + 1).getValue() || "0");
          var firstOpen    = sheet.getRange(i + 1, firstOpenIdx + 1).getValue().toString().trim();

          // Increment count
          sheet.getRange(i + 1, openCntIdx + 1).setValue(currentCount + 1);

          // Set first open only once
          if (!firstOpen || firstOpen === "-" || firstOpen === "0" || firstOpen === "") {
            sheet.getRange(i + 1, firstOpenIdx + 1).setValue(now);
          }

          // Always update last open
          sheet.getRange(i + 1, lastOpenIdx + 1).setValue(now);

          Logger.log("Open tracked: UID " + uid + " Row " + (i+1) + " Email " + (e+1) + " Count: " + (currentCount + 1));
          return;
        }
      }
    }

    Logger.log("UID not found in sheet: " + uid + " — old email, ignored.");

  } catch (err) {
    Logger.log("logEmailOpen error: " + err.toString());
  }
}

// ============================================================
// AUTO-REPLY DETECTION
// ============================================================

function isAutoReply(message) {
  var subject = message.getSubject().toLowerCase();
  var body    = message.getPlainBody().toLowerCase();
  var from    = message.getFrom().toLowerCase();

  for (var s = 0; s < AUTO_REPLY_SENDER_PATTERNS.length; s++) {
    if (from.indexOf(AUTO_REPLY_SENDER_PATTERNS[s]) !== -1) {
      Logger.log("Auto-reply via sender: " + AUTO_REPLY_SENDER_PATTERNS[s]);
      return true;
    }
  }

  for (var sk = 0; sk < AUTO_REPLY_SUBJECT_KEYWORDS.length; sk++) {
    if (subject.indexOf(AUTO_REPLY_SUBJECT_KEYWORDS[sk]) !== -1) {
      Logger.log("Auto-reply via subject: " + AUTO_REPLY_SUBJECT_KEYWORDS[sk]);
      return true;
    }
  }

  for (var bk = 0; bk < AUTO_REPLY_BODY_KEYWORDS.length; bk++) {
    if (body.indexOf(AUTO_REPLY_BODY_KEYWORDS[bk]) !== -1) {
      Logger.log("Auto-reply via body: " + AUTO_REPLY_BODY_KEYWORDS[bk]);
      return true;
    }
  }

  return false;
}

// ============================================================
// HELPER FUNCTIONS
// ============================================================

function generateUID() {
  return new Date().getTime().toString() + "_" + Math.floor(Math.random() * 100000).toString();
}

function countEmailColumns(headers) {
  var count = 0;
  for (var i = EMAIL_START_COL - 1; i < headers.length; i++) {
    var h = headers[i].toString().toLowerCase();
    if (h.indexOf("email") !== -1 && h.indexOf("date") !== -1) count++;
  }
  return count;
}

function getDocContent(docId) {
  var doc = DocumentApp.openById(docId);
  return doc.getBody().getText();
}

function getDocHtml(docId) {
  var url = "https://docs.google.com/document/d/" + docId + "/export?format=html";
  var response = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
    followRedirects: true
  });
  return response.getContentText();
}

function getEmailFromDoc(plainText, brandName, emailNum) {
  try {
    var searchKey = "[" + brandName.toUpperCase() + " - EMAIL " + emailNum + "]";
    var upperText = plainText.toUpperCase();
    var keyIndex  = upperText.indexOf(searchKey);

    if (keyIndex === -1) return null;

    var afterKey    = plainText.substring(keyIndex + searchKey.length);
    var nextBracket = afterKey.indexOf("[");
    var section     = nextBracket === -1 ? afterKey : afterKey.substring(0, nextBracket);

    var subjectMatch = section.match(/Subject:\s*([^\n\r]+)/i);
    var subject      = subjectMatch ? subjectMatch[1].trim() : "(No Subject)";

    var bodyMatch = section.match(/Body:\s*([\s\S]+)/i);
    var bodyText  = bodyMatch ? bodyMatch[1].trim() : "";

    var htmlContent = getDocHtml(CONFIG.DOC_ID);
    var htmlBody    = extractHtmlSection(htmlContent, brandName, emailNum, bodyText);

    Logger.log("Subject: " + subject.substring(0, 80));
    Logger.log("Body length: " + htmlBody.length);

    return { subject: subject, body: htmlBody };

  } catch (e) {
    Logger.log("getEmailFromDoc error: " + e.toString());
    return null;
  }
}

function extractHtmlSection(htmlContent, brandName, emailNum, fallbackText) {
  try {
    var styleMatch = htmlContent.match(/<style[^>]*>([\s\S]*?)<\/style>/i);
    var cssText    = styleMatch ? styleMatch[1] : "";

    var classMap  = {};
    var ruleRegex = /\.([\w-]+)\s*\{([^}]+)\}/g;
    var match;
    while ((match = ruleRegex.exec(cssText)) !== null) {
      classMap[match[1]] = match[2].trim();
    }

    var searchKey = brandName.toUpperCase() + " - EMAIL " + emailNum;
    var upperHtml = htmlContent.toUpperCase();
    var keyPos    = upperHtml.indexOf(searchKey);

    if (keyPos === -1) return fallbackText.replace(/\n/g, "<br>");

    var nextKey   = upperHtml.indexOf("- EMAIL ", keyPos + searchKey.length);
    var endPos    = nextKey === -1 ? htmlContent.length : htmlContent.lastIndexOf("[", nextKey);
    var htmlChunk = htmlContent.substring(keyPos, endPos === -1 ? htmlContent.length : endPos);

    var bodyPos = htmlChunk.toUpperCase().indexOf("BODY:");
    if (bodyPos === -1) return fallbackText.replace(/\n/g, "<br>");

    var htmlBody = htmlChunk.substring(bodyPos + 5);
    htmlBody     = htmlBody.replace(/^[^<]*/, "").trim();

    htmlBody = htmlBody.replace(/class="([^"]+)"/g, function(m, classes) {
      var inlineStyle = "";
      classes.split(/\s+/).forEach(function(cls) {
        if (classMap[cls]) inlineStyle += classMap[cls] + ";";
      });
      return inlineStyle ? 'style="' + inlineStyle + '"' : "";
    });

    return htmlBody || fallbackText.replace(/\n/g, "<br>");

  } catch (e) {
    Logger.log("extractHtmlSection error: " + e.toString());
    return fallbackText.replace(/\n/g, "<br>");
  }
}

function formatTimeFromSheet(value) {
  if (!value) return "";
  if (value instanceof Date) {
    return Utilities.formatDate(value, CONFIG.TIMEZONE, "HH:mm");
  }
  return value.toString().trim();
}

function parseTime(sendTime) {
  var str = sendTime.toString().trim();

  var match12 = str.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
  if (match12) {
    var h = parseInt(match12[1]);
    var m = parseInt(match12[2]);
    var p = match12[3].toUpperCase();
    if (p === "PM" && h !== 12) h += 12;
    if (p === "AM" && h === 12) h = 0;
    return { hours: h, minutes: m };
  }

  var match24s = str.match(/^(\d{1,2}):(\d{2}):(\d{2})$/);
  if (match24s) return { hours: parseInt(match24s[1]), minutes: parseInt(match24s[2]) };

  var match24 = str.match(/^(\d{1,2}):(\d{2})$/);
  if (match24) return { hours: parseInt(match24[1]), minutes: parseInt(match24[2]) };

  return null;
}

function isEmailDue(scheduledDate, sendTime) {
  try {
    var d      = scheduledDate instanceof Date ? scheduledDate : new Date(scheduledDate);
    var parsed = parseTime(sendTime);
    if (!parsed) return false;

    var scheduled    = new Date(d.getFullYear(), d.getMonth(), d.getDate(), parsed.hours, parsed.minutes, 0);
    var scheduledStr = Utilities.formatDate(scheduled, CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm");
    var nowStr       = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm");

    return scheduledStr <= nowStr;
  } catch (e) {
    Logger.log("isEmailDue error: " + e.toString());
    return false;
  }
}

function isValidDate(value) {
  if (!value) return false;
  if (value instanceof Date) return !isNaN(value.getTime());
  return !isNaN(new Date(value).getTime());
}

function isValidTime(value) {
  if (!value) return false;
  var str = value.toString().trim();
  return /^\d{1,2}:\d{2}\s*(AM|PM)$/i.test(str) ||
         /^\d{1,2}:\d{2}:\d{2}$/.test(str) ||
         /^\d{1,2}:\d{2}$/.test(str);
}

function canSendNow() {
  var lastSent = PropertiesService.getScriptProperties().getProperty("LAST_SENT");
  if (!lastSent) return true;
  return (new Date().getTime() - parseInt(lastSent)) >= (CONFIG.EMAIL_INTERVAL_SECONDS * 1000);
}

function addTrackingPixel(body, uid) {
  if (!CONFIG.TRACKING_URL || CONFIG.TRACKING_URL === "") return body;
  var pixel = '<img src="' + CONFIG.TRACKING_URL + '?track=' + uid + '" width="1" height="1" style="display:none;border:0;" alt=""/>';
  return body + "<br><br>" + pixel;
}

function sendAlert(subject, message) {
  try {
    GmailApp.sendEmail(
      CONFIG.ALERT_EMAIL,
      "HawkTok Automation Alert: " + subject,
      "Hi Muhammad,\n\nAn issue was detected in your HawkTok Email Automation:\n\n" + message +
      "\n\nTime: " + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "M/d/yyyy h:mm a") +
      "\n\nThis is an automated alert."
    );
  } catch (e) {
    Logger.log("Alert email failed: " + e.toString());
  }
}

// ============================================================
// SETUP
// ============================================================

function setupTriggers() {
  var existing = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existing.length; i++) ScriptApp.deleteTrigger(existing[i]);

  ScriptApp.newTrigger("sendScheduledEmails").timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger("checkReplies").timeBased().everyMinutes(5).create();

  Logger.log("Triggers created: sendScheduledEmails (1 min), checkReplies (5 min)");
}

function clearRetries() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  Logger.log("All retry counters and properties cleared.");
}

function testDoc() {
  var plain  = getDocContent(CONFIG.DOC_ID);
  var result = getEmailFromDoc(plain, "Natural Life", 1);
  if (result) {
    Logger.log("Subject: " + result.subject);
    Logger.log("Body (first 300): " + result.body.substring(0, 300));
  } else {
    Logger.log("No template found");
  }
}
