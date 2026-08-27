/**
 * Strauch Agency — Quote Application intake
 * Backend for Personal_Quote_Application.html and Commercial_Quote_Application.html
 *
 * SETUP (one time, ~4 minutes)
 *  1. Create a new Google Sheet. Name it "Quote Applications".
 *  2. Extensions > Apps Script. Delete the sample code, paste this whole file.
 *  3. Change NOTIFY_EMAIL below if you want alerts somewhere other than the
 *     account that owns the Sheet.
 *  4. Deploy > New deployment > type: Web app
 *       Execute as: Me
 *       Who has access: Anyone      <-- must be "Anyone", not "Anyone with a Google Account"
 *  5. Copy the /exec URL it gives you and paste it into ENDPOINT at the top of
 *     BOTH html files, then publish them.
 *
 * Any time you change this file you must Deploy > NEW deployment (not "manage"),
 * and re-paste the new /exec URL into the forms.
 */

var INGEST_KEY   = "strauch-quote-2026";   // must match INGEST_KEY in both html files
var NOTIFY_EMAIL = "";                      // blank = the Sheet owner's address
var MAX_FIELDS   = 400;                     // spam / abuse ceiling
var MAX_LEN      = 5000;                    // per-answer character ceiling

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    var body = JSON.parse(e.postData.contents);
    if (body.key !== INGEST_KEY) return json({ ok: false, error: "unauthorized" });

    var type = body.type === "Commercial" ? "Commercial" : "Personal";
    var data = body.data || {};
    var keys = Object.keys(data);
    if (!keys.length)          return json({ ok: false, error: "empty submission" });
    if (keys.length > MAX_FIELDS) return json({ ok: false, error: "too many fields" });

    var row = { "Received": new Date(), "Status": "New" };
    keys.forEach(function (k) {
      row[k] = String(data[k]).slice(0, MAX_LEN);
    });

    writeRow(type, row);
    notify(type, row);
    return json({ ok: true });

  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/** Appends a row, growing the header row whenever a new field shows up. */
function writeRow(tabName, row) {
  var ss  = SpreadsheetApp.getActive();
  var sh  = ss.getSheetByName(tabName) || ss.insertSheet(tabName);

  var headers = sh.getLastColumn()
    ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].filter(String)
    : [];

  var added = Object.keys(row).filter(function (k) { return headers.indexOf(k) === -1; });
  if (added.length) {
    headers = headers.concat(added);
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight("bold").setBackground("#E01933").setFontColor("#FFFFFF");
    sh.setFrozenRows(1);
  }

  var values = headers.map(function (h) { return row.hasOwnProperty(h) ? row[h] : ""; });
  sh.appendRow(values);
}

/** Emails the highlights so nobody has to babysit the Sheet. */
function notify(type, row) {
  var to = NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
  if (!to) return;

  var who = type === "Commercial"
    ? (row.legalName || row.contactName || "New business")
    : ((row.firstName || "") + " " + (row.lastName || "")).trim();
  var phone = row.phone || row.contactPhone || "";
  var email = row.email || row.contactEmail || "";

  var lines = [];
  Object.keys(row).forEach(function (k) {
    if (k === "Received" || k === "Status") return;
    if (String(row[k]).length) lines.push(k + ": " + row[k]);
  });

  MailApp.sendEmail({
    to: to,
    subject: "New " + type + " quote application — " + (who || "unnamed"),
    body: who + "\n" + phone + "   " + email + "\n" +
          "Wants: " + (row.coverageWanted || "").split("|").join(", ") + "\n" +
          "\n----------------------------------------\n\n" +
          lines.join("\n") +
          "\n\n----------------------------------------\n" +
          "Full record is in the Quote Applications sheet, tab \"" + type + "\"."
  });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Visiting the /exec URL in a browser should say this. Handy deploy check. */
function doGet() {
  return ContentService.createTextOutput("Strauch quote intake is live.");
}

/** Run this from the editor once to confirm writes and email work end to end. */
function testSubmission() {
  writeRow("Personal", {
    Received: new Date(), Status: "TEST — delete me",
    firstName: "Testy", lastName: "McTestface", phone: "555-0100",
    email: "test@example.com", coverageWanted: "Auto|Home"
  });
  Logger.log("Wrote a test row to the Personal tab. Delete it when you're satisfied.");
}
