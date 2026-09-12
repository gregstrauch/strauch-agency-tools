/**
 * Strauch Agency — Job application intake
 * Backend for https://strauchagency.com/careers/
 *
 * Deliberately SEPARATE from Quote_Intake_Code.gs: applicant data is HR data and
 * must not sit in the same Sheet as client quote data, which gets shared with
 * producers.
 *
 * SETUP (one time, ~3 minutes)
 *  1. Go to script.new  (do NOT use Extensions > Apps Script — Sheets hides that menu)
 *  2. Delete the sample code, paste this whole file, name it "Strauch Careers Intake"
 *  3. Deploy > New deployment > type: Web app
 *       Execute as: Me
 *       Who has access: Anyone      <-- must be "Anyone", NOT "Anyone with a Google Account"
 *  4. Copy the /exec URL and send it to Claude, or paste it into ENDPOINT in
 *     careers/index.html yourself, then publish.
 *
 * The Sheet and the resume folder already exist — the IDs below are filled in.
 * Any time you change this file you must Deploy > NEW deployment (not "manage").
 */

var SHEET_ID     = "1ETQt7SStHHBd3ggr9QgBSh0hIdc4NJBTLtz_yanvhic";  // "Strauch Agency — Job Applications"
var RESUME_FOLDER= "1PGokTKBRCXJDDgEz6MDEHx1OBCckkWLD";              // "Strauch Agency — Resumes"
var INGEST_KEY   = "strauch-careers-2026";  // must match INGEST_KEY in careers/index.html
var NOTIFY_EMAIL = "gstrauch@farmersagent.com";  // where application alerts go
var TAB          = "Applications";
var MAX_FIELDS   = 60;
var MAX_LEN      = 5000;
var MAX_RESUME_MB= 5;

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    var body = JSON.parse(e.postData.contents);
    if (body.key !== INGEST_KEY) return json({ ok: false, error: "unauthorized" });

    var data = body.data || {};
    var keys = Object.keys(data);
    if (!keys.length)             return json({ ok: false, error: "empty submission" });
    if (keys.length > MAX_FIELDS) return json({ ok: false, error: "too many fields" });

    var row = { "Received": new Date(), "Status": "New" };
    keys.forEach(function (k) { row[k] = String(data[k]).slice(0, MAX_LEN); });

    // Resume arrives as base64. Store it in Drive, keep only the link in the Sheet.
    if (body.resume && body.resume.b64) {
      var r = saveResume(body.resume, row.firstName, row.lastName);
      row["Resume"]     = r.url;
      row["ResumeName"] = r.name;
    }

    writeRow(row);
    notify(row);
    return json({ ok: true });

  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/** Writes the base64 resume to Drive and returns its link. */
function saveResume(resume, first, last) {
  var bytes = Utilities.base64Decode(resume.b64);
  if (bytes.length > MAX_RESUME_MB * 1024 * 1024) throw new Error("resume too large");

  var safe = String((first || "") + "_" + (last || "applicant"))
               .replace(/[^\w.-]+/g, "_").replace(/^_+|_+$/g, "") || "applicant";
  var ext  = (String(resume.name || "").match(/\.[A-Za-z0-9]{1,5}$/) || [""])[0];
  var name = safe + "_" + Utilities.formatDate(new Date(), "America/Denver", "yyyy-MM-dd") + ext;

  var blob = Utilities.newBlob(bytes, resume.type || "application/octet-stream", name);
  var file = DriveApp.getFolderById(RESUME_FOLDER).createFile(blob);
  return { url: file.getUrl(), name: name };
}

/** Appends a row, growing the header row whenever a new field shows up. */
function writeRow(row) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(TAB) || ss.insertSheet(TAB);

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

  sh.appendRow(headers.map(function (h) { return row.hasOwnProperty(h) ? row[h] : ""; }));
}

/** Emails the application so nobody has to babysit the Sheet. */
function notify(row) {
  var to = NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
  if (!to) return;

  var who = ((row.firstName || "") + " " + (row.lastName || "")).trim() || "Unnamed applicant";
  var lines = [];
  Object.keys(row).forEach(function (k) {
    if (k === "Received" || k === "Status") return;
    if (String(row[k]).length) lines.push(k + ": " + row[k]);
  });

  MailApp.sendEmail({
    to: to,
    subject: "Job application — " + (row.role || "open application") + " — " + who,
    body: who + "\n" + (row.phone || "") + "   " + (row.email || "") + "\n" +
          "Role: " + (row.role || "") + "\n" +
          "Licensed: " + (row.licensed || "") + "\n" +
          (row.Resume ? "Resume: " + row.Resume + "\n" : "No resume attached\n") +
          "\n----------------------------------------\n\n" +
          lines.join("\n") +
          "\n\n----------------------------------------\n" +
          "Full record: https://docs.google.com/spreadsheets/d/" + SHEET_ID + "/edit"
  });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Visiting the /exec URL in a browser should say this. Handy deploy check. */
function doGet() {
  return ContentService.createTextOutput("Strauch careers intake is live.");
}

/** Run once from the editor to confirm the Sheet write and the email work. */
function testApplication() {
  var row = {
    Received: new Date(), Status: "TEST — delete me",
    firstName: "Testy", lastName: "McTestface", phone: "555-0100",
    email: "test@example.com", role: "Producer", licensed: "No — willing to get licensed"
  };
  writeRow(row);
  notify(row);
  Logger.log("Wrote a test row to the Applications tab and emailed it. Delete the row when satisfied.");
}
