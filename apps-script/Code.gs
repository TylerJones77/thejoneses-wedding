/**
 * The Joneses – form receiver.
 * Paste into the Google Sheet via Extensions → Apps Script, then
 * Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone).
 * RSVPs and song requests become rows; each one emails NOTIFY.
 */
const NOTIFY = "tylerjones30@icloud.com";

const TABS = {
  rsvp: { name: "RSVPs", headers: ["Submitted", "Name", "Email", "Attending", "Number attending", "Dietary needs", "Note"] },
  song: { name: "Song Requests", headers: ["Submitted", "Song", "Artist", "Requested by"] }
};

function clean(v) {
  v = v == null ? "" : String(v).slice(0, 1000);
  return /^[=+\-@]/.test(v) ? "'" + v : v; // keep guest text from running as a formula
}

function tab(type) {
  const t = TABS[type];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(t.name);
  if (!sh) {
    sh = ss.insertSheet(t.name);
    sh.appendRow(t.headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, t.headers.length).setFontWeight("bold");
  }
  return sh;
}

function doPost(e) {
  const p = (e && e.parameter) || {};
  const type = p.type === "song" ? "song" : "rsvp";
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (type === "song") {
      tab("song").appendRow([new Date(), clean(p.title), clean(p.artist), clean(p.by)]);
      MailApp.sendEmail(NOTIFY, "New song request: " + p.title,
        p.title + " by " + p.artist + (p.by ? "\nRequested by " + p.by : ""));
    } else {
      tab("rsvp").appendRow([new Date(), clean(p.name), clean(p.email), clean(p.attending), clean(p.count), clean(p.diet), clean(p.note)]);
      MailApp.sendEmail(NOTIFY, "RSVP from " + p.name + (p.attending === "yes" ? " – attending (" + p.count + ")" : " – can't make it"),
        "Name: " + p.name + "\nEmail: " + (p.email || "") + "\nAttending: " + p.attending +
        "\nNumber: " + p.count + "\nDietary: " + (p.diet || "") + "\nNote: " + (p.note || ""));
    }
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}

/** Run once from the editor to create both tabs and approve permissions. */
function setup() { tab("rsvp"); tab("song"); }
