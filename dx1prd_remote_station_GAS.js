// =============================================================
//  DX1PRD Remote Station Portal — Google Apps Script Backend
//  Deploy as: Web App → Execute as: Me → Who has access: Anyone
// =============================================================

const SHEET_NAME   = "Applications";
const STATUS_COL   = 12;   // Column L — status field
const CLUB_EMAIL   = "dx1prd@gmail.com";      // ← change to club address
const CLUB_CALLSIGN = "DX1PRD";
const STATION_NAME  = "DX1PRD Remote Station";
const ADMIN_NAME    = "Lot / DW1IRF";

// Column order in the sheet (1-indexed)
// A:timestamp B:name C:callsign D:license E:country F:experience
// G:intendedUse H:date I:slot J:email K:notes L:status

function doGet(e) {
  const action = e.parameter.action || "submit";

  if (action === "submit")       return handleSubmit(e.parameter);
  if (action === "getAll")       return handleGetAll();
  if (action === "updateStatus") return handleUpdateStatus(e.parameter);

  return jsonResponse({ error: "Unknown action" });
}

// ─── 1. SUBMIT a new application ────────────────────────────
function handleSubmit(p) {
  const sheet = getSheet();

  // Append row
  sheet.appendRow([
    p.timestamp || new Date().toISOString(),
    p.name, p.callsign, p.license, p.country,
    p.experience, p.intendedUse, p.date, p.slot,
    p.email, p.notes || "",
    "Pending"
  ]);

  // Notify club of new application
  sendNewApplicationAlert(p);

  return jsonResponse({ status: "ok" });
}

// ─── 2. GET ALL applications (for admin panel) ───────────────
function handleGetAll() {
  const sheet = getSheet();
  const rows  = sheet.getDataRange().getValues();

  // Row 1 is header; data starts at row 2 (index 1)
  const applications = rows.slice(1).map((r, i) => ({
    row:          i + 2,          // actual sheet row number
    timestamp:    r[0],
    name:         r[1],
    callsign:     r[2],
    license:      r[3],
    country:      r[4],
    experience:   r[5],
    intendedUse:  r[6],
    date:         r[7],
    slot:         r[8],
    email:        r[9],
    notes:        r[10],
    status:       r[11] || "Pending"
  }));

  return jsonResponse({ applications });
}

// ─── 3. UPDATE STATUS and send email notification ────────────
function handleUpdateStatus(p) {
  const row    = parseInt(p.row, 10);
  const status = p.status;   // "Approved" or "Denied"

  if (!row || !status) return jsonResponse({ error: "Missing row or status" });

  const sheet = getSheet();
  sheet.getRange(row, STATUS_COL).setValue(status);

  // Send email to applicant
  if (status === "Approved") {
    sendApprovalEmail(p);
  } else if (status === "Denied") {
    sendDenialEmail(p);
  }

  return jsonResponse({ status: "ok", row, newStatus: status });
}

// ─── Email: New application alert to club ───────────────────
function sendNewApplicationAlert(p) {
  const subject = `[${CLUB_CALLSIGN}] New Remote Station Application — ${p.callsign}`;
  const body = `
A new application has been submitted to the ${STATION_NAME} portal.

─────────────────────────────────
  Callsign   : ${p.callsign}
  Name       : ${p.name}
  License    : ${p.license} (${p.country})
  Experience : ${p.experience}
  Intended Use: ${p.intendedUse}
  Requested  : ${p.date}  ${p.slot}
  Email      : ${p.email}
  Notes      : ${p.notes || "(none)"}
─────────────────────────────────

Log in to the Admin Panel to approve or deny this application.
An email notification will be sent to the applicant automatically.

73 de ${STATION_NAME} Portal
  `.trim();

  GmailApp.sendEmail(CLUB_EMAIL, subject, body);
}

// ─── Email: Approval notification to applicant ──────────────
function sendApprovalEmail(p) {
  const subject = `[${CLUB_CALLSIGN}] Remote Station Access APPROVED — ${p.callsign}`;
  const body = `
Dear ${p.name} / ${p.callsign},

Congratulations! Your application to operate the ${STATION_NAME} has been APPROVED.

─────────────────────────────────
  Your Callsign : ${p.callsign}
  Scheduled     : ${p.date}  ${p.slot}
  Intended Use  : ${p.intendedUse}
─────────────────────────────────

NEXT STEPS:
1. You will receive a separate email with your login credentials for the remote station web interface.
2. Please review the operating policies at the portal before your session.
3. Reply to this email if you need to reschedule or have any questions.

OPERATING REMINDERS:
• You must be actively present at the keyboard at all times while transmitting.
• Follow ITU and Philippine regulations (or your home country's, whichever is more restrictive).
• Log any QSOs made during your session and send the ADIF/log to ${CLUB_EMAIL}.
• No unattended or automated operation without prior coordination.

We look forward to having you on the air with us!

73 de ${ADMIN_NAME}
${STATION_NAME}
${CLUB_EMAIL}
  `.trim();

  GmailApp.sendEmail(p.email, subject, body, {
    name: STATION_NAME,
    replyTo: CLUB_EMAIL,
    cc: CLUB_EMAIL
  });
}

// ─── Email: Denial notification to applicant ────────────────
function sendDenialEmail(p) {
  const subject = `[${CLUB_CALLSIGN}] Remote Station Application Update — ${p.callsign}`;
  const body = `
Dear ${p.name} / ${p.callsign},

Thank you for your interest in operating the ${STATION_NAME}.

After reviewing your application, we are unable to approve access at this time.

This may be due to scheduling conflicts, eligibility requirements, or capacity constraints. We encourage you to re-apply for a future date or contact us directly if you have questions.

Feel free to reach out at ${CLUB_EMAIL} and we'll do our best to assist you.

73 de ${ADMIN_NAME}
${STATION_NAME}
${CLUB_EMAIL}
  `.trim();

  GmailApp.sendEmail(p.email, subject, body, {
    name: STATION_NAME,
    replyTo: CLUB_EMAIL
  });
}

// ─── Helpers ─────────────────────────────────────────────────
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      "Timestamp","Name","Callsign","License","Country",
      "Experience","Intended Use","Date","Slot","Email","Notes","Status"
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
