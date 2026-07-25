/**
 * VEEHIVES EVENT PH — Google Sheets backend for the website.
 *
 * SETUP (one time):
 * 1. Open your Google Sheet: https://docs.google.com/spreadsheets/d/1TVGXQLbcuwej_d9R_ZaX-Qp7B4_wZMn0EiayFFyAbqY/edit
 * 2. Extensions > Apps Script.
 * 3. Delete anything in the editor, paste this whole file in, then save (disk icon).
 * 4. In your Sheet, create two tabs (bottom of the screen) named exactly:
 *      "Bookings"  — leave empty, this fills in automatically when someone books.
 *      "Events"    — add one header row with these exact column names:
 *                    title | date | description | image
 *                    Then one row per upcoming event below the header.
 *                    "image" must be a link to a photo (see notes.txt for how to get one from Google Drive).
 * 5. In the Apps Script editor: Deploy > New deployment > gear icon > "Web app".
 *      - Execute as: Me
 *      - Who has access: Anyone
 *    Click Deploy, authorize when asked, then copy the "Web app URL" it gives you
 *    (looks like https://script.google.com/macros/s/AKfycb.../exec).
 * 6. Paste that URL into gas-config.js (replace the placeholder), re-upload the site files.
 *
 * That's it — every booking form submission becomes a new row in "Bookings", and every
 * row you add to "Events" (with a title, date, description, and image link) shows up
 * automatically on the Upcoming Events page.
 */

const BOOKINGS_SHEET_NAME = 'Bookings';
const EVENTS_SHEET_NAME = 'Events';
const BOOKINGS_HEADERS = ['Timestamp', 'Full Name', 'Contact Number', 'Email', 'Event Type', 'Date & Time Requested', 'Vision / Notes'];

// Handles the booking form (POST from the website).
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(BOOKINGS_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(BOOKINGS_SHEET_NAME);
      sheet.appendRow(BOOKINGS_HEADERS);
    }
    sheet.appendRow([
      new Date(),
      data.fullName || '',
      data.contactNumber || '',
      data.email || '',
      data.eventType || '',
      data.bookingSummary || '',
      data.vision || ''
    ]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Serves the "Events" tab as JSON (GET from the Upcoming Events page).
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(EVENTS_SHEET_NAME);
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({ events: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) {
    return ContentService.createTextOutput(JSON.stringify({ events: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const headers = rows.shift().map(h => String(h).trim().toLowerCase());
  const events = rows
    .filter(r => r.some(c => String(c).trim().length))
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = r[i] !== undefined ? String(r[i]) : ''; });
      return obj;
    });
  return ContentService.createTextOutput(JSON.stringify({ events }))
    .setMimeType(ContentService.MimeType.JSON);
}
