/**
 * ============================================================
 * FOOD TRACKER — Google Apps Script Backend
 * ============================================================
 *
 * SETUP INSTRUCTIONS:
 *
 * 1. Open your "Stomach Stuff" spreadsheet in Google Sheets.
 *
 * 2. Go to  Extensions > Apps Script  in the menu bar.
 *
 * 3. Delete any code already in the editor and paste this
 *    entire file in its place.
 *
 * 4. Click the 💾 Save icon (or Ctrl+S / Cmd+S).
 *
 * 5. Deploy as a web app:
 *    a. Click  Deploy > New deployment
 *    b. Click the ⚙ gear next to "Select type" and choose "Web app"
 *    c. Set:
 *         Description:  Food Tracker API
 *         Execute as:   Me (your-email@gmail.com)
 *         Who has access:  Anyone
 *    d. Click  Deploy
 *    e. Authorize the app when prompted (review permissions, allow).
 *    f. Copy the Web app URL that appears — it looks like:
 *       https://script.google.com/macros/s/XXXX.../exec
 *
 * 6. Open your index.html file and replace YOUR_APPS_SCRIPT_URL
 *    with the URL you just copied.
 *
 * That's it! The SPA can now write to your spreadsheet.
 *
 * NOTE: If you edit this script later, you must create a
 * NEW deployment (Deploy > New deployment) for changes to
 * take effect — or update the existing deployment version.
 * ============================================================
 */

/** Only this Google account is allowed to write data. */
var ALLOWED_EMAIL = 'dillmann.brian@gmail.com';

/**
 * Verify a Google ID token and check the email matches ALLOWED_EMAIL.
 * Returns the parsed token payload on success, or null on failure.
 */
function verifyToken(idToken) {
  if (!idToken) return null;
  try {
    var response = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + idToken,
      { muteHttpExceptions: true }
    );
    if (response.getResponseCode() !== 200) return null;
    var payload = JSON.parse(response.getContentText());
    if (payload.email === ALLOWED_EMAIL && payload.email_verified === 'true') {
      return payload;
    }
    return null;
  } catch (e) {
    return null;
  }
}

/**
 * Handle GET requests — simple health-check endpoint.
 */
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'Food Tracker API is running.' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Handle POST requests — append a row to the appropriate sheet.
 */
function doPost(e) {
  try {
    // Support both JSON (fetch) and form submission (iframe fallback)
    var data;
    if (e.postData && e.postData.type === 'application/x-www-form-urlencoded') {
      data = JSON.parse(e.parameter.payload);
    } else {
      data = JSON.parse(e.postData.contents);
    }

    // Verify the caller's identity before writing anything.
    var tokenPayload = verifyToken(data.idToken);
    if (!tokenPayload) {
      return _error('Unauthorized: invalid token or wrong account.');
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dateTime = data.dateTime || new Date().toISOString();

    if (data.type === 'meal') {
      var sheet = ss.getSheetByName('Meals');
      if (!sheet) {
        return _error('Sheet "Meals" not found. Make sure the tab is named exactly "Meals".');
      }
      sheet.appendRow([
        dateTime,
        data.mealType || '',
        data.contents || '',
        data.notes || ''
      ]);
      return _success('Meal logged successfully.');

    } else if (data.type === 'symptom') {
      var sheet = ss.getSheetByName('Symptoms');
      if (!sheet) {
        return _error('Sheet "Symptoms" not found. Make sure the tab is named exactly "Symptoms".');
      }
      sheet.appendRow([
        dateTime,
        data.severity || '',
        data.symptoms || '',
        data.notes || ''
      ]);
      return _success('Symptoms logged successfully.');

    } else {
      return _error('Unknown type: "' + data.type + '". Expected "meal" or "symptom".');
    }

  } catch (err) {
    return _error('Server error: ' + err.message);
  }
}

function _success(message) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: message }))
    .setMimeType(ContentService.MimeType.JSON);
}

function _error(message) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'error', message: message }))
    .setMimeType(ContentService.MimeType.JSON);
}
