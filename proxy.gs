/**
 * Travel Planner Proxy
 * Handles Calendar, Sheets, and Gemini API calls securely.
 */

const SPREADSHEET_ID = '1jghXxfXv8ia-S3re7PYwL3YR96jvcvyES33hb_KX-No';
const GEMINI_API_KEY = 'AIzaSyCAuo3hp8CVUyOmEWirVbbosQbSHISpTDo';

function doGet(e) {
  const action = e.parameter.action;
  
  try {
    if (action === 'getCalendar') {
      const events = CalendarApp.getDefaultCalendar().getEvents(
        new Date('2026-06-06T00:00:00Z'), 
        new Date('2026-06-27T23:59:59Z')
      );
      const result = events.map(ev => ({
        id: ev.getId(),
        summary: ev.getTitle(),
        location: ev.getLocation(),
        description: ev.getDescription(),
        start: { dateTime: ev.getStartTime().toISOString() }
      }));
      return jsonResponse(result);
    }
    
    if (action === 'getBudget') {
      const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheets()[0];
      const values = sheet.getDataRange().getValues();
      return jsonResponse(values);
    }

    if (action === 'getComments') {
      const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('comments');
      if (!sheet) return jsonResponse([]);
      const values = sheet.getDataRange().getValues();
      return jsonResponse(values);
    }
  } catch (err) {
    return jsonResponse({ error: err.message }, 500);
  }
}

function doPost(e) {
  const data = e.postData ? JSON.parse(e.postData.contents) : {};
  
  try {
    if (data.action === 'saveComment') {
      const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('comments') || 
                    SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet('comments');
      sheet.appendRow([data.key, data.summary, data.date, data.comment, new Date().toISOString()]);
      return jsonResponse({ success: true });
    }
    
    if (data.action === 'gemini') {
      const url = `https://generativelanguage.googleapis.com/v1/models/gemini-1.5-flash:generateContent?key=${GEMINI_API_KEY}`;
      const payload = {
        contents: [{ parts: [{ text: data.prompt }] }]
      };
      const response = UrlFetchApp.fetch(url, {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify(payload)
      });
      return jsonResponse(JSON.parse(response.getContentText()));
    }
  } catch (err) {
    return jsonResponse({ error: err.message }, 500);
  }
}

function jsonResponse(data, code = 200) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}