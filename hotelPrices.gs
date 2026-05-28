// ============================================================
// Hotel Prices Web App
// Reads city names + hotel costs from Source Data tab and
// returns them as JSON so the travel calculator can load
// live rates instead of using hardcoded values.
//
// Deploy as: Execute as → Me, Who has access → Anyone
// ============================================================

const HOTEL_SHEET_ID  = '1eX2xQ89R8IHTgJmdW7oukOK1aPgVMfYLHdox2hJ7wro';
const HOTEL_SOURCE_TAB = 'Source Data';

function doGet() {
  try {
    const ss     = SpreadsheetApp.openById(HOTEL_SHEET_ID);
    const sheet  = ss.getSheetByName(HOTEL_SOURCE_TAB);

    if (!sheet) {
      return jsonResponse({ error: 'Source Data tab not found' }, 404);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return jsonResponse({ error: 'No data rows found' }, 404);
    }

    // Col A = city name, Col B = hotel cost ($/night)
    const data   = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    const rates  = {};

    data.forEach(([city, rate]) => {
      const name = city ? city.toString().trim() : '';
      const cost = parseFloat(rate);
      if (name && !isNaN(cost) && cost > 0) {
        rates[name] = cost;
      }
    });

    const payload = {
      updatedAt: new Date().toISOString(),
      source:    'Google Sheets – Source Data col B',
      rates
    };

    return jsonResponse(payload, 200);

  } catch (e) {
    return jsonResponse({ error: e.toString() }, 500);
  }
}

function jsonResponse(data, statusCode) {
  const output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}
