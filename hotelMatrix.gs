// ============================================================
// Hotel Rate Auto-Updater
// Estimates nightly hotel rates for each city using tiered
// pricing + live CAD→USD exchange rate + seasonal adjustment.
// Writes results to Source Data col B.
// ============================================================

const HM_SHEET_ID    = '1eX2xQ89R8IHTgJmdW7oukOK1aPgVMfYLHdox2hJ7wro';
const HM_SOURCE_TAB  = 'Source Data';
const HM_FALLBACK_RATE = 0.7233;  // CAD→USD fallback

// ── City base rates (USD/night) ──────────────────────────────
// Canadian cities are in CAD and converted using live rate.
// International cities use approximate USD equivalents.
const HOTEL_RATES_USD = {
  // ── Ultra-premium metros ─────────────────────────────────
  'New York, NY':          285,
  'Boston, MA':            230,
  'San Francisco, CA':     275,
  'Nashville, TN':         220,
  'Washington, DC':        210,
  'Miami, FL':             200,
  'Providence, RI':        175,

  // ── High-cost metros ────────────────────────────────────
  'Chicago, IL':           185,
  'Fort Lauderdale, FL':   185,
  'Seattle, WA':           195,
  'Cleveland, OH':         190,
  'Denver, CO':            180,
  'Austin, TX':            175,
  'Philadelphia, PA':      170,
  'Charlotte, NC':         170,
  'San Diego, CA':         175,
  'Portland, OR':          165,
  'Dallas, TX':            160,
  'Indianapolis, IN':      160,
  'New Orleans, LA':       160,
  'Salt Lake City, UT':    155,
  'Minneapolis, MN':       150,
  'Orlando, FL':           150,
  'Baltimore, MD':         150,
  'Atlantic City, NJ':     150,
  'Fort Worth, TX':        150,
  'Phoenix, AZ':           155,
  'Scottsdale, AZ':        160,
  'Las Vegas, NV':         130,
  'Portland, ME':          150,
  'Hershey, PA':           145,

  // ── Mid-tier cities ─────────────────────────────────────
  'Atlanta, GA':           140,
  'Pittsburgh, PA':        140,
  'Louisville, KY':        140,
  'Columbus, OH':          145,
  'St. Louis, MO':         135,
  'Cincinnati, OH':        135,
  'Tampa, FL':             145,
  'Virginia Beach, VA':    135,
  'Williamsburg, VA':      130,
  'Kansas City, MO':       130,
  'Des Moines, IA':        130,
  'Detroit, MI':           130,
  'Jacksonville, FL':      130,
  'Corpus Christi, TX':    130,
  'Anaheim, CA':           135,
  'Mt. Laurel, NJ':        130,
  'Sacramento, CA':        140,
  'Albuquerque, NM':       115,
  'Reno, NV':              115,
  'Omaha, NE':             120,
  'Houston, TX':           120,
  'San Antonio, TX':       125,
  'Lexington, KY':         125,
  'Hooksett, NH':          125,
  'Spokane, WA':           115,
  'Tucson, AZ':            105,
  'Waukesha, WI':          115,
  'Concord, NC':           120,

  // ── Smaller cities ──────────────────────────────────────
  'Birmingham, AL':        115,
  'Midland, TX':           115,
  'Loveland, CO':          115,
  'Grand Rapids, MI':      120,
  'Fort Wayne, IN':        110,
  'Lakeville, MN':         110,
  'Marietta, GA':          110,
  'Omaha, NE':             120,
  'Lafayette, IN':         100,
  'Muncie, IN':             90,

  // ── Caribbean ───────────────────────────────────────────
  'San Juan, PR':          155,

  // ── International (USD-equivalent estimates) ─────────────
  'London, UK':            250,
  'Berlin, Germany':       130,
  'Madrid, Spain':         150,
};

// Canadian cities — rates in CAD, converted to USD at live rate
const HOTEL_RATES_CAD = {
  'Toronto, ON':           205,
  'Mississauga, ON':       195,
  'Oakville, ON':          210,
  'Kitchener, ON':         185,
  'Vancouver, BC':         270,
  'Calgary, AB':           195,
  'Edmonton, AB':          160,
  'Ottawa, ON':            220,
  'Sudbury, ON':           155,
  'Petawawa, ON':          130,
  'Winnipeg, MB':          175,
  'Halifax, NS':           190,
  'Moncton, NB':           155,
};

// ── Seasonal multiplier ──────────────────────────────────────
// Adjusts rates up in peak travel months, down in slow months.
function getSeasonalFactor() {
  const month = new Date().getMonth(); // 0 = Jan
  const factors = [0.88, 0.88, 0.93, 0.97, 1.00, 1.08,
                   1.12, 1.10, 1.02, 0.97, 0.90, 0.92];
  return factors[month];
}

// ── Fetch live CAD→USD rate ──────────────────────────────────
function getHotelCadUsdRate() {
  try {
    const resp = UrlFetchApp.fetch(
      'https://api.frankfurter.app/latest?from=CAD&to=USD',
      { muteHttpExceptions: true }
    );
    if (resp.getResponseCode() === 200) {
      const rate = JSON.parse(resp.getContentText()).rates.USD;
      if (rate && rate > 0) return rate;
    }
  } catch(e) { Logger.log('Exchange rate fetch failed: ' + e); }
  return HM_FALLBACK_RATE;
}

// ── Main update function ─────────────────────────────────────
function updateHotelRates() {
  const ss       = SpreadsheetApp.openById(HM_SHEET_ID);
  const sheet    = ss.getSheetByName(HM_SOURCE_TAB);

  if (!sheet) {
    Logger.log('ERROR: Source Data tab not found.');
    return;
  }

  const cadUsd   = getHotelCadUsdRate();
  const seasonal = getSeasonalFactor();
  Logger.log(`CAD→USD: ${cadUsd} | Seasonal factor: ${seasonal.toFixed(2)}`);

  const lastRow  = sheet.getLastRow();
  const cityData = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  let updated = 0, skipped = 0;

  cityData.forEach((row, i) => {
    const city = row[0] ? row[0].toString().trim() : '';
    if (!city) return;

    let rateUsd = null;

    if (HOTEL_RATES_USD[city] !== undefined) {
      rateUsd = HOTEL_RATES_USD[city] * seasonal;
    } else if (HOTEL_RATES_CAD[city] !== undefined) {
      rateUsd = HOTEL_RATES_CAD[city] * cadUsd * seasonal;
    }

    if (rateUsd !== null) {
      const rounded = Math.round(rateUsd / 5) * 5; // round to nearest $5
      sheet.getRange(i + 2, 2).setValue(rounded);
      updated++;
    } else {
      Logger.log(`No rate defined for: ${city}`);
      skipped++;
    }
  });

  // Stamp last updated info
  const stampRow = lastRow + 2;
  const now      = new Date();
  sheet.getRange(stampRow, 1).setValue(
    `Hotel rates last updated: ${now.toDateString()} | CAD→USD: ${cadUsd.toFixed(4)} | Seasonal: ${(seasonal * 100).toFixed(0)}% | Updated: ${updated} cities | Skipped: ${skipped}`
  );

  Logger.log(`Done. Updated: ${updated} | Skipped: ${skipped}`);
  SpreadsheetApp.flush();
}

// ── Menu + triggers ──────────────────────────────────────────
function onOpenHotel() {
  SpreadsheetApp.getUi()
    .createMenu('🏨 Hotel Rates')
    .addItem('Update Now', 'updateHotelRates')
    .addItem('Set Daily Auto-Update', 'setupHotelDailyTrigger')
    .addItem('Remove Auto-Update', 'removeHotelTriggers')
    .addToUi();
}

function setupHotelDailyTrigger() {
  removeHotelTriggers();
  ScriptApp.newTrigger('updateHotelRates')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();
  SpreadsheetApp.getUi().alert('✅ Daily hotel rate update set for every day at 7am.');
}

function removeHotelTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'updateHotelRates')
    .forEach(t => ScriptApp.deleteTrigger(t));
}
