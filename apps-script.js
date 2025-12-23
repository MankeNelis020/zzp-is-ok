/**
 * Google Apps Script: Website logging tool for Sheets
 *
 * Configure the logging window here. Logging only occurs when the current local time
 * (in LOCAL_TIMEZONE) is between LOG_START_LOCAL and LOG_END_LOCAL (inclusive).
 */
const LOG_START_LOCAL = '2025-03-18 05:00'; // Logging window start (local time)
const LOG_END_LOCAL = '2025-03-18 07:30';   // Logging window end (local time)
const LOCAL_TIMEZONE = 'Europe/Amsterdam';   // Timezone used for logging window and display

/**
 * Optional body signature check configuration.
 * When SIGNATURE_CHECK_ENABLED is true and an expected string exists for the URL,
 * body_signature_ok will be set to true/false based on a case-insensitive match.
 */
const SIGNATURE_CHECK_ENABLED = true; // Set to false to disable signature checks entirely
const EXPECTED_STRING_MAP = {
  'https://www.leidscongresbureau.nl/': 'wp-content',
  'https://www.pitactief.nl/': 'wp-content',
};

/**
 * Entry point: checks both sites and appends logs when within the configured window.
 */
function runWebsiteCheck() {
  const now = new Date();
  if (!isWithinLoggingWindow(now)) {
    return;
  }

  const runId = Utilities.getUuid();
  const urls = [
    { url: 'https://www.leidscongresbureau.nl/', sheet: 'Logs_leidscongresbureau' },
    { url: 'https://www.pitactief.nl/', sheet: 'Logs_pitactief' },
  ];

  const rows = urls.map((item) => {
    const result = fetchUrlWithMetrics(item.url);
    result.values[0] = runId;
    return result;
  });

  rows.forEach((row) => {
    appendRows(row.sheetName, [row.values]);
  });
}

/**
 * Checks whether the provided date is within the hardcoded logging window.
 * @param {Date} now
 * @returns {boolean}
 */
function isWithinLoggingWindow(now) {
  const nowLocal = Utilities.formatDate(now, LOCAL_TIMEZONE, 'yyyy-MM-dd HH:mm');
  const startLocal = LOG_START_LOCAL;
  const endLocal = LOG_END_LOCAL;
  return nowLocal >= startLocal && nowLocal <= endLocal;
}

/**
 * Performs an HTTP fetch with metrics and structured logging fields.
 * @param {string} url
 * @returns {{sheetName: string, values: any[]}}
 */
function fetchUrlWithMetrics(url) {
  const started = Date.now();
  let response = null;
  let errorType = '';
  let errorMessage = '';

  try {
    response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      validateHttpsCertificates: true,
      headers: {
        'User-Agent': 'SheetsUptimeLogger/1.0',
        'Accept': '*/*',
      },
    });
  } catch (err) {
    const message = String(err && err.message ? err.message : err).substring(0, 500);
    errorMessage = message;
    if (/timed out/i.test(message)) {
      errorType = 'timeout';
    } else if (/dns|ssl|tls|certificate/i.test(message)) {
      errorType = 'dns_or_tls';
    } else {
      errorType = 'exception';
    }
  }

  const ended = Date.now();
  const responseTimeMs = ended - started;
  const now = new Date();
  const nowUtc = now.toISOString();
  const timestampLocal = Utilities.formatDate(now, LOCAL_TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");

  let status = '';
  let statusClass = '';
  let success = false;
  let headers = {};
  let bodySample200 = '';
  let responseSizeBytes = '';
  let redirectLocation = '';
  let contentType = '';
  let server = '';
  let dateHeader = '';
  let cacheControl = '';
  let expires = '';
  let age = '';
  let via = '';
  let xCache = '';
  let xServedBy = '';
  let xTimer = '';
  let xRequestId = '';
  let cfRay = '';
  let cfCacheStatus = '';
  let retryAfter = '';
  let locationHeader = '';
  let bodySignatureOk = '';
  let xPoweredBy = '';

  if (response) {
    status = response.getResponseCode();
    statusClass = getStatusClass(status);
    success = status >= 200 && status < 300;
    headers = response.getAllHeaders() || {};
    const normalizedHeaders = normalizeHeaders(headers);

    const contentBytes = response.getContent();
    responseSizeBytes = contentBytes ? contentBytes.length : '';
    const bodyText = response.getContentText();
    if (status === 200 && bodyText) {
      bodySample200 = bodyText.replace(/\s+/g, ' ').substring(0, 200);
      bodySignatureOk = computeSignatureCheck(url, bodyText);
    }

    redirectLocation = normalizedHeaders['location'] || '';
    contentType = normalizedHeaders['content-type'] || '';
    server = normalizedHeaders['server'] || '';
    dateHeader = normalizedHeaders['date'] || '';
    cacheControl = normalizedHeaders['cache-control'] || '';
    expires = normalizedHeaders['expires'] || '';
    age = normalizedHeaders['age'] || '';
    via = normalizedHeaders['via'] || '';
    xCache = normalizedHeaders['x-cache'] || '';
    xServedBy = normalizedHeaders['x-served-by'] || '';
    xTimer = normalizedHeaders['x-timer'] || '';
    xRequestId = normalizedHeaders['x-request-id'] || '';
    cfRay = normalizedHeaders['cf-ray'] || '';
    cfCacheStatus = normalizedHeaders['cf-cache-status'] || '';
    retryAfter = normalizedHeaders['retry-after'] || '';
    locationHeader = normalizedHeaders['location'] || '';
    xPoweredBy = normalizedHeaders['x-powered-by'] || '';
  }

  const otherHeadersJson = buildOtherHeadersJson(headers);

  return {
    sheetName: url.indexOf('leidscongresbureau') !== -1 ? 'Logs_leidscongresbureau' : 'Logs_pitactief',
    values: [
      '', // run_id placeholder; filled in caller
      nowUtc,
      timestampLocal,
      LOCAL_TIMEZONE,
      url,
      'GET',
      status,
      statusClass,
      success,
      responseTimeMs,
      responseSizeBytes,
      errorType,
      errorMessage,
      redirectLocation,
      bodySignatureOk,
      bodySample200,
      contentType,
      server,
      dateHeader,
      cacheControl,
      expires,
      age,
      via,
      xCache,
      xServedBy,
      xTimer,
      xRequestId,
      cfRay,
      cfCacheStatus,
      retryAfter,
      locationHeader,
      xPoweredBy,
      otherHeadersJson,
    ],
  };
}

/**
 * Appends rows to the specified sheet, ensuring headers exist.
 * @param {string} sheetName
 * @param {any[][]} rows
 */
function appendRows(sheetName, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  ensureHeaders(sheet);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

/**
 * Ensures the sheet has the expected headers in the first row.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function ensureHeaders(sheet) {
  const headers = [
    'run_id', 'timestamp_utc', 'timestamp_local', 'timezone', 'url_checked', 'method',
    'http_status', 'status_class', 'success', 'response_time_ms', 'response_size_bytes',
    'error_type', 'error_message', 'redirect_location', 'body_signature_ok', 'body_sample_200',
    'content_type', 'server_header', 'date_header', 'cache_control', 'expires', 'age',
    'via_header', 'x_cache', 'x_served_by', 'x_timer', 'x_request_id', 'cf_ray',
    'cf_cache_status', 'retry_after', 'location_header', 'x_powered_by', 'other_headers_json_compact'
  ];

  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const needsHeaders = headers.some((header, idx) => current[idx] !== header);
  if (needsHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

/**
 * Builds the HTTP status class bucket.
 * @param {number|string} status
 * @returns {string}
 */
function getStatusClass(status) {
  if (!status) return '';
  if (status >= 200 && status < 300) return '2xx';
  if (status >= 300 && status < 400) return '3xx';
  if (status >= 400 && status < 500) return '4xx';
  if (status >= 500 && status < 600) return '5xx';
  return '';
}

/**
 * Normalize headers to a lowercase key map for consistent lookups.
 * @param {Object} headers
 * @returns {Object}
 */
function normalizeHeaders(headers) {
  const normalized = {};
  Object.keys(headers || {}).forEach((key) => {
    const lower = key.toLowerCase();
    if (!normalized[lower]) {
      normalized[lower] = headers[key];
    }
  });
  return normalized;
}

/**
 * Computes the optional body signature check result.
 * Returns '' when disabled or when no expected string exists for the URL.
 * @param {string} url
 * @param {string} bodyText
 * @returns {boolean|string}
 */
function computeSignatureCheck(url, bodyText) {
  if (!SIGNATURE_CHECK_ENABLED) {
    return '';
  }
  const expected = EXPECTED_STRING_MAP[url];
  if (!expected) {
    return '';
  }
  const haystack = bodyText.toLowerCase().substring(0, 5000);
  const needle = String(expected).toLowerCase();
  return haystack.indexOf(needle) !== -1;
}

/**
 * Builds a compact JSON of remaining headers after filtering reserved and sensitive ones.
 * @param {Object} headers
 * @returns {string}
 */
function buildOtherHeadersJson(headers) {
  const reservedLower = new Set([
    'server', 'date', 'cache-control', 'expires', 'age', 'via', 'x-cache', 'x-served-by',
    'x-timer', 'x-request-id', 'cf-ray', 'cf-cache-status', 'retry-after', 'location',
    'content-type', 'x-powered-by'
  ]);

  const excludedLower = new Set([
    'report-to', 'nel', 'cf-nel', 'set-cookie', 'permissions-policy', 'cookie',
    'authorization', 'proxy-authorization'
  ]);

  const otherHeaders = {};
  Object.keys(headers || {}).forEach((key) => {
    const lower = key.toLowerCase();
    if (reservedLower.has(lower)) return;
    if (excludedLower.has(lower)) return;
    otherHeaders[key] = headers[key];
  });

  let otherHeadersJson = '';
  try {
    otherHeadersJson = JSON.stringify(otherHeaders);
    const limit = 1800; // within 1500-2000 chars requirement
    if (otherHeadersJson.length > limit) {
      otherHeadersJson = otherHeadersJson.substring(0, limit) + '...truncated';
    }
  } catch (err) {
    otherHeadersJson = '';
  }
  return otherHeadersJson;
}

/**
 * Adds custom menu for trigger management.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Monitoring')
    .addItem('Setup trigger (every minute)', 'setupTrigger')
    .addItem('Remove trigger', 'removeTrigger')
    .addItem('Run now', 'runWebsiteCheck')
    .addToUi();
}

/**
 * Removes existing triggers for runWebsiteCheck and creates a new one that runs every minute.
 */
function setupTrigger() {
  removeTrigger();
  ScriptApp.newTrigger('runWebsiteCheck').timeBased().everyMinutes(1).create();
}

/**
 * Removes all triggers for runWebsiteCheck.
 */
function removeTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => {
    if (trigger.getHandlerFunction() === 'runWebsiteCheck') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}
