# Website logging via Google Sheets + Apps Script

Dit project bevat een Apps Script dat twee websites (leidscongresbureau en pitactief) elke minuut logt in een Google Spreadsheet, maar **uitsluitend** binnen een hard gedefinieerd logvenster.

## Installatie
1. Maak of open een Google Spreadsheet en geef het een naam.
2. Open **Extensions → Apps Script** en vervang de standaard `Code.gs` inhoud door de code uit [`apps-script.js`](apps-script.js).
3. Sla op. Kies **Run → onOpen** of herlaad de sheet zodat het menu **Monitoring** verschijnt.
4. Gebruik **Monitoring → Setup trigger (every minute)** om de tijdgestuurde trigger te installeren (autoriseer wanneer gevraagd). Gebruik **Monitoring → Remove trigger** om hem later te verwijderen.
5. Gebruik **Monitoring → Run now** voor een directe handmatige run.

## Logvenster instellen
Bovenaan `apps-script.js` staan de enige bron van waarheid voor het logvenster:
```javascript
const LOG_START_LOCAL = '2025-03-18 05:00'; // Logging window start (local time)
const LOG_END_LOCAL   = '2025-03-18 07:30'; // Logging window end (local time)
const LOCAL_TIMEZONE  = 'Europe/Amsterdam'; // Timezone for logging window
const SIGNATURE_CHECK_ENABLED = true; // Enable/disable body signature check
const EXPECTED_STRING_MAP = {
  'https://www.leidscongresbureau.nl/': 'wp-content',
  'https://www.pitactief.nl/': 'wp-content',
};
```
Pas deze drie constanten aan om het gewenste start- en eindmoment (in lokale tijd) te bepalen; buiten dit venster wordt niets gelogd.

Zet `SIGNATURE_CHECK_ENABLED` op `false` om de body-signaturecheck volledig uit te schakelen. Laat hem op `true` en pas `EXPECTED_STRING_MAP` aan als je per URL een andere verwachte substring (case-insensitive, op de eerste 5000 chars) wilt controleren.

## Werking
- De functie `runWebsiteCheck()` wordt via de trigger elke minuut gestart.
- Binnen het logvenster wordt per URL exact één rij toegevoegd aan de bijbehorende tab (`Logs_leidscongresbureau` of `Logs_pitactief`).
- De headers worden automatisch aangemaakt indien nog niet aanwezig.
- Er worden uitgebreide velden gelogd: statuscodes, timing, headers, body-sample, foutdetails enz.

## Test-checklist (handmatig)
- **Signature check**: zet `SIGNATURE_CHECK_ENABLED = true`, forceer een fetch (Monitoring → Run now) en controleer in de sheet dat `body_signature_ok` `TRUE/FALSE` wordt voor een 200-response afhankelijk van de aanwezigheid van de verwachte substring; zet daarna de flag op `false` en controleer dat het veld leeg blijft.
- **Header filtering**: kies een run met headers als `Set-Cookie` of `Report-To` (of voeg tijdelijk test-headers via een mock endpoint) en controleer dat deze niet in `other_headers_json_compact` verschijnen; `X-Powered-By` moet in de aparte kolom `x_powered_by` staan en de JSON moet afgekapt eindigen met `...truncated` wanneer hij erg lang is.
