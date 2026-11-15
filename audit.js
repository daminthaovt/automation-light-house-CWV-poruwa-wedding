import lighthouse from "lighthouse";
import * as chromeLauncher from "chrome-launcher";
import { google } from "googleapis";

const TARGET_URL = "https://weddings.poruwa.lk/";
const SPREADSHEET_ID = process.env.LIGHTHOUSE_SHEET_ID;
const TAB_NAME = "Lighthouse Audit Points - Porwua Wedding";
const NUM_COLS = 14; // Columns A through N (Application -> Agentic Browsing)

async function runAudit() {
  // 1. Run Headless Chrome + Lighthouse (including Agentic Browsing)
  const chrome = await chromeLauncher.launch({
    chromeFlags: ["--headless", "--no-sandbox"],
  });
  const options = {
    logLevel: "error",
    output: "json",
    port: chrome.port,
    onlyCategories: [
      "performance",
      "accessibility",
      "best-practices",
      "seo",
      "agentic-browsing",
    ],
  };

  const runnerResult = await lighthouse(TARGET_URL, options);
  await chrome.kill();

  const { audits, categories } = runnerResult.lhr;

  // Calculate Agentic Browsing pass fraction (e.g., "1/2")
  const agenticRefs = categories["agentic-browsing"]?.auditRefs || [];
  let passed = 0;
  let total = 0;
  for (const ref of agenticRefs) {
    const audit = audits[ref.id];
    if (audit && audit.scoreDisplayMode === "binary") {
      total++;
      if (audit.score === 1) passed++;
    }
  }
  const agenticFraction = `${passed}/${total}`;

  // 2. Authenticate with Google Sheets API
  const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  const sheets = google.sheets({ version: "v4", auth });

  // Get the internal numeric sheetId for TAB_NAME (needed for formatting/merging)
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const targetSheet = meta.data.sheets.find(
    (s) => s.properties.title === TAB_NAME,
  );
  if (!targetSheet) {
    throw new Error(`Tab "${TAB_NAME}" not found in spreadsheet.`);
  }
  const sheetId = targetSheet.properties.sheetId;

  // 3. Read Column A of ONLY TAB_NAME to calculate next Saturday–Friday week banner
  const existing = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${TAB_NAME}'!A:A`,
  });
  const colA = (existing.data.values || []).flat();
  const lastRowBefore = colA.length;

  // Format date in Sri Lanka time (Asia/Colombo, UTC+5:30)
  const formatDotDate = (dateObj) => {
    const parts = new Intl.DateTimeFormat("en-CA", {
      timeZone: "Asia/Colombo",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
    }).format(dateObj);
    return parts.replace(/-/g, ".");
  };

  const todayStr = formatDotDate(new Date());
  let nextWeekNum = 1;
  let periodStart = new Date(Date.UTC(2026, 8, 26)); // Anchored to 2026.09.26

  for (let i = colA.length - 1; i >= 0; i--) {
    const match = String(colA[i]).match(
      /Week\s+(\d+)\s*-\s*(\d{4}\.\d{2}\.\d{2})\s*-\s*(\d{4}\.\d{2}\.\d{2})/i,
    );
    if (match) {
      nextWeekNum = parseInt(match[1], 10) + 1;
      const [y, m, d] = match[3].split(".").map(Number);
      periodStart = new Date(Date.UTC(y, m - 1, d + 1)); // Saturday after previous Friday
      break;
    }
  }
  const periodEnd = new Date(periodStart.getTime() + 6 * 86400000); // Friday
  const weekPadded = String(nextWeekNum).padStart(2, "0");
  const bannerText = `Poruwa 2.0 Week ${weekPadded} - ${formatDotDate(periodStart)} - ${formatDotDate(periodEnd)} - Measured on ${todayStr}`;

  // 4. Append ONLY the 2 rows (Banner + Data row — no repeated header)
  const bannerRowIndex = lastRowBefore; // 0-based index for batchUpdate
  const dataRowIndex = lastRowBefore + 1; // 0-based index for batchUpdate

  const rows = [
    [bannerText],
    [
      "Poruwa Wedding",
      "Landing Page",
      "Production",
      "Google Lighthouse",
      audits["largest-contentful-paint"].displayValue,
      audits["first-contentful-paint"].displayValue,
      audits["total-blocking-time"].displayValue,
      audits["cumulative-layout-shift"].displayValue,
      audits["speed-index"].displayValue,
      Math.round(categories.performance.score * 100),
      Math.round(categories.accessibility.score * 100),
      Math.round(categories["best-practices"].score * 100),
      Math.round(categories.seo.score * 100),
      agenticFraction,
    ],
  ];

  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${TAB_NAME}'!A1`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: rows },
  });

  // 5. Copy exact colors, alignments, borders, merge, and dropdowns from Row 1 & Row 3
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [
        // Copy Banner formatting (Row 1 -> bannerRow)
        {
          copyPaste: {
            source: {
              sheetId,
              startRowIndex: 0,
              endRowIndex: 1,
              startColumnIndex: 0,
              endColumnIndex: NUM_COLS,
            },
            destination: {
              sheetId,
              startRowIndex: bannerRowIndex,
              endRowIndex: bannerRowIndex + 1,
              startColumnIndex: 0,
              endColumnIndex: NUM_COLS,
            },
            pasteType: "PASTE_FORMAT",
          },
        },
        // Merge Banner cells across A:N
        {
          mergeCells: {
            range: {
              sheetId,
              startRowIndex: bannerRowIndex,
              endRowIndex: bannerRowIndex + 1,
              startColumnIndex: 0,
              endColumnIndex: NUM_COLS,
            },
            mergeType: "MERGE_ALL",
          },
        },
        // Copy Data Row colors, borders & alignments (Row 3 -> dataRow)
        {
          copyPaste: {
            source: {
              sheetId,
              startRowIndex: 2,
              endRowIndex: 3,
              startColumnIndex: 0,
              endColumnIndex: NUM_COLS,
            },
            destination: {
              sheetId,
              startRowIndex: dataRowIndex,
              endRowIndex: dataRowIndex + 1,
              startColumnIndex: 0,
              endColumnIndex: NUM_COLS,
            },
            pasteType: "PASTE_FORMAT",
          },
        },
        // Copy Dropdowns (Poruwa Wedding, Landing Page, Production) from Row 3 -> dataRow
        {
          copyPaste: {
            source: {
              sheetId,
              startRowIndex: 2,
              endRowIndex: 3,
              startColumnIndex: 0,
              endColumnIndex: NUM_COLS,
            },
            destination: {
              sheetId,
              startRowIndex: dataRowIndex,
              endRowIndex: dataRowIndex + 1,
              startColumnIndex: 0,
              endColumnIndex: NUM_COLS,
            },
            pasteType: "PASTE_DATA_VALIDATION",
          },
        },
      ],
    },
  });

  console.log(`Successfully appended: ${bannerText}`);
}

runAudit();
