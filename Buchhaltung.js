// ==========================================
// BUCHHALTUNG: KONFIGURATION & KATEGORIEN
// ==========================================
const ACCOUNTING_YEAR = "2026";
const MAX_ACCOUNTING_RUNTIME_MS = 280000; // Sicherheitsstopp bei ~4 Min. 40 Sek. (vor dem 6-Min.-Limit)

const ACCOUNTING_CATEGORIES = [
  "Pflegeleistungen (§36 / §37 / §39 / §45b SGB XI)",
  "Beratungsbesuch (§37 Abs. 3 SGB XI)",
  "Krankenkasse / Pflegekasse (SBK)",
  "Berufshaftpflicht / BGW / DPV",
  "Kontoführung & Bankgebühren",
  "Website / Domain / Google Workspace",
  "Steuerberater / Anwalt / Behörde",
  "Büromaterial / Berufsbekleidung / Pflegematerial",
  "Auto (Finanzierung / Versicherung / Steuer)",
  "Auto (Benzin / Reparaturen / Parking)",
  "Werbungsmaterial & Porto",
  "Kooperationsdienstleistung",
  "Fortbildung & Sonderkosten",
  "Privatentnahmen",
  "Steuer / Rücklage Steuer",
  "Sonstiges"
];

const ACCOUNTING_SOURCE_PREFIXES = [
  "01_Finanzen"
];

const ACCOUNTING_EXCLUDE_SUBSTRINGS = [
  "Innaktive Kunden",
  "Volks Bank",
  "Honorarkalkulation",
  "Infos zum Steuer"
];

const MONTH_SHEET_NAMES = [
  "Januar", "Februar", "März", "April", "Mai", "Juni",
  "Juli", "August", "September", "Oktober", "November", "Dezember"
];


// ==========================================
// EIGENES MENÜ FÜR DIE BUCHHALTUNG
// ==========================================
function createAccountingMenu() {
  SpreadsheetApp.getUi()
    .createMenu("📊 Buchhaltung")
    .addItem("📂 Bestandsordner 2026 auslesen & Monatsblätter füllen", "scanExistingFoldersForAccounting")
    .addItem("📄 Kontoauszug-PDF aus Drive einlesen & abgleichen", "importBankStatementPdfsFromDrive")
    .addItem("🏦 Bankabgleich starten (Kontoauszug vs. Belege)", "reconcileBankStatement")
    .addItem("🔄 Aus Buchungsjournal in Monatsblätter synchronisieren", "syncJournalToMonthlySheets")
    .addSeparator()
    .addItem("🛠️ Tabellenstruktur reparieren (Einmalig)", "repairAllAccountingSheets")
    .addItem("🧹 Scan-Fortschritt zurücksetzen (Belege)", "resetAccountingScanLog")
    .addItem("🧹 Scan-Fortschritt zurücksetzen (Kontoauszug-PDFs)", "resetBankPdfScanLog")
    .addToUi();
}


// ==========================================
// PROMPT & SCHEMA FÜR DIE KI
// ==========================================
function getAccountingPromptExtension(companyName) {
  return `
--- BUCHHALTUNG ---
Firma/Name des Nutzers: "${companyName}"
Analysiere die Rechnungsdaten (Ignoriere Netto/MwSt, nutze ausschließlich BRUTTO):
- t: Belegtyp ("Ausgabe", "Einnahme" oder "Kein Beleg"). Rechnungen/Abrechnungen an Kunden oder Pflegekassen sind "Einnahme". Rechnungen/Quittungen an den Nutzer sind "Ausgabe". Kostenvoranschläge, Verträge, Abtretungserklärungen oder Kontoauszüge sind "Kein Beleg".
- d: Rechnungs-/Belegdatum im Format JJJJ-MM-DD (falls nicht im Text, nutze das Fallback-Datum).
- due: Fälligkeitsdatum im Format JJJJ-MM-DD (leer lassen, wenn keins angegeben).
- p: Geschäftspartner / Kunde (Name der Gegenseite oder des Pflegekunden).
- b: Rechnungsnummer (leer lassen, wenn keine gefunden).
- v: Brutto-Endbetrag als Zahl (z.B. 229.00). 0 eintragen, wenn kein Betrag gefunden wird.
- c: Währung (z.B. EUR).
- k: Passendste Kategorie aus dieser Liste: ${ACCOUNTING_CATEGORIES.join(", ")}
- s: Zahlungsstatus ("Bezahlt", "Offen" oder "Unbekannt").
`;
}

function getAccountingSchemaExtension() {
  return {
    "t": { "type": "STRING" },
    "d": { "type": "STRING" },
    "due": { "type": "STRING" },
    "p": { "type": "STRING" },
    "b": { "type": "STRING" },
    "v": { "type": "NUMBER" },
    "c": { "type": "STRING" },
    "k": { "type": "STRING" },
    "s": { "type": "STRING" }
  };
}


// ==========================================
// TABELLENBLÄTTER PRÜFEN, REPARIEREN & BUCHEN
// ==========================================
function setupAccountingSheets(ss) {
  getOrCreateBuchungsjournal(ss);
  if (typeof setupBankSheet === "function") setupBankSheet(ss);
}

function repairAllAccountingSheets() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    "Tabellenstruktur zurücksetzen?",
    "Achtung: Dadurch werden 'Buchungsjournal', 'Jahresübersicht 2026' und alle 12 Monatsblätter geleert und neu aufgebaut. Fortfahren?",
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Buchungsjournal sauber in Zeile 1 (Spalten A bis M) aufbauen
  let journal = ss.getSheetByName("Buchungsjournal");
  if (!journal) journal = ss.insertSheet("Buchungsjournal");
  journal.clear();
  const jHeaders = [
    "Monat (Auswertung)", "Belegdatum", "Typ", "Geschäftspartner",
    "Rechnungs-Nr.", "Brutto-Betrag", "Währung", "Kategorie",
    "Zahlungsstatus", "Beleg-Link", "Fälligkeitsdatum", "Datei-ID", "Token-Verbrauch"
  ];
  journal.getRange(1, 1, 1, jHeaders.length).setValues([jHeaders])
    .setFontWeight("bold").setBackground("#d9ead3");
  journal.setFrozenRows(1);
  journal.getRange("F:F").setNumberFormat("#,##0.00 €");

  // 2. Alle 12 Monatsblätter (Januar - Dezember) sauber in Zeile 3 aufbauen
  const monthHeaders = [
    "Datum", "Einnahmen", "Ausgaben", "Beschreibung", "Saldo", "",
    "Kunde / Einnahmequelle", "Rechnungs-Nr.", "Fälligkeits-Datum",
    "Rechnungsbetrag", "Zahlungseingang am", "Eingezahlter Betrag",
    "Zahlungserinnerung", "1. Mahnung", "2. Mahnung"
  ];

  MONTH_SHEET_NAMES.forEach(monthName => {
    let ws = ss.getSheetByName(monthName) || ss.getSheetByName("Kopie von " + monthName);
    if (!ws) {
      ws = ss.insertSheet(monthName);
    } else if (ws.getName() !== monthName) {
      ws.setName(monthName);
    }

    ws.clear();
    ws.getRange("A1:O1").merge()
      .setValue(`Monatsabrechnung - ${monthName} ${ACCOUNTING_YEAR}`)
      .setFontSize(14).setFontWeight("bold");

    ws.getRange(3, 1, 1, monthHeaders.length).setValues([monthHeaders])
      .setFontWeight("bold").setBackground("#333333").setFontColor("#ffffff");
    ws.getRange(3, 6).setBackground(null);
    ws.setFrozenRows(3);

    ws.getRange("A4:A500").setNumberFormat("dd.mm.yyyy");
    ws.getRange("I4:I500").setNumberFormat("dd.mm.yyyy");
    ws.getRange("K4:K500").setNumberFormat("dd.mm.yyyy");
    ws.getRange("B4:C500").setNumberFormat("#,##0.00 €");
    ws.getRange("E4:E500").setNumberFormat("#,##0.00 €");
    ws.getRange("J4:J500").setNumberFormat("#,##0.00 €");
    ws.getRange("L4:L500").setNumberFormat("#,##0.00 €");
    ws.setColumnWidth(6, 25);
  });

  // 3. Jahresübersicht 2026 sauber verknüpfen
  let dash = ss.getSheetByName(`Jahresübersicht ${ACCOUNTING_YEAR}`);
  if (!dash) dash = ss.insertSheet(`Jahresübersicht ${ACCOUNTING_YEAR}`);
  dash.clear();
  dash.getRange("A1:D1").setValues([["Monat", "Gesamteinnahmen", "Gesamtausgaben", "Saldo (Gewinn/Verlust)"]])
    .setFontWeight("bold").setBackground("#333333").setFontColor("#ffffff");

  MONTH_SHEET_NAMES.forEach((m, idx) => {
    const r = idx + 2;
    dash.getRange(r, 1).setValue(m);
    dash.getRange(r, 2).setFormula(`=SUM('${m}'!B4:B500)`);
    dash.getRange(r, 3).setFormula(`=SUM('${m}'!C4:C500)`);
    dash.getRange(r, 4).setFormula(`=B${r}-C${r}`);
  });
  dash.getRange(14, 1).setValue("GESAMT").setFontWeight("bold");
  dash.getRange(14, 2).setFormula("=SUM(B2:B13)").setFontWeight("bold");
  dash.getRange(14, 3).setFormula("=SUM(C2:C13)").setFontWeight("bold");
  dash.getRange(14, 4).setFormula("=SUM(D2:D13)").setFontWeight("bold");
  dash.getRange("B2:D14").setNumberFormat("#,##0.00 €");
  dash.setFrozenRows(1);

  if (typeof setupBankSheet === "function") setupBankSheet(ss);

  ss.toast("Alle Buchhaltungs-Blätter wurden repariert und verknüpft!", "Reparatur abgeschlossen ✓", 8);
}

function getOrCreateBuchungsjournal(ss) {
  let sheet = ss.getSheetByName("Buchungsjournal");
  const headers = [
    "Monat (Auswertung)", "Belegdatum", "Typ", "Geschäftspartner",
    "Rechnungs-Nr.", "Brutto-Betrag", "Währung", "Kategorie",
    "Zahlungsstatus", "Beleg-Link", "Fälligkeitsdatum", "Datei-ID", "Token-Verbrauch"
  ];

  if (!sheet) {
    sheet = ss.insertSheet("Buchungsjournal");
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#d9ead3");
    sheet.setFrozenRows(1);
    sheet.getRange("F:F").setNumberFormat("#,##0.00 €");
  } else {
    const currentLastCol = sheet.getLastColumn();
    if (currentLastCol === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#d9ead3");
      sheet.setFrozenRows(1);
      sheet.getRange("F:F").setNumberFormat("#,##0.00 €");
    } else if (currentLastCol < headers.length) {
      const existingHeaders = sheet.getRange(1, 1, 1, currentLastCol).getValues()[0];
      headers.forEach(h => {
        if (!existingHeaders.includes(h)) {
          const nextCol = sheet.getLastColumn() + 1;
          sheet.getRange(1, nextCol).setValue(h).setFontWeight("bold").setBackground("#d9ead3");
        }
      });
    }
  }
  return sheet;
}

function getMonthSheet(ss, monthIndex) {
  const baseName = MONTH_SHEET_NAMES[monthIndex];
  return ss.getSheetByName(baseName) || ss.getSheetByName("Kopie von " + baseName);
}

function normalizeDateToIso(val) {
  if (!val) return "";
  const parsed = typeof parseFlexibleDate === "function" ? parseFlexibleDate(val) : parseIsoDate(val);
  if (parsed && !isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, "Europe/Berlin", "yyyy-MM-dd");
  }
  return "";
}

function parseAccountingAmount(val) {
  if (typeof val === "number") return val;
  if (typeof parseGermanAmount === "function") return parseGermanAmount(val);
  if (!val) return 0;
  let str = val.toString().replace(/[€EUR\s]/gi, "").trim();
  if (str.includes(".") && str.includes(",")) str = str.replace(/\./g, "").replace(",", ".");
  else if (str.includes(",")) str = str.replace(",", ".");
  const num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}

function bookApprovedInvoice(ss, data) {
  const numAmount = parseAccountingAmount(data ? data.amount : 0);
  if (!data || data.type === "Kein Beleg" || data.type === "" || numAmount <= 0) return;

  const journal = getOrCreateBuchungsjournal(ss);
  const cleanFileId = (data.fileId || "").toString().trim();
  const cleanInvNr = (data.invNr || "").toString().trim();
  const cleanPartner = (data.partner || "").toString().trim();

  if (journal.getLastRow() > 1) {
    const jCols = Math.max(journal.getLastColumn(), 12);
    const jData = journal.getRange(2, 1, journal.getLastRow() - 1, jCols).getValues();
    for (let r = 0; r < jData.length; r++) {
      const existingFileId = (jData[r][11] || "").toString().trim();
      const existingInvNr = (jData[r][4] || "").toString().trim();
      const existingPartner = (jData[r][3] || "").toString().trim();
      if (cleanFileId && existingFileId === cleanFileId) return;
      if (cleanInvNr && cleanInvNr.length >= 4 && existingInvNr === cleanInvNr && existingPartner === cleanPartner) return;
    }
  }

  const isoDateStr = normalizeDateToIso(data.bDate) || Utilities.formatDate(new Date(), "Europe/Berlin", "yyyy-MM-dd");
  const isoDueStr = normalizeDateToIso(data.dueDate);
  const monthStr = isoDateStr.substring(0, 7);

  data.bDate = isoDateStr;
  data.dueDate = isoDueStr;

  journal.appendRow([
    monthStr, isoDateStr, data.type, cleanPartner, cleanInvNr,
    numAmount, data.currency || "EUR", data.category || "Sonstiges",
    data.payStatus || "Offen", data.link || "", isoDueStr, cleanFileId, data.tokens || 0
  ]);

  if (monthStr.startsWith(ACCOUNTING_YEAR + "-")) {
    writeEntryToMonthlySheet(ss, data, monthStr);
  }
}

function writeEntryToMonthlySheet(ss, data, monthStr) {
  const parts = monthStr.split("-");
  if (parts.length < 2) return;
  if (parts[0] !== ACCOUNTING_YEAR) return;

  const monthIndex = parseInt(parts[1], 10) - 1;
  if (isNaN(monthIndex) || monthIndex < 0 || monthIndex > 11) return;

  const monthSheet = getMonthSheet(ss, monthIndex);
  if (!monthSheet) return;

  const dateObj = (typeof parseFlexibleDate === "function" ? parseFlexibleDate(data.bDate) : parseIsoDate(data.bDate)) || new Date();
  const dueObj = (typeof parseFlexibleDate === "function" ? parseFlexibleDate(data.dueDate) : parseIsoDate(data.dueDate)) || "";
  const amount = parseAccountingAmount(data.amount);
  const isIncome = data.type === "Einnahme";

  const descParts = [data.partner];
  if (data.invNr) descParts.push(`Rechn. ${data.invNr}`);
  if (data.category && !isIncome) descParts.push(`(${data.category})`);
  const description = descParts.filter(Boolean).join(" – ");

  // Linke Tabelle (Spalten A-E): nächste freie Zeile suchen
  const maxAvailableRows = Math.max(1, monthSheet.getMaxRows() - 3);
  const scanLimit = Math.min(Math.max(monthSheet.getLastRow() - 3, 50), maxAvailableRows);
  const leftValues = monthSheet.getRange(4, 1, scanLimit, 4).getValues();
  let targetLeftRow = 4;
  for (let i = 0; i < leftValues.length; i++) {
    if (leftValues[i][0] === "" && leftValues[i][1] === "" && leftValues[i][2] === "" && leftValues[i][3] === "") {
      targetLeftRow = i + 4;
      break;
    }
    targetLeftRow = i + 5;
  }

  if (targetLeftRow > monthSheet.getMaxRows()) {
    monthSheet.insertRowsAfter(monthSheet.getMaxRows(), 10);
  }

  monthSheet.getRange(targetLeftRow, 1, 1, 4).setValues([[
    dateObj,
    isIncome ? amount : "",
    !isIncome ? amount : "",
    description
  ]]);
  monthSheet.getRange(targetLeftRow, 1).setNumberFormat("dd.mm.yyyy");
  monthSheet.getRange(targetLeftRow, 2, 1, 2).setNumberFormat("#,##0.00 €");

  const saldoFormula = (targetLeftRow === 4)
    ? "=B4-C4"
    : `=E${targetLeftRow - 1}+B${targetLeftRow}-C${targetLeftRow}`;
  monthSheet.getRange(targetLeftRow, 5).setFormula(saldoFormula).setNumberFormat("#,##0.00 €");

  // Rechte Tabelle (Spalten G-L) nur bei Einnahmen
  if (isIncome) {
    const rightValues = monthSheet.getRange(4, 7, scanLimit, 4).getValues();
    let targetRightRow = 4;
    for (let i = 0; i < rightValues.length; i++) {
      if (rightValues[i][0] === "" && rightValues[i][1] === "" && rightValues[i][3] === "") {
        targetRightRow = i + 4;
        break;
      }
      targetRightRow = i + 5;
    }

    if (targetRightRow > monthSheet.getMaxRows()) {
      monthSheet.insertRowsAfter(monthSheet.getMaxRows(), 10);
    }

    const isPaid = (data.payStatus || "").toLowerCase() === "bezahlt";
    monthSheet.getRange(targetRightRow, 7, 1, 6).setValues([[
      data.partner || "",
      data.invNr || "",
      dueObj,
      amount,
      isPaid ? dateObj : "",
      isPaid ? amount : 0
    ]]);
    if (dueObj) monthSheet.getRange(targetRightRow, 9).setNumberFormat("dd.mm.yyyy");
    monthSheet.getRange(targetRightRow, 10).setNumberFormat("#,##0.00 €");
    if (isPaid) monthSheet.getRange(targetRightRow, 11).setNumberFormat("dd.mm.yyyy");
    monthSheet.getRange(targetRightRow, 12).setNumberFormat("#,##0.00 €");
  }
}

function parseIsoDate(dateStr) {
  if (!dateStr) return null;
  if (dateStr instanceof Date) return dateStr;
  const str = dateStr.toString().trim();
  const match = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) {
    return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  }
  return null;
}


// ==========================================
// BESTANDS-SCANNER: EINZEL-VERARBEITUNG ALLER BELEGE
// ==========================================
function notifyAccountingProgress(ss, message, title = "📊 Buchhaltung", seconds = 5) {
  if (typeof updateLiveStatus === "function") {
    updateLiveStatus(ss, message, title, seconds);
  } else if (ss) {
    ss.toast(message, title, seconds);
  }
}

function scanExistingFoldersForAccounting() {
  const startTime = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const whitelistSheet = ss.getSheetByName("Ordner_Whitelist");
  if (!whitelistSheet || whitelistSheet.getLastRow() < 2) {
    ss.toast("Bitte zuerst in 'Ordner_Whitelist' die Ordner prüfen.", "Hinweis", 6);
    return;
  }

  let apiKey;
  try {
    apiKey = getApiKey();
  } catch (e) {
    ss.toast(e.message, "API-Schlüssel fehlt ⚠️", 8);
    return;
  }

  setupAccountingSheets(ss);

  const props = PropertiesService.getScriptProperties();
  const scannedRaw = props.getProperty("ACCOUNTING_SCANNED_FILE_IDS");
  const scannedIds = new Set(scannedRaw ? JSON.parse(scannedRaw) : []);

  // Bereits im Buchungsjournal erfasste Datei-IDs laden
  const journalSheet = ss.getSheetByName("Buchungsjournal");
  if (journalSheet && journalSheet.getLastRow() > 1) {
    const jRows = journalSheet.getRange(2, 12, journalSheet.getLastRow() - 1, 1).getValues();
    jRows.forEach(r => {
      const id = (r[0] || "").toString().trim();
      if (id) scannedIds.add(id);
    });
  }

  const wlData = whitelistSheet.getRange(2, 1, whitelistSheet.getLastRow() - 1, 2).getValues();
  const targetFolders = [];
  const seenFolderIds = new Set();
  const seedFolders = [];

  wlData.forEach(row => {
    const path = (row[0] || "").toString().trim().replace(/\s*\/\s*/g, " / ");
    const id = (row[1] || "").toString().trim();
    if (!path || !id || id.length < 10) return;

    const matchesPrefix = ACCOUNTING_SOURCE_PREFIXES.some(prefix => path.startsWith(prefix));
    const isExcluded = ACCOUNTING_EXCLUDE_SUBSTRINGS.some(ex => path.includes(ex));
    if (matchesPrefix && !isExcluded) {
      seedFolders.push({ path, id });
    }
  });

  seedFolders.forEach(seed => {
    try {
      const folder = DriveApp.getFolderById(seed.id);
      if (!folder.isTrashed()) {
        collectAccountingSubfoldersDirect(folder, seed.path, targetFolders, seenFolderIds);
      }
    } catch (e) {}
  });

  if (targetFolders.length === 0) {
    ss.toast("Keine Ordner unter '01_Finanzen' gefunden. Bitte zuerst '0. Ordner & Unterordner einlesen' ausführen.", "Hinweis", 6);
    return;
  }

  notifyAccountingProgress(ss, `Durchsuche ${targetFolders.length} Ordner unter '01_Finanzen' nach neuen Dateien...`, "📊 Bestands-Scan");

  // Alle neuen Dateien ohne Mengenbegrenzung einsammeln
  const filesQueue = [];
  const seenInQueue = new Set();

  for (let f = 0; f < targetFolders.length; f++) {
    const { path, id } = targetFolders[f];
    try {
      const folder = DriveApp.getFolderById(id);
      const files = folder.getFiles();
      while (files.hasNext()) {
        const file = files.next();
        const fileId = file.getId();
        if (scannedIds.has(fileId) || seenInQueue.has(fileId)) continue;
        seenInQueue.add(fileId);
        filesQueue.push({ file, folderPath: path });
      }
    } catch (folderErr) {}
  }

  if (filesQueue.length === 0) {
    notifyAccountingProgress(ss, "Alle Belege unter '01_Finanzen' wurden bereits ausgelesen!", "Alles aktuell ✓", 8);
    return;
  }

  notifyAccountingProgress(ss, `${filesQueue.length} neue Datei(en) gefunden. Starte Einzel-Auslesung...`, "📊 Bestands-Scan");

  let models = ["models/gemini-2.5-flash"];
  if (typeof getActiveGeminiModels === "function") {
    try { models = getActiveGeminiModels(ss, "Buchhaltung"); } catch (e) {}
  }

  let bookedCount = 0;
  let skippedNonInvoices = 0;
  let errorCount = 0;
  let totalTokensRun = 0;
  let processedCount = 0;
  let stoppedByTimeout = false;

  // JEDE Datei einzeln vorbereiten, an Gemini senden und sofort verbuchen
  for (let i = 0; i < filesQueue.length; i++) {
    if (Date.now() - startTime > MAX_ACCOUNTING_RUNTIME_MS) {
      stoppedByTimeout = true;
      notifyAccountingProgress(
        ss,
        `Zeitlimit (4 Min. 40 Sek.) erreicht. ${filesQueue.length - processedCount} Datei(en) verbleiben für den nächsten Lauf.`,
        "⏱️ Sicherheitspause",
        8
      );
      break;
    }

    const { file, folderPath } = filesQueue[i];
    const fileId = file.getId();
    const fileName = file.getName();
    const stepPrefix = `📊 Beleg ${i + 1}/${filesQueue.length}`;

    notifyAccountingProgress(ss, `[1/3] Lese Inhalt aus "${fileName}" (${folderPath})...`, stepPrefix, 5);

    const payload = buildAccountingPayloadForFile(file, folderPath);
    if (!payload) {
      scannedIds.add(fileId);
      skippedNonInvoices++;
      processedCount++;
      notifyAccountingProgress(ss, `[2/3] "${fileName}" nicht auslesbar – übersprungen.`, stepPrefix, 4);
      continue;
    }

    notifyAccountingProgress(ss, `[2/3] Sende "${fileName}" an Gemini KI...`, stepPrefix, 6);

    const res = sendAccountingRequest(payload, models, apiKey, ss, stepPrefix, fileName);
    if (!res || res.getResponseCode() !== 200) {
      errorCount++;
      processedCount++;
      notifyAccountingProgress(ss, `[3/3] API-Fehler bei "${fileName}" – wird beim nächsten Lauf erneut geprüft.`, stepPrefix, 5);
      continue;
    }

    let extracted;
    let usage = {};
    try {
      const json = JSON.parse(res.getContentText());
      const rawText = json.candidates && json.candidates[0] && json.candidates[0].content &&
                      json.candidates[0].content.parts && json.candidates[0].content.parts[0] &&
                      json.candidates[0].content.parts[0].text;
      if (!rawText) {
        errorCount++;
        processedCount++;
        continue;
      }
      const cleanJson = rawText.replace(/^\s*```(?:json)?\s*/i, "").replace(/\s*```\s*$/i, "").trim();
      extracted = JSON.parse(cleanJson);
      usage = json.usageMetadata || {};
    } catch (parseErr) {
      errorCount++;
      processedCount++;
      continue;
    }

    const totalTok = usage.totalTokenCount || 0;
    const tokenDisplay = `${totalTok} (In:${usage.promptTokenCount || 0} Out:${usage.candidatesTokenCount || 0})`;

    totalTokensRun += totalTok;
    scannedIds.add(fileId);

    const invAmount = parseAccountingAmount(extracted.v);
    if (extracted.t && extracted.t !== "Kein Beleg" && invAmount > 0) {
      const fallbackDate = Utilities.formatDate(file.getDateCreated(), "Europe/Berlin", "yyyy-MM-dd");
      const linkFormula = `=HYPERLINK("${file.getUrl()}", "${fileName.replace(/"/g, '""')}")`;
      bookApprovedInvoice(ss, {
        link: linkFormula,
        type: extracted.t,
        bDate: extracted.d || fallbackDate,
        dueDate: extracted.due || "",
        partner: extracted.p || fileName,
        invNr: extracted.b || "",
        amount: invAmount,
        currency: extracted.c || "EUR",
        category: extracted.k || "Sonstiges",
        payStatus: extracted.s || "Offen",
        fileId: fileId,
        tokens: tokenDisplay
      });
      SpreadsheetApp.flush();
      bookedCount++;
      notifyAccountingProgress(
        ss,
        `[3/3] Verbucht: ${extracted.t} | ${extracted.p || fileName} | ${invAmount.toFixed(2)} € (${tokenDisplay})`,
        stepPrefix,
        5
      );
    } else {
      skippedNonInvoices++;
      notifyAccountingProgress(
        ss,
        `[3/3] Übersprungen (${extracted.t || "Kein Beleg"}): "${fileName}" (${tokenDisplay})`,
        stepPrefix,
        4
      );
    }

    processedCount++;

    if (i + 1 < filesQueue.length) {
      Utilities.sleep(300);
    }
  }

  // Fortschritt sicher in Properties puffern (max. 150 IDs, um 9-KB-Limit nicht zu verletzen)
  try {
    const safeRecentIds = [...scannedIds].slice(-150);
    props.setProperty("ACCOUNTING_SCANNED_FILE_IDS", JSON.stringify(safeRecentIds));
  } catch (propErr) {}

  const elapsedSec = ((Date.now() - startTime) / 1000).toFixed(1);
  const errInfo = errorCount > 0 ? ` | ❌ ${errorCount} Fehler` : "";
  const remInfo = stoppedByTimeout ? ` | ⏱️ ${filesQueue.length - processedCount} verbleibend` : "";

  notifyAccountingProgress(
    ss,
    `Fertig in ${elapsedSec}s: ${bookedCount} Beleg(e) gebucht, ${skippedNonInvoices} übersprungen (${totalTokensRun} Tokens)${errInfo}${remInfo}`,
    "✅ Bestands-Scan abgeschlossen",
    10
  );
}

function collectAccountingSubfoldersDirect(folder, currentPath, targetList, seenIds) {
  if (ACCOUNTING_EXCLUDE_SUBSTRINGS.some(ex => currentPath.includes(ex))) return;
  if (!seenIds.has(folder.getId())) {
    seenIds.add(folder.getId());
    targetList.push({ path: currentPath, id: folder.getId() });
  }
  const subs = folder.getFolders();
  while (subs.hasNext()) {
    const sub = subs.next();
    if (!sub.isTrashed()) {
      collectAccountingSubfoldersDirect(sub, `${currentPath} / ${sub.getName()}`, targetList, seenIds);
    }
  }
}

function sendAccountingRequest(basePayload, models, apiKey, ss, stepInfo, fileName) {
  for (let m = 0; m < models.length; m++) {
    const modelName = models[m];
    const shortName = modelName.replace("models/", "");
    const maxRetries = (m === 0) ? 3 : 2;

    const reqPayload = JSON.parse(JSON.stringify(basePayload));
    if (!reqPayload.generationConfig) reqPayload.generationConfig = {};

    const lowerModel = shortName.toLowerCase();
    let appliedCustomConfig = false;
    if (lowerModel.includes("2.5") || lowerModel.includes("thinking")) {
      reqPayload.generationConfig.thinkingConfig = { thinkingBudget: 0 };
      appliedCustomConfig = true;
    }
    if (reqPayload.generationConfig.mediaResolution) appliedCustomConfig = true;

    const url = `https://generativelanguage.googleapis.com/v1beta/${modelName}:generateContent?key=${apiKey}`;

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        if (attempt > 1 || m > 0) {
          notifyAccountingProgress(
            ss,
            `Wiederhole Anfrage (${shortName}, Versuch ${attempt}/${maxRetries}) für "${fileName}"...`,
            stepInfo,
            4
          );
        }

        const res = UrlFetchApp.fetch(url, {
          method: "post",
          contentType: "application/json",
          payload: JSON.stringify(reqPayload),
          muteHttpExceptions: true
        });
        const code = res.getResponseCode();

        if (code === 200) return res;

        if (code === 400 && appliedCustomConfig) {
          delete reqPayload.generationConfig.thinkingConfig;
          delete reqPayload.generationConfig.mediaResolution;
          appliedCustomConfig = false;
          continue;
        }

        if ((code === 429 || code === 503) && attempt < maxRetries) {
          const waitSec = attempt * 2;
          notifyAccountingProgress(ss, `API-Limit (HTTP ${code}) – warte ${waitSec}s...`, stepInfo, waitSec + 2);
          Utilities.sleep(waitSec * 1000);
          continue;
        }
        break;
      } catch (netErr) {
        if (attempt < maxRetries) Utilities.sleep(1500);
      }
    }
  }
  return null;
}

function buildAccountingPayloadForFile(file, folderPath) {
  const lowerName = file.getName().toLowerCase();
  const driveMime = (file.getMimeType() || "").toLowerCase();
  const blobMime = (file.getBlob().getContentType() || "").toLowerCase();
  const fallbackDate = Utilities.formatDate(file.getDateCreated(), "Europe/Berlin", "yyyy-MM-dd");
  const maxChars = typeof MAX_EXTRACTED_CHARS !== "undefined" ? MAX_EXTRACTED_CHARS : 2000;

  let folderHint = "Hinweis aus Ordnerpfad: Bestimme Einnahme oder Ausgabe anhand des Dokumentinhalts.";
  if (folderPath.includes("Belege Ausgabe") || folderPath.toLowerCase().includes("ausgabe")) {
    folderHint = "Hinweis aus Ordnerpfad: Dies liegt im AUSGABEN-Ordner.";
  } else if (folderPath.toLowerCase().includes("kunde") || folderPath.toLowerCase().includes("einnahme") || folderPath.toLowerCase().includes("abrechnung")) {
    folderHint = "Hinweis aus Ordnerpfad: Dies liegt im KUNDEN-ABRECHNUNGS-Ordner (Einnahme, sofern Rechnung/Abrechnung).";
  }

  const basePrompt = `Extrahiere Buchhaltungsdaten aus diesem Dokument.
Dateiname: "${file.getName()}" (Fallback-Datum: ${fallbackDate})
Speicherort in Drive: "${folderPath}"
${folderHint}
${getAccountingPromptExtension(typeof COMPANY_NAME !== "undefined" ? COMPANY_NAME : "Pflegedienst")}`;

  const schemaProps = getAccountingSchemaExtension();
  const genConfig = {
    response_mime_type: "application/json",
    response_schema: {
      type: "OBJECT",
      properties: schemaProps,
      required: Object.keys(schemaProps)
    }
  };

  const isGoogleDoc = driveMime === "application/vnd.google-apps.document";
  const isGoogleSheet = driveMime === "application/vnd.google-apps.spreadsheet";
  const isWord = driveMime.includes("wordprocessingml") || driveMime === "application/msword" || lowerName.endsWith(".docx") || lowerName.endsWith(".doc");
  const isExcel = driveMime.includes("spreadsheetml") || driveMime === "application/vnd.ms-excel" || lowerName.endsWith(".xlsx") || lowerName.endsWith(".xls");
  const isPdf = driveMime === "application/pdf" || blobMime === "application/pdf" || lowerName.endsWith(".pdf");
  const isImage = driveMime.startsWith("image/") || /\.(jpg|jpeg|png|webp|heic)$/i.test(lowerName);

  let textContent = "";
  if (isGoogleDoc) {
    try {
      const res = UrlFetchApp.fetch(`https://www.googleapis.com/drive/v3/files/${file.getId()}/export?mimeType=text/plain`, {
        headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
      });
      textContent = res.getResponseCode() === 200
        ? res.getContentText("UTF-8").replace(/\s+/g, " ").trim().substring(0, maxChars)
        : DocumentApp.openById(file.getId()).getBody().getText().replace(/\s+/g, " ").trim().substring(0, maxChars);
    } catch (e) {}
  } else if (isGoogleSheet) {
    try {
      const vals = SpreadsheetApp.openById(file.getId()).getDataRange().getDisplayValues();
      textContent = vals.slice(0, 35).map(r => r.join(" | ")).join("\n").replace(/\s+/g, " ").trim().substring(0, maxChars);
    } catch (e) {}
  } else if ((isWord || isExcel) && typeof extractTextFromOfficeFile === "function") {
    try { textContent = extractTextFromOfficeFile(file, isWord); } catch (e) {}
  } else if (isPdf && typeof extractTextFromPdfViaDrive === "function") {
    const ocr = extractTextFromPdfViaDrive(file);
    if (ocr && ocr.text) textContent = ocr.text;
  }

  if (textContent && textContent.trim().length >= 30) {
    return {
      contents: [{ parts: [{ text: basePrompt + "\n\nINHALT:\n" + textContent }] }],
      generationConfig: genConfig
    };
  }

  if (isPdf || isImage) {
    let previewBlob = typeof getPageOnePreviewImage === "function" ? getPageOnePreviewImage(file) : null;
    if (!previewBlob && isImage) previewBlob = file.getBlob();
    if (previewBlob) {
      genConfig.mediaResolution = "MEDIA_RESOLUTION_LOW";
      return {
        contents: [{
          parts: [
            { text: basePrompt },
            { inline_data: { mime_type: previewBlob.getContentType(), data: Utilities.base64Encode(previewBlob.getBytes()) } }
          ]
        }],
        generationConfig: genConfig
      };
    }
  }

  return null;
}

function syncJournalToMonthlySheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const journal = getOrCreateBuchungsjournal(ss);
  if (journal.getLastRow() < 2) {
    ss.toast("Das Buchungsjournal enthält noch keine Einträge.", "Hinweis", 5);
    return;
  }

  for (let m = 0; m < 12; m++) {
    const mSheet = getMonthSheet(ss, m);
    if (mSheet && mSheet.getLastRow() >= 4) {
      const numRows = mSheet.getLastRow() - 3;
      mSheet.getRange(4, 1, numRows, 5).clearContent();
      mSheet.getRange(4, 7, numRows, 6).clearContent();
    }
  }

  const headers = journal.getRange(1, 1, 1, journal.getLastColumn()).getValues()[0];
  const dueIndex = headers.indexOf("Fälligkeitsdatum");
  const fileIdIndex = headers.indexOf("Datei-ID");

  const rows = journal.getRange(2, 1, journal.getLastRow() - 1, Math.max(12, journal.getLastColumn())).getValues();
  rows.forEach(r => {
    // Date-Objekte aus Google Sheets wieder in "yyyy-MM" umwandeln
    let monthStr = "";
    if (r[0] instanceof Date && !isNaN(r[0].getTime())) {
      monthStr = Utilities.formatDate(r[0], "Europe/Berlin", "yyyy-MM");
    } else {
      monthStr = (r[0] || "").toString().trim();
    }
    if (!monthStr && r[1]) {
      const parsedBDate = normalizeDateToIso(r[1]);
      if (parsedBDate) monthStr = parsedBDate.substring(0, 7);
    }

    if (!monthStr.startsWith(ACCOUNTING_YEAR + "-")) return;

    writeEntryToMonthlySheet(ss, {
      bDate: r[1],
      type: r[2],
      partner: r[3],
      invNr: r[4],
      amount: r[5],
      currency: r[6],
      category: r[7],
      payStatus: r[8],
      link: r[9],
      dueDate: dueIndex > -1 ? r[dueIndex] : "",
      fileId: fileIdIndex > -1 ? r[fileIdIndex] : ""
    }, monthStr);
  });

  ss.toast(`${rows.length} Buchung(en) aus dem Journal in die Monatsblätter übertragen.`, "Synchronisation fertig ✓", 6);
}

function resetAccountingScanLog() {
  PropertiesService.getScriptProperties().deleteProperty("ACCOUNTING_SCANNED_FILE_IDS");
  SpreadsheetApp.getActiveSpreadsheet().toast("Scan-Verlauf wurde zurückgesetzt.", "Zurückgesetzt ✓", 5);
}