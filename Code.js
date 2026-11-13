// ==========================================
// KONFIGURATION
// ==========================================
const COMPANY_NAME = "fb-Pflege / Fernanda Bogado";
const INBOX_FOLDER_ID = "1ytvHfJqTRFvrf78N0VNmpq-ltJuFQtGK";

// Gmail-Einstellungen
const GMAIL_LABEL_NAME = "KI-Ablage";
const AUTO_FETCH_GMAIL_ON_ANALYZE = true;
const SAVE_EMAIL_BODY_AS_PDF = true;
const MAX_GMAIL_THREADS_PER_RUN = 15;

// KI & Archiv-Einstellungen
const PRIMARY_MODEL = "auto";
const ARCHIVE_SHEET_NAME = "Archiv_Log";
const LIVE_LOG_SHEET_NAME = "Live_Log";
const MAX_LOG_ROWS = 250;
const MAX_RUNTIME_MS = 280000; // Sicherheitsstopp bei ~4 Min. 40 Sek.

// Token-Optimierungs-Limits
const MAX_EXCEL_ROWS = 35;
const MAX_EXTRACTED_CHARS = 2000;
const MAX_IMAGE_DIMENSION = 1024;
const MAX_MATCHED_FOLDERS = 60;

// Namensschema
const NAMING_CONVENTION = "JJJJ-MM-TT_Absender-oder-Thema_Dokumentart (z. B. 2026-11-12_Telekom_Rechnung)";

function getApiKey() {
  const key = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!key || key.trim() === "") {
    throw new Error("Kein API-Schlüssel hinterlegt. Bitte 'GEMINI_API_KEY' in den Projekteinstellungen unter Skripteigenschaften eintragen.");
  }
  return key.trim();
}


// ==========================================
// MENÜ & LIVE-PROTOKOLLIERUNG (TOAST + LIVE_LOG)
// ==========================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🤖 KI-Ablage")
    .addItem("0. Ordner & Unterordner einlesen", "syncSubfolders")
    .addItem("📥 Gmail-Anhänge in Inbox laden", "runGmailImportOnly")
    .addSeparator()
    .addItem("1. Posteingang prüfen & analysieren", "analyzeInboxFiles")
    .addItem("2. Genehmigte Dateien umbenennen & verschieben", "executeApprovedMoves")
    .addSeparator()
    .addItem("🧹 Live-Log leeren", "clearLiveLog")
    .addToUi();

  if (typeof createModelMenu === "function") createModelMenu();
  if (typeof createAccountingMenu === "function") createAccountingMenu();
  if (typeof setupAccountingSheets === "function") {
    setupAccountingSheets(SpreadsheetApp.getActiveSpreadsheet());
  }
}

function updateLiveStatus(ss, message, title = "🤖 KI-Ablage", seconds = 5, flushNow = true) {
  console.log(`[${title}] ${message}`);
  if (!ss) return;

  ss.toast(message, title, seconds);

  try {
    const logSheet = getOrCreateLiveLogSheet(ss);
    const timeStr = Utilities.formatDate(new Date(), "Europe/Berlin", "HH:mm:ss");
    logSheet.insertRowBefore(2);
    logSheet.getRange(2, 1, 1, 3).setValues([[timeStr, title, message]]);

    if (logSheet.getLastRow() > MAX_LOG_ROWS + 1) {
      logSheet.deleteRows(MAX_LOG_ROWS + 2, logSheet.getLastRow() - (MAX_LOG_ROWS + 1));
    }
    if (flushNow) SpreadsheetApp.flush();
  } catch (e) {}
}

function getOrCreateLiveLogSheet(ss) {
  let logSheet = ss.getSheetByName(LIVE_LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = ss.insertSheet(LIVE_LOG_SHEET_NAME);
    logSheet.getRange(1, 1, 1, 3)
      .setValues([["Uhrzeit", "Schritt / Phase", "Live-Status & Details"]])
      .setFontWeight("bold")
      .setBackground("#e8f0fe");
    logSheet.setFrozenRows(1);
    logSheet.setColumnWidth(1, 95);
    logSheet.setColumnWidth(2, 230);
    logSheet.setColumnWidth(3, 680);
  }
  return logSheet;
}

function clearLiveLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(LIVE_LOG_SHEET_NAME);
  if (logSheet && logSheet.getLastRow() > 1) {
    logSheet.deleteRows(2, logSheet.getLastRow() - 1);
  }
  ss.toast("Das Live-Log wurde geleert.", "Live-Log bereinigt ✓", 4);
}

function formatFileSize(bytes) {
  if (!bytes || bytes <= 0) return "0 KB";
  if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}


// ==========================================
// GMAIL-IMPORT: ANHÄNGE & E-MAILS IN 00_INBOX LADEN
// ==========================================
function runGmailImportOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const count = fetchGmailAttachments(ss);
  updateLiveStatus(ss, `Gmail-Import beendet: ${count} Datei(en) in '00_Inbox' gespeichert.`, "📥 Gmail-Import fertig ✓", 6);
}

function fetchGmailAttachments(ss) {
  updateLiveStatus(ss, `Prüfe Gmail-Label "${GMAIL_LABEL_NAME}" auf neue E-Mails...`, "📥 Gmail-Import", 4);

  let label = GmailApp.getUserLabelByName(GMAIL_LABEL_NAME);
  if (!label) {
    label = GmailApp.createLabel(GMAIL_LABEL_NAME);
    updateLiveStatus(ss, `Gmail-Label "${GMAIL_LABEL_NAME}" wurde neu angelegt (0 E-Mails).`, "📥 Gmail-Import", 4);
    return 0;
  }

  const threads = label.getThreads(0, MAX_GMAIL_THREADS_PER_RUN);
  if (threads.length === 0) {
    updateLiveStatus(ss, `Keine neuen E-Mails unter "${GMAIL_LABEL_NAME}" gefunden.`, "📥 Gmail-Import", 3, false);
    return 0;
  }

  updateLiveStatus(ss, `${threads.length} Gmail-Konversation(en) gefunden. Starte Import...`, "📥 Gmail-Import", 5);

  const inboxFolder = DriveApp.getFolderById(INBOX_FOLDER_ID);
  const allowedMimeTypes = new Set([
    "application/pdf", "image/jpeg", "image/png", "image/webp", "image/heic",
    "text/plain", "text/csv",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "application/msword",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/vnd.ms-excel"
  ]);

  let importedFilesCount = 0;

  for (let t = 0; t < threads.length; t++) {
    const thread = threads[t];
    const messages = thread.getMessages();
    const subject = messages.length > 0 ? (messages[0].getSubject() || "Ohne Betreff") : "Ohne Betreff";
    updateLiveStatus(ss, `E-Mail ${t + 1}/${threads.length}: "${subject.substring(0, 50)}"...`, "📥 Gmail-Import", 4);

    let savedInThisThread = 0;

    for (let m = 0; m < messages.length; m++) {
      const msg = messages[m];
      const attachments = msg.getAttachments({ includeInlineImages: false, includeAttachments: true });

      for (let a = 0; a < attachments.length; a++) {
        const att = attachments[a];
        const mime = (att.getContentType() || "").toLowerCase();
        const sizeBytes = att.getSize();

        if (!allowedMimeTypes.has(mime)) continue;
        if (mime.startsWith("image/") && sizeBytes < 5120) continue;

        let ext = "";
        if (mime === "application/vnd.openxmlformats-officedocument.wordprocessingml.document") ext = ".docx";
        else if (mime === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") ext = ".xlsx";

        const fileName = att.getName() || `Gmail-Anhang_${Utilities.formatDate(msg.getDate(), "Europe/Berlin", "yyyy-MM-dd")}${ext}`;
        inboxFolder.createFile(att.copyBlob().setName(fileName));
        importedFilesCount++;
        savedInThisThread++;
        updateLiveStatus(ss, `Anhang gespeichert: "${fileName}" (${formatFileSize(sizeBytes)})`, "📥 Gmail-Import", 4, false);
      }
    }

    if (savedInThisThread === 0 && SAVE_EMAIL_BODY_AS_PDF && messages.length > 0) {
      const pdfBlob = convertEmailToPdfBlob(messages[messages.length - 1]);
      if (pdfBlob) {
        inboxFolder.createFile(pdfBlob);
        importedFilesCount++;
        updateLiveStatus(ss, `E-Mail-Text als PDF gespeichert: "${pdfBlob.getName()}"`, "📥 Gmail-Import", 4, false);
      }
    }
    thread.removeLabel(label);
  }
  return importedFilesCount;
}

function convertEmailToPdfBlob(message) {
  const subject = message.getSubject() || "Ohne_Betreff";
  const sender = message.getFrom() || "Unbekannter_Absender";
  const dateStr = Utilities.formatDate(message.getDate(), "Europe/Berlin", "yyyy-MM-dd");
  const safeSubject = subject.replace(/[\/\\:*?"<>|]/g, "-").substring(0, 60).trim();
  const pdfFileName = `${dateStr}_Email_${safeSubject}.pdf`;

  const headerHtml = `
    <div style="font-family: Arial, sans-serif; font-size: 12px; color: #333; border-bottom: 1px solid #ccc; padding-bottom: 10px; margin-bottom: 15px;">
      <strong>Betreff:</strong> ${escapeHtml(subject)}<br>
      <strong>Von:</strong> ${escapeHtml(sender)}<br>
      <strong>Datum:</strong> ${Utilities.formatDate(message.getDate(), "Europe/Berlin", "dd.MM.yyyy HH:mm")}
    </div>
  `;

  const rawBody = message.getBody() || `<pre>${escapeHtml(message.getPlainBody() || "")}</pre>`;
  const fullHtml = `<html><head><meta charset="UTF-8"></head><body>${headerHtml}${rawBody}</body></html>`;

  try {
    return Utilities.newBlob(fullHtml, "text/html", pdfFileName).getAs("application/pdf");
  } catch (e) {
    const cleanedBody = rawBody.replace(/<img[^>]*>/gi, "");
    try {
      return Utilities.newBlob(`<html><head><meta charset="UTF-8"></head><body>${headerHtml}${cleanedBody}</body></html>`, "text/html", pdfFileName).getAs("application/pdf");
    } catch (e2) {
      return null;
    }
  }
}

function escapeHtml(text) {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}


// ==========================================
// 1. EINZEL-DATEI-ANALYSE FÜR 00_INBOX
// ==========================================
function analyzeInboxFiles() {
  const startTime = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Prüfung & Freigabe");
  const whitelistSheet = ss.getSheetByName("Ordner_Whitelist");

  if (!sheet || !whitelistSheet) {
    ss.toast("Tabellenblatt 'Prüfung & Freigabe' oder 'Ordner_Whitelist' fehlt.", "Konfigurationsfehler ⚠️", 8);
    return;
  }

  updateLiveStatus(ss, "Starte Posteingangs-Prüfung & prüfe Tabellenstruktur...", "🚀 Start", 4);

  if (typeof setupAccountingSheets === "function") setupAccountingSheets(ss);

  try {
    getApiKey();
  } catch (keyErr) {
    updateLiveStatus(ss, keyErr.message, "Konfigurationsfehler ⚠️", 10);
    return;
  }

  let gmailImportedCount = 0;
  if (AUTO_FETCH_GMAIL_ON_ANALYZE) {
    gmailImportedCount = fetchGmailAttachments(ss);
  }

  updateLiveStatus(ss, "Lade Zielordner aus 'Ordner_Whitelist' & Duplikat-Datenbank...", "📋 Vorbereitung", 4);

  const { allowedFolders, folderMap } = loadWhitelistData(whitelistSheet);
  if (allowedFolders.length === 0) {
    updateLiveStatus(ss, "Keine Zielordner in 'Ordner_Whitelist' gefunden. Bitte zuerst Ordner einlesen.", "Hinweis ⚠️", 6);
    return;
  }

  const { existingRowsMap, knownHashes, knownNames } = loadKnownDuplicates(ss, sheet);

  const inbox = DriveApp.getFolderById(INBOX_FOLDER_ID);
  const files = inbox.getFiles();
  const filesToProcess = [];
  let alreadyListedCount = 0;

  while (files.hasNext()) {
    const file = files.next();
    const existingEntry = existingRowsMap.get(file.getId());
    if (existingEntry && !existingEntry.isError) {
      alreadyListedCount++;
      continue;
    }
    filesToProcess.push({ file, existingEntry });
  }

  if (filesToProcess.length === 0) {
    updateLiveStatus(
      ss,
      `Keine neuen Dateien in '00_Inbox' (${alreadyListedCount} Datei(en) warten bereits in 'Prüfung & Freigabe').`,
      "Alles aktuell ✓",
      6
    );
    return;
  }

  updateLiveStatus(
    ss,
    `${filesToProcess.length} neue Datei(en) in '00_Inbox' gefunden (${allowedFolders.length} Zielordner, ${knownHashes.size} bekannte Hashes geladen).`,
    "📂 Posteingang bereit",
    5
  );

  let totalTokensRun = 0;
  let duplicateCount = 0;
  let processedCount = 0;
  let errorCount = 0;
  let stoppedByTimeout = false;

  for (let i = 0; i < filesToProcess.length; i++) {
    if (Date.now() - startTime > MAX_RUNTIME_MS) {
      stoppedByTimeout = true;
      updateLiveStatus(
        ss,
        `Zeitlimit (4 Min. 40 Sek.) erreicht. ${filesToProcess.length - processedCount} Datei(en) verbleiben für den nächsten Durchlauf.`,
        "⏱️ Sicherheitspause",
        8
      );
      break;
    }

    const { file, existingEntry } = filesToProcess[i];
    const fileId = file.getId();
    const fileName = file.getName();
    const sizeLabel = formatFileSize(file.getSize());
    const stepPrefix = `📄 Datei ${i + 1}/${filesToProcess.length}`;

    updateLiveStatus(ss, `[1/4] Prüfe Fingerabdruck (MD5) für "${fileName}" (${sizeLabel})...`, stepPrefix, 4);

    const fileHash = computeFileHash(file);
    let isExactDup = false;
    let aiResult;

    if (fileHash && knownHashes.has(fileHash)) {
      const dupMatch = knownHashes.get(fileHash);
      isExactDup = true;
      updateLiveStatus(
        ss,
        `[2/4] Exaktes Duplikat erkannt! Identisch mit "${dupMatch.name}" (${dupMatch.source}) – 0 Tokens verbraucht.`,
        stepPrefix,
        5
      );
      aiResult = {
        folder_name: dupMatch.folder || "MANUELLE_PRUEFUNG",
        suggested_filename: dupMatch.name || fileName,
        reasoning: `⚠️ EXAKTES DUPLIKAT (${dupMatch.source}): Identischer Inhalt existiert als "${dupMatch.name}".`,
        tokens: 0,
        tokenDisplay: "0"
      };
    } else {
      updateLiveStatus(ss, `[2/4] Lese Dateiinhalt aus "${fileName}" aus (OCR / Text / Vorschau)...`, stepPrefix, 5);
      const prep = prepareClassificationPayload(file, allowedFolders);

      if (prep && prep.payload) {
        updateLiveStatus(ss, `[3/4] Extraktion fertig (${prep.modeTag}). Sende Anfrage an Gemini KI...`, stepPrefix, 8);
        aiResult = executeGeminiCall(prep.payload, allowedFolders, ss, stepPrefix, fileName);
        aiResult.reasoning = `${aiResult.reasoning} [${prep.modeTag}]`;
      } else {
        aiResult = {
          folder_name: "Fehler bei Analyse",
          suggested_filename: fileName,
          reasoning: "Dokument konnte nicht vorbereitet werden",
          tokens: 0,
          tokenDisplay: "0"
        };
      }
    }

    totalTokensRun += (aiResult.tokens || 0);
    const hasApiError = aiResult.folder_name.startsWith("API-Fehler") || aiResult.folder_name === "Fehler bei Analyse";
    if (hasApiError) errorCount++;

    let statusText = isExactDup ? "⚠️ Duplikat" : "Offen";
    let isDuplicate = isExactDup;
    if (isDuplicate) duplicateCount++;

    const cleanSuggestedName = ensureFileExtension(fileName, aiResult.suggested_filename || fileName);

    if (!isDuplicate && !hasApiError) {
      const lowerSuggested = cleanSuggestedName.toLowerCase();
      if (knownNames.has(lowerSuggested)) {
        const nameMatch = knownNames.get(lowerSuggested);
        isDuplicate = true;
        duplicateCount++;
        statusText = "⚠️ Duplikat";
        aiResult.reasoning = `⚠️ NAMENS-DUPLIKAT (${nameMatch.source}): ${aiResult.reasoning}`;
      } else if (folderMap[aiResult.folder_name]) {
        try {
          const targetFolder = DriveApp.getFolderById(folderMap[aiResult.folder_name]);
          if (targetFolder.getFilesByName(cleanSuggestedName).hasNext()) {
            isDuplicate = true;
            duplicateCount++;
            statusText = "⚠️ Duplikat";
            aiResult.reasoning = `⚠️ IM ZIELORDNER VORHANDEN: ${aiResult.reasoning}`;
          }
        } catch (driveErr) {}
      }
    }

    if (fileHash && !knownHashes.has(fileHash)) {
      knownHashes.set(fileHash, { name: cleanSuggestedName, folder: aiResult.folder_name, source: "Aktueller Lauf" });
    }
    if (cleanSuggestedName && !knownNames.has(cleanSuggestedName.toLowerCase())) {
      knownNames.set(cleanSuggestedName.toLowerCase(), { folder: aiResult.folder_name, source: "Aktueller Lauf" });
    }

    const rowValues = [
      `=HYPERLINK("${file.getUrl()}"; "${fileName.replace(/"/g, '""')}")`,
      file.getDateCreated(),
      aiResult.folder_name,
      cleanSuggestedName,
      aiResult.reasoning,
      false,
      "",
      statusText,
      fileId,
      aiResult.tokenDisplay || aiResult.tokens || 0,
      fileHash,
      aiResult.t || "",
      aiResult.d || "",
      aiResult.p || "",
      aiResult.b || "",
      aiResult.v || 0,
      aiResult.c || "",
      aiResult.k || "",
      aiResult.s || "",
      aiResult.due || ""
    ];

    const targetRow = (existingEntry && existingEntry.isError)
      ? existingEntry.rowNumber
      : sheet.getLastRow() + 1;

    sheet.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
    sheet.getRange(targetRow, 6).insertCheckboxes();
    sheet.getRange(targetRow, 8).setBackground(isDuplicate ? "#fff2cc" : (hasApiError ? "#f4cccc" : null));

    const outcomeInfo = hasApiError
      ? `Fehler (${aiResult.folder_name}): ${aiResult.reasoning}`
      : `Ziel: "${aiResult.folder_name}" | Name: "${cleanSuggestedName}" | Tokens: ${aiResult.tokenDisplay}`;

    updateLiveStatus(ss, `[4/4] In Zeile ${targetRow} eingetragen → ${outcomeInfo}`, stepPrefix, 5);
    processedCount++;

    if (!isExactDup && i + 1 < filesToProcess.length) {
      Utilities.sleep(300);
    }
  }

  if (typeof applyFolderDropdownValidation === "function") {
    applyFolderDropdownValidation(ss);
  }

  const elapsedSec = ((Date.now() - startTime) / 1000).toFixed(1);
  const gmailInfo = gmailImportedCount > 0 ? ` | 📥 ${gmailImportedCount} aus Gmail` : "";
  const dupInfo = duplicateCount > 0 ? ` | ⚠️ ${duplicateCount} Duplikat(e)` : "";
  const errInfo = errorCount > 0 ? ` | ❌ ${errorCount} Fehler` : "";
  const remainingInfo = stoppedByTimeout ? ` | ⏱️ ${filesToProcess.length - processedCount} verbleibend` : "";

  updateLiveStatus(
    ss,
    `Fertig in ${elapsedSec}s: ${processedCount} Datei(en) geprüft, ${totalTokensRun} Tokens gesamt${gmailInfo}${dupInfo}${errInfo}${remainingInfo}`,
    "✅ Analyse abgeschlossen",
    10
  );
}

function loadWhitelistData(whitelistSheet) {
  const allowedFolders = [];
  const folderMap = {};
  const lastRow = whitelistSheet.getLastRow();
  if (lastRow === 0) return { allowedFolders, folderMap };

  const rows = whitelistSheet.getRange(1, 1, lastRow, 2).getValues();
  rows.forEach(row => {
    const name = row[0] ? row[0].toString().trim() : "";
    const id = row[1] ? row[1].toString().trim() : "";
    if (name && id && id.length > 10) {
      allowedFolders.push(name);
      folderMap[name] = id;
    }
  });
  return { allowedFolders, folderMap };
}

function loadKnownDuplicates(ss, reviewSheet) {
  const existingRowsMap = new Map();
  const knownHashes = new Map();
  const knownNames = new Map();

  const lastRow = reviewSheet.getLastRow();
  if (lastRow > 1) {
    const maxCol = Math.max(reviewSheet.getLastColumn(), 11);
    const existingData = reviewSheet.getRange(2, 1, lastRow - 1, maxCol).getValues();
    existingData.forEach((row, index) => {
      const folderSuggestion = row[2] ? row[2].toString().trim() : "";
      const suggestedName = row[3] ? row[3].toString().trim() : "";
      const existingId = row[8] ? row[8].toString().trim() : "";
      const existingHash = row[10] ? row[10].toString().trim() : "";

      if (existingId) {
        const isError = folderSuggestion.startsWith("API-Fehler") || folderSuggestion === "Fehler bei Analyse";
        existingRowsMap.set(existingId, { rowNumber: index + 2, isError });

        if (!isError) {
          if (existingHash) knownHashes.set(existingHash, { name: suggestedName, folder: folderSuggestion, source: "Prüfliste" });
          if (suggestedName) knownNames.set(suggestedName.toLowerCase(), { folder: folderSuggestion, source: "Prüfliste" });
        }
      }
    });
  }

  const archiveSheet = ss.getSheetByName(ARCHIVE_SHEET_NAME);
  if (archiveSheet && archiveSheet.getLastRow() > 1) {
    const archCols = archiveSheet.getLastColumn();
    const archData = archiveSheet.getRange(2, 1, archiveSheet.getLastRow() - 1, archCols).getValues();
    archData.forEach(row => {
      const archNewName = row[0] ? row[0].toString().trim() : "";
      const archFolder = row[2] ? row[2].toString().trim() : "";
      const archHash = (archCols >= 9 && row[8]) ? row[8].toString().trim() : "";

      if (archHash && !knownHashes.has(archHash)) {
        knownHashes.set(archHash, { name: archNewName, folder: archFolder, source: "Archiv" });
      }
      if (archNewName && !knownNames.has(archNewName.toLowerCase())) {
        knownNames.set(archNewName.toLowerCase(), { folder: archFolder, source: "Archiv" });
      }
    });
  }

  return { existingRowsMap, knownHashes, knownNames };
}


// ==========================================
// 2. UMBENENNEN, VERSCHIEBEN & ARCHIVIEREN
// ==========================================
function executeApprovedMoves() {
  const startTime = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Prüfung & Freigabe");
  const whitelistSheet = ss.getSheetByName("Ordner_Whitelist");

  if (!sheet || !whitelistSheet) {
    ss.toast("Tabellenblatt 'Prüfung & Freigabe' oder 'Ordner_Whitelist' fehlt.", "Konfigurationsfehler ⚠️", 8);
    return;
  }

  const archiveSheet = getOrCreateArchiveSheet(ss);
  const { folderMap } = loadWhitelistData(whitelistSheet);

  const data = sheet.getDataRange().getValues();
  let totalToMove = 0;
  for (let i = 1; i < data.length; i++) {
    const status = (data[i][7] || "").toString().trim();
    if (data[i][5] === true && (status === "Offen" || status === "⚠️ Duplikat")) totalToMove++;
  }

  if (totalToMove === 0) {
    ss.toast("Keine freigegebenen Dateien zum Verschieben markiert.", "Hinweis", 5);
    return;
  }

  updateLiveStatus(ss, `Starte Verschiebung von ${totalToMove} freigegebenen Datei(en)...`, "📂 Verschieben", 5);

  let movedCount = 0;
  let currentStep = 0;
  const rowsToDelete = [];
  const archiveRowsToAdd = [];

  for (let i = 1; i < data.length; i++) {
    const originalName = data[i][0];
    const dateCreated = data[i][1];
    const aiFolder = (data[i][2] || "").toString().trim();
    const newNameInput = (data[i][3] || "").toString().trim();
    const reasoning = data[i][4];
    const isApproved = data[i][5];
    const manualOverride = (data[i][6] || "").toString().trim();
    const status = (data[i][7] || "").toString().trim();
    const fileId = (data[i][8] || "").toString().trim();
    const tokens = data[i][9] || 0;
    const fileHash = data[i][10] || "";

    const bType = data[i][11] || "";
    const bDate = data[i][12] || "";
    const bPartner = data[i][13] || "";
    const bInvNr = data[i][14] || "";
    const bAmount = data[i][15] || 0;
    const bCurrency = data[i][16] || "";
    const bCategory = data[i][17] || "";
    const bStatus = data[i][18] || "";
    const bDueDate = data[i][19] || "";

    if (isApproved === true && (status === "Offen" || status === "⚠️ Duplikat")) {
      currentStep++;
      const targetFolderName = manualOverride !== "" ? manualOverride : aiFolder;
      const targetFolderId = folderMap[targetFolderName];

      if (targetFolderId) {
        try {
          const file = DriveApp.getFileById(fileId);
          const oldFileName = file.getName();
          const targetFolder = DriveApp.getFolderById(targetFolderId);

          const finalName = newNameInput !== ""
            ? ensureFileExtension(oldFileName, newNameInput)
            : oldFileName;

          updateLiveStatus(
            ss,
            `(${currentStep}/${totalToMove}) "${oldFileName}" → "${finalName}" nach "${targetFolderName}"...`,
            "📂 Verschieben",
            5
          );

          if (finalName !== oldFileName) file.setName(finalName);
          file.moveTo(targetFolder);

          const finalLink = `=HYPERLINK("${file.getUrl()}"; "${finalName.replace(/"/g, '""')}")`;
          archiveRowsToAdd.push([
            finalLink, originalName, targetFolderName, reasoning, dateCreated, new Date(), fileId, tokens, fileHash
          ]);

          const isFinanceFolder = targetFolderName.replace(/\s*\/\s*/g, " / ").startsWith("01_Finanzen");
          if (isFinanceFolder && typeof bookApprovedInvoice === "function" && bType !== "Kein Beleg" && bType !== "") {
            updateLiveStatus(ss, `Buche Finanzbeleg "${finalName}" ins Buchungsjournal...`, "📊 Buchhaltung", 4, false);
            bookApprovedInvoice(ss, {
              link: finalLink, type: bType, bDate: bDate, dueDate: bDueDate, partner: bPartner,
              invNr: bInvNr, amount: bAmount, currency: bCurrency,
              category: bCategory, payStatus: bStatus, fileId: fileId, tokens: tokens
            });
          }

          rowsToDelete.push(i + 1);
          movedCount++;
        } catch (e) {
          sheet.getRange(i + 1, 8).setValue("Fehler beim Verschieben").setBackground("#f4cccc");
          updateLiveStatus(ss, `Fehler bei Zeile ${i + 1}: ${e.message}`, "📂 Fehler ⚠️", 6);
        }
      } else {
        sheet.getRange(i + 1, 8).setValue("Fehler: Ordner nicht in Whitelist").setBackground("#f4cccc");
        updateLiveStatus(ss, `Fehler bei Zeile ${i + 1}: Zielordner "${targetFolderName}" nicht in Whitelist.`, "📂 Fehler ⚠️", 6);
      }
    }
  }

  if (archiveRowsToAdd.length > 0) {
    const startArchRow = archiveSheet.getLastRow() + 1;
    archiveSheet.getRange(startArchRow, 1, archiveRowsToAdd.length, archiveRowsToAdd[0].length).setValues(archiveRowsToAdd);
  }

  if (rowsToDelete.length > 0) {
    for (let r = rowsToDelete.length - 1; r >= 0; r--) sheet.deleteRow(rowsToDelete[r]);
  }

  const elapsedSec = ((Date.now() - startTime) / 1000).toFixed(1);
  updateLiveStatus(
    ss,
    `${movedCount} von ${totalToMove} Datei(en) in ${elapsedSec}s erfolgreich umbenannt, verschoben und archiviert ✓`,
    "✅ Verschieben fertig",
    8
  );
}

function getOrCreateArchiveSheet(ss) {
  let archiveSheet = ss.getSheetByName(ARCHIVE_SHEET_NAME);
  const headers = ["Dateiname (Neu)", "Original-Dateiname", "Zielordner", "KI-Begründung", "Eingangsdatum", "Verschoben am", "Datei-ID", "Token-Verbrauch", "Datei-Hash"];

  if (!archiveSheet) {
    archiveSheet = ss.insertSheet(ARCHIVE_SHEET_NAME);
    archiveSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#f3f3f3");
    archiveSheet.setFrozenRows(1);
  } else {
    const lastArchCol = archiveSheet.getLastColumn();
    if (lastArchCol === 0) {
      archiveSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#f3f3f3");
      archiveSheet.setFrozenRows(1);
    } else {
      const currentHeaders = archiveSheet.getRange(1, 1, 1, lastArchCol).getValues()[0];
      headers.forEach(headerTitle => {
        if (!currentHeaders.includes(headerTitle)) {
          const nextCol = archiveSheet.getLastColumn() + 1;
          archiveSheet.getRange(1, nextCol).setValue(headerTitle).setFontWeight("bold").setBackground("#f3f3f3");
        }
      });
    }
  }
  return archiveSheet;
}


// ==========================================
// DOKUMENT-EXTRAKTION & DRIVE-HILFSFUNKTIONEN
// ==========================================
function computeFileHash(file) {
  try {
    if (file.getSize() / (1024 * 1024) > 25) return "";
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, file.getBlob().getBytes());
    return digest.map(b => ("0" + ((b < 0 ? b + 256 : b).toString(16))).slice(-2)).join("");
  } catch (e) {
    return "";
  }
}

function prepareClassificationPayload(file, allowedFolders) {
  const lowerName = file.getName().toLowerCase();
  const driveMime = (file.getMimeType() || "").toLowerCase();
  const sizeInMB = file.getSize() / (1024 * 1024);

  const isText = driveMime === "text/plain" || driveMime === "text/csv" ||
                 lowerName.endsWith(".txt") || lowerName.endsWith(".csv");
  const isWord = driveMime.includes("wordprocessingml") || driveMime === "application/msword" ||
                 lowerName.endsWith(".docx") || lowerName.endsWith(".doc");
  const isExcel = driveMime.includes("spreadsheetml") || driveMime === "application/vnd.ms-excel" ||
                  lowerName.endsWith(".xlsx") || lowerName.endsWith(".xls");
  const isPdf = driveMime === "application/pdf" || lowerName.endsWith(".pdf");
  const isImage = driveMime.startsWith("image/") || /\.(jpg|jpeg|png|webp|heic)$/i.test(lowerName);

  if (isText) {
    const rawText = file.getBlob().getDataAsString("UTF-8").substring(0, MAX_EXTRACTED_CHARS);
    if (rawText.trim().length > 10) {
      return buildTextPayload(file, rawText, allowedFolders, `Text (${rawText.trim().length} Zeichen)`);
    }
  }

  if (isWord || isExcel) {
    try {
      const extractedText = extractTextFromOfficeFile(file, isWord);
      if (extractedText && extractedText.trim().length > 15) {
        return buildTextPayload(file, extractedText, allowedFolders, `Office-Text (${extractedText.length} Zeichen)`);
      }
    } catch (officeErr) {}
    return buildMetadataPayload(file, allowedFolders, "Office-Fehler");
  }

  if (isPdf) {
    let ocrErrorNote = "";
    if (sizeInMB <= 15) {
      const ocrResult = extractTextFromPdfViaDrive(file);
      if (ocrResult.text && ocrResult.text.length >= 40) {
        return buildTextPayload(file, ocrResult.text, allowedFolders, `PDF-Text (${ocrResult.text.length} Zeichen)`);
      }
      ocrErrorNote = ocrResult.error;
    }

    const previewBlob = getPageOnePreviewImage(file);
    if (previewBlob) {
      const hint = ocrErrorNote ? `Seite-1-Bild (${ocrErrorNote})` : "Seite-1-Bild";
      return buildInlinePayload(file, allowedFolders, hint, previewBlob, true);
    }

    return buildMetadataPayload(file, allowedFolders, ocrErrorNote || "PDF ohne Vorschau");
  }

  if (isImage) {
    const resizedBlob = (file.getSize() > 200 * 1024) ? getPageOnePreviewImage(file) : null;
    return buildInlinePayload(file, allowedFolders, resizedBlob ? "Bild (komprimiert)" : "Bild", resizedBlob, true);
  }

  return buildMetadataPayload(file, allowedFolders, "Unbekanntes Format");
}

function buildTextPayload(file, extractedText, allowedFolders, modeTag) {
  const filteredFolders = selectRelevantFolders(allowedFolders, file.getName(), extractedText);
  return {
    modeTag: `${modeTag} | ${filteredFolders.length}/${allowedFolders.length} Ordner`,
    payload: {
      contents: [{ parts: [{ text: buildPrompt(file, filteredFolders) + "\n\nINHALT:\n" + extractedText }] }],
      generationConfig: getJsonConfig(false)
    }
  };
}

function buildInlinePayload(file, allowedFolders, modeTag, customBlob, useLowMediaRes) {
  const filteredFolders = selectRelevantFolders(allowedFolders, file.getName(), "");
  const fileBlob = customBlob || file.getBlob();
  return {
    modeTag: `${modeTag} | ${filteredFolders.length}/${allowedFolders.length} Ordner`,
    payload: {
      contents: [{
        parts: [
          { text: buildPrompt(file, filteredFolders) },
          { inline_data: { mime_type: fileBlob.getContentType(), data: Utilities.base64Encode(fileBlob.getBytes()) } }
        ]
      }],
      generationConfig: getJsonConfig(useLowMediaRes)
    }
  };
}

function buildMetadataPayload(file, allowedFolders, reason) {
  const filteredFolders = selectRelevantFolders(allowedFolders, file.getName(), "");
  const fallbackDate = Utilities.formatDate(file.getDateCreated(), "Europe/Berlin", "yyyy-MM-dd");
  const compactList = formatIndexedFolders(filteredFolders);
  const prompt = `Archivierung nur per Dateiname ("${reason}"). Datei: "${file.getName()}" (Datum: ${fallbackDate}).
Ordner-Index:
${compactList}
Regeln:
- f: Passende Ordner-Nummer (Integer) aus Liste, oder -1 falls unklar.
- n: Dateiname nach Schema "${NAMING_CONVENTION}".
- r: Deutsche Begründung (max. 7 Wörter).`;

  return {
    modeTag: `Nur-Dateiname: ${reason} | ${filteredFolders.length}/${allowedFolders.length} Ordner`,
    payload: { contents: [{ parts: [{ text: prompt }] }], generationConfig: getJsonConfig(false) }
  };
}

function createTempConvertedDriveFile(blob, tempName, targetMimeType, useOcr = false) {
  if (typeof Drive === "undefined" || !Drive.Files) return null;
  if (typeof Drive.Files.create === "function") {
    return Drive.Files.create({ name: tempName, mimeType: targetMimeType }, blob).id;
  }
  if (typeof Drive.Files.insert === "function") {
    const options = useOcr ? { ocr: true, ocrLanguage: "de" } : {};
    return Drive.Files.insert({ title: tempName, mimeType: targetMimeType }, blob, options).id;
  }
  return null;
}

function deleteTempDriveFile(tempFileId) {
  if (!tempFileId) return;
  try {
    if (typeof Drive !== "undefined" && Drive.Files) {
      if (typeof Drive.Files.delete === "function") {
        Drive.Files.delete(tempFileId);
        return;
      }
      if (typeof Drive.Files.remove === "function") {
        Drive.Files.remove(tempFileId);
        return;
      }
    }
    DriveApp.getFileById(tempFileId).setTrashed(true);
  } catch (e) {}
}

function exportGoogleDocAsPlainText(fileId) {
  const exportUrl = `https://www.googleapis.com/drive/v3/files/${fileId}/export?mimeType=text/plain`;
  const res = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() === 200) return res.getContentText("UTF-8");
  return DocumentApp.openById(fileId).getBody().getText() || "";
}

function extractTextFromPdfViaDrive(file) {
  if (typeof Drive === "undefined" || !Drive.Files) {
    return { text: "", error: "Drive API nicht aktiv" };
  }

  let tempFileId = null;
  try {
    const pdfBlob = file.getBlob().setContentType("application/pdf");
    tempFileId = createTempConvertedDriveFile(
      pdfBlob,
      "Temp_PDF_OCR_" + file.getName(),
      "application/vnd.google-apps.document",
      true
    );
    if (!tempFileId) return { text: "", error: "OCR-Konvertierung fehlgeschlagen" };

    const rawText = exportGoogleDocAsPlainText(tempFileId);
    const cleanText = rawText.replace(/\s+/g, " ").trim().substring(0, MAX_EXTRACTED_CHARS);
    return { text: cleanText, error: cleanText.length < 40 ? "Kaum Text im PDF" : "" };
  } catch (e) {
    return { text: "", error: "OCR-Fehler" };
  } finally {
    deleteTempDriveFile(tempFileId);
  }
}

function extractTextFromOfficeFile(file, isWord) {
  if (typeof Drive === "undefined" || !Drive.Files) throw new Error("Drive-Dienst nicht aktiviert");

  const targetMimeType = isWord ? "application/vnd.google-apps.document" : "application/vnd.google-apps.spreadsheet";
  let tempFileId = null;

  try {
    tempFileId = createTempConvertedDriveFile(file.getBlob(), "Temp_Convert_" + file.getName(), targetMimeType, false);
    if (!tempFileId) return "";

    let extractedText = "";
    if (isWord) {
      extractedText = exportGoogleDocAsPlainText(tempFileId);
    } else {
      const sheet = SpreadsheetApp.openById(tempFileId);
      const values = sheet.getDataRange().getDisplayValues();
      const maxRows = Math.min(values.length, MAX_EXCEL_ROWS);
      const lines = [];
      for (let i = 0; i < maxRows; i++) lines.push(values[i].join(" | "));
      extractedText = lines.join("\n");
    }
    return extractedText.replace(/\s+/g, " ").trim().substring(0, MAX_EXTRACTED_CHARS);
  } finally {
    deleteTempDriveFile(tempFileId);
  }
}

function getPageOnePreviewImage(file) {
  try {
    const nativeThumb = file.getThumbnail();
    if (nativeThumb && nativeThumb.getBytes().length > 1000) return nativeThumb;
  } catch (e) {}

  if (typeof Drive !== "undefined" && Drive.Files) {
    try {
      let meta;
      try { meta = Drive.Files.get(file.getId(), { fields: "thumbnailLink" }); }
      catch (v2Err) { meta = Drive.Files.get(file.getId()); }

      if (meta && meta.thumbnailLink) {
        const resizedUrl = meta.thumbnailLink.replace(/=s\d+.*$/, `=s${MAX_IMAGE_DIMENSION}`);
        let res = UrlFetchApp.fetch(resizedUrl, { muteHttpExceptions: true });
        if (res.getResponseCode() !== 200) {
          res = UrlFetchApp.fetch(resizedUrl, {
            headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
            muteHttpExceptions: true
          });
        }
        if (res.getResponseCode() === 200) {
          const thumbBlob = res.getBlob();
          if (thumbBlob.getContentType().startsWith("image/") && thumbBlob.getBytes().length > 1000) return thumbBlob;
        }
      }
    } catch (e2) {}
  }
  return null;
}


// ==========================================
// INTELLIGENTE ORDNER-VORFILTERUNG
// ==========================================
function selectRelevantFolders(allowedFolders, fileName, extractedText) {
  if (allowedFolders.length <= 45) {
    return allowedFolders.map((path, idx) => ({ idx, path }));
  }

  const stopWords = new Set([
    "der", "die", "das", "und", "für", "von", "mit", "ein", "eine", "einer", "eines", "einem", "einen",
    "dem", "den", "des", "aus", "bei", "nach", "zur", "zum", "als", "auf", "oder", "nicht", "wird",
    "werden", "sind", "hat", "haben", "sich", "auch", "noch", "nur", "über", "unter", "durch", "gegen",
    "ohne", "bis", "seit", "vor", "zwischen", "pdf", "jpg", "png", "docx", "xlsx", "dokument", "datei"
  ]);

  const sourceText = `${fileName} ${extractedText || ""}`.toLowerCase();
  const rawTokens = sourceText.match(/[a-zäöüß0-9]{3,}/g) || [];

  const exactTokens = new Set();
  const stemTokens = new Set();

  rawTokens.forEach(tok => {
    if (stopWords.has(tok)) return;
    exactTokens.add(tok);
    if (tok.length >= 5) stemTokens.add(tok.substring(0, 5));
  });

  const scored = allowedFolders.map((path, idx) => {
    const cleanPath = path.replace(/\s*\/\s*/g, "/");
    const depth = cleanPath.split("/").length;
    const pathWords = cleanPath.toLowerCase().match(/[a-zäöüß0-9]{3,}/g) || [];

    let score = 0;
    pathWords.forEach(pw => {
      if (stopWords.has(pw)) return;
      if (exactTokens.has(pw)) score += 3;
      else if (pw.length >= 5 && stemTokens.has(pw.substring(0, 5))) score += 2;
    });

    return { idx, path, depth, score };
  });

  const matched = scored
    .filter(item => item.score > 0)
    .sort((a, b) => b.score - a.score || a.depth - b.depth)
    .slice(0, MAX_MATCHED_FOLDERS);

  const matchedIdxSet = new Set(matched.map(m => m.idx));
  const topLevel = scored
    .filter(item => item.depth <= 2 && !matchedIdxSet.has(item.idx))
    .slice(0, 25);

  const combined = [...matched, ...topLevel];
  combined.sort((a, b) => a.idx - b.idx);
  return combined.map(item => ({ idx: item.idx, path: item.path }));
}


// ==========================================
// GEMINI API-KOMMUNIKATION & ANTWORT-PARSER
// ==========================================
function parseGeminiJsonResponse(responseText, allowedFolders) {
  try {
    const json = JSON.parse(responseText);
    const rawText = json.candidates && json.candidates[0] && json.candidates[0].content &&
                    json.candidates[0].content.parts && json.candidates[0].content.parts[0] &&
                    json.candidates[0].content.parts[0].text;
    if (!rawText) {
      return {
        folder_name: "Fehler bei Analyse",
        suggested_filename: "",
        reasoning: "Leere oder blockierte KI-Antwort",
        tokens: 0,
        tokenDisplay: "0"
      };
    }

    const rawParsed = JSON.parse(rawText);
    const idx = Number(rawParsed.f);
    const resolvedFolder = (!isNaN(idx) && idx >= 0 && idx < allowedFolders.length)
      ? allowedFolders[idx]
      : "MANUELLE_PRUEFUNG";

    const usage = json.usageMetadata || {};
    const totalTok = usage.totalTokenCount || 0;
    const promptTok = usage.promptTokenCount || 0;
    const outTok = usage.candidatesTokenCount || 0;
    const thoughtTok = usage.thoughtsTokenCount || 0;

    const detailStr = thoughtTok > 0
      ? `${totalTok} (In:${promptTok} Out:${outTok} Denk:${thoughtTok})`
      : `${totalTok} (In:${promptTok} Out:${outTok})`;

    return {
      folder_name: resolvedFolder,
      suggested_filename: rawParsed.n || "",
      reasoning: rawParsed.r || "",
      tokens: totalTok,
      tokenDisplay: detailStr,
      t: rawParsed.t,
      d: rawParsed.d,
      due: rawParsed.due,
      p: rawParsed.p,
      b: rawParsed.b,
      v: rawParsed.v,
      c: rawParsed.c,
      k: rawParsed.k,
      s: rawParsed.s
    };
  } catch (e) {
    return {
      folder_name: "Fehler bei Analyse",
      suggested_filename: "",
      reasoning: "Ungültiges JSON von KI erhalten",
      tokens: 0,
      tokenDisplay: "0"
    };
  }
}

function formatIndexedFolders(indexedFolderItems) {
  return indexedFolderItems
    .map(item => `${item.idx}:${item.path.replace(/\s*\/\s*/g, "/")}`)
    .join("\n");
}

function buildPrompt(file, indexedFolderItems) {
  const fallbackDate = Utilities.formatDate(file.getDateCreated(), "Europe/Berlin", "yyyy-MM-dd");
  const compactList = formatIndexedFolders(indexedFolderItems);

  let prompt = `Archiviere Dokument (Original: "${file.getName()}", Fallback-Datum: ${fallbackDate}).
Ordner-Index:
${compactList}
Regeln:
- f: Passende Ordner-Nummer (Integer, spezifischsten Unterordner wählen, oder -1 falls unklar).
- n: Dateiname nach "${NAMING_CONVENTION}" mit Dokumentendatum (sonst ${fallbackDate}). Ohne / : * ? " < > |.
- r: Deutsche Begründung (max. 7 Wörter).`;

  if (typeof getAccountingPromptExtension === "function") {
    prompt += "\n" + getAccountingPromptExtension(COMPANY_NAME);
  }
  return prompt;
}

function getJsonConfig(useLowMediaRes = false) {
  const props = {
    f: { type: "INTEGER" },
    n: { type: "STRING" },
    r: { type: "STRING" }
  };
  const requiredFields = ["f", "n", "r"];

  if (typeof getAccountingSchemaExtension === "function") {
    const accProps = getAccountingSchemaExtension();
    Object.assign(props, accProps);
    requiredFields.push(...Object.keys(accProps));
  }

  const config = {
    response_mime_type: "application/json",
    response_schema: {
      type: "OBJECT",
      properties: props,
      required: requiredFields
    }
  };
  if (useLowMediaRes) config.mediaResolution = "MEDIA_RESOLUTION_LOW";
  return config;
}

function getActiveGeminiModels(ss, stepInfo) {
  if (PRIMARY_MODEL !== "auto") return [PRIMARY_MODEL.startsWith("models/") ? PRIMARY_MODEL : "models/" + PRIMARY_MODEL];

  const props = PropertiesService.getScriptProperties();
  const cachedJson = props.getProperty("CACHED_GEMINI_MODELS");
  if (cachedJson) {
    try {
      const parsed = JSON.parse(cachedJson);
      if (Array.isArray(parsed) && parsed.length > 0) return parsed;
    } catch (e) {}
  }

  updateLiveStatus(ss, "Rufe verfügbare Gemini-Modelle von Google API ab...", "🤖 Modell-Check", 4);

  const listUrl = "https://generativelanguage.googleapis.com/v1beta/models?key=" + getApiKey();
  const res = UrlFetchApp.fetch(listUrl, { method: "get", muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error("Modellliste konnte nicht abgerufen werden");

  const data = JSON.parse(res.getContentText());
  const models = (data.models || []).filter(m => {
    const methods = m.supportedGenerationMethods || [];
    const name = (m.name || "").toLowerCase();
    return methods.includes("generateContent") && name.includes("gemini") &&
           !name.includes("embedding") && !name.includes("imagen") && !name.includes("veo") &&
           !name.includes("tts") && !name.includes("native-audio") && !name.includes("live") && !name.includes("aqa");
  });

  const standardFlash = models.map(m => m.name).filter(n => n.includes("flash") && !n.includes("lite") && !n.includes("latest")).sort().reverse();
  const flashLatest = models.map(m => m.name).filter(n => n.includes("flash-latest"));
  const flashLite = models.map(m => m.name).filter(n => n.includes("flash") && n.includes("lite")).sort().reverse();
  const proFallback = models.map(m => m.name).filter(n => n.includes("pro") && !n.includes("latest")).sort().reverse();

  const rankedModels = [...new Set([...standardFlash, ...flashLatest, ...flashLite, ...proFallback])].slice(0, 3);
  if (rankedModels.length === 0) throw new Error("Kein kompatibles Modell gefunden.");

  props.setProperty("CACHED_GEMINI_MODELS", JSON.stringify(rankedModels));
  updateLiveStatus(ss, `Aktive Modell-Kette: ${rankedModels.map(m => m.replace("models/", "")).join(" → ")}`, "🤖 Modell-Check", 4, false);
  return rankedModels;
}

function executeGeminiCall(payload, allowedFolders, ss, stepInfo, fileName) {
  let apiKey, modelsToTry;
  try {
    apiKey = getApiKey();
    modelsToTry = getActiveGeminiModels(ss, stepInfo);
  } catch (err) {
    return { folder_name: "API-Fehler", suggested_filename: "", reasoning: err.message, tokens: 0, tokenDisplay: "0" };
  }

  let lastErrorCode = 0;
  let lastErrorMsg = "Unbekannter Fehler";

  for (let m = 0; m < modelsToTry.length; m++) {
    const modelName = modelsToTry[m];
    const shortName = modelName.replace("models/", "");
    const maxRetries = (m === 0) ? 3 : 2;

    const requestPayload = JSON.parse(JSON.stringify(payload));
    if (!requestPayload.generationConfig) requestPayload.generationConfig = {};

    const lowerModel = shortName.toLowerCase();
    let appliedCustomConfig = false;
    if (lowerModel.includes("flash")) {
      requestPayload.generationConfig.thinkingConfig = (lowerModel.includes("2.5") || lowerModel.includes("2.0"))
        ? { thinkingBudget: 0 }
        : { thinkingLevel: "MINIMAL" };
      appliedCustomConfig = true;
    }
    if (requestPayload.generationConfig.mediaResolution) appliedCustomConfig = true;

    const url = "https://generativelanguage.googleapis.com/v1beta/" + modelName + ":generateContent?key=" + apiKey;

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        if (attempt > 1 || m > 0) {
          updateLiveStatus(
            ss,
            `KI-Anfrage (${shortName}, Versuch ${attempt}/${maxRetries}) für "${fileName}"...`,
            stepInfo,
            5
          );
        }

        const response = UrlFetchApp.fetch(url, {
          method: "post",
          contentType: "application/json",
          payload: JSON.stringify(requestPayload),
          muteHttpExceptions: true
        });
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();

        if (responseCode === 200) {
          return parseGeminiJsonResponse(responseText, allowedFolders);
        }

        if (responseCode === 400 && appliedCustomConfig) {
          delete requestPayload.generationConfig.thinkingConfig;
          delete requestPayload.generationConfig.mediaResolution;
          appliedCustomConfig = false;
          updateLiveStatus(ss, `Parameter-Anpassung (HTTP 400) für ${shortName}, wiederhole sofort...`, stepInfo, 4, false);
          continue;
        }

        lastErrorCode = responseCode;
        try { lastErrorMsg = JSON.parse(responseText).error.message || responseText; } catch (e) { lastErrorMsg = responseText; }

        if (responseCode === 404) {
          PropertiesService.getScriptProperties().deleteProperty("CACHED_GEMINI_MODELS");
          updateLiveStatus(ss, `Modell ${shortName} nicht verfügbar (404), wechsle auf Ausweichmodell...`, stepInfo, 4);
          break;
        }
        if ((responseCode === 503 || responseCode === 429) && attempt < maxRetries) {
          const waitSec = attempt * 2;
          updateLiveStatus(
            ss,
            `API-Limit/Überlastung (HTTP ${responseCode}) bei ${shortName} – warte ${waitSec}s vor Versuch ${attempt + 1}/${maxRetries}...`,
            stepInfo,
            waitSec + 2
          );
          Utilities.sleep(waitSec * 1000);
          continue;
        }
        updateLiveStatus(ss, `Modell ${shortName} meldet HTTP ${responseCode} – prüfe nächstes Modell...`, stepInfo, 4);
        break;
      } catch (netErr) {
        lastErrorCode = "Netzwerk";
        lastErrorMsg = netErr.toString();
        if (attempt < maxRetries) {
          updateLiveStatus(ss, `Netzwerkfehler bei "${fileName}" – erneuter Versuch in 1,5s...`, stepInfo, 4);
          Utilities.sleep(1500);
        }
      }
    }
  }
  return { folder_name: "API-Fehler (" + lastErrorCode + ")", suggested_filename: "", reasoning: lastErrorMsg, tokens: 0, tokenDisplay: "0" };
}

function ensureFileExtension(originalName, newName) {
  let sanitized = newName.replace(/[\/\\:*?"<>|]/g, "-");
  const lastDotIndex = originalName.lastIndexOf(".");
  if (lastDotIndex > 0 && (originalName.length - lastDotIndex) <= 6) {
    const ext = originalName.substring(lastDotIndex);
    if (!sanitized.toLowerCase().endsWith(ext.toLowerCase())) return sanitized + ext;
  }
  return sanitized;
}