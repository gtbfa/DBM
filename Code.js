// ==========================================
// KONFIGURATION
// ==========================================
const COMPANY_NAME = "fb-Pflege / Fernanda Bogado";
const INBOX_FOLDER_ID = "1ekDmimq4LJ41-tAOu3Ls610VAAk28zXq";

// KI & Archiv-Einstellungen
const PRIMARY_MODEL = "auto";
const ARCHIVE_SHEET_NAME = "Archiv_Log";
const LIVE_LOG_SHEET_NAME = "Live_Log";
const MAX_LOG_ROWS = 250;
const MAX_RUNTIME_MS = 280000; // Sicherheitsstopp bei ~4 Min. 40 Sek.

// Token-Optimierungs-Limits
const MAX_EXCEL_ROWS = 35;
const MAX_IMAGE_DIMENSION = 1024;

// Namensschema
const NAMING_CONVENTION = "JJJJ-MM-TT_Absender-oder-Thema_Dokumentart (z. B. 2026-11-12_Telekom_Rechnung)";

function getApiKey() {
  const key = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!key || key.trim() === "") {
    throw new Error("Kein API-Schlüssel hinterlegt. Bitte 'GEMINI_API_KEY' in den Projekteinstellungen unter Skripteigenschaften eintragen.");
  }
  return key.trim();
}

function normalizeDupKey(rawName) {
  if (!rawName) return "";
  const withoutExt = rawName.toString().replace(/\.[a-z0-9]{1,6}$/i, "");
  if (typeof normalizeRuleText === "function") {
    return normalizeRuleText(withoutExt).replace(/\s+/g, "_");
  }
  return withoutExt
    .toLowerCase()
    .replace(/ä/g, "ae")
    .replace(/ö/g, "oe")
    .replace(/ü/g, "ue")
    .replace(/ß/g, "ss")
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function buildInvoiceBusinessKey(invNr, amount, bDate) {
  const cleanInv = (invNr || "").toString().trim().toLowerCase().replace(/[^a-z0-9]/g, "");
  const numAmt = Number(amount) || 0;
  if (!cleanInv || cleanInv.length < 3 || numAmt <= 0) return "";
  const datePart = (bDate || "").toString().trim().substring(0, 10);
  return `${cleanInv}|${numAmt.toFixed(2)}|${datePart}`;
}


// ==========================================
// MENÜ & LIVE-PROTOKOLLIERUNG (TOAST + LIVE_LOG)
// ==========================================
function onOpen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureReviewSheetHeaders(ss);

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
// 1. EINZEL-DATEI-ANALYSE FÜR 00_INBOX (INKL. UNTERORDNER)
// ==========================================
function collectInboxFilesRecursive(folder, relativePath, existingRowsMap, filesToProcess, stats, visitedFolderIds) {
  if (!folder || folder.isTrashed()) return;

  const folderId = folder.getId();
  if (visitedFolderIds.has(folderId)) return;
  visitedFolderIds.add(folderId);
  stats.scannedFolders++;

  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.isTrashed()) continue;

    const existingEntry = existingRowsMap.get(file.getId());
    if (existingEntry && !existingEntry.isError) {
      stats.alreadyListedCount++;
      continue;
    }
    filesToProcess.push({
      file: file,
      existingEntry: existingEntry,
      subFolderPath: relativePath
    });
  }

  const subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    const sub = subFolders.next();
    if (sub.isTrashed()) continue;
    const nextPath = relativePath ? `${relativePath} / ${sub.getName()}` : sub.getName();
    collectInboxFilesRecursive(sub, nextPath, existingRowsMap, filesToProcess, stats, visitedFolderIds);
  }
}

function analyzeInboxFiles() {
  const startTime = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Prüfung & Freigabe");
  const whitelistSheet = ss.getSheetByName("Ordner_Whitelist");

  if (!sheet || !whitelistSheet) {
    ss.toast("Tabellenblatt 'Prüfung & Freigabe' oder 'Ordner_Whitelist' fehlt.", "Konfigurationsfehler ⚠️", 8);
    return;
  }

  updateLiveStatus(ss, "Starte Posteingangs-Prüfung (inkl. Unterordner) & prüfe Tabellenstruktur...", "🚀 Start", 4);

  if (typeof setupAccountingSheets === "function") setupAccountingSheets(ss);

  try {
    getApiKey();
  } catch (keyErr) {
    updateLiveStatus(ss, keyErr.message, "Konfigurationsfehler ⚠️", 10);
    return;
  }

  let gmailImportedCount = 0;
  if (typeof AUTO_FETCH_GMAIL_ON_ANALYZE !== "undefined" && AUTO_FETCH_GMAIL_ON_ANALYZE && typeof fetchGmailAttachments === "function") {
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
  const filesToProcess = [];
  const scanStats = { alreadyListedCount: 0, scannedFolders: 0 };

  collectInboxFilesRecursive(inbox, "", existingRowsMap, filesToProcess, scanStats, new Set());

  if (filesToProcess.length === 0) {
    updateLiveStatus(
      ss,
      `Keine neuen Dateien in '00_Inbox' und ${Math.max(0, scanStats.scannedFolders - 1)} Unterordner(n) (${scanStats.alreadyListedCount} Datei(en) warten bereits in 'Prüfung & Freigabe').`,
      "Alles aktuell ✓",
      6
    );
    return;
  }

  const subFolderInfo = scanStats.scannedFolders > 1
    ? ` across ${scanStats.scannedFolders} Ordner(n) inkl. Unterordner`
    : "";
  updateLiveStatus(
    ss,
    `${filesToProcess.length} neue Datei(en)${subFolderInfo} gefunden (${allowedFolders.length} Zielordner, ${knownHashes.size} bekannte Hashes geladen).`,
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

    const { file, existingEntry, subFolderPath } = filesToProcess[i];
    const fileId = file.getId();
    const fileName = file.getName();
    const displayLocation = subFolderPath ? `${subFolderPath} / ${fileName}` : fileName;
    const sizeLabel = formatFileSize(file.getSize());
    const stepPrefix = `📄 Datei ${i + 1}/${filesToProcess.length}`;

    updateLiveStatus(ss, `[1/4] Prüfe Fingerabdruck (MD5) für "${displayLocation}" (${sizeLabel})...`, stepPrefix, 4);

    const fileHash = computeFileHash(file);
    let isExactDup = false;
    let aiResult;

    if (fileHash && knownHashes.has(fileHash)) {
      const dupMatch = knownHashes.get(fileHash);
      isExactDup = true;
      const origRef = dupMatch.originalName ? ` (Original: "${dupMatch.originalName}")` : "";
      const folderRef = dupMatch.folder ? ` in "${dupMatch.folder}"` : "";
      updateLiveStatus(
        ss,
        `[2/4] Exaktes Duplikat erkannt! Identisch mit "${dupMatch.name}"${origRef} (${dupMatch.source}${folderRef}) – 0 Tokens verbraucht.`,
        stepPrefix,
        5
      );
      aiResult = {
        folder_name: dupMatch.folder || "MANUELLE_PRUEFUNG",
        suggested_filename: dupMatch.name || fileName,
        reasoning: `⚠️ EXAKTES DUPLIKAT (${dupMatch.source}${folderRef}): Identischer Inhalt wie "${dupMatch.name}"${origRef}.`,
        tokens: 0,
        tokenDisplay: "0"
      };
    } else {
      let ruleEval = {
        mode: "FALLBACK_LLM",
        boostedFolders: [],
        promptHint: "",
        isInvoiceCandidate: false
      };

      if (typeof buildSortingContext === "function" && typeof evaluateSortingRules === "function") {
        const ctx = buildSortingContext(file, "", subFolderPath, folderMap, "");
        ruleEval = evaluateSortingRules(ctx, folderMap);
        ruleEval.isInvoiceCandidate = Boolean(ctx.isInvoiceCandidate);
      }

      if (ruleEval.mode === "HARD_ROUTE" && ruleEval.folder_name) {
        updateLiveStatus(
          ss,
          `[2/4] Regel-Treffer (${ruleEval.ruleId}) → "${ruleEval.folder_name}" (0 Tokens).`,
          stepPrefix,
          5
        );
        aiResult = {
          folder_name: ruleEval.folder_name,
          suggested_filename: fileName,
          reasoning: `${ruleEval.reasoning} [Regel-Routing]`,
          tokens: 0,
          tokenDisplay: ruleEval.tokenDisplay || "0 (Regel)",
          t: ruleEval.t || "",
          d: "",
          due: "",
          p: ruleEval.p || "",
          b: "",
          v: typeof ruleEval.v === "number" ? ruleEval.v : 0,
          c: "EUR",
          k: ruleEval.k || "",
          s: ""
        };
      } else {
        updateLiveStatus(ss, `[2/4] Lese Dateiinhalt aus "${displayLocation}" aus (OCR / Text / Vorschau)...`, stepPrefix, 5);
        const prep = prepareClassificationPayload(file, allowedFolders, subFolderPath, {
          boostedFolders: ruleEval.boostedFolders || [],
          promptHint: ruleEval.promptHint || "",
          isInvoiceCandidate: ruleEval.isInvoiceCandidate
        });

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
    }

    totalTokensRun += (aiResult.tokens || 0);
    const hasApiError = aiResult.folder_name.startsWith("API-Fehler") || aiResult.folder_name === "Fehler bei Analyse";
    if (hasApiError) errorCount++;

    let statusText = isExactDup ? "⚠️ Duplikat" : "Offen";
    let isDuplicate = isExactDup;
    if (isDuplicate) duplicateCount++;

    const cleanSuggestedName = ensureFileExtension(fileName, aiResult.suggested_filename || fileName);
    const normSuggestedKey = normalizeDupKey(cleanSuggestedName);
    const normOriginalKey = normalizeDupKey(fileName);

    if (!isDuplicate && !hasApiError) {
      const matchedNameEntry =
        (normSuggestedKey && knownNames.get(normSuggestedKey)) ||
        (normOriginalKey && knownNames.get(normOriginalKey));

      if (matchedNameEntry) {
        isDuplicate = true;
        duplicateCount++;
        statusText = "⚠️ Duplikat";
        const refName = matchedNameEntry.name || fileName;
        const refOrig = matchedNameEntry.originalName && matchedNameEntry.originalName !== refName
          ? ` [Original: "${matchedNameEntry.originalName}"]`
          : "";
        aiResult.reasoning = `⚠️ NAMENS-/VERSION-DUPLIKAT (${matchedNameEntry.source} → "${refName}"${refOrig}): ${aiResult.reasoning}`;
      } else if (folderMap[aiResult.folder_name]) {
        try {
          const targetFolder = DriveApp.getFolderById(folderMap[aiResult.folder_name]);
          const existingMatches = targetFolder.getFilesByName(cleanSuggestedName);
          let existsInTarget = false;
          while (existingMatches.hasNext()) {
            if (!existingMatches.next().isTrashed()) {
              existsInTarget = true;
              break;
            }
          }
          if (existsInTarget) {
            isDuplicate = true;
            duplicateCount++;
            statusText = "⚠️ Duplikat";
            aiResult.reasoning = `⚠️ IM ZIELORDNER VORHANDEN: ${aiResult.reasoning}`;
          }
        } catch (driveErr) {}
      }
    }

    if (fileHash && !knownHashes.has(fileHash)) {
      knownHashes.set(fileHash, {
        name: cleanSuggestedName,
        originalName: fileName,
        folder: aiResult.folder_name,
        source: "Aktueller Lauf"
      });
    }
    if (normSuggestedKey && !knownNames.has(normSuggestedKey)) {
      knownNames.set(normSuggestedKey, {
        name: cleanSuggestedName,
        originalName: fileName,
        folder: aiResult.folder_name,
        source: "Aktueller Lauf"
      });
    }
    if (normOriginalKey && !knownNames.has(normOriginalKey)) {
      knownNames.set(normOriginalKey, {
        name: cleanSuggestedName,
        originalName: fileName,
        folder: aiResult.folder_name,
        source: "Aktueller Lauf"
      });
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

function extractPlainLinkLabel(val) {
  if (!val) return "";
  const str = val.toString().trim();
  const m = str.match(/=HYPERLINK\("[^"]*";\s*"([^"]*)"\)/i);
  return m ? m[1].replace(/""/g, '"') : str;
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
      const origLabel = extractPlainLinkLabel(row[0]);
      const folderSuggestion = row[2] ? row[2].toString().trim() : "";
      const suggestedName = row[3] ? row[3].toString().trim() : "";
      const existingId = row[8] ? row[8].toString().trim() : "";
      const existingHash = row[10] ? row[10].toString().trim() : "";

      if (existingId) {
        const isError = folderSuggestion.startsWith("API-Fehler") || folderSuggestion === "Fehler bei Analyse";
        existingRowsMap.set(existingId, { rowNumber: index + 2, isError });

        if (!isError) {
          if (existingHash && !knownHashes.has(existingHash)) {
            knownHashes.set(existingHash, {
              name: suggestedName || origLabel,
              originalName: origLabel,
              folder: folderSuggestion,
              source: "Prüfliste"
            });
          }
          [suggestedName, origLabel].forEach(rawCandidate => {
            const normKey = normalizeDupKey(rawCandidate);
            if (normKey && !knownNames.has(normKey)) {
              knownNames.set(normKey, {
                name: suggestedName || origLabel,
                originalName: origLabel,
                folder: folderSuggestion,
                source: "Prüfliste"
              });
            }
          });
        }
      }
    });
  }

  const archiveSheet = ss.getSheetByName(ARCHIVE_SHEET_NAME);
  if (archiveSheet && archiveSheet.getLastRow() > 1) {
    const archCols = archiveSheet.getLastColumn();
    const archData = archiveSheet.getRange(2, 1, archiveSheet.getLastRow() - 1, archCols).getValues();
    archData.forEach(row => {
      const archNewName = extractPlainLinkLabel(row[0]);
      const archOrigName = extractPlainLinkLabel(row[1]);
      const archFolder = row[2] ? row[2].toString().trim() : "";
      const archHash = (archCols >= 9 && row[8]) ? row[8].toString().trim() : "";

      if (archHash && !knownHashes.has(archHash)) {
        knownHashes.set(archHash, {
          name: archNewName || archOrigName,
          originalName: archOrigName,
          folder: archFolder,
          source: "Archiv"
        });
      }
      [archNewName, archOrigName].forEach(rawCandidate => {
        const normKey = normalizeDupKey(rawCandidate);
        if (normKey && !knownNames.has(normKey)) {
          knownNames.set(normKey, {
            name: archNewName || archOrigName,
            originalName: archOrigName,
            folder: archFolder,
            source: "Archiv"
          });
        }
      });
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
  const bookedBusinessKeys = new Set();

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

    const bType = (data[i][11] || "").toString().trim();
    const bDate = data[i][12] || "";
    const bPartner = data[i][13] || "";
    const bInvNr = data[i][14] || "";
    const bAmount = Number(data[i][15]) || 0;
    const bCurrency = data[i][16] || "";
    const bCategory = data[i][17] || "";
    const bStatus = data[i][18] || "";
    const bDueDate = data[i][19] || "";

    if (isApproved === true && (status === "Offen" || status === "⚠️ Duplikat")) {
      currentStep++;
      const isAiSuggestion = manualOverride === "";
      const targetFolderName = !isAiSuggestion ? manualOverride : aiFolder;
      const targetFolderId = folderMap[targetFolderName];

      const financeGate = typeof canMoveToFinanceFolder === "function"
        ? canMoveToFinanceFolder(targetFolderName, isAiSuggestion, bType, bAmount, originalName, newNameInput, reasoning)
        : { allowed: true, isFinanceTarget: false, isConfirmedInvoice: false, isBankStatement: false };

      if (!financeGate.allowed) {
        sheet.getRange(i + 1, 8).setValue("Blockiert: Keine Rechnung").setBackground("#fff2cc");
        updateLiveStatus(
          ss,
          `(${currentStep}/${totalToMove}) Blockiert: Weder Rechnung noch Kontoauszug für "${targetFolderName}".`,
          "🛡️ Sortierschutz",
          6
        );
        continue;
      }

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

          const bizKey = buildInvoiceBusinessKey(bInvNr, bAmount, bDate);
          const alreadyBookedInRun = bizKey ? bookedBusinessKeys.has(bizKey) : false;

          if (
            financeGate.isFinanceTarget &&
            typeof bookApprovedInvoice === "function" &&
            financeGate.isConfirmedInvoice &&
            !financeGate.isBankStatement &&
            !alreadyBookedInRun
          ) {
            if (bizKey) bookedBusinessKeys.add(bizKey);
            updateLiveStatus(ss, `Buche Rechnungsbeleg "${finalName}" ins Buchungsjournal...`, "📊 Buchhaltung", 4, false);
            bookApprovedInvoice(ss, {
              link: finalLink, type: bType, bDate: bDate, dueDate: bDueDate, partner: bPartner,
              invNr: bInvNr, amount: bAmount, currency: bCurrency,
              category: bCategory, payStatus: bStatus, fileId: fileId, tokens: tokens
            });
          } else if (financeGate.isFinanceTarget && financeGate.isConfirmedInvoice && !financeGate.isBankStatement && alreadyBookedInRun) {
            updateLiveStatus(
              ss,
              `Doppelbuchung verhindert für "${finalName}" (Geschäftsschlüssel ${bizKey} bereits gebucht).`,
              "📊 Buchhaltung",
              4,
              false
            );
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
    rowsToDelete.sort((a, b) => a - b);
    let end = rowsToDelete[rowsToDelete.length - 1];
    let count = 1;
    for (let r = rowsToDelete.length - 2; r >= 0; r--) {
      if (rowsToDelete[r] === end - count) {
        count++;
      } else {
        sheet.deleteRows(end - count + 1, count);
        end = rowsToDelete[r];
        count = 1;
      }
    }
    sheet.deleteRows(end - count + 1, count);
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

function ensureReviewSheetHeaders(ss) {
  const sheet = (ss || SpreadsheetApp.getActiveSpreadsheet()).getSheetByName("Prüfung & Freigabe");
  if (!sheet) return;

  const extraHeaders = [
    "Typ (Einnahme/Ausgabe)",
    "Belegdatum",
    "Geschäftspartner",
    "Rechnungs-Nr.",
    "Brutto-Betrag",
    "Währung",
    "Kategorie",
    "Zahlungsstatus",
    "Fälligkeitsdatum"
  ];

  const targetRange = sheet.getRange(1, 12, 1, extraHeaders.length);
  targetRange
    .setValues([extraHeaders])
    .setFontWeight("bold")
    .setBackground("#1f4e78")
    .setFontColor("#ffffff");

  sheet.getRange("P2:P").setNumberFormat("#,##0.00 €");
}


// ==========================================
// ALLE ORDNER EINBEZIEHEN (OHNE MAX-BEGRENZUNG)
// ==========================================
function selectRelevantFolders(allowedFolders) {
  return allowedFolders.map((path, idx) => ({ idx, path }));
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

    const cleanedRawText = rawText.replace(/^\s*```(?:json)?\s*/i, "").replace(/\s*```\s*$/i, "").trim();
    const rawParsed = JSON.parse(cleanedRawText);

    const isIdxValid = rawParsed.f !== null && rawParsed.f !== "" && !isNaN(Number(rawParsed.f));
    const idx = isIdxValid ? Number(rawParsed.f) : -1;
    let resolvedFolder = (idx >= 0 && idx < allowedFolders.length)
      ? allowedFolders[idx]
      : "MANUELLE_PRUEFUNG";

    const isAltIdxValid = rawParsed.alt_f !== null && rawParsed.alt_f !== "" && !isNaN(Number(rawParsed.alt_f));
    const altIdx = isAltIdxValid ? Number(rawParsed.alt_f) : -1;
    const altFolder = (altIdx >= 0 && altIdx < allowedFolders.length)
      ? allowedFolders[altIdx]
      : "MANUELLE_PRUEFUNG";

    const guarded = typeof applyFinanceSortingGuardrail === "function"
      ? applyFinanceSortingGuardrail(rawParsed, resolvedFolder, altFolder, rawParsed.r || "", rawParsed.t || "")
      : {
          folder_name: resolvedFolder,
          reasoning: rawParsed.r || "",
          docType: rawParsed.t || "",
          amount: rawParsed.v
        };

    const usage = json.usageMetadata || {};
    const totalTok = usage.totalTokenCount || 0;
    const promptTok = usage.promptTokenCount || 0;
    const outTok = usage.candidatesTokenCount || 0;
    const thoughtTok = usage.thoughtsTokenCount || 0;

    const detailStr = thoughtTok > 0
      ? `${totalTok} (In:${promptTok} Out:${outTok} Denk:${thoughtTok})`
      : `${totalTok} (In:${promptTok} Out:${outTok})`;

    return {
      folder_name: guarded.folder_name,
      suggested_filename: rawParsed.n || "",
      reasoning: guarded.reasoning,
      tokens: totalTok,
      tokenDisplay: detailStr,
      t: guarded.docType,
      d: rawParsed.d,
      due: rawParsed.due,
      p: rawParsed.p,
      b: rawParsed.b,
      v: guarded.amount,
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

function buildPrompt(file, indexedFolderItems, subFolderPath = "", promptHint = "") {
  const fallbackDate = Utilities.formatDate(file.getDateCreated(), "Europe/Berlin", "yyyy-MM-dd");
  const compactList = formatIndexedFolders(indexedFolderItems);
  const subFolderHint = subFolderPath ? `, Inbox-Unterordner: "${subFolderPath}"` : "";
  const ruleHintBlock = promptHint ? `\n${promptHint}\n` : "\n";
  const financeRulesBlock = typeof getFinancePromptRules === "function" ? getFinancePromptRules() : "";

  let prompt = `Archiviere Dokument (Original: "${file.getName()}", Fallback-Datum: ${fallbackDate}${subFolderHint}).
${ruleHintBlock}
${financeRulesBlock}

Ordner-Index:
${compactList}

Ausgabe-Regeln:
- is_rechnung: true NUR wenn es sich um eine Rechnung, Quittung, einen Kassenbon oder eine Gutschrift mit Brutto-/Endbetrag > 0 handelt (Dokumente mit "Abrechnung" sind KEINE Rechnungen, sofern nicht explizit als Rechnung/Quittung/Gutschrift ausgewiesen).
- is_kontoauszug: true NUR wenn es sich um einen Bank-Kontoauszug, Finanzstatus oder eine Umsatzübersicht handelt, sonst zwingend false.
- f: Passende Ordner-Nummer (Integer, spezifischsten Unterordner wählen, oder -1 falls unklar). Darf nur bei is_rechnung = true ODER is_kontoauszug = true auf "01_Finanzen" zeigen! Dokumente mit "Abrechnung" (ohne Rechnung/Quittung/Gutschrift/Kontoauszug) gehören in einen passenden Ordner AUSSERHALB von "01_Finanzen".
- alt_f: Beste alternative Ordner-Nummer außerhalb von "01_Finanzen" (oder -1, falls unklar).
- n: Dateiname nach "${NAMING_CONVENTION}" mit Dokumentendatum (sonst ${fallbackDate}). Ohne / : * ? " < > |.
- r: Kurze deutsche Begründung (max. 8 Wörter).`;

  if (typeof getAccountingPromptExtension === "function") {
    prompt += "\n" + getAccountingPromptExtension(COMPANY_NAME);
  }
  return prompt;
}

function getJsonConfig(useLowMediaRes = false) {
  const props = {
    is_rechnung: { type: "BOOLEAN" },
    is_kontoauszug: { type: "BOOLEAN" },
    f: { type: "INTEGER" },
    alt_f: { type: "INTEGER" },
    n: { type: "STRING" },
    r: { type: "STRING" }
  };
  const requiredFields = ["is_rechnung", "is_kontoauszug", "f", "alt_f", "n", "r"];

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
  if (useLowMediaRes) config.mediaResolution = "MEDIA_RESOLUTION_MEDIUM";
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
    if (lowerModel.includes("2.5") || lowerModel.includes("thinking")) {
      requestPayload.generationConfig.thinkingConfig = { thinkingBudget: 256 };
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