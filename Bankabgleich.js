// ==========================================
// BANKABGLEICH & PDF-KONTOAUSZUG-EXTRAKTION
// ==========================================
const BANK_SHEET_NAME = "Kontoauszug";
const MAX_BANK_PDF_CHARS = 15000;
const BANK_FOLDER_KEYWORDS = ["volks bank", "volksbank", "vb 2026", "vb_", "kontoauszug", "bank"];

function setupBankSheet(ss) {
  let sheet = ss.getSheetByName(BANK_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(BANK_SHEET_NAME);
    const headers = [
      "Buchungsdatum",
      "Auftraggeber / Empfänger",
      "Verwendungszweck",
      "Betrag (+/-)",
      "Abgleich-Status",
      "Zugeordneter Beleg",
      "Beleg-Link / Hinweis"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#cfe2f3");
    sheet.setFrozenRows(1);
    sheet.getRange("A:A").setNumberFormat("dd.mm.yyyy");
    sheet.getRange("D:D").setNumberFormat("#,##0.00 €");
    sheet.setColumnWidth(1, 115);
    sheet.setColumnWidth(2, 220);
    sheet.setColumnWidth(3, 300);
    sheet.setColumnWidth(4, 120);
    sheet.setColumnWidth(5, 210);
    sheet.setColumnWidth(6, 220);
    sheet.setColumnWidth(7, 220);
  }
  return sheet;
}


// ==========================================
// KONTOAUSZUG-PDFs AUS DRIVE AUSLESEN & ABGLEICHEN
// ==========================================
function importBankStatementPdfsFromDrive() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bankSheet = setupBankSheet(ss);

  try {
    if (typeof getApiKey === "function") getApiKey();
  } catch (e) {
    ss.toast(e.message, "API-Schlüssel fehlt ⚠️", 8);
    return;
  }

  const props = PropertiesService.getScriptProperties();
  const scannedRaw = props.getProperty("BANK_PDF_SCANNED_IDS");
  const scannedPdfIds = new Set(scannedRaw ? JSON.parse(scannedRaw) : []);

  const candidateFiles = [];
  const seenFileIds = new Set();
  const searchedFolderIds = new Set();

  // 1. Alle Ordner unter "01_Finanzen" sammeln (inkl. rekursiver Unterordnersuche)
  const financeFolders = [];
  const whitelistSheet = ss.getSheetByName("Ordner_Whitelist");
  let finRootId = "17Wte3udpVOb5gHSSYsSmdbs9TNlxlJ6i";

  if (whitelistSheet && whitelistSheet.getLastRow() >= 2) {
    const wlData = whitelistSheet.getRange(2, 1, whitelistSheet.getLastRow() - 1, 2).getValues();
    wlData.forEach(row => {
      const path = (row[0] || "").toString().trim().replace(/\s*\/\s*/g, " / ");
      const folderId = (row[1] || "").toString().trim();
      if (!path || !folderId || folderId.length < 10) return;

      if (path === "01_Finanzen") finRootId = folderId;
      if (path.startsWith("01_Finanzen")) {
        financeFolders.push({ path, id: folderId });
      }
    });
  }

  // Falls in Ordner_Whitelist nur der Hauptordner "01_Finanzen" steht, alle Unterordner direkt aus Drive holen
  if (financeFolders.length <= 1) {
    try {
      const rootFolder = DriveApp.getFolderById(finRootId);
      if (!rootFolder.isTrashed()) {
        collectBankSubfoldersRecursive(rootFolder, "01_Finanzen", financeFolders, new Set());
      }
    } catch (e) {}
  }

  financeFolders.forEach(({ path, id }) => {
    if (searchedFolderIds.has(id)) return;
    searchedFolderIds.add(id);

    const lowerPath = path.toLowerCase();
    if (lowerPath.includes("kreditkarte")) return;

    if (BANK_FOLDER_KEYWORDS.some(kw => lowerPath.includes(kw))) {
      try {
        const folder = DriveApp.getFolderById(id);
        const files = folder.getFiles();
        while (files.hasNext()) {
          const f = files.next();
          if (!seenFileIds.has(f.getId()) && isPdfFile(f)) {
            seenFileIds.add(f.getId());
            candidateFiles.push(f);
          }
        }
      } catch (err) {}
    }
  });

  // 2. Zusätzlich in 00_Inbox nach Kontoauszug-/Volksbank-PDFs suchen
  if (typeof INBOX_FOLDER_ID !== "undefined" && INBOX_FOLDER_ID) {
    try {
      const inbox = DriveApp.getFolderById(INBOX_FOLDER_ID);
      const files = inbox.getFiles();
      while (files.hasNext()) {
        const f = files.next();
        const lName = f.getName().toLowerCase();
        const isBankName = ["kontoauszug", "finanzstatus", "umsatz", "volksbank", "auszug", "vb_", "vb "].some(k => lName.includes(k));
        if (!seenFileIds.has(f.getId()) && isPdfFile(f) && isBankName) {
          seenFileIds.add(f.getId());
          candidateFiles.push(f);
        }
      }
    } catch (err) {}
  }

  const newPdfs = candidateFiles.filter(f => !scannedPdfIds.has(f.getId()));

  if (newPdfs.length === 0) {
    if (candidateFiles.length === 0) {
      SpreadsheetApp.getUi().alert(
        "Keine Kontoauszug-PDFs gefunden",
        "Es wurden weder unter '01_Finanzen / ... / Volks Bank' noch in '00_Inbox' PDF-Dateien gefunden.",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    ss.toast("Alle gefundenen Kontoauszug-PDFs wurden bereits eingelesen. Starte Abgleich...", "Hinweis", 5);
    reconcileBankStatement();
    return;
  }

  const existingSignatures = new Set();
  if (bankSheet.getLastRow() >= 2) {
    const existingRows = bankSheet.getRange(2, 1, bankSheet.getLastRow() - 1, 4).getValues();
    existingRows.forEach(r => {
      const d = parseFlexibleDate(r[0]);
      const amt = parseGermanAmount(r[3]);
      const name = (r[1] || "").toString().toLowerCase().trim().substring(0, 15);
      if (d && amt !== 0) {
        const dKey = Utilities.formatDate(d, "Europe/Berlin", "yyyy-MM-dd");
        existingSignatures.add(`${dKey}|${amt.toFixed(2)}|${name}`);
      }
    });
  }

  let totalAddedRows = 0;
  let totalTokens = 0;
  const newRowsToAppend = [];
  const errorsList = [];
  const batch = newPdfs.slice(0, 5);

  for (let i = 0; i < batch.length; i++) {
    const pdfFile = batch[i];
    const stepInfo = `Kontoauszug ${i + 1}/${batch.length}`;

    if (typeof updateLiveStatus === "function") {
      updateLiveStatus(ss, `${stepInfo}: Analysiere "${pdfFile.getName()}"...`, "📄 Kontoauszug-KI");
    }

    const extraction = extractTransactionsFromBankPdf(pdfFile, ss, stepInfo);
    if (extraction.error) {
      errorsList.push(`${pdfFile.getName()}: ${extraction.error}`);
      continue;
    }

    totalTokens += (extraction.tokens || 0);
    const txList = extraction.transactions || [];

    if (txList.length === 0) {
      errorsList.push(`${pdfFile.getName()}: 0 Umsätze im PDF erkannt`);
      continue;
    }

    txList.forEach(tx => {
      const dateObj = parseFlexibleDate(tx.d);
      const amount = Number(tx.a) || 0;
      const partner = (tx.p || "").toString().trim();
      const purpose = (tx.z || "").toString().trim();

      if (!dateObj || amount === 0) return;

      const dKey = Utilities.formatDate(dateObj, "Europe/Berlin", "yyyy-MM-dd");
      const sig = `${dKey}|${amount.toFixed(2)}|${partner.toLowerCase().substring(0, 15)}`;
      if (existingSignatures.has(sig)) return;

      existingSignatures.add(sig);
      newRowsToAppend.push([dateObj, partner, purpose, amount, "", "", ""]);
      totalAddedRows++;
    });

    scannedPdfIds.add(pdfFile.getId());
  }

  if (newRowsToAppend.length > 0) {
    newRowsToAppend.sort((a, b) => a[0] - b[0]);
    const startRow = bankSheet.getLastRow() + 1;
    bankSheet.getRange(startRow, 1, newRowsToAppend.length, 7).setValues(newRowsToAppend);
    bankSheet.getRange(startRow, 1, newRowsToAppend.length, 1).setNumberFormat("dd.mm.yyyy");
    bankSheet.getRange(startRow, 4, newRowsToAppend.length, 1).setNumberFormat("#,##0.00 €");
  }

  props.setProperty("BANK_PDF_SCANNED_IDS", JSON.stringify([...scannedPdfIds]));

  if (errorsList.length > 0 && totalAddedRows === 0) {
    SpreadsheetApp.getUi().alert("Kontoauszug konnte nicht ausgelesen werden ⚠️", errorsList.join("\n"), SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  ss.toast(
    `${totalAddedRows} Buchung(en) aus ${batch.length} PDF(s) importiert (${totalTokens} Tokens). Starte Abgleich...`,
    "PDF-Import fertig ✓",
    6
  );

  reconcileBankStatement();
}

function collectBankSubfoldersRecursive(folder, currentPath, targetList, seenIds) {
  if (!seenIds.has(folder.getId())) {
    seenIds.add(folder.getId());
    targetList.push({ path: currentPath, id: folder.getId() });
  }
  const subs = folder.getFolders();
  while (subs.hasNext()) {
    const sub = subs.next();
    if (!sub.isTrashed()) {
      collectBankSubfoldersRecursive(sub, `${currentPath} / ${sub.getName()}`, targetList, seenIds);
    }
  }
}

function extractTransactionsFromBankPdf(file, ss, stepInfo) {
  const fallbackYear = Utilities.formatDate(file.getDateCreated(), "Europe/Berlin", "yyyy");
  const baseInstructions = `Extrahiere alle einzelnen Kontobewegungen (Umsätze) aus diesem Bank-Kontoauszug.
Dateiname: "${file.getName()}" (Fallback-Jahr: ${fallbackYear})

Regeln:
- Ignoriere Zeilen wie "Alter Kontostand", "Neuer Kontostand", "Übertrag", "Summe Soll/Haben" oder reine Hinweise.
- d: Buchungsdatum im Format JJJJ-MM-DD. WICHTIG: Steht bei den Buchungen nur Tag und Monat (z.B. "02.01."), ermittle das richtige Jahr aus dem Zeitraum im Kopfbereich des Auszugs (sonst ${fallbackYear}).
- p: Auftraggeber oder Empfänger (Name der Gegenseite).
- z: Verwendungszweck / Buchungstext (kompakt, inkl. evtl. Rechnungsnummern).
- a: Betrag als Zahl. WICHTIG ZUM VORZEICHEN: Ausgaben, Lastschriften, Kartenzahlungen, Gebühren und ausgehende Überweisungen (Soll / "-") MÜSSEN zwingend NEGATIV sein (z.B. -283.26). Einnahmen, Gutschriften und eingehende Überweisungen (Haben / "+") MÜSSEN POSITIV sein (z.B. 321.00).`;

  const fullText = extractFullTextFromPdfViaDrive(file, MAX_BANK_PDF_CHARS);
  let parts = [];

  if (fullText && fullText.trim().length >= 80) {
    parts = [{ text: `${baseInstructions}\n\nKONTOAUSZUG-TEXT:\n${fullText}` }];
  } else {
    const sizeMB = file.getSize() / (1024 * 1024);
    if (sizeMB > 15) {
      return { error: "PDF ist größer als 15 MB und enthält keinen kopierbaren Text", transactions: [], tokens: 0 };
    }
    const base64Pdf = Utilities.base64Encode(file.getBlob().getBytes());
    parts = [
      { text: baseInstructions },
      { inline_data: { mime_type: "application/pdf", data: base64Pdf } }
    ];
  }

  const payload = {
    contents: [{ parts: parts }],
    generationConfig: {
      response_mime_type: "application/json",
      response_schema: {
        type: "OBJECT",
        properties: {
          transactions: {
            type: "ARRAY",
            items: {
              type: "OBJECT",
              properties: {
                d: { type: "STRING" },
                p: { type: "STRING" },
                z: { type: "STRING" },
                a: { type: "NUMBER" }
              },
              required: ["d", "p", "z", "a"]
            }
          }
        },
        required: ["transactions"]
      }
    }
  };

  const apiKey = getApiKey();
  let models = ["models/gemini-2.5-flash"];
  if (typeof getActiveGeminiModels === "function") {
    try { models = getActiveGeminiModels(ss, stepInfo); } catch (e) {}
  }

  let lastErr = "API nicht erreichbar";

  for (let m = 0; m < models.length; m++) {
    const modelName = models[m];
    const shortName = modelName.replace("models/", "").toLowerCase();
    const reqPayload = JSON.parse(JSON.stringify(payload));

    let hasThinkingCfg = false;
    if (shortName.includes("flash")) {
      reqPayload.generationConfig.thinkingConfig = (shortName.includes("2.5") || shortName.includes("2.0"))
        ? { thinkingBudget: 0 }
        : { thinkingLevel: "MINIMAL" };
      hasThinkingCfg = true;
    }

    const url = `https://generativelanguage.googleapis.com/v1beta/${modelName}:generateContent?key=${apiKey}`;

    for (let attempt = 1; attempt <= 2; attempt++) {
      let res = UrlFetchApp.fetch(url, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(reqPayload),
        muteHttpExceptions: true
      });

      if (res.getResponseCode() === 400 && hasThinkingCfg) {
        delete reqPayload.generationConfig.thinkingConfig;
        hasThinkingCfg = false;
        continue;
      }

      if (res.getResponseCode() === 200) {
        const json = JSON.parse(res.getContentText());
        const parsed = JSON.parse(json.candidates[0].content.parts[0].text);
        const tokens = json.usageMetadata ? json.usageMetadata.totalTokenCount : 0;
        return { transactions: parsed.transactions || [], tokens: tokens, error: "" };
      }

      lastErr = `HTTP ${res.getResponseCode()}`;
      if ((res.getResponseCode() === 429 || res.getResponseCode() === 503) && attempt < 2) {
        Utilities.sleep(2500);
        continue;
      }
      break;
    }
  }

  return { error: lastErr, transactions: [], tokens: 0 };
}

function extractFullTextFromPdfViaDrive(file, maxChars) {
  if (typeof Drive === "undefined" || !Drive.Files) return "";
  let tempFileId = null;

  try {
    const pdfBlob = file.getBlob().setContentType("application/pdf");
    if (typeof Drive.Files.create === "function") {
      tempFileId = Drive.Files.create(
        { name: "Temp_Bank_OCR_" + file.getName(), mimeType: "application/vnd.google-apps.document" },
        pdfBlob
      ).id;
    } else if (typeof Drive.Files.insert === "function") {
      tempFileId = Drive.Files.insert(
        { title: "Temp_Bank_OCR_" + file.getName(), mimeType: "application/vnd.google-apps.document" },
        pdfBlob,
        { ocr: true, ocrLanguage: "de" }
      ).id;
    }

    if (!tempFileId) return "";

    const exportRes = UrlFetchApp.fetch(`https://www.googleapis.com/drive/v3/files/${tempFileId}/export?mimeType=text/plain`, {
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });

    const rawText = exportRes.getResponseCode() === 200
      ? exportRes.getContentText("UTF-8")
      : (DocumentApp.openById(tempFileId).getBody().getText() || "");

    return rawText.replace(/[ \t]+/g, " ").trim().substring(0, maxChars);
  } catch (e) {
    return "";
  } finally {
    if (tempFileId) {
      try {
        if (typeof Drive.Files.delete === "function") Drive.Files.delete(tempFileId);
        else if (typeof Drive.Files.remove === "function") Drive.Files.remove(tempFileId);
        else DriveApp.getFileById(tempFileId).setTrashed(true);
      } catch (delErr) {}
    }
  }
}

function isPdfFile(file) {
  const name = file.getName().toLowerCase();
  const mime = (file.getMimeType() || "").toLowerCase();
  return name.endsWith(".pdf") || mime === "application/pdf";
}

function resetBankPdfScanLog() {
  PropertiesService.getScriptProperties().deleteProperty("BANK_PDF_SCANNED_IDS");
  SpreadsheetApp.getActiveSpreadsheet().toast("Kontoauszug-PDF-Verlauf wurde zurückgesetzt.", "Zurückgesetzt ✓", 5);
}


// ==========================================
// ABGLEICH-ENGINE: KONTOAUSZUG VS. BUCHUNGSJOURNAL (0 TOKENS)
// ==========================================
function reconcileBankStatement() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bankSheet = setupBankSheet(ss);
  const journal = typeof getOrCreateBuchungsjournal === "function"
    ? getOrCreateBuchungsjournal(ss)
    : ss.getSheetByName("Buchungsjournal");

  if (!journal) {
    ss.toast("Tabellenblatt 'Buchungsjournal' nicht gefunden.", "Fehler ⚠️", 6);
    return;
  }

  normalizeVolksbankCsvIfNeeded(bankSheet);

  const lastBankRow = bankSheet.getLastRow();
  if (lastBankRow < 2) {
    ss.toast("Keine Umsätze im Blatt 'Kontoauszug' vorhanden.", "Hinweis", 5);
    return;
  }

  if (typeof updateLiveStatus === "function") {
    updateLiveStatus(ss, "Gleiche Kontoauszug mit Buchungsjournal ab...", "🏦 Bankabgleich");
  }

  const journalEntries = [];
  if (journal.getLastRow() >= 2) {
    const numJRows = journal.getLastRow() - 1;
    const jCols = Math.max(journal.getLastColumn(), 12);
    const jData = journal.getRange(2, 1, numJRows, jCols).getValues();
    const jFormulas = journal.getRange(2, 10, numJRows, 1).getFormulas();
    jData.forEach((r, idx) => {
      const amount = parseGermanAmount(r[5]);
      if (amount <= 0) return;
      journalEntries.push({
        rowNumber: idx + 2,
        month: (r[0] || "").toString().trim(),
        bDate: parseFlexibleDate(r[1]),
        type: (r[2] || "").toString().trim(),
        partner: (r[3] || "").toString().trim(),
        invNr: (r[4] || "").toString().trim(),
        amount: amount,
        category: (r[7] || "").toString().trim(),
        status: (r[8] || "").toString().trim(),
        linkFormula: jFormulas[idx][0] || (r[9] || "").toString(),
        dueDate: parseFlexibleDate(r[10]),
        matchedInThisRun: false
      });
    });
  }

  const bankData = bankSheet.getRange(2, 1, lastBankRow - 1, 7).getValues();
  const statusUpdates = [];
  const bgColors = [];

  let matchedIncomeCount = 0;
  let matchedExpenseCount = 0;
  let missingReceiptCount = 0;
  let journalChanged = false;

  for (let i = 0; i < bankData.length; i++) {
    const bDate = parseFlexibleDate(bankData[i][0]);
    const counterparty = (bankData[i][1] || "").toString().trim();
    const purpose = (bankData[i][2] || "").toString().trim();
    const amount = parseGermanAmount(bankData[i][3]);

    if ((!bDate && !counterparty && !purpose) || amount === 0 || counterparty.toLowerCase() === "total") {
      statusUpdates.push(["", "", ""]);
      bgColors.push([null, null, null]);
      continue;
    }

    const fullText = `${counterparty} ${purpose}`;
    const lowerText = fullText.toLowerCase();

    if (amount > 0) {
      const matchedInvoices = findMatchingIncomeEntries(journalEntries, fullText, amount, bDate);

      if (matchedInvoices.length > 0) {
        matchedIncomeCount += matchedInvoices.length;
        const invLabels = matchedInvoices.map(m => m.invNr ? `${m.partner} (${m.invNr})` : m.partner).join(", ");
        const firstLink = matchedInvoices[0].linkFormula || "Im Journal verbucht";

        matchedInvoices.forEach(entry => {
          entry.matchedInThisRun = true;
          if (entry.status !== "Bezahlt") {
            entry.status = "Bezahlt";
            journal.getRange(entry.rowNumber, 9).setValue("Bezahlt");
            if (bDate && !entry.bDate) journal.getRange(entry.rowNumber, 2).setValue(bDate);
            journalChanged = true;
          }
        });

        statusUpdates.push(["✅ Einnahme bezahlt", invLabels, firstLink]);
        bgColors.push(["#d9ead3", "#d9ead3", "#d9ead3"]);
      } else {
        statusUpdates.push(["⚠️ Einnahme ohne Beleg", "Keine offene Rechnung gefunden", "Prüfen / Rechnung erfassen"]);
        bgColors.push(["#fff2cc", "#fff2cc", "#fff2cc"]);
      }
    } else {
      const absAmount = Math.abs(amount);
      const matchedExpense = findMatchingExpenseEntry(journalEntries, fullText, absAmount, bDate);

      if (matchedExpense) {
        matchedExpenseCount++;
        matchedExpense.matchedInThisRun = true;
        if (matchedExpense.status !== "Bezahlt") {
          matchedExpense.status = "Bezahlt";
          journal.getRange(matchedExpense.rowNumber, 9).setValue("Bezahlt");
          journalChanged = true;
        }

        const label = matchedExpense.invNr
          ? `${matchedExpense.partner} (${matchedExpense.invNr})`
          : `${matchedExpense.partner} [${matchedExpense.category}]`;

        statusUpdates.push(["✅ Beleg vorhanden", label, matchedExpense.linkFormula || "Im Journal vorhanden"]);
        bgColors.push(["#d9ead3", "#d9ead3", "#d9ead3"]);
      } else if (lowerText.includes("privatentnahme") || lowerText.includes("saque")) {
        statusUpdates.push(["ℹ️ Privatentnahme", "Kein Beleg erforderlich", "Privatbuchung"]);
        bgColors.push(["#efefef", "#efefef", "#efefef"]);
      } else {
        missingReceiptCount++;
        statusUpdates.push(["❌ Beleg in Drive fehlt!", "Kein passender Beleg im Journal", "Bitte Beleg in 00_Inbox ablegen"]);
        bgColors.push(["#f4cccc", "#f4cccc", "#f4cccc"]);
      }
    }
  }

  const targetRange = bankSheet.getRange(2, 5, statusUpdates.length, 3);
  targetRange.setValues(statusUpdates);
  targetRange.setBackgrounds(bgColors);

  renderOpenItemsSummary(bankSheet, journalEntries);

  if (journalChanged && typeof syncJournalToMonthlySheets === "function") {
    syncJournalToMonthlySheets();
  }

  ss.toast(
    `✅ ${matchedIncomeCount} Einnahme(n) & ${matchedExpenseCount} Ausgabe(n) zugeordnet | ❌ ${missingReceiptCount} fehlende Belege`,
    "Bankabgleich abgeschlossen ✓",
    10
  );
}


// ==========================================
// MATCHING-LOGIK (EINNAHMEN & AUSGABEN)
// ==========================================
function findMatchingIncomeEntries(journalEntries, fullText, bankAmount, bankDate) {
  const openIncomes = journalEntries.filter(e => e.type === "Einnahme" && !e.matchedInThisRun);
  if (openIncomes.length === 0) return [];

  const byInvNr = [];
  openIncomes.forEach(entry => {
    if (!entry.invNr) return;
    const cleanInv = entry.invNr.trim();
    const shortInv = cleanInv.includes("-") ? cleanInv.split("-").pop() : cleanInv;
    if (cleanInv.length >= 4 && fullText.includes(cleanInv)) {
      byInvNr.push(entry);
    } else if (shortInv.length >= 4 && new RegExp(`\\b${escapeRegex(shortInv)}\\b`).test(fullText)) {
      byInvNr.push(entry);
    }
  });

  if (byInvNr.length > 0) {
    const sumMatched = byInvNr.reduce((acc, item) => acc + item.amount, 0);
    if (Math.abs(sumMatched - bankAmount) <= 1.00) return byInvNr;
    const singleExact = byInvNr.find(item => Math.abs(item.amount - bankAmount) <= 0.50);
    if (singleExact) return [singleExact];
  }

  let bestCandidate = null;
  let bestScore = 0;

  openIncomes.forEach(entry => {
    const diff = Math.abs(entry.amount - bankAmount);
    if (diff > 0.50) return;

    let score = (diff <= 0.02) ? 10 : 6;
    score += computeNameOverlapScore(fullText, entry.partner);

    if (bankDate && entry.bDate) {
      const daysDiff = (bankDate - entry.bDate) / (1000 * 60 * 60 * 24);
      if (daysDiff >= -5 && daysDiff <= 60) score += 4;
    }

    if (score > bestScore) {
      bestScore = score;
      bestCandidate = entry;
    }
  });

  return (bestCandidate && bestScore >= 12) ? [bestCandidate] : [];
}

function findMatchingExpenseEntry(journalEntries, fullText, absBankAmount, bankDate) {
  const openExpenses = journalEntries.filter(e => e.type === "Ausgabe" && !e.matchedInThisRun);
  if (openExpenses.length === 0) return null;

  let bestCandidate = null;
  let bestScore = 0;

  openExpenses.forEach(entry => {
    if (entry.invNr && entry.invNr.length >= 4 && fullText.includes(entry.invNr)) {
      bestCandidate = entry;
      bestScore = 99;
      return;
    }

    const diff = Math.abs(entry.amount - absBankAmount);
    if (diff > 0.50) return;

    let score = (diff <= 0.02) ? 10 : 5;
    score += computeNameOverlapScore(fullText, `${entry.partner} ${entry.category}`);

    if (bankDate && entry.bDate) {
      const daysDiff = Math.abs((bankDate - entry.bDate) / (1000 * 60 * 60 * 24));
      if (daysDiff <= 45) score += 4;
    }

    if (score > bestScore) {
      bestScore = score;
      bestCandidate = entry;
    }
  });

  return (bestCandidate && bestScore >= 12) ? bestCandidate : null;
}

function computeNameOverlapScore(bankText, partnerText) {
  if (!bankText || !partnerText) return 0;
  const ignore = new Set(["gmbh", "ag", "ug", "kg", "ohg", "frau", "herr", "hr", "fr", "sgb", "xi", "v", "rechnung", "rechn"]);
  const bLower = bankText.toLowerCase();
  const words = partnerText.toLowerCase().match(/[a-zäöüß]{4,}/g) || [];

  let hits = 0;
  words.forEach(w => {
    if (!ignore.has(w) && bLower.includes(w)) {
      hits += 5;
    }
  });
  return Math.min(hits, 10);
}


// ==========================================
// OFFENE POSTEN ÜBERSICHT (SPALTEN I BIS N)
// ==========================================
function renderOpenItemsSummary(bankSheet, journalEntries) {
  const maxRows = Math.max(bankSheet.getMaxRows(), 50);
  bankSheet.getRange(1, 9, maxRows, 6).clear();

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const unpaidEntries = journalEntries.filter(e => e.status.toLowerCase() !== "bezahlt");
  const unpaidIncomes = unpaidEntries.filter(e => e.type === "Einnahme");
  const unpaidExpenses = unpaidEntries.filter(e => e.type === "Ausgabe");

  bankSheet.getRange("I1:N1").merge()
    .setValue(`📌 OFFENE KUNDENRECHNUNGEN / EINNAHMEN (${unpaidIncomes.length})`)
    .setFontWeight("bold")
    .setBackground("#fff2cc");

  bankSheet.getRange("I2:N2")
    .setValues([["Rechnungsdatum", "Kunde / Partner", "Rechnungs-Nr.", "Fällig am", "Offener Betrag", "Status"]])
    .setFontWeight("bold")
    .setBackground("#f3f3f3");

  let currentRow = 3;
  if (unpaidIncomes.length === 0) {
    bankSheet.getRange(currentRow, 9, 1, 6).merge().setValue("Alle erfassten Kundenrechnungen sind bezahlt ✓");
    currentRow += 2;
  } else {
    unpaidIncomes.forEach(item => {
      const isOverdue = item.dueDate && item.dueDate < today;
      const statusLabel = isOverdue ? "⚠️ Überfällig!" : "⏳ Offen";
      bankSheet.getRange(currentRow, 9, 1, 6).setValues([[
        item.bDate || "",
        item.partner,
        item.invNr || "-",
        item.dueDate || "",
        item.amount,
        statusLabel
      ]]);
      bankSheet.getRange(currentRow, 9).setNumberFormat("dd.mm.yyyy");
      if (item.dueDate) bankSheet.getRange(currentRow, 12).setNumberFormat("dd.mm.yyyy");
      bankSheet.getRange(currentRow, 13).setNumberFormat("#,##0.00 €");
      if (isOverdue) {
        bankSheet.getRange(currentRow, 14).setBackground("#f4cccc").setFontWeight("bold");
      }
      currentRow++;
    });

    bankSheet.getRange(currentRow, 12).setValue("Summe Offen:").setFontWeight("bold");
    bankSheet.getRange(currentRow, 13)
      .setFormula(`=SUM(M3:M${currentRow - 1})`)
      .setFontWeight("bold")
      .setNumberFormat("#,##0.00 €");
    currentRow += 2;
  }

  bankSheet.getRange(currentRow, 9, 1, 6).merge()
    .setValue(`📌 NOCH NICHT ABGEBUCHTE AUSGABEN (${unpaidExpenses.length})`)
    .setFontWeight("bold")
    .setBackground("#fce5cd");
  currentRow++;

  bankSheet.getRange(currentRow, 9, 1, 6)
    .setValues([["Belegdatum", "Lieferant / Partner", "Rechnungs-Nr.", "Kategorie", "Offener Betrag", "Status"]])
    .setFontWeight("bold")
    .setBackground("#f3f3f3");
  currentRow++;

  if (unpaidExpenses.length === 0) {
    bankSheet.getRange(currentRow, 9, 1, 6).merge().setValue("Keine offenen Ausgaben-Rechnungen vorhanden ✓");
  } else {
    unpaidExpenses.forEach(item => {
      bankSheet.getRange(currentRow, 9, 1, 6).setValues([[
        item.bDate || "",
        item.partner,
        item.invNr || "-",
        item.category,
        item.amount,
        "⏳ Noch nicht abgebucht"
      ]]);
      bankSheet.getRange(currentRow, 9).setNumberFormat("dd.mm.yyyy");
      bankSheet.getRange(currentRow, 13).setNumberFormat("#,##0.00 €");
      currentRow++;
    });
  }

  bankSheet.setColumnWidth(8, 30);
  bankSheet.autoResizeColumns(9, 6);
}


// ==========================================
// CSV-AUTO-ERKENNUNG & PARSER-HELFER
// ==========================================
function normalizeVolksbankCsvIfNeeded(bankSheet) {
  const lastRow = bankSheet.getLastRow();
  const lastCol = bankSheet.getLastColumn();
  if (lastRow < 2) return;

  const firstCell = (bankSheet.getRange(1, 1).getValue() || "").toString();
  if (lastCol === 1 && firstCell.includes(";")) {
    const rawLines = bankSheet.getRange(1, 1, lastRow, 1).getValues().map(r => r[0].toString());
    const parsedGrid = rawLines.map(line => line.split(";").map(cell => cell.replace(/^"|"$/g, "").trim()));
    const maxCols = Math.max(...parsedGrid.map(r => r.length));
    const paddedGrid = parsedGrid.map(r => {
      while (r.length < maxCols) r.push("");
      return r;
    });
    bankSheet.clearContents();
    bankSheet.getRange(1, 1, paddedGrid.length, maxCols).setValues(paddedGrid);
  }

  const scanRows = Math.min(bankSheet.getLastRow(), 15);
  const scanCols = bankSheet.getLastColumn();
  const topData = bankSheet.getRange(1, 1, scanRows, scanCols).getValues();

  let headerRowIdx = -1;
  let colDate = -1;
  let colName = -1;
  let colPurpose = -1;
  let colAmount = -1;

  for (let r = 0; r < topData.length; r++) {
    const rowLower = topData[r].map(c => (c || "").toString().toLowerCase().trim());
    const dIdx = rowLower.findIndex(h => h === "buchungstag" || h === "buchungsdatum" || h === "datum");
    const aIdx = rowLower.findIndex(h => h === "betrag" || h === "betrag (+/-)" || h === "umsatz");

    if (dIdx !== -1 && aIdx !== -1) {
      headerRowIdx = r;
      colDate = dIdx;
      colAmount = aIdx;
      colName = rowLower.findIndex(h => h.includes("zahlungsbeteiligter") || h.includes("auftraggeber") || h.includes("empfänger") || h === "name");
      colPurpose = rowLower.findIndex(h => h.includes("verwendungszweck") || h.includes("buchungstext") || h.includes("beschreibung"));
      break;
    }
  }

  if (headerRowIdx === 0 && colDate === 0 && colName === 1 && colPurpose === 2 && colAmount === 3) {
    return;
  }

  if (headerRowIdx !== -1) {
    const allRows = bankSheet.getRange(headerRowIdx + 2, 1, bankSheet.getLastRow() - headerRowIdx - 1, scanCols).getValues();
    const cleanRows = [];

    allRows.forEach(r => {
      const dVal = parseFlexibleDate(r[colDate]);
      const amtVal = parseGermanAmount(r[colAmount]);
      const nameVal = colName !== -1 ? (r[colName] || "").toString().trim() : "";
      const purpVal = colPurpose !== -1 ? (r[colPurpose] || "").toString().trim() : "";
      if (!dVal && amtVal === 0 && !nameVal && !purpVal) return;

      cleanRows.push([dVal || r[colDate], nameVal, purpVal, amtVal, "", "", ""]);
    });

    bankSheet.clear();
    const headers = [
      "Buchungsdatum", "Auftraggeber / Empfänger", "Verwendungszweck",
      "Betrag (+/-)", "Abgleich-Status", "Zugeordneter Beleg", "Beleg-Link / Hinweis"
    ];
    bankSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    bankSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#cfe2f3");
    bankSheet.setFrozenRows(1);

    if (cleanRows.length > 0) {
      bankSheet.getRange(2, 1, cleanRows.length, 7).setValues(cleanRows);
      bankSheet.getRange(2, 1, cleanRows.length, 1).setNumberFormat("dd.mm.yyyy");
      bankSheet.getRange(2, 4, cleanRows.length, 1).setNumberFormat("#,##0.00 €");
    }
  }
}

function parseGermanAmount(val) {
  if (typeof val === "number") return val;
  if (!val) return 0;
  let str = val.toString().replace(/[€EUR\s]/gi, "").trim();
  if (!str) return 0;

  let isNegative = false;
  if (str.endsWith("-")) {
    isNegative = true;
    str = str.slice(0, -1);
  } else if (str.startsWith("-")) {
    isNegative = true;
    str = str.slice(1);
  }

  if (str.includes(".") && str.includes(",")) {
    str = str.replace(/\./g, "").replace(",", ".");
  } else if (str.includes(",")) {
    str = str.replace(",", ".");
  }

  const num = parseFloat(str);
  if (isNaN(num)) return 0;
  return isNegative ? -Math.abs(num) : num;
}

function parseFlexibleDate(val) {
  if (!val) return null;
  if (val instanceof Date && !isNaN(val.getTime())) return val;
  const str = val.toString().trim();

  let m = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));

  m = str.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})/);
  if (m) {
    let year = Number(m[3]);
    if (year < 100) year += 2000;
    return new Date(year, Number(m[2]) - 1, Number(m[1]));
  }
  return null;
}

function escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}