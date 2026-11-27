// ==========================================
// DOKUMENT-EXTRAKTION & DRIVE-HILFSFUNKTIONEN
// ==========================================
function computeFileHash(file) {
  try {
    if (typeof Drive !== "undefined" && Drive.Files) {
      try {
        const meta = Drive.Files.get(file.getId(), { fields: "md5Checksum" });
        if (meta && meta.md5Checksum) return meta.md5Checksum.toLowerCase();
      } catch (driveErr) {}
    }
    if (file.getSize() / (1024 * 1024) > 15) return "";
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, file.getBlob().getBytes());
    return digest.map(b => ("0" + ((b < 0 ? b + 256 : b).toString(16))).slice(-2)).join("");
  } catch (e) {
    return "";
  }
}

function prepareClassificationPayload(file, allowedFolders, subFolderPath = "", ruleOptions = {}) {
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
    const rawText = file.getBlob().getDataAsString("UTF-8");
    if (rawText.trim().length > 10) {
      return buildTextPayload(file, rawText, allowedFolders, `Text (${rawText.trim().length} Zeichen)`, subFolderPath, ruleOptions);
    }
  }

  if (isWord || isExcel) {
    try {
      const extractedText = extractTextFromOfficeFile(file, isWord);
      if (extractedText && extractedText.trim().length > 15) {
        return buildTextPayload(file, extractedText, allowedFolders, `Office-Text (${extractedText.length} Zeichen)`, subFolderPath, ruleOptions);
      }
    } catch (officeErr) {}
    return buildMetadataPayload(file, allowedFolders, "Office-Fehler", subFolderPath, ruleOptions);
  }

  if (isPdf) {
    let ocrErrorNote = "";
    if (sizeInMB <= 15) {
      const ocrResult = extractTextFromPdfViaDrive(file);
      if (ocrResult.text && ocrResult.text.length >= 40) {
        return buildTextPayload(file, ocrResult.text, allowedFolders, `PDF-Text (${ocrResult.text.length} Zeichen)`, subFolderPath, ruleOptions);
      }
      ocrErrorNote = ocrResult.error;
    }

    const previewBlob = getPageOnePreviewImage(file);
    if (previewBlob) {
      const hint = ocrErrorNote ? `Seite-1-Bild (${ocrErrorNote})` : "Seite-1-Bild";
      return buildInlinePayload(file, allowedFolders, hint, previewBlob, true, subFolderPath, ruleOptions);
    }

    return buildMetadataPayload(file, allowedFolders, ocrErrorNote || "PDF ohne Vorschau", subFolderPath, ruleOptions);
  }

  if (isImage) {
    const resizedBlob = (file.getSize() > 200 * 1024) ? getPageOnePreviewImage(file) : null;
    return buildInlinePayload(file, allowedFolders, resizedBlob ? "Bild (komprimiert)" : "Bild", resizedBlob, true, subFolderPath, ruleOptions);
  }

  return buildMetadataPayload(file, allowedFolders, "Unbekanntes Format", subFolderPath, ruleOptions);
}

function buildTextPayload(file, extractedText, allowedFolders, modeTag, subFolderPath = "", ruleOptions = {}) {
  const folderSearchHint = subFolderPath ? `${subFolderPath} ${file.getName()}` : file.getName();
  const filteredFolders = selectRelevantFolders(
    allowedFolders,
    folderSearchHint,
    extractedText,
    ruleOptions.boostedFolders || [],
    Boolean(ruleOptions.isInvoiceCandidate)
  );
  return {
    modeTag: `${modeTag} | ${filteredFolders.length}/${allowedFolders.length} Ordner`,
    payload: {
      contents: [{
        parts: [{
          text: buildPrompt(file, filteredFolders, subFolderPath, ruleOptions.promptHint || "") + "\n\nINHALT:\n" + extractedText
        }]
      }],
      generationConfig: getJsonConfig(false)
    }
  };
}

function buildInlinePayload(file, allowedFolders, modeTag, customBlob, useLowMediaRes, subFolderPath = "", ruleOptions = {}) {
  const folderSearchHint = subFolderPath ? `${subFolderPath} ${file.getName()}` : file.getName();
  const filteredFolders = selectRelevantFolders(
    allowedFolders,
    folderSearchHint,
    "",
    ruleOptions.boostedFolders || [],
    Boolean(ruleOptions.isInvoiceCandidate)
  );
  const fileBlob = customBlob || file.getBlob();
  return {
    modeTag: `${modeTag} | ${filteredFolders.length}/${allowedFolders.length} Ordner`,
    payload: {
      contents: [{
        parts: [
          { text: buildPrompt(file, filteredFolders, subFolderPath, ruleOptions.promptHint || "") },
          { inline_data: { mime_type: fileBlob.getContentType(), data: Utilities.base64Encode(fileBlob.getBytes()) } }
        ]
      }],
      generationConfig: getJsonConfig(useLowMediaRes)
    }
  };
}

function buildMetadataPayload(file, allowedFolders, reason, subFolderPath = "", ruleOptions = {}) {
  const folderSearchHint = subFolderPath ? `${subFolderPath} ${file.getName()}` : file.getName();
  const filteredFolders = selectRelevantFolders(
    allowedFolders,
    folderSearchHint,
    "",
    ruleOptions.boostedFolders || [],
    Boolean(ruleOptions.isInvoiceCandidate)
  );
  const fallbackDate = Utilities.formatDate(file.getDateCreated(), "Europe/Berlin", "yyyy-MM-dd");
  const compactList = formatIndexedFolders(filteredFolders);
  const folderContext = subFolderPath ? `, Inbox-Unterordner: "${subFolderPath}"` : "";
  const ruleLine = ruleOptions.promptHint ? `\n${ruleOptions.promptHint}` : "";
  const prompt = `Archivierung nur per Dateiname ("${reason}"). Datei: "${file.getName()}" (Datum: ${fallbackDate}${folderContext}).${ruleLine}
Ordner-Index:
${compactList}
Regeln:
- is_rechnung: true NUR wenn der Dateiname eindeutig auf eine Rechnung/Quittung/Gutschrift hinweist (Dokumente mit "Abrechnung" sind KEINE Rechnungen, sofern nicht explizit als Rechnung/Quittung/Gutschrift ausgewiesen), sonst false.
- is_kontoauszug: true NUR wenn es sich um einen Bank-Kontoauszug, Finanzstatus oder eine Umsatzübersicht handelt, sonst false.
- f: Passende Ordner-Nummer (Integer) aus Liste, oder -1 falls unklar. WICHTIG: Ordner unter "01_Finanzen" dürfen NUR bei is_rechnung = true ODER is_kontoauszug = true gewählt werden! Dokumente mit "Abrechnung" gehören in einen passenden Ordner AUSSERHALB von "01_Finanzen".
- alt_f: Beste alternative Ordner-Nummer außerhalb von "01_Finanzen" (oder -1).
- n: Dateiname nach Schema "${NAMING_CONVENTION}".
- r: Deutsche Begründung (max. 8 Wörter).`;

  return {
    modeTag: `Nur-Dateiname: ${reason} | ${filteredFolders.length}/${allowedFolders.length} Ordner`,
    payload: { contents: [{ parts: [{ text: prompt }] }], generationConfig: getJsonConfig(false) }
  };
}

function createTempConvertedDriveFile(blob, tempName, targetMimeType, useOcr = false) {
  if (typeof Drive === "undefined" || !Drive.Files) return null;
  if (typeof Drive.Files.create === "function") {
    const options = useOcr ? { ocrLanguage: "de" } : {};
    return Drive.Files.create({ name: tempName, mimeType: targetMimeType }, blob, options).id;
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
    const cleanText = rawText.replace(/\s+/g, " ").trim();
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
    return extractedText.replace(/\s+/g, " ").trim();
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

function ensureFileExtension(originalName, newName) {
  let sanitized = newName.replace(/[\/\\:*?"<>|]/g, "-").trim();
  const lastDotIndex = originalName.lastIndexOf(".");
  if (lastDotIndex > 0 && (originalName.length - lastDotIndex) <= 6) {
    const ext = originalName.substring(lastDotIndex).toLowerCase();
    sanitized = sanitized.replace(/\.(pdf|jpe?g|png|webp|heic|txt|csv|docx?|xlsx?)$/i, "");
    return sanitized + ext;
  }
  return sanitized;
}