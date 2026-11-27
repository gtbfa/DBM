// ==========================================
// KONFIGURATION: GMAIL-IMPORT
// ==========================================
const GMAIL_LABEL_NAME = "KI-Ablage";
const AUTO_FETCH_GMAIL_ON_ANALYZE = false;
const SAVE_EMAIL_BODY_AS_PDF = true;
const MAX_GMAIL_THREADS_PER_RUN = 15;


// ==========================================
// GMAIL-IMPORT: ANHÄNGE & E-MAILS IN 00_INBOX LADEN
// ==========================================
function runGmailImportOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const count = fetchGmailAttachments(ss);
  updateLiveStatus(
    ss,
    `Gmail-Import beendet: ${count} Datei(en) in '00_Inbox' gespeichert.`,
    "📥 Gmail-Import fertig ✓",
    6
  );
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
        const rawMime = (att.getContentType() || "").toLowerCase();
        const mime = rawMime.split(";")[0].trim();
        const sizeBytes = att.getSize();
        const rawAttName = att.getName() || "";
        const lowerAttName = rawAttName.toLowerCase();
        const isAllowedExt = /\.(pdf|jpe?g|png|webp|heic|txt|csv|docx?|xlsx?)$/i.test(lowerAttName);

        if (!allowedMimeTypes.has(mime) && !isAllowedExt) continue;
        if (mime.startsWith("image/") && sizeBytes < 5120) continue;

        let ext = "";
        if (mime === "application/vnd.openxmlformats-officedocument.wordprocessingml.document" || lowerAttName.endsWith(".docx")) ext = ".docx";
        else if (mime === "application/msword" || lowerAttName.endsWith(".doc")) ext = ".doc";
        else if (mime === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" || lowerAttName.endsWith(".xlsx")) ext = ".xlsx";
        else if (mime === "application/vnd.ms-excel" || lowerAttName.endsWith(".xls")) ext = ".xls";
        else if (mime === "application/pdf" || lowerAttName.endsWith(".pdf")) ext = ".pdf";
        else if (mime === "image/jpeg" || lowerAttName.endsWith(".jpg") || lowerAttName.endsWith(".jpeg")) ext = ".jpg";
        else if (mime === "image/png" || lowerAttName.endsWith(".png")) ext = ".png";

        const fileName = rawAttName || `Gmail-Anhang_${Utilities.formatDate(msg.getDate(), "Europe/Berlin", "yyyy-MM-dd")}${ext}`;
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