// Lädt pdf-lib einmalig gecacht in den Apps Script Speicher
function getPdfLib() {
  const cdnUrl = "https://cdn.jsdelivr.net/npm/pdf-lib@1.17.9/dist/pdf-lib.min.js";
  const code = UrlFetchApp.fetch(cdnUrl).getContentText();
  eval(code);
  return PDFLib;
}