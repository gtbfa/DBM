// ==========================================
// KONFIGURATION: ORDNER-SYNC
// ==========================================
const FOLDER_SYNC_WHITELIST_SHEET = "Ordner_Whitelist";
const FOLDER_SYNC_REVIEW_SHEET = "Prüfung & Freigabe";
const FOLDER_SYNC_DROPDOWN_COL = 7; // Spalte G in "Prüfung & Freigabe"

// Optionales eigenes Menü (nur aktivieren, falls diese Datei OHNE code.gs genutzt wird!)
/*
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📁 Ordner-Sync")
    .addItem("0. Ordner & Unterordner einlesen", "syncSubfolders")
    .addToUi();
}
*/

function notifySyncStatus(ss, message, title = "📁 Ordner-Sync", seconds = 5) {
  if (!ss) return;
  if (typeof updateLiveStatus === "function") {
    updateLiveStatus(ss, message, title, seconds);
  } else {
    ss.toast(message, title, seconds);
  }
}

// ==========================================
// HAUPTFUNKTION: ORDNER & UNTERORDNER EINLESEN
// ==========================================
function syncSubfolders() {
  const startTime = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const whitelistSheet = ss.getSheetByName(FOLDER_SYNC_WHITELIST_SHEET);

  if (!whitelistSheet) {
    SpreadsheetApp.getUi().alert(
      "Tabellenblatt fehlt ⚠️",
      `Das Tabellenblatt "${FOLDER_SYNC_WHITELIST_SHEET}" wurde nicht gefunden.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const lastRow = whitelistSheet.getLastRow();
  if (lastRow === 0) {
    notifySyncStatus(ss, `Bitte trage zuerst deine Haupt-Zielordner in '${FOLDER_SYNC_WHITELIST_SHEET}' ein.`, "Hinweis", 5);
    return;
  }

  notifySyncStatus(ss, "Lese Ordnerstruktur aus Google Drive ein...", "📁 Ordner-Sync", 5);

  const existingData = whitelistSheet.getRange(1, 1, lastRow, 2).getValues();
  const rootSeeds = [];
  let hasHeader = false;
  let headerRow = [];

  existingData.forEach((row, idx) => {
    const name = row[0] ? row[0].toString().trim() : "";
    const rawId = row[1] ? row[1].toString().trim() : "";
    if (!rawId) return;

    const idMatch = rawId.match(/[-\w]{25,}/);
    if (!idMatch) {
      if (idx === 0) {
        hasHeader = true;
        headerRow = [name, rawId];
      }
      return;
    }
    const id = idMatch[0];
    const depth = name.split(" / ").length;
    rootSeeds.push({ id, name, depth });
  });

  // Hauptordner zuerst verarbeiten
  rootSeeds.sort((a, b) => a.depth - b.depth);

  const folderRegistry = new Map();
  let lastDriveError = "";

  // Schnellmodus: Alle Drive-Ordner mit einer Abfrage über die Drive API laden
  const childrenMap = fetchAllDriveFoldersMap();

  if (childrenMap) {
    rootSeeds.forEach(seed => {
      if (folderRegistry.has(seed.id)) return;
      try {
        const rootFolder = DriveApp.getFolderById(seed.id);
        if (!rootFolder.isTrashed()) {
          const basePath = seed.name !== "" ? seed.name : rootFolder.getName();
          traverseFolderMapInMemory(seed.id, basePath, childrenMap, folderRegistry);
        }
      } catch (e) {
        lastDriveError = e.message || e.toString();
      }
    });
  } else {
    // Fallback über DriveApp, falls erweiterter Drive-Dienst nicht aktiv ist
    rootSeeds.forEach(seed => {
      if (folderRegistry.has(seed.id)) return;
      try {
        const folder = DriveApp.getFolderById(seed.id);
        if (!folder.isTrashed()) {
          const basePath = seed.name !== "" ? seed.name : folder.getName();
          collectFoldersRecursively(folder, basePath, folderRegistry);
        }
      } catch (e) {
        lastDriveError = e.message || e.toString();
      }
    });
  }

  if (folderRegistry.size === 0) {
    const activeUser = Session.getActiveUser().getEmail() || "Unbekanntes Konto";
    const reason = lastDriveError
      ? `Google Drive meldet: "${lastDriveError}" (Aktives Konto: ${activeUser})`
      : `In '${FOLDER_SYNC_WHITELIST_SHEET}' stehen aktuell keine gültigen Ordner-IDs.`;
    SpreadsheetApp.getUi().alert("0 Ordner eingelesen ⚠️", reason, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const outputRows = [];
  if (hasHeader) outputRows.push(headerRow);
  folderRegistry.forEach((path, id) => outputRows.push([path, id]));

  whitelistSheet.clearContents();
  whitelistSheet.getRange(1, 1, outputRows.length, 2).setValues(outputRows);
  applyFolderDropdownValidation(ss);

  const totalFolders = hasHeader ? outputRows.length - 1 : outputRows.length;
  const elapsedSec = ((Date.now() - startTime) / 1000).toFixed(1);
  notifySyncStatus(ss, `${totalFolders} Ordner in ${elapsedSec}s eingelesen ✓`, "📁 Synchronisation fertig ✓", 5);
}

// ==========================================
// HILFSFUNKTIONEN (DRIVE API, REKURSION, DROPDOWN)
// ==========================================
function fetchAllDriveFoldersMap() {
  if (typeof Drive === "undefined" || !Drive.Files || typeof Drive.Files.list !== "function") return null;
  try {
    const childrenMap = new Map();
    let pageToken = null;
    do {
      const params = {
        q: "mimeType = 'application/vnd.google-apps.folder' and trashed = false",
        pageSize: 1000,
        fields: "nextPageToken, files(id, name, parents)"
      };
      if (pageToken) params.pageToken = pageToken;

      let res;
      try {
        res = Drive.Files.list(params);
      } catch (v2Err) {
        const v2Params = {
          q: "mimeType = 'application/vnd.google-apps.folder' and trashed = false",
          maxResults: 1000,
          fields: "nextPageToken, items(id, title, parents(id))"
        };
        if (pageToken) v2Params.pageToken = pageToken;
        res = Drive.Files.list(v2Params);
      }

      const items = res.files || res.items || [];
      items.forEach(item => {
        const fName = item.name || item.title || "";
        const parents = item.parents || [];
        parents.forEach(p => {
          const pId = typeof p === "string" ? p : p.id;
          if (!pId) return;
          if (!childrenMap.has(pId)) childrenMap.set(pId, []);
          childrenMap.get(pId).push({ id: item.id, name: fName });
        });
      });

      pageToken = res.nextPageToken;
    } while (pageToken);

    return childrenMap;
  } catch (e) {
    return null;
  }
}

function traverseFolderMapInMemory(folderId, currentPath, childrenMap, registry) {
  if (!registry.has(folderId)) {
    registry.set(folderId, currentPath);
  }
  const kids = childrenMap.get(folderId) || [];
  kids.sort((a, b) => a.name.localeCompare(b.name));
  kids.forEach(kid => {
    traverseFolderMapInMemory(kid.id, `${currentPath} / ${kid.name}`, childrenMap, registry);
  });
}

function collectFoldersRecursively(folder, currentPath, registry) {
  if (!registry.has(folder.getId())) {
    registry.set(folder.getId(), currentPath);
  }
  const subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    const sub = subFolders.next();
    if (!sub.isTrashed()) {
      collectFoldersRecursively(sub, `${currentPath} / ${sub.getName()}`, registry);
    }
  }
}

function applyFolderDropdownValidation(ss) {
  const reviewSheet = ss.getSheetByName(FOLDER_SYNC_REVIEW_SHEET);
  const whitelistSheet = ss.getSheetByName(FOLDER_SYNC_WHITELIST_SHEET);
  if (!reviewSheet || !whitelistSheet) return;

  const lastWhitelistRow = whitelistSheet.getLastRow();
  if (lastWhitelistRow === 0) return;

  const data = whitelistSheet.getRange(1, 1, lastWhitelistRow, 2).getValues();
  let firstValidRow = -1;
  let lastValidRow = -1;

  for (let i = 0; i < data.length; i++) {
    const name = data[i][0] ? data[i][0].toString().trim() : "";
    const id = data[i][1] ? data[i][1].toString().trim() : "";
    if (name && id && id.length > 10) {
      if (firstValidRow === -1) firstValidRow = i + 1;
      lastValidRow = i + 1;
    }
  }
  if (firstValidRow === -1) return;

  const sourceRange = whitelistSheet.getRange(firstValidRow, 1, lastValidRow - firstValidRow + 1, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sourceRange, true)
    .setAllowInvalid(false)
    .setHelpText("Bitte wählen Sie einen gültigen Zielordner aus der Whitelist.")
    .build();

  const maxTargetRows = Math.max(reviewSheet.getMaxRows() - 1, 50);
  reviewSheet.getRange(2, FOLDER_SYNC_DROPDOWN_COL, maxTargetRows, 1).setDataValidation(rule);
}