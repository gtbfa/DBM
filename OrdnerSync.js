// ==========================================
// KONFIGURATION: ORDNER-SYNC
// ==========================================
const FOLDER_SYNC_WHITELIST_SHEET = "Ordner_Whitelist";
const FOLDER_SYNC_REVIEW_SHEET = "Prüfung & Freigabe";
const FOLDER_SYNC_DROPDOWN_COL = 7; // Spalte G in "Prüfung & Freigabe"
const FOLDER_SYNC_EXCLUDED_NAMES = ["00_Inbox"];

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

function isExcludedSyncFolder(folderId, folderName) {
  const cleanName = (folderName || "").toString().trim().toLowerCase();
  if (FOLDER_SYNC_EXCLUDED_NAMES.some(ex => cleanName === ex.toLowerCase())) {
    return true;
  }
  if (typeof INBOX_FOLDER_ID !== "undefined" && INBOX_FOLDER_ID && folderId === INBOX_FOLDER_ID) {
    return true;
  }
  return false;
}

// ==========================================
// INITIALISIERUNG: VERZEICHNIS AUSWÄHLEN ODER EINGEBEN
// ==========================================
function syncSubfolders(directPathOrId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const whitelistSheet = ss.getSheetByName(FOLDER_SYNC_WHITELIST_SHEET);

  if (!whitelistSheet) {
    ui.alert(
      "Tabellenblatt fehlt ⚠️",
      `Das Tabellenblatt "${FOLDER_SYNC_WHITELIST_SHEET}" wurde nicht gefunden.`,
      ui.ButtonSet.OK
    );
    return;
  }

  // Falls ein Pfad/Link/ID direkt übergeben wurde, ohne Dialog ausführen
  if (typeof directPathOrId === "string" && directPathOrId.trim() !== "") {
    try {
      const resolvedDir = resolveDriveDirectory(directPathOrId);
      executeFolderSync(resolvedDir);
    } catch (err) {
      ui.alert("Ungültiges Verzeichnis ⚠️", err.message || String(err), ui.ButtonSet.OK);
    }
    return;
  }

  // Hauptordner aus Google Drive laden und als nummerierte Auswahl anzeigen
  const topFolders = getTopLevelDriveFolders();
  let folderListText = "";

  if (topFolders.length > 0) {
    const visibleFolders = topFolders.slice(0, 20);
    folderListText =
      "Verfügbare Hauptordner in Google Drive:\n" +
      visibleFolders.map((f, idx) => `  ${idx + 1}: ${f.name}`).join("\n") +
      (topFolders.length > 20 ? `\n  ... (${topFolders.length - 20} weitere Ordner)` : "") +
      "\n\n";
  }

  const promptMessage =
    `${folderListText}` +
    `Gib eine Nummer (z. B. 1), einen Verzeichnispfad (z. B. 02_Finanzen / Abrechnung), einen Drive-Link oder eine Ordner-ID ein:`;

  const response = ui.prompt(
    "📁 Zielverzeichnis für Ordner-Sync",
    promptMessage,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const userInput = response.getResponseText().trim();
  if (!userInput) {
    ui.alert("Eingabe fehlt ⚠️", "Bitte wähle eine Nummer aus oder gib einen gültigen Verzeichnispfad an.", ui.ButtonSet.OK);
    return;
  }

  try {
    let resolvedDir = null;

    // Prüfung, ob eine Nummer aus der angezeigten Hauptordner-Liste gewählt wurde
    if (/^\d{1,3}$/.test(userInput)) {
      const num = parseInt(userInput, 10);
      if (num >= 1 && num <= topFolders.length) {
        resolvedDir = { id: topFolders[num - 1].id, path: topFolders[num - 1].name };
      }
    }

    if (!resolvedDir) {
      resolvedDir = resolveDriveDirectory(userInput);
    }

    executeFolderSync(resolvedDir);
  } catch (err) {
    ui.alert("Ungültiges Verzeichnis ⚠️", err.message || String(err), ui.ButtonSet.OK);
  }
}

function getTopLevelDriveFolders() {
  const folders = [];
  try {
    const iter = DriveApp.getRootFolder().getFolders();
    while (iter.hasNext()) {
      const f = iter.next();
      if (!f.isTrashed() && !isExcludedSyncFolder(f.getId(), f.getName())) {
        folders.push({ id: f.getId(), name: f.getName() });
      }
    }
    folders.sort((a, b) => a.name.localeCompare(b.name, "de"));
  } catch (e) {}
  return folders;
}

function resolveDriveDirectory(rawInput) {
  const input = (rawInput || "").toString().trim();
  if (!input) {
    throw new Error("Bitte gib ein gültiges Verzeichnis, einen Drive-Link oder eine Ordner-ID an.");
  }

  // 1. Prüfung auf Google-Drive-URL oder direkte Ordner-ID
  const urlIdMatch = input.match(/\/folders\/([-\w]{25,})/) || input.match(/[?&]id=([-\w]{25,})/);
  const directIdCandidate = urlIdMatch ? urlIdMatch[1] : (/^[-\w]{25,}$/.test(input) ? input : null);

  if (directIdCandidate) {
    try {
      const folder = DriveApp.getFolderById(directIdCandidate);
      if (folder.isTrashed()) {
        throw new Error("Der angegebene Ordner befindet sich im Papierkorb.");
      }
      if (isExcludedSyncFolder(folder.getId(), folder.getName())) {
        throw new Error(`Der Ordner "${folder.getName()}" ist ausgeschlossen und wird nicht eingelesen.`);
      }
      return { id: folder.getId(), path: folder.getName() };
    } catch (e) {
      if (urlIdMatch || (e.message && e.message.includes("ausgeschlossen"))) {
        throw e;
      }
    }
  }

  // 2. Prüfung als Verzeichnispfad (z. B. "01_Kunden" oder "02_Finanzen / Abrechnung")
  const segments = input
    .split(/[\/\\]+/)
    .map(s => s.trim())
    .filter(s => s.length > 0 && s.toLowerCase() !== "meine ablage" && s.toLowerCase() !== "my drive");

  if (segments.length === 0) {
    throw new Error("Bitte gib einen gültigen Verzeichnispfad an.");
  }

  let currentFolder = null;
  const rootIter = DriveApp.getRootFolder().getFoldersByName(segments[0]);
  while (rootIter.hasNext()) {
    const candidate = rootIter.next();
    if (!candidate.isTrashed()) {
      currentFolder = candidate;
      break;
    }
  }

  // Fallback: Falls der Startordner ein freigegebener Ordner außerhalb von "Meine Ablage" ist
  if (!currentFolder) {
    const sharedIter = DriveApp.getFoldersByName(segments[0]);
    while (sharedIter.hasNext()) {
      const candidate = sharedIter.next();
      if (!candidate.isTrashed()) {
        currentFolder = candidate;
        break;
      }
    }
  }

  if (!currentFolder) {
    throw new Error(`Das Startverzeichnis "${segments[0]}" wurde in Google Drive nicht gefunden.`);
  }

  if (isExcludedSyncFolder(currentFolder.getId(), currentFolder.getName())) {
    throw new Error(`Der Ordner "${currentFolder.getName()}" ist ausgeschlossen und wird nicht eingelesen.`);
  }

  const resolvedSegments = [currentFolder.getName()];

  for (let i = 1; i < segments.length; i++) {
    const subName = segments[i];
    const subIter = currentFolder.getFoldersByName(subName);
    let nextFolder = null;
    while (subIter.hasNext()) {
      const candidate = subIter.next();
      if (!candidate.isTrashed()) {
        nextFolder = candidate;
        break;
      }
    }
    if (!nextFolder) {
      throw new Error(`Der Unterordner "${subName}" wurde unter "${resolvedSegments.join(" / ")}" nicht gefunden.`);
    }
    if (isExcludedSyncFolder(nextFolder.getId(), nextFolder.getName())) {
      throw new Error(`Der Ordner "${nextFolder.getName()}" ist ausgeschlossen und wird nicht eingelesen.`);
    }
    currentFolder = nextFolder;
    resolvedSegments.push(currentFolder.getName());
  }

  return {
    id: currentFolder.getId(),
    path: resolvedSegments.join(" / ")
  };
}

// ==========================================
// HAUPTFUNKTION: NUR END-UNTERORDNER EINLESEN
// ==========================================
function executeFolderSync(rootDir) {
  const startTime = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const whitelistSheet = ss.getSheetByName(FOLDER_SYNC_WHITELIST_SHEET);

  if (!whitelistSheet) {
    throw new Error(`Das Tabellenblatt "${FOLDER_SYNC_WHITELIST_SHEET}" wurde nicht gefunden.`);
  }

  notifySyncStatus(ss, `Lese Ziel-Unterordner unter "${rootDir.path}" ein...`, "📁 Ordner-Sync", 5);

  // Bestehende Kopfzeile in Zeile 1 beibehalten, falls vorhanden
  let hasHeader = false;
  let headerRow = [];
  if (whitelistSheet.getLastRow() >= 1) {
    const firstRow = whitelistSheet.getRange(1, 1, 1, 2).getValues()[0];
    const colA = firstRow[0] ? firstRow[0].toString().trim() : "";
    const colB = firstRow[1] ? firstRow[1].toString().trim() : "";
    if (colA && colB && !/[-\w]{25,}/.test(colB)) {
      hasHeader = true;
      headerRow = [colA, colB];
    }
  }

  const folderRegistry = new Map();
  let lastDriveError = "";

  // Schnellmodus: Alle Drive-Ordner mit einer Abfrage über die Drive API laden
  const childrenMap = fetchAllDriveFoldersMap();

  if (childrenMap) {
    try {
      const rootFolder = DriveApp.getFolderById(rootDir.id);
      if (!rootFolder.isTrashed()) {
        // Hauptordner selbst und ausgeschlossene Ordner (z. B. 00_Inbox) überspringen
        const directSubfolders = (childrenMap.get(rootDir.id) || [])
          .filter(sub => !isExcludedSyncFolder(sub.id, sub.name));
        directSubfolders.sort((a, b) => a.name.localeCompare(b.name, "de"));
        directSubfolders.forEach(sub => {
          traverseFolderMapInMemory(sub.id, sub.name, childrenMap, folderRegistry);
        });
      }
    } catch (e) {
      lastDriveError = e.message || e.toString();
    }
  } else {
    // Fallback über DriveApp, falls erweiterter Drive-Dienst nicht aktiv ist
    try {
      const rootFolder = DriveApp.getFolderById(rootDir.id);
      if (!rootFolder.isTrashed()) {
        const subIter = rootFolder.getFolders();
        const directSubs = [];
        while (subIter.hasNext()) {
          const sub = subIter.next();
          if (!sub.isTrashed() && !isExcludedSyncFolder(sub.getId(), sub.getName())) {
            directSubs.push(sub);
          }
        }
        directSubs.sort((a, b) => a.getName().localeCompare(b.getName(), "de"));
        directSubs.forEach(sub => {
          collectFoldersRecursively(sub, sub.getName(), folderRegistry);
        });
      }
    } catch (e) {
      lastDriveError = e.message || e.toString();
    }
  }

  if (folderRegistry.size === 0) {
    const activeUser = Session.getActiveUser().getEmail() || "Unbekanntes Konto";
    const reason = lastDriveError
      ? `Google Drive meldet: "${lastDriveError}" (Aktives Konto: ${activeUser})`
      : `Unter "${rootDir.path}" wurden keine gültigen Unterordner gefunden.`;
    throw new Error(reason);
  }

  const outputRows = [];
  if (hasHeader) outputRows.push(headerRow);
  folderRegistry.forEach((path, id) => outputRows.push([path, id]));

  whitelistSheet.clearContents();
  whitelistSheet.getRange(1, 1, outputRows.length, 2).setValues(outputRows);
  applyFolderDropdownValidation(ss);

  const totalFolders = hasHeader ? outputRows.length - 1 : outputRows.length;
  const elapsedSec = ((Date.now() - startTime) / 1000).toFixed(1);
  notifySyncStatus(ss, `${totalFolders} Zielordner aus "${rootDir.path}" in ${elapsedSec}s eingelesen ✓`, "📁 Synchronisation fertig ✓", 5);
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
  const currentName = currentPath.split(" / ").pop();
  if (isExcludedSyncFolder(folderId, currentName)) return;

  const kids = (childrenMap.get(folderId) || [])
    .filter(kid => !isExcludedSyncFolder(kid.id, kid.name));

  // Wenn Unterordner vorhanden sind, diesen Oberordner ignorieren und nur weiter absteigen
  if (kids.length > 0) {
    kids.sort((a, b) => a.name.localeCompare(b.name, "de"));
    kids.forEach(kid => {
      traverseFolderMapInMemory(kid.id, `${currentPath} / ${kid.name}`, childrenMap, registry);
    });
    return;
  }

  // Nur Ordner ohne weitere Unterordner (Endordner) in die Whitelist eintragen
  if (!registry.has(folderId)) {
    registry.set(folderId, currentPath);
  }
}

function collectFoldersRecursively(folder, currentPath, registry) {
  if (isExcludedSyncFolder(folder.getId(), folder.getName())) return;

  const subIter = folder.getFolders();
  const validSubs = [];
  while (subIter.hasNext()) {
    const sub = subIter.next();
    if (!sub.isTrashed() && !isExcludedSyncFolder(sub.getId(), sub.getName())) {
      validSubs.push(sub);
    }
  }

  // Wenn Unterordner vorhanden sind, diesen Oberordner ignorieren und nur weiter absteigen
  if (validSubs.length > 0) {
    validSubs.sort((a, b) => a.getName().localeCompare(b.getName(), "de"));
    validSubs.forEach(sub => {
      collectFoldersRecursively(sub, `${currentPath} / ${sub.getName()}`, registry);
    });
    return;
  }

  // Nur Ordner ohne weitere Unterordner (Endordner) in die Whitelist eintragen
  if (!registry.has(folder.getId())) {
    registry.set(folder.getId(), currentPath);
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