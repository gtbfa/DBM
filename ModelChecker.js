// ==========================================
// EIGENES MENÜ FÜR DIE MODELL-ÜBERSICHT
// ==========================================
function createModelMenu() {
  SpreadsheetApp.getUi()
    .createMenu("🔍 Gemini-Modelle")
    .addItem("Verfügbare Modelle anzeigen", "showAvailableModels")
    .addItem("Modell-Cache zurücksetzen", "clearModelCache")
    .addToUi();
}


// ==========================================
// VERFÜGBARE MODELLE ALS POPUP ANZEIGEN
// ==========================================
function showAvailableModels() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty("GEMINI_API_KEY");

  if (!apiKey || apiKey.trim() === "") {
    ui.alert("Fehlender API-Schlüssel", "Bitte hinterlegen Sie zuerst 'GEMINI_API_KEY' in den Projekteinstellungen unter Skripteigenschaften.", ui.ButtonSet.OK);
    return;
  }

  const url = "https://generativelanguage.googleapis.com/v1beta/models?key=" + apiKey.trim();
  const res = UrlFetchApp.fetch(url, { method: "get", muteHttpExceptions: true });

  if (res.getResponseCode() !== 200) {
    ui.alert("API-Fehler (" + res.getResponseCode() + ")", res.getContentText(), ui.ButtonSet.OK);
    return;
  }

  const data = JSON.parse(res.getContentText());
  const cachedRaw = props.getProperty("CACHED_GEMINI_MODELS");
  let activeCachedModels = [];
  if (cachedRaw) {
    try { activeCachedModels = JSON.parse(cachedRaw); } catch (e) {}
  }

  const models = (data.models || []).filter(m => {
    const methods = m.supportedGenerationMethods || [];
    const name = (m.name || "").toLowerCase();
    return methods.includes("generateContent") && name.includes("gemini");
  });

  // Sortierung: Aktuelle Versionen nach oben
  models.sort((a, b) => b.name.localeCompare(a.name));

  let rowsHtml = "";
  models.forEach(m => {
    const cleanId = m.name.replace("models/", "");
    const displayName = m.displayName || cleanId;
    const tokenLimit = m.inputTokenLimit ? m.inputTokenLimit.toLocaleString("de-DE") : "-";
    const isCached = activeCachedModels.includes(m.name);
    const badge = isCached ? `<span style="background:#d9ead3;color:#274e13;padding:2px 6px;border-radius:4px;font-size:11px;font-weight:bold;">Aktiv (Auto)</span>` : "";

    rowsHtml += `
      <tr style="border-bottom: 1px solid #eee;">
        <td style="padding: 8px; font-family: monospace; font-size: 13px;">
          <strong>${cleanId}</strong> ${badge}
        </td>
        <td style="padding: 8px; font-size: 13px;">${displayName}</td>
        <td style="padding: 8px; font-size: 12px; color: #555; text-align: right;">${tokenLimit}</td>
        <td style="padding: 8px; text-align: center;">
          <button onclick="copyId('${cleanId}', this)" style="cursor:pointer;padding:4px 8px;border:1px solid #ccc;border-radius:4px;background:#f8f9fa;">Kopieren</button>
        </td>
      </tr>
    `;
  });

  const htmlContent = `
    <div style="font-family: Arial, sans-serif; color: #222;">
      <p style="font-size: 13px; color: #555; margin-top: 0;">
        Für Ihren API-Schlüssel freigeschaltete Text-/Multimodal-Modelle (${models.length} gefunden):
      </p>
      <div style="max-height: 360px; overflow-y: auto; border: 1px solid #ddd; border-radius: 6px;">
        <table style="width: 100%; border-collapse: collapse; text-align: left;">
          <thead style="background: #f3f3f3; position: sticky; top: 0;">
            <tr>
              <th style="padding: 8px; font-size: 12px;">Modell-ID (für Code)</th>
              <th style="padding: 8px; font-size: 12px;">Bezeichnung</th>
              <th style="padding: 8px; font-size: 12px; text-align: right;">Input-Limit</th>
              <th style="padding: 8px; font-size: 12px; text-align: center;">Aktion</th>
            </tr>
          </thead>
          <tbody>${rowsHtml}</tbody>
        </table>
      </div>
      <script>
        function copyId(text, btn) {
          navigator.clipboard.writeText(text);
          const orig = btn.innerText;
          btn.innerText = "Kopiert ✓";
          btn.style.background = "#d9ead3";
          setTimeout(() => { btn.innerText = orig; btn.style.background = "#f8f9fa"; }, 1500);
        }
      </script>
    </div>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(720)
    .setHeight(450);

  ui.showModalDialog(htmlOutput, "Verfügbare Gemini-Modelle");
}


// ==========================================
// ZWISCHENSPEICHER FÜR AUTO-MODELLE LÖSCHEN
// ==========================================
function clearModelCache() {
  PropertiesService.getScriptProperties().deleteProperty("CACHED_GEMINI_MODELS");
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Der Modell-Cache wurde geleert. Beim nächsten Analyselauf wird die Modellliste frisch von Google abgerufen.",
    "Cache zurückgesetzt ✓",
    5
  );
}