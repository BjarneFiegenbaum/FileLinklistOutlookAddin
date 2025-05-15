
Office.onReady(() => {
  // Ensure Office.js is ready
});

function mapDriveToUNC(path) {
  if (path.startsWith("S:")) {
    return path.replace("S:", "\\MDZ-FS1\Zone 1");
  } else if (path.startsWith("T:")) {
    return path.replace("T:", "\\MDZ-FS1\Zone 2");
  } else if (path.startsWith("U:")) {
    return path.replace("U:", "\\MDZ-FS1\Zone 3");
  }
  return path;
}

function extractFileName(path) {
  const parts = path.split("\\");
  return parts[parts.length - 1];
}

async function insertFileLinks() {
  await Office.context.mailbox.item.body.getAsync("text", async function(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const body = result.value;
      const pathRegex = /([STU]:\\[^\s<>"']+\.(docx|xlsx|pdf|pptx|txt|csv))/gi;
      const matches = [...body.matchAll(pathRegex)];

      if (matches.length === 0) {
        Office.context.mailbox.item.notificationMessages.addAsync("noMatches", {
          type: "informationalMessage",
          message: "Keine Dateipfade gefunden.",
          icon: "Icon.16x16",
          persistent: false
        });
        return;
      }

      let html = "<p><strong>📎 Verlinkte Dateien:</strong><ul>";

      matches.forEach(match => {
        const original = match[1];
        const unc = mapDriveToUNC(original).replace(/ /g, "%20");
        const name = extractFileName(original);
        html += `<li><a href="file://${unc}">${name}</a></li>`;
      });

      html += "</ul></p>";

      Office.context.mailbox.item.body.setSelectedDataAsync(html, {
        coercionType: Office.CoercionType.Html,
        asyncContext: null
      });
    }
  });
}

window.insertFileLinks = insertFileLinks;
