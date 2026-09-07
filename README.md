## CS Internal Search Engine

A custom-built internal search engine using Google Apps Script (GAS) to centralize operational data, announcements, and Google Sheets for the Customer Support team.

### Features
- **Role-Based Access Control:** Validates user email domains and roles to restrict access to sensitive sheets (e.g., Senior-level only documents).
- **Multi-Source Aggregation:** Pulls data from various Google Sheets and Google Drive folders (via ID) into a single, searchable interface.
- **Auto-Sync Trigger:** Automatically syncs new files from specified Google Drive folders every hour using Time-Driven Triggers.
- **Data Masking (Multi-Masking):** Built-in logic to automatically mask sensitive customer data (like Emails and Phone Numbers) before outputting or displaying it.
- **Tailwind CSS & Glassmorphism UI:** A modern, responsive, and fast web app interface deployed via `HtmlService`.

### Code Snippets

**1. Data Masking Function:**
```javascript
function maskDataFromWeb(input) {
  if (!input) return "";
  var items = input.split(/[\n,]+/).filter(item => item.trim() !== "");
  var results = items.map(function(item) {
    var trimmed = item.trim();
    if (trimmed.includes("@") && trimmed.includes(".")) {
      var parts = trimmed.toLowerCase().split("@");
      var nameRaw = parts[0]; var domainRaw = parts[1];
      var lastDotIndex = domainRaw.lastIndexOf(".");
      var domainName = domainRaw.substring(0, lastDotIndex);
      var domainExt = domainRaw.substring(lastDotIndex + 1);
      var maskedName = nameRaw.charAt(0) + "*".repeat(Math.max(1, nameRaw.length - 1));
      var maskedDomain = domainName.charAt(0) + "*".repeat(Math.max(1, domainName.length - 1));
      return maskedName + "@" + maskedDomain + "." + domainExt;
    } else if (/^\d{8,12}$/.test(trimmed)) {
      var len = trimmed.length;
      var mask1 = trimmed.substring(0, 3) + "x".repeat(len - 5) + trimmed.slice(-2);
      var mask2 = "x".repeat(len - 2) + trimmed.slice(-2);
      var mask3 = trimmed.substring(0, 2) + "x".repeat(len - 4) + trimmed.slice(-2);
      return mask1 + " / " + mask2 + " / " + mask3;
    }
    return trimmed;
  });
  return results.join("\n");
}
