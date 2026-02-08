
// 1. Open your Google Sheet
// 2. Go to Extensions > Apps Script
// 3. Delete any code in the editor and paste this code
// 4. Click 'Deploy' > 'New deployment'
// 5. Select type 'Web app'
// 6. Set Description to 'Contact Form'
// 7. Set 'Execute as' to 'Me'
// 8. Set 'Who has access' to 'Anyone' (IMPORTANT)
// 9. Click 'Deploy', then 'Authorize access' if prompted
// 10. Copy the 'Web app URL' and paste it into the script in index.html

function doPost(e) {
    try {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
        if (!sheet) {
            // Fallback if Sheet1 is not found, though user requested Sheet1.
            sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        }
        var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        var nextRow = sheet.getLastRow() + 1;

        // Parse the data
        var params = e.parameter;
        var newRow = [];
        var timestamp = new Date();

        // If headers are missing, set them up
        if (headers.length === 0 || headers[0] === "") {
            headers = ["Timestamp", "Name", "Email", "Service", "Message"];
            sheet.appendRow(headers);
        }

        // Map existing headers to values
        for (var i = 0; i < headers.length; i++) {
            var header = headers[i];
            var value = "";

            switch (header) {
                case "Timestamp":
                    value = timestamp;
                    break;
                case "Name":
                    value = params.Name;
                    break;
                case "Email":
                    value = params.Email;
                    break;
                case "Service":
                    value = params.Service;
                    break;
                case "Message":
                    value = params["Text Area"]; // "Text Area" is the name in the form
                    break;
                default:
                    value = params[header] || "";
            }
            newRow.push(value);
        }

        sheet.appendRow(newRow);

        return ContentService.createTextOutput(JSON.stringify({ "result": "success" }))
            .setMimeType(ContentService.MimeType.JSON);

    } catch (error) {
        return ContentService.createTextOutput(JSON.stringify({ "result": "error", "error": error }))
            .setMimeType(ContentService.MimeType.JSON);
    }
}
