# Run BigQuery and Update Google Sheet

This Google Apps Script project helps you to automatically run a BigQuery SQL and update a Google Sheet with the result.
The script can also be set up to perform this operation daily, at a particular time, for multiple spreadsheets.

## Getting Started

To get started with this script, you need to:

1. Create a new Google Apps Script project in the Google Developer Console.
2. Copy and paste the content of the `Code.gs` file into your new Apps Script project.
3. Add the BigQuery API v2 to the Apps Script project.

## Usage

The script functions in the following way:

1. The `doGet(e)` function accepts HTTP GET requests with a `spreadsheetId` URL parameter to specify the Google
   Spreadsheet that the BigQuery results should be written to.
2. The `setupDailyTrigger()` function sets up a daily trigger to run the `processMultipleSpreadsheets()` function.
3. The `processMultipleSpreadsheets()` function runs the BigQuery SQL for multiple spreadsheets.
4. The `runBigQueryAndUpdateSheet_(spreadsheetId)` function runs the BigQuery SQL and updates the specified Google
   Sheet.

You need to replace all `'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX-XXXXXXXX'` placeholders with actual Google Spreadsheet
IDs. Additionally, replace the `'XXXXXXXX'` placeholders with your Google Cloud Platform project ID and BigQuery
location.

### Using Bookmarklet

One of the most convenient ways to use this script is to trigger it from a bookmarklet. This JavaScript snippet can be
saved as a bookmark in your browser. When clicked, it retrieves the ID of the currently open Google Spreadsheet and
triggers the web app with this ID as a parameter.

Here's the bookmarklet code:

```javascript
javascript: (function () {
    var spreadsheetId = null;
    var matches = window.location.href.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (matches) {
        spreadsheetId = matches[1];
    }
    if (spreadsheetId) {
        var webAppUrl = "https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec";
        var url = webAppUrl + "?spreadsheetId=" + spreadsheetId;
        window.open(url, "_blank");
    } else {
        alert("Unable to retrieve the Spreadsheet ID.");
    }
})();
```

In the `webAppUrl` variable, replace `YOUR_SCRIPT_ID` with the actual ID of your web app. You can find this ID in the "
Publish" > "Deploy as web app" dialog of the Apps Script editor. After deploying the project as a web app, the URL of
the latest version will be displayed. The ID is the long string of characters in the URL.

This ID is unique to your project and keeps your script secure, so don't share it publicly.

## Configuration

This script requires you to have two sheets in your Google Spreadsheet:

1. 'SQL': This sheet should contain your BigQuery SQL statement in cell A1.
2. 'Result': This is where the results of the BigQuery will be written to.
