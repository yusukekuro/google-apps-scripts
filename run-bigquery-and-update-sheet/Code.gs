let spreadsheet;

function doGet(e) {
  const spreadsheetId = e.parameter.spreadsheetId; // Get the spreadsheetId from the URL parameters
  const resultMessage = runBigQueryAndUpdateSheet_(spreadsheetId);
  return ContentService.createTextOutput(resultMessage);
}

function setupDailyTrigger() {
  ScriptApp.newTrigger('processMultipleSpreadsheets')
      .timeBased()
      .everyDays(1)
      .atHour(8)
      .create();
}

function processMultipleSpreadsheets() {
  // Replace this value with the spreadsheet ID that can be found in the URL
  runBigQueryAndUpdateSheet_('XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX-XXXXXXXX')
  runBigQueryAndUpdateSheet_('XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX-XXXXXXXX')
  runBigQueryAndUpdateSheet_('XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX-XXXXXXXX')
}

function runBigQueryAndUpdateSheet_(spreadsheetId) {
  const startTime = new Date();

  spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sql = getSql_()
  const queryResult = runQuery_(sql)
  clearResultSheet_()
  appendHeader_(queryResult.header)
  appendRows_(queryResult.rows)

  const endTime = new Date();
  const executionTime = (endTime - startTime) / 1000;
  const executionTimeString = executionTime.toFixed(3);
  const resultMessage = `Spreadsheet ID: ${spreadsheetId}, Spreadsheet Name: ${spreadsheet.getName()}, Execution Time: ${executionTimeString} s.`;
  Logger.log(resultMessage);
  return resultMessage;
}

function getSql_() {
  const sqlSheet = getSheetByName_('SQL');
  return sqlSheet.getRange(1, 1).getValue();
}

// https://developers.google.com/apps-script/advanced/bigquery
function runQuery_(sql) {
  // Replace this value with the project ID listed in the Google Cloud Platform project.
  const projectId = 'XXXXXXXX';
  // Replace this value with the location to run the job
  // https://cloud.google.com/bigquery/docs/reference/rest/v2/jobs/getQueryResults#query-parameters
  const location = 'XXXXXXXX'

  const request = {
    query: sql,
    useLegacySql: false,
    location: location
  };
  let queryResults = BigQuery.Jobs.query(request, projectId);
  const jobId = queryResults.jobReference.jobId;

  // Check on status of the Query Job.
  let sleepTimeMs = 500;
  while (!queryResults.jobComplete) {
    Utilities.sleep(sleepTimeMs);
    sleepTimeMs *= 2;
    queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId, {
      location: location
    });
  }

  // Get all the rows of results.
  let rows = queryResults.rows;
  while (queryResults.pageToken) {
    queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId, {
      pageToken: queryResults.pageToken,
      location: location
    });
    rows = rows.concat(queryResults.rows);
  }

  const header = queryResults.schema.fields.map(function(field) {
    return field.name;
  });

  return {
    header: header,
    rows: rows
  }
}

function clearResultSheet_() {
  const resultSheet = getSheetByName_('Result');
  resultSheet.clear({ contentsOnly: true });
}

function appendHeader_(header) {
  const resultSheet = getSheetByName_('Result');
  resultSheet.appendRow(header);
}

function appendRows_(rows) {
  const resultSheet = getSheetByName_('Result');

  const data = new Array(rows.length);
  for (let i = 0; i < rows.length; i++) {
    const cols = rows[i].f;
    data[i] = new Array(cols.length);
    for (let j = 0; j < cols.length; j++) {
      data[i][j] = cols[j].v;
    }
  }
  resultSheet.getRange(2, 1, rows.length, rows[0].f.length).setValues(data);
}

function getSheetByName_(sheetName) {
  return spreadsheet.getSheetByName(sheetName);
}
