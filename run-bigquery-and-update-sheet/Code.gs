function runBigQueryAndUpdateSheet() {
  const sql = getSql_()
  const queryResult = runQuery_(sql)
  clearResultSheet_()
  appendHeader_(queryResult.header)
  appendRows_(queryResult.rows)
}

function getSql_() {
  const sqlSheet = getSheetByName_('SQL');
  return sqlSheet.getRange(1, 1).getValue();
}

// https://developers.google.com/apps-script/advanced/bigquery
function runQuery_(sql) {
  // Replace this value with the project ID listed in the Google
  // Cloud Platform project.
  const projectId = 'XXXXXXXX';

  const request = {
    query: sql,
    useLegacySql: false
  };
  let queryResults = BigQuery.Jobs.query(request, projectId);
  const jobId = queryResults.jobReference.jobId;

  // Check on status of the Query Job.
  let sleepTimeMs = 500;
  while (!queryResults.jobComplete) {
    Utilities.sleep(sleepTimeMs);
    sleepTimeMs *= 2;
    queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId);
  }

  // Get all the rows of results.
  let rows = queryResults.rows;
  while (queryResults.pageToken) {
    queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId, {
      pageToken: queryResults.pageToken
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
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet.getSheetByName(sheetName);
}
