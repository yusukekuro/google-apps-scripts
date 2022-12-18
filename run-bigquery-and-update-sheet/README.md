# run-bigquery-and-update-sheet

This script runs specified SQL in BigQuery and update spread sheet with the result of the SQL

## Preparation

1. Create a google spread sheet
2. In the spread sheet, create 2 sheets "Result" and "SQL"
3. Write SQL in the top left cell of "SQL" sheet
    - This SQL will be passed to the BiqQuery API
4. From the google spread sheet, create apps script
5. Paste the content of Code.gs to the apps script project
    - Change `projectId` in the script
6. Create a trigger to run `runBigQueryAndUpdateSheet` function
