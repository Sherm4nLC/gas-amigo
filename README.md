# gas-amigo
Google Apps Script Useful Functions

## Example 1: BigQuery to Sheets
Import library, execute query and get data and headers.
```
var gami = GASAmigo;
var queryStr = 'select lat, lon, price from `properati-data-public.properties_ar.properties_rent_201501` limit 100';
var projectId = 'glc-01'; // this is the project that gets billed
var results = gami.getQueryData(queryStr, projectId, false);
var data = results["data"]; // this is a 2d array
var headers = results["headers"]; // these are the column names
```
Now send to sheet
```
var sheetId = '1B6uA-hE8LDLmOtbqtvM-vc8AHuDkLXGYpclwJCFdOSc';
var tabName = 'Sheet1';
gami.exportDataToSheet(data, headers, sheetId, tabName); // this onverrides any existing content
```


## Installation in 3 steps
1. Import library
Go to your GAS project, from the main menu on the top<br>
"Resources" > "Libraries"
add the id of the library
`MeEY_ICVdfxPmGkVzM0iAwLQZHlojl1UC`
Select the latest version.
![alt text](https://raw.githubusercontent.com/Sherm4nLC/gas-amigo/master/docs/docs01.png)

2. Set Scopes Explicitly


https://developers.google.com/apps-script/concepts/scopes