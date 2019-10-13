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
You are done 🧙✨📈


## Installation in 3 steps
### 1. Import library
Go to your GAS project, from the main menu on the top and do "Resources" > "Libraries" add the id of the library<br>
**`MeEY_ICVdfxPmGkVzM0iAwLQZHlojl1UC`**<br>
Select the latest version.
<img src="https://raw.githubusercontent.com/Sherm4nLC/gas-amigo/master/docs/docs01.png" width="400"/>
<img src="https://raw.githubusercontent.com/Sherm4nLC/gas-amigo/master/docs/docs02.png" width="400"/>
<img src="https://raw.githubusercontent.com/Sherm4nLC/gas-amigo/master/docs/docs03.png" width="400"/>

### 2. Set Scopes Explicitly
Before proceeding make sure you understand the [concepts](https://developers.google.com/apps-script/concepts/scopes) on setting scopes.

On the top menu select "View" > "Show manifest file".
<img src="https://raw.githubusercontent.com/Sherm4nLC/gas-amigo/master/docs/docs04.png" width="400"/>
This will make visible an additional file on your project called `appsscript.json`.
<img src="https://raw.githubusercontent.com/Sherm4nLC/gas-amigo/master/docs/docs05.png" width="400"/>
<br>
On the top part of the file add the following lines.<br>
**Remove any scopes you consider you don't need, depending on the functions you want to use.**
```
"oauthScopes": [
    "https://mail.google.com/",
    "https://www.googleapis.com/auth/bigquery",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"
    ]
```
<img src="https://raw.githubusercontent.com/Sherm4nLC/gas-amigo/master/docs/docs06.png" width="400"/>

