// Google Apps Script Functions Useful for Analysts


function getWeekNumber(d) {
  // Copy date so don't modify original
  d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  // Set to nearest Thursday: current date + 4 - current day number
  // Make Sunday's day number 7
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay()||7));
  // Get first day of year
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(),0,1));
  // Calculate full weeks to nearest Thursday
  var weekNo = Math.ceil(( ( (d - yearStart) / 86400000) + 1)/7);
  // Return array of year and week number
  return [d.getUTCFullYear(), weekNo];
}

// It takes as date string and converts to yyyy-mm-dd 
function formatDate(date) {
    var d = new Date(date),
        month = '' + (d.getMonth() + 1),
        day = '' + d.getDate(),
        year = d.getFullYear();
    if (month.length < 2) month = '0' + month;
    if (day.length < 2) day = '0' + day;
    return [year, month, day].join('-');
}

// Function Leading Zeroes
function zeroPad(num, places) {
  var zero = places - num.toString().length + 1;
  return Array(+(zero > 0 && zero)).join("0") + num;
}


// Where data is a 2d array
function convertArrayToCsv(data) {
   var csvFile = undefined;
    // loop through the data in the range and build a string with the csv data
    if (data.length > 1) {
      var csv = "";
      for (var row = 0; row < data.length; row++) {
        for (var col = 0; col < data[row].length; col++) {
          if (data[row][col].toString().indexOf(",") != -1) {
            data[row][col] = "\"" + data[row][col] + "\"";
          }
        }
        // join each row's columns
        // add a carriage return to end of each row, except for the last one
        if (row < data.length-1) {
          csv += data[row].join(",") + "\r\n";
        }
        else {
          csv += data[row];
        }
      }
      csvFile = csv;
   }  
  return csvFile;
}

// Takes a sheet object (not a sheet id -- an instantiated sheet object)
// and returns a string with the format csv
function convertRangeToCsvFile_(sheet) {
  // get available data range in the spreadsheet
  var activeRange = sheet.getDataRange();
  try {
    var data = activeRange.getValues();
    var csvFile = undefined;
    // loop through the data in the range and build a string with the csv data
    if (data.length > 1) {
      var csv = "";
      for (var row = 0; row < data.length; row++) {
        for (var col = 0; col < data[row].length; col++) {
          if (data[row][col].toString().indexOf(",") != -1) {
            data[row][col] = "\"" + data[row][col] + "\"";
          }
        }
        // join each row's columns
        // add a carriage return to end of each row, except for the last one
        if (row < data.length-1) {
          csv += data[row].join(",") + "\r\n";
        }
        else {
          csv += data[row];
        }
      }
      csvFile = csv;
    }
    return csvFile;
  }
  catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
  }
}


// Gets data from google big query returning an object like this
// {results: queryResults, data:data, headers:headers}
function getQueryData(queryStr, projectId, useLegacySql) {
  if (projectId === undefined) {projectId = 'tmg-plat-dev'};
  if (useLegacySql === undefined) {useLegacySql = false};
  DriveApp.getRootFolder();' to get Drive Auth Scope'
  var job = {
    configuration: {
      query: {
        query: queryStr,
        writeDisposition: null,
        destinationTable: null,
        useLegacySql: useLegacySql
      }
    }
  };

  var queryResults = BigQuery.Jobs.insert(job, projectId);
  var jobId = queryResults.jobReference.jobId;

  // Check on status of the Query Job and wait if not completed.
  var sleepTimeMs = 500;
  while (!queryResults.jobComplete) {
    Utilities.sleep(sleepTimeMs);
    sleepTimeMs *= 2;
    queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId);
  }
  
  // Get all the rows of results.
  var rows = queryResults.rows;
  while (queryResults.pageToken) {
    queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId, {
      pageToken: queryResults.pageToken
    });
    rows = rows.concat(queryResults.rows);
  };
  
  // If rows is undefined then finish function returning empty arrays
  if (rows === undefined) {
    return {results: queryResults, data:[[]], headers:[]};
  }

  // This is when we actually got any results  
  else {
  
  // Sort rows in a nice array of arrays
  // and build the array of arrays `data`
  var data = new Array(rows.length);
  for (var i = 0; i < rows.length; i++) {
    var cols = rows[i].f;
    data[i] = new Array(cols.length);
    for (var j = 0; j < cols.length; j++) {
      data[i][j] = cols[j].v;
    }
  }

  // Append the headers.
    var headers = queryResults.schema.fields.map(function(field) {
      return field.name;
    });
  
  return {results: queryResults, data:data, headers:headers};
  }
}

// takes data as 2d array and headers, clears contents on destination sheet, and paste values
function exportDataToSheet(data, headers, sheetId, tabName) {
  
    var spreadsheet = SpreadsheetApp.openById(sheetId);
    var sheet = spreadsheet.getSheetByName(tabName);
    sheet.clearContents();
    sheet.appendRow(headers);
    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
}

function writeQueryToTable(queryStr, projectId, datasetId, tableId, useLegacySql) {

  if (useLegacySql === undefined) {useLegacySql = false};
  var job = {
    configuration: {
      query: {
        query: queryStr,
        writeDisposition: 'WRITE_TRUNCATE',
        useLegacySql: useLegacySql,
        destinationTable: {
          projectId: projectId,
          datasetId: datasetId,
          tableId: tableId
        }
      }
    }
  };

  var queryResults = BigQuery.Jobs.insert(job, projectId);
  var jobId = queryResults.jobReference.jobId;

  // Try disabling status check to avoid using too much computer time
  // Check on status of the Query Job.
  //  var sleepTimeMs = 500;
  //  while (!queryResults.jobComplete) {
  //    Utilities.sleep(sleepTimeMs);
  //    sleepTimeMs *= 2;
  //    queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId);
  //  }
  //  
  //  Logger.log(queryResults.status);

}


// This function takes a sheet (not a spreadsheet) and an array of string patterns
// Then it matches each cell on the range 1:1 (header), and returns the ranges that match the patterns
// for example if we want to get the columns where the header contains 'Date ' or ' Date'
// it will return an array like ['A','C','X'] where the columns match those patterns.
// This is useful when editors might change the position of certain columns, although still relying on consistent names
function findRangeInHeader(sheet, patterns) {
  var searchRange = sheet.getRange('1:1');
  var values = searchRange.getValues()[0];
  var addressesFound = [];
  // for value index in
  for (vi in values) {
    var value = values[vi].toString();
    var match = false;
    // for pattern index in patterns
    for (pi in patterns) {
      var pattern = patterns[pi];
      // if value matches any of the patterns
      if (value.indexOf(pattern) > -1) {match = true; break}
    }
    // if for this header value, we have a match
    if (match === true) {
      // crazy enough we need to convert the index to number otherwise it concatenates strings (?)
      var colPos = Number(vi) + 1;
      var range = sheet.getRange(1, colPos);
      var rangeAddress = range.getA1Notation().replace('1','');
      addressesFound.push(rangeAddress);
    }
  }
  return addressesFound;
}

// This fuction returns a string that can be used to create a Blob
// for an  Excel CSV file from headers and rows (data)
// It appends at the top of the file the instruction "sep=|"
// which allows excel to recognize the pipe delimiter which is more convenient

function dataToExcelCsvString(data, headers, addSep) {
  // Basically if undefined or true, then yes, add the Excel explicit delimiter
  // We also remove the pipe character "|" from strings 
  var queryDataStr;
   if (addSep === undefined || addSep === true) {
     queryDataStr = 'sep=|'+"\r\n";   
   }
  else {
    queryDataStr = ''  ;
  }
  queryDataStr += headers.join("|") + "\r\n";
  // row index
  for (ri in data) {
    // remove pipe
    queryDataStr += data[ri].join("|") + "\r\n";
  }
  return queryDataStr;
}

// Create 2d Array to HTML Table
// Taken from https://stackoverflow.com/questions/34494032/loading-a-csv-file-into-an-html-table-using-javascript
function arrayToHTMLTable(array) {
//  var content = "<table border=\"1\"";
//    array.forEach(function(row) {
//        content += "<tr>";
//        row.forEach(function(cell) {
//            content += "<td>" + cell + "</td>";
//        });
//        content += "</tr>";
//    });
//    return content+'</table>';
  var content = '<table border=\"1\" bordercolor=\"black\" rules=\"cols\">';
  for (var i = 0; i < array.length; i++) {
    var row = array[i];
    if (i === 0) {
      content += "<tr style=\"background-color: #53DCD0;\">";
      for (var j = 0; j < row.length; j++) {
        var cell = row[j];
        content += "<td>" + cell + "</td>";
      }
      content += "</tr>";
    }
    else {
      content += "<tr style=\"background-color: #E0E0E0;\">";
      for (var j = 0; j < row.length; j++) {
        var cell = row[j];
        content += "<td>" + cell + "</td>";
      }
      content += "</tr>";
    }
  }
  return content+"</table>";
}

// Upload array to BigQueryTable
// We could consider having a function that creates the table if it does not exists
// but that might end up in further complications assuming we need to either
// specify table headers + schema
// or use autodetect, in which case the column names end up with generic names
function upload2dArrayToExistingTable(data, projectId, datasetId, tableId, writeDisposition) {
  DriveApp.getRootFolder();' to get Drive Auth Scope';
  var csvData = convertArrayToCsv(data);
  var blob = Utilities.newBlob(csvData, 'application/octet-stream');
  // Firstly try to upload assuming the table exists
  try {
    // Create the data upload job.
  var job = {
    configuration: {
      load: {
        //schema:{fields:[{name:'col1', type:'string'},{name:'col2', type:'string'},{name:'col3', type:'string'}]},
        //autodetect: true,
        writeDisposition: writeDisposition,
        destinationTable: {
          projectId: projectId,
          datasetId: datasetId,
          tableId: tableId,
        },
        skipLeadingRows: 0,
      },
    },
  };
  var queryResults = BigQuery.Jobs.insert(job, projectId, blob);
  Logger.log('Load job started. Check on the status of it here: ' +
      'https://bigquery.cloud.google.com/jobs/%s', projectId);
    
  var jobId = queryResults.jobReference.jobId;

  // Check on status of the Query Job.
  var sleepTimeMs = 500;
  while (!queryResults.jobComplete) {
    Utilities.sleep(sleepTimeMs);
    sleepTimeMs *= 2;
    queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId);
  }
  Logger.log(queryResults.status);
  }
  catch(err) {
  Logger.log(err);
  }

}


function getAttachmentsBySearch(searchPattern, fileNamePattern, nFilesToFind) {
  //search gmail with the given query(partial name using * as a wildcard to find anything in the current subject name).
  var threads = GmailApp.search(searchPattern);
  //retrieve all messages in the specified threads.
  // This is arranged in a 2d array e.g. [[message,message],[message],[message,message,message]]
  var msgs = GmailApp.getMessagesForThreads(threads);
  var filesFound = [];
  for (var i = 0; i < msgs.length; i++) {
    var thread = msgs[i];
    for (var j = 0; j < thread.length; j++) {
      var msg = thread[j];
      var attachments = msg.getAttachments();
      for (var k = 0; k < attachments.length; k++) {
        var attachment = attachments[k];
        var fileName = attachment.getName();
        var match = fileName.match(fileNamePattern);
        if (match !== null) {
          filesFound.push(attachment);        
        
        }
        // If limit reached return
        if (filesFound.length === nFilesToFind) {return filesFound}
      }
    }
  }
  return filesFound;
}

// Adding missing range function as in Python
function range(start, stop, step){
  var a=[start], b=start;
  while(b<stop-1){b+=(step || 1);a.push(b)}
  return a;
};

function getSheetsIdsFromFolder(folderId, idsToNeglect) {
    var folder = DriveApp.getFolderById(folderId);
    var data = [];
    var files = folder.getFiles();
    while (files.hasNext()){
      file = files.next();
      var row = [];
      row.push(file.getName(),file.getId(),file.getSize(),date);
      // if the file id is not in the list of ids to neglect
      if  (idsToNeglect.indexOf(file.getId()) === -1) {
        data.push(row);
      }
    }
  return data
}

function extractTableToStorage(projectId, datasetId, tableId, destinationUris) {
  // https://cloud.google.com/bigquery/docs/reference/rest/v2/Job#JobConfigurationExtract
  var sourceTable = BigQuery.Tables.get(projectId, datasetId, tableId);
  var projectId = projectId;
  var job = {
    jobType: 'EXTRACT',
    configuration: {
      extract: {
        destinationUris: destinationUris,
        printHeader:true,
        fieldDelimiter: null, // default is CSV
        destinationFormat: null, // default is CSV
        compression: null, // default is 'NONE'
        sourceTable: sourceTable.tableReference
      }
    }
  };
  var queryResults = BigQuery.Jobs.insert(job=job, projectId=projectId);
  var jobId = queryResults.jobReference.jobId;
  Logger.log(jobId);
}

function testFoo() {
  var projectId = 'tmg-plat-dev';
  var datasetId = 'glc';
  var tableId = 'riviera_2';
  var destinationUris = ['gs://tmg-finance/riviera_test_4.csv'];
  extractTableToStorage(projectId, datasetId, tableId, destinationUris);  
}