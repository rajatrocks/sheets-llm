var apiEndpoint = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=';

function callLLM() {
  // first check that we have an API Key. Error out if we don't. 
  var APIKey = getAPIKey();
  // if there's nothing there, error out
  if (APIKey == "") {
    Browser.msgBox ("You need to put in a valid API Key in cell A1 on the SETUP sheet!");
    return;
  }
  // if there are spaces in it, we know it hasn't been updated. 
  var parts = APIKey.split(" ");
  if (parts.length > 1) {
    Browser.msgBox ("You need to put in a valid API Key in cell A1 on the SETUP sheet!");
    return;
  }
  else {
    var fullEndpoint = apiEndpoint + APIKey;
  }

  // next make sure we're on a sheet that has a prompt
  var promptLabel = getCellValue('A1');
  if (promptLabel != "Prompt") {
    Browser.msgBox ("You can't send this sheet to an LLM!");
    return;  
  }

  // get the text in the prompt field
  var prompt = getCellValue('B1');
  if (prompt == "") {
    Browser.msgBox ("You need to enter a prompt!");
    return;
  }

  // clear out the response field
  setCellValue('B2', '');

  // replace any variables in the prompt field
  var matches = findWordsInDoubleBrackets(prompt);

  // for each variable, find the corresponding cell with the value
  matches.forEach(function(item) {
    var dataCell = findCellWithValue(item);
    if (dataCell != "") {
      // replace the variable in prompt with the contents of the data cell
      prompt = prompt.replace("[[" + item + "]]", getCellValue(dataCell));
    }
  });

  // get the temperature
  var temp = getCellValue('B3');
  if (temp == "") {
    temp = "0.1";
  }

  console.log ("Calling with prompt: " + prompt);

  // make the API call
  // because this is just for experiementation, remove all the safety blocks
    var nlData = {
      'contents': [{
        'parts': [
          {'text': prompt}
        ]
      }],
      'generationConfig': {
        'temperature': parseFloat(temp),
      },
      'safetySettings':[
          {
              "category": "HARM_CATEGORY_HATE_SPEECH",
              "threshold": "BLOCK_NONE"
          },
          {
              "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT",
              "threshold": "BLOCK_NONE"
          },
          {
              "category": "HARM_CATEGORY_DANGEROUS_CONTENT",
              "threshold": "BLOCK_NONE"
          },
          {
              "category": "HARM_CATEGORY_HARASSMENT",
              "threshold": "BLOCK_NONE"
          }
      ]
    };

    // Packages all of the options and the data together for the API call.
    const nlOptions = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(nlData)
    };
    // Makes the API call.
    let response = UrlFetchApp.fetch(fullEndpoint, nlOptions);
    const jsonData = JSON.parse(response);
    console.log (jsonData);

    if ("candidates" in jsonData) {
      // Put the response into the Result field. 
      setCellValue('B2', jsonData.candidates[0].content.parts[0].text);
      console.log(jsonData.candidates[0].content.parts[0].text);
    }
    else {
      Browser.msgBox ("There was an error of some kind");
    }
}

// Add the Sheets LLM menu item
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries0 = [
      {
          name: "Run",
          functionName: "callLLM",
      },
  ];

  ss.addMenu("Sheets LLM", menuEntries0);
}

// Get a cell value
function getCellValue(cell) {
  var sheet = SpreadsheetApp.getActiveSheet();

  var range = sheet.getRange(cell);
  var value = range.getValue();
  return value;
}

// set a cell value
function setCellValue(cell, value) {
  var sheet = SpreadsheetApp.getActiveSheet();

  var range = sheet.getRange(cell);
  range.setValue(value);
}

// find the variables in the prompt
// function findWordsInDoubleBrackets(prompt) {
//   var matches = prompt.match(/\[\[(.*?)\]\]/g).map(function(val){
//     return val.replace(/\[\[(.*?)\]\]/g, '\$1');
//   });
//   return matches;
// }

function findWordsInDoubleBrackets(prompt) {
  // Use matchAll instead of match to get an iterator that we can directly loop over
  const matchesIterator = [...prompt.matchAll(/\[\[(.*?)\]\]/g)];

  // Check if there were any matches at all
  if (matchesIterator.length === 0) {
    return []; // Return an empty array if there are no matches
  }

  // Map over the matches to replace the double brackets with just the content inside them
  const matches = matchesIterator.map(match => match[0].replace(/\[\[(.*?)\]\]/g, '$1'));
  
  return matches;
}


// find the right data cell for a variable. 
function findCellWithValue(val) {
  var sheet = SpreadsheetApp.getActiveSheet();

  var range = sheet.getRange(1, 1, sheet.getLastRow(), 1);
  var textFinder = range.createTextFinder(val);
  var range = textFinder.findNext();
  
  if (range) {
    var row = range.getRow();
    var column = range.getColumn();
    return "B" + row;
  } else {
    return "";
  }
}

// Get the API Key from the SETUP sheet
function getAPIKey() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('SETUP');
  var range = sheet.getRange('A1');
  var value = range.getValue();
  return value;
}


