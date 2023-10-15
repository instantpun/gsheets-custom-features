function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('New Features')
      .addSubMenu(ui.createMenu('Duplicates')
        .addItem('Highlight in Selection', 'highlightDuplicatesInSelection')
        .addItem('Remove in Selection', 'removeDuplicatesInSelection')
      )
      .addToUi();
}

function highlightDuplicatesInSelection() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  const [dupIndex, err] = getDuplicatesInSelection(sheet);
  if (err != null) {
    console.log(err);
    Browser.msgBox(err);
  }
  if (dupIndex === null || dupIndex === undefined) {
    console.log("dupIndex",dupIndex);
    return;
  }
  
  highlightColorMap = {
    "Cerulean Blue": [152, 180, 212],
    "Fuschia Rose": [195, 68, 122],
    "True Red": [188, 36, 60],
    "Aqua Sky": [127, 205, 205],
    "Tigerlily": [225, 93, 68],
    "Turquosie": [85, 180, 176],
    "Sand Dollar": [223, 207, 190],
    "Chili pepper": [155, 35, 53],
    "Blue Izis": [91, 94, 166],
    "Mimosa": [239, 192, 80],
    "Honeysuckle": [214, 80, 118],
    "Emerald": [0, 155, 119],
    "Rose Quartz": [247, 202, 201],
    "Serenity": [146, 168, 209],
    "Greenery": [136, 176, 75]
  };

  var randomProperty = function (obj) {
      var keys = Object.keys(obj);
      random_key = keys.length * Math.random() << 0
      return keys[random_key], obj[keys[random_key]]; // random key, value
  };

  highlights: {
    for (const [key, values] of Object.entries(dupIndex)) {
      // console.log(`${key}: ${values}`);
      if (values.length > 1) {
        var _, rgb = randomProperty(highlightColorMap);
        for (var i = 0; i < values.length; i++) {
          sheet.getRange(values[i]).setBackgroundRGB(...rgb);
          // toSelect.push(values[i]);
        }
      }
    }
  }

}

function removeDuplicatesInSelection() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  const [dupIndex, err] = getDuplicatesInSelection(sheet);
  if (err) {
    Browser.msgBox(err);
  }
  if (dupIndex === null || dupIndex === undefined) {
    return;
  }

  removal: {
    dropList = [];

    for (const [key, values] of Object.entries(dupIndex)) {
      // console.log(`${key}: ${values}`);
      if (values.length > 1) {
        for (var i = 1; i < values.length; i++) {
          dropList.push(values[i])
        }
      }
    }

    sheet.getRangeList(dropList).clearContent();
  }
}

function getDuplicatesInSelection(activeSpreadSheet) {
  var sheet = activeSpreadSheet;
  var selection = sheet.getSelection();
  var selected_range = selection.getActiveRange()
  // console.log('Active Range: ' + selected_range.getA1Notation());
  var start_row = selected_range.getRow();
  var start_col = selected_range.getColumn();
  // console.log("starting pos", start_row, start_col);
  var end_row = selected_range.getLastRow();
  var end_col = selected_range.getLastColumn();
  var numRows = end_row - start_row + 1;
  var numCols = end_col - start_col + 1;
  // console.log("end",end_row, end_col)
  var err = null;

  if (selected_range.getValues().length == 0) {
    err = "Please make a selection and try again.";
    return null, err;
  } else if ( selected_range.getValues().length == 1) {
    err = "Selection is too small to find matches. Please try again.";
    return null, err;
  }

  // console.log(selected_range.getValues());
  var index = {};

  var emptyCellStopLimit = 30;
  var emptyCellCount = 0;
  var noDupesFound = true;

  indexing: {
    for (var row = 1; row <= numRows; row++) {
      for (var col = 1; col <= numCols; col++) {
        console.log(row, col)
        cell = selected_range.getCell(row,col);
        cellA1name = cell.getA1Notation();
        value = cell.getValues()[0][0]; // always singular value
        console.log(cellA1name, value);
        if (!value) {
          emptyCellCount++;
          console.log(emptyCellCount);
          if (emptyCellCount >= emptyCellStopLimit) {
            err = `Stopped searching after ${emptyCellStopLimit} empty cells.`;
            break indexing; // break loop when scanning too many empty cells
          } else {
            continue; // skip processing of empty cells
          }
        } else {
          emptyCellCount = 0;

          if (Array.isArray(index[value])) {
            if (index[value].length > 1) {
              noDupesFound = false;
            }
            index[value].push(cellA1name)
          } else {
            index[value] = [cellA1name]
          };
        }
        
      };
    };
  }
  if (!noDupesFound) {
    err = "No duplicate entries were found.";
  }
  console.log(index, err);
  return [index, err];
}