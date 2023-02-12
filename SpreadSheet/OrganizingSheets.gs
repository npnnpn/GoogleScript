function shiftUpOnBlank_getFormula() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var data = sheet.getDataRange().getValues();
    var newData = [];
    var formula = [];
    var displayText = [];
    
    for (var j = 0; j < data.length; j++) {
      if (data[j].join("").length > 0) {
        newData.push(data[j]);
        var cell = sheet.getRange(j + 1, 1);
        var hyperlink = cell.getFormula();
        if (hyperlink.indexOf('HYPERLINK') === 0) {
          displayText.push(cell.getDisplayValue());
          var link = hyperlink.match(/"(.*?)"/)[1];
          formula.push([j + 1, 1, '=HYPERLINK("' + link + '", "' + displayText[j] + '")']);
        }
      } else {
        for (var k = 0; k < formula.length; k++) {
          if (formula[k][0] > j + 1) {
            formula[k][0] = formula[k][0] - 1;
          }
        }
      }
    }
    
    sheet.clearContents();
    sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
    
    for (var j = 0; j < formula.length; j++) {
      sheet.getRange(formula[j][0], formula[j][1]).setFormula(formula[j][2]);
    }
  }
}


function shiftUpOnBlank_getFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var data = sheet.getDataRange().getValues();
    var newData = [];
    var formulas = sheet.getFormulas();
    var newFormulas = [];

    for (var j = 0; j < data.length; j++) {
      if (data[j].join("").length > 0) {
        newData.push(data[j]);
        for (var k = 0; k < formulas.length; k++) {
          if (formulas[k][0] == j + 1 && formulas[k][1].indexOf("HYPERLINK") == 0) {
            newFormulas.push([j + 1, formulas[k][1]]);
          }
        }
      } else {
        for (var k = 0; k < formulas.length; k++) {
          if (formulas[k][0] > j + 1 && formulas[k][1].indexOf("HYPERLINK") == 0) {
            formulas[k][0]--;
            newFormulas.push([formulas[k][0], formulas[k][1]]);
          }
        }
      }
    }

    sheet.clearContents();
    sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);

    for (var j = 0; j < newFormulas.length; j++) {
      sheet.getRange(newFormulas[j][0], 1).setFormula(newFormulas[j][1]);
    }
  }
}

function shiftUpOnBlank_WithoutLink() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var data = sheet.getDataRange().getValues();
    var newData = [];
    
    for (var j = 0; j < data.length; j++) {
      if (data[j].join("").length > 0) {
        newData.push(data[j]);
      }
    }
    
    sheet.clearContents();
    sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
  }
}



function shiftUpOnBlank_getHyperlinks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var data = sheet.getDataRange().getValues();
    var newData = [];
    var links = sheet.getHyperlinks();
    var newLinks = [];
    
    for (var j = 0; j < data.length; j++) {
      if (data[j].join("").length > 0) {
        newData.push(data[j]);
        for (var k = 0; k < links.length; k++) {
          if (links[k].getRow() == j + 1) {
            newLinks.push([links[k].getRow() - j, links[k].getColumn(), links[k].getUrl()]);
          }
        }
      } else {
        for (var k = 0; k < links.length; k++) {
          if (links[k].getRow() > j + 1) {
            newLinks.push([links[k].getRow() - 1, links[k].getColumn(), links[k].getUrl()]);
          }
        }
      }
    }
    
    sheet.clearContents();
    sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
    
    for (var j = 0; j < newLinks.length; j++) {
      sheet.getRange(newLinks[j][0], newLinks[j][1]).setFormula('=HYPERLINK("' + newLinks[j][2] + '", "Link")');
    }
  }
}
