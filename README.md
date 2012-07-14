windage
=======

Goofy Google Docs javascript for a yacht racing results spreadsheet

To use it, you want to do e.g.:

  function onOpen() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menuItems = [ { name: "Calculate", functionName: "calc" } ];
    ss.addMenu("Results", menuItems);
  }

  function calc() {
    return calculateResults();
  }

  var windage = UrlFetchApp.fetch("https://raw.github.com/markmc/windage/master/windage.js");
  eval(windage.getContentText());
