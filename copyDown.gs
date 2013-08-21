var scriptTitle = "copyDown Script V2.2 (11/7/12)";
// Written by Andrew Stillman for New Visions for Public Schools
// Published under GNU General Public License, version 3 (GPL-3.0)
// See restrictions at http://www.opensource.org/licenses/gpl-3.0.html
// Support and contact at http://www.youpd.org/copydown
var scriptTrackingId = "UA-31030730-1"
var scriptName = "copyDown"


var ss = SpreadsheetApp.getActiveSpreadsheet();
var alphabet = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ"];
var COPYDOWNIMAGEURL = 'https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/searchable-docs-collection/copyDown_icon.gif?attachauth=ANoY7cq0pvQqsSiCo2XLcXxjrzNDmlWtNTjFIDM60Wz9CUdiFUYR6UiFB-CF81KKwD7T2EIjdA1JNd65Ndp-d_KypSbOqTv2QdduiiEIwLm3AuaH-iF6kjf5GK-ir7ew5UPWbqiAxl5cVjhvlXZaZGHNpKOb0I78JbuAVcPmoc8uMzAChZ_iHuS_7b6IN_IYF1VgeOWBnIjel6ZCWgGlyfIR65MWvv0bhs1ztQCZRYdQQj96D3ZdcCWugeiHtYCS13_cY-7VU4KT&attredirects=0';

function onInstall() {
  onOpen();
}

function onOpen() {
  var ss = SpreadsheetApp.getActive();
  var menuEntries = [];
      menuEntries.push({name: "What is copyDown?", functionName: "copyDown_whatIs"});
      menuEntries.push({name: "Run initial configuration", functionName: "copyDown_preconfig"});
  ss.addMenu("copyDown", menuEntries);
  copyDown_initialize();
}

function copyDown_initialize() {
  var ss = SpreadsheetApp.getActive();
  var menuEntries = [];
  menuEntries.push({name: "What is copyDown?", functionName: "copyDown_whatIs"});
  var preconfigStatus = ScriptProperties.getProperty('preconfigStatus');
  if (preconfigStatus) {
    menuEntries.push({name: "Create/manage copyDown jobs", functionName: "copyDown_createJob"});
    menuEntries.push({name: "Run copydown", functionName: "runCopyDown"});
    menuEntries.push(null);
    menuEntries.push({name: "Advanced options", functionName: "copyDown_advanced"});
  } else {
    menuEntries.push({name: "Run initial configuration", functionName: "copyDown_preconfig"});
  }
  ss.addMenu("copyDown", menuEntries);
}


function copyDown_createJob() {
  copyDown_getInstitutionalTrackerObject();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = ScriptProperties.getProperties();
  var sheets = ss.getSheets();
  var sheetNames = [];
  for (var i=0; i<sheets.length; i++) {
    sheetNames.push(sheets[i].getName());
  }
  var activeSheet = ss.getActiveSheet();
  var activeSheetIndex = activeSheet.getIndex();
  var activeSheetName = activeSheet.getName();
  var activeSheetSelectIndex = sheetNames.indexOf(activeSheetName);
  var app = UiApp.createApplication().setTitle("Manage copyDown jobs").setHeight(400).setWidth(700);
  var panel = app.createVerticalPanel().setId("panel").setWidth("680px").setHeight("200px");
  var helpLabel = app.createLabel("Be aware that any changes you make to the column order or structure of this sheet will require you to redo your copyDown settings.")
  var refreshSheetHandler = app.createServerHandler('refreshSheet').addCallbackElement(panel);
  var columns = activeSheet.getLastColumn();
  var grid = app.createGrid(4, columns+1).setId('grid');
  var spinner = app.createImage(COPYDOWNIMAGEURL).setWidth("115px").setId('spinner');
  spinner.setVisible(false);
  spinner.setStyleAttribute("position", "absolute");
  spinner.setStyleAttribute("top", "120px");
  spinner.setStyleAttribute("left", "220px");
  var refreshOpacityHandler = app.createClientHandler().forTargets(panel).setStyleAttribute('opacity', '0.5').forTargets(spinner).setVisible(true);
  var sheetSelect = app.createListBox().setName('sheetSelected');
  for (var i=0; i<sheetNames.length; i++) {
    sheetSelect.addItem(sheetNames[i]);
  }
  if (activeSheetSelectIndex!=-1) {
    sheetSelect.setSelectedIndex(activeSheetSelectIndex);
  }
  sheetSelect.addChangeHandler(refreshOpacityHandler).addChangeHandler(refreshSheetHandler);
  
  
  panel.add(sheetSelect);
  var sheetId = activeSheet.getSheetId();
  var hiddenSheetIdBox = app.createTextBox().setValue(sheetId).setId('sheetId').setName('sheetId').setVisible(false);  
  var hiddenColNumBox = app.createTextBox().setValue(columns).setId('numCols').setName('numCols').setVisible(false);
  panel.add(hiddenSheetIdBox);
  panel.add(hiddenColNumBox);
  var noDataLabel = app.createLabel("No data in sheet").setId('noDataLabel');
  noDataLabel.setVisible(false);
  panel.add(noDataLabel);
  panel.add(grid);
  app.add(panel);
  app.add(spinner);
  copyDown_returnSheetUi(activeSheet, properties);
  ss.show(app);        
}


function manualSave(e) {
  var app = UiApp.getActiveApplication();
  saveCopyDownSettings(e);
  app.close();
  return app;
}


function refreshSheet(e) {
  var app = UiApp.getActiveApplication();
  saveCopyDownSettings(e);
  var properties = ScriptProperties.getProperties();
  var sheetName = e.parameter.sheetSelected;
  var sheet = ss.getSheetByName(sheetName);
  copyDown_refreshSheetUi(sheet, properties);
  return app;
}


function copyDown_returnSheetUi(sheet, properties) {
  var app = UiApp.getActiveApplication();
  var panel = app.getElementById('panel');
  panel.setStyleAttribute('opacity','1');
  var spinner = app.getElementById('spinner');
  spinner.setVisible(false);
  var scrollPanel = app.createScrollPanel().setWidth("660px").setStyleAttribute('margin', '10px');
  var grid = app.getElementById('grid');
  var sheetId = sheet.getSheetId();
  var sheetProperties = properties['sheet-'+sheetId];
  if ((sheetProperties)&&(sheetProperties!='')) {
    sheetProperties = Utilities.jsonParse(sheetProperties);
  } else {
    sheetProperties = new Object();
  }
  var hiddenSheetIdBox = app.getElementById('sheetId').setValue(sheetId);
  var columns = sheet.getLastColumn();
  var hiddenColNumBox = app.getElementById('numCols').setValue(columns);
  if (sheetProperties['formulaRow']) {
    var row = parseInt(sheetProperties['formulaRow']);
  } else {
    var row = 2;
  }
  var rowsArray = [2,3,4];
  var copyCol = sheetProperties['copyCol'];
  if (!copyCol) {
    copyCol = "A";
  }
  var noDataLabel = app.getElementById('noDataLabel');
  if (columns>0) {
    var colValues = sheetProperties['colValues'];
    if (colValues) {
      colValues = Utilities.jsonParse(colValues);
    } else {
      colValues = new Object();
    }
    noDataLabel.setVisible(false);
    grid.resize(4, columns+1);
    grid.setVisible(true);
    var horizontalPanel = app.createHorizontalPanel().setId('horizPanel');
    var formulaRowLabel = app.createLabel("Row containing master values/formulas");
    var formulaRowSelect = app.createListBox().setId('formulaRowSelect').setName("formulaRow");
    var rowSelectHandler = app.createServerHandler('copyDown_refreshRowFormulas').addCallbackElement(panel);
    var lastRow = sheet.getLastRow();
    for (var i=2; i<lastRow+1; i++) {
      formulaRowSelect.addItem(i);
    }
    var rowSelectIndex = rowsArray.indexOf(row);
    formulaRowSelect.setSelectedIndex(rowSelectIndex);
    grid.setBorderWidth(1).setCellSpacing(0).setStyleAttribute('borderColor','#E5E5E5').setStyleAttribute('opacity', '1');
    formulaRowSelect.addChangeHandler(rowSelectHandler);
    horizontalPanel.add(formulaRowLabel);
    horizontalPanel.add(formulaRowSelect);
    var formulaRowLabel = app.createLabel("Column containing last row to copy down to");
    var copyColSelect = app.createListBox().setId('copyColSelect').setName("copyCol");
    for (var i=0; i<columns; i++) {
      copyColSelect.addItem(this.alphabet[i]);
    }
    var colSelectIndex = this.alphabet.indexOf(copyCol);
    if (colSelectIndex!=-1) {
      copyColSelect.setSelectedIndex(colSelectIndex);
    }
    horizontalPanel.add(formulaRowLabel);
    horizontalPanel.add(copyColSelect);
    panel.add(horizontalPanel);
   
    var headers = sheet.getRange(1,1,1,columns).getValues()[0];
    
    grid.setWidget(0, 0, app.createLabel("Column"));
    grid.setWidget(1, 0, app.createLabel("Header"));  
    grid.setWidget(2, 0, app.createLabel("Value/Formula"));  
    grid.setWidget(3, 0, app.createLabel("Paste as values?"));  
    var onButtonHandlers = [];
    var onButtonServerHandlers = [];
    var offButtonHandlers = [];
    var offButtonServerHandlers = [];
    var onButtons = [];
    var offButtons = [];
    var buttonValues = [];
    var formulaLabels = [];
    var asValuesCheckBoxes = [];
    for (var i=0; i<columns; i++) {
      onButtons[i] = app.createButton(this.alphabet[i]).setId('onButton-'+sheetId+'-'+i).setStyleAttribute('background', 'whiteSmoke').setWidth("50px");
      offButtons[i] = app.createButton(this.alphabet[i]).setId('offButton-'+sheetId+'-'+i).setStyleAttribute('background', '#E5E5E5').setStyleAttribute('border', '2px solid grey').setWidth("50px").setVisible(false);
      buttonValues[i] = app.createTextBox().setVisible(false).setText(i+'-off').setName('bv-'+sheetId+'-'+i);
      var buttonPanel = app.createHorizontalPanel();
      buttonPanel.add(onButtons[i])
                 .add(offButtons[i])
                 .add(buttonValues[i])
                 .setStyleAttribute('width',"80px")
                 .setHorizontalAlignment(UiApp.HorizontalAlignment.CENTER);
      grid.setWidget(0, i+1, buttonPanel).setStyleAttribute(0, i+1, 'backgroundColor', 'whiteSmoke').setStyleAttribute(0, i+1, 'textAlign', 'center');
      grid.setWidget(1, i+1, app.createLabel(headers[i]));
      var formulas = copyDown_getSheetFormulas(sheet, row);
      var formulaLabel = app.createLabel(formulas[i]).setId('formula-'+i).setStyleAttribute('opacity', '0.5');
      grid.setWidget(2, i+1, formulaLabel); 
      asValuesCheckBoxes[i] = app.createCheckBox().setId('asValues-'+sheetId+'-'+i).setName('av-'+sheetId+'-'+i).setEnabled(false).setValue(false);
      grid.setWidget(3, i+1, asValuesCheckBoxes[i]);       
      onButtonHandlers[i] = app.createClientHandler().forTargets(onButtons[i]).setVisible(false).forTargets(offButtons[i]).setVisible(true).forTargets(asValuesCheckBoxes[i]).setEnabled(true).forTargets(buttonValues[i]).setText(i+'-on');
      onButtonServerHandlers[i] = app.createServerHandler('toggleOpacity').addCallbackElement(panel);
      offButtonHandlers[i] = app.createClientHandler().forTargets(offButtons[i]).setVisible(false).forTargets(onButtons[i]).setVisible(true).forTargets(asValuesCheckBoxes[i]).setEnabled(false).setValue(false).forTargets(buttonValues[i]).setText(i+'-off');
      offButtonServerHandlers[i] = app.createServerHandler('toggleOpacity').addCallbackElement(panel);
      onButtons[i].addClickHandler(onButtonHandlers[i]).addClickHandler(onButtonServerHandlers[i]);
      offButtons[i].addClickHandler(offButtonHandlers[i]).addClickHandler(offButtonServerHandlers[i]);
      if (colValues['col-'+(i+1)]) {
        onButtons[i].setVisible(false);
        offButtons[i].setVisible(true);
        buttonValues[i].setText(i+'-on');
        formulaLabel.setStyleAttribute('opacity','1');
        asValuesCheckBoxes[i].setEnabled(true);
        if (colValues['col-'+(i+1)]=='1') {
          asValuesCheckBoxes[i].setValue(true);
        }
      }
    }
    scrollPanel.add(grid);
    panel.add(scrollPanel);
    var saveHandler = app.createServerHandler('manualSave').addCallbackElement(panel);
    var saveClientHandler = app.createClientHandler().forTargets(spinner).setVisible(true).forTargets(panel).setStyleAttribute('opacity', '0.5');
    var button = app.createButton("Save settings").setId('button').addClickHandler(saveHandler).addClickHandler(saveClientHandler);
    panel.add(button);
  } else {
    noDataLabel.setVisible(true);
    var saveHandler = app.createServerHandler('manualSave').addCallbackElement(panel);
    var button = app.createButton("Save settings").setId('button').addClickHandler(saveHandler).setVisible(false);
    panel.add(button);
  }
  return app;
}



function copyDown_refreshSheetUi(sheet, properties) {
  var app = UiApp.getActiveApplication().setWidth(700);
  var panel = app.getElementById('panel');
  panel.setStyleAttribute('opacity','1');
  var spinner = app.getElementById('spinner');
  spinner.setVisible(false);
  var scrollPanel = app.createScrollPanel().setWidth("660px");
  var grid = app.getElementById('grid');
  var sheetId = sheet.getSheetId();
  var sheetProperties = properties['sheet-'+sheetId];
  if (sheetProperties) {
    sheetProperties = Utilities.jsonParse(sheetProperties);
  } else {
    sheetProperties = new Object();
  }
  if (sheetProperties['formulaRow']) {
    var row = parseInt(sheetProperties['formulaRow']);
  } else {
    var row = 2;
  }
  var rowsArray = [2,3,4];
  var copyCol = sheetProperties['copyCol'];
  if (!copyCol) {
    copyCol = "A";
  }
  var hiddenSheetIdBox = app.getElementById('sheetId').setValue(sheetId);
  var columns = sheet.getLastColumn();
  var hiddenColNumBox = app.getElementById('numCols').setValue(columns);
  var noDataLabel = app.getElementById('noDataLabel');
  var horizontalPanel = app.getElementById('horizPanel');
  var button = app.getElementById('button');
  if (columns>0) {
    var colValues = sheetProperties['colValues'];
    if (colValues) {
      colValues = Utilities.jsonParse(colValues);
    } else {
      colValues = new Object();
    }
    grid.resize(4, columns+1);
    horizontalPanel.setVisible(true);
    grid.setVisible(true);
    noDataLabel.setVisible(false);
    var formulaRowSelect = app.getElementById('formulaRowSelect');
    var rowSelectIndex = rowsArray.indexOf(row);
    formulaRowSelect.setSelectedIndex(rowSelectIndex);
    var rowSelectHandler = app.createServerHandler('copyDown_refreshRowFormulas').addCallbackElement(panel);
    grid.setBorderWidth(1).setCellSpacing(0).setStyleAttribute('borderColor','#E5E5E5');
    formulaRowSelect.addChangeHandler(rowSelectHandler);
    var copyColSelect = app.getElementById('copyColSelect');
    copyColSelect.clear();
    for (var i=0; i<columns; i++) {
      copyColSelect.addItem(this.alphabet[i]);
    }
    var colSelectIndex = this.alphabet.indexOf(copyCol);
    if (colSelectIndex!=-1) {
      copyColSelect.setSelectedIndex(colSelectIndex);
    }
    var headers = sheet.getRange(1,1,1,columns).getValues()[0];
    grid.setWidget(0, 0, app.createLabel("Column"));
    grid.setWidget(1, 0, app.createLabel("Header"));  
    grid.setWidget(2, 0, app.createLabel("Value/Formula"));  
    grid.setWidget(3, 0, app.createLabel("Paste as values?"));  
    var onButtonHandlers = [];
    var onButtonServerHandlers = [];
    var offButtonHandlers = [];
    var offButtonServerHandlers = [];
    var onButtons = [];
    var offButtons = [];
    var buttonValues = [];
    var formulaLabels = [];
    var asValuesCheckBoxes = [];
    for (var i=0; i<columns; i++) {
      onButtons[i] = app.createButton(this.alphabet[i]).setId('onButton-'+sheetId+'-'+i).setStyleAttribute('background', 'whiteSmoke').setWidth("50px");
      offButtons[i] = app.createButton(this.alphabet[i]).setId('offButton-'+sheetId+'-'+i).setStyleAttribute('background', '#E5E5E5').setStyleAttribute('border', '2px solid grey').setWidth("50px").setVisible(false);
      buttonValues[i] = app.createTextBox().setVisible(false).setText(i+'-off').setName('bv-'+sheetId+'-'+i);
      var buttonPanel = app.createHorizontalPanel();
      buttonPanel.add(onButtons[i])
                 .add(offButtons[i])
                 .add(buttonValues[i])
                 .setStyleAttribute('width',"80px")
                 .setHorizontalAlignment(UiApp.HorizontalAlignment.CENTER);
      grid.setWidget(0, i+1, buttonPanel).setStyleAttribute(0, i+1, 'backgroundColor', 'whiteSmoke').setStyleAttribute(0, i+1, 'textAlign', 'center');
      grid.setWidget(1, i+1, app.createLabel(headers[i]));
      var formulas = copyDown_getSheetFormulas(sheet, row);
      var formulaLabel = app.createLabel(formulas[i]).setId('formula-'+i).setStyleAttribute('opacity', '0.5').setWidth("100px");
      grid.setWidget(2, i+1, formulaLabel); 
      asValuesCheckBoxes[i] = app.createCheckBox().setId('asValues-'+sheetId+'-'+i).setName('av-'+sheetId+'-'+i).setEnabled(false).setValue(false);
      grid.setWidget(3, i+1, asValuesCheckBoxes[i]);       
      onButtonHandlers[i] = app.createClientHandler().forTargets(onButtons[i]).setVisible(false).forTargets(offButtons[i]).setVisible(true).forTargets(asValuesCheckBoxes[i]).setEnabled(true).forTargets(buttonValues[i]).setText(i+'-on');
      onButtonServerHandlers[i] = app.createServerHandler('toggleOpacity').addCallbackElement(panel);
      offButtonHandlers[i] = app.createClientHandler().forTargets(offButtons[i]).setVisible(false).forTargets(onButtons[i]).setVisible(true).forTargets(asValuesCheckBoxes[i]).setEnabled(false).setValue(false).forTargets(buttonValues[i]).setText(i+'-off');
      offButtonServerHandlers[i] = app.createServerHandler('toggleOpacity').addCallbackElement(panel);
      onButtons[i].addClickHandler(onButtonHandlers[i]).addClickHandler(onButtonServerHandlers[i]);
      offButtons[i].addClickHandler(offButtonHandlers[i]).addClickHandler(offButtonServerHandlers[i]);
      if (colValues['col-'+(i+1)]) {
        onButtons[i].setVisible(false);
        offButtons[i].setVisible(true);
        buttonValues[i].setText(i+'-on');
        formulaLabel.setStyleAttribute('opacity','1');
        asValuesCheckBoxes[i].setEnabled(true);
        if (colValues['col-'+(i+1)]=='1') {
          asValuesCheckBoxes[i].setValue(true);
        }
      }
    }
    button.setVisible(true);
  } else {
    horizontalPanel.setVisible(false);
    grid.clear().setVisible(false);
    noDataLabel.setVisible(true);
    button.setVisible(false);
  }
  return app;
}


function saveCopyDownSettings(e) {
  var app = UiApp.getActiveApplication();
  var properties = ScriptProperties.getProperties();
  var sheetId = e.parameter.sheetId;
  var sheet = copyDown_getSheetFromId(sheetId);
  if (properties['sheet-'+sheetId]) {
    var sheetProperties = Utilities.jsonParse(properties['sheet-'+sheetId]);
  } else {
    sheetProperties = new Object();
  }
  copyDown_clearSheetComments(sheet, sheetProperties);
  var colProperties = new Object();
  var numCols = e.parameter.numCols;
  var headerString = fetchHeaderString(sheet);
  sheetProperties['headerString'] = headerString;
  var formulaRow = e.parameter['formulaRow']
    sheetProperties["formulaRow"] = formulaRow;
    sheetProperties["copyCol"] = e.parameter['copyCol'];
    for (var j=0; j<numCols; j++) {
      var buttonValue = e.parameter['bv-'+sheetId+'-'+j];
      if (buttonValue==j+'-on') {
        var asValuesOption = e.parameter['av-'+sheetId+'-'+j];
        if (asValuesOption == "false") {
          colProperties["col-" + (j+1)] = 2;
        } else {
          colProperties["col-" + (j+1)] = 1; 
        }
      } 
    }
  if (colProperties) {
    sheetProperties['colValues'] = Utilities.jsonStringify(colProperties);
  }
  if (sheetProperties) {
    properties['sheet-'+sheetId] = Utilities.jsonStringify(sheetProperties);
  }
  ScriptProperties.setProperties(properties);
  copyDown_setSheetComments(sheet, sheetProperties);
  return app;
}

function fetchHeaderString(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var headerString = [];
  for (var i=0; i<headers.length; i++) {
    headers[i] = String(headers[i]);
    headerString += headers[i].replace(/\s+/g, ' ');
  }
  return headerString;
}


function validateHeaderStrings(sheet, properties) {
  var sheetId = sheet.getSheetId();
  var sheetProperties = properties['sheet-'+sheetId];
  if (sheetProperties) {
    sheetProperties = Utilities.jsonParse(sheetProperties);
    var colValues = sheetProperties.colValues;
    if (colValues) {
      colValues = Utilities.jsonParse(colValues);
      var count=0;
      for (var key in colValues) {
        count++;
      }
      if (count==0) {
        return true;
      }
      var savedHeaderString = sheetProperties.headerString;
      var currentHeaderString = fetchHeaderString(sheet);
      if (savedHeaderString != currentHeaderString) {
        return false;
      }
      return true;
    }
    return true;
  }
  return true;
}



function copyDown_clearSheetComments(sheet, sheetProperties) {
  if (sheetProperties.colValues) {
    var colProperties = Utilities.jsonParse(sheetProperties.colValues);
  }
  var row = parseInt(sheetProperties.formulaRow);
  if (!colProperties) {
    return;
  } else {
    sheet.getRange(row, 1, 1, sheet.getLastColumn()).clearComment(); 
  }
}


function copyDown_setSheetComments(sheet, sheetProperties) {
  if (sheetProperties.colValues) {
    var colProperties = Utilities.jsonParse(sheetProperties.colValues);
  }
  var copyCol = sheetProperties.copyCol;
  var row = parseInt(sheetProperties.formulaRow);
  if (!colProperties) {
    return;
  } else {
  for (var key in colProperties) {
      var col = parseInt(key.split('-')[1]);
      var copyDownOption = colProperties[key];
      var range = sheet.getRange(row, col);
      var formula = range.getFormula();
      var type = 'formula';
      if (formula=='') {
        type = 'value';
      }
    if ((copyDownOption=="1")&&(type=='formula')) {
        range.setComment('This cell\'s ' + type + ' and formats will be copied down to the last row in column ' + copyCol + ' AS VALUES when copydown script runs');
    } else {
      sheet.getRange(row, col).setComment('This cell\'s ' + type + ' and formats will be copied down to the last row of column ' + copyCol + ' when copydown script runs');
    }
    }
  }
}

function waitingIcon() {
  var app = UiApp.createApplication().setHeight(250).setWidth(200);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var waitingImageUrl = this.COPYDOWNIMAGEURL;
  var image = app.createImage(waitingImageUrl).setWidth("125px").setStyleAttribute('marginLeft', '25px');
  app.add(image);
  app.add(app.createLabel('Please be patient as copydown formulas are recalculated and pasted down their designated columns. For complex spreadsheets this can take some time.'));
  ss.show(app);
  return app;
}

function closeIcon(e) {
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}


function runCopyDown() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = ScriptProperties.getProperties();
  var lastSheetName = ScriptProperties.getProperty('lastSheetName');
  var startTime = new Date();
  startTime = startTime.getTime();
  waitingIcon();
  var app = UiApp.createApplication();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheetNames = [];
  for (var i=0; i<sheets.length; i++) {
    sheetNames.push(sheets[i].getName());
  }
  var startIndex = 0;
  var lastCol = '';
  if ((lastSheetName)&&(lastSheetName != "completed")) {
   startIndex = sheetNames.indexOf(lastSheetName);
   lastCol = ScriptProperties.getProperty('lastCol');
  }
  for (var i=startIndex; i<sheets.length; i++) {
    var sheetId = sheets[i].getSheetId();
    var sheetProperties = properties['sheet-'+sheetId];
    if (sheetProperties) {
      sheetProperties = Utilities.jsonParse(sheetProperties);
      var colValues = sheetProperties.colValues;
      if (colValues!="{}") {
        var headerCheck = validateHeaderStrings(sheets[i], properties);
        if (headerCheck!=true) {
          ss.setActiveSheet(sheets[i]);
          Browser.msgBox("Warning: copyDown discovered a change in the headers of this sheet.  To avoid accidentally over-writing your data, please confirm and re-save your settings before running.");
          copyDown_createJob();
          app.close();
          return app;
        }
      var timeout = runSheetCopyDown(sheets[i], startTime, lastCol, properties);
      }
    }
  }
  ScriptProperties.setProperty('lastSheetName', 'completed');
  if (timeout!=true) {
    copyDown_logCopyDown();
    Browser.msgBox("copyDown completed successfully");
  }
  return app;
}

function copyDown_advanced() {
  var ss = SpreadsheetApp.getActive();
  var app = UiApp.createApplication().setTitle("Advanced options").setHeight(130).setWidth(290);
  var quitHandler = app.createServerHandler('copyDown_quitUi');
  var handler2 = app.createServerHandler('copyDown_extractorWindow');
  var button2 = app.createButton('Package copyDown settings for distribution').addClickHandler(quitHandler).addClickHandler(handler2);
  var handler3 = app.createServerHandler('copyDown_institutionalTrackingUi');
  var button3 = app.createButton('Manage your usage tracker settings').addClickHandler(quitHandler).addClickHandler(handler3);
  var panel = app.createVerticalPanel();
  panel.add(button2);
  panel.add(button3);
  app.add(panel);
  ss.show(app);
  return app;
}

function copyDown_quitUi(e) {
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}

function recoverCopyDown () {
 var allTriggers = ScriptApp.getScriptTriggers();
 for (var i=0; i<allTriggers.length; i++) {
   if (allTriggers[i].getHandlerFunction()=="recoverCopyDown") {
     ScriptApp.deleteTrigger(allTriggers[i]);
  }
 }
 runCopyDown();
}



function runSheetCopyDown(sheet, startTime, lastCol, properties) {
  var sheetId = sheet.getSheetId();
  var sheetProperties = properties['sheet-'+sheetId];
  if (sheetProperties) {
    sheetProperties = Utilities.jsonParse(sheetProperties);
  } else {
    return;
  }
  var copyCol = sheetProperties.copyCol;
  var copyColNum = this.alphabet.indexOf(copyCol) + 1;
  var colValues = sheetProperties.colValues;
  if (colValues) {
    colValues = Utilities.jsonParse(colValues);
  } else {
    colValues = new Object();
  }
  var keys = [];
  for (var key in colValues) {
    keys.push(key);
  }
  keys.sort(function(a,b){return parseInt(a.split('-')[1])-parseInt(b.split('-')[1])});
  var nextCol = 0;
  if (lastCol) {
    nextCol = keys.indexOf(lastCol) + 1;
  }
  if ((sheet.getLastRow()>0)&&(keys.length>0)) {
  var copyColValues = sheet.getRange(1, copyColNum, sheet.getLastRow(), 1).getValues();
  for (var i = nextCol; i<copyColValues.length; i++) {
    if (i>5) {
      if ((copyColValues[i][0]=='')&&(copyColValues[i-1][0]=='')&&(copyColValues[i-2][0]=='')&&(copyColValues[i-3][0]=='')&&(copyColValues[i-4][0]=='')) {
        var lastCopyRow = i-4;
        break;
      }
    } else {
      var lastCopyRow = sheet.getLastRow();
    }
  }
    var lastRow = sheet.getLastRow();
  var formulaRow = parseInt(sheetProperties.formulaRow);
  if ((lastRow-formulaRow>=1)&&(lastCopyRow-formulaRow>=1)) {
  for (var i=0; i<keys.length; i++) {
    var col = keys[i].split("-")[1];
    var masterCellFormula = sheet.getRange(formulaRow, col);
    var comment = masterCellFormula.getComment();
    masterCellFormula.clearComment();
    var clearRange = sheet.getRange(formulaRow+1,col, lastRow-formulaRow, 1).clear();
    var destRange = sheet.getRange(formulaRow+1, col, lastCopyRow-formulaRow, 1);
    masterCellFormula.copyTo(destRange);
    if (colValues[keys[i]]=="1") {
      var destValues = sheet.getRange(formulaRow+1, col, lastCopyRow-formulaRow, 1).getValues();
      destRange.setValues(destValues);
    }
    masterCellFormula.setComment(comment);
    masterCellFormula.copyFormatToRange(sheet, col, 1, formulaRow+1, lastCopyRow-formulaRow);
    var timeNow = new Date();
    timeNow = timeNow.getTime();
    var timeDiff = (timeNow - startTime)/1000;
    if (timeDiff > 250) {
      ScriptProperties.setProperty('lastSheetName', sheet.getName());
      ScriptProperties.setProperty('lastCol', keys[i]);
      var newDateObj = new Date(timeNow + (0.5)*60000);
      ScriptApp.newTrigger('recoverCopyDown').timeBased().at(newDateObj).create();
      Browser.msgBox("The copyDown script will resume in 30 seconds to avoid timeout.");
      return true;
    }
    }
   }
  } 
}

function copyDown_getSheetIds() {
  var sheets = ss.getSheets();
  var sheetIds = [];
  for (var i=0; i<sheets.length; i++) {
    sheetIds.push(sheets[i].getSheetId());
  }
  return sheetIds;
}

function copyDown_refreshRowFormulas(e) {
  var properties = ScriptProperties.getProperties();
  var sheetId = e.parameter.sheetId;
  var sheetProperties = properties['sheet-'+sheetId];
  if (sheetProperties) {
    sheetProperties = Utilities.jsonParse(sheetProperties);
  } else {
    sheetProperties = new Object();
  }
  var colValues = sheetProperties['colValues'];
  if (colValues) {
    colValues = Utilities.jsonParse(colValues);
  } else {
    colValues = new Object();
  }
  var app = UiApp.getActiveApplication();
  var grid = app.getElementById('grid');
  var row = e.parameter["formulaRow"];
  var sheet = copyDown_getSheetFromId(sheetId);
  var formulas = copyDown_getSheetFormulas(sheet, row);
  
  for (var i=0; i<formulas.length; i++) {
    var formulaLabel = app.createLabel(formulas[i]).setId('formula-'+i).setStyleAttribute('opacity', '0.35');
    if (colValues['col-'+(i+1)]) {
      formulaLabel.setStyleAttribute('opacity', '1');
    }
    grid.setWidget(2, i+1, formulaLabel); 
  }
  
  
  return app;
}


function toggleOpacity(e) {
  var app = UiApp.getActiveApplication();
  var sheetId = e.parameter.sheetId;
  var numCols = e.parameter.numCols;
  for (var i=0; i<numCols; i++ ) {
  var buttonValue = e.parameter['bv-'+sheetId+'-'+i];
  if (buttonValue) {
    buttonValue = buttonValue.split("-");
  }
  var label = app.getElementById('formula-'+buttonValue[0]);
  if (buttonValue[1] == "on") {
    label.setStyleAttribute('opacity','1');
  } else {
    label.setStyleAttribute('opacity','0.5');
  }
}
  return app;
}

function copyDown_getSheetFromId(sheetId) { 
  var sheets = ss.getSheets();
  for (var i=0; i<sheets.length; i++) {
    if (sheets[i].getSheetId()==sheetId) {
      var sheet = sheets[i];
      break;
    }
  }
  return sheet;
}

function copyDown_getSheetFormulas(sheet, row) {
  var columns = sheet.getLastColumn();
  var range = sheet.getRange(row, 1, 1, columns)
  var formulas = range.getFormulas()[0];
  var values = range.getValues()[0];
  for (var i=0; i<formulas.length; i++) {
    if (formulas[i]=='') {
      if (typeof values[i] == 'string') {
        formulas[i]=values[i].substring(0,15);
      }  else {
        formulas[i]=values[i];
      }
    } else {
      formulas[i] = formulas[i].substring(0,15);
    }
    if (formulas[i].length==15) {
      formulas[i] += "...";
    }
  }
  return formulas;
}
  
  
