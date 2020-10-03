//This Creates the menu at document opening
function onOpen(e) {
    SpreadsheetApp.getUi()
        .createMenu('GButcher')
        .addItem('Batch Delete Sheet', 'BatchDeleteUXLogic')
        .addSeparator()
        .addItem('credits', 'credits')
        .addToUi();
}

//Credits dialog
function credits() {
    var URL1 = 'https://www.linkedin.com/in/pierpaolocanini/'
    var URL2 = 'https://github.com/pierpaolo-canini/GButcher'
    var htmlOutput = HtmlService
        .createHtmlOutput('<p style="font-family:Roboto; "><a href="' + URL1 + '" target="_blank";>Pierpaolo Canini</a></br></br>Check for updates on my <a href="' + URL2 + '"target="_blank";>GitHub</a></p>')
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setWidth(260)
        .setHeight(100);
    var dl = SpreadsheetApp.getUi().showModelessDialog(htmlOutput, "Credits")
}

//Everything starts from here. This Checks if you have more than one sheet and if the selection sheet already exist
function BatchDeleteUXLogic() {
    var ui = SpreadsheetApp.getUi()
    var ss = SpreadsheetApp.getActiveSpreadsheet()
    var BDcklst = ss.getSheetByName("00 - BATCH DELETE CHECKLIST")
    if (ss.getSheets().length == 1) {
        ui.alert("YOU HAVE ONLY ONE SHEET, YOU CANT DELETE IT. AT LEAST ONE SHEET MUST ALWAYS EXIST")
    } else {
        if (BDcklst == null) {
            BatchDelete()
        } else {
            ss.deleteSheet(BDcklst);
            BatchDelete()
        }
    }
}

//This creates a new sheet with all the existing sheets names and place all the checkboxes to select them
function BatchDelete() {
    var ss = SpreadsheetApp.getActiveSpreadsheet()
    var allss = SpreadsheetApp.getActiveSpreadsheet().getSheets()
    var cks = ss.insertSheet().setName("00 - BATCH DELETE CHECKLIST")
    
    Logger.log(allss.length,allss.map(function(itm){return itm.getName()}),allss)

    cks.getRange("A2:A"+((allss.length)+1)).setValues(allss.map(function(sheet){return [sheet.getName()]}))
    cks.getRange("B2:B"+((allss.length)+1)).insertCheckboxes()

    cks.getRange("A1").setValue("CHECK THE SHEETS YOU WANT TO DELETE")
    cks.getRange("E1").setValue("CHECK WHEN YOU'RE DONE")
    cks.getRange("E2").insertCheckboxes()
    ss.setActiveSheet(cks)
}

//This waits for the "DONE" checkbox to be checked in order to erase all sheets that have been marked.
function onEdit(e) {
    var ui = SpreadsheetApp.getUi()
    var ss = SpreadsheetApp.getActiveSpreadsheet()
    var BDcklst = ss.getSheetByName("00 - BATCH DELETE CHECKLIST")
    if (BDcklst == null) {} else {
        var done = BDcklst.getRange("E2").getValue()
        if (done == true) {
            var lr = BDcklst.getDataRange().getLastRow()
            var nf = BDcklst.getRange("A2:A" + lr).getValues()
            var ckalltrue = BDcklst.getRange("B2:B" + lr).getValues().sort()
            if (ckalltrue[0][0] == true) {
                BDcklst.getRange("E2").setValue(false)
                ui.alert("AT LEAST ONE SHEET MUST EXIST OR YOU WILL CAUSE A BLACK HOLE")
            } else {
              
              var flatBoolValues = BDcklst.getRange("B2:B" + lr).getValues().map(function(val){return val[0]})
              var deleteIndices = flatBoolValues.map(function(bol,index){if(bol){return index}}).filter(e => typeof e === 'number')           
              var flatSheetNames = BDcklst.getRange("A2:A" + lr).getValues().map(function(val){return val[0]})
              var deleteSheets = deleteIndices.map(x=>flatSheetNames[x])
              
              deleteSheets.map(x=>ss.deleteSheet(ss.getSheetByName(x)))
              
              ss.deleteSheet(ss.getSheetByName("00 - BATCH DELETE CHECKLIST"))
              ui.alert("DONE")
            }
        }
    }
}