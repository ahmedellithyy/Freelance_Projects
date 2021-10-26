// Global Variables
var map = {}


function onOpen() {
  ui.createAddonMenu()
    .addItem("BOL Creditor Import txt Generator", "openDialog")
    .addToUi()
}

function openDialog() {
  clients = ui.prompt('Enter list of client IDs as a comma-seprated list: \n', ui.ButtonSet.OK).getResponseText()
  var clientIDs = clients.split(",")
  
  if(main(clientIDs)!=-1) return -1

}
function main(clientIDs){
  var today = Utilities.formatDate(dateTime, "GMT+2", dateTimeFormat)
  var fileName = outputFilename.replace(/{{dateTime}}/g,today)
  var outputFile = createFile(clientIDs)

  if(outputFile == -1){
    return;
  }
  var fileToDrive = Utilities.newBlob(outputFile, "text/txt", fileName)
  const file = DriveApp.getFolderById(outputFolderId).createFile(fileToDrive);
  
}

function createFile(clientIDs){
  if (getRowsForClients(clientIDs)==-1) return -1;

  var txtFile = ""
  for(id of clientIDs)
  {
    for(key in accountTypes)
    {
      var row = createRow(id,key)
      if(row != -1){
        txtFile = txtFile + row +"\n"
      }
      else{
        return -1
      }
    }    
  }
  if(forceUpperCase == 1){
    txtFile = txtFile.toUpperCase()
  }
  Logger.log(txtFile)

  return txtFile
}


function createRow(clientID,accountType){
  var row = ""
  for(let i = 0; i < creditorFileTemplate.length;i++)
  {
    row = row + creditorFileTemplate[i].value + "|"
  }
  
  var accountMap = accountTypes[accountType]
  var entries = Object.entries(accountMap)
  for(let spreadsheet in map){
    for(let sheet in map[spreadsheet]){
      for(var k=0; k< map[spreadsheet][sheet]["sheetMap"].length-1;k++){ 
        if(row.includes(map[spreadsheet][sheet]["sheetMap"][k][0])){
          var regExp = new RegExp(`{{${map[spreadsheet][sheet]["sheetMap"][k][0]}}}`, "g")
          var columnIndex = -1
          columnIndex = map[spreadsheet][sheet]["sheetData"][0][map[spreadsheet][sheet]["headerRow"]-1].indexOf(map[spreadsheet][sheet]["sheetMap"][k][1])
          if(columnIndex == -1) {
            ui.alert("Mismatch of Columns Headers", `"${map[spreadsheet][sheet]["sheetMap"][k][1]}" column header is  written wrongly in the variable map.`, ui.ButtonSet.OK);
            return -1;
          }
          row = row.replace(regExp, map[spreadsheet][sheet]["sheetData"][getRowIndex(map[spreadsheet][sheet]["sheetData"],clientID)][columnIndex])
        }
        else{
          if(Object.values(accountMap).indexOf(map[spreadsheet][sheet]["sheetMap"][k][0]) > -1){
            var regExp = new RegExp(`{{${Object.keys(accountMap).find(key => accountMap[key] === map[spreadsheet][sheet]["sheetMap"][k][0])}}}`, "g")
            var columnIndex = -1
            columnIndex = map[spreadsheet][sheet]["sheetData"][0][map[spreadsheet][sheet]["headerRow"]-1].indexOf(map[spreadsheet][sheet]["sheetMap"][k][1])
            if(columnIndex == -1) {
              ui.alert("Mismatch of Columns Headers", `"${map[spreadsheet][sheet]["sheetMap"][k][1]}" column header is  written wrongly in the variable map.`, ui.ButtonSet.OK);
              return -1;
            }
            row = row.replace(regExp, map[spreadsheet][sheet]["sheetData"][getRowIndex(map[spreadsheet][sheet]["sheetData"],clientID)][columnIndex])
          }
        }
      }
      
    }
  }
  for(let i = 0; i < entries.length;i++)
  {
    if(row.includes(entries[i][0])){
      var regExp = new RegExp(`{{${entries[i][0]}}}`, "g")
      row = row.replace(regExp,entries[i][1])
    }
  }
  row = row.split("|")
  var final = ""
  for(let i = 0; i < creditorFileTemplate.length;i++)
  {
    if(creditorFileTemplate[i].type == "Number"){
      final = final + String(row[i]).padStart(creditorFileTemplate[i].len, '0')
    }
    else if (creditorFileTemplate[i].type == "String"){
      let pad = Array(creditorFileTemplate[i].len).join(" ")
      final = final + (row[i] + pad).substring(0, pad.length)

    }
  }
  return final
}

function getRowIndex(data,clientID){
  for(var i =0; i< data.length;i++)
  {
    if(data[i][clientIdColumnIndex-1] == clientID)
    {
      return i
    }
  }
}

function getRowsForClients(clientIDs)
{
  var missingIDs = []
  var obj = {}
  for(i in variableMap){
    for(j in variableMap[i]){
      obj[j] = {
        "sheetData":[],
        "headerRow":variableMap[i][j]["headerRow"],
        "sheetMap": Object.entries(variableMap[i][j])
      }
    }
    map[i] = obj
    obj ={}
  }
  for(i in map)
  {
    var clientDateSpreadSheet = SpreadsheetApp.openById(i)
    for(j in map[i])
    {
      var clientDataSheet = getSheetById(clientDateSpreadSheet,j);
      // Get sheet headers
      map[i][j]["sheetData"].push(clientDataSheet.getRange(1, 1, map[i][j]["headerRow"], clientDataSheet.getLastColumn()).getValues())
      // Get sheet data for the specific client IDs
      for(var k = 0; k < clientIDs.length;k++)
      {
        var row = clientDataSheet.createTextFinder(clientIDs[k]).matchEntireCell(true).findAll().map(x => x.getRow())
        if(row.length == 0){
          if(!missingIDs.includes(clientIDs[k])) missingIDs.push(clientIDs[k]);
        }
        else{
          map[i][j]["sheetData"].push(clientDataSheet.getRange(row, 1, 1, clientDataSheet.getLastColumn()).getValues()[0])
        }
      }
    }
  }
  if(missingIDs.length > 0){
    ui.alert(`These client IDs "${missingIDs}" don't exist.`, 'Please enter valid client IDs.', ui.ButtonSet.OK);
    return -1;
  }
  if(missingIDs.length == clientIDs.length){
    return -1;
  }
}






function getSheetById(spreadSheet,sheetID) {
  return spreadSheet.getSheets().filter(function(sheet) {return sheet.getSheetId() == sheetID})[0]
}
