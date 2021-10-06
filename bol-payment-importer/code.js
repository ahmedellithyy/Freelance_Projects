// Global Variables
var spreadsheetIDs
var startingSpreadSheet
var startingSheet
var startingData
var startingSheetHeaderRows
var startingMap
var map = {}
var batchNo

function createCsv(batchNo) {
  spreadsheetIDs = Object.keys(variableMap)
  const templateSheet = DriveApp.getFileById(csvTemplateFileId).getAs("text/csv").getDataAsString()
  const templateCsv = Utilities.parseCsv(templateSheet)
  const csvHeaders = templateCsv.shift()
  let csvFile = csvHeaders + "\n"
  const templateRow2String = templateCsv[0].join(",")
  const templateRow3String = templateCsv[1].join(",")
  const templateRow4String = templateCsv[2].join(",")
  
  var batchIdFound = false
  loop:
  for(let i =0; i<spreadsheetIDs.length;i++){
    var sheetsIDs = Object.keys(variableMap[spreadsheetIDs[i]])
    for(let j=0;j<sheetsIDs.length;j++){
      var startingKeys = Object.keys(variableMap[spreadsheetIDs[i]][sheetsIDs[j]])
      for(let k =0;k<startingKeys.length;k++){
        if(startingKeys[k] =="{{batchId}}"){
          startingSpreadSheet = spreadsheetIDs[i]
          startingSheet = sheetsIDs[j]
          batchIdFound = true
          break loop
        }
      }
    }
  }
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



  if(batchIdFound == true){
    startingMap = Object.entries(variableMap[startingSpreadSheet][startingSheet])
    startingSheetHeaderRows = variableMap[startingSpreadSheet][startingSheet]["headerRow"];
    startingData = getRowsForBatchNo(batchNo,startingSpreadSheet,startingSheet)


    var paymentMethodColumnName;
    var paymentMethodColumnNumber;

    for(let j=0; j<startingMap.length;j++)
    {
        if(startingMap[j][0]=="{{paymentMethod}}"){
          paymentMethodColumnName = startingMap[j][1]
        }    
    }
    for(let i = 0; i <startingData[0].length;i++)
    {
      if(startingData[0][i]== paymentMethodColumnName){
        paymentMethodColumnNumber = i;
        break;
      }
    }

    for(let i = 1;i<startingData.length;i++)
    {
      if(startingData[i][paymentMethodColumnNumber]=="Cash")
      {
        var row = create_csvRow(templateRow2String,i,batchNo)
        if(row!= -1){
          csvFile += row + "\n"
          }
        else{return -1}
        row = create_csvRow(templateRow3String,i,batchNo)
        if(row != -1){csvFile += row + "\n"

        } else{return -1}
      }
      else{
        row = create_csvRow(templateRow2String,i,batchNo)
        if(row != -1){csvFile += row + "\n"
                 } else{return -1}
        row = create_csvRow(templateRow4String,i,batchNo)
        if(row != -1){csvFile += row + "\n"
} else{return -1}
      }

    }
    return csvFile;
  }
  else{
    ui.alert("'batchId' isn't found in variable map.", "Check variable map of the script.", ui.ButtonSet.OK);
    return -1 
  }
}

function create_csvRow(templateRowString,rowIndex,batchNo){
  getRowsForClients(startingData)

  for(i in map)
  {
    for(j in map[i])
    {
        for(var k=0; k<map[i][j]["sheetMap"].length-1;k++)
        { 
          var replacedData
          var regExp = new RegExp(`${map[i][j]["sheetMap"][k][0]}`, "g")

          if(j == startingSheet){
            if(map[i][j]["sheetMap"][k][0] =="{{date}}") {
              replacedData = variableMap[i][j]["{{date}}"]
            }
            else if(map[i][j]["sheetMap"][k][0]=="{{batchId}}"){
              replacedData = batchNo
            }
            else
            {
              var columnIndex = -1
              columnIndex = map[i][j]["sheetData"][0][map[i][j]["headerRow"]-1].indexOf(map[i][j]["sheetMap"][k][1])
              if(columnIndex == -1) {
                ui.alert("Mismatch of Columns Headers", `${map[i][j]["sheetMap"][k][1]} is  written wrongly in the variable map.`, ui.ButtonSet.OK);
                return -1;
              }
              replacedData = map[i][j]["sheetData"][rowIndex][columnIndex]
              if(Number(replacedData) === replacedData && replacedData % 1 !== 0)
              {
                replacedData = replacedData.toFixed(2)
              } 
            }
            templateRowString = templateRowString.replace(regExp, replacedData)
          }
          else
          {
              columnIndex = -1
              columnIndex = map[i][j]["sheetData"][0][map[i][j]["headerRow"]-1].indexOf(map[i][j]["sheetMap"][k][1])
              if(columnIndex == -1) {
                ui.alert("Mismatch of Columns Headers", `${map[i][j]["sheetMap"][k][1]} is  written wrongly in the variable map.`, ui.ButtonSet.OK);
                return -1;
              }
            templateRowString = templateRowString.replace(regExp, map[i][j]["sheetData"][getRowIndex(map[i][j]["sheetData"],startingData,rowIndex)][columnIndex])
          }
        }
      
      
    }
  }
    return templateRowString      
}


function getRowIndex(data,tradeData,index){
  for(var i =0; i< data.length;i++)
  {
    if(data[i][clientIdColumnIndex]==tradeData[index][clientIdColumnIndex])
    {
      return i
    }
  }
}

function getRowsForClients(startingData)
{
  for(i in map)
  {
    var clientDateSpreadSheet = SpreadsheetApp.openById(i)
    for(j in map[i])
    {
      var clientDataSheet = getSheetById(clientDateSpreadSheet,j);
      var clientDataSheetHeaderRows = map[i][j]["headerRow"];
      map[i][j]["sheetData"].push(clientDataSheet.getRange(1, 1, clientDataSheetHeaderRows, clientDataSheet.getLastColumn()).getValues())
      for(var k = startingSheetHeaderRows; k < startingData.length;k++)
      {
        var row = clientDataSheet.createTextFinder(startingData[k][clientIdColumnIndex]).matchEntireCell(true).findAll().map(x => x.getRow())
        row.forEach(
        function(row) 
        {
          if(i == startingSpreadSheet && j == startingSheet){
            if(clientDataSheet.getRange(row, 4, 1, clientDataSheet.getLastColumn()).getValue() == batchNo){
                map[i][j]["sheetData"].push(clientDataSheet.getRange(row, 1, 1, clientDataSheet.getLastColumn()).getValues()[0])
            }
          }
          else{
            map[i][j]["sheetData"].push(clientDataSheet.getRange(row, 1, 1, clientDataSheet.getLastColumn()).getValues()[0])
          }
        })
      }
    }
  }
  return;
}

function getRowsForBatchNo(batchNo,startingSpreadSheet,startingSheet) {

  const tradeSpreadSheet = SpreadsheetApp.openById(startingSpreadSheet)
  const tradesSheet = getSheetById(tradeSpreadSheet,startingSheet);
  
  const tradeRows = tradesSheet.createTextFinder(batchNo).matchEntireCell(true).findAll().map(x => x.getRow())
  if (tradeRows.length == 0) {
    ui.alert('Batch ID does not exist.', 'Please enter a valid batch ID.', ui.ButtonSet.OK);
  } 
  const rowData = [tradesSheet.getRange(1, 1, startingSheetHeaderRows, tradesSheet.getLastColumn()).getValues()[0]]
 
  tradeRows.forEach(
    function(row) 
    {
      rowData.push(tradesSheet.getRange(row, 1, 1, tradesSheet.getLastColumn()).getValues()[0])
    }
  )

  return rowData
}


function getSheetById(spreadSheet,sheetID) {
  return spreadSheet.getSheets().filter(function(sheet) {return sheet.getSheetId() == sheetID})[0]
}


function onOpen() {
  ui.createAddonMenu()
    .addItem("Bol payment importer", "openDialog")
    .addToUi()
}

function openDialog() {
  batchNo = ui.prompt('Enter Batch Number: \n', ui.ButtonSet.OK).getResponseText()
  if(isNaN(batchNo)){
    ui.alert('Batch must be a number.', 'Please enter a valid batch ID.', ui.ButtonSet.OK);
  }
  else{
    if(main(batchNo)!=-1) return -1
  }
}



function main(batchNo){
  var today = Utilities.formatDate(new Date(), "GMT+2", dateFormat)
  var fileName = outputFilename.replace(/{{batchId}}/g,batchNo).replace(/{{date}}/g,today)+".csv"
  var outputFile = createCsv(batchNo)
  if ( outputFile!= -1)
  {
     var fileToDrive = Utilities.newBlob(outputFile, "text/txt", fileName)
      const file = DriveApp.getFolderById(folderId).createFile(fileToDrive);

  }
  else{
    return -1
  }
  //Logger.log(createCsv(batchNo))
}
