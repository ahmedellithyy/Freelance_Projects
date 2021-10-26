const ui = SpreadsheetApp.getUi();

const clientIdColumnIndex = 1


// format to use for timestamps for filename creation
const dateTimeFormat = "yyyy-mm-dd hh:mm"
 
const dateTime = new Date()
 
// toggle formatting text as UPPERCASE
const forceUpperCase = 1
 
// creditor file data template metadata definition 
const creditorFileTemplate = [
   {
      "type":"Number",
      "len":13,
      "value":"{{accountNumber}}",
      "description":"acct no"
   },
   {
      "type":"Number",
      "len":6,
      "value":"{{branchCode}}",
      "description":"branch code"
   },
   {
      "type":"String",
      "len":30,
      "value":"KRPT (PTY) LTD",
      "description":"statement ref"
   },
   {
      "type":"Number",
      "len":15,
      "value":10000000,
      "description":"creditor item limit"
   },
   {
      "type":"String",
      "len":16,
      "value":"{{clientId}}-{{accountTypeSuffix}}",
      "description":"creditor code"
   },
   {
      "type":"Number",
      "len":13,
      "value":0,
      "description":"CDI no"
   },
   {
      "type":"String",
      "len":30,
      "value":"{{lastName}}, {{firstName}} - {{accountTypeSuffix}}",
      "description":"acct name"
   }
]
 
// Map of bank account type metadata for matching with variableMap in creditor file data generation 
const accountTypes = {
   "personal":{
      "accountNumber":"personalAccountNumber",
      "branchCode":"personalBranchCode",
      "accountTypeSuffix": "PERS"
   },
   "mercantile":{
      "accountNumber":"mercantileAccountNumber",
      "branchCode":"mercantileBranchCode",
      "accountTypeSuffix": "MERC"
   }
}
 
// Map of variable names in templace file to corresponding data columns in the trade sheet
const variableMap = {
   "XXXXXXXXXXXXXXXX":{
      "0":{
         "clientId":"Client ID",
         "firstName":"First Name",
         "lastName":"Last Name",
         "headerRow":1
      },
      "59065364":{
         "mercantileAccountNumber":"Mercantile Bank\nAccount No.",
         "mercantileBranchCode":"Mercantile Bank\nBranch Code",
         "personalAccountNumber":"Personal Bank\nAccount No.",
         "personalBranchCode":"Personal Bank\nBranch Code",
         "headerRow":2
      }
   }
}
 
// ID of folder in which to save the exported documents
const outputFolderId = "XXXXXXXXXXXXXXXX"
 
// ID of folder in which to save the exported documents
const outputFilename = "bol_creditor_import_{{dateTime}}"
