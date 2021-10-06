const ui = SpreadsheetApp.getUi();


// Wrtie the the number of column that has client ID values but decrement it by 1, like if it's in column A so index will be 0
const clientIdColumnIndex = 0


// File ID for CSV template on Drive
const csvTemplateFileId = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
 
// Date format for output file
const dateFormat = "yyyyMMdd"
 
// Map of variable names in templace file to corresponding data columns in the trade sheet
const variableMap = {
 "1No34Grg88R6d5fZ0joUJUzmTUsGrov8ENXFC3S_3Fhg": {
   "2026891114": {
     "{{batchId}}": "Batch Trade ID",
     "{{dealRef}}": "Deal Reference",
     "{{tradedCapital}}": "FX Amount Traded\n(ZAR)",
     "{{netClientProfit}}": "Net Client Profit\n(ZAR)",
     "{{paymentMethod}}": "Payment\nMethod",
     "{{date}}": Utilities.formatDate(new Date(), "GMT+2", dateFormat),
     "headerRow": 1
   }
 },
 "19tInAXiaWe7MM5D6j6sRWOG_VCfMyee32E__W-GMR6c": 
 {
    "0": 
    {
     "{{firstName}}": "First Name",
     "{{lastName}}": "Last Name",
     "{{email}}": "Email Address",
     "headerRow": 1
    },
   "59065364": 
   {
     "{{mercantileAccountNumber}}": "Mercantile Bank\nAccount No.",
     "{{mercantileBranchCode}}": "Mercantile Bank\nBranch Code",
     "{{personalAccountNumber}}": "Personal Bank\nAccount No.",
     "{{personalBranchCode}}": "Personal Bank\nBranch Code",
     "headerRow": 2
   }
 }
}


// ID of folder in which to save the exported documents
const folderId = "xxxxxxxxxxxxxxxxxxxxxxx"
 
// ID of folder in which to save the exported documents
const outputFilename = "{{batchId}} - {{date}}"
