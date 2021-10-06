var sheetName = "Sheet2"

function getMailBody() {  
  var titles = []
  var thread = GmailApp.search("from:pant3ra.strike@gmail.com",0,20);
  //var thread = GmailApp.search("from:info@send.bluethumb.com.au",0,20);
  //var thread = GmailApp.search("from:emails@songkick.com");
  var messages = GmailApp.getMessagesForThreads(thread);
  for (var i = 0 ; i < messages.length; i++) 
  {
    for (var j = 0; j < messages[i].length; j++)
    {
      var body = messages[i][j].getPlainBody()
      var regexToken = new RegExp("\\b((?:[a-z][\\w-]+:(?:/{1,3}|[a-z0-9%])|www\\d{0,3}[.]|[a-z0-9.\\-]+[.][a-z]{2,4}/)(?:[^\\s()<>]+|\\(([^\\s()<>]+|(\\([^\\s()<>]+\\)))*\\))+(?:\\(([^\\s()<>]+|(\\([^\\s()<>]+\\)))*\\)|[^\\s`!()\\[\\]{};:'\"\".,<>?«»“”‘’]))", "i");
      var url = body.match(regexToken)
      var urlData = UrlFetchApp.fetch(url[0]).getContentText();
      var title = urlData.match(/<title[^>]*>([^<]+)<\/title>/)[0];
      title = title.replace("<title>","").replace("</title>","")
     // Logger.log(title)
      titles.push([title])
    }
  }
  return titles
}

function sheetUpdating(){
  var titles = getMailBody()
  var newTitles = []
  var spreadsheet = SpreadsheetApp.openById("16xlY3HPay3SItBCalpKElEkS21EmpBMUEB7nSaGQjnE")
  var sheet = spreadsheet.getSheetByName(sheetName)
  var lr = sheet.getLastRow()
  var currentTitles = sheet.getRange(2,3,lr,1).getValues()
  for(var i of titles){
    var duplicated = false
    for(var j of currentTitles){
      if(i[0] == j){
        duplicated = true
      }
    }
    if(!duplicated){
      newTitles.push([i[0]])
    }
  }
  edit = sheet.getRange(lr+1,3,newTitles.length,1).setValues(newTitles)
}
