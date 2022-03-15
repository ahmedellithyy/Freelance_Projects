function dataExtraction(query)
{
  query = JSON.stringify({query: query})
  var requestOptions = {
    'method': 'post',
    'payload': query,
    'contentType':'application/json',
    'headers':{
      'access_token': "ACCESS_TOKEN"
    }
  };
  var response = UrlFetchApp.fetch(`https://gis-api.aiesec.org/graphql?access_token=${requestOptions["headers"]["access_token"]}`, requestOptions);
  var recievedDate = JSON.parse(response.getContentText());
  return recievedDate.data.allOpportunityApplication.data;
}

function approvalsUpdating(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("APDs")
  var startDate = "01/11/2021"
  var query = `query{allOpportunityApplication(\n\t\tfilters:\n\t\t{\n\t\t\topportunity_home_mc:1609\n\t\tdate_approved:{from:\"${startDate}\"}\n\t\t}\n  \n\tpage:1\n    per_page:4000\n\t  \n\t)\n\t{\n  \t\n\t\tdata{\n\t\t\tperson{\n\t\t\t\tid\n\t\t\t\tfull_name\n\t\t\t\t\n\t\t\t\temail\n\t\t\t\thome_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n\t\t\t}\n\t\t\topportunity{\n\t\t\t\tid\n\t\t\t}\n\t\t\tdate_approved\n\t\t\tslot{\n\t\t\t\tstart_date\n\t\t\t\tend_date\n\t\t\t}\n\t\t\thost_lc{\n\t\tid\n\t\tname\n\t\t\t}\n\t\t\t\n\t\t\t\n\t\t}\n\t\t\n  }\n}`
  var data = dataExtraction(query)
  var rows = []
    for(let i = 0; i < data.length ;i++){
      var searchingID = data[i].person.id +"_"+ data[i].opportunity.id
      var rowIndex = sheet.createTextFinder(`${searchingID}`).matchEntireCell(true).findAll().map(x => x.getRow())
      if(rowIndex.length == 0){
      //GmailApp.sendEmail("ahmed.ellithy4@aiesec.net","Invitation Letter of AIESEC in Egypt Form", "Greeting from AIESEC in Egypt!\n\nIn this mail you will find invitation letter for your internship. In case of any questions, please contact local committee.\nInvitation Link Form: https://docs.google.com/forms/d/e/1FAIpQLSe28BzFSfZng53hMcnjBNsbgxNo9sH5sTqUfUVpfvUw9Dd4Jg/viewform \n\nPlease do not reply to this mail, it's automatical.")
      rows.push([data[i].person.id +"_"+ data[i].opportunity.id, data[i].person.id, data[i].opportunity.id,data[i].person.full_name,data[i].person.email,data[i].date_approved !=null?data[i].date_approved:"-",data[i].slot.start_date,data[i].slot.end_date,data[i].person.home_lc.name,data[i].person.home_mc.name,data[i].host_lc.name])
    }
  }
  if(rows.length>0)sheet.getRange(sheet.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows)
}

function sendIL(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RESPONSES")
  const approvalSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("APDs")
  const referenceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reference")
  var template = DocumentApp.openById("1x_qmEGfNba6crnW5xh3YmxeNLUqNVQ48S2P2hkRIi6U")
  var data = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues()
  var folder = DriveApp.getFolderById("1fU-UYIuHhXnkDJU8_7_Bdquo9mH6rd_X")

  for(let i = 1; i< data.length; i++)
  {
    if(data[i][22] != true && data[i][1]!= "")
    { 
      try{
        var row = approvalSheet.createTextFinder(`${data[i][18]}`).matchEntireCell(true).findAll().map(x => x.getRow())
        row.push("0")
        if(row.length>0){  
          var title = template.getName().toString().replace("{{full_name}}",`${data[i][1]}`)
          var invitationLetter = DocumentApp.create(title)
          Logger.log(invitationLetter.getUrl())
          sheet.getRange(i+2,23).setValue(true)
          var invitationLetterBody = invitationLetter.getBody()
          var total_items = template.getBody().getNumChildren()
          for(let i = 0;i< total_items;i++){
            switch(template.getBody().getChild(i).getType()){
              case DocumentApp.ElementType.PARAGRAPH: invitationLetterBody.appendParagraph(template.getBody().getChild(i).copy());
                break;
              case DocumentApp.ElementType.INLINE_IMAGE: invitationLetterBody.appendImage(template.getBody().getChild(i).copy());
                break;
            }
          }
          DriveApp.getFileById(invitationLetter.getId()).moveTo(folder)
          sheet.getRange(i+2,24).setValue(invitationLetter.getId())
          var today = Utilities.formatDate(new Date(),"GMT+2","dd-MM-yyyy")
          invitationLetterBody.replaceText("{{date_of_creation}}",today)
          var numOfIL = referenceSheet.getRange(referenceSheet.createTextFinder(data[i][19]).matchEntireCell(true).findAll().map(x => x.getRow()),2).getValue()
          var id =  String(lc_codes[`${data[i][19]}`]) + String(numOfIL+1).padStart(4, '0')
          invitationLetterBody.replaceText("{{IL_ID}}",id)
          referenceSheet.getRange(referenceSheet.createTextFinder(data[i][19]).matchEntireCell(true).findAll().map(x => x.getRow()),2).setValue(numOfIL+1)
          sheet.getRange(i+2,21).setValue(id)
          invitationLetterBody.replaceText("{{full_name}}",data[i][1])
          if(data[i][3].getTimezoneOffset() <0){
            invitationLetterBody.replaceText("{{b_day}}",Utilities.formatDate(data[i][3],`GMT+${data[i][3].getTimezoneOffset()/-60}`,"dd/MM/yyyy"))
          }
          else if(data[i][3].getTimezoneOffset() >0){
            invitationLetterBody.replaceText("{{b_day}}",Utilities.formatDate(data[i][3],`GMT-${data[i][3].getTimezoneOffset()/-60}`,"dd/MM/yyyy"))
          }
          else{
            invitationLetterBody.replaceText("{{b_day}}",Utilities.formatDate(data[i][3],`GMT`,"dd/MM/yyyy"))
          }
          invitationLetterBody.replaceText("{{place_of_birth}}",data[i][6])
          invitationLetterBody.replaceText("{{eng_citizenship}}",data[i][4])
          invitationLetterBody.replaceText("{{passport_id}}",data[i][9])
          invitationLetterBody.replaceText("{{date_of_issue}}",Utilities.formatDate(data[i][7],"GMT+2","dd/MM/yyyy"))
          invitationLetterBody.replaceText("{{expire_date}}",Utilities.formatDate(data[i][8],"GMT+2","dd/MM/yyyy"))
          invitationLetterBody.replaceText("{{living_address}}",data[i][5])
          invitationLetterBody.replaceText("{{project_city}}",data[i][13])
          invitationLetterBody.replaceText("{{project_start_date}}",Utilities.formatDate(data[i][14],"GMT+2","dd/MM/yyyy"))
          invitationLetterBody.replaceText("{{project_end_date}}",Utilities.formatDate(data[i][15],"GMT+2","dd/MM/yyyy"))
          invitationLetter.saveAndClose()
          var folder = DriveApp.getFolderById("1fU-UYIuHhXnkDJU8_7_Bdquo9mH6rd_X")
          var file = DriveApp.getFileById(invitationLetter.getId())
          var docblob = file.getBlob().getAs('application/pdf')
          var newDoc = DriveApp.createFile(docblob)
          newDoc.setName(title)
          newDoc.moveTo(folder)
          newDoc.setSharing(DriveApp.Access.ANYONE,DriveApp.Permission.VIEW)
          sheet.getRange(i+2,22).setValue(newDoc.getUrl())
          sheet.getRange(i+2,25).setValue(true)
          var message = `Hello ${data[i][1]},\nGreeting from AIESEC in Egypt!\n\nIn this mail you will find invitation letter for your internship. In case of any questions, please contact local committee. You can also download it to print it: ${newDoc.getUrl()}.\n\n\nPlease do not reply to this mail, it's automatical.`
          //MailApp.sendEmail(`${data[i][11]}`, "Invitation Letter from AIESEC in Egypt",message);
          sheet.getRange(i+2,26).setValue(true)
        }
      } 
      catch(e){
          Logger.log(e.toString())
          if(e.toString().includes("Invalid email")){
            var message = `Hello ${data[i][1]},\nGreeting from AIESEC in Egypt!\n\nIn this mail you will find invitation letter for your internship. In case of any questions, please contact local committee. You can also download it to print it: ${newDoc.getUrl()}.\n\n\nPlease do not reply to this mail, it's automatical.`
          MailApp.sendEmail(`${data[i][17]}`, "Invitation Letter from AIESEC in Egypt",message);
          sheet.getRange(i+2,26).setValue(true)
          }
      } 
      
    }
  }
 
  Logger.log("done") 
}

const lc_codes = {
     "6th October":	"2820",
     "AAST-Alex": "1788",
     "AAST-Cairo":	"1322",
     "Ain Shams University":	"1789",
     "Alexandria":	"0899",
     "AUC":	"1489",
     "Beni Suef": "2126",
     "Cairo University": "1064",
     "GUC":	"0257",
     "Helwan":	"2124",
      "KSU":	"2524",
      "Luxor & Aswan":	"2114",
      "Mansoura":	"0171",
     "Menofia":	"1727",
     "MIU" :	"2125",
     "MSA":	"2817",
     "MUST":	"2818",
     "Suez":	"0015",
     "Tanta": "1725",
     "Zagazig":"1114",
     "Damietta":	"0109"

}



