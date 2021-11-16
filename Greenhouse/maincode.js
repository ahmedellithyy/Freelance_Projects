var ui = SpreadsheetApp.getUi();
function onOpen() {  
  ui.createMenu('Functions')
      .addItem("List Offers", 'listAllOffers')
      .addItem("List Jobs", 'listAllJobs')
      .addItem("List Applications", 'listAllApplications')
      .addToUi();
}


function listAllApplications(){
  var createdAfterDate = ui.prompt('Enter the date that you want to extract applications after it (in this format MM/dd/yyyy): \n', ui.ButtonSet.OK).getResponseText()
  if (createdAfterDate == ""){
    ui.alert("Invalid Date", `You didn't enter any date, please re-run agaian and enter a date in this format MM/dd/yyyy`, ui.ButtonSet.OK);
    return
  }
  var date = new Date(createdAfterDate)
  var created_after = Utilities.formatDate(date,"GMT-4","yyyy-MM-dd'T'HH:mm:ss'Z'")


  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Applications")
  var options = { 
     "method" : "GET",
     "headers" : {
    "Authorization": ""
    }
  };

  var allData = []
  var page_num = 1 
  do{    
    var response = UrlFetchApp.fetch(`https://harvest.greenhouse.io/v1/applications?page=${page_num}&per_page=500&created_after=${created_after}`, options) 
    if(JSON.parse(response.getContentText()).length!= 0){
      allData.push(JSON.parse(response.getContentText()))
    }
    page_num++
  }while (JSON.parse(response.getContentText()).length != 0)

  var newRows = []
  var newRowIndex = 0
  for(var data of allData){
    for(var i = 0; i < data.length; i++){
      let duplicated = sheet.createTextFinder(data[i].id).matchEntireCell(true).findAll().map(x => x.getRow()).length == 0? true:false
      if(duplicated){
        Logger.log(i)
        newRows.push([
          data[i].status,data[i].source.public_name,data[i].source.id,
          
          data[i].rejection_reason != null? data[i].rejection_reason.type.name:"-",data[i].rejection_reason != null? data[i].rejection_reason.name:"-",

          data[i].rejection_details,data[i].rejected_at,data[i].prospective_office,data[i].prospective_department,data[i].prospect_detail.prospect_stage,data[i].prospect_detail.prospect_pool,
          data[i].prospect_detail.prospect_owner,data[i].prospect,
          data[i].location != null? data[i].location.address:"-",
          data[i].last_activity_at,data[i].jobs[0].name,data[i].jobs[0].id,data[i].job_post_id,data[i].id,

          data[i].current_stage != null ? data[i].current_stage.name:"-",
          data[i].credited_to != null? data[i].credited_to.id:"-",data[i].credited_to != null? data[i].credited_to.first_name:"-",data[i].credited_to != null?data[i].credited_to.last_name:"-", 
          data[i].credited_to != null? data[i].credited_to.employee_id:"-",
          
          data[i].candidate_id,
          data[i].attachments[0] != null? data[i].attachments[0].url:"-",data[i].attachments[0] != null? data[i].attachments[0].type:"-",
          data[i].attachments[0] != null && data[i].attachments[1] != null? data[i].attachments[1].url:"-",
          data[i].attachments[0] != null && data[i].attachments[1] != null? data[i].attachments[1].type:"-",
          
          data[i].applied_at,
          "-","-",    "-","-",    "-","-",    "-","-"
        ])
        if(data[i].answers[0] != null){
          for(let j = 0; j < data[i].answers.length;j++){
            newRows[newRowIndex][30+(j*2)] = data[i].answers[j].question
            newRows[newRowIndex][31+(j*2)] = data[i].answers[j].answer
          }
        }
        newRowIndex++
      }
      else{
        var row = []
        row.push([
        data[i].status,data[i].source.public_name,data[i].source.id,
          
          data[i].rejection_reason != null? data[i].rejection_reason.type.name:"-",data[i].rejection_reason != null? data[i].rejection_reason.name:"-",

          data[i].rejection_details,data[i].rejected_at,data[i].prospective_office,data[i].prospective_department,data[i].prospect_detail.prospect_stage,data[i].prospect_detail.prospect_pool,
          data[i].prospect_detail.prospect_owner,data[i].prospect,
          data[i].location != null? data[i].location.address:"-",
          data[i].last_activity_at,data[i].jobs[0].name,data[i].jobs[0].id,data[i].job_post_id,data[i].id,

          data[i].current_stage != null ? data[i].current_stage.name:"-",
          data[i].credited_to != null? data[i].credited_to.id:"-",data[i].credited_to != null? data[i].credited_to.first_name:"-",data[i].credited_to != null?data[i].credited_to.last_name:"-", 
          data[i].credited_to != null? data[i].credited_to.employee_id:"-",
          
          data[i].candidate_id,
          data[i].attachments[0] != null? data[i].attachments[0].url:"-",data[i].attachments[0] != null? data[i].attachments[0].type:"-",
          data[i].attachments[0] != null && data[i].attachments[1] != null? data[i].attachments[1].url:"-",
          data[i].attachments[0] != null && data[i].attachments[1] != null? data[i].attachments[1].type:"-",
          
          data[i].applied_at,
          "-","-",    "-","-",    "-","-",    "-","-"
        ])
        if(data[i].answers[0] != null){
          for(let j = 0; j < data[i].answers.length;j++){
            row[0][30+(j*2)] = data[i].answers[j].question
            row[0][31+(j*2)] = data[i].answers[j].answer
          }
        }
        var rowIndex = sheet.createTextFinder(data[i].id).matchEntireCell(true).findAll().map(x => x.getRow())
        update = sheet.getRange(rowIndex[0],1,row.length,row[0].length).setValues(row) 
      }
    }
  }
  if(newRows.length>0){
    update = sheet.getRange(sheet.getLastRow()+1,1,newRows.length,newRows[0].length).setValues(newRows)
  }
  
}


function listAllOffers(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Offers")
  var options = {     
     "method" : "GET",
     "headers" : {
    "Authorization": ""
  }
  };

  var allData = []
  var page_num = 1 
  do{    
    var response = UrlFetchApp.fetch(`https://harvest.greenhouse.io/v1/offers?page=${page_num}&per_page=500`, options) 
    if(JSON.parse(response.getContentText()).length!= 0){
      allData.push(JSON.parse(response.getContentText()))
    }
    page_num++
  }while (JSON.parse(response.getContentText()).length != 0)

  var newRows = []
  for(var data of allData){
    for(var i = 0; i < data.length; i++){
      let duplicated = sheet.createTextFinder(data[i].id).matchEntireCell(true).findAll().map(x => x.getRow()).length == 0? false:true
      if(data[i].opening != null){
        var applicationData = JSON.parse(UrlFetchApp.fetch(`https://harvest.greenhouse.io/v1/applications/${data[i].opening.application_id}`, options).getContentText())
      }
      var candidateData = JSON.parse(UrlFetchApp.fetch(`https://harvest.greenhouse.io/v1/candidates/${data[i].candidate_id}`, options).getContentText())
      if(!duplicated){
        Logger.log(i)
        newRows.push([
         data[i].id/*0*/,data[i].version/*1*/,data[i].application_id/*2*/,data[i].created_at/*3*/,data[i].updated_at/*4*/,data[i].sent_at/*5*/,data[i].resolved_at/*6*/,data[i].starts_at/*7*/,
          data[i].status/*8*/,data[i].job_id/*9*/,
          
          data[i].candidate_id/*10*/, candidateData.first_name, candidateData.last_name, 
          candidateData.recruiter != null? candidateData.recruiter.name:"-", candidateData.recruiter != null? candidateData.recruiter.employee_id: "-",
          candidateData.recruiter != null && candidateData.recruiter.coordinators != null? candidateData.recruiter.coordinators.name:"-",

          data[i].opening != null? data[i].opening.id: "-",data[i].opening != null? data[i].opening.opening_id:"-",data[i].opening != null?data[i].opening.status:"-",
          data[i].opening != null?data[i].opening.opened_at:"-",data[i].opening != null?data[i].opening.closed_at:"-",
          data[i].opening != null?data[i].opening.application_id:"-", applicationData.prospect, applicationData.applied_at, applicationData.rejected_at, applicationData.source != null? applicationData.source.id:"-", applicationData.source != null? applicationData.source.public_name:"-", 

          
          data[i].opening != null && data[i].opening.close_reason != null? data[i].opening.close_reason.id:"-",
          data[i].opening != null && data[i].opening.close_reason != null? data[i].opening.close_reason.name:"-",

          data[i].custom_fields != null? data[i].custom_fields.employment_type:"-",data[i].custom_fields != null?data[i].custom_fields.responsibilities:"-" ,data[i].custom_fields != null? data[i].custom_fields.bonus:"-",data[i].custom_fields != null? data[i].custom_fields.salary:"-",data[i].custom_fields != null? data[i].custom_fields.manager:"-",
          data[i].custom_fields != null? data[i].custom_fields.city:"-",data[i].custom_fields != null? data[i].custom_fields.expiration:"-",
 

        ])
      }
      else{
        Logger.log(i)
        var row = []
        row.push([
           data[i].id/*0*/,data[i].version/*1*/,data[i].application_id/*2*/,data[i].created_at/*3*/,data[i].updated_at/*4*/,data[i].sent_at/*5*/,data[i].resolved_at/*6*/,data[i].starts_at/*7*/,
          data[i].status/*8*/,data[i].job_id/*9*/,
          
          data[i].candidate_id/*10*/, candidateData.first_name, candidateData.last_name, 

          candidateData.recruiter != null? candidateData.recruiter.name:"-", candidateData.recruiter != null? candidateData.recruiter.employee_id: "-",
          candidateData.recruiter != null && candidateData.recruiter.coordinator != null? candidateData.recruiter.coordinator.name:"-",

          data[i].opening != null? data[i].opening.id: "-",data[i].opening != null? data[i].opening.opening_id:"-",data[i].opening != null?data[i].opening.status:"-",
          data[i].opening != null?data[i].opening.opened_at:"-",data[i].opening != null?data[i].opening.closed_at:"-",
          data[i].opening != null?data[i].opening.application_id:"-", applicationData.prospect, applicationData.applied_at, applicationData.rejected_at, applicationData.source != null? applicationData.source.id:"-", applicationData.source != null? applicationData.source.public_name:"-", 

          
          data[i].opening != null && data[i].opening.close_reason != null? data[i].opening.close_reason.id:"-",
          data[i].opening != null && data[i].opening.close_reason != null? data[i].opening.close_reason.name:"-",

          data[i].custom_fields != null? data[i].custom_fields.employment_type:"-",data[i].custom_fields != null?data[i].custom_fields.responsibilities:"-" ,data[i].custom_fields != null? data[i].custom_fields.bonus:"-",data[i].custom_fields != null? data[i].custom_fields.salary:"-",data[i].custom_fields != null? data[i].custom_fields.manager:"-",
          data[i].custom_fields != null? data[i].custom_fields.city:"-",data[i].custom_fields != null? data[i].custom_fields.expiration:"-",
 
        ])
        var rowIndex = sheet.createTextFinder(data[i].id).matchEntireCell(true).findAll().map(x => x.getRow())
        update = sheet.getRange(rowIndex[0],1,row.length,row[0].length).setValues(row) 
      }
    }
  }
  if(newRows.length>0){
    update = sheet.getRange(sheet.getLastRow()+1,1,newRows.length,newRows[0].length).setValues(newRows)
  }
}

function listAllJobs(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Jobs")
  var options = {
     
     "method" : "GET",
     "headers" : {
    "Authorization": ""
  }
  };

  var allData = []
  var page_num = 1 
  do{    
    var response = UrlFetchApp.fetch(`https://harvest.greenhouse.io/v1/jobs?page=${page_num}&per_page=500`, options)
    if(JSON.parse(response.getContentText()).length!= 0){
      allData.push(JSON.parse(response.getContentText()))
    }
    page_num++
  }while (JSON.parse(response.getContentText()).length != 0)

  var newRows = []
  var newRowIndex = 0
  for(var data of allData){
    for(var i = 0; i < data.length; i++){
      let duplicated = sheet.createTextFinder(data[i].id).matchEntireCell(true).findAll().map(x => x.getRow()).length == 0? true:false
      if(duplicated){
        Logger.log("New: "+i)
        newRows.push([
          data[i].id,data[i].name,data[i].requisition_id,data[i].notes,data[i].confidential,data[i].is_template,data[i].copied_from_id,data[i].status,data[i].created_at,data[i].opened_at,
          data[i].closed_at,data[i].updated_at,
          "-"/** departments.id - 12*/,"-"/**departments.name - 13*/,"-"/**departments.parent_id - 14*/,"-"/**departments.parent_department_external_id - 15*/,"-"/**departments.child_ids - 16*/,
          "-"/**departments.child_department_external_ids - 17*/,"-"/**departments.external_id - 18*/,
          "-"/** offices.id - 19*/,"-"/**offices.name - 20*/,"-"/**offices.location.name */,"-"/**offices.primary_contact_user_id */,"-"/**offices.parent_id - 23*/,
          "-"/**offices.parent_department_external_id - 24*/,"-"/**offices.child_ids - 25*/,"-"/**offices.child_department_external_ids - 26*/,"-"/**offices.external_id - 27*/,
        
          "-"/** hiring_team.hiring_managers.id - 28*/,"-"/**hiring_team.hiring_managers.first_name - 29*/,"-"/**hiring_team.hiring_managers.last_name - 30*/,
          "-"/**hiring_team.hiring_managers.employee_id - 31*/,
          "-"/** hiring_team.recruiters.id - 32*/,"-"/**hiring_team.recruiters.first_name - 33*/,"-"/**hiring_team.recruiters.last_name - 34*/,
          "-"/**hiring_team.recruiters.employee_id - 35*/,"-"/**hiring_team.recruiters.responsible - 36*/,

          "-"/** hiring_team.coordinators.id - 37*/,"-"/**hiring_team.coordinators.first_name - 38*/,"-"/**hiring_team.coordinators.last_name - 39*/,
          "-"/**hiring_team.coordinators.employee_id - 40*/,"-"/**hiring_team.coordinators.responsible - 41*/,
          "-"/** hiring_team.sourcers.id - 42*/,"-"/**hiring_team.sourcers.first_name - 43*/,"-"/**hiring_team.sourcers.last_name - 44*/,
          "-"/**hiring_team.sourcers.employee_id - 45*/,
          
          "-"/**custom_fields.employment_type - 46 */,"-" /**custom_fields.reason_for_headcount_request */,"-",/**custom_fields.business_case_justification */,
          
          "-" /** openings.id */, "-" /**openings.opening_id */, "-" /**openings.status */,"-" /**openings.opened_at */,"-"/**openings.closed_at */,"-"/** openings.application_id*/,
          "-"/**openings.close_reason.id */,"-"/**openings.close_reason.name */,
          "-" /** openings.id */, "-" /**openings.opening_id */, "-" /**openings.status */,"-" /**openings.opened_at */,"-"/**openings.closed_at */,"-"/** openings.application_id*/,
          "-"/**openings.close_reason.id */,"-"/**openings.close_reason.name */


        ])
        if(data[i].departments != null){
          newRows[newRowIndex][12] = data[i].departments[0].id
          newRows[newRowIndex][13] = data[i].departments[0].name
          newRows[newRowIndex][14] = data[i].departments[0].parent_id
          newRows[newRowIndex][15] = data[i].departments[0].parent_department_external_id
          newRows[newRowIndex][16] = data[i].departments[0].child_ids != null? data[i].departments[0].child_ids.join(','): null
          newRows[newRowIndex][17] = data[i].departments[0].child_department_external_ids != null? data[i].departments[0].child_department_external_ids.join(','): null
          newRows[newRowIndex][18] = data[i].departments[0].external_id
        }
        if(data[i].offices != null){
          newRows[newRowIndex][19] = data[i].offices[0].id
          newRows[newRowIndex][20] = data[i].offices[0].name
          newRows[newRowIndex][21] = data[i].offices[0].location != null? data[i].offices[0].location.name: null
          newRows[newRowIndex][22] = data[i].offices[0].primary_contact_user_id
          newRows[newRowIndex][23] = data[i].offices[0].parent_id
          newRows[newRowIndex][24] = data[i].offices[0].parent_department_external_id
          newRows[newRowIndex][25] = data[i].offices[0].child_ids != null? data[i].offices[0].child_ids.join(','):null
          newRows[newRowIndex][26] = data[i].offices[0].child_department_external_ids != null? data[i].offices[0].child_department_external_ids.join(','):null 
          newRows[newRowIndex][27] = data[i].offices[0].external_id
        }
        if(data[i].hiring_team != null){
          if(data[i].hiring_team.hiring_managers[0] != null){
            newRows[newRowIndex][28] = data[i].hiring_team.hiring_managers[0].id
            newRows[newRowIndex][29] = data[i].hiring_team.hiring_managers[0].first_name
            newRows[newRowIndex][30] = data[i].hiring_team.hiring_managers[0].last_name
            newRows[newRowIndex][31] = data[i].hiring_team.hiring_managers[0].employee_id
          }
          if(data[i].hiring_team.recruiters[0] != null){
            newRows[newRowIndex][32] = data[i].hiring_team.recruiters[0].id
            newRows[newRowIndex][33] = data[i].hiring_team.recruiters[0].first_name
            newRows[newRowIndex][34] = data[i].hiring_team.recruiters[0].last_name
            newRows[newRowIndex][35] = data[i].hiring_team.recruiters[0].employee_id
            newRows[newRowIndex][36] = data[i].hiring_team.recruiters[0].responsible
          }
          if(data[i].hiring_team.coordinators[0] != null){
            newRows[newRowIndex][37] = data[i].hiring_team.coordinators[0].id
            newRows[newRowIndex][38] = data[i].hiring_team.coordinators[0].first_name
            newRows[newRowIndex][39] = data[i].hiring_team.coordinators[0].last_name
            newRows[newRowIndex][40] = data[i].hiring_team.coordinators[0].employee_id
            newRows[newRowIndex][41] = data[i].hiring_team.coordinators[0].responsible  
          }
          if(data[i].hiring_team.sourcers[0] != null){
            newRows[newRowIndex][42] = data[i].hiring_team.sourcers[0].id
            newRows[newRowIndex][43] = data[i].hiring_team.sourcers[0].first_name
            newRows[newRowIndex][44] = data[i].hiring_team.sourcers[0].last_name
            newRows[newRowIndex][45] = data[i].hiring_team.sourcers[0].employee_id
          }
        }
        if(data[i].custom_fields != null){
          newRows[newRowIndex][46] = data[i].custom_fields.employment_type
          newRows[newRowIndex][47] = data[i].custom_fields.reason_for_headcount_request
          newRows[newRowIndex][48] = data[i].custom_fields.business_case_justification 
        }
        if(data[i].openings[0]!= null){
          newRows[newRowIndex][49] = data[i].openings[0].id
          newRows[newRowIndex][50] = data[i].openings[0].opening_id
          newRows[newRowIndex][51] = data[i].openings[0].status
          newRows[newRowIndex][52] = data[i].openings[0].opened_at
          newRows[newRowIndex][53] = data[i].openings[0].closed_at
          newRows[newRowIndex][54] = data[i].openings[0].application_id
          if(data[i].openings[0].close_reason != null){
            newRows[newRowIndex][55] = data[i].openings[0].close_reason.id
            newRows[newRowIndex][56] = data[i].openings[0].close_reason.name
          }

        }
        if(data[i].openings[1]!= null){
          newRows[newRowIndex][57] = data[i].openings[1].id
          newRows[newRowIndex][58] = data[i].openings[1].opening_id
          newRows[newRowIndex][59] = data[i].openings[1].status
          newRows[newRowIndex][60] = data[i].openings[1].opened_at
          newRows[newRowIndex][61] = data[i].openings[1].closed_at
          newRows[newRowIndex][62] = data[i].openings[1].application_id
          if(data[i].openings[1].close_reason != null){
            newRows[newRowIndex][63] = data[i].openings[1].close_reason.id
            newRows[newRowIndex][64] = data[i].openings[1].close_reason.name
          }

        }
        newRowIndex++
      }
      else{
        Logger.log("Old: "+i)
        var row = []
        row.push([
          data[i].id, data[i].name, data[i].requisition_id, data[i].notes,  data[i].confidential, data[i].is_template,  data[i].copied_from_id, data[i].status, data[i].created_at, data[i].opened_at,
          data[i].closed_at,  data[i].updated_at,
          "-"/** departments.id - 12*/,"-"/**departments.name - 13*/,"-"/**departments.parent_id - 14*/,"-"/**departments.parent_department_external_id - 15*/,"-"/**departments.child_ids - 16*/,
          "-"/**departments.child_department_external_ids - 17*/,"-"/**departments.external_id - 18*/,
          "-"/** offices.id - 19*/,"-"/**offices.name - 20*/,"-"/**offices.location.name - 21*/,"-"/**offices.primary_contact_user_id - 22*/,"-"/**offices.parent_id - 23*/,
          "-"/**offices.parent_department_external_id - 24*/,"-"/**offices.child_ids - 25*/,"-"/**offices.child_department_external_ids - 26*/,"-"/**offices.external_id - 27*/,

          "-"/** hiring_team.hiring_managers.id - 28*/,"-"/**hiring_team.hiring_managers.first_name - 29*/,"-"/**hiring_team.hiring_managers.last_name - 30*/,
          "-"/**hiring_team.hiring_managers.employee_id - 31*/,
          "-"/** hiring_team.recruiters.id - 32*/,"-"/**hiring_team.recruiters.first_name - 33*/,"-"/**hiring_team.recruiters.last_name - 34*/,
          "-"/**hiring_team.recruiters.employee_id - 35*/,"-"/**hiring_team.recruiters.responsible - 36*/,

          "-"/** hiring_team.coordinators.id - 37*/,"-"/**hiring_team.coordinators.first_name - 38*/,"-"/**hiring_team.coordinators.last_name - 39*/,
          "-"/**hiring_team.coordinators.employee_id - 40*/,"-"/**hiring_team.coordinators.responsible - 41*/,
          "-"/** hiring_team.sourcers.id - 42*/,"-"/**hiring_team.sourcers.first_name - 43*/,"-"/**hiring_team.sourcers.last_name - 44*/,
          "-"/**hiring_team.sourcers.employee_id - 45*/,
          
          "-"/**custom_fields.employment_type - 46 */,"-" /**custom_fields.reason_for_headcount_request */,"-",/**custom_fields.business_case_justification */
          
          "-" /** openings.id */, "-" /**openings.opening_id */, "-" /**openings.status */,"-" /**openings.opened_at */,"-"/**openings.closed_at */,"-"/** openings.application_id*/,
          "-"/**openings.close_reason.id */,"-"/**openings.close_reason.name */,
          "-" /** openings.id */, "-" /**openings.opening_id */, "-" /**openings.status */,"-" /**openings.opened_at */,"-"/**openings.closed_at */,"-"/** openings.application_id*/,
          "-"/**openings.close_reason.id */,"-"/**openings.close_reason.name */


        ])
        if(data[i].departments != null){
          row[0][12] = data[i].departments[0].id
          row[0][13] = data[i].departments[0].name
          row[0][14] = data[i].departments[0].parent_id
          row[0][15] = data[i].departments[0].parent_department_external_id
          row[0][16] = data[i].departments[0].child_ids != null? data[i].departments[0].child_ids.join(','): null
          row[0][17] = data[i].departments[0].child_department_external_ids != null? data[i].departments[0].child_department_external_ids.join(','):null
          row[0][18] = data[i].departments[0].external_id
        }
        if(data[i].offices != null){
          row[0][19] = data[i].offices[0].id
          row[0][20] = data[i].offices[0].name
          row[0][21] = data[i].offices[0].location != null? data[i].offices[0].location.name: null
          row[0][22] = data[i].offices[0].primary_contact_user_id
          row[0][23] = data[i].offices[0].parent_id
          row[0][24] = data[i].offices[0].parent_office_external_id
          row[0][25] = data[i].offices[0].child_ids != null? data[i].offices[0].child_ids.join(','):null
          row[0][26] = data[i].offices[0].child_department_external_ids != null? data[i].offices[0].child_department_external_ids.join(','):null 
          row[0][27] = data[i].offices[0].external_id
        }
        if(data[i].hiring_team != null){
          if(data[i].hiring_team.hiring_managers[0] != null){
            row[0][28] = data[i].hiring_team.hiring_managers[0].id
            row[0][29] = data[i].hiring_team.hiring_managers[0].first_name
            row[0][30] = data[i].hiring_team.hiring_managers[0].last_name
            row[0][31] = data[i].hiring_team.hiring_managers[0].employee_id
          }
          if(data[i].hiring_team.recruiters[0] != null){
            row[0][32] = data[i].hiring_team.recruiters[0].id
            row[0][33] = data[i].hiring_team.recruiters[0].first_name
            row[0][34] = data[i].hiring_team.recruiters[0].last_name
            row[0][35] = data[i].hiring_team.recruiters[0].employee_id
            row[0][36] = data[i].hiring_team.recruiters[0].responsible
          }
          if(data[i].hiring_team.coordinators[0] != null){
            row[0][37] = data[i].hiring_team.coordinators[0].id
            row[0][38] = data[i].hiring_team.coordinators[0].first_name
            row[0][39] = data[i].hiring_team.coordinators[0].last_name
            row[0][40] = data[i].hiring_team.coordinators[0].employee_id
            row[0][41] = data[i].hiring_team.coordinators[0].responsible  
          }
          if(data[i].hiring_team.sourcers[0] != null){
            row[0][42] = data[i].hiring_team.sourcers[0].id
            row[0][43] = data[i].hiring_team.sourcers[0].first_name
            row[0][44] = data[i].hiring_team.sourcers[0].last_name
            row[0][45] = data[i].hiring_team.sourcers[0].employee_id
          }
        }
        if(data[i].custom_fields != null){
          row[0][46] = data[i].custom_fields.employment_type
          row[0][47] = data[i].custom_fields.reason_for_headcount_request
          row[0][48] = data[i].custom_fields.business_case_justification 
        }
        if(data[i].openings[0] != null){
          row[0][49] = data[i].openings[0].id
          row[0][50] = data[i].openings[0].opening_id
          row[0][51] = data[i].openings[0].status
          row[0][52] = data[i].openings[0].opened_at
          row[0][53] = data[i].openings[0].closed_at
          row[0][54] = data[i].openings[0].application_id
          if(data[i].openings[0].close_reason != null){
            row[0][55] = data[i].openings[0].close_reason.id
            row[0][56] = data[i].openings[0].close_reason.name
          }
        }
        if(data[i].openings[1]!= null){
          row[0][57] = data[i].openings[1].id
          row[0][58] = data[i].openings[1].opening_id
          row[0][59] = data[i].openings[1].status
          row[0][60] = data[i].openings[1].opened_at
          row[0][61] = data[i].openings[1].closed_at
          row[0][62] = data[i].openings[1].application_id
          if(data[i].openings[1].close_reason != null){
            row[0][63] = data[i].openings[1].close_reason.id
            row[0][64] = data[i].openings[1].close_reason.name
          }

        }
        var rowIndex = sheet.createTextFinder(data[i].id).matchEntireCell(true).findAll().map(x => x.getRow())
        update = sheet.getRange(rowIndex[0],1,row.length,row[0].length).setValues(row) 
      }
    }
  }
  if(newRows.length>1){
    update = sheet.getRange(sheet.getLastRow()+1,1,newRows.length,newRows[0].length).setValues(newRows)
  }
}
