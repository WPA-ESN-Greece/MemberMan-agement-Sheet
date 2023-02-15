//Global Variables
const DOCUMENTATION_LINK = 'https://docs.google.com/document/d/1uCqoSNN5fHieTUz-2lBLeBr2hMXIyG_3NlG9rC10zhM/edit?usp=sharing'
var ui = SpreadsheetApp.getUi()
var ss = SpreadsheetApp.getActiveSpreadsheet()
var UsersSheet = ss.getSheetByName('users')
var settingsSheet = ss.getSheetByName('Settings')
var membersSheet = ss.getSheetByName('Members')
var formResSheet = ss.getSheetByName('Form responses')
var ss = SpreadsheetApp.getActiveSpreadsheet()
var FORM_ID = settingsSheet.getRange('C3').getValue()

function spreadsheetInfo(){

  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var UsersSheet = ss.getSheetByName('users')
  var settingsSheet = ss.getSheetByName('Settings')
  var membersSheet = ss.getSheetByName('Members')
  var formResSheet = ss.getSheetByName('Form responses')

}


//var PARENT_FOLDER = ""

// Sends Emails to new users with ESN Email, single use password and Google log in link.
function emailCredentials() {
  spreadsheetInfo()

  var lastRow = UsersSheet.getLastRow()-1
  var userData = UsersSheet.getRange(2,1,lastRow,26).getValues()

  var member = {
    name:'',
    lastName:'',
    esnEmail:'',
    password:'',
    recoveryEmail:''
    }

  var esnMail = ""
  var esnMailPassword = ""
  var subject = "Your New ESN Google Account Credentials"

  //Data Loop
  userData.forEach(function (row) {
  
  if(row[2] == "" && row[3] == ""){return} //Checks if there is an email address in each row.

   member.name = row[0];
   member.lastName = row[1];
   esnMail = row[2];
   esnMailPassword = row[3];
   member.recoveryEmail = row[7];
  
  /*
  var message = 
    `<h2>Your New ESN Google Account is Ready!</h2>`+
    `<p><b>ESN Email Address: </b> ${member.esnEmail}</p>` +
    `<p><b>Single-Use Password: </b> ${member.password}</p>`+
    `<p><i>After the first sign in to your new Google Account, you will be asked to change the password above with one only you will know.
     You can sign in <a href="shorturl.at/erBX3">here</a>.</i></p>`*/


  var htmlTemplate = HtmlService.createTemplateFromFile("email_template")
  htmlTemplate.esnMail = row[2]
  htmlTemplate.esnMailPassword = row[3]

  var message = htmlTemplate.evaluate().getContent()

  MailApp.sendEmail(
    {
            to: member.recoveryEmail,
            cc: "",
            subject: subject,
            htmlBody: message,
    }
  )})

toast("Now the new ESNers can log in to Google.","🎉 Emails sent")
}

// Generates the 'users' sheet link in the Settings
function generateUsersLink(){
  settingsSheet.getRange('C10').setValue(ss.getUrl()+'#gid='+UsersSheet.getSheetId())
  toast("The users sheet link is ready in the Settings.","🎉 Done!")
}


//Creates the "Recruiting Status" as the first Column of the Form responses
function createRecruitingStatusCol(formSheet) {
  spreadsheetInfo()

  formSheet.insertColumnBefore(1)
  formSheet.getRange('A1').setValue('Status')

  var recStatusRange = formSheet.getRange('A2:A')
  var sourceRange = settingsSheet.getRange('E3:E20')

  var rule = SpreadsheetApp.newDataValidation().requireValueInRange(sourceRange).requireValueInRange(sourceRange, true).build()
  var rules = recStatusRange.getDataValidations()

  for (var i = 0; i < rules.length; i++) {
    for (var j = 0; j < rules[i].length; j++) {
    rules[i][j] = rule}
  }
  recStatusRange.setDataValidations(rules)
  
}


function condtionalFormating(sheet){
  var formSheet = ss.getSheetByName(sheet.getName())
  //Sets Conditional formating rules
  var lastColumn = formSheet.getLastColumn()
  var lastRow = formSheet.getLastRow()
  var conditionaFormatRange = formSheet.getRange(2,1,lastRow - 1,lastColumn)


  var formatRule1 = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2=INDIRECT("Settings!E4")') //Contacted
      .setBackground("#c9daf8")
      .setRanges([conditionaFormatRange])
      .build()

  var formatRule2 = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2=INDIRECT("Settings!E5")') //Pending Coctact
      .setBackground("#fff2cc")
      .setRanges([conditionaFormatRange])
      .build()

  var formatRule3 = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2=INDIRECT("Settings!E6")') //Accepted
      .setBackground("#d9ead3")
      .setRanges([conditionaFormatRange])
      .build()

  var formatRule4 = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2=INDIRECT("Settings!E7")') //Rejected
      .setBackground("#f4cccc")
      .setRanges([conditionaFormatRange])
      .build()

    var formatRule5 = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2=INDIRECT("Settings!E8")') //Accepted & Transferred
      .setBackground("#d9ead3")
      .setFontColor("#274e13")
      .setItalic(true)
      .setRanges([conditionaFormatRange])
      .build()



  var conditionalFormatRules = formSheet.getConditionalFormatRules()
  conditionalFormatRules.push(formatRule1, formatRule2, formatRule3, formatRule4, formatRule5)
  formSheet.setConditionalFormatRules(conditionalFormatRules)

  var recStatusRange = formSheet.getRange('A2:A')
  recStatusRange.setHorizontalAlignment("left")
}


function createAgeCol(formSheet) {
  spreadsheetInfo()

  var lastColumn = formSheet.getLastColumn()
  formSheet.insertColumnAfter(lastColumn)

  formSheet.getRange(1,lastColumn+1)
  .setFormula(`={"Age";ARRAYFORMULA(IF(INDIRECT(Settings!K16)<>"",YEAR(TODAY()) - Year(DATE(INDIRECT(Settings!K16),1,1)),""))}`)

  var a1AgeRange = formSheet.getRange(1,lastColumn+1,formSheet.getLastRow()).getA1Notation()
  formSheet.getRange(a1AgeRange).setHorizontalAlignment("center")

}

//Formats Headers in Form Responses Sheet
function formatHeaders(formSheet){
  spreadsheetInfo()

  var lastColumn = formSheet.getLastColumn()
  formSheet.getRange(1,1,1,lastColumn)
  .setBackground('#2e3192')
  .setFontColor('#ffffff')
  .setFontWeight("bold")
  .setHorizontalAlignment("center")
  .setVerticalAlignment("middle")
  .setFontFamily("Roboto")
  .setWrap(true)
}

// Creates a new form from an existing file 
function createNewRecruitmentForm() {
  var ui = SpreadsheetApp.getUi()
  spreadsheetInfo()

  var parentFolderID = DriveApp.getFileById(ss.getId()).getParents().next().getId() //Spreadsheet Parent folder
  var destinationFolder = DriveApp.getFolderById(parentFolderID)
  var recruitForm = DriveApp.getFileById(FORM_ID).makeCopy("🦸‍♂️ MemberMan(agement) Form", destinationFolder)
  var formUrl = recruitForm.getUrl()
  //var recruitFormID = recruitForm.getId()

  var formCreationMessage = HtmlService.createHtmlOutput(`<p style="font-family: 'Open Sans'">You can find your new recruiting form <a href="${formUrl+"edit#responses"}"target="_blank">here</a>.
  Don't forget to Link the responses to this sheet from the responses tab and rename the new sheet "Form responses".</p>`).setWidth(400).setHeight(120)

  SpreadsheetApp.getUi().showModalDialog(formCreationMessage,"Your Form is ready!")
toast("","🎉 The Form is ready!")
}


function renameFormResponses(){
  spreadsheetInfo()
  var searchText = "Form responses"
  var sheets = ss.getSheets()
  var sheet = sheets.filter(s => s.getSheetName().includes(searchText))
  if (sheet.length > 0){sheet[0].setName("Form responses")}

}


function deleteBlankColumns(formSheet)
{
  spreadsheetInfo()

  var maxColumn = formSheet.getMaxColumns()
  var lastColumn = formSheet.getLastColumn()

  if(maxColumn-lastColumn > 0)
  {
    formSheet.deleteColumns(lastColumn+1, maxColumn-lastColumn)
  }

}


function deleteMostBlankRows(formSheet)
{

  spreadsheetInfo()

  var maxRow = formSheet.getMaxRows()
  var lastRow = formSheet.getLastRow()

  if( maxRow-lastRow -100 > 0)
  {
    formSheet.deleteRows(lastRow+1, maxRow-lastRow -100)
  }
  //Logger.log(maxRow-lastRow -100)
  
}


function toast(message, tittle, timeoutSeconds){
  ss.toast(message, tittle, timeoutSeconds)
}


function setRangesInSettings()
{
  for(var i=4;i < 17;i+=2)
  {

    settingsSheet.getRange(`K${i}`)
  .setFormula(`=LEFT(ADDRESS(1,MATCH(I${i},INDIRECT("Form responses!1:1"),0),4),1)&"2:"&LEFT(ADDRESS(1,MATCH(I${i},INDIRECT("Form responses!1:1"),0),4),1)`)
  Logger.log("K"+i)
  }

}
