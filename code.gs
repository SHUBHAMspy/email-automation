function createMenu() {
  let menu = SpreadsheetApp.getUi().createMenu('Process RFQ');
  menu.addItem('Track Sent Email','scanSent');
  menu.addItem('Track Received Email','scanReceived')
  menu.addToUi()
  Logger.log(Session.getActiveUser().getEmail());
}

let emailInfo = []

function scanSent(){
  let subject = "RFQ"
  let lastTwentyFourHours = new Date();
  lastTwentyFourHours.setHours(lastTwentyFourHours.getHours() - 24);
  Logger.log(lastTwentyFourHours);
  
  let threads = GmailApp.search("from:me in:sent subject:" + subject);
  Logger.log("There are " + threads.length + " threads in Sent matching with the specified query.");
  Logger.log(threads)

  for(let i = 0; i < threads.length; i++){
    if(threads[i].getLastMessageDate() > lastTwentyFourHours && threads[i].getLastMessageDate() <= new Date() ){
      let messages = threads[i].getMessages();
      let info;
      
        info = extractDetails(messages[0])
      
      //emailInfo.push(info)
      Logger.log(emailInfo)
      let sheetSize = getSheetSize()
      Logger.log(sheetSize)
      let sheet = getSheet()
      let requiredRange = sheet.getRange(2,2,sheetSize)
      if(!hasValue(requiredRange,info.rfqNumber)){
        modifySheet(sheetSize === 0 ? i+1: sheetSize + 1,info)
      }
    }
  }
}

function scanReceived(){
  // let sheet = getSheet()
  // let sheetSize = getSheetSize()
  // let emailRange = sheet.getRange(2,5,sheetSize)
  // let emails = emailRange.getValues().flat()
  // let threads
  // for(let email of emails){
  //   Logger.log(email)
  // }
  threads = GmailApp.search(`subject:RFQ`)
  Logger.log(threads.length)
  Logger.log(threads)
  let infos = []
  for(let thread of threads){
    let messages = thread.getMessages();
    Logger.log(messages.length)
    if(messages.length >=2){
      info = extractResponse(messages[messages.length-1])
      Logger.log(info)
      infos.push(info)
    }
  }
  Logger.log(infos)  
  for(let info of infos){
    if(!info.responseRead) recordResponse(info)
    Logger.log(info)
  }
}

function getSheetSize(){
  let sheet = getSheet()
  let range = sheet.getDataRange()
  //Logger.log(range.getValues())
  return range.getValues().length - 1
}

function hasValue(range, value) {
  let rfqValues = range.getValues().flat()  
  for(let rfqValue of rfqValues){
    if(rfqValue.toFixed(0) === value) return true
    else false
  }
  //return range.getValues().flat().includes(value)
}
function extractResponse(message){
  let body = message.getPlainBody();
  let senderEmail = message.getFrom().replace(/^.+<([^>]+)>$/,"$1")
  let subject = message.getSubject();
  
  let subjectStrings = subject.split(" ")
  // let length = subjectStrings.length;
  let rfqNumberReply = subjectStrings[2]
  let attachments = message.getAttachments({
    includeAttachments:true
  })
  Logger.log(attachments)
  let hasAttachment = attachments.length > 0 ? "Yes":"No"
  let replyDate = message.getDate()
  return {
    senderEmail,
    rfqNumberReply,
    body,
    hasAttachment,
    attachments,
    replyDate,
    response: "yes",
    responseRecorded: false
  }
}

function extractDetails(message){
  let subject = message.getSubject();
  let body = message.getPlainBody();
  let receiverEmail = message.getTo().replace(/^.+<([^>]+)>$/,"$1")
  //Logger.log(receiverEmail)
  let subjectStrings = subject.split(" ")
  let length = subjectStrings.length;
  // Logger.log(Math.floor(length))
  //Logger.log(subjectStrings)
  let rfqNumber = subjectStrings[1]
  //Logger.log(rfqNumber)
  let sender = ''
  for(let i = 2; i < length - 2; i++){
    sender += subjectStrings[i] + ' ';
  }
  //Logger.log(sender)
  return {
    rfqNumber,
    sender,
    receiverEmail,
    subject,
    body
  }
}

function recordResponse(response){
  Logger.log(response.attachments)
  let mailBody = response.body.split('\n')[0]
  let searchValue = response.senderEmail
  let sheet = getSheet();
  let lastColumn = sheet.getLastColumn()
  //Logger.log(lastColumn)
  let dataRange = sheet.getDataRange();
  let values = dataRange.getValues();
  Logger.log(values)
  let rangeValues = getRowRange(values,searchValue)
  let range = sheet.getRange(rangeValues.rowNumber,rangeValues.columnNumber,1, lastColumn -(rangeValues.columnNumber - 1))
  //Logger.log(range.isBlank())
  if(range.isBlank()){
    let attachmentLink
    if(response.attachments.length > 0) attachmentLink = saveGmailToDrive(response.attachments)
    
    let data = [[response.response,response.hasAttachment,attachmentLink?attachmentLink:"",mailBody,response.replyDate]]
    range.setValues(data)
  }

  // 
  Logger.log(range.getValues().flat())
  response.responseRecorded = true
  
  // Logger.log(values.flat())
  // let i = values.flat().indexOf(value)
  // Logger.log(i)
  // var columnIndex = i % lastColumn
  // Logger.log(columnIndex)
  // var rowIndex = ((i - columnIndex) / lastColumn);
  // Logger.log(rowIndex)

  //let range = sheet.getRange(2,5,sheetSize)
}

function modifySheet(serialNumber,info){
  let sheet = getSheet();
  sheet.appendRow([serialNumber,info.rfqNumber,info.sender,info.subject,info.receiverEmail,info.body])
}

function getRowRange(values,value){
  for (let i = 0; i < values.length; i++) {
    let row = "";
    for (let j = 0; j < values[i].length; j++) {     
      if (values[i][j] == value) {
        // row = values[i][j+1];
        Logger.log(j + 3);
        Logger.log(i + 1); // This is your row number
        return {
          rowNumber: i+1,
          columnNumber: j+3
        }
      }
    }    
  }  
}

function saveGmailToDrive(attachments){
  let folderId = '1aDQHDF5zM_cX6WgdUnLMNcMt5EszMKDs'

  for(let attachment of attachments){
    Drive.Files.insert({
      title: attachment.getName(),
      mimeType: attachment.getContentType(),
      parents: [{id:folderId}]
    })
    attachment.copyBlob()
  }
  let folder = DriveApp.getFolderById(folderId)
  let driveFolderFiles = folder.getFiles()
  Logger.log(driveFolderFiles)
  let lastFileId
  let fileUrl
  if(driveFolderFiles.hasNext()){
    lastFileId = driveFolderFiles.next().getId()
    fileUrl = driveFolderFiles.next().getUrl()
  }
  Logger.log(lastFileId.toString());
  return fileUrl
}

function getSheet(){
  let spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadSheet.getActiveSheet();
  return sheet;
}
// scanSent()