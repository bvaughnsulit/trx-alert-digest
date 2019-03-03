function main() {
  
  const sheetId = getSheetId()
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName("1")
 
  var threads = getThreads()
  var msgs = extractMsgData(threads)
  var trxs = parseMsgs(msgs)
  trxs = dedupe(trxs, sheet)
  
  emailTrxs(trxs)
 
  writeToSheet(trxs, sheet)
}


function getSheetId(){
  return PropertiesService.getUserProperties().getProperty('sheetId')
}


function getThreads() {
  var queryTimeframe = '48h'
  var searchQuery = '(from:chase subject:"your single transaction alert from chase" newer_than:' + queryTimeframe + ')'
  var threads = []
  var start = 0
  do {
    const max = 500
    var results = GmailApp.search(searchQuery, start, max)
    threads = threads.concat(results)
    start = start + max
  } while (results.length > 0)
  
  return threads

}


function extractMsgData(threads){
  var threadMsgs = GmailApp.getMessagesForThreads(threads)
  var msgs = []
  
  for (var threadIndex in threadMsgs) {
  
    for (var msgIndex in threadMsgs[threadIndex]) {
      var msg = threadMsgs[threadIndex][msgIndex]
      var msgData = {}
      msgData.msgId = msg.getId()
      msgData.msgDate = msg.getDate()
      msgData.body = msg.getBody()
      msgs.push(msgData)
    }    
  }
  return msgs
}


function parseMsgs(msgs) {
  var trxs = []
  
  for (var msgIndex in msgs){
    var trx = {}
    var normalizedText = msgs[msgIndex].body.replace(/\r?\n|\r/g, " ")
  
    var extractedData = new RegExp(/ending in (\d{4}).+ A charge of \(([^)]+)\) (\d+\.\d+) at (.+) has been authorized on (\d{2}\/\d{2}\/\d{4}) (\d{1,2}\:\d{2}\:\d{2} \w{2} \w+)\./).exec(normalizedText)
  
    trx.trxCardNumber = extractedData[1]
    trx.trxCurrency = extractedData[2]
    trx.trxAmount = extractedData[3]
    trx.trxMerchant = extractedData[4]
    trx.trxDate = extractedData[5]
    trx.trxTime = extractedData[6]
    trx.msgId = msgs[msgIndex].msgId
    trx.msgDate = msgs[msgIndex].msgDate
    
    trxs.push(trx)
  }
  
  return trxs
}


function dedupe(trxs, sheet) {
  // gets all ids currently in the spreadsheet to dedupe new trxs
  // currently gets all ids, so this could get slow depending on how many old trxs are saved
  var trxRange = sheet.getRange(1, 1, sheet.getLastRow()).getValues()
  var ids = []
  for (var j in trxRange){
    ids.push(trxRange[j][0])
  }
  
  deduped = []
  
  for (var i in trxs) {
    if (ids.indexOf(trxs[i].msgId) == -1){
      deduped.push(trxs[i])
    }
  } 
  return deduped
}


function emailTrxs(trxs) {
  var htmlrows = ""
  
  for (var i in trxs) {
   htmlrows = htmlrows + "<tr><td>" + trxs[i].trxDate + "</td><td>" + trxs[i].trxAmount + "</td><td>" + trxs[i].trxMerchant + "</td></tr>"
  }
  
  var html = "<html><head><style type='text/css'>table,td{ font-family: monospace; border: 1px solid black; border-collapse: collapse; }</style></head><table>" + htmlrows + "</table></html>"
  
  GmailApp.sendEmail("bvaughnsulit@gmail.com", "Transactions - " + new Date, "", {
    htmlBody : html
  } )
}


function writeToSheet(trxs, sheet) {  
  for (var i in trxs) {
    sheet.appendRow([trxs[i].msgId, trxs[i].msgDate, trxs[i].trxAmount,trxs[i].trxCardNumber, trxs[i].trxCurrency,trxs[i].trxDate, trxs[i].trxMerchant, trxs[i].trxTime])
    }
}


