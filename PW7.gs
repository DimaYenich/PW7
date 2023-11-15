function createBulkPDFs()
{
  const docFile = DriveApp.getFileById("1VH6lZH9zREZo3PUofNGsjfCMWhy1ApVJOvIXSQzCP8w")
  const tempFolder = DriveApp.getFolderById("1wWoXAyQzhA7_REy8EdQmM51pWnTCEqWj")
  const pdfFolder = DriveApp.getFolderById("1LjY-gdQsdxBgxEeVweJvXze_xaykERW4")
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("students")
  
  const data = spreadsheet.getRange(2, 1, spreadsheet.getLastRow()-1, spreadsheet.getLastColumn()).getDisplayValues()
  const lessons = spreadsheet.getRange(1, 4, 1, spreadsheet.getLastColumn() - 3).getValues()[0]

  data.forEach(row=>
  {
    createPDF(row[0],row[2],row[1], row[0],[row[3], row[4], row[5]], docFile, tempFolder, pdfFolder, lessons)
  })
}

function createPDF(name, email, group, pdfName, array, docFile, tempFolder, pdfFolder, lessons)
{ 
  const tempFile = docFile.makeCopy(tempFolder)
  const tempDocFile = DocumentApp.openById(tempFile.getId())
  const body = tempDocFile.getBody()

  let sum = 0;
  array.forEach(num => {
    sum += parseInt(num);
  });
  let length = array.length;
  var table = []
  for(var i = 0; i < array.length; i++)
  {
    table.push([lessons[i],array[i]])
  }

  table.push(["Сереній бал",(sum/length).toFixed(2)])
  body.appendTable(table)
  body.replaceText("{name}",name)
  body.replaceText("{email}",email)
  body.replaceText("{group}", group)
  tempDocFile.saveAndClose()
  const pdfContentBlob = tempFile.getAs(MimeType.PDF)
  const pdfFile = pdfFolder.createFile(pdfContentBlob).setName(pdfName)
  const pdfFileId = pdfFile.getId();
  const file = DriveApp.getFileById(pdfFileId);
  file.addViewer(email)
  tempFile.setTrashed(true);
}


