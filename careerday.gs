 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var firstBlockSheet = ss.getSheetByName('First Block');
 var secondBlockSheet = ss.getSheetByName('Second Block');
 var thirdBlockSheet = ss.getSheetByName('Third Block');
 var fourthBlockSheet = ss.getSheetByName('Fourth Block');
 var responseSheet = ss.getSheetByName('Form Responses 1');
 var form = FormApp.openByUrl('https://docs.google.com/a/mustangsla.org/forms/d/17b99wl8auOOoyn2wlJrGT58bbLkllt7kMV2xAL52GZM/edit?usp=drive_web')
 var firstBlockDocId = '1fzZAIGZj-GBgOB9k-zwYE6G1qZp4JFbIvKUIfceSDWM';
 var secondBlockDocId = '1HiuGB9AVd8DUGrkCmiY4JyKrSRJHWZo9JuNQC9xIH34';
 var thirdBlockDocId = '1k-pf1rmhzg02UrmxuvmN_y3OcstLluqLM0wr8aq1o2k';
 var fourthBlockDocId = '1l-ko_2u2dh7vgS67O9MqyCgKlJylBvR1KMCYr9MDXXg';
 
var studentData = getRowsData(responseSheet);
var jobs = [
'Architecture/Landscape - Rm. 208', 'Bureau of Prisons - Rm. 315', 'Dentist - Rm. 308', 'Domestic & International Business - Rm. 302', 'Insurance Representative - Rm. 316',
  'Management / Business - Rm. 303', 'Police Officer - Rm. 206', 'USC Local Government Relations - Rm. 301', 'Public Office - Rm. 201', 'Electrical-Quality Engineer - Rm. 215', 
  'Engineer - Boeing - Rm. 209', 'Municipal/Government Attorney - Rm. 306', 'Military - Rm. 216', 'Health Care - Rm. 304', 'Author - Rm. 307', 'Banking - Rm. 217', 'Business Consultant - Rm. 318', 
  'Apple Executive - Rm. 311', 'Brain Surgeon - Rm. 314'
];


// This function iterates job by job in the jobs array, searching the student response array
//    for students who selected the given job iteration in the specified class block.
//    
//    When a student's job request matches the current job iteration, the script writes the
//    student's name to the appropriate blockSheet.
//
//    The function takes two arguments: 
//        @blockSheet is a reference to the ***sheet*** in the active spreadsheet
//            corresponding to the appropriate schedule block
//        @block is a number corresponding to the appropriate schedule block.
function assignBlock(blockSheet, block){
  
  var overflowStudents = [];

  // Loop through jobs
  for (var i = 0; i<jobs.length; i++){
    var job = jobs[i];
    var newColumn = blockSheet.getRange(1, (i+1));
    var activeRange = blockSheet.setActiveRange(newColumn);
    Logger.log(' ');
    Logger.log(job);
    activeRange.setValue(job).setBackground('yellow').setFontSize(14);
    
    
    var counter = 0;
    // Loop through each student response
    for (var j = 0; j<studentData.length; j++){
      var student = studentData[j];
      var currentColumn = blockSheet.getRange((3), (i+1), (counter+1)).getValues();
      
      // Associate the numeric argument to a specific key in the studentData object elements
      switch(block){
        case 1: 
            var studentBlock = student.firstBlock;
            break;
        case 2: 
            studentBlock = student.secondBlock;
            break;
        case 3: 
            studentBlock = student.thirdBlock;
            break;
        case 4: 
            studentBlock = student.fourthBlock;
            break;
        }
        
        // If student.firstBlock matches the current job, setValue to last row of current column
        if (studentBlock == job){
          counter++;

          var studentGetRange = blockSheet.getRange((counter + 2), (i+1));
          var studentActiveRange = blockSheet.setActiveRange(studentGetRange);
          studentActiveRange.setValue(student.name);
          
        }
      }
      Logger.log(currentColumn.length);
    }

}
/*
function pushToDoc(docName, blockSheet){
  var careerDayFile = DriveApp.getFilesByName(docName).next();
  var docId = careerDayFile.getId();
  var firstBlockDoc = DocumentApp.openById(docId);
  var docBody = firstBlockDoc.getBody();
  Logger.log(docBody.getText());
  
  for (var i = 0; i<jobs.length; i++){
    var job = jobs[i];
    var jobTitle = docBody.appendParagraph(job);
    // Set Attributes of jobTitle here
    // Need to declare an object of attributes
    var student = studentData[i];
    var lastRow = blockSheet.getLastRow()
    var currentColumn = blockSheet.getRange((3), (i+1), (lastRow)).getValues();
    
    Logger.log(currentColumn);
    for (var j = 0; j<currentColumn.length; j++){
      if (currentColumn[j][0] == ''){
        continue;
      }
      else{
      docBody.appendListItem(currentColumn[j][0]);
      }
    }
    docBody.appendPageBreak();
  }
}
*/
function pushToDoc(docId, blockSheet, blockAsText){
  var firstBlockDoc = DocumentApp.openById(docId);
  var docBody = firstBlockDoc.getBody();
  
  var titleStyle = {}
  titleStyle[DocumentApp.Attribute.FONT_SIZE] = 14;
  titleStyle[DocumentApp.Attribute.BOLD] = true;
 
  if (!firstBlockDoc.getHeader()){
    var docHeader = firstBlockDoc.addHeader();
  }
  else{
    docHeader = firstBlockDoc.getHeader();
    docHeader.setAttributes(titleStyle);
  }
  docHeader.setText('Career Day -- '+blockAsText+' Assignments -- Attendance Sheet');

  var listStyle = {}
  listStyle[DocumentApp.Attribute.FONT_SIZE] = 10;
  listStyle[DocumentApp.Attribute.BOLD] = false;
  
  for (var i = 0; i<jobs.length; i++){
    var job = jobs[i];
    var jobTitle = docBody.appendParagraph(job).setAttributes(titleStyle);
    
    var student = studentData[i];
    var lastRow = blockSheet.getLastRow();
    var currentColumn = blockSheet.getRange((3), (i+1), (lastRow)).getValues();
    
    Logger.log(currentColumn);
    for (var j = 0; j<currentColumn.length; j++){
      if (currentColumn[j][0] == ''){
        continue;
      }
      else{
      docBody.appendListItem(currentColumn[j][0]).setAttributes(listStyle);
      }
    }
    docBody.appendPageBreak();
  }
}

function assignAllBlocks (){
  assignBlock(firstBlockSheet, 1);
  assignBlock(secondBlockSheet, 2);
  assignBlock(thirdBlockSheet, 3);
  assignBlock(fourthBlockSheet, 4);
}

function reduceAllChoices(){
  reduceChoices(firstBlockSheet, 1, 30);
  reduceChoices(secondBlockSheet, 2, 30);
  reduceChoices(thirdBlockSheet, 3, 30);
  reduceChoices(fourthBlockSheet, 4, 30);
}

function pushAllBlocks(){
  pushToDoc(firstBlockDocId, firstBlockSheet, 'First Block');
  pushToDoc(secondBlockDocId, secondBlockSheet, 'Second Block');
  pushToDoc(thirdBlockDocId, thirdBlockSheet, 'Third Block');
  pushToDoc(fourthBlockDocId, fourthBlockSheet, 'Fourth Block');
}

//  This function analyzes the given blockSheet, going job by job (or column by column).
//    When a given job contains 30 student names, the script removes that job choice
//    from the form
//
//    The function takes two arguments: 
//        @blockSheet is a reference to the ***sheet*** in the active spreadsheet
//            corresponding to the appropriate schedule block
//        @block is a number corresponding to the appropriate schedule block.
//        @maxSpace is a number that will cap the class at a certain number of seats
function reduceChoices(blockSheet, block, maxSpace){
  var items = form.getItems();
  var blockItem = items[block];
  var blockListItem = blockItem.asListItem();
  var updatedChoices = [];

  // Loop through jobs
  for (var i = 0; i<jobs.length; i++){
    var job = jobs[i];
    var currentColumn = blockSheet.getRange((3), (i+1), (maxSpace-1)).getValues();
    
    // Loop through the contents of the current column (current job iteration)
    for (var k = (maxSpace-2); k<currentColumn.length; k++){
      if (currentColumn[k][0] == ''){
        
      // Add the current job iteration to updated choice array if there is space left in the class (max 30 students)
      updatedChoices.push(job);
      }
    }

  }

  blockListItem.setChoiceValues(updatedChoices);
  
}

    
function resetJobsChoices(){
  var items = form.getItems();
  
  for (var i = 1; i<5; i++){
    items[i].asListItem().setChoiceValues(jobs);
  }
}
