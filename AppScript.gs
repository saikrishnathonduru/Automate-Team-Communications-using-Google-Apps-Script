/* 
  Create a communication rule between infulence and interest
  Based on the mendelow power matrix
*/

// Create some global variables which will be useful in further process

const sheet = SpreadsheetApp.getActiveSpreadsheet();
const ss0 = sheet.getSheetByName("Status");
const ss1 = sheet.getSheetByName("Stakeholders");
const ss2 = sheet.getSheetByName("Communication Log");

// Automate the data values based on the infulence and interest
function stakeHolders() {
var LastRow = ss1.getLastRow();
var data = ss1.getRange(2,1,LastRow-1,ss1.getLastColumn()).getValues();
// console.log(data);

// Create empty list for communication rule,Review note,Mail subject,Message Header, Footer
comRule = [];
notes   = [];
mailSub = [];
msgHead = [];
msgFoot = [];

// Looping for each row to iterate
data.forEach((row,i)=>{
  if (row[3]==="High" & row[4]==="High"){
    row[5] = "Manage closely";
    row[6] = "Daily email including update notes";
    row[7] = "Your daily update";
    row[8] = "Hi "+row[0] +", somebody has put an update, please have a look.";
    row[9] = "Thank you for your amazing inputs to this project! ";
    comRule.push([row[5]]);
    notes.push([row[6]]);
    mailSub.push([row[7]]);
    msgHead.push([row[8]]);
    msgFoot.push([row[9]]);
    ///console.log(row[3]);
    

}
  else if (row[3]==="Low" & row[4]==="High"){
    row[5] = "Keep satisfied";
    row[6] = "Monthly email including update notes";
    row[7] = "Your monthly update";
    row[8] = "Hi "+row[0] +", somebody has put an update, please have a look.";
    row[9] = "Thank you for your support to this project! ";
    comRule.push([row[5]]);
    notes.push([row[6]]);
    mailSub.push([row[7]]);
    msgHead.push([row[8]]);
    msgFoot.push([row[9]]);
    //console.log(row[3]);

} else if (row[3]==="Low" & row[4]==="Low"){
    row[5] = "Monitor";
    row[6] = "Monthly email without including update notes";
    row[7] = "Your Monthly update";
    row[8] = "Hi "+row[0] +", somebody has put an update, please have a look.";
    row[9] = "Thank you for being part of this project! ";
    comRule.push([row[5]]);
    notes.push([row[6]]);
    mailSub.push([row[7]]);
    msgHead.push([row[8]]);
    msgFoot.push([row[9]]);
    //console.log(row[3]);

} else if (row[3]==="High" & row[4]==="Low"){
    row[5] = "Keep informed";
    row[6] = "Weekly email including update notes";
    row[7] = "Your weekly update";
    row[8] = "Hi "+row[0] +", somebody has put an update, please have a look.";
    row[9] = "Thank you for your amazing support to this project! ";
    comRule.push([row[5]]);
    notes.push([row[6]]);
    mailSub.push([row[7]]);
    msgHead.push([row[8]]);
    msgFoot.push([row[9]]);
    //console.log(row[3]);

}
  else if (row[3]=== "Core team" & row[4]=== "Core team"){
    row[5] = "Core team";
    row[6] = "Email any update";
    row[7] = "Recent update";
    row[8] = "Hi "+row[0] +", somebody has put an update, please have a look.";
    row[9] = "Thank you for your amazing impact to this project! ";
    comRule.push([row[5]]);
    notes.push([row[6]]);
    mailSub.push([row[7]]);
    msgHead.push([row[8]]);
    msgFoot.push([row[9]]);
    //console.log(row[3]);
}

})

// console.log(values);
// Updating the communication Rule based on the mendelow power matrix
ss1.getRange(2,6,LastRow-1,1).setValues(comRule);
ss1.getRange(2,7,LastRow-1,1).setValues(notes);
ss1.getRange(2,8,LastRow-1,1).setValues(mailSub);
ss1.getRange(2,9,LastRow-1,1).setValues(msgHead);
ss1.getRange(2,10,LastRow-1,1).setValues(msgFoot);


  
}

/* 
  Create a function called mailSender 
  which should send an email to the respective stakeholder 
  with the status of the project
*/

function mailSender(){
  var ss = ss1.getDataRange().getValues();
  var stat = ss0.getDataRange().getValues();
  
  ss.shift();
  var review = "Email Sent";
  var message = ""; //container to hold final email content
  
  //set up the HTML to be rendered by sendEmail eventually.
  //adding a table for the column headings and content inside it.
  
  message = "  <table>";
  message += "<tr>"
  message += "<td><b>Task</td>"
  message += "<td><b>owner</td>"
  message += "<td><b>status</td>"
  message += "<td><b>Last update</td>"
  message += "</b> </tr>"

  //Loop through the content in data grid and set them up in each column with new row for each.
  for(var i=1;i<stat.length; i++)
  {
    message+= "<tr>"
    message+= "<td>" + stat[i][0] + "</td>"
    message+= "<td>" + stat[i][1] + "</td>"
    message+= "<td>" + stat[i][2] + "</td>"
    message+= "<td>" + stat[i][3] + "</td>"
    message+= "</tr>"
  }
  //end of table tag
  message +="</table>"
  // console.log(ss);
  /* 
      Loop the data range of stakeholders 
      check whether email is sent or not
      if not send an email
  */ 
  for (var i = 0; i < ss.length; ++i){
    var row = ss[i];
    // Check whether email is sent or not
    if(row[10]!==review){
      var recep = row[2];
      // console.log(recep);
      var sub = row[7];   
      var body = row[8] + '\n\n' + message + '\n\n' + row[9];
      // Send an email
      GmailApp.sendEmail(recep,sub,'',{htmlBody:body});
      // row[10]="mail sent";
      // review.push([row[10]]);
      ss1.getRange(2+i,11).setValue(review);
      // create dateTime variable and get the time using Utilities function
      var dateTime = Utilities.formatDate(new Date(),Session.getTimeZone(),"dd-MM-yyyy '|' HH:mm:ss");
      ss2.appendRow([dateTime,recep,sub]);
      SpreadsheetApp.flush();
    }
    // otherwise log as updated
    else{
      console.log("Updated");      
    }
  }
  
  
}

function bodyHtml(){
    //set up the HTML to be rendered by sendEmail eventually.
  //adding a table for the column headings and content inside it.
  var stat = ss0.getDataRange().getValues();
  message = "  <table>";
  message += "<tr>"
  message += "<td><b>Task</td>"
  message += "<td><b>owner</td>"
  message += "<td><b>status</td>"
  message += "<td><b>Last update</td>"
  message += "</b> </tr>"

  //Loop through the content in data grid and set them up in each column with new row for each.
  for(var i=1;i<stat.length; i++)
  {
    message+= "<tr>"
    message+= "<td>" + stat[i][0] + "</td>"
    message+= "<td>" + stat[i][1] + "</td>"
    message+= "<td>" + stat[i][2] + "</td>"
    message+= "<td>" + stat[i][3] + "</td>"
    message+= "</tr>"
  }
  //end of table tag
  message +="</table>"
}
// Manage closely
function manageClosely(){
  // create a loop system for trigger : Manage Closely
  var data = ss1.getRange(2,1,ss1.getLastRow(),ss1.getLastColumn()).getValues();
  for(var i =0;i < data.length; i++){
    var row = data[0];
    if (row[5]==='Manage closely' & row[10]===''){
      var body = new bodyHtml();
      var msg = row[8] + '\n\n' + body + '\n\n' + row[9];
      GmailApp.sendEmail(row[2],row[7],'',{htmlbody:msg});
      Browser.msgBox("Email Sent!");
    }

  }

}

// Keep satisfied
function keepSatisfied(){
  // create a loop system for trigger : Keep Satisfied
  var data = ss1.getRange(2,1,ss1.getLastRow(),ss1.getLastColumn()).getValues();
  for(var i =0;i < data.length; i++){
    var row = data[i];
    if (row[5]==='Keep satisfied' & row[10]===''){
      var body = new bodyHtml();
      var msg = row[8] + '\n\n' + body + '\n\n' + row[9];
      GmailApp.sendEmail(row[2],row[7],'',{htmlbody:msg});
      Browser.msgBox("Email Sent!");
    }

  }
}

// Moniotor
function monitoR(){
  // create a loop system for trigger : Monitor
  var data = ss1.getRange(2,1,ss1.getLastRow(),ss1.getLastColumn()).getValues();
  
  for(var i =0;i < data.length; i++){
    var row = data[i];
    if (row[5]==='Monitor' & row[10]===''){
      var body = new bodyHtml();
      var msg = row[8] + '\n\n' + body + '\n\n' + row[9];
      GmailApp.sendEmail(row[2],row[7],'',{htmlbody:msg});
      
      Browser.msgBox("Email Sent!");
    }

  }
}

//Keep informed
function keepInformed(){
  // create a loop system for trigger : Keep Informed
  var data = ss1.getRange(2,1,ss1.getLastRow()-1,ss1.getLastColumn()).getValues();
  // console.log(data);
  for(var i =0;i < data.length; ++i){
    var row = data[i];
    if (row[5]==='Keep informed' & row[10]!=='Email Sent'){
      var body = new bodyHtml();
      var msg = row[8] + '\n\n' + body + '\n\n' + row[9];
      console.log(msg);
      MailApp.sendEmail(row[2],row[7],'',{htmlbody:msg});
      ss1.getRange(2+i,11).setValue('Email Sent');
      // create dateTime variable and get the time using Utilities function
      var dateTime = Utilities.formatDate(new Date(),Session.getTimeZone(),"dd-MM-yyyy '|' HH:mm:ss");
      ss2.appendRow([dateTime,row[2],row[7]]);
      SpreadsheetApp.flush();
      // console.log('sent');
      // Browser.msgBox("Email Sent!");
    }

  }

}

// core team
function coreTeam(){
  // create a loop system for trigger : Core Team
  var data = ss1.getRange(2,1,ss1.getLastRow(),ss1.getLastColumn()).getValues();
  for(var i =0;i < data.length; i++){
    var row = data[i];
    if (row[5]==='core team' & row[10]===''){
      var body = new bodyHtml();
      var msg = row[8] + '\n\n' + body + '\n\n' + row[9];
      MailApp.sendEmail(row[2],row[7],'',{htmlbody:msg});
      Browser.msgBox("Email Sent!");
    }

  }

}

// Create a custom function

function onOpen(){
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Custom Functions")
  .addItem("Stakeholders Update","stakeHolders")
  .addItem("Send Mail","mailSender")
  .addToUi();
}
