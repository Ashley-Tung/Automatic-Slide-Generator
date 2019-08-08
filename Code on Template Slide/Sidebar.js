/** 
Sidebar for Feedback Email for CIS Staff Meetings
Written in Summer 2019 by
Ashley Tung
atung@g.hmc.edu
**/

function onOpen(e) {
  var ui = SlidesApp.getUi();
  ui.createMenu('Feedback Email Sidebar')
      .addItem('Launch Sidebar', 'showSidebar')
      .addToUi();
}


function showSidebar() {
  var ss=SlidesApp.getActivePresentation();
  var slides=ss.getSlides();
  var userInterface=HtmlService.createHtmlOutputFromFile('Button')
            .setTitle('Feedback Email Sidebar')
            .setWidth(300);
  SlidesApp.getUi().showSidebar(userInterface);
}



function formEmail() {
  // Sends email with the feedback form after the meeting
  // Gets the date in "month/year" format
  var date = new Date();
  var month = date.getMonth() +1+"";
  var year = date.getFullYear() +"";
  
  // Updated Date Information
  var conversion = {
                     1:'January',
                     2: 'February',
                     3:'March',
                     4:'April',
                     5:"May",
                     6:'June',
                     7:"July",
                     8:"August",
                     9:"September",
                     10:"October",
                     11:"November",
                     12:"December"
                   };
  
  date = conversion[month] + " " + year;
  
  // Email list for all CIS Staff
  var staffEmail = "cis-staff@g.hmc.edu";
  
  // Draws email format from the HTML template
  var template = HtmlService.createTemplateFromFile("FeedbackEmail.html");
  
  // Sends email with link to the feedback form
  template.date = date;
  var dataSpreadsheetId = '1FXL4ID6EPI5hoCti3rZIrXd8GJeUXv4X_f4oOe54p44';
  template.agendaLink = "https://docs.google.com/spreadsheets/d/" + dataSpreadsheetId;
  MailApp.sendEmail({
    to: staffEmail,
    name: "CIS Automated Slide Generator",
    subject: "CIS Staff Meeting Feedback - " + template.date,
    htmlBody: template.evaluate().getContent(),
    noReply:true
  });
}