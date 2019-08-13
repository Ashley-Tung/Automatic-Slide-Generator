/*
Automated Slide Generation for CIS Staff Meetings
Written in Summer 2018 by
Samuel Nunoo and Ashley Tung
snunoo@g.hmc.edu // atung@g.hmc.edu
*/

/*
Notes:
If you want to make changes to the default agenda items present in spreadsheet, be sure to update spreadsheet range
*/

function onOpen(e) {
  // onOpen(e) creates the custom menu in the agenda spreadsheet
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('CIS Automated Slide Generator')
      .addItem('Create Slides and Email', 'slidesGenerator')
      .addToUi();
}

// Declares IDs
var dataSpreadsheetId = '_spreadsheet_id_here_'; //spreadsheet id the main code is bound to
var templatePresentationId = '_template_id_here_'; //slides id of template presentation
var presentationCopyId;

// Uses the Sheets API to load data from agenda sheet
var sheet = SpreadsheetApp.openById(dataSpreadsheetId).getSheetByName('Current Agenda');
var sheetLastRow = sheet.getLastRow()+1;

// Gets the date in "month/year" format
var date = new Date();
var month = date.getMonth() +1+"";
var year = date.getFullYear() +"";

// Updated Date Information - Samuel
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

function slidesGenerator() {
  //Creates meeting slides, emails them to presenters, and archives the agenda
  addContent();
  slideEmail();
  archiveAgenda();
}
 
function getName(presenterEmail){
    // Calls a User Information API, a Google Apps Script deployed as a web app returning JSON
    var url = "_script_link_?email=\'"+presenterEmail+"\'";
    var userDisplayName = UrlFetchApp.fetch(url);

    // Removes beginning and end quotes from the returned JSON
    // Returns First Name only 
    var presenter = JSON.parse(userDisplayName).split(" ")[0];
    Logger.log(presenter)
  
  // Returns the name 
  return presenter;
}

function addContent() {
  // Copies the template presentation and populates it with agenda items
  // Duplicates the template presentation using the Drive API
  var copyTitle = date + ' CIS Staff Meeting';
  var driveResponse = DriveApp.getFileById(templatePresentationId).makeCopy(copyTitle);
  presentationCopyId = driveResponse.getId()
  
  // Declares IDs for the template, new copy, the template for each topic's slide
  var copyPresentation = SlidesApp.openById(presentationCopyId);
  var templatePresentation = SlidesApp.openById(templatePresentationId);
  // Gets the empty 3rd slide template
  var templateSlide = templatePresentation.getSlides()[2];
    
  // Accesses the agenda slide and its text
  var agendaSlide = copyPresentation.getSlides()[1];
  var placeholder = agendaSlide.getPlaceholder(SlidesApp.PlaceholderType.BODY);
  var shape = placeholder.asShape();
  
  // Substitutes the current date into the presentation
  copyPresentation.replaceAllText('{{Date}}', date);
  
  // Sets up the variables for each record on the agenda 
  var sheet = SpreadsheetApp.getActiveSheet();
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    Logger.log(row);
    var topic = row[0]; // topic in column 1 (A)
    var time = row[1]; // time in column 2 (B)
    time = time +""; //Make time a string
    var presenterEmail = row[2]; // email in column 3 (C)
    var notes = row[3]; // notes in column 4 (D)
    
    // Checks for invalid email addresses and throws errors accordingly
    try {
      var presenter = getName(presenterEmail);
    } 
    catch(exception){
      Logger.log(exception);
      throw "Only valid _(domain)_ email addresses are allowed. Check email addresses and try again. Other problems may be: typing in cells outside the labeled columns";
    }
    
    
    // Inserts lines on the agenda slide for each each topic
    var textRange = shape.getText()
    //If we are on the last topic, do not create a new line, else create new line
    var agendaText = (i == values.length-1) ? '{{Topic}} - {{Presenter}} ({{Time}} min)': '{{Topic}} - {{Presenter}} ({{Time}} min) \n';
    var insertedText =  textRange.appendText(agendaText);
    insertedText.getTextStyle().setFontSize(18);
    insertedText.getTextStyle().setFontFamily('Helvetica');
    textRange.getListStyle().applyListPreset(SlidesApp.ListPreset.DISC_CIRCLE_SQUARE);
    
    // Text tags on each topic slide to be replaced
    var topicReplacement = '{{Topic}}';
    var timeReplacement = '{{Time}}';
    var presReplacement = '{{Presenter}}';
    
    // Replaces the text tags on the agenda slide with the corresponding text - Ashley
    copyPresentation.replaceAllText(topicReplacement, topic);
    copyPresentation.replaceAllText(timeReplacement, time);
    copyPresentation.replaceAllText(presReplacement, presenter); 

    // Inserts topic and presenter to next topic slide and creates a new topic slide after - Ashley
    var topicSlideIndex = 1+i;
    copyPresentation.getSlides()[topicSlideIndex].replaceAllText(topicReplacement, topic);
    copyPresentation.getSlides()[topicSlideIndex].replaceAllText(presReplacement, presenter);
    //Keep adding topic slides until we reach the last topic
    if (topicSlideIndex != values.length){copyPresentation.insertSlide(topicSlideIndex+1, templateSlide);}
  }
  

  // Insert recurring agenda items after the loop
  var textRange = shape.getText();
  var insertedText = textRange.appendText('\nReview\nOpen Forum');
  insertedText.getTextStyle().setFontSize(18);
  insertedText.getTextStyle().setFontFamily('Helvetica');
  var st = textRange.getListStyle().applyListPreset(SlidesApp.ListPreset.DISC_CIRCLE_SQUARE);
  
  
}

function slideEmail() {
  // slideEmail() sends an email (with a link to the meeting slides) to all presenters
  
  // Draws email format from the html template
  var template = HtmlService.createTemplateFromFile("PresenterEmail.html");
  
  // Sets up the variables for each record on the agenda
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var topic = row[0]; // topic in column 1 (A)
    var time = row[1]; // time in column 2 (B)
    time = time +""; //Make time a string
    var presenterEmail = row[2]; // email in column 3 (C)
    var notes = row[3]; // notes in column 4 (D)
      
    // Sends email with link to template slides
    template.presenter = getName(presenterEmail);
    template.topic = topic;
    template.date = date;
    template.slideLink = "https://docs.google.com/presentation/d/" + presentationCopyId;
    template.notes = notes;
      
    var subject = "Staff Meeting Slides";
    MailApp.sendEmail({
      to: presenterEmail,
      name: "CIS Automated Slide Generator",
      subject: "CIS Staff Meeting Slides - " + template.date,
      htmlBody: template.evaluate().getContent(),
      noReply: true
    }); 
  }
}

function archiveAgenda() {
  // Clears the current agenda and archives it
  
  var archiveSheet = SpreadsheetApp.openById(dataSpreadsheetId).getSheetByName('Archive');
  // The last row of content in the archive sheet
  var archiveLastRow = archiveSheet.getLastRow();
  
  // Removes the agenda items from the current agenda
  var agendaLastRow = sheet.getLastRow();
  var agendaItems = sheet.getRange("A6:D"+agendaLastRow );
  agendaItems.clearContent();
    
  // Inserts the timestamp for each newly archived agenda item into the archive sheet
  var dateRange = archiveSheet.getRange(archiveLastRow+1, 1, values.length-1, 1);
  dateRange.setValue(date);
 
  // Inserts the newly archived agenda items into the archive sheet
  var range = archiveSheet.getRange(archiveLastRow+1, 2, values.length, 4);
  range.setValues(values);
    
}