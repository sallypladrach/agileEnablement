function createCaseStudyPresentation() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Proposal Requests"); // Replace with your sheet name
    var data = sheet.getDataRange().getValues();
  var lastRow = sheet.getLastRow();
  var row = data[lastRow-1]
  var projectTypeColumn = 1; // Assuming project type is in the second column (index 1)
  var requesterEmailColumn = 2; // Assuming email of the person making the request is in the third column (index 1)
  var clientNameColumn = 3; // Assuming project type is in the second column (index 1)
  var technologiesColumn = 4; // Assuming project type is in the second column (index 1)
    var companyNameColumn = 3; // Assuming project type is in the second column (index 3)

  var sourcePresentationId = 'XXXXXXXXXXX'; // Replace with your source presentation ID
  var sourcePresentation = SlidesApp.openById(sourcePresentationId);
  var sourceSlides = sourcePresentation.getSlides();


  var projectType = row[projectTypeColumn];
  var clientName = row[clientNameColumn];
  var requesterEmail = row[requesterEmailColumn];
  var companyName = row[companyNameColumn];
  var technologiesArray = row[technologiesColumn].split(", ");

  var newPresentation = SlidesApp.create(`${projectType} - ${clientName}`);

  var folderId = 'XXXXXXXXXXXXXX'; // Replace with your shared folder ID

  moveFileToFolder(newPresentation.getId(), folderId); 


  
  sourceSlides.forEach(function(slide){
    var slideTitle = slide.getNotesPage().getSpeakerNotesShape().getText().asString();
    console.log(slideTitle)
    let technologyMatch = false;
    technologiesArray.map((tech) => {
      if (slideTitle.includes(tech) && tech !== "") {
        technologyMatch = true
      }
    })


    if(slideTitle.includes(projectType) || slideTitle.includes("all") || technologyMatch === true) { // Check if slide title matches the project type
      var importedSlide = newPresentation.appendSlide(slide);
      // Customize importedSlide as needed
    }
  });

  newPresentation.getSlides().forEach(function(slide){
    slide.getNotesPage().getSpeakerNotesShape().getText().clear();

    var shapes = slide.getShapes();
    shapes.forEach(function(shape){
      if(shape.getText().asString().includes('{{companyName}}')){ // Check for placeholder
        shape.getText().setText(companyName); // Replace placeholder with form response
      }
    });
  })

  newPresentation.getSlides().forEach(function(slide){
    slide.getNotesPage().getSpeakerNotesShape().getText().clear();
  })

  const removeFirstSlide = newPresentation.getSlides().shift()
  removeFirstSlide.remove()

  var emailAddress = requesterEmail; // Replace with the recipient's email address
  var subject = 'New Case Study Presentation Created: ' + `${projectType} - ${clientName}`;
  var message = 'A new case study presentation has been generated.\n\n' +
                'You can view it here: ' + newPresentation.getUrl();

  // Send the email
  sendEmailNotification(emailAddress, subject, message);

  Logger.log(`${projectType} - ${clientName}: ` + newPresentation.getUrl());

  function sendEmailNotification(emailAddress, subject, message) {
    MailApp.sendEmail(emailAddress, subject, message);
  }

  function moveFileToFolder(fileId, folderId) {
    var file = DriveApp.getFileById(fileId);
    var folder = DriveApp.getFolderById(folderId);

    file.moveTo(folder);
  }
 
}
