function showEmailPreview() {
  var result = doGet();
  var agentSelect = result.agentSelect;
  var selectMenu = HtmlService.createTemplateFromFile('template-select-menu.html');
  selectMenu.agentSelect = agentSelect;
  var selectMenuHtml = selectMenu.evaluate().getContent();


  // Get values from the active sheet and active row
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = sheet.getActiveRange().getRowIndex();
  var rate = sheet.getLastRow();
  var userEmail = sheet.getRange(row, getColIndexByName("Primary Email")).getValue(); 
  var userFullName = sheet.getRange(row, getColIndexByName("Contact Full Name")).getValue();  
  var userCompanyName = sheet.getRange(row, getColIndexByName("Company Name")).getValue();   
  var subjectLine = "Company Proposal - " + userCompanyName ;
  var aliases = GmailApp.getAliases(); // add this line here
    
  var selectMenu = HtmlService.createTemplateFromFile('template-select-menu.html');
  var selectMenuHtml = selectMenu.evaluate().getContent();

  // Create an HTML output page that displays the email template selection menu and a button to send the email
  var output = HtmlService.createHtmlOutput(selectMenuHtml)
//var output = HtmlService.createHtmlOutput()
  .append('<h2>You will send emails from: '+ aliases[0] + '.  If the email above is correct you can continue. </h2> <h3><br>1.  Please choose a template and click preview. <br>2. Check that all of the information is correct before sending the email.<br>3.  Please bear in mind that the email function is currently in development and does not yet function; however, you can now preview templates. </h3>')
  .append(result.agentSelect)
  .setWidth(650)
  .setHeight(400);

  // Display the output page in a modal dialog box
  SpreadsheetApp.getUi().showModalDialog(output, 'Email Preview');
}
