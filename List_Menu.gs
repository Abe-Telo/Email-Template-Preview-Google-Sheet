function onOpen(e) {
  var subMenus = [       
                      {name:"Show Email Preview", functionName: "showEmailPreview"}  
                  ];
  SpreadsheetApp.getActiveSpreadsheet().addMenu("SendEmail Template", subMenus);
  Logger.log("Menu Started");
}
