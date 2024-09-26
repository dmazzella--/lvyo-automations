function generateAndSendTickets() {
  // Open the active sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow(); // Get the last row in the sheet
  
  for (var i = 2; i <= lastRow; i++) {
    var firstName = sheet.getRange(i, 1).getValue(); // Column A (First Name)
    var lastName = sheet.getRange(i, 2).getValue();  // Column B (Last Name)
    
    // Get email addresses from columns C, D, and E
    var email1 = sheet.getRange(i, 3).getValue();    // Column C (Email 1)
    var email2 = sheet.getRange(i, 4).getValue();    // Column D (Email 2)
    var email3 = sheet.getRange(i, 5).getValue();    // Column E (Email 3)
    
    // Filter out any blank emails
    var emailList = [email1, email2, email3].filter(String);  // Only keep non-empty emails
    
    if (emailList.length > 0) {
      // Create the ticket URL
      var ticketUrl = `https://lvyo.org/ticket?fname=${firstName}&lname=${lastName}&email=${email1}`;
      
      // Generate the QR code using quickchart.io API
      var qrCodeUrl = `https://quickchart.io/qr?text=${encodeURIComponent(ticketUrl)}&size=150`;
      
      // Set the subject and message for the email
      var subject = `[LVYO] Your Comp Ticket for the Fall 2024 Concert`;
      var message = `Dear ${firstName} ${lastName} (and parents),\n\nHere is your comp ticket for two people for the LVYO Fall Concert. Please show this email with QR code at the front door on your phone (Save paper!).\nMore tickets can be purchased on https://lvyo.org/\n\nOct 09, 2024, 6:30 PM\n\nLas Vegas Academy of the Arts, 315 S 7th St, Las Vegas, NV 89101, USA\n\nTicket URL: ${ticketUrl}`;
      
      // HTML body for the email with the QR code image
      var htmlBody = `<p>Dear ${firstName} ${lastName} (and parents),</p>
                      <p>Here is your comp ticket for two people for the LVYO Fall Concert. Please show this email with QR code at the front door on your phone (Save paper!)</p>
                      <p>More tickets can be purchased on https://lvyo.org/</p>
                      <p>Oct 09, 2024, 6:30 PM</p>
                      <p>Las Vegas Academy of the Arts, 315 S 7th St, Las Vegas, NV 89101, USA</p>
                      <p><a href="${ticketUrl}">${ticketUrl}</a></p>
                      <img src="${qrCodeUrl}" alt="QR Code Ticket">`;
      
      // Send the email to all email addresses in the list
      for (var j = 0; j < emailList.length; j++) {
        MailApp.sendEmail({
          to: emailList[j],
          subject: subject,
          body: message,
          htmlBody: htmlBody
        });
      }
      
      // Optionally log that emails were sent (you can remove this line if not needed)
      Logger.log('Email sent to: ' + emailList.join(', '));
    }
  }
}

