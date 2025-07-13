// Code.gs
// Compatibility-focused: ES5 only, robust error handling, avoid modern JS, fallback for spreadsheet access

function doGet(e) {
  // Use only standard HTML output for compatibility
  var templateFile = (e && e.parameter && e.parameter.poc) ? "poc" : "index";
  return HtmlService.createHtmlOutputFromFile(templateFile)
    .setTitle("Automated PoC System")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  // Compatibility: avoid template includes if not used
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSpreadsheet_() {
  // Use openById for reliability if possible, fallback to getActiveSpreadsheet
  // Replace 'YOUR_SPREADSHEET_ID' with actual ID for best compatibility
  try {
    // return SpreadsheetApp.openById('YOUR_SPREADSHEET_ID');
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    throw new Error('Unable to access spreadsheet. Please contact admin.');
  }
}

function generateAndSendOTP(email) {
  try {
    // Basic email validation
    if (!email || email.indexOf('@ashoka.edu.in') === -1) {
      return {
        success: false,
        message: "Please use a valid Ashoka University email address."
      };
    }

    var ss = getSpreadsheet_();
    var requestSheet = ss.getSheetByName("Request");
    var pocsSheet = ss.getSheetByName("pocs");
    var pendingSheet = ss.getSheetByName("Pending");
    if (!requestSheet || !pocsSheet || !pendingSheet) {
      return {
        success: false,
        message: "System error: Required sheet(s) missing. Please contact admin."
      };
    }
    
    // Check if email already has pending request
    var pendingData = pendingSheet.getDataRange().getValues();
    for (var i = 1; i < pendingData.length; i++) {
      if (pendingData[i][3] === email && pendingData[i][5] === 'Pending') {
        return {
          success: false,
          message: "You already have a pending request. Please wait for approval."
        };
      }
    }
    
    // Check if email exists in POCs database
    var pocsData = pocsSheet.getDataRange().getValues();
    var emailInDatabase = false;
    var studentName = "";
    
    for (var i = 1; i < pocsData.length; i++) {
      if (pocsData[i][2] === email) {
        emailInDatabase = true;
        studentName = pocsData[i][1];
        break;
      }
    }
    
    if (!emailInDatabase) {
      return {
        success: false,
        message: "Your email ID was not found in our database. Please contact the administrator."
      };
    }
    
    var otp = Math.floor(1000 + Math.random() * 9000).toString();
    
    // Update or add OTP in Request sheet
    var requestData = requestSheet.getDataRange().getValues();
    var rowFound = -1;
    
    for (var i = 0; i < requestData.length; i++) {
      if (requestData[i][0] === email) {
        rowFound = i + 1;
        break;
      }
    }
    
    var expiryTime = new Date(new Date().getTime() + 10 * 60 * 1000); // OTP expires in 10 minutes

    if (rowFound > 0) {
      requestSheet.getRange(rowFound, 2).setValue(otp);
      requestSheet.getRange(rowFound, 3).setValue(expiryTime);
    } else {
      requestSheet.appendRow([email, otp, expiryTime]);
    }
    
    sendOTPEmail(email, studentName, otp);
    
    return { 
      success: true,
      message: "OTP sent successfully"
    };
  } catch (e) {
    Logger.log("generateAndSendOTP Error: " + e.toString());
    return {
      success: false,
      message: "A system error occurred. Please try again later or contact admin."
    };
  }
}

function verifyAndSubmit(email, otp, fullName, rollNumber, pocName, studentMessage) {
  try {
    var ss = getSpreadsheet_();
    var requestSheet = ss.getSheetByName("Request");
    if (!requestSheet) {
      return { success: false, message: "System error: Request sheet not found." };
    }

    var requestData = requestSheet.getDataRange().getValues();
    var rowFound = -1;
    var storedOTP = "";
    var expiryTime = null;

    for (var i = 1; i < requestData.length; i++) { // Start from 1 to skip header
      if (requestData[i][0] === email) {
        rowFound = i + 1;
        storedOTP = requestData[i][1].toString();
        expiryTime = new Date(requestData[i][2]);
        break;
      }
    }

    if (rowFound < 0) {
      return { success: false, message: "Email not found. Please request a new OTP." };
    }

    if (new Date() > expiryTime) {
      return { success: false, message: "Your OTP has expired. Please request a new one." };
    }

    if (storedOTP !== otp) {
      return { success: false, message: "Invalid OTP. Please try again." };
    }

    // OTP is correct, now add to pending
    var pendingSheet = ss.getSheetByName("Pending");
    if (!pendingSheet) {
      return { success: false, message: "System error: Pending sheet not found." };
    }

    var timestamp = new Date();
    var newRequest = [
      timestamp,
      fullName,
      rollNumber,
      email,
      pocName,
      "Pending",
      "", // POC Message
      studentMessage || "" // Student Message
    ];
    pendingSheet.appendRow(newRequest);

    // Clear the OTP
    requestSheet.getRange(rowFound, 2).setValue("");

    // Send notification to the PoC
    sendNewRequestEmail(fullName, email, pocName, studentMessage);

    return {
      success: true,
      message: "Your verification request has been submitted successfully. You will be notified once it is reviewed."
    };
  } catch (e) {
    return {
      success: false,
      message: "An error occurred during submission: " + e.message
    };
  }
}

function sendNewRequestEmail(studentName, studentEmail, pocName, studentMessage) {
  try {
    var subject = "New Verification Request from " + studentName;
    var body = "A new verification request has been submitted by " + studentName + " (" + studentEmail + ").\n\n";
    if (studentMessage) {
      body += "Message from student: " + studentMessage + "\n\n";
    }
    body += "Please visit the PoC Dashboard to review this request.";
    
    // This is a placeholder. You need a way to get the PoC's email address.
    // Assuming you have a "POCs" sheet with names and emails.
    var ss = getSpreadsheet_();
    var pocsSheet = ss.getSheetByName("pocs");
    if (!pocsSheet) return; // Or handle error

    var pocsData = pocsSheet.getDataRange().getValues();
    var pocEmail = "";
    for (var i = 1; i < pocsData.length; i++) {
      // Assuming PoC Name is in the first column (index 0) and Email in the second (index 1)
      // Adjust column indices as per your sheet structure
      if (pocsData[i][0].toLowerCase().trim() === pocName.toLowerCase().trim()) {
        pocEmail = pocsData[i][2]; // Assuming email is in the 3rd column
        break;
      }
    }

    if (pocEmail) {
      MailApp.sendEmail(pocEmail, subject, body);
    }
  } catch (e) {
    // Optional: log error
  }
}


function getPendingRequests() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pending");
    if (!sheet) {
      throw new Error("Pending sheet not found.");
    }
    var data = sheet.getDataRange().getValues();
    // Return all data including headers, let client handle it
    // Or slice(1) to exclude headers
    return data;
  } catch (e) {
    return [];
  }
}

function testConnection() {
  return "Connection working!";
}

function getApprovedRejectedCounts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var approvedSheet = ss.getSheetByName("Approved");
  var rejectedSheet = ss.getSheetByName("Rejected");
  var approvedCount = approvedSheet ? Math.max(approvedSheet.getLastRow() - 1, 0) : 0;
  var rejectedCount = rejectedSheet ? Math.max(rejectedSheet.getLastRow() - 1, 0) : 0;
  return { approved: approvedCount, rejected: rejectedCount };
}

function updateStatus(id, email, status, message) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pendingSheet = ss.getSheetByName("Pending");
  var approvedSheet = ss.getSheetByName("Approved");
  var rejectedSheet = ss.getSheetByName("Rejected");

  var pendingData = pendingSheet.getDataRange().getValues();
  var rowFound = false;

  for (var i = 1; i < pendingData.length; i++) {
    if (pendingData[i][1] == id && pendingData[i][3] == email) {
      var rowData = pendingData[i].slice();
      var studentName = rowData[2];
      var studentEmail = rowData[3];
      var pocName = rowData[4];
      var pocEmail = rowData[5];

      if (status === "Approved") {
        var approvedRow = [rowData[0], new Date(), rowData[1], rowData[2], rowData[3], rowData[4], rowData[5], status];
        approvedSheet.insertRowBefore(2);
        approvedSheet.getRange(2, 1, 1, approvedRow.length).setValues([approvedRow]);
        var approvalMessage = "Your Superset verification request has been approved.";
        if (message && message.trim()) {
          approvalMessage += "\n\nNote from your verifier: " + message;
        }
        sendEmailToPOC(pocEmail, pocName, studentName, studentEmail, "Request Approved", "You have approved " + studentName + "'s verification request.");
        sendEmailToStudent(studentEmail, studentName, "Request Approved", approvalMessage);
      } else if (status === "Rejected") {
        var rejectedRow = [rowData[0], new Date(), rowData[1], rowData[2], rowData[3], rowData[4], rowData[5], status];
        rejectedSheet.insertRowBefore(2);
        rejectedSheet.getRange(2, 1, 1, rejectedRow.length).setValues([rejectedRow]);
        var rejectionMessage = "Your Superset verification request has been rejected. Please review the feedback on your Superset profile. Once all suggested changes are made, kindly resubmit for verification via Duperset.";
        if (message) {
          rejectionMessage += "\n\nNote from your verifier: " + message;
        }
        sendEmailToPOC(pocEmail, pocName, studentName, studentEmail, "Request Rejected", "You have rejected " + studentName + "'s verification request.");
        sendEmailToStudent(studentEmail, studentName, "Request Rejected", rejectionMessage);
      } else if (status === "Pending") {
        return "Status remains pending";
      }
      pendingSheet.deleteRow(i + 1);
      rowFound = true;
      return "Status updated successfully!";
    }
  }
  if (!rowFound) {
    return "Request not found!";
  }
}

function updatePocMessage(id, message) {
  try {
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName("Pending");
    if (!sheet) {
      throw new Error("Pending sheet not found.");
    }
    var data = sheet.getDataRange().getValues();
    var requestRow = -1;

    for (var i = 1; i < data.length; i++) {
      var sheetDate = new Date(data[i][0]);
      var idDate = new Date(id);
      if (sheetDate.getTime() === idDate.getTime()) {
        requestRow = i + 1;
        break;
      }
    }

    if (requestRow !== -1) {
      sheet.getRange(requestRow, 7).setValue(message); // Update PoC message
      return { success: true, message: "Message updated successfully." };
    } else {
      return { success: false, message: "Request not found." };
    }
  } catch (e) {
    return { success: false, message: "Error updating message: " + e.message };
  }
}

function updateRequestStatus(id, status, pocMessage) {
  try {
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName("Pending");
    if (!sheet) {
      throw new Error("Pending sheet not found.");
    }
    var data = sheet.getDataRange().getValues();
    var requestRow = -1;

    for (var i = 1; i < data.length; i++) {
      var sheetDate = new Date(data[i][0]);
      var idDate = new Date(id);
      if (sheetDate.getTime() === idDate.getTime()) {
        requestRow = i + 1;
        break;
      }
    }

    if (requestRow !== -1) {
      sheet.getRange(requestRow, 6).setValue(status); // Update status
      // Only update the message if it's not null.
      // This prevents overwriting an existing message when just changing status.
      if (pocMessage !== null && pocMessage !== undefined) {
        sheet.getRange(requestRow, 7).setValue(pocMessage); // Update PoC message
      }
      
      // Send notification email to student
      var studentEmail = sheet.getRange(requestRow, 4).getValue();
      var studentName = sheet.getRange(requestRow, 2).getValue();
      var finalPocMessage = (pocMessage !== null && pocMessage !== undefined) ? pocMessage : sheet.getRange(requestRow, 7).getValue();

      sendApprovalEmail(studentEmail, studentName, status, finalPocMessage);
      
      return { success: true, message: "Status updated successfully." };
    } else {
      return { success: false, message: "Request not found." };
    }
  } catch (e) {
    return { success: false, message: "Error updating status: " + e.message };
  }
}

function bulkUpdateRequestStatus(ids, status, message) {
  try {
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName("Pending");
    if (!sheet) {
      throw new Error("Pending sheet not found.");
    }
    var data = sheet.getDataRange().getValues();
    var updatedCount = 0;

    // Create a map for faster lookups
    var idMap = {};
    for (var j = 0; j < ids.length; j++) {
        idMap[new Date(ids[j]).getTime()] = true;
    }

    for (var i = 1; i < data.length; i++) {
      var sheetDate = new Date(data[i][0]);
      if (idMap[sheetDate.getTime()]) {
        var requestRow = i + 1;
        sheet.getRange(requestRow, 6).setValue(status); // Update status
        if (message) {
            sheet.getRange(requestRow, 7).setValue(message); // Update PoC message
        }
        
        var studentEmail = sheet.getRange(requestRow, 4).getValue();
        var studentName = sheet.getRange(requestRow, 2).getValue();
        var finalMessage = message || sheet.getRange(requestRow, 7).getValue();
        sendApprovalEmail(studentEmail, studentName, status, finalMessage);
        
        updatedCount++;
      }
    }

    if (updatedCount > 0) {
      return { success: true, message: updatedCount + " requests updated successfully." };
    } else {
      return { success: false, message: "No matching requests found to update." };
    }
  } catch (e) {
    return { success: false, message: "Error during bulk update: " + e.message };
  }
}

function sendApprovalEmail(studentEmail, studentName, status, pocMessage) {
  try {
    var subject = "Your verification request has been " + status;
    var body = "Dear " + studentName + ",\n\n";
    body += "Your verification request has been " + status.toLowerCase() + ".\n\n";
    if (pocMessage) {
      body += "Message from your verifier: " + pocMessage + "\n\n";
    }
    body += "Thank you,\nPlaceCom";

    GmailApp.sendEmail(studentEmail, subject, body);
  } catch (e) {
    // Handle email sending error
  }
}

function sendEmailToStudent(email, name, subject, message) {
  try {
    var folderLink = "https://drive.google.com/drive/folders/1y-LXYX-_E3mx5DGlqQYxgZG0pXdYmu9I";
    var htmlBody = '<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e1e1e1; border-radius: 5px;">' + 
                   '<div style="text-align: center; margin-bottom: 20px;">' + 
                     '<h2 style="color: #2c3e50; margin-bottom: 5px;">Superset Verification System</h2>' + 
                     '<div style="height: 3px; background-color: #3498db; margin: 0 auto;"></div>' + 
                   '</div>' + 
                   '<p style="font-size: 16px; color: #2c3e50;">Hello ' + name + ',</p>' + 
                   '<div style="background-color: #f9f9f9; padding: 15px; border-left: 4px solid #3498db; margin: 20px 0;">' + 
                     '<p style="font-size: 16px; color: #2c3e50; margin: 0;">' + message + '</p>' + 
                   '</div>' + 
                   '<div style="text-align: center; margin: 30px 0;">' + 
                     '<a href="' + folderLink + '" target="_blank" style="background-color: #28a745; color: white; padding: 12px 25px; text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block;">Resume Building Resources</a>' + 
                   '</div>' + 
                   '<p style="font-size: 14px; color: #7f8c8d; margin-top: 30px;">Best regards,<br>PlaceCom</p>' + 
                 '</div>';
    var plainText = "Hello " + name + ",\n\n" + message + "\n\nAccess Resume Building Resources here: " + folderLink + "\n\nBest regards,\nPlaceCom";
    GmailApp.sendEmail(email, "Superset Verification: " + subject, plainText, {htmlBody: htmlBody});
    return true;
  } catch (e) {
    return false;
  }
}

function sendEmailToPOC(pocEmail, pocName, studentName, studentEmail, subject, message) {
  try {
    var portalLink = ScriptApp.getService().getUrl() + "?poc=true";
    var htmlBody = '<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e1e1e1; border-radius: 5px;"><div style="text-align: center; margin-bottom: 20px;"><h2 style="color: #2c3e50; margin-bottom: 5px;">Superset Verification System</h2><div style="height: 3px; background-color: #3498db; margin: 0 auto;"></div></div><p style="font-size: 16px; color: #2c3e50;">Hello ' + pocName + ',</p><div style="background-color: #f9f9f9; padding: 15px; border-left: 4px solid #3498db; margin: 20px 0;"><p style="font-size: 16px; color: #2c3e50; margin: 0;">' + message + '</p></div><div style="margin: 20px 0; padding: 15px; border: 1px solid #e1e1e1; border-radius: 5px;"><h3 style="color: #2c3e50; margin-top: 0; font-size: 18px;">Student Details</h3><p style="margin: 5px 0; font-size: 15px;"><strong>Name:</strong> ' + studentName + '<br><strong>Email:</strong> ' + studentEmail + '</p></div><div style="text-align: center; margin: 25px 0;"><a href="' + portalLink + '" style="background-color: #3498db; color: white; padding: 12px 25px; text-decoration: none; border-radius: 4px; font-weight: bold; display: inline-block;">Access Verification Portal</a></div><p style="font-size: 14px; color: #7f8c8d; margin-top: 30px;">Best regards,<br>PlaceCom</p></div>';
    var plainText = "Hello " + pocName + ",\n\n" + message + "\n\nStudent Details:\nName: " + studentName + "\nEmail: " + studentEmail + "\n\nPlease use the following link to access the verification portal:\n" + portalLink + "\n\nBest regards,\nPlaceCom";
    GmailApp.sendEmail(pocEmail, "Superset Verification: " + subject, plainText, {htmlBody: htmlBody});
    return true;
  } catch (e) {
    return false;
  }
}
