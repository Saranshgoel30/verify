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
  // Forcing getActiveSpreadsheet() as per user request.
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    Logger.log('getSpreadsheet_ Error: ' + e.toString());
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
    var pocsSheet = ss.getSheetByName("pocs"); // FIX: Corrected sheet name
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
      // SCHEMA: Pending[3] is Email, Pending[6] is Status
      if (pendingData[i][3] === email && pendingData[i][6] === 'Pending') {
        return {
          success: false,
          message: "You already have a pending request. Please wait for approval."
        };
      }
    }
    
    // Check if email exists in POCs database and get details
    var pocsData = pocsSheet.getDataRange().getValues(); 
    var emailInDatabase = false;
    var studentName = "";
    var rollNumber = "";
    var pocName = "";
    
    for (var i = 1; i < pocsData.length; i++) { 
      // SCHEMA: pocs[2] is student Email
      if (pocsData[i][2] === email) {
        emailInDatabase = true;
        // SCHEMA: pocs[0]=ID, pocs[1]=Name, pocs[3]=POC Name
        rollNumber = pocsData[i][0]; 
        studentName = pocsData[i][1]; 
        pocName = pocsData[i][3];     
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
      message: "OTP sent successfully",
      // Return student details to the client
      studentName: studentName,
      rollNumber: rollNumber,
      pocName: pocName
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
      // SCHEMA: Request[0]=Email, Request[1]=OTP, Request[2]=Date
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
    var pocsSheet = ss.getSheetByName("pocs");
    if (!pendingSheet || !pocsSheet) {
      return { success: false, message: "System error: Pending or PoCs sheet not found." };
    }

    // Find PoC Email by looking up the student's email
    var pocsData = pocsSheet.getDataRange().getValues();
    var pocEmail = "";
    for (var i = 1; i < pocsData.length; i++) {
        // SCHEMA: pocs[2] is student Email, pocs[4] is POC Mail ID
        if (pocsData[i][2] && typeof pocsData[i][2] === 'string' && pocsData[i][2].toLowerCase().trim() === email.toLowerCase().trim()) {
            pocEmail = pocsData[i][4]; // PoC Email is in Column E
            break;
        }
    }

    if (!pocEmail) {
        Logger.log("verifyAndSubmit: Could not find PoC email for student: " + email);
        return { success: false, message: "System error: Your assigned PoC was not found. Please contact the administrator." };
    }

    var timestamp = new Date();
    // SCHEMA: Pending -> [Date, ID, Name, Email, POC, POC Mail ID, Status, Message]
    var newRequest = [
      timestamp,
      rollNumber,
      fullName,
      email,
      pocName,
      pocEmail,
      "Pending",
      studentMessage || "" // Student's message
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

function sendOTPEmail(email, studentName, otp) {
  var subject = "Your Verification OTP";
  var message = "Your OTP for student verification is: " + otp + ". This OTP is valid for 10 minutes.";
  // Using the enhanced email function for consistency and better formatting
  sendEmailToStudent(email, studentName, subject, message);
}

function sendNewRequestEmail(studentName, studentEmail, pocName, studentMessage) {
  try {
    var subject = "New Verification Request from " + studentName;
    var message = "A new verification request has been submitted by " + studentName + " (" + studentEmail + ").";
    if (studentMessage) {
      message += "\n\nMessage from student: " + studentMessage;
    }
    
    var ss = getSpreadsheet_();
    var pocsSheet = ss.getSheetByName("pocs");
    if (!pocsSheet) {
        Logger.log("sendNewRequestEmail Error: 'pocs' sheet not found.");
        return; 
    }

    var pocsData = pocsSheet.getDataRange().getValues();
    var pocEmail = "";
    for (var i = 1; i < pocsData.length; i++) {
      // SCHEMA: pocs[2] is student Email, pocs[4] is POC Mail ID
      if (pocsData[i][2] && typeof pocsData[i][2] === 'string' && pocsData[i][2].toLowerCase().trim() === studentEmail.toLowerCase().trim()) {
        pocEmail = pocsData[i][4]; // PoC Email is in Column E
        break;
      }
    }

    if (pocEmail) {
      // Use the modern sendEmailToPOC function
      sendEmailToPOC(pocEmail, pocName, studentName, studentEmail, subject, message);
    } else {
      Logger.log("sendNewRequestEmail Error: PoC Email not found for PoC Name: " + pocName);
    }
  } catch (e) {
    Logger.log("sendNewRequestEmail Error: " + e.toString());
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
    // SCHEMA: Pending[3] is Email, Pending[0] is Date of Request
    if (pendingData[i][3] == email && new Date(pendingData[i][0]).getTime() == new Date(id).getTime()) {
      var rowData = pendingData[i].slice();
      var studentName = rowData[2]; // Pending[2] is Name
      var studentEmail = rowData[3]; // Pending[3] is Email
      var studentMessage = rowData[7]; // Pending[7] is Student Message

      var targetSheet;
      if (status === "Approved") {
        targetSheet = approvedSheet;
      } else if (status === "Rejected") {
        targetSheet = rejectedSheet;
      }

      if (targetSheet) {
        // SCHEMA: Approved/Rejected -> [Req Date, Action Date, ID, Name, Email, POC, POC Mail, Status, Student Msg, POC Msg]
        var newRow = [
          rowData[0], // Date of Request
          new Date(), // Date Approved/Rejected
          rowData[1], // ID
          rowData[2], // Name
          rowData[3], // Email
          rowData[4], // POC
          rowData[5], // POC Mail ID
          status,     // Status
          studentMessage, // Student Message
          message || ""   // Note from POC
        ];
        targetSheet.appendRow(newRow);
        
        sendApprovalEmail(studentEmail, studentName, status, message);
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
      // SCHEMA: Pending[0] is Date of Request
      var sheetDate = new Date(data[i][0]);
      var idDate = new Date(id);
      if (sheetDate.getTime() === idDate.getTime()) {
        requestRow = i + 1;
        break;
      }
    }

    if (requestRow !== -1) {
      // This function is intended to add a note from the PoC.
      // Per the schema, the "Pending" sheet does not have a "Note from POC" column.
      // This action is now handled correctly by the `updateStatus` function, which adds the note
      // when moving the request to the Approved/Rejected sheet.
      // This function is therefore redundant and potentially confusing.
      // Returning a message to indicate this.
      return { success: false, message: "PoC messages can only be added upon approval or rejection." };
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
    // This function is deprecated by the more robust `updateStatus` function.
    // It is not aligned with the new sheet structure and workflow.
    // It is recommended to remove calls to this function from the client-side.
    // For now, it will do nothing to prevent data corruption.
    Logger.log("Deprecated function updateRequestStatus called. No action taken.");
    return { success: false, message: "This function is deprecated. Please use the main dashboard actions." };
  } catch (e) {
    return { success: false, message: "Error updating status: " + e.message };
  }
}

function bulkUpdateRequestStatus(ids, status, message) {
  try {
    var ss = getSpreadsheet_();
    var pendingSheet = ss.getSheetByName("Pending");
    var approvedSheet = ss.getSheetByName("Approved");
    var rejectedSheet = ss.getSheetByName("Rejected");

    if (!pendingSheet || !approvedSheet || !rejectedSheet) {
      throw new Error("One or more required sheets (Pending, Approved, Rejected) not found.");
    }
    
    var data = pendingSheet.getDataRange().getValues();
    var updatedCount = 0;
    var rowsToDelete = [];

    // Create a map for faster lookups
    var idMap = {};
    for (var j = 0; j < ids.length; j++) {
        idMap[new Date(ids[j]).getTime()] = true;
    }

    // First, identify rows to move and prepare new rows
    var newRows = [];
    for (var i = 1; i < data.length; i++) {
      var sheetDate = new Date(data[i][0]);
      if (idMap[sheetDate.getTime()]) {
        var rowData = data[i];
        
        var newRow = [
          rowData[0], // Date of Request
          new Date(), // Date Approved/Rejected
          rowData[1], // ID
          rowData[2], // Name
          rowData[3], // Email
          rowData[4], // POC
          rowData[5], // POC Mail ID
          status,     // Status
          rowData[7], // Student Message
          message || ""   // Note from POC
        ];
        newRows.push(newRow);

        // Send notification email
        var studentEmail = rowData[3];
        var studentName = rowData[2];
        sendApprovalEmail(studentEmail, studentName, status, message);
        
        rowsToDelete.push(i + 1);
        updatedCount++;
      }
    }

    // Now, append the new rows to the correct sheet
    if (newRows.length > 0) {
      var targetSheet;
      if (status === "Approved") {
        targetSheet = approvedSheet;
      } else if (status === "Rejected") {
        targetSheet = rejectedSheet;
      }
      if (targetSheet) {
        targetSheet.getRange(targetSheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
      }
    }

    // Finally, delete the rows from the Pending sheet, from bottom to top
    for (var k = rowsToDelete.length - 1; k >= 0; k--) {
      pendingSheet.deleteRow(rowsToDelete[k]);
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

    MailApp.sendEmail(studentEmail, subject, body);
  } catch (e) {
    Logger.log("sendApprovalEmail Error: " + e.toString());
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
    Logger.log("sendEmailToStudent Error: " + e.toString());
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
    Logger.log("sendEmailToPOC Error: " + e.toString());
    return false;
  }
}
