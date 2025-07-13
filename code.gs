// Batch update status for multiple requests
function batchUpdateStatus(updates) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pendingSheet = ss.getSheetByName("Pending");
  var approvedSheet = ss.getSheetByName("Approved");
  var rejectedSheet = ss.getSheetByName("Rejected");
  var pendingData = pendingSheet.getDataRange().getValues();
  var results = [];
  for (var u = 0; u < updates.length; u++) {
    var update = updates[u];
    var id = update.id;
    var email = update.email;
    var status = update.status;
    var message = update.message || "";
    var rowFound = false;
    for (var i = 1; i < pendingData.length; i++) {
      if (pendingData[i][1] == id && pendingData[i][3] == email) {
        var rowData = pendingData[i].slice();
        var studentName = rowData[2];
        var studentEmail = rowData[3];
        var pocName = rowData[4];
        var pocEmail = rowData[5];
        if (status === "Approved") {
          var approvedRow = [rowData[0], new Date(), rowData[1], rowData[2], rowData[3], rowData[4], rowData[5], status, message];
          approvedSheet.insertRowBefore(2);
          approvedSheet.getRange(2, 1, 1, approvedRow.length).setValues([approvedRow]);
          var approvalMessage = "Your Superset verification request has been approved.";
          if (message && message.trim()) {
            approvalMessage += "\n\nNote from your verifier: " + message;
          }
          sendEmailToPOC(pocEmail, pocName, studentName, studentEmail, "Request Approved", "You have approved " + studentName + "'s verification request.");
          sendEmailToStudent(studentEmail, studentName, "Request Approved", approvalMessage);
        } else if (status === "Rejected") {
          var rejectedRow = [rowData[0], new Date(), rowData[1], rowData[2], rowData[3], rowData[4], rowData[5], status, message];
          rejectedSheet.insertRowBefore(2);
          rejectedSheet.getRange(2, 1, 1, rejectedRow.length).setValues([rejectedRow]);
          var rejectionMessage = "Your Superset verification request has been rejected. Please review the feedback on your Superset profile. Once all suggested changes are made, kindly resubmit for verification via Duperset.";
          if (message) {
            rejectionMessage += "\n\nNote from your verifier: " + message;
          }
          sendEmailToPOC(pocEmail, pocName, studentName, studentEmail, "Request Rejected", "You have rejected " + studentName + "'s verification request.");
          sendEmailToStudent(studentEmail, studentName, "Request Rejected", rejectionMessage);
        }
        pendingSheet.deleteRow(i + 1);
        rowFound = true;
        break;
      }
    }
    if (rowFound) {
      results.push("ID " + id + ": Updated");
    } else {
      results.push("ID " + id + ": Not found");
    }
    // Refresh data after each update to avoid row index issues
    pendingData = pendingSheet.getDataRange().getValues();
  }
  return "Batch update complete. " + results.join(", ");
}
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
      if (pendingData[i][3] === email) {
        return {
          success: false,
          message: "You already have a pending request. Please wait for approval."
        };
      }
    }
    
    // Check if email exists in POCs database
    var pocsData = pocsSheet.getDataRange().getValues();
    var emailInDatabase = false;
    
    for (var i = 1; i < pocsData.length; i++) {
      if (pocsData[i][2] === email) {
        emailInDatabase = true;
        break;
      }
    }
    
    if (!emailInDatabase) {
      return {
        success: false,
        message: "Email ID not found in our database. Please contact the administrator."
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
    
    if (rowFound > 0) {
      requestSheet.getRange(rowFound, 2).setValue(otp);
    } else {
      requestSheet.appendRow([email, otp]);
    }
    
    // Get student name for email
    var studentName = "";
    for (var i = 1; i < pocsData.length; i++) {
      if (pocsData[i][2] === email) {
        studentName = pocsData[i][1];
        break;
      }
    }
    sendOTPEmail(email, studentName, otp);
    
    return { 
      success: true,
      message: "OTP sent successfully"
    };
  } catch (e) {
    return {
      success: false,
      message: "A system error occurred. Please try again later or contact admin."
    };
  }
}

function verifyOTPAndSubmit(email, otp, studentMessage) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var requestSheet = ss.getSheetByName("Request");
    var requestData = requestSheet.getDataRange().getValues();
    var rowFound = -1;
    var storedOTP = "";
    for (var i = 0; i < requestData.length; i++) {
      if (requestData[i][0] === email) {
        rowFound = i + 1;
        storedOTP = requestData[i][1].toString();
        break;
      }
    }
    if (rowFound < 0) {
      return {
        success: false,
        message: "Email not found. Please request a new OTP."
      };
    }
    if (storedOTP !== otp) {
      return {
        success: false,
        message: "Invalid OTP. Please try again."
      };
    }
    var result = submitEmail(email, studentMessage);
    return {
      success: true,
      message: result
    };
  } catch (e) {
    return {
      success: false,
      message: "An error occurred. Please try again later."
    };
  }
}

function submitEmail(email, studentMessage) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var requestSheet = ss.getSheetByName("Request");
  var pocsSheet = ss.getSheetByName("pocs");
  var pendingSheet = ss.getSheetByName("Pending");
  var pocsData = pocsSheet.getDataRange().getValues();
  var found = false;
  var studentId = "";
  var studentName = "";
  var pocName = "";
  var pocEmail = "";
  for (var i = 1; i < pocsData.length; i++) {
    if (pocsData[i][2] === email) {
      studentId = pocsData[i][0];
      studentName = pocsData[i][1];
      pocName = pocsData[i][3];
      pocEmail = pocsData[i][4];
      // Add message in column H (index 7)
      var newRow = [new Date()].concat(pocsData[i].slice(0, 5)).concat(["Pending", studentMessage || ""]);
      pendingSheet.appendRow(newRow);
      // Delete from Request sheet
      var requestData = requestSheet.getDataRange().getValues();
      for (var j = 0; j < requestData.length; j++) {
        if (requestData[j][0] === email) {
          requestSheet.deleteRow(j + 1);
          break;
        }
      }
      sendEmailToStudent(email, studentName, "Request Submitted", 
                        "Your verification request has been submitted successfully. Your Point of Contact (POC) will review your request shortly.");
      sendEmailToPOC(pocEmail, pocName, studentName, email, "New Verification Request", 
                     "You have received a new verification request from " + studentName + " (" + email + "). Please review and approve or reject this request at your earliest convenience.");
      found = true;
      break;
    }
  }
  if (found) {
    return "Request submitted successfully!";
  } else {
    return "Email ID not found in our database. Please contact the administrator.";
  }
}

function sendOTPEmail(email, name, otp) {
  try {
    var htmlBody = '<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e1e1e1; border-radius: 5px;"><div style="text-align: center; margin-bottom: 20px;"><h2 style="color: #2c3e50; margin-bottom: 5px;">Superset Verification System</h2><div style="height: 3px; background-color: #3498db; margin: 0 auto;"></div></div><p style="font-size: 16px; color: #2c3e50;">Hello ' + (name || "there") + ',</p><div style="background-color: #f9f9f9; padding: 15px; border-left: 4px solid #3498db; margin: 20px 0;"><p style="font-size: 16px; color: #2c3e50; margin: 0;">Here is your one-time verification code:</p></div><div style="text-align: center; margin: 30px 0;"><div style="background-color: #e8f0fe; display: inline-block; padding: 15px 40px; border-radius: 4px; letter-spacing: 10px; font-size: 32px; font-weight: bold; color: #1a73e8;">' + otp + '</div><p style="margin-top: 15px; color: #7f8c8d; font-size: 14px;">This code will expire in 10 minutes.</p></div><div style="margin: 20px 0; padding: 15px; border: 1px solid #e1e1e1; border-radius: 5px; background-color: #fafafa;"><p style="margin: 0; color: #555; font-size: 14px;">If you didn\'t request this code, please ignore this email.</p></div><p style="font-size: 14px; color: #7f8c8d; margin-top: 30px;">Best regards,<br>PlaceCom</p></div>';
    
    var plainText = "Hello " + (name || "there") + ",\n\nHere is your one-time verification code: " + otp + "\n\nThis code will expire in 10 minutes.\n\nIf you didn't request this code, please ignore this email.\n\nBest regards,\nPlaceCom";
    
    GmailApp.sendEmail(email, "Superset Verification: Your OTP Code", plainText, {htmlBody: htmlBody});
    return true;
  } catch (e) {
    return false;
  }
}

function getPendingRequests() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pending");
  if (!sheet) {
    return [];
  }

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return [];
  }

  var result = data.slice(1);
  var cleanResult = [];
  for (var i = 0; i < result.length; i++) {
    var row = [];
    for (var j = 0; j < result[i].length; j++) {
      if (result[i][j] instanceof Date) {
        row.push(result[i][j].toString());
      } else {
        row.push(result[i][j]);
      }
    }
    cleanResult.push(row);
  }
  
  return cleanResult;
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