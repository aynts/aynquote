/**
 * ==============================================================================
 * AYE YEIK NYO CO., LTD. - GOOGLE APPS SCRIPT WEB APP FOR DRIVE & SHEETS SYNC
 * ==============================================================================
 * 
 * INSTRUCTIONS FOR DEPLOYMENT:
 * 1. Go to https://script.google.com and click "New project".
 * 2. Delete any code in Code.gs, paste this entire script, and click Save (Ctrl + S).
 * 3. Click "Deploy" > "New deployment".
 * 4. Select type: "Web app".
 * 5. Configuration:
 *    - Description: "Aye Yeik Nyo Drive Sync"
 *    - Execute as: "Me" (your Google account)
 *    - Who has access: "Anyone" (allows client-side browser uploads without login prompts)
 * 6. Click "Deploy", grant permissions when prompted.
 * 7. Copy the "Web app URL" (e.g. https://script.google.com/macros/s/.../exec).
 * 8. In the Aye Yeik Nyo Quotation/Invoice Web App, click "Drive Setup" and paste your URL.
 * ==============================================================================
 */

// Configuration Defaults
const FOLDER_NAME = "Aye Yeik Nyo - Quotations & Invoices";
const SPREADSHEET_NAME = "AYN_Master_Records_Database";

function doPost(e) {
  try {
    const rawData = e.postData.contents;
    const data = JSON.parse(rawData);

    // 1. Find or create Master Folder in Google Drive
    const parentFolder = getOrCreateFolder(FOLDER_NAME);
    
    // Subfolder by Doc Type & Year: e.g. "Quotations/2026" or "Invoices/2026"
    const docType = data.docType || "Quotation";
    const year = (data.date ? data.date.slice(0, 4) : new Date().getFullYear().toString());
    const typeFolder = getOrCreateSubFolder(parentFolder, docType + "s");
    const targetFolder = getOrCreateSubFolder(typeFolder, year);

    // 2. Decode and Save PDF File
    let fileUrl = "";
    let fileId = "";
    if (data.pdfBase64) {
      const pdfBytes = Utilities.base64Decode(data.pdfBase64.replace(/^data:application\/pdf;base64,/, ''));
      const blob = Utilities.newBlob(pdfBytes, 'application/pdf', data.filename || `${data.docId || 'Document'}.pdf`);
      const file = targetFolder.createFile(blob);
      
      // Set to anyone with link can view (public shared folder record)
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (err) {
        // Fallback if domain restricted
      }
      
      fileUrl = file.getUrl();
      fileId = file.getId();
    }

    // 3. Find or create Master Google Spreadsheet
    const spreadsheet = getOrCreateSpreadsheet(parentFolder, SPREADSHEET_NAME);
    let sheet = spreadsheet.getSheetByName(docType + "s");
    if (!sheet) {
      sheet = spreadsheet.insertSheet(docType + "s");
      // Create Standard Table Header
      sheet.appendRow([
        "Timestamp",
        "Doc ID",
        "Date",
        "Doc Type",
        "Client / Customer Name",
        "Location / Address",
        "Phone",
        "System Spec / Project",
        "Total Price (MMK)",
        "Advance (MMK)",
        "Balance Due (MMK)",
        "Validity / Due Date",
        "Drive PDF Link",
        "File ID"
      ]);
      sheet.getRange(1, 1, 1, 14).setFontWeight("bold").setBackground("#10b981").setFontColor("#ffffff");
      sheet.setFrozenRows(1);
    }

    // 4. Append Structured Record Row
    sheet.appendRow([
      new Date(),
      data.docId || "",
      data.date || "",
      data.docType || "",
      data.clientName || "",
      data.clientAddress || "",
      data.clientPhone || "",
      data.projectDesc || data.systemTitle || "",
      data.totalPrice || 0,
      data.advance || 0,
      data.grandTotal || 0,
      data.validity || "",
      fileUrl,
      fileId
    ]);

    // Format currency columns
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 9, 1, 3).setNumberFormat("#,##0");

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "Saved successfully to Google Drive & Google Sheets record!",
      fileUrl: fileUrl,
      folderUrl: targetFolder.getUrl(),
      spreadsheetUrl: spreadsheet.getUrl(),
      docId: data.docId
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "online",
    service: "Aye Yeik Nyo Google Drive & Sheets Sync API",
    time: new Date()
  })).setMimeType(ContentService.MimeType.JSON);
}

// Helpers
function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}

function getOrCreateSubFolder(parent, subName) {
  const folders = parent.getFoldersByName(subName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parent.createFolder(subName);
}

function getOrCreateSpreadsheet(folder, sheetName) {
  const files = folder.getFilesByName(sheetName);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }
  const ss = SpreadsheetApp.create(sheetName);
  const file = DriveApp.getFileById(ss.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  return ss;
}
