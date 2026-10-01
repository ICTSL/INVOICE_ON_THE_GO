---

### File 3: `Code.gs`

Save this file as `Code.gs` in your repository or copy it into **Extensions > Apps Script**:

```javascript
/**
 * ICT Solutions Limited - Automated Invoice Generator, Logger & Dispatcher
 * Project: ALTS-GoogleVersion
 */

// CONFIGURATION CONSTANTS
const DRIVE_FOLDER_ID = "YOUR_GOOGLE_DRIVE_FOLDER_ID_HERE"; 
const INITIAL_BILL_NO = 1001; 

/**
 * 1. DESKTOP MENU TRIGGER
 * Automatically creates custom menu on spreadsheet open.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Invoice System")
    .addItem("Send Invoice", "generateAndDispatchInvoice")
    .addToUi();
}

/**
 * 2. MOBILE INSTALLABLE TRIGGER
 * Fires when tick box in cell H1 is checked (TRUE).
 */
function installedOnEdit(e) {
  if (!e || !e.range) return;
  
  const range = e.range;
  const sheet = range.getSheet();
  
  // Verify trigger location (INVOICE_GENERATOR tab, cell H1)
  if (sheet.getName() === "INVOICE_GENERATOR" && 
      range.getA1Notation() === "H1" && 
      (e.value === "TRUE" || e.value === true)) {
    generateAndDispatchInvoice();
  }
}

/**
 * 3. MAIN WORKFLOW FUNCTION
 * Handles validation, PDF generation, Gmail dispatch, sales logging, and form reset.
 */
function generateAndDispatchInvoice() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invoiceSheet = ss.getSheetByName("INVOICE_GENERATOR");
  const salesLogSheet = ss.getSheetByName("SALES_LOG");
  const inventorySheet = ss.getSheetByName("INVENTORY");

  if (!invoiceSheet || !salesLogSheet) {
    SpreadsheetApp.getUi().alert("Error: Required sheets not found.");
    resetTriggerCheckbox(invoiceSheet);
    return;
  }

  // Force sheet formula updates
  SpreadsheetApp.flush();

  // Validate status in cell G1
  const validationStatus = invoiceSheet.getRange("G1").getValue().toString().trim();

  if (validationStatus !== "OK") {
    SpreadsheetApp.getUi().alert(
      "Invoice cannot be sent. Please ensure customer email in cell A9 is valid and all selected items are marked 'Available' in Inventory."
    );
    resetTriggerCheckbox(invoiceSheet);
    return; 
  }

  // Fetch client credentials
  let customerEmail = invoiceSheet.getRange("A9").getValue().toString().trim();
  const customerName  = invoiceSheet.getRange("A7").getValue();
  const customerPhone = invoiceSheet.getRange("A8").getValue();

  // Map inventory data
  let inventoryMap = {};
  let priceMap = {};

  if (inventorySheet) {
    const invData = inventorySheet.getDataRange().getValues();
    for (let i = 1; i < invData.length; i++) {
      let serial = invData[i][0] ? invData[i][0].toString().trim() : "";
      if (serial !== "") {
        let brand = invData[i][1] || "";
        let model = invData[i][2] || "";
        let processor = invData[i][3] || "";
        let ram = invData[i][5] || "";
        let storage = invData[i][6] || "";
        let condition = invData[i][7] || "";
        
        let rawPrice = invData[i][8] || 0;
        let cleanPrice = Number(rawPrice.toString().replace(/[^0-9.-]+/g, "")) || 0;
        
        inventoryMap[serial] = `${brand} ${model} (${processor}, ${ram}, ${storage}, ${condition}) - S/N: ${serial}`;
        priceMap[serial] = cleanPrice;
      }
    }
  }

  // Read line items (Rows 14 to 18)
  let serialNumbers = [];
  let fullDescriptions = [];

  for (let r = 14; r <= 18; r++) {
    let rawSerial = invoiceSheet.getRange("B" + r).getValue();
    if (rawSerial && rawSerial.toString().trim() !== "") {
      let serial = rawSerial.toString().trim();
      serialNumbers.push(serial);
      
      let descInCellA = invoiceSheet.getRange("A" + r).getValue();
      let fullDesc = descInCellA || inventoryMap[serial] || serial;
      fullDescriptions.push(fullDesc);

      let qty = invoiceSheet.getRange("C" + r).getValue();
      if (!qty || qty === "" || isNaN(qty)) {
        invoiceSheet.getRange("C" + r).setValue(1);
      }

      let unitPrice = invoiceSheet.getRange("D" + r).getValue();
      if (!unitPrice || unitPrice === "" || isNaN(unitPrice)) {
        if (priceMap[serial] !== undefined) {
          invoiceSheet.getRange("D" + r).setValue(priceMap[serial]);
        }
      }
    }
  }

  // Generate Bill and Invoice numbers
  invoiceSheet.getRange("D6").setValue("BILL_NO");
  invoiceSheet.getRange("E6").setValue("INVOICE_NO");
  invoiceSheet.getRange("F6").setValue("DATE");

  let currentBillNo = invoiceSheet.getRange("D7").getValue();
  
  if (!currentBillNo || isNaN(currentBillNo)) {
    const lastRow = salesLogSheet.getLastRow();
    if (lastRow > 1) {
      const lastBill = salesLogSheet.getRange(lastRow, 6).getValue();
      currentBillNo = (!isNaN(lastBill) && lastBill !== "") ? Number(lastBill) + 1 : INITIAL_BILL_NO;
    } else {
      currentBillNo = INITIAL_BILL_NO;
    }
  }

  const currentYear = new Date().getFullYear();
  const generatedInvoiceNo = `ICTSL/${currentYear}/${currentBillNo}`;
  const todayDate = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");

  invoiceSheet.getRange("D7").setValue(currentBillNo);      
  invoiceSheet.getRange("E7").setValue(generatedInvoiceNo); 
  invoiceSheet.getRange("F7").setValue(todayDate);         

  let safeCustomerName = (customerName || "Customer")
    .toString()
    .trim()
    .replace(/[^a-zA-Z0-9\s-]/g, "") 
    .replace(/\s+/g, "_");           

  const pdfFileName = `${safeCustomerName}_${generatedInvoiceNo.replace(/\//g, "_")}.pdf`;

  SpreadsheetApp.flush();

  const subtotal = invoiceSheet.getRange("F19").getValue();
  const vat = invoiceSheet.getRange("F21").getValue();
  const grandTotal = invoiceSheet.getRange("F22").getValue();

  // Generate PDF and upload to Google Drive
  const pdfBlob = createInvoicePDF(ss.getId(), invoiceSheet.getSheetId(), pdfFileName);
  let pdfUrl = "";

  try {
    let folder;
    try {
      folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    } catch (err) {
      folder = DriveApp.getRootFolder();
    }
    const pdfFile = folder.createFile(pdfBlob).setName(pdfFileName);
    pdfUrl = pdfFile.getUrl();
  } catch (e) {
    Logger.log("Drive save error: " + e.message);
  }

  // Dispatch Email via GmailApp
  const emailSubject = `Invoice ${generatedInvoiceNo} from ICT Solutions Limited`;
  const emailBody = `Dear ${customerName},\n\nPlease find attached your invoice (${generatedInvoiceNo}) for your purchase from ICT Solutions Limited.\n\nTotal Amount: ₦${Number(grandTotal).toLocaleString('en-US', {minimumFractionDigits: 2})}\n\nThank you for partnering with us.\n\nBest regards,\nICT Solutions Limited`;

  GmailApp.sendEmail(customerEmail, emailSubject, emailBody, {
    attachments: [pdfBlob],
    name: "ICT Solutions Limited"
  });

  // Update INVENTORY status to 'Sold'
  if (inventorySheet) {
    const invData = inventorySheet.getDataRange().getValues();
    for (let i = 1; i < invData.length; i++) {
      let invSerial = invData[i][0] ? invData[i][0].toString().trim() : "";
      if (serialNumbers.includes(invSerial)) {
        inventorySheet.getRange(i + 1, 10).setValue("Sold");
      }
    }
  }

  // Append entry to SALES_LOG
  salesLogSheet.appendRow([
    customerName,
    customerPhone,
    customerEmail,
    todayDate,
    generatedInvoiceNo,
    currentBillNo,
    serialNumbers.join(", "),
    fullDescriptions.join("; "),
    subtotal,
    vat,
    grandTotal,
    pdfUrl
  ]);

  // Reset INVOICE_GENERATOR form
  invoiceSheet.getRange("D6").setValue("BILL_NO");
  invoiceSheet.getRange("E6").setValue("INVOICE_NO");
  invoiceSheet.getRange("F6").setValue("DATE");

  invoiceSheet.getRange("D7").setValue(Number(currentBillNo) + 1);
  invoiceSheet.getRange("E7").clearContent();
  invoiceSheet.getRange("A7:A9").clearContent();
  
  invoiceSheet.getRange("A14:C18").clearContent();
  invoiceSheet.getRange("E14:E18").clearContent();

  // Re-apply item lookup formulas
  for (let r = 14; r <= 18; r++) {
    invoiceSheet.getRange("A" + r).setFormula(`=IF(ISBLANK(B${r}), "", IFERROR(VLOOKUP(B${r}, INVENTORY!A2:J, 2, FALSE) & " " & VLOOKUP(B${r}, INVENTORY!A2:J, 3, FALSE) & " (" & VLOOKUP(B${r}, INVENTORY!A2:J, 4, FALSE) & ", " & VLOOKUP(B${r}, INVENTORY!A2:J, 6, FALSE) & ", " & VLOOKUP(B${r}, INVENTORY!A2:J, 7, FALSE) & ")", "Item Not Found"))`);
    invoiceSheet.getRange("D" + r).setFormula(`=IF(ISBLANK(B${r}), "", IFERROR(VALUE(SUBSTITUTE(SUBSTITUTE(VLOOKUP(B${r}, INVENTORY!A2:J, 9, FALSE), "₦", ""), ",", "")), 0))`);
    invoiceSheet.getRange("F" + r).setFormula(`=IF(OR(ISBLANK(B${r}), D${r}=""), "", ((IF(C${r}="", 1, C${r})) * D${r}) - IF(E${r}="", 0, E${r}))`);
  }

  // Re-apply summary formulas
  invoiceSheet.getRange("F22").setFormula(`=IF(SUM(F14:F18)=0, "", SUM(F14:F18))`);        
  invoiceSheet.getRange("F21").setFormula(`=IF(F22="", "", ROUND(F22 * 0.075, 2))`);       
  invoiceSheet.getRange("F19").setFormula(`=IF(OR(F22="",F21=""), "", F22 - F21)`);          

  // Reset trigger tick box
  resetTriggerCheckbox(invoiceSheet);

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert("Invoice " + generatedInvoiceNo + " sent and logged successfully!");
}

/**
 * 4. HELPER: RESET TICK BOX
 */
function resetTriggerCheckbox(sheet) {
  if (sheet) {
    sheet.getRange("H1").setValue(false);
  }
}

/**
 * 5. HELPER: GENERATE PDF BLOB
 * Renders A4 portrait PDF bounded to printable area.
 */
function createInvoicePDF(spreadsheetId, sheetId, pdfFileName) {
  const url = "https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/export?" +
    "exportFormat=pdf&format=pdf&size=A4&portrait=true&fitw=true&gridlines=false&printtitle=false&sheetnames=false&fzr=false" +
    "&gid=" + sheetId + "&r1=0&c1=0&r2=23&c2=6";

  const params = {
    method: "GET",
    headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  };

  return UrlFetchApp.fetch(url, params).getBlob().setName(pdfFileName);
}