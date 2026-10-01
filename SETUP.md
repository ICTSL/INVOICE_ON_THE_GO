### File 2: `SETUP_GUIDE.md`

```markdown
# Workbook Setup & Formula Reference Guide

This guide details the step-by-step creation of the Google Sheet workbook required for the ICTSL Invoice Dispatcher system.

---

## 1. Sheet: `INVOICE_GENERATOR`

This sheet serves as the primary data entry interface and customer receipt layout.

### Cell Formatting & Inputs
* **Cell A7**: Customer Name (Text)
* **Cell A8**: Customer Phone Number (Text/Number)
* **Cell A9**: Customer Email Address (Valid email format: `name@domain.com`)
* **Cell D7**: Bill Number (`BILL_NO`)
* **Cell E7**: Invoice Number (`INVOICE_NO`)
* **Cell F7**: Date (`DATE`)
* **Range B14:B18**: Item Serial Numbers (Selection or Manual Entry)
* **Range C14:C18**: Quantity (Defaults to `1` if left blank)
* **Range E14:E18**: Discount Amount per item (Optional)
* **Cell H1**: Insert **Tick box** (Go to `Insert > Tick box`)

---

### Cell Formulas

#### Cell G1 (Validation Status Formula)
Ensures customer email is valid and all selected items are available in inventory:
```excel
=IF(AND(A9<>"", ISNUMBER(MATCH("*@*.*", A9, 0)), COUNTA(B14:B18)>0, COUNTIF(B14:B18, "") + SUMPRODUCT(--(ISNUMBER(MATCH(B14:B18, FILTER(INVENTORY!A2:A, INVENTORY!J2:J="Available"), 0)))) = 5), "OK", "NOT_READY") ```

Line Items Table (Rows 14 to 18)
Apply these formulas across rows 14 through 18 (replace 14 with the respective row number):

Column A (Description - Cell A14):
```
=IF(ISBLANK(B14), "", IFERROR(VLOOKUP(B14, INVENTORY!A2:J, 2, FALSE) & " " & VLOOKUP(B14, INVENTORY!A2:J, 3, FALSE) & " (" & VLOOKUP(B14, INVENTORY!A2:J, 4, FALSE) & ", " & VLOOKUP(B14, INVENTORY!A2:J, 6, FALSE) & ", " & VLOOKUP(B14, INVENTORY!A2:J, 7, FALSE) & ")", "Item Not Found"))
```
Column D (Unit Price - Cell D14):
```
=IF(ISBLANK(B14), "", IFERROR(VALUE(SUBSTITUTE(SUBSTITUTE(VLOOKUP(B14, INVENTORY!A2:J, 9, FALSE), "₦", ""), ",", "")), 0))
```
Column F (Total Amount - Cell F14):
```
=IF(OR(ISBLANK(B14), D14=""), "", ((IF(C14="", 1, C14)) * D14) - IF(E14="", 0, E14))
```
Financial Summaries
Cell F19 (Subtotal):
```
=IF(OR(F22="",F21=""), "", F22 - F21)
```
Cell F21 (VAT - 7.5%):
```
=IF(F22="", "", ROUND(F22 * 0.075, 2))
```
Cell F22 (Grand Total):
```
=IF(SUM(F14:F18)=0, "", SUM(F14:F18))
```
2. Sheet: INVENTORY
Stores stock inventory records and availability statuses.

Table Headers (Row 1)
```
ColumnHeader NameDescriptionASERIAL_NOUnique hardware serial numberBBRANDLaptop/Device ManufacturerCMODELModel Name/NumberDPROCESSORCPU specificationsEGENERATIONHardware GenerationFRAMInstalled RAMGSTORAGEStorage capacity and typeHCONDITIONGrade/ConditionIPRICEUnit PriceJSTATUSAvailability Status (Available or Sold)
```
3. Sheet: SALES_LOG
Stores processed sales records appended by Apps Script.

Table Headers (Row 1)
```
ColumnHeader NameACUSTOMER_NAMEBPHONECEMAILDDATEEINVOICE_NOFBILL_NOGSERIAL_NUMBERSHITEM_DESCRIPTIONSISUBTOTALJVATKGRAND_TOTALLPDF_DRIVE_URL
```

4. UI Elements Setup
A. Mobile Tick Box
Select cell H1 on INVOICE_GENERATOR.

Click Insert > Tick box.

Add label in G1 or adjacent cell indicating SEND INVOICE.

B. Desktop Button (Drawing)
Go to Insert > Drawing.

Draw a rectangle shape, style it, and add text: "SEND INVOICE".

Click Save and Close.

Click the three dots menu (⋮) on the top-right of the drawing object.

Click Assign script and enter: generateAndDispatchInvoice.
Done

Next, take a look at the Apps Script file, code.gs. 
Code Architecture Explanation
onOpen(): Runs automatically when the spreadsheet is opened in a desktop web browser to inject a UI menu (Invoice System > Send Invoice).

installedOnEdit(e): Serves as the listener for an Installable On edit trigger. It detects when cell H1 on INVOICE_GENERATOR changes to TRUE. Using an installable trigger bypasses security limitations imposed on simple triggers, granting authorization for GmailApp and DriveApp.

generateAndDispatchInvoice():

Evaluates validation cell G1. If G1 !== "OK", execution halts.

Compiles inventory data, generates an incremental invoice ID (ICTSL/YYYY/BILL_NO), and exports rows 0–23 / columns A–F into an A4 PDF blob via UrlFetchApp.

Sends the PDF attachment via GmailApp and stores a copy in Google Drive.

Updates inventory item statuses to "Sold" in INVENTORY!J:J and logs sales metadata into SALES_LOG.

Clears form inputs, re-applies formula strings to prevent breaking sheet bindings, and unchecks cell H1.