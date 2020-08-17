/** 
* 
* Create a menu option for script functions
*
* References
* https://developers.google.com/apps-script/reference/document/document-app#getui
*/

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Functions')
  .addItem('Function: Create New Ledger', 'createLedger')
  .addToUi();
}

//******************************************************************************************************************************************************************

/**
* Create ledger from Tab 1 and Tab 2, output to sheet
* https://www.reddit.com/r/sheets/comments/i5flov/any_automation_possible/
*
* Directions
* 1. Make sure words in ColA on Tab 2 and header row 1 on Tab 1 match! E.G:
*   a. Trim all column headers on Tab 1 and Payroll Transaction values in ColA on Tab 2. No extra spaces in front and behind.
*   b. Correct Allownance to Allowance on Tab 1 and Tab 2.
*   c. Capitaliza major words in ColA on Tab 2 and header row 1 on Tab 1 (e.g. "Benefit").
*   d. Change "Mobility & Housing Subsidy Allowance" on A12!'Tab 2' to "Housing Subsidy Allowance" to match Tab 1.
* 2. Once they match, run the function "onOpen" to create the function menu.
* 3. Run the function "createLedger" to create a new ledger sheet based on the content of Tab 1 and Tab 2.
* 4. After ledger is created, modify the manual entries accordingly and check for errors.
*
*/ 

function createLedger() {
  
  //  Declare variables
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
  var tab1 = spreadsheet.getSheetByName("Tab 1");
  var tab1Range = tab1.getDataRange();
  var tab1Values = tab1Range.getDisplayValues();
  var tab2 = spreadsheet.getSheetByName("Tab 2");
  var tab2Range = tab2.getDataRange();
  var tab2Values = tab2Range.getDisplayValues();
  var ledger = [];
  var date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM-dd-yyyy_HH:mm:ss");
  
  //  Get transactions
  var deptCodeString = "Department code";
  var deptBaseSalaryString = "Base Salary";
  var deptFringeBenefitString = "Fringe Benefit";
  var deptFixedAllowanceString = "Fixed Allowance";
  var deptHousingSubsidyAllowanceString = "Housing Subsidy Allowance"; //Original had typo, "Allownace", make sure you correct on sheets
  var deptOvertimeString = "Overtime";
  var deptTaxPayExpatriateString = "Tax Payment for Expatriate";
  var deptSocialSecurityString = "Social Security Contribution (Employee Share)";
  var deptIncomeTaxString = "Personal Income Tax";
  
  
  //  Get Columns from Tab 1, make sure the column headers in Tab 1 do not have spaces in front of them! 
  //  If they need to stay, modify the below values to match exactly (e.g. put a space in front of them)
  var deptCodeColumn = tab1Values[0].indexOf(deptCodeString);
  var deptBaseSalaryColumn = tab1Values[0].indexOf(deptBaseSalaryString);
  var deptFringeBenefitColumn = tab1Values[0].indexOf(deptFringeBenefitString);
  var deptFixedAllowanceColumn = tab1Values[0].indexOf(deptFixedAllowanceString);
  var deptHousingSubsidyAllowanceColumn = tab1Values[0].indexOf(deptHousingSubsidyAllowanceString);
  var deptOvertimeColumn = tab1Values[0].indexOf(deptOvertimeString);  
  var deptTaxPayExpatriateColumn = tab1Values[0].indexOf(deptTaxPayExpatriateString);
  var deptSocialSecurityColumn = tab1Values[0].indexOf(deptSocialSecurityString);
  var deptIncomeTaxColumn = tab1Values[0].indexOf(deptIncomeTaxString);
  
  //  Flatten array to grab indexes easier
  var tab2ValuesFlattened = tab2.getRange("A:A").getDisplayValues().join().split(',');
  
  //  Get Payroll Transaction rows from Tab 2
  var deptBaseSalaryRow = tab2ValuesFlattened.indexOf(deptBaseSalaryString);
  var deptFringeBenefitRow = tab2ValuesFlattened.indexOf(deptFringeBenefitString);
  var deptFixedAllowanceRow = tab2ValuesFlattened.indexOf(deptFixedAllowanceString);
  var deptHousingSubsidyAllowanceRow = tab2ValuesFlattened.indexOf(deptHousingSubsidyAllowanceString);
  var deptOvertimeRow = tab2ValuesFlattened.indexOf(deptOvertimeString);  
  var deptTaxPayExpatriateRow = tab2ValuesFlattened.indexOf(deptTaxPayExpatriateString);
  var deptSocialSecurityRow = tab2ValuesFlattened.indexOf(deptSocialSecurityString);
  var deptIncomeTaxRow = tab2ValuesFlattened.indexOf(deptIncomeTaxString);
  
  //  Get sum values
  var deptBaseSalary = 0;
  var deptFringeBenefit = 0;
  var deptFixedAllowance = 0;
  var deptHousingSubsidyAllowance = 0;
  var deptOvertime = 0;
  
  //  Get deduction values
  var deptTaxPayExpatriate = 0;
  var deptSocialSecurity = 0;
  var deptIncomeTax = 0;
  
  //  Get ledger values
  var deptCode = "";
  var transactionDate = ""; // Manual entry after ledger is generated
  var currency = "";        // Manual entry after ledger is generated
  var GLAccount = "";
  var GLDescription = "";
  var sumAmount = "";
  var deduction = "";
  var netAmount = "";
  var ledgerSheet = "";
  
  //  Add header row
  ledger.push(["Department Code", "Transaction Date", "Currency", "GL Account", "GL Description", "Sum Amount", "Deduction", "Net Amount"]);
  
  //  Parse Tab 1 to grab depts and update ledger
  for (var x = 1; x < tab1Values.length; x++){
    
    //    reset values    
    deptCode = "";
    transactionDate = "";
    currency = "";        
    GLAccount = "";
    GLDescription = "";
    sumAmount = "";
    deduction = "";
    netAmount = "";
    
    //    Grab new values
    deptCode = tab1Values[x][deptCodeColumn];
    deptBaseSalary = tab1Values[x][deptBaseSalaryColumn];
    deptFringeBenefit = tab1Values[x][deptFringeBenefitColumn];
    deptFixedAllowance = tab1Values[x][deptFixedAllowanceColumn];
    deptHousingSubsidyAllowance = tab1Values[x][deptHousingSubsidyAllowanceColumn];
    deptOvertime = tab1Values[x][deptOvertimeColumn];
    deptTaxPayExpatriate = tab1Values[x][deptTaxPayExpatriateColumn];
    deptSocialSecurity = tab1Values[x][deptSocialSecurityColumn];
    deptIncomeTax = tab1Values[x][deptIncomeTaxColumn];    
    
    //    Update ledger with Salary
    if (deptBaseSalary && deptBaseSalary != "0"){
      GLDescription = tab2Values[deptBaseSalaryRow][1];
      GLAccount = tab2Values[deptBaseSalaryRow][2];
      sumAmount = deptBaseSalary;
      netAmount = sumAmount;
      ledger.push([deptCode, transactionDate, currency, GLAccount, GLDescription, sumAmount, deduction, netAmount]);
    }
    
    //    Update ledger with Fringe Benefit
    if (deptFringeBenefit && deptFringeBenefit != "0"){
      GLDescription = tab2Values[deptFringeBenefitRow][1];
      GLAccount = tab2Values[deptFringeBenefitRow][2];
      sumAmount = deptFringeBenefit;
      netAmount = sumAmount;
      ledger.push([deptCode, transactionDate, currency, GLAccount, GLDescription, sumAmount, deduction, netAmount]);
    }
    
    //    Update ledger with Fixed Allowance
    if (deptFixedAllowance && deptFixedAllowance != "0"){
      GLDescription = tab2Values[deptFixedAllowanceRow][1];
      GLAccount = tab2Values[deptFixedAllowanceRow][2];
      sumAmount = deptFixedAllowance;
      netAmount = sumAmount;
      ledger.push([deptCode, transactionDate, currency, GLAccount, GLDescription, sumAmount, deduction, netAmount]);
    }
    
    //    Update ledger with Housing Subsidy Allowance
    if (deptHousingSubsidyAllowance && deptHousingSubsidyAllowance != "0"){
      GLDescription = tab2Values[deptHousingSubsidyAllowanceRow][1];
      GLAccount = tab2Values[deptHousingSubsidyAllowanceRow][2];
      sumAmount = deptHousingSubsidyAllowance;
      netAmount = sumAmount;
      ledger.push([deptCode, transactionDate, currency, GLAccount, GLDescription, sumAmount, deduction, netAmount]);
    }
    
    //    Update ledger with Overtime
    if (deptOvertime && deptOvertime != "0"){
      GLDescription = tab2Values[deptOvertimeRow][1];
      GLAccount = tab2Values[deptOvertimeRow][2];
      sumAmount = deptOvertime;
      netAmount = sumAmount;
      ledger.push([deptCode, transactionDate, currency, GLAccount, GLDescription, sumAmount, deduction, netAmount]);
    }
    
    //    Update ledger with Tax Payment for Expatriate
    if (deptTaxPayExpatriate && deptTaxPayExpatriate != "0"){
      GLDescription = tab2Values[deptTaxPayExpatriateRow][1];
      GLAccount = tab2Values[deptTaxPayExpatriateRow][2];
      deduction = deptOvertime;
      netAmount = "-" + deduction;
      ledger.push([deptCode, transactionDate, currency, GLAccount, GLDescription, sumAmount, deduction, netAmount]);
    }
    
    //    Update ledger with Social Security Contribution (Employee Share)
    if (deptSocialSecurity && deptSocialSecurity != "0"){
      GLDescription = tab2Values[deptSocialSecurityRow][1];
      GLAccount = tab2Values[deptSocialSecurityRow][2];
      deduction = deptSocialSecurity;
      netAmount = "-" + deduction;
      ledger.push([deptCode, transactionDate, currency, GLAccount, GLDescription, sumAmount, deduction, netAmount]);
    }
    
    //    Update ledger with Personal Income Tax
    if (deptIncomeTax && deptIncomeTax != "0"){
      GLDescription = tab2Values[deptIncomeTaxRow][1];
      GLAccount = tab2Values[deptIncomeTaxRow][2];
      deduction = deptIncomeTax;
      netAmount = "-" + deduction;
      ledger.push([deptCode, transactionDate, currency, GLAccount, GLDescription, sumAmount, deduction, netAmount]);
    }    
  }  
  
  //  Create ledger
  ledgerSheet = spreadsheet.insertSheet("Ledger_" + date);
  ledgerSheet.getRange(1, 1, ledger.length, ledger[0].length).setValues(ledger);
  
  //  Format ledger sheet
  formatLedger(ledgerSheet);
}

//******************************************************************************************************************************************************************

/**
* Make ledger all pretty
*
* @param {Object} ledgerSheet The created sheet that needs to be formatted
*
*/

function formatLedger (ledgerSheet){
  
  //  Declare variables
  var maxColumns = ledgerSheet.getMaxColumns(); 
  var lastColumn = ledgerSheet.getLastColumn();
  var maxRows = ledgerSheet.getMaxRows(); 
  var lastRow = ledgerSheet.getLastRow();
  var ledgerSheetRange = ledgerSheet.getDataRange();
  var boldArray = [];
  var horizontalArray = [];
  var subBoldArray = [];
  var subHorizontalArray = [];
  
  //  Create bold and centered text style for row 1
  for (var x = 0; x < lastColumn; x++){
    subBoldArray.push("bold");
    subHorizontalArray.push("center");
  }
  boldArray.push(subBoldArray);
  horizontalArray.push(subHorizontalArray);
  
  //Delete empty columns
  if (maxColumns - lastColumn != 0){
    ledgerSheet.deleteColumns(lastColumn + 1, maxColumns - lastColumn);
  }
  
  //Delete empty rows
  if (maxRows - lastRow > 1){
    ledgerSheet.deleteRows(lastRow + 1, maxRows - lastRow);
  }
  
  // Freezes the first row
  ledgerSheet.setFrozenRows(1);
  
  // Create alternating rows
  ledgerSheetRange.applyRowBanding();  
  
  //  Modify header row
  ledgerSheet.getRange("A1:H1").setBackground('#fbe4d5').setFontWeights(boldArray).setHorizontalAlignments(horizontalArray);
  
  //  Set border
  ledgerSheetRange.setBorder(true, true, true, true, true, true);
  
  //  Autoresize columns
  ledgerSheet.autoResizeColumns(1, lastColumn)
}



















