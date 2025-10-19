function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Fee Tracker')
    .addItem('Generate Invoice', 'generateInvoice')
    .addItem('Refresh Dashboard', 'refreshDashboard')
    .addToUi();
  
  refreshDashboard();
}

function generateInvoice() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invoiceSheet = ss.getSheetByName('Invoice Generator');
  const studentsSheet = ss.getSheetByName('Students 2025');
  const settingsSheet = ss.getSheetByName('Settings');
  
  // Get settings
  const settings = {};
  const settingsData = settingsSheet.getRange('A2:B11').getValues();
  settingsData.forEach(row => {
    settings[row[0]] = row[1];
  });
  
  const ui = SpreadsheetApp.getUi();
  
  // Prompt for student phone
  const phoneResponse = ui.prompt('Generate Invoice', 'Enter student phone number:', ui.ButtonSet.OK_CANCEL);
  
  if (phoneResponse.getSelectedButton() != ui.Button.OK) return;
  
  const phone = phoneResponse.getResponseText();
  
  // Prompt for month
  const monthResponse = ui.prompt('Generate Invoice', 'Enter month (e.g., Jan, Feb, Mar...):', ui.ButtonSet.OK_CANCEL);
  
  if (monthResponse.getSelectedButton() != ui.Button.OK) return;
  
  const monthInput = monthResponse.getResponseText().toLowerCase();
  const months = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  
  const monthIndex = months.findIndex(m => m.startsWith(monthInput.substring(0, 3)));
  
  if (monthIndex === -1) {
    ui.alert('Invalid month. Please use Jan, Feb, Mar, etc.');
    return;
  }
  
  // Find student
  const studentsData = studentsSheet.getDataRange().getValues();
  const studentRow = studentsData.find(row => row[0].toString() === phone);
  
  if (!studentRow) {
    ui.alert('Student not found!');
    return;
  }
  
  const studentName = studentRow[1];
  const batch = studentRow[2];
  const fee = studentRow[3];
  
  // Increment invoice number
  const currentInvoiceNum = parseInt(settings['Last Invoice Number']) + 1;
  const invoiceNumber = `INV-2025-${String(currentInvoiceNum).padStart(3, '0')}`;
  
  // Update last invoice number in settings
  settingsSheet.getRange('B11').setValue(currentInvoiceNum);
  
  // Clear previous invoice
  invoiceSheet.clear();
  
  // Generate invoice layout
  invoiceSheet.setColumnWidth(1, 500);
  
  // Header
  invoiceSheet.getRange('A1').setValue(settings['Institute Name']).setFontSize(24).setFontWeight('bold');
  invoiceSheet.getRange('A2').setValue(settings['Address']).setFontSize(10);
  invoiceSheet.getRange('A3').setValue('Phone: ' + settings['Contact Number']).setFontSize(10);
  invoiceSheet.getRange('A4').setValue('Email: ' + settings['Email']).setFontSize(10);
  
  invoiceSheet.getRange('D1').setValue('INVOICE').setFontSize(28).setFontWeight('bold').setFontColor('#059669');
  invoiceSheet.getRange('D2').setValue('Invoice #: ' + invoiceNumber).setFontSize(10);
  invoiceSheet.getRange('D3').setValue('Date: ' + new Date().toLocaleDateString('en-IN')).setFontSize(10);
  
  // Bill To
  invoiceSheet.getRange('A7').setValue('Bill To:').setFontWeight('bold').setFontSize(12);
  invoiceSheet.getRange('A8').setValue(studentName).setFontWeight('bold');
  invoiceSheet.getRange('A9').setValue('Phone: ' + phone);
  invoiceSheet.getRange('A10').setValue('Batch: ' + batch);
  
  // Fee Details Table
  invoiceSheet.getRange('A13:D13').setValues([['Description', '', '', 'Amount']]).setFontWeight('bold').setBackground('#f3f4f6');
  invoiceSheet.getRange('A14:C14').merge().setValue('Tuition Fee - ' + monthNames[monthIndex] + ' 2025');
  invoiceSheet.getRange('D14').setValue('₹' + fee.toLocaleString('en-IN'));
  
  invoiceSheet.getRange('A16:C16').merge().setValue('Total Amount Due').setFontWeight('bold').setBackground('#f9fafb');
  invoiceSheet.getRange('D16').setValue('₹' + fee.toLocaleString('en-IN')).setFontWeight('bold').setFontSize(14);
  
  // Payment Details
  invoiceSheet.getRange('A19').setValue('Payment Details:').setFontWeight('bold').setFontSize(12);
  invoiceSheet.getRange('A20').setValue('Bank: ' + settings['Bank Account']);
  invoiceSheet.getRange('A21').setValue('UPI: ' + settings['UPI ID']);
  invoiceSheet.getRange('A22').setValue('Due Date: ' + settings['Payment Due Date']);
  
  // Footer
  invoiceSheet.getRange('A25').setValue('Thank you for your payment!').setFontSize(10).setHorizontalAlignment('center');
  invoiceSheet.getRange('A26').setValue('Powered by mazaclass.com').setFontSize(9).setFontColor('#059669').setFontWeight('bold').setHorizontalAlignment('center');
  
  // Borders
  invoiceSheet.getRange('A13:D16').setBorder(true, true, true, true, true, true);
  
  ui.alert('Invoice generated successfully! Invoice #: ' + invoiceNumber);
}

function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName('Dashboard');
  const studentsSheet = ss.getSheetByName('Students 2025');
  
  // Get current month dynamically
  const now = new Date();
  const currentMonthIndex = now.getMonth(); // 0-11
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  
  // Check if month selector exists, if not use current month
  let selectedMonth = dashboardSheet.getRange('B2').getValue();
  if (!selectedMonth || !months.includes(selectedMonth)) {
    selectedMonth = months[currentMonthIndex];
  }
  
  const selectedMonthIndex = months.indexOf(selectedMonth);
  const selectedMonthName = monthNames[selectedMonthIndex];
  
  // Clear dashboard
  dashboardSheet.clear();
  
  // Get student data
  const data = studentsSheet.getDataRange().getValues();
  const headers = data[0];
  const students = data.slice(1);
  
  // Find the status column for selected month dynamically
  const monthStatusColumn = headers.findIndex(h => h === `${selectedMonth} Status`);
  
  if (monthStatusColumn === -1) {
    dashboardSheet.getRange('A1').setValue('Error: Month column not found in Students sheet');
    return;
  }
  
  let currentPaid = 0;
  let currentPending = 0;
  let currentOverdue = 0;
  let yearTotal = 0;
  let yearPending = 0;
  
  students.forEach(student => {
    const fee = student[3] || 0;
    
    // Current/Selected month stats
    if (student[monthStatusColumn] === 'Paid') currentPaid++;
    if (student[monthStatusColumn] === 'Pending') currentPending++;
    if (student[monthStatusColumn] === 'Overdue') currentOverdue++;
    
    // Year totals - check all month status columns dynamically
    months.forEach(month => {
      const colIndex = headers.findIndex(h => h === `${month} Status`);
      if (colIndex !== -1) {
        if (student[colIndex] === 'Paid') {
          yearTotal += fee;
        } else if (student[colIndex] === 'Pending' || student[colIndex] === 'Overdue') {
          yearPending += fee;
        }
      }
    });
  });
  
  // Create dashboard header
  dashboardSheet.getRange('A1').setValue('Dashboard Overview').setFontSize(20).setFontWeight('bold');
  
  // Month selector
  dashboardSheet.getRange('A2').setValue('Select Month:').setFontWeight('bold');
  const monthDropdown = dashboardSheet.getRange('B2');
  monthDropdown.setValue(selectedMonth);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(months)
    .setAllowInvalid(false)
    .build();
  monthDropdown.setDataValidation(rule);
  
  // Add instruction
  dashboardSheet.getRange('A3').setValue('(Change month above and click Fee Tracker → Refresh Dashboard)').setFontSize(9).setFontColor('#666666');
  
  // Current/Selected Month Stats
  dashboardSheet.getRange('A5').setValue(`${selectedMonthName} 2025`).setFontSize(14).setFontWeight('bold');
  dashboardSheet.getRange('A6:D6').setValues([['Metric', 'Count', '', '']]).setFontWeight('bold').setBackground('#f3f4f6');
  dashboardSheet.getRange('A7').setValue('Paid');
  dashboardSheet.getRange('B7').setValue(currentPaid).setBackground('#d1fae5');
  dashboardSheet.getRange('A8').setValue('Pending');
  dashboardSheet.getRange('B8').setValue(currentPending).setBackground('#fef3c7');
  dashboardSheet.getRange('A9').setValue('Overdue');
  dashboardSheet.getRange('B9').setValue(currentOverdue).setBackground('#fee2e2');
  
  // Year to Date
  dashboardSheet.getRange('A12').setValue('Year to Date (2025)').setFontSize(14).setFontWeight('bold');
  dashboardSheet.getRange('A13:D13').setValues([['Metric', 'Amount (₹)', '', '']]).setFontWeight('bold').setBackground('#f3f4f6');
  dashboardSheet.getRange('A14').setValue('Total Collected');
  dashboardSheet.getRange('B14').setValue('₹' + yearTotal.toLocaleString('en-IN')).setBackground('#dbeafe');
  dashboardSheet.getRange('A15').setValue('Total Pending');
  dashboardSheet.getRange('B15').setValue('₹' + yearPending.toLocaleString('en-IN')).setBackground('#fed7aa');
  
  // Total students
  dashboardSheet.getRange('A18').setValue('Total Students: ' + students.length).setFontWeight('bold');
  
  // Last updated timestamp
  dashboardSheet.getRange('A20').setValue('Last Updated: ' + new Date().toLocaleString('en-IN')).setFontSize(9).setFontColor('#666666');
  
  // Auto-resize
  dashboardSheet.autoResizeColumns(1, 4);
}

// Optional: Auto-refresh when month dropdown changes (requires installable trigger)
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  
  // If Dashboard B2 (month selector) is edited, refresh
  if (sheet.getName() === 'Dashboard' && range.getA1Notation() === 'B2') {
    refreshDashboard();
  }
}