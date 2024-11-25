function calculateLeadScore() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  // Set the column header for Lead Score if it's not already set
  if (data[0][7] !== 'Lead Score') {
    sheet.getRange(1, 8).setValue('Lead Score');
  }
  
  // Loop through each row of the spreadsheet, starting from row 2 to skip the header
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    var companySize = row[3]; // Assuming companySize is in column D
    var budget = row[4];      // Assuming budget is in column E
    var industry = row[5];    // Assuming industry is in column F
    var urgency = row[6];     // Assuming urgency is in column G
    
    var leadScore = 0;
    
    switch (companySize) {
      case '1-50 employees':
        leadScore += 10;
        break;
      case '51-200 employees':
        leadScore += 20;
        break;
      case '201-1000 employees':
        leadScore += 30;
        break;
      case '1000+ employees':
        leadScore += 40;
        break;
    }
    
    switch (budget) {
      case 'Less than $10,000':
        leadScore += 10;
        break;
      case '$10,000 - $50,000':
        leadScore += 20;
        break;
      case '$50,001 - $100,000':
        leadScore += 30;
        break;
      case 'More than $100,000':
        leadScore += 40;
        break;
    }
    
    switch (industry) {
      case 'Technology':
        leadScore += 30;
        break;
      case 'Finance':
        leadScore += 20;
        break;
      case 'Healthcare':
        leadScore += 40;
        break;
      case 'Retail':
        leadScore += 20;
        break;
      case 'Other':
        leadScore += 10;
        break;
    }
    
    switch (urgency) {
      case 'Immediate (within 1 month)':
        leadScore += 40;
        break;
      case 'Short-term (1-3 months)':
        leadScore += 30;
        break;
      case 'Medium-term (3-6 months)':
        leadScore += 20;
        break;
      case 'Long-term (6+ months)':
        leadScore += 10;
        break;
    }
    
    sheet.getRange(i + 1, 8).setValue(leadScore); // Column H for Lead Score
  }
}
