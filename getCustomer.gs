function runCustomerDataFetch() {
  var customerIds = [1319, 1320, 1321]; // Replace with actual customer IDs
  fetchNetSuiteCustomerData(customerIds);
}
function fetchNetSuiteCustomerData(customerIds) {
  var service = getOAuthService();
  
  // Check if authorized
  if (!service.hasAccess()) {
    Logger.log('Access not authorized. Please re-run the authorization process.');
    return;
  }
  
  Logger.log('Authorization successful, fetching data...');
  
  // Validate the customerIds array
  if (!Array.isArray(customerIds) || customerIds.length === 0) {
    Logger.log('Invalid input: customerIds is not a valid array.');
    return;
  }

  // Get the active spreadsheet and sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet5'); // Adjust to your sheet name
  
  // Loop through customer IDs and fetch data for each
  customerIds.forEach(function(customerId) {
    // Define the API endpoint for retrieving customer record
    var url = 'https://XXXXXX.suitetalk.api.netsuite.com/services/rest/record/v1/customer/' + customerId;
    
    // Create request headers
    var headers = {
      Authorization: 'Bearer ' + service.getAccessToken(),
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    };
    
    try {
      var response = UrlFetchApp.fetch(url, {
        method: 'GET',
        headers: headers,
        muteHttpExceptions: true
      });
      
      var responseCode = response.getResponseCode();
      Logger.log('Response code for customer ' + customerId + ': ' + responseCode);
      
      if (responseCode == 200) {
        var jsonResponse = JSON.parse(response.getContentText());
        Logger.log('Parsed JSON for customer ' + customerId + ': ' + JSON.stringify(jsonResponse));
        
        // Find the last row in the sheet
        var lastRow = sheet.getLastRow() + 1; // Find the next available row
        
        // Insert data into the sheet at the next available row
        sheet.getRange(lastRow, 1).setValue('Customer ID');
        sheet.getRange(lastRow, 2).setValue(jsonResponse.id);
        
        sheet.getRange(lastRow + 1, 1).setValue('Company Name');
        sheet.getRange(lastRow + 1, 2).setValue(jsonResponse.companyName);
        
        sheet.getRange(lastRow + 2, 1).setValue('Email');
        sheet.getRange(lastRow + 2, 2).setValue(jsonResponse.email);
        
        sheet.getRange(lastRow + 3, 1).setValue('Phone');
        sheet.getRange(lastRow + 3, 2).setValue(jsonResponse.phone);
        
        sheet.getRange(lastRow + 4, 1).setValue('Address');
        sheet.getRange(lastRow + 4, 2).setValue(jsonResponse.defaultAddress);
        
        sheet.getRange(lastRow + 5, 1).setValue('Balance');
        sheet.getRange(lastRow + 5, 2).setValue(jsonResponse.balance);
        
        sheet.getRange(lastRow + 6, 1).setValue('Subsidiary');
        sheet.getRange(lastRow + 6, 2).setValue(jsonResponse.subsidiary.refName);
        
        sheet.getRange(lastRow + 7, 1).setValue('Terms');
        sheet.getRange(lastRow + 7, 2).setValue(jsonResponse.terms.refName);
        
        sheet.getRange(lastRow + 8, 1).setValue('Sales Rep');
        sheet.getRange(lastRow + 8, 2).setValue(jsonResponse.salesRep.refName);
        
        sheet.getRange(lastRow + 9, 1).setValue('Price Level');
        sheet.getRange(lastRow + 9, 2).setValue(jsonResponse.priceLevel.refName);
        
      } else {
        Logger.log('Request failed with status code: ' + responseCode);
      }
      
    } catch (e) {
      Logger.log('Error fetching data for customer ' + customerId + ': ' + e.message);
    }
  });
}
