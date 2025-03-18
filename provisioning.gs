function configureDevices() {
  // Get the spreadsheet and sheets
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = spreadsheet.getSheetByName('Bulk Provisioning');
  var gettingStartedSheet = spreadsheet.getSheetByName('Getting started');
  
  // Get OrgID and API Key from Getting Started sheet
  var orgId = gettingStartedSheet.getRange("B1").getValue();
  var apiKey = gettingStartedSheet.getRange("B2").getValue();
  
  // Get Bulk Provisioning data
  var dataRange = configSheet.getDataRange();
  var data = dataRange.getValues();
  
  // Base URL
  var baseUrl = "https://api.meraki.com/api/v1";
  
  // Headers for API request
  var headers = {
    "X-Cisco-Meraki-API-Key": apiKey,
    "Content-Type": "application/json"
  };
  
  // Skip header row and process each device
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    // Extract values from row
    var serial = row[0];           // Device Serial
    var usingStaticIp = row[1];    // Using static IP
    var ip = row[2];              // IP
    var subnetMask = row[3];      // Subnet mask
    var gateway = row[4];         // Gateway
    var vlan = row[5];            // VLAN
    var primaryDns = row[6];      // Primary DNS
    var secondaryDns = row[7];    // Secondary DNS
    
    // Prepare payload
    var payload = {
      "wan1": {
        "usingStaticIp": usingStaticIp,
        "staticIp": ip,
        "staticSubnetMask": subnetMask,
        "staticGatewayIp": gateway,
        "staticDns": [primaryDns, secondaryDns || ""],
        "vlan": vlan === "" ? null : vlan
      }
    };
    
    // API request options
    var options = {
      "method": "put",
      "headers": headers,
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };
    
    try {
      // Make API call
      var response = UrlFetchApp.fetch(
        `${baseUrl}/devices/${serial}/managementInterface`,
        options
      );
      
      var responseCode = response.getResponseCode();
      // Update status in column I (9th column, index 8)
      configSheet.getRange(i + 1, 9).setValue(responseCode);
      
    } catch (error) {
      // Log error and put error code/message in column I
      Logger.log(`Error for serial ${serial}: ${error}`);
      configSheet.getRange(i + 1, 9).setValue(
        error.toString().includes("HTTP") ? 
        error.toString().match(/\d{3}/)[0] : 
        "Error: " + error.message
      );
    }
    
    // Add small delay to avoid rate limiting
    Utilities.sleep(1000);
  }
}