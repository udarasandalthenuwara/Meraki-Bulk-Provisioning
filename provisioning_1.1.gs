// Function to create a custom menu when the spreadsheet is opened
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Meraki Tools') // Name of the custom menu
    .addItem('Configure Devices', 'configureDevices') // Menu item to run the function
    .addToUi(); // Add the menu to the UI
}

// Main function to configure devices
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
    
    // Extract values from row (updated column indices with Host Name)
    var serial = row[0];           // Column A: Device Serial
    var hostName = row[1];         // Column B: Host Name
    var usingStaticIp = row[2];    // Column C: Using static IP
    var ip = row[3];              // Column D: IP
    var subnetMask = row[4];      // Column E: Subnet mask
    var gateway = row[5];         // Column F: Gateway
    var vlan = row[6];            // Column G: VLAN
    var primaryDns = row[7];      // Column H: Primary DNS
    var secondaryDns = row[8];    // Column I: Secondary DNS
    
    // Get the cell in Column J (10th column, index 9) for API status
    var statusCell = configSheet.getRange(i + 1, 10);
    
    try {
      // Step 1: Update Host Name (if provided)
      if (hostName && hostName.trim() !== "") {
        var namePayload = {
          "name": hostName
        };
        var nameOptions = {
          "method": "put",
          "headers": headers,
          "payload": JSON.stringify(namePayload),
          "muteHttpExceptions": true
        };
        
        var nameResponse = UrlFetchApp.fetch(
          `${baseUrl}/devices/${serial}`,
          nameOptions
        );
        
        var nameResponseCode = nameResponse.getResponseCode();
        if (nameResponseCode !== 200) {
          throw new Error(`Host Name update failed with code ${nameResponseCode}`);
        }
      }
      
      // Step 2: Update Management Interface (static IP config)
      var configPayload = {
        "wan1": {
          "usingStaticIp": usingStaticIp,
          "staticIp": ip,
          "staticSubnetMask": subnetMask,
          "staticGatewayIp": gateway,
          "staticDns": [primaryDns, secondaryDns || ""],
          "vlan": vlan === "" ? null : vlan
        }
      };
      
      var configOptions = {
        "method": "put",
        "headers": headers,
        "payload": JSON.stringify(configPayload),
        "muteHttpExceptions": true
      };
      
      var configResponse = UrlFetchApp.fetch(
        `${baseUrl}/devices/${serial}/managementInterface`,
        configOptions
      );
      
      var configResponseCode = configResponse.getResponseCode();
      var statusMessage = "";
      
      // Interpret the response code for the config update
      switch (configResponseCode) {
        case 200:
          statusMessage = "Success (200)";
          statusCell.setFontColor("green"); // Green for success
          break;
        case 201:
          statusMessage = "Created (201)";
          statusCell.setFontColor("green"); // Green for success
          break;
        case 400:
          statusMessage = "Bad Request (400)";
          statusCell.setFontColor("red");   // Red for error
          break;
        case 404:
          statusMessage = "Not Found (404)";
          statusCell.setFontColor("red");   // Red for error
          break;
        case 429:
          statusMessage = "Rate Limit Exceeded (429)";
          statusCell.setFontColor("red");   // Red for error
          break;
        default:
          statusMessage = `Response: ${configResponseCode}`;
          statusCell.setFontColor("black"); // Default color for unexpected codes
      }
      
      // Update status in Column J
      statusCell.setValue(statusMessage);
      
    } catch (error) {
      // Log error and put descriptive error message in Column J
      Logger.log(`Error for serial ${serial}: ${error}`);
      var errorMessage = error.toString();
      var statusMessage = "Failed: ";
      
      if (errorMessage.includes("HTTP")) {
        var errorCode = errorMessage.match(/\d{3}/) ? errorMessage.match(/\d{3}/)[0] : "Unknown";
        statusMessage += `${errorCode} - ${error.message}`;
      } else {
        statusMessage += error.message;
      }
      
      // Set error status and color it red
      statusCell.setValue(statusMessage);
      statusCell.setFontColor("red"); // Red for all errors
    }
    
    // Flush changes to the sheet to display immediately
    SpreadsheetApp.flush();
    
    // Add small delay to avoid rate limiting
    Utilities.sleep(1000);
  }
}