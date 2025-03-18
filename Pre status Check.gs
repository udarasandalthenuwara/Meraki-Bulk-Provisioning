function pre_getMerakiDeviceStatus() {
  // Get the spreadsheet and sheets
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var gettingStartedSheet = spreadsheet.getSheetByName('Getting started');
  var statusSheet = spreadsheet.getSheetByName('Pre Status Check');
  
  // Get OrgID and API Key from Getting Started sheet
  var orgId = gettingStartedSheet.getRange("B1").getValue();
  var apiKey = gettingStartedSheet.getRange("B2").getValue();
  
  // API endpoint
  var url = "https://api.meraki.com/api/v1/organizations/" + orgId + "/devices/availabilities";
  
  // Set up headers for API request
  var headers = {
    "X-Cisco-Meraki-API-Key": apiKey,
    "Content-Type": "application/json"
  };
  
  // Set up parameters for API call
  var options = {
    "method": "get",
    "headers": headers,
    "muteHttpExceptions": true
  };
  
  // Initialize status counters
  var statusCounts = {
    'online': 0,
    'offline': 0,
    'dormant': 0,
    'alerting': 0
  };
  
  // Make API call
  try {
    var response = UrlFetchApp.fetch(url, options);
    var json = response.getContentText();
    var data = JSON.parse(json);
    
    // Clear existing data (keep headers)
    statusSheet.getRange(2, 1, statusSheet.getLastRow(), statusSheet.getLastColumn()).clear();
    
    // Prepare data array
    var dataArray = [];
    
    // Process each device
    for (var i = 0; i < data.length; i++) {
      var device = data[i];
      var row = [
        device.serial || '',
        device.name || '',
        device.mac || '',
        device.network.id || '',
        device.status || ''
      ];
      dataArray.push(row);
      
      // Count statuses
      var status = device.status.toLowerCase();
      if (status in statusCounts) {
        statusCounts[status]++;
      }
      
      // Apply color coding based on status
      var rowNum = i + 2;
      var statusRange = statusSheet.getRange(rowNum, 1, 1, 5);
      
      switch(status) {
        case 'online':
          statusRange.setBackground('#00FF00'); // Green
          break;
        case 'offline':
          statusRange.setBackground('#FF0000'); // Red
          break;
        case 'dormant':
          statusRange.setBackground('#FFFF00'); // Yellow
          break;
        case 'alerting':
          statusRange.setBackground('#FFA500'); // Orange
          break;
        default:
          statusRange.setBackground(null);
      }
    }
    
    // Write device data to sheet if there is any
    if (dataArray.length > 0) {
      statusSheet.getRange(2, 1, dataArray.length, 5).setValues(dataArray);
    }
    
    // Create summary table starting at G1 (column 7)
    var summaryData = [
      ['Status', 'Count'],
      ['No of Online', statusCounts.online],
      ['No of Offline', statusCounts.offline],
      ['No of Dormant', statusCounts.dormant],
      ['No of Alerting', statusCounts.alerting]
    ];
    
    // Write summary table
    statusSheet.getRange(1, 7, 5, 2).setValues(summaryData);
    
    // Format summary table
    statusSheet.getRange(1, 7, 1, 2).setFontWeight('bold'); // Header row bold
    statusSheet.getRange(1, 7, 5, 2).setBorder(true, true, true, true, true, true); // Add borders
    
    // Auto-resize columns
    statusSheet.autoResizeColumns(1, 8);
    
  } catch (e) {
    Logger.log('Error: ' + e.toString());
    SpreadsheetApp.getUi().alert('Error occurred: ' + e.toString());
  }
}

// Optional: Add a menu to run the script from the spreadsheet
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Meraki Tools')
    .addItem('Get Device Status', 'pre_getMerakiDeviceStatus')
    .addToUi();
}