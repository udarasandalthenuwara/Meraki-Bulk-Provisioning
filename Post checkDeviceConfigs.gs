function Post_fetchAllMerakiManagementInterfaces() {
  // Get the active spreadsheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get OrgID and API Key from 'Getting started' sheet
  const gettingStartedSheet = spreadsheet.getSheetByName('Getting started');
  const orgId = gettingStartedSheet.getRange('B1').getValue();
  const apiKey = gettingStartedSheet.getRange('B2').getValue();
  
  // Get the target sheet and clear existing data (except header)
  const targetSheet = spreadsheet.getSheetByName('Post config check');
  targetSheet.getRange("A2:H").clear(); // Clear content below header row
  targetSheet.getRange("J2:K").clear(); // Clear summary area
  
  // Set up headers with API key
  const headers = {
    'X-Cisco-Meraki-API-Key': apiKey,
    'Content-Type': 'application/json'
  };
  
  const options = {
    'method': 'get',
    'headers': headers,
    'muteHttpExceptions': true
  };
  
  // Counters for summary
  let dhcpCount = 0;
  let staticCount = 0;
  
  try {
    // First get all devices in the organization
    const devicesUrl = `https://api.meraki.com/api/v1/organizations/${orgId}/devices`;
    const devicesResponse = UrlFetchApp.fetch(devicesUrl, options);
    const devicesJson = devicesResponse.getContentText();
    const devices = JSON.parse(devicesJson);
    
    // Start writing from row 2
    let currentRow = 2;
    
    // Loop through each device and add row-by-row
    for (let device of devices) {
      const serial = device.serial;
      
      // Get management interface for each device
      const interfaceUrl = `https://api.meraki.com/api/v1/devices/${serial}/managementInterface`;
      const interfaceResponse = UrlFetchApp.fetch(interfaceUrl, options);
      const interfaceJson = interfaceResponse.getContentText();
      const interfaceData = JSON.parse(interfaceJson);
      
      // Extract wan1 data (with fallback for missing data)
      const wan1 = interfaceData.wan1 || {};
      
      // Explicitly handle usingStaticIp
      const usingStaticIp = (typeof wan1.usingStaticIp === 'boolean') ? wan1.usingStaticIp : '';
      
      // Update counters
      if (usingStaticIp === true) {
        staticCount++;
      } else if (usingStaticIp === false) {
        dhcpCount++;
      }
      
      // Prepare row data in the specified order
      const rowData = [
        serial,                    // Device Serial
        usingStaticIp,            // Using static IP (true or false)
        wan1.staticIp || '',      // IP
        wan1.staticSubnetMask || '', // Subnet mask
        wan1.staticGatewayIp || '',  // Gateway
        wan1.vlan || '',          // VLAN
        wan1.staticDns && wan1.staticDns[0] ? wan1.staticDns[0] : '', // Primary DNS
        wan1.staticDns && wan1.staticDns[1] ? wan1.staticDns[1] : ''  // Secondary DNS
      ];
      
      // Write row immediately
      targetSheet.getRange(currentRow, 1, 1, 8).setValues([rowData]);
      currentRow++;
      
      // Add small delay to respect API rate limits
      Utilities.sleep(200); // 200ms delay between requests
      
      // Flush changes to sheet to show progress
      SpreadsheetApp.flush();
    }
    
    // Write summary: labels in J, values in K
    targetSheet.getRange("J2").setValue("No of devices using DHCP:");
    targetSheet.getRange("K2").setValue(dhcpCount);
    targetSheet.getRange("J3").setValue("No of devices using Static IP:");
    targetSheet.getRange("K3").setValue(staticCount);
    
    SpreadsheetApp.getUi().alert('Success', 
      `Data for ${devices.length} devices has been added to the sheet!\n` +
      `DHCP: ${dhcpCount}, Static IP: ${staticCount}`, 
      SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 
      `Failed to fetch data: ${error.message}`, 
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// Add custom menu
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Meraki Tools')
    .addItem('Fetch All Management Interfaces', 'Post_fetchAllMerakiManagementInterfaces')
    .addToUi();
}