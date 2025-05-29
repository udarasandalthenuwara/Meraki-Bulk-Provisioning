# Overview:


This tool is a Google Apps Script-based solution designed to streamline the management and configuration of Meraki devices within an organization. It allows network administrators to efficiently retrieve device status, check configurations, and apply bulk static IP configurations using a Google Sheets interface. By interacting with the Meraki API, this tool saves time and reduces manual effort.



https://github.com/user-attachments/assets/83c5851e-a06d-4531-bb17-f41a6dea26a8



# Prerequisites:


A Google account with access to Google Sheets and Google Apps Script.
A valid Meraki Organization ID (OrgID) and API Key.
Basic knowledge of Meraki device management and IP configurations.

# Setup and Usage:


# Getting Started:

Input your Meraki OrgID and API Key in the "Getting Started" sheet of "Meraki Bulk provisioning_1.1.xlsx".
# Step 1: Pre-Status Check

Purpose: Retrieve the current status of all devices.
Action: Click "Pre Status Check" in the "Getting Started" sheet.
Result: View device statuses (Online, Offline, etc.) in the "Pre Status Check" sheet.

# Step 2: Pre-Config Check

Purpose: Retrieve current device configurations.
Action: Click "Pre Config Check" in the "Getting Started" sheet.
Result: View configuration details like IP address, subnet mask, and DNS in the "Pre Config Check" sheet.

# Step 3: Bulk Provisioning

Purpose: Apply static IP configurations to devices in bulk.
Action: Fill in device details in the "Bulk Provisioning" sheet and click "Bulk Provisioning" in the "Getting Started" sheet.
Result: Static IP configurations are applied via the Meraki API.

# Step 4: Post-Status Check

Purpose: Verify device statuses after provisioning.
Action: Click "Post Status Check" in the "Getting Started" sheet.
Result: View updated device statuses in the "Post Status Check" sheet.

# Step 5: Post-Config Check

Purpose: Confirm correct application of static IP configurations.
Action: Click "Post Config Check" in the "Getting Started" sheet.
Result: Verify updated configuration details in the "Post Config Check" sheet.

# Notes:


Ensure the Meraki API key has necessary permissions.
Double-check entries in the "Bulk Provisioning" sheet to avoid misconfigurations.
Review the status of any devices that fail to update.

## License
This project is licensed under the MIT License - see the [LICENSE](https://github.com/udarasandalthenuwara/Meraki-Bulk-Provisioning/blob/main/License.md)file for details.
