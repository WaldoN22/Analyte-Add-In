---
page_type: sample
urlFragment: outlook-email-count-add-in
products:
  - office
  - office-outlook
languages:
  - C#
  - Blazor
extensions:
  contentType: samples
  technologies: 
    - Add-ins
  createdDate: '01/30/2025 11:13:00 PM'
description: 'Guide to using an Outlook add-in that counts the number of emails received on the current day.'
---

# Outlook Email Count Add-In - User Guide

## Overview
This guide provides step-by-step instructions on how to use the **Outlook Email Count Add-In**, which counts the number of emails you have received on the current day. The add-in adds a button to your Outlook interface, allowing you to easily track your email count.

### Key Features:
- Count emails received on the current day.
- Display email count within the Outlook interface.
- Integrate with Microsoft Graph API or Outlook REST API to fetch email data.

## Prerequisites

- **Microsoft 365 Account**: You'll need a Microsoft 365 subscription to use the add-in.
- **Outlook (Desktop or Web)**: This add-in is compatible with both Outlook Desktop (Windows) and Outlook Web (Outlook on the Web).

To get started, you will need to add the plugin to Outlook. Follow the instructions for the appropriate version below.

## How to Add the Add-In to Outlook

### For Outlook Desktop (Windows):
1. **Open Outlook**: Launch the Outlook application on your computer.
2. **Go to File**: In the top left corner, click on the "File" tab.
3. **Manage Add-ins**: Under the "Manage Add-ins" section, click on "Manage Add-ins" or "Options" to open the Outlook Web App.
4. **Add Custom Add-in**:
   - Click the **+** icon at the top to add a new add-in.
   - Select **Add from file** (this allows you to add an add-in file such as an .xml manifest file).
5. **Select the Add-in**: Locate and select the manifest file for the add-in on your computer.
6. **Confirm**: Once added, the add-in will appear in your list of installed add-ins. You can enable or disable it as needed.

For additional help with managing add-ins, refer to [Microsoft's link for Outlook add-ins](https://aka.ms/olksideload).

### For Outlook Web (Outlook on the Web):
1. **Go to Settings**: Open Outlook Web and click on the gear icon in the top right corner to open Settings.
2. **View All Outlook Settings**: At the bottom of the settings panel, click on "View all Outlook settings."
3. **Mail > Customize Actions**: In the settings window, go to Mail > Customize Actions.
4. **Manage Add-ins**: Scroll to the Add-ins section and click on "Manage add-ins."
5. **Add Custom Add-in**:
   - Click the **+** sign at the top and choose **Add from file**.
   - Select the manifest file on your computer, then confirm.

## Solution Overview

The **Outlook Email Count Add-In** is developed using **Blazor WebAssembly** and **C#**. It provides a user-friendly interface that allows users to view the number of emails received on the current day. It integrates with Outlook using Microsoft Graph API or the Outlook REST API.

## Key Files:
- **Manifest File**: The **manifest.json** or **manifest.xml** defines the settings and capabilities of the add-in.
- **Task Pane Files**: These files contain HTML, CSS, and JavaScript that make up the interface and interaction with Outlook.
  - **taskpane.html**: Defines the structure of the task pane.
  - **taskpane.css**: Styles the task pane's content.
  - **taskpane.js**: Contains JavaScript code that uses the Office JavaScript API to interact with Outlook.

### Example Folder Structure:
./manifest.json ./src/taskpane/taskpane.html ./src/taskpane/taskpane.css ./src/taskpane/taskpane.js



## Adding the Add-In Manually

### Step 1: Sideload the Add-In
If you're developing or testing the add-in locally, you can sideload it by following these steps:

1. **Download or Clone the Repository**: Clone the repository from GitHub and open it in Visual Studio.
2. **Set Start Action**:
   - In **Solution Explorer**, select the **outlook-email-count-sideloader** project.
   - Set the **Start Action** to **Office Desktop Client**.
3. **Set Email Address**:
   - In the **Properties** window, set the email address for the account you want to use.
4. **Debug and Start**:
   - Press **F5** to start debugging, and Outlook will open with the add-in ready to use.
5. **Show Task Pane**: In Outlook, go to the **Home** tab and select **Show Taskpane** to view the email count.

### Troubleshooting
If you encounter any issues:
1. **Ensure Task Pane is Enabled**: Go to **File > Options > Add-ins** in Outlook and enable the add-in if it's disabled.
2. **Clear Cache**: Close Outlook and run the following command in Command Prompt (as Administrator) to clear the cache:
   ```powershell
   taskkill /IM outlook.exe /F
   ipconfig /flushdns


Additional Resources

1. https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/blazor-add-in/outlook-blazor-add-in
2. https://youtu.be/pabal0sqzrM?si=y60j0D1WuQ1uZTzx
3. https://learn.microsoft.com/en-gb/office/dev/add-ins/overview/explore-with-script-lab
4. https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/outlook-quickstart-vs
5. https://learn.microsoft.com/en-us/office/dev/add-ins/overview/office-add-in-code-samples
