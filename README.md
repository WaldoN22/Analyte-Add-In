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
description: 'Create an Outlook add-in that counts the number of emails received on the current day.'
---

# Outlook Email Count Add-In

## Project Scope: Outlook Plugin for Daily Email Count

### Project Overview
The objective of this project is to develop an Outlook plugin that adds a button to the Outlook interface. When clicked, this button will display the total number of emails received by the user on the current day.

### Objectives
- **Develop an Outlook plugin compatible with the latest versions of Outlook.**
- **Integrate a user-friendly button into the Outlook toolbar.**
- **Implement functionality to count and display the number of emails received in a day.**

### Deliverables
1. **Outlook Plugin**: A fully functional plugin that can be installed and used within Outlook.
2. **User Interface**: A button added to the Outlook toolbar.
3. **Email Counting Functionality**: Code to count the number of emails received in the current day.
4. **Documentation**: Detailed documentation covering installation, usage, and troubleshooting.

### Technical Requirements
- **Programming Language**: C# and Blazor
- **Development Environment**: Visual Studio.
- **Outlook API**: Utilize the Outlook REST API or Microsoft Graph API to access email data.

### Resources
- [GitHub - Office Add-ins Samples](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/blazor-add-in/outlook-blazor-add-in)

### Success Criteria
- The plugin is successfully installed and used by the target audience.
- The button accurately counts and displays the number of emails received in a day.
- Positive feedback from users regarding the plugin's functionality and ease of use.

## Features

- Count emails received on the current day.
- Display email count in a user-friendly interface within Outlook.
- Interact with Microsoft Graph API or Outlook REST API to fetch email data.

## Applies to

- Outlook on the web, Windows, and Mac.

## Prerequisites

- Microsoft 365 - Get a [free developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) that provides a renewable 90-day Microsoft 365 E5 developer subscription.

## Run the sample

1. Download or clone the [Office Add-ins samples repository](https://github.com/OfficeDev/Office-Add-in-samples).
1. Open Visual Studio 2022 and open the **Office-Add-in-samples\Samples\outlook-email-count-add-in\outlook-email-count-add-in.sln** solution.
1. In **Solution Explorer**, select the **outlook-email-count-sideloader** project. Then display the project properties (F4).
1. In the **Properties** window, set the **Start Action** to **Office Desktop Client**.
   ![Start Action](./images/StartAction.png)
1. In the **Properties** window, set the **Email Address** to the email address of the account you want to use with this sample.
   In case this was not set before running, you might also see this dialog:
   ![Start Action](./images/Connect.png)
1. To start the solution, choose **Debug** > **Start Debugging** or press **F5**.
1. When Outlook opens, choose **Home** > **Show Taskpane**.

The task pane will display the email count for the current day.

## Understand the Add-In's Functionality

This Outlook add-in is a web application that adds functionality to Outlook, allowing users to count and display the number of emails received on the current day. It leverages Microsoft Graph API or Outlook REST API to retrieve email data, and Blazor WebAssembly to build the user interface.

## Key parts of this sample

The add-in is built using Blazor WebAssembly, and it utilizes C# and JavaScript interop to interact with the Outlook APIs.

### Blazor pages

The **Pages** folder contains the Blazor pages, such as **Index.razor**. Each **.razor** page also has two code-behind pages, for example, **Index.razor.cs** and **Index.razor.js**. The C# file sets up the interop connection with the JavaScript file.

```csharp
protected override async Task OnAfterRenderAsync(bool firstRender)
{
  if (firstRender)
  {
    JSModule = await JSRuntime.InvokeAsync<IJSObjectReference>("import", "./Pages/Index.razor.js");
  }
}
