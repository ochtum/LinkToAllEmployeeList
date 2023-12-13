# LinkToAllEmployeeList
Sample source to reflect the process to GoogleWorkSpace based on the data of Microsoft365 Lists [All Employee List].
- Currently, it is only applied to the shared drive of GoogleDrive.
- Assuming that the trigger is executed every 5 minutes on the gas side.

## 1.Setup Procedure for Development Environment
### (1) Install Node.js
[Node.js website](https://nodejs.org/en)

### (2) Execute the following command in PowerShell to install Clasp.
```bash
npm install @google/clasp -g
```
### (3) Execute the following command in PowerShell to install npm.
```bash
npm install -g npm@latest
```

### (4) In the folder containing .clasp.json, execute the command to log in to Clasp.
```bash
clasp login
```

## 2.The following settings are required before execution
### (1) Enable Google Drive API in GCP
### (2) Create authentication credentials in GCP (API Key, OAuth 2.0 Client ID, Service Account)
### (3) Create a new app from "Application" -> "Register an app" in Microsoft Entra Management Center.
### (4) Create a client secret under "New Client Secret" in the "Client Secret" tab under "Certificates and Secrets" in the app you created (remember to note the value at this time. Once you close it or something, it will not be visible.).
### (5) Select Microsoft Graph from the "API Permissions" of the created application and add "User.Read" and "Sites.Read.All" from the application permissions.
### (6) Modify the values in the Source Code
- .clasp.json<br>
YOUR_SCRIPT_ID<br>
YOUR_ROOT_DIRECTORY<br>
- index.js<br>

The following sets out the information found in the Microsoft Entra Management Center.<br>
YOUR_MICROSOFT-ENTRA_CLIENT_ID<br>
YOUR_MICROSOFT-ENTRA_TENANT_ID<br>
YOUR_MICROSOFT-ENTRA_CLIENT_SECRET<br>
<br>
Below is a Graph Explorer to see and set up what is actually used in your Micrsoft365.<br>
YOUR_MICROSOFT365_SHAREPOINT_SITE_ID<br>
YOUR_MICROSOFT365_LISTS_LIST_ID<br>
<br>
The following should be the column names used in Lists. Note that if the column name is in Japanese, the alphanumeric symbols are used.<br>
USER_EMAIL_ADDRESS<br>
<br>
The following is set up as confirmed by GCP.<br>
YOUR_GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY<br>
YOUR_GOOGLE_SERVICE_ACCOUNT_CLIENT_EMAIL<br>
YOUR_GOOGLE_API_KEY<br>
<br>
The following sets the shared drive ID. If you want to set more than one, copy and paste "folderList.push('SHARE_DRIVE_FOLDER_ID');" to increase the number.<br>
SHARE_DRIVE_FOLDER_ID<br>

## 3.Deployment Method
### Verify the settings in .clasp.json and adjust the [rootDir] setting to match the local path where the source code is located.

### Execute the following command
```bash
clasp push
```

## 4.Materials
- Postman's MicrosoftGraphAPI collection
[MicrosoftGraphAPICollection](https://www.postman.com/microsoftgraph/workspace/microsoft-graph/collection/455214-085f7047-1bec-4570-9ed0-3a7253be148c?action=share&creator=19182434)

- Postman's MicrosoftGraphAPI preferences
[Fork environment.](https://www.postman.com/microsoftgraph/workspace/microsoft-graph/environment/455214-efbc69b2-69bd-402e-9e72-850b3a49bb21/fork)

- Microsoft Entra Management Center（旧名 Microsoft Active Directory)
https://entra.microsoft.com/#home

- Graph Explorer
https://developer.microsoft.com/en-us/graph/graph-explorer

- Microsoft Graph Rest API Reference
https://learn.microsoft.com/ja-jp/graph/api/overview?view=graph-rest-1.0

- Google Drive API v3 Reference
https://developers.google.com/drive/api/reference/rest/v3
