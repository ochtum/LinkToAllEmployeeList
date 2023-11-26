var MICROSOFT_BASE_URL = 'https://graph.microsoft.com/v1.0/';
var CLIENT_ID = 'YOUR_MICROSOFT-ENTRA_CLIENT_ID';
var TENANT_ID = 'YOUR_MICROSOFT-ENTRA_TENANT_ID';
var CLIENT_SECRET = 'YOUR_MICROSOFT-ENTRA_CLIENT_SECRET';
var SITE_ID = 'YOUR_MICROSOFT365_SHAREPOINT_SITE_ID';
var LIST_ID = 'YOUR_MICROSOFT365_LISTS_LIST_ID';

var GOOGLE_BASE_URL = 'https://www.googleapis.com/';
var GOOGLE_OATH_URL = 'https://oauth2.googleapis.com/token';
var PRIVATE_KEY = "YOUR_GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY";
var CLIENT_EMAIL = "YOUR_GOOGLE_SERVICE_ACCOUNT_CLIENT_EMAIL";
var API_KEY = 'YOUR_GOOGLE_API_KEY';

function main() {
    var now = new Date();
    now.setSeconds(0);
    now.setMilliseconds(0);

    
    var result = getListsData(now);
    if (result.length == 0) {
        Logger.log('There were no objects to be processed.');
    } else {
        setGoogleDrivePermission(result);
    }
}

// ++++ Microsoft Graph API process ++++
/*
* get Microsoft OAth2.0 token
* return access_token
*/
function getMicrosoftOAthToken() {

    var options = {
        method: 'post',
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded'
        },
        payload: {
          grant_type: 'client_credentials',
          scope: 'https://graph.microsoft.com/.default',
          client_id: CLIENT_ID,
          client_secret: CLIENT_SECRET
        }
    };

    var response = UrlFetchApp.fetch('https://login.microsoftonline.com/' + TENANT_ID + '/oauth2/v2.0/token', options);
    return JSON.parse(response.getContentText()).access_token;
}

/*
* get Microsoft Lists Data
* return Microsoft Lists data
*/
function getMicrosoftListsData() {
    var access_token = getMicrosoftOAthToken();

    var url = MICROSOFT_BASE_URL + 'sites/' + SITE_ID + '/lists/' + LIST_ID + '/items?$expand=fields';
    var options = {
        method: 'get',
        headers: {
            'Authorization': 'Bearer ' + access_token
        }
    };
    var response = UrlFetchApp.fetch(url, options);
    return JSON.parse(response.getContentText());
}

/*
* Create Google Drive Setting List
* param @now gas execute time
* return Google Drive Setting List
*/
function getListsData(now) {
    var listsJsonData = getMicrosoftListsData();
    var listsData = [];

    listsJsonData.value.forEach(function(listData) {        
        if (isWithinFiveMinutes(listData.fields['Modified'], now)) {
            //Only settings that change updated data and shared drives within 5 minutes before execution are subject to acquisition.

            listsData.push(listData.fields);
            Logger.log(listData.fields);
        }
    });
    return listsData;
}

// ++++ Google API process ++++
/*
* get Google OAuth2.0 token
* return access_token
*/
function getGoogleOAthToken() {
    //scope setting
    var scope = GOOGLE_BASE_URL + 'auth/drive';
    var jwt = {
      alg:'RS256',
      typ:'JWT'
    };
    var claimSet = {
      iss: CLIENT_EMAIL,
      scope: scope,
      aud: GOOGLE_OATH_URL,
      exp: Math.floor(Date.now() / 1000) + 3600,
      iat: Math.floor(Date.now() / 1000)
    };
    var key = PRIVATE_KEY;
    var encodedJwt = Utilities.base64EncodeWebSafe(JSON.stringify(jwt)) + '.'
                    + Utilities.base64EncodeWebSafe(JSON.stringify(claimSet));
    var signature = Utilities.computeRsaSha256Signature(encodedJwt, key);
    var jwtSigned =  encodedJwt + '.' + Utilities.base64EncodeWebSafe(signature);
  
    var options = {
      method: 'post',
      headers: {'Content-Type': 'application/x-www-form-urlencoded'},
      payload: {
        grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
        assertion: jwtSigned
      }
    };
    var response = UrlFetchApp.fetch(GOOGLE_OATH_URL, options);
    return JSON.parse(response.getContentText()).access_token;
}

/*
* set Google Drive permission
* param @microsoftListsData Microsoft Lists data
*/
function setGoogleDrivePermission(microsoftListsData) {
    var access_token = getGoogleOAthToken();

    var headers = {
        "Authorization": "Bearer " + access_token,
        "Accept": "application/json",
        "Content-Type": "application/json"
    };

    var folderList = getFolderList();
    microsoftListsData.forEach(function(listData) {
        var newPermission = {
            "role": "organizer",
            "type": "user",
            "emailAddress": listData['USER_EMAIL_ADDRESS']
        };

        // Shared Drive Add Permission
        var postOptions = {
            "method": "post",
            "headers": headers,
            "muteHttpExceptions": true,
            "payload": JSON.stringify(newPermission)            
        };

        // Shared Drive Remove Permission
        var deleteOptions = {
            "method": "delete",
            "headers": headers,
            "muteHttpExceptions": true    
        };

        folderList.forEach(function(folderId) {

            try {
                    // Shared Drive Add Permission
                    var url = GOOGLE_BASE_URL + 'drive/v3/files/' + folderId + '/permissions?sendNotificationEmail=true&supportsAllDrives=true&key=' + API_KEY;
                    var response = UrlFetchApp.fetch(url, postOptions);
                    Logger.log('Add Permission: ' + response.getContentText() + ', MailAddress: ' + listData['USER_EMAIL_ADDRESS']);

                    /*
                    var PermissionId = getPermissionId(folderId, listData['USER_EMAIL_ADDRESS'], headers);
                    if (PermissionId != '') {
                        // Shared Drive Remove Permission
                        var url = GOOGLE_BASE_URL + 'drive/v3/files/' + folderId + '/permissions/' + PermissionId + '?supportsAllDrives=true&key=' + API_KEY;
                        Logger.log('Remove Permission: ' + listData['USER_EMAIL_ADDRESS'] + ', ' + folderId);
                    }
                    */
            } catch (e) {
                Logger.log(e);
            }
        });
    });
}

/* 
* get Shared Drive Permission ID
* param @folderId folder ID
* param @mailAddress mail address
* param @headers request headers
* return PermissionId
*/
function getPermissionId(folderId, mailAddress, headers) {
    var options = {
        "method": "get",
        "headers": headers,
        "muteHttpExceptions": true
    };

    var url = GOOGLE_BASE_URL + 'drive/v3/files/' + folderId + '/permissions?supportsAllDrives=true&fields=*&key=' + API_KEY;
    var response = UrlFetchApp.fetch(url, options);
    var permissionList = JSON.parse(response.getContentText()).permissions;
    var permissionId = '';

    permissionList.forEach(function(permission) {
        if (permission.emailAddress == mailAddress) {
            Logger.log(permission.id);
            permissionId = permission.id;
            return true;
        }
    });
    return permissionId;
}

// ++++ Common Process ++++
/* 
* Determine if Microsoft Lists data update is within 5 minutes prior to gas runtime
* param @targetDateTimeString Microsoft Lists update date and time
* param @now Date and time of gas execution
* return true: after 5 min. false: before 5 min.
*/
function isWithinFiveMinutes(targetDateTimeString, now) {
    var targetDate = new Date(targetDateTimeString);

    var fiveMinutes = (5 * 60 * 1000) - 1000;
    if (now - targetDate <= fiveMinutes) {
        return true;
    } else {
        return false;
    }
}

/*
* Setting Folder List
* return Folder List
*/
function getFolderList() {
    var folderList = [];

    folderList.push('SHARE_DRIVE_FOLDER_ID');
  
    return folderList;
}
  