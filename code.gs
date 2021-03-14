var _sheetToken = '&&&&&&&&&&&&&&';
var _dataTablegId= '515344256';

function doGet() {
    return HtmlService.createHtmlOutputFromFile('example').setTitle(
        '建議資料上傳系統'
    );
}


function GetSheetData(sheetId, sheetName, range) {
    var spreadsheetId = sheetId;
    var rangeName = `${sheetName}${range}`;
    var values = Sheets.Spreadsheets.Values.get(spreadsheetId, rangeName).values;
    if (!values) {
        return "No data fond in google sheet";
    }
    return values;
}


function saveFile(obj, fileName, folderName) {
    try {
        var dropbox = '資料夾名稱';
        var folder,
            folders = DriveApp.getFoldersByName(dropbox);

        if (folders.hasNext()) {
            folder = folders.next();
        } else {
            folder = DriveApp.createFolder(dropbox);
        }

        var blob = Utilities.newBlob(Utilities.base64Decode(obj.data), obj.mimeType, fileName);


        var cfolder,
            cfolders = folder.getFoldersByName(folderName);
        if (!cfolders.hasNext()) {
            folder.createFolder(folderName).createFile(blob);
        } else {
            cfolders.next().createFile(blob);
        }
        // 照片存完後，存資料到sheet
        return 'OK';
    } catch (e) {
        console.log(e);
        return e.toString();
    }

}


function getSheetDataExample() {
    return GetSheetData('%%%%%%%%%%%%%%%', '專案', '!A2:C')
}

function getProjectList() {
    console.log(GetSheetData(this._sheetToken, '專案', '!A2:C'));
    return GetSheetData(this._sheetToken, '專案', '!A2:C');
}

//取得會計科目清單
function getAccounting() {
    var data =  GetSheetData(this._sheetToken, '會計科目', '!A2:B');
    console.log(data);
    return data ; 
}

function getLogin(account, password) {
    var userData = GetSheetData(_sheetToken, '使用者', '!A2:E');
    console.log(userData);

    var userData = userData.filter(x => x[0].toUpperCase() == account.toUpperCase());
    console.log(userData);
    if (userData.length == 0) {
        return {
            status: false,
            message: '找不到使用者...(你是誰派來的?)'
        };
    }
    if (userData[0][3].toUpperCase() == password.toUpperCase()) {
        return {
            status: true,
            message: '成功登入'
        };
    } else {
        return {
            status: false,
            message: '密碼錯誤...(請吃杏仁或找榮庭處理一下!)'
        };
    }
}

function appendFee(obj) {
  try{
    console.log(obj);
    var sheet = SpreadsheetApp.openById(_sheetToken);
    var dataSheetTable = getSheetByGid(sheet,_dataTablegId );
    dataSheetTable.appendRow([
        obj.id,
        obj.account,
        `=VLOOKUP("${obj.account}",'使用者'!$A$2:$C$999,2,false)`,
        obj.project.toString(),
        `=VLOOKUP(${obj.project.toString()},'專案'!$A$2:$C$999,3,false)`,
        `=text(${obj.accounting.toString()},"00")`,
        `=VLOOKUP(text(${obj.accounting.toString()},"00"),'會計科目'!$A$1:$B$29,2,false)`,
        obj.remark.toString(),
        obj.fee==''?0:obj.fee,
        obj.buyDate,
        obj.type,
        `=HYPERLINK("https://drive.google.com/drive/folders/${obj.fId}", "檔案連結")`
    ])
    return 'ok';
  }catch(e){
    console.log(e);
    return e.toString();
  }

}

//Creates a folder as a child of the Parent folder with the ID: FOLDER_ID
function createFolderBasic(folderID, folderName) {
    var folder = DriveApp.getFolderById(folderID);
    var newFolder = folder.createFolder(folderName);
    return newFolder.getId();
};
//Create folder if does not exists only
function createFolder(folderID, folderName) {
    var parentFolder = DriveApp.getFolderById(folderID);
    var subFolders = parentFolder.getFolders();
    var doesntExists = true;
    var newFolder = '';

    // Check if folder already exists.
    while (subFolders.hasNext()) {
        var folder = subFolders.next();

        //If the name exists return the id of the folder
        if (folder.getName() === folderName) {
            doesntExists = false;
            newFolder = folder;
            return newFolder.getId();
        };
    };
    //If the name doesn't exists, then create a new folder
    if (doesntExists == true) {
        //If the file doesn't exists
        newFolder = parentFolder.createFolder(folderName);
        return newFolder.getId();
    };
};

function start() {
    //Add your own folder ID here: 
    var FOLDER_ID = '130KbplcZX1AzJUD3vq2uAaRXmN-PaTuK';
    //Add the name of your folder here:
    var NEW_FOLDER_NAME = "The New Folder";

    var myFolderID = createFolder(FOLDER_ID, NEW_FOLDER_NAME);

    Logger.log(myFolderID);
};

//Create folder if does not exists only
function createFolder(folderID, folderName) {
    var parentFolder = DriveApp.getFolderById(folderID);
    var subFolders = parentFolder.getFolders();
    var doesntExists = true;
    var newFolder = '';

    // Check if folder already exists.
    while (subFolders.hasNext()) {
        var folder = subFolders.next();
        //If the name exists return the id of the folder
        if (folder.getName() === folderName) {
            doesntExists = false;
            newFolder = folder;
            return newFolder.getId();
        };
    };
    //If the name doesn't exists, then create a new folder
    if (doesntExists == true) {
        //If the file doesn't exists
        newFolder = parentFolder.createFolder(folderName);
        return newFolder.getId();
    };
};


function saveFiles(obj,date) {
  try{
        var baseFolderId = '%%%%%%%%%%%%%%%%%%%%%%%%%%';
        var sId = createFolder(baseFolderId, date.split('_')[0]);
        console.log(sId)
        var fId = createFolder(sId, date.split('_')[1]);
        console.log(fId)
        console.log(fId);
        var fFolder = DriveApp.getFolderById(fId);
        var blobs = obj.map(function(e) {
        var blob = Utilities.newBlob(Utilities.base64Decode(e.data), e.mimeType, e.fileName);
        fFolder.createFile(blob);
    });
//回傳資料夾ID
return {status:'ok',fId:fId};
  }catch(e){
    console.log(e);
return {message:e.toString()};
  }
}


function createZip(obj,date) {
  var baseFolderId = '%%%%%%%%%%%%%%%%%%%%%';
        var sId = createFolder(baseFolderId, date.split('_')[0]);
        console.log(sId)
        var fId = createFolder(sId, date.split('_')[1]);
        console.log(fId)
    var blobs = obj.map(function(e) {
        return Utilities.newBlob(Utilities.base64Decode(e.data), e.mimeType, e.fileName);
    });
    var zip = Utilities.zip(blobs, "收據.zip");
    var fFolder = DriveApp.getFolderById(fId);
    return    fFolder.createFile(zip).getId();
}