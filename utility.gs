function getSheetByGid(spreadsheet, gid){
    gid = +gid || 0;
    var res_ = undefined;
    var sheets_ = spreadsheet.getSheets();
    for(var i = sheets_.length; i--; ){
        if(sheets_[i].getSheetId() === gid){
            res_ = sheets_[i];
            break;
        }
    }
    return res_;
}


function test(){
  var formData = {
    'title': 'Ithelp Ironman 9',
    'name': 'Bacon'
};
var options = {
    'method': 'post',
    'payload': formData,
    'headers': {
        Cookie:  'session=aXRoZWxwaXJvbm1hbjl0aA;'
    }
};
var res = UrlFetchApp.fetch("https://google.com", options);
console.log(res);
}