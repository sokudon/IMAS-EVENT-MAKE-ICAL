/**
 * Return a list of sheet names in the Spreadsheet with the given ID.
 * @param {String} a Spreadsheet ID.
 * @return {Array} A list of sheet names.
 */


var sid="1CpwNLrurUVVLX2dmMgZHU-uQC7WQfyfWqLlaiooRaN8";
var sname="EXたいま";


function doGet() {
  var ss = SpreadsheetApp.openById(sid);
  var sheets = ss.getSheetByName(sname);
  
　var last_row = sheets.getRange(1,6).getValue();
　var last_col = 6;
  var vst='BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//はんようたいまー//NONSGML v1.0//EN\r\n';
  var tmp='BEGIN:VEVENT\r\nUID:\r\nDTSTART:20200423T150000Z\r\nDTEND:20200424T150000Z\r\nSUMMARY:うづき\r\nEND:VEVENT\r\n';
  var vend="END:VCALENDAR";
var name="X-WR-CALNAME:imasgameEVENT\r\nX-WR-TIMEZONE:Asia/Tokyo\r\n";

  var values= sheets.getRange(1,1,last_row ,last_col).getValues();
 var value = JSON.parse(JSON.stringify(values));
 var tmp2="";
 var tmp3="";
  var moment = Moment.load();
  
  for(var k=1;k<value.length+1;k++){
 
 if(value[k-1][0]=="ミリシタ海外版"){
 var type="["+value[k-1][1]+"]";
 var titleraw=type + "偶像大師百萬人演唱會！ 劇場時光"; 
 var ibekaishi= moment(value[k-1][2]).add("h",1);
 var ibeowari= moment(value[k-1][3]).add("h",1);      
 }
 else if(value[k][0]=="ミリシタ海外版"){
 var   titleraw=type + "아이돌마스터 밀리언 라이브! 시어터 데이즈"; 
 var type="["+value[k][1]+"]";
 var ibekaishi= moment(value[k][2]);
 var ibeowari= moment(value[k][3]);     
 }
 else{
 var type="["+value[k][1]+"]";
 var titleraw=type + value[k][0]; 
 var ibekaishi= moment(value[k][2]);
 var ibeowari= moment(value[k][3]);     
 }
 var pendingend="";
    
  
  if(ibeowari=="" || ibeowari=="--" || ibeowari=="Invalid date"){
  ibeowari = ibekaishi;
  pendingend ="(※終了時未定,ENDisPENGING)";
  }
  
tmp2=tmp.toString().replace(/DTSTART:\d+T\d+Z/,"DTSTART:"+moment.utc(ibekaishi).format("YYYYMMDDTHHmmss[Z]"));
tmp2=tmp2.toString().replace(/DTEND:\d+T\d+Z/,"DTEND:"+moment.utc(ibeowari).format("YYYYMMDDTHHmmss[Z]"));
tmp2=tmp2.replace(/SUMMARY:うづき/,"SUMMARY:"+titleraw+pendingend);
  tmp2=tmp2.replace(/UID:/,"UID:MIRISITA"+moment.utc(ibekaishi).format("YYYYMMDDTHHmmss[Z]"));
  
    tmp3=tmp3+tmp2;
  }
  
  tmp3=vst + tmp3 +vend;
  tmp3=tmp3.replace(/BEGIN:VE/, name +"BEGIN:VE");
  
  //return ContentService.createTextOutput(tmp3).setMimeType(ContentService.MimeType.TEXT);
  return ContentService.createTextOutput(tmp3).setMimeType(ContentService.MimeType.ICAL);
  

   
    

}

//<>#"%　"#"はURI参照として、"%"はエスケープ用文字として使われます。
//除外されている記号 (RFC2396 に定義がないもの)
//以下の文字は使用できません。
// {}|\^[]`<>#"%

function hyperlink(link,a){
  link= "<a href='" + link +"' target=\"_blank\" rel=\"noopener\">" +a +"</a>";
  
  return link;
}

function wmap_getSheetsName(sheets){
  //var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var sheet_names = new Array();
  
  if (sheets.length >= 1) {  
    for(var i = 0;i < sheets.length; i++)
    {
      sheet_names.push(sheets[i].getName());
    }
  }
  return sheet_names;
}