// Leafy sync - paste into Extensions > Apps Script of your Google Sheet, then Deploy > New deployment > Web app
// (Execute as: Me, Who has access: Anyone). Paste the /exec link into Leafy's "What's new" > Sheet sync.
// Generic: each app saves its data in a tab named after the app, so one deployment can serve several apps.
function doGet(e){return out(handle(e.parameter.app,null))}
function doPost(e){return out(handle(JSON.parse(e.postData.contents).app,JSON.parse(e.postData.contents)))}
function out(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
function handle(app,b){
  if(!app)return{ok:false,error:'no app'};
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const sh=ss.getSheetByName(app)||ss.insertSheet(app);
    const rev=Number(sh.getRange(1,2).getValue())||0;
    const n=sh.getLastRow();
    const cur=n>=2?sh.getRange(2,2,n-1,1).getValues().map(r=>r[0]).join(''):'';
    if(!b||b.op!=='put')return{ok:true,rev:rev,data:cur?JSON.parse(cur):null};
    if(Number(b.rev)!==rev)return{ok:false,conflict:true,rev:rev,data:cur?JSON.parse(cur):null};
    const s=JSON.stringify(b.data),chunks=[];
    for(let i=0;i<s.length;i+=40000)chunks.push([i/40000,s.substr(i,40000)]);
    sh.clearContents();
    sh.getRange(1,1,1,2).setValues([['rev',rev+1]]);
    if(chunks.length)sh.getRange(2,1,chunks.length,2).setValues(chunks);
    return{ok:true,rev:rev+1};
  }finally{lock.releaseLock()}
}
