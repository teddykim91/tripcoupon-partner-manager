const SPREADSHEET_ID='1XF3RSj63uMDt0yIxtm5aAiZZOUFwttJfloEQ8ronX0M';
const MASTER='① MASTER', PERFORMANCE='② Monthly Performance', CHECKS='⑤ Partnership Check', HISTORY='④ Change History';
const STAFF_SPREADSHEET_ID='1P032ye5-nJqAVMjs_tQ8SmR-IlxeKMqCwMHRpAyC6KY', STAFF_SHEET='Tripcoupon Affiliate';
function out_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
function key_(r,n){return (String(r||'').trim()+'||'+String(n||'').trim()).toLowerCase()}
function sheet_(name,headers){const ss=SpreadsheetApp.openById(SPREADSHEET_ID);let sh=ss.getSheetByName(name);if(!sh){sh=ss.insertSheet(name);if(headers){sh.getRange(1,1,1,headers.length).setValues([headers]);sh.setFrozenRows(1)}}return sh}
function readMaster_(){const sh=sheet_(MASTER),n=sh.getLastRow();if(n<2)return [];const v=sh.getRange(1,1,n,14).getDisplayValues(),head=v.shift();return v.filter(r=>r[1]).map(r=>{let o={};head.forEach((h,i)=>o[h]=r[i]);const ps=String(r[11]||'');o._status=(r[10]||ps.includes('Discontinued'))?'Discontinued':ps.includes('New')?'New':'Existing';return o})}
function readChecks_(){const sh=sheet_(CHECKS,['Key','Region','Partner Name (KR)','Status','Last Checked','Updated At']),n=sh.getLastRow(),o={};if(n<2)return o;sh.getRange(2,1,n-1,Math.max(8,sh.getLastColumn())).getDisplayValues().forEach(r=>{if(r[0])o[r[0]]={status:r[3],lastChecked:r[4],updatedAt:r[5],checkedBy:r[6]||'',notes:r[7]||''}});return o}
function readPerformance_() {
  var sh = sheet_(PERFORMANCE, [
    'Month','Key','Region','Partner Name (KR)','Partner Name (EN)',
    'Usage','Checked By','Checked Date','Notes','Source File','Updated At'
  ]);
  var n = sh.getLastRow();
  if (n < 2) return [];

  var rg = sh.getRange(2, 1, n - 1, 11);
  var raw = rg.getValues();
  var disp = rg.getDisplayValues();
  var result = [];

  for (var i = 0; i < raw.length; i++) {
    var r = raw[i];
    var d = disp[i];
    var month = '';
    var checkedDate = '';

    if (r[0] instanceof Date) {
      month = Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'yyyy-MM');
    } else {
      month = String(d[0] || r[0] || '').substring(0, 7);
    }

    if (r[7] instanceof Date) {
      checkedDate = Utilities.formatDate(r[7], Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else {
      checkedDate = String(d[7] || r[7] || '');
    }

    var key = String(r[1] || '');
    if (!month || !key) continue;

    result.push({
      month: month,
      key: key,
      region: String(r[2] || ''),
      partner: String(r[3] || ''),
      partnerEn: String(r[4] || ''),
      usage: Number(r[5]) || 0,
      checkedBy: String(r[6] || ''),
      checkedDate: checkedDate,
      notes: String(r[8] || ''),
      sourceFile: String(r[9] || ''),
      updatedAt: String(r[10] || '')
    });
  }

  return result;
}

function upsertPerformance_(p) {
  if (!p || !Array.isArray(p.items) || p.items.length === 0) {
    throw new Error('performance items missing');
  }

  var headers = [
    'Month','Key','Region','Partner Name (KR)','Partner Name (EN)',
    'Usage','Checked By','Checked Date','Notes','Source File','Updated At'
  ];
  var sh = sheet_(PERFORMANCE, headers);
  var n = sh.getLastRow();
  var existing = {};

  if (n >= 2) {
    var current = sh.getRange(2, 1, n - 1, 2).getDisplayValues();
    for (var i = 0; i < current.length; i++) {
      if (current[i][0] && current[i][1]) {
        existing[current[i][0] + '||' + current[i][1]] = i + 2;
      }
    }
  }

  for (var j = 0; j < p.items.length; j++) {
    var item = p.items[j];
    var usage = Number(item.usage);

    if (!isFinite(usage) || usage < 0) {
      throw new Error('Invalid usage for ' + (item.partner || item.key || 'unknown partner'));
    }

    var month = String(item.month || '');
    var partnerKey = String(item.key || '');
    if (!month || !partnerKey) {
      throw new Error('Month or partner key missing');
    }

    var checkedDate = item.checkedDate ||
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    var updatedAt = item.updatedAt || new Date().toISOString();

    var vals = [
      month,
      partnerKey,
      item.region || '',
      item.partner || '',
      item.partnerEn || '',
      Math.round(usage),
      item.checkedBy || '',
      checkedDate,
      item.notes || '',
      item.sourceFile || '',
      updatedAt
    ];

    var compositeKey = month + '||' + partnerKey;
    var row = existing[compositeKey];

    if (row) {
      var target = sh.getRange(row, 1, 1, 11);
      target.clearDataValidations();
      target.setValues([vals]);
    } else {
      row = sh.getLastRow() + 1;
      var target = sh.getRange(row, 1, 1, 11);
      target.clearDataValidations();
      target.setValues([vals]);
      existing[compositeKey] = row;
    }

    sh.getRange(row, 1).setNumberFormat('@');
    sh.getRange(row, 6).setNumberFormat('0');
    sh.getRange(row, 8).setNumberFormat('@');
  }
}

function readStaffInput_(){const ss=SpreadsheetApp.openById(STAFF_SPREADSHEET_ID),sh=ss.getSheetByName(STAFF_SHEET);if(!sh)throw new Error('Staff input sheet not found');const n=sh.getLastRow(),c=sh.getLastColumn();if(n<2)return [];const v=sh.getRange(1,1,n,c).getDisplayValues(),head=v.shift();return v.filter(r=>String(r[1]||'').trim()).map(r=>{let o={};head.forEach((h,i)=>o[h]=r[i]);return o})}
function readHistory_(){const sh=sheet_(HISTORY,['Date','Status','Region','Partner','Details','File']),n=sh.getLastRow();if(n<2)return [];return sh.getRange(2,1,n-1,6).getDisplayValues().reverse().map(r=>({date:r[0],status:r[1],region:r[2],partner:r[3],details:r[4],file:r[5]}))}
function doGet(e){try{if(((e&&e.parameter&&e.parameter.action)||'bootstrap')==='bootstrap'){let snapshot=null;try{snapshot=JSON.parse(PropertiesService.getScriptProperties().getProperty('LAST_IMPORT_SNAPSHOT')||'null')}catch(_e){}return out_({ok:true,master:readMaster_(),checks:readChecks_(),history:readHistory_(),performance:readPerformance_(),snapshot:snapshot})};return out_({ok:false,error:'Unknown action'})}catch(err){return out_({ok:false,error:String(err)})}}
function doPost(e){const lock=LockService.getScriptLock();lock.waitLock(30000);try{const p=JSON.parse((e&&e.postData&&e.postData.contents)||'{}');if(p.action==='setCheck'){setCheck_(p);return out_({ok:true})}if(p.action==='bulkConfirm'){bulkConfirm_(p);return out_({ok:true})}if(p.action==='applyImport'){applyImport_(p);return out_({ok:true})}if(p.action==='patchMasterFields'){patchMasterFields_(p);return out_({ok:true})}if(p.action==='readStaffInput'){return out_({ok:true,rows:readStaffInput_()})}if(p.action==='upsertPerformance'){upsertPerformance_(p);return out_({ok:true})}return out_({ok:false,error:'Unknown action'})}catch(err){return out_({ok:false,error:String(err)})}finally{lock.releaseLock()}}
function setCheck_(p){const sh=sheet_(CHECKS,['Key','Region','Partner Name (KR)','Status','Last Checked','Updated At']),n=sh.getLastRow();let row=0;if(n>=2){const keys=sh.getRange(2,1,n-1,1).getDisplayValues().flat(),i=keys.indexOf(p.key);if(i>=0)row=i+2}const m=readMaster_().find(x=>key_(x['Region'],x['Partner Name (KR)'])===p.key)||{},vals=[p.key,m['Region']||'',m['Partner Name (KR)']||'',p.status||'',p.lastChecked||'',p.updatedAt||new Date().toISOString()];if(row)sh.getRange(row,1,1,6).setValues([vals]);else sh.appendRow(vals);const ms=sheet_(MASTER),last=ms.getLastRow();if(last>=2){const rows=ms.getRange(2,1,last-1,14).getDisplayValues(),i=rows.findIndex(r=>key_(r[0],r[1])===p.key);if(i>=0){const status=p.status==='Confirmed'?'🟢 Active':p.status==='Check Required'?'🟡 Check Required':p.status==='New'?'🔵 New':p.status==='Discontinued'?'🔴 Discontinued':'⚫ Closed';ms.getRange(i+2,12).setValue(status);if(p.status==='Confirmed')ms.getRange(i+2,13).setValue(p.lastChecked||Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd')).setNumberFormat('yyyy-mm-dd')}}}
function bulkConfirm_(p){if(!Array.isArray(p.items)||!p.items.length)throw new Error('items missing');const sh=sheet_(CHECKS,['Key','Region','Partner Name (KR)','Status','Last Checked','Updated At','Checked By','Notes']);sh.getRange(1,1,sh.getMaxRows(),Math.max(8,sh.getMaxColumns())).clearDataValidations();SpreadsheetApp.flush();const master=readMaster_(),n=sh.getLastRow(),existing={};if(n>=2){sh.getRange(2,1,n-1,Math.max(8,sh.getLastColumn())).getDisplayValues().forEach((r,i)=>{if(r[0])existing[r[0]]=i+2})}p.items.forEach(item=>{const m=master.find(x=>key_(x['Region'],x['Partner Name (KR)'])===item.key)||{},st=String(item.status||'Confirmed').trim(),vals=[item.key,m['Region']||'',m['Partner Name (KR)']||'',st,item.lastChecked||new Date().toISOString(),item.updatedAt||new Date().toISOString(),item.checkedBy||'',item.notes||''];let row=existing[item.key];if(row)sh.getRange(row,1,1,8).setValues([vals]);else sh.appendRow(vals)});const ms=sheet_(MASTER),last=ms.getLastRow();if(last>=2){const rows=ms.getRange(2,1,last-1,14).getDisplayValues(),map={};rows.forEach((r,i)=>map[key_(r[0],r[1])]=i+2);p.items.forEach(item=>{let row=map[item.key];if(row){let st=String(item.status||'Confirmed').trim().toLowerCase(),masterStatus=st==='closed'?'⚫ Closed':st==='discontinued'?'🔴 Discontinued':'🟢 Active';let statusCell=ms.getRange(row,12);statusCell.clearDataValidations();statusCell.setValue(masterStatus);let dateCell=ms.getRange(row,13);dateCell.clearDataValidations();dateCell.setValue(item.lastChecked||Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd')).setNumberFormat('yyyy-mm-dd');let old=ms.getRange(row,14).getDisplayValue(),note=item.notes?('[확인 '+(item.checkedBy||'')+'] '+item.notes):'';if(note)ms.getRange(row,14).setValue(old?old+' / '+note:note)}})}}
function patchMasterFields_(p){if(!p||!Array.isArray(p.items))throw new Error('patch items missing');const sh=sheet_(MASTER),last=sh.getLastRow();if(last<2)return;const rows=sh.getRange(2,1,last-1,14).getDisplayValues(),rowMap={};rows.forEach((r,i)=>rowMap[key_(r[0],r[1])]=i+2);const colMap={'Partner Name (KR)':2,'Partner Name (EN)':3,'Category':4,'Contact Person':5,'Phone':6,'Email':7,'Messenger':8,'TripCoupon Benefit':9};p.items.forEach(item=>{const row=rowMap[item.key];if(!row)return;const col=colMap[item.field];if(!col)return;const cell=sh.getRange(row,col);cell.clearDataValidations();let v=item.value==null?'':item.value;cell.setValue(v)});if(Array.isArray(p.events)&&p.events.length){const hs=sheet_(HISTORY,['Date','Status','Region','Partner','Details','File']),vals=p.events.map(x=>[x.date||new Date(),x.status||'Changed',x.region||'',x.partner||'',x.details||'',x.file||p.file||'']);hs.getRange(hs.getLastRow()+1,1,vals.length,6).setValues(vals)}if(p.snapshot)PropertiesService.getScriptProperties().setProperty('LAST_IMPORT_SNAPSHOT',JSON.stringify(p.snapshot))}
function applyImport_(p){if(!Array.isArray(p.master))throw new Error('master missing');const sh=sheet_(MASTER),last=sh.getLastRow(),old=last>=2?sh.getRange(2,1,last-1,14).getDisplayValues():[],oldMap={};old.forEach(r=>oldMap[key_(r[0],r[1])]=r);const rows=p.master.map(x=>{const k=key_(x['Region'],x['Partner Name (KR)']),o=oldMap[k]||Array(14).fill('');let st=o[11]||'🟡 Check Required';if(x._status==='New')st='🔵 New';if(x._status==='Discontinued'||x['Partnership End Date'])st='🔴 Discontinued';return [x['Region']||'',x['Partner Name (KR)']||'',x['Partner Name (EN)']||'',x['Category']||'',x['Contact Person']||'',x['Phone']||'',x['Email']||'',x['Messenger']||'',x['TripCoupon Benefit']||'',x['Registration Date']||'',x['Partnership End Date']||'',st,o[12]||'',o[13]||'']});if(last>1)sh.getRange(2,1,last-1,14).clearContent();if(rows.length)sh.getRange(2,1,rows.length,14).setValues(rows);const hs=sheet_(HISTORY,['Date','Status','Region','Partner','Details','File']);if(Array.isArray(p.events)&&p.events.length){const vals=p.events.map(x=>[x.date||new Date(),x.status||'',x.region||'',x.partner||'',x.details||'',x.file||p.file||'']);hs.getRange(hs.getLastRow()+1,1,vals.length,6).setValues(vals)}if(p.snapshot)PropertiesService.getScriptProperties().setProperty('LAST_IMPORT_SNAPSHOT',JSON.stringify(p.snapshot))}
