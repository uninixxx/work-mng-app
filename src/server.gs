/* Private Google Apps Script backend. Never create public sharing links. */
var SalaryServer=(function(){
'use strict';
var STORE_KEY='SALARY_STORE_V1',VERSION=1,MAX_PHOTO_BYTES=2*1024*1024,MAX_JSON_CHARS=45000;
var APPLICATION='salary-store-v1',ACCESS_POLICY='drive.file-v1',FILE_FIELDS='id,name,mimeType,size,parents,trashed,appProperties,description';
var SHEETS=['Records','SettingsVersions','Photos'],ACTIONS=['bootstrap','saveRecord','deleteRecord','saveSettings','uploadPhoto','readPhoto','ocrPhoto','getExport'];
function fail(code,message){var e=new Error(message);e.salaryCode=code;throw e;}
function clone(v){try{return JSON.parse(JSON.stringify(v));}catch(e){fail('INVALID_INPUT','入力内容を確認して、もう一度保存してください。');}}
function timestamp(){return new Date().toISOString();}
function object(v,label){if(!v||typeof v!=='object'||Array.isArray(v))fail('INVALID_INPUT',label+'の入力内容を確認してください。');return v;}
function string(v,max,label){if(v===undefined||v===null)return '';if(typeof v!=='string'||v.length>max)fail('INVALID_INPUT',label+'の入力内容を確認してください。');return v;}
function number(v,label,max,integer){if(typeof v!=='number'||!isFinite(v)||v<0||v>max||(integer&&Math.floor(v)!==v))fail('INVALID_INPUT',label+'には正しい数値を入力してください。');return v;}
function validDate(v){if(typeof v!=='string'||!/^\d{4}-\d{2}-\d{2}$/.test(v))return false;var d=new Date(v+'T00:00:00Z');return !isNaN(d.getTime())&&d.toISOString().slice(0,10)===v;}
function requestId(p){if(typeof p.requestId!=='string'||!/^[A-Za-z0-9_.:-]{8,128}$/.test(p.requestId))fail('INVALID_REQUEST','保存の識別情報がありません。画面を開き直してから、もう一度お試しください。');return p.requestId;}
function canonical(v){if(Array.isArray(v))return '['+v.map(canonical).join(',')+']';if(v&&typeof v==='object')return '{'+Object.keys(v).sort().filter(function(k){return v[k]!==undefined;}).map(function(k){return JSON.stringify(k)+':'+canonical(v[k]);}).join(',')+'}';return JSON.stringify(v);}
function hex(bytes){return bytes.map(function(v){return ('0'+((v+256)%256).toString(16)).slice(-2);}).join('');}
function digest(v){return hex(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,v,Utilities.Charset.UTF_8));}
function append(sheet,event){
 event.storeVersion=VERSION;var json=JSON.stringify(event);if(json.length>MAX_JSON_CHARS)fail('RECORD_TOO_LARGE','1件の記録が大きすぎます。メモや明細の行数を減らして保存してください。');
 // Results and request identifiers share one cell, so uncertain retries can be recovered.
 Sheets.Spreadsheets.Values.append({values:[[json]]},sheet.spreadsheetId,"'"+sheet.name+"'!A:A",{valueInputOption:'RAW',insertDataOption:'INSERT_ROWS'});
}
function readEvents(sheet){
 var rows=Sheets.Spreadsheets.Values.get(sheet.spreadsheetId,"'"+sheet.name+"'!A2:A").values||[];
 return rows.map(function(row,i){
  if(!row[0])return null;try{var event=JSON.parse(String(row[0]));if(!event||typeof event!=='object'||event.storeVersion!==VERSION)throw new Error('Invalid event');return event;}
  catch(e){fail('STORE_DAMAGED','保存データを読み込めません。シート「'+sheet.name+'」の'+(i+2)+'行目を確認してください。');}
 }).filter(function(v){return v!==null;});
}
function fileId(id){if(typeof id!=='string'||!/^[A-Za-z0-9_-]{10,200}$/.test(id))fail('INVALID_FILE','保存ファイルを確認できません。アプリの保存先設定を確認してください。');return id;}
function tags(purpose,extra){var value={application:APPLICATION,purpose:purpose};Object.keys(extra||{}).forEach(function(k){value[k]=String(extra[k]);});return value;}
function tagged(file,purpose){var p=file.appProperties||{};return p.application===APPLICATION&&p.purpose===purpose;}
function metadata(id){return Drive.Files.get(fileId(id),{fields:FILE_FIELDS});}
function createFolder(name,parent,purpose){var resource={name:name,mimeType:'application/vnd.google-apps.folder',appProperties:tags(purpose)};if(parent)resource.parents=[fileId(parent)];return Drive.Files.create(resource).id;}
function storeFile(id,purpose,parent,mimeType){
 var file;try{file=metadata(id);}catch(e){fail('STORE_MISSING','アプリの保存先にアクセスできません。保存ファイルが残っているか確認してください。');}
 if(file.trashed)fail('STORE_MISSING','アプリの保存先がゴミ箱にあります。Googleドライブで元に戻してください。');
 if(!tagged(file,purpose)||file.mimeType!==mimeType||(parent&&(file.parents||[]).indexOf(parent)<0))fail('STORE_FORBIDDEN','アプリ専用の保存先を確認できません。保存先設定を確認してください。');return file;
}
function setup(){
 var props=PropertiesService.getScriptProperties(),saved=props.getProperty(STORE_KEY),meta;
 if(saved){try{meta=JSON.parse(saved);}catch(e){fail('STORE_DAMAGED','保存先の設定を読み込めません。初期設定を確認してください。');}}
 else meta={version:VERSION,createdAt:timestamp()};
 if(meta.version!==VERSION)fail('STORE_VERSION','保存先の形式に対応していません。アプリの更新を確認してください。');
 function remember(){props.setProperty(STORE_KEY,JSON.stringify(meta));}
 if(meta.rootFolderId&&meta.accessPolicy!==ACCESS_POLICY)fail('STORE_MIGRATION_REQUIRED','権限変更前の保存先が設定されています。既存のデータを保持したまま、保存先の移行を確認してください。');
 if(!meta.rootFolderId){meta.rootFolderId=createFolder('給与記録アプリ',null,'root');meta.accessPolicy=ACCESS_POLICY;remember();}
 storeFile(meta.rootFolderId,'root',null,'application/vnd.google-apps.folder');
 if(!meta.photoFolderId){meta.photoFolderId=createFolder('給与明細の写真',meta.rootFolderId,'photos');remember();}
 if(!meta.ocrFolderId){meta.ocrFolderId=createFolder('読み取り用一時ファイル',meta.rootFolderId,'ocr');remember();}
 storeFile(meta.photoFolderId,'photos',meta.rootFolderId,'application/vnd.google-apps.folder');storeFile(meta.ocrFolderId,'ocr',meta.rootFolderId,'application/vnd.google-apps.folder');
 if(!meta.spreadsheetId){meta.spreadsheetId=Drive.Files.create({name:'給与記録アプリのデータ',mimeType:'application/vnd.google-apps.spreadsheet',parents:[meta.rootFolderId],appProperties:tags('records')}).id;remember();}
 storeFile(meta.spreadsheetId,'records',meta.rootFolderId,'application/vnd.google-apps.spreadsheet');
 var book=Sheets.Spreadsheets.get(meta.spreadsheetId,{fields:'sheets(properties(sheetId,title))'}),existing={},requests=[{updateSpreadsheetProperties:{properties:{timeZone:'Asia/Tokyo'},fields:'timeZone'}}],sheets={};
 (book.sheets||[]).forEach(function(s){existing[s.properties.title]=s.properties.sheetId;});
 SHEETS.forEach(function(name){if(existing[name]===undefined)requests.push({addSheet:{properties:{title:name,gridProperties:{frozenRowCount:1}}}});sheets[name]={spreadsheetId:meta.spreadsheetId,name:name};});
 Sheets.Spreadsheets.batchUpdate({requests:requests},meta.spreadsheetId);
 SHEETS.forEach(function(name){var range="'"+name+"'!A1",first=Sheets.Spreadsheets.Values.get(meta.spreadsheetId,range).values||[];if(!first.length)Sheets.Spreadsheets.Values.update({values:[['json']]},meta.spreadsheetId,range,{valueInputOption:'RAW'});else if(first[0][0]!=='json')fail('STORE_DAMAGED','保存データの見出しを確認できません。保存先設定を確認してください。');});
 if(!readEvents(sheets.SettingsVersions).length){var initial=SalaryCore.validateSettings(SalaryCore.defaultSettings('2000-01-01'));initial.id=Utilities.getUuid();initial.createdAt=timestamp();append(sheets.SettingsVersions,{kind:'settings',settings:initial,createdAt:initial.createdAt});}
 return {meta:meta,sheets:sheets};
}
function load(store){
 var records=readEvents(store.sheets.Records),settings=readEvents(store.sheets.SettingsVersions),photos=readEvents(store.sheets.Photos),latest=Object.create(null),requests=Object.create(null);
 records.forEach(function(e){if(!e.record||!e.record.id)fail('STORE_DAMAGED','明細の保存データを確認してください。');latest[e.record.id]=e;});
 records.concat(settings,photos).forEach(function(e){if(e.action&&e.requestId)requests[e.action+':'+e.requestId]=e;});
 return {latest:latest,requests:requests,recordEvents:records,photoEvents:photos,settingsVersions:settings.map(function(e){return e.settings;})};
}
function responseFor(e){if(e.action==='saveRecord')return e.record;if(e.action==='saveSettings')return e.settings;if(e.action==='uploadPhoto')return e.photo;if(e.action==='deleteRecord')return {id:e.record.id,revision:e.record.revision,deleted:true};fail('STORE_DAMAGED','保存結果を読み込めません。画面を開き直してください。');}
function replay(state,action,id,hash){var e=state.requests[action+':'+id];if(!e)return null;if(e.requestHash!==hash)fail('REQUEST_REUSED','保存の識別情報が重複しています。画面を開き直して、入力を確認してください。');return {found:true,data:responseFor(e)};}
function activeRecords(state){return Object.keys(state.latest).map(function(id){return state.latest[id];}).filter(function(e){return e.kind!=='delete';}).map(function(e){return e.record;}).sort(function(a,b){return b.businessDate.localeCompare(a.businessDate)||b.createdAt.localeCompare(a.createdAt);});}
function settingsForDate(versions,date){
 var eligible=versions.map(function(s,i){return {settings:s,index:i};}).filter(function(e){return e.settings.effectiveFrom<=date;});
 if(!eligible.length)fail('SETTINGS_NOT_FOUND','この営業日に適用できる計算設定がありません。設定の適用開始日を確認してください。');
 eligible.sort(function(a,b){return b.settings.effectiveFrom.localeCompare(a.settings.effectiveFrom)||b.index-a.index;});return clone(eligible[0].settings);
}
function storageQuota(){try{var q=Drive.About.get({fields:'storageQuota'}).storageQuota||{};return {usedBytes:q.usage===undefined?null:Number(q.usage),limitBytes:q.limit===undefined?null:Number(q.limit)};}catch(e){return {usedBytes:null,limitBytes:null};}}
function photoFile(store,state,id){
 if(typeof id!=='string'||!/^[A-Za-z0-9_-]{10,200}$/.test(id))fail('PHOTO_NOT_FOUND','写真が見つかりません。写真を選び直してください。');
 if(!state.photoEvents.some(function(e){return e.photo&&e.photo.id===id;}))fail('PHOTO_FORBIDDEN','この写真はアプリに登録されていません。写真を登録し直してください。');
 var f;try{f=metadata(id);}catch(e){fail('PHOTO_NOT_FOUND','写真が見つかりません。Googleドライブで写真が残っているか確認してください。');}
 if(f.trashed)fail('PHOTO_NOT_FOUND','写真がゴミ箱にあります。Googleドライブで元に戻してください。');
 if(!tagged(f,'photo')||(f.parents||[]).indexOf(store.meta.photoFolderId)<0)fail('PHOTO_FORBIDDEN','アプリ専用フォルダ内の保存写真を確認できません。写真を登録し直してください。');
 if(['image/jpeg','image/png'].indexOf(f.mimeType)<0)fail('PHOTO_TYPE','JPEGまたはPNGの写真を選び直してください。');
 if(!isFinite(Number(f.size))||Number(f.size)<=0||Number(f.size)>MAX_PHOTO_BYTES)fail('PHOTO_TOO_LARGE','写真のサイズを確認できません。2MB以下の写真を登録し直してください。');return f;
}
function photoMetadata(file){
 var name=file.name;try{var description=JSON.parse(file.description||'{}');if(description.application===APPLICATION&&typeof description.originalName==='string')name=description.originalName;}catch(ignore){}
 return {id:file.id,name:name,mimeType:file.mimeType,size:Number(file.size)};
}
function photoBlob(file){
 // Apps Script's advanced Drive service does not return binary media reliably.
 // Keep the OAuth token on this fixed Google endpoint; never follow redirects.
 var response=UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/'+fileId(file.id)+'?alt=media',{method:'get',headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},followRedirects:false,muteHttpExceptions:true});
 if(response.getResponseCode()!==200)fail('PHOTO_READ_FAILED','写真を読み込めませんでした。時間をおいて再度お試しください。');
 var blob=response.getBlob(),bytes=blob.getBytes(),mime=String(blob.getContentType()||'').split(';')[0].toLowerCase();
 if(mime!==file.mimeType||bytes.length!==Number(file.size)||bytes.length>MAX_PHOTO_BYTES)fail('PHOTO_READ_FAILED','保存写真の内容を確認できません。写真を登録し直してください。');return blob;
}
function prepareRecord(input){
 object(input,'給与明細');if(!validDate(input.businessDate))fail('INVALID_DATE','営業日を正しい日付で入力してください。');
 var r=clone(input);r.memo=string(r.memo,4000,'メモ');r.startTime=string(r.startTime,5,'出勤時刻');r.endTime=string(r.endTime,5,'退勤時刻');
 [r.startTime,r.endTime].forEach(function(v){if(v&&!/^([01]\d|2[0-3]):[0-5]\d$/.test(v))fail('INVALID_TIME','出勤・退勤時刻を正しく入力してください。');});
 if((r.startTime&&!r.endTime)||(!r.startTime&&r.endTime))fail('INVALID_TIME','出勤時刻と退勤時刻を両方入力してください。');
 if(r.workMinutes!==undefined)number(r.workMinutes,'勤務時間',1439,true);if(r.transport===undefined)r.transport=0;number(r.transport,'交通費',10000000,true);
 if(!Array.isArray(r.backs)||r.backs.length>100)fail('INVALID_INPUT','バックの明細は100行以内で入力してください。');
 var types=['drink','apo','companion','catch','keep','original','champagne','custom'];
 r.backs.forEach(function(b){
  object(b,'バック');if(types.indexOf(b.type)<0)fail('INVALID_INPUT','バックの種類を選び直してください。');b.label=string(b.label,100,'バックの名称');if(!b.id)b.id=Utilities.getUuid();
  [['amount','バック金額',1000000000,true],['quantity','バックの数量',100000,true],['minutes','同伴時間',1440,true],['salePrice','販売価格',1000000000,true],['ratePercent','バック率',100,false]].forEach(function(rule){var v=b[rule[0]];if(v!==undefined&&v!==null&&v!=='')number(v,rule[1],rule[2],rule[3]);});
 });
 if(r.basePayManual)number(r.basePay,'基本給',1000000000,true);
 r.photos=r.photos||[];if(!Array.isArray(r.photos)||r.photos.length>10)fail('INVALID_INPUT','写真は1件につき10枚以内で登録してください。');
 r.source=r.source==='photo'?'photo':'manual';return r;
}
function saveRecord(store,state,p){
 var id=requestId(p);object(p.record,'給与明細');var hash=digest(canonical(p.record)),replayed=replay(state,'saveRecord',id,hash);if(replayed)return replayed.data;
 var candidate=prepareRecord(p.record),previous=candidate.id?state.latest[candidate.id]:null;
 if(candidate.id&&!previous)fail('RECORD_NOT_FOUND','変更する明細が見つかりません。一覧を開き直してください。');
 if(previous&&previous.kind==='delete')fail('RECORD_DELETED','この明細は削除されています。一覧を開き直してください。');
 if(previous&&candidate.revision!==previous.record.revision)fail('REVISION_CONFLICT','別の画面でこの明細が更新されました。一覧を開き直してから変更してください。');
 if(!previous&&candidate.revision!==undefined&&candidate.revision!==0)fail('REVISION_CONFLICT','明細の保存状態を確認できません。画面を開き直してください。');
 var settings=previous&&previous.record.ruleSnapshot?clone(previous.record.ruleSnapshot):candidate.ruleSnapshot?clone(candidate.ruleSnapshot):settingsForDate(state.settingsVersions,candidate.businessDate);
 settings=SalaryCore.validateSettings(settings);candidate.ruleSnapshot=settings;
 var seen=Object.create(null);candidate.photos=candidate.photos.map(function(photo){object(photo,'写真');if(seen[photo.id])fail('PHOTO_DUPLICATE','同じ写真が重複しています。写真を1枚ずつ選び直してください。');seen[photo.id]=true;return photoMetadata(photoFile(store,state,photo.id));});
 var record=SalaryCore.calculateRecord(candidate,settings);record.id=previous?previous.record.id:Utilities.getUuid();record.revision=previous?previous.record.revision+1:1;record.createdAt=previous?previous.record.createdAt:timestamp();record.updatedAt=timestamp();record.ruleSnapshot=settings;delete record.deletedAt;delete record.requestId;
 append(store.sheets.Records,{kind:'record',action:'saveRecord',requestId:id,requestHash:hash,record:record,createdAt:record.updatedAt});return record;
}
function deleteRecord(store,state,p){
 var id=requestId(p),hash=digest(canonical({id:p.id,revision:p.revision})),replayed=replay(state,'deleteRecord',id,hash);if(replayed)return replayed.data;
 var previous=state.latest[p.id];if(!previous||previous.kind==='delete')fail('RECORD_NOT_FOUND','削除する明細が見つかりません。一覧を開き直してください。');
 if(p.revision!==previous.record.revision)fail('REVISION_CONFLICT','別の画面でこの明細が更新されました。内容を確認してから削除してください。');
 var record=clone(previous.record);record.revision++;record.updatedAt=timestamp();record.deletedAt=record.updatedAt;
 append(store.sheets.Records,{kind:'delete',action:'deleteRecord',requestId:id,requestHash:hash,record:record,createdAt:record.updatedAt});return {id:record.id,revision:record.revision,deleted:true};
}
function saveSettings(store,state,p){
 var id=requestId(p);object(p.settings,'設定');var hash=digest(canonical(p.settings)),replayed=replay(state,'saveSettings',id,hash);if(replayed)return replayed.data;
 if(!validDate(p.settings.effectiveFrom))fail('INVALID_DATE','設定の適用開始日を正しく入力してください。');
 var s=SalaryCore.validateSettings(clone(p.settings));s.id=Utilities.getUuid();s.createdAt=timestamp();
 append(store.sheets.SettingsVersions,{kind:'settings',action:'saveSettings',requestId:id,requestHash:hash,settings:s,createdAt:s.createdAt});return s;
}
function uploadPhoto(store,state,p){
 var id=requestId(p),name=string(p.name||'給与明細.jpg',200,'写真の名前'),mime=p.mimeType;
 if(['image/jpeg','image/png'].indexOf(mime)<0)fail('PHOTO_TYPE','JPEGまたはPNGの写真を選んでください。');
 var encoded=p.dataBase64;if(typeof encoded!=='string'||!encoded||encoded.length>Math.ceil(MAX_PHOTO_BYTES/3)*4+4||!/^[A-Za-z0-9+/]+={0,2}$/.test(encoded))fail('PHOTO_INVALID','写真を読み込めません。2MB以下のJPEGまたはPNGを選び直してください。');
 var bytes;try{bytes=Utilities.base64Decode(encoded);}catch(e){fail('PHOTO_INVALID','写真を読み込めません。写真を選び直してください。');}
 if(!bytes.length||bytes.length>MAX_PHOTO_BYTES)fail('PHOTO_TOO_LARGE','写真が2MBを超えています。写真を小さくして登録してください。');
 var sig=bytes.slice(0,8).map(function(v){return (v+256)%256;}),jpeg=bytes.length>=4&&sig[0]===255&&sig[1]===216&&sig[2]===255,png=bytes.length>=24&&sig.join(',')==='137,80,78,71,13,10,26,10';
 if((mime==='image/jpeg'&&!jpeg)||(mime==='image/png'&&!png))fail('PHOTO_TYPE','写真の形式を確認できません。JPEGまたはPNGに変換して選び直してください。');
 var contentHash=hex(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,bytes)),hash=digest(canonical({name:name,mimeType:mime,contentHash:contentHash,size:bytes.length})),replayed=replay(state,'uploadPhoto',id,hash);if(replayed)return replayed.data;
 var storedName='photo_'+digest('uploadPhoto:'+id)+(mime==='image/jpeg'?'.jpg':'.png'),file=null;
 var matches=Drive.Files.list({q:"'"+fileId(store.meta.photoFolderId)+"' in parents and name = '"+storedName+"' and trashed = false",fields:'files('+FILE_FIELDS+')',pageSize:100}).files||[];
 matches.forEach(function(existing){if(!tagged(existing,'photo')||existing.appProperties.requestHash!==hash)fail('REQUEST_REUSED','写真の保存情報が重複しています。画面を開き直してください。');file=existing;});
 if(!file)file=Drive.Files.create({name:storedName,mimeType:mime,parents:[store.meta.photoFolderId],appProperties:tags('photo',{requestHash:hash}),description:JSON.stringify({application:APPLICATION,originalName:name,uploadedAt:timestamp()})},Utilities.newBlob(bytes,mime,storedName),{fields:FILE_FIELDS});
 var photo=photoMetadata(file);append(store.sheets.Photos,{kind:'photo',action:'uploadPhoto',requestId:id,requestHash:hash,photo:photo,createdAt:timestamp()});return photo;
}
function readPhoto(store,state,p){var file=photoFile(store,state,p.id),photo=photoMetadata(file);return {mimeType:photo.mimeType,dataBase64:Utilities.base64Encode(photoBlob(file).getBytes()),name:photo.name};}
function documentText(doc){
 var parts=[];
 function contents(items){(items||[]).forEach(function(item){if(item.paragraph)(item.paragraph.elements||[]).forEach(function(e){if(e.textRun)parts.push(e.textRun.content||'');});if(item.table)(item.table.tableRows||[]).forEach(function(row){(row.tableCells||[]).forEach(function(cell){contents(cell.content);parts.push('\n');});});if(item.tableOfContents)contents(item.tableOfContents.content);});}
 function tabs(items){(items||[]).forEach(function(tab){if(tab.documentTab&&tab.documentTab.body)contents(tab.documentTab.body.content);tabs(tab.childTabs);});}
 if(doc.tabs&&doc.tabs.length)tabs(doc.tabs);else if(doc.body)contents(doc.body.content);return parts.join('').slice(0,100000);
}
function ocrPhoto(store,state,p){
 var file=photoFile(store,state,p.photoId),temporaryId=null,result,pendingError=null;
 try{
  var converted=Drive.Files.create({name:'OCR_'+Utilities.getUuid(),mimeType:'application/vnd.google-apps.document',parents:[store.meta.ocrFolderId],appProperties:tags('ocr-document')},photoBlob(file),{ocrLanguage:'ja',fields:'id'});temporaryId=converted.id;
  var text='';for(var attempt=0;attempt<3;attempt++){try{text=documentText(Docs.Documents.get(temporaryId,{includeTabsContent:true}));if(text.trim())break;}catch(e){if(attempt===2)throw e;}if(attempt<2)Utilities.sleep(500*(attempt+1));}
  if(!text.trim())fail('OCR_EMPTY','文字を読み取れませんでした。明るい場所で撮り直すか、手入力してください。');result={text:text.slice(0,100000)};
 }catch(e){pendingError=e.salaryCode?e:{salaryCode:'OCR_FAILED',message:'写真の読み取りに失敗しました。時間をおいて試すか、手入力してください。'};}
 finally{if(temporaryId){try{storeFile(temporaryId,'ocr-document',store.meta.ocrFolderId,'application/vnd.google-apps.document');Drive.Files.update({trashed:true},temporaryId);}catch(e){pendingError={salaryCode:'OCR_CLEANUP_FAILED',message:'読み取り用の一時ファイルをゴミ箱へ移せませんでした。時間をおいて試してください。元の写真は保存されています。'};}}}
 if(pendingError)throw pendingError;return result;
}
function exportData(state){return {format:'salary-store',version:VERSION,exportedAt:timestamp(),records:activeRecords(state),settingsVersions:state.settingsVersions,deletedRecords:Object.keys(state.latest).filter(function(id){return state.latest[id].kind==='delete';}).map(function(id){return state.latest[id].record;}),photos:state.photoEvents.map(function(e){return e.photo;})};}
function dispatch(action,payload){
 if(ACTIONS.indexOf(action)<0)fail('UNKNOWN_ACTION','この操作に対応していません。画面を開き直してください。');object(payload,'操作');
 var lock=LockService.getScriptLock();if(!lock.tryLock(25000))fail('BUSY','別の保存処理が続いています。少し待ってから、もう一度お試しください。');
 try{var store=setup(),state=load(store);if(action==='bootstrap')return {records:activeRecords(state),settingsVersions:state.settingsVersions,storage:storageQuota(),backend:'google'};
 if(action==='saveRecord')return saveRecord(store,state,payload);if(action==='deleteRecord')return deleteRecord(store,state,payload);if(action==='saveSettings')return saveSettings(store,state,payload);if(action==='uploadPhoto')return uploadPhoto(store,state,payload);if(action==='readPhoto')return readPhoto(store,state,payload);if(action==='ocrPhoto')return ocrPhoto(store,state,payload);return exportData(state);
 }finally{lock.releaseLock();}
}
return {dispatch:dispatch};
})();
function doGet(){return HtmlService.createHtmlOutputFromFile('Index').setTitle('給与明細').addMetaTag('viewport','width=device-width, initial-scale=1, viewport-fit=cover');}
function rpc(action,payload){
 try{return {ok:true,data:SalaryServer.dispatch(action,payload||{})};}
 catch(e){
  if(e&&e.salaryCode)return {ok:false,error:{code:e.salaryCode,message:e.message}};
  if(e&&e.name==='ValidationError')return {ok:false,error:{code:'VALIDATION_ERROR',message:e.message||'入力内容を確認して保存してください。',field:e.field}};
  var message=String(e&&e.message||'');
  if(/quota|limit exceeded|too many times|storage/i.test(message))return {ok:false,error:{code:'SERVICE_LIMIT',message:'Googleの利用上限または保存容量に達しました。空き容量を確認するか、時間をおいて再度お試しください。'}};
  if(/permission|authorization|not authorized|access denied/i.test(message))return {ok:false,error:{code:'AUTH_REQUIRED',message:'Googleへの接続を確認できません。利用するGoogleアカウントで開き直し、必要なアクセスを許可してください。'}};
  if(/(?:Drive|Sheets|Docs) is not defined|advanced service|API has not been used|API.*disabled/i.test(message))return {ok:false,error:{code:'SETUP_REQUIRED',message:'Googleへの接続設定が完了していません。アプリの初期設定を確認してください。'}};
  return {ok:false,error:{code:'SERVER_ERROR',message:'処理を完了できませんでした。入力内容を残したまま、少し待ってからもう一度お試しください。'}};
 }
}
