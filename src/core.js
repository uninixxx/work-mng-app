/* Shared payroll calculations for browser, Apps Script and Node. */
var SalaryCore = (function () {
'use strict';
var TYPES=['drink','apo','companion','catch','keep','original','champagne','custom'],MONEY_LIMIT=1000000000,QUANTITY_LIMIT=100000;
function fail(message,field){var e=new Error(message);e.name='ValidationError';e.field=field;throw e;}
function clone(v){return JSON.parse(JSON.stringify(v));}
function has(v){return v!==undefined&&v!==null&&v!=='';}
function number(v,field,o){
 o=o||{};if(!has(v)){if(o.optional)return o.fallback===undefined?0:o.fallback;fail('数値を入力してください。',field);}
 if((typeof v!=='number'&&typeof v!=='string')||(typeof v==='string'&&!/^\s*\d+(?:\.\d+)?\s*$/.test(v)))fail('0以上の数値を入力してください。',field);
 var n=Number(v),max=o.max===undefined?MONEY_LIMIT:o.max;
 if(!isFinite(n)||n<(o.min===undefined?0:o.min)||n>max)fail('入力した数値が範囲外です。数値を確認してください。',field);
 if(o.integer!==false&&Math.floor(n)!==n)fail('整数で入力してください。',field);return n;
}
function dateOnly(v,field){
 if(typeof v!=='string'||!/^\d{4}-\d{2}-\d{2}$/.test(v))fail('日付を選択してください。',field||'businessDate');
 var p=v.split('-').map(Number),d=new Date(Date.UTC(p[0],p[1]-1,p[2]));
 if(p[0]<1900||p[0]>9999||d.getUTCFullYear()!==p[0]||d.getUTCMonth()!==p[1]-1||d.getUTCDate()!==p[2])fail('存在する日付を選択してください。',field||'businessDate');return v;
}
function businessDate(v,cutoff){
 cutoff=number(cutoff===undefined?6:cutoff,'businessDayCutoffHour',{max:23});
 if(typeof v==='string'&&/^\d{4}-\d{2}-\d{2}$/.test(v))return dateOnly(v);
 var instant;
 if(v instanceof Date)instant=v.getTime();
 else if(typeof v==='string'){
  if(/^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?$/.test(v))v=v.replace(' ','T')+'+09:00';
  if(!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:Z|[+-]\d{2}:?\d{2})$/.test(v))fail('日時を確認してください。','businessDate');
  dateOnly(v.substring(0,10));
  if(Number(v.substring(11,13))>23||Number(v.substring(14,16))>59||(v.charAt(16)===':'&&Number(v.substring(17,19))>59))fail('日時を確認してください。','businessDate');
  instant=new Date(v).getTime();
 }
 if(!isFinite(instant))fail('日時を確認してください。','businessDate');
 return new Date(instant+(9-cutoff)*3600000).toISOString().substring(0,10);
}
function defaultSettings(from){
 from=from||'2026-01-01';dateOnly(from,'effectiveFrom');
 return {id:'rules-'+from,effectiveFrom:from,hourlyRate:1700,workUnitMinutes:30,workRounding:'floor',drinkRate:200,apoRate:1000,companionHourlyRate:2000,companionUnitMinutes:60,companionRounding:'ceil',catchRate:200,keepPercent:20,champagnePercent:20,originalTiers:[{min:20000,ratePercent:30},{min:30000,ratePercent:40},{min:40000,ratePercent:50}],businessDayCutoffHour:6,moneyRounding:'floor'};
}
function rounding(v,field){if(['floor','ceil','round'].indexOf(v)<0)fail('端数の計算方法を選択してください。',field);return v;}
function validateSettings(s){
 if(!s||typeof s!=='object'||Array.isArray(s))fail('設定を読み込めません。設定を確認してください。','settings');
 var r=clone(s);r.effectiveFrom=dateOnly(s.effectiveFrom,'effectiveFrom');
 ['hourlyRate','drinkRate','apoRate','companionHourlyRate','catchRate'].forEach(function(k){r[k]=number(s[k],k);});
 ['workUnitMinutes','companionUnitMinutes'].forEach(function(k){r[k]=number(s[k],k,{min:1,max:1440});});
 ['keepPercent','champagnePercent'].forEach(function(k){r[k]=number(s[k],k,{max:100,integer:false});});
 r.businessDayCutoffHour=number(s.businessDayCutoffHour,'businessDayCutoffHour',{max:23});
 r.workRounding=rounding(s.workRounding,'workRounding');r.companionRounding=rounding(s.companionRounding,'companionRounding');r.moneyRounding=rounding(s.moneyRounding,'moneyRounding');
 if(!Array.isArray(s.originalTiers)||s.originalTiers.length>100)fail('オリシャンの価格と歩合率を確認してください。','originalTiers');
 var used={};r.originalTiers=s.originalTiers.map(function(t,i){
  if(!t||typeof t!=='object')fail('価格と歩合率を入力してください。','originalTiers.'+i);
  var min=number(t.min,'originalTiers.'+i+'.min');if(used[min])fail('同じ販売価格のルールが重複しています。','originalTiers.'+i+'.min');used[min]=true;
  return {min:min,ratePercent:number(t.ratePercent,'originalTiers.'+i+'.ratePercent',{max:100,integer:false})};
 }).sort(function(a,b){return a.min-b.min;});return r;
}
function pickSettings(versions,date){
 dateOnly(date);if(!Array.isArray(versions))fail('設定履歴を読み込めません。','settings');
 var selected=null;versions.forEach(function(v){var c=validateSettings(v);if(c.effectiveFrom<=date&&(!selected||c.effectiveFrom>=selected.effectiveFrom))selected=c;});return selected;
}
function roundMoney(v,s,field){return number(Math[s.moneyRounding](v),field||'amount');}
function calculateBackWithRules(b,s){
 if(!b||typeof b!=='object'||Array.isArray(b))fail('バックの入力を確認してください。','backs');
 if(TYPES.indexOf(b.type)<0)fail('バックの種類を選択してください。','type');
 var r=clone(b);r.type=b.type;r.manual=b.manual===true||b.type==='custom';
 r.quantity=has(b.quantity)?number(b.quantity,'quantity',{max:QUANTITY_LIMIT}):r.manual?null:b.type==='companion'?0:1;
 if(has(b.minutes))r.minutes=number(b.minutes,'minutes',{max:1440});
 if(has(b.salePrice))r.salePrice=number(b.salePrice,'salePrice');
 if(has(b.ratePercent))r.ratePercent=number(b.ratePercent,'ratePercent',{max:100,integer:false});
 if(r.manual){r.amount=number(b.amount,'amount');return r;}
 if(b.type==='companion'){
  r.minutes=number(b.minutes,'minutes',{optional:true,max:1440});r.paidMinutes=Math[s.companionRounding](r.minutes/s.companionUnitMinutes)*s.companionUnitMinutes;
  r.hourlyRate=s.companionHourlyRate;r.amount=roundMoney(r.paidMinutes*s.companionHourlyRate/60,s);return r;
 }
 var key={drink:'drinkRate',apo:'apoRate',catch:'catchRate'}[b.type];
 if(key){r.unitRate=s[key];r.amount=roundMoney(r.quantity*r.unitRate,s);return r;}
 r.salePrice=number(b.salePrice,'salePrice',{optional:r.quantity===0});
 if(b.type==='original'){
  if(!has(b.ratePercent)){r.ratePercent=null;s.originalTiers.forEach(function(t){if(r.salePrice>=t.min)r.ratePercent=t.ratePercent;});
   if(r.ratePercent===null&&r.quantity>0)fail('この販売額の歩合率が未設定です。設定を追加するか、歩合率または明細の金額を入力してください。','ratePercent');}
 }else if(!has(b.ratePercent))r.ratePercent=b.type==='keep'?s.keepPercent:s.champagnePercent;
 r.amount=roundMoney(r.quantity*r.salePrice*(r.ratePercent||0)/100,s);return r;
}
function calculateBack(b,s){return calculateBackWithRules(b,validateSettings(s||defaultSettings()));}
function clockMinutes(v,field){
 if(typeof v!=='string'||!/^([01]\d|2[0-3]):[0-5]\d$/.test(v))fail('時刻を入力してください。',field);
 var p=v.split(':').map(Number);return p[0]*60+p[1];
}
function workingMinutes(start,end){return (clockMinutes(end,'endTime')-clockMinutes(start,'startTime')+1440)%1440;}
function calculateRecord(record,settings){
 if(!record||typeof record!=='object'||Array.isArray(record))fail('明細の入力を確認してください。','record');
 var s=validateSettings(record.ruleSnapshot||settings||defaultSettings()),r=clone(record);
 if(has(record.businessDate))r.businessDate=dateOnly(record.businessDate);
 var explicit=has(record.explicitWorkMinutes)?record.explicitWorkMinutes:record.explicitworkMinutes;
 if(!has(explicit))explicit=record.workMinutes;
 if(has(explicit))r.workMinutes=number(explicit,'workMinutes',{max:1439});
 else if(has(record.startTime)||has(record.endTime))r.workMinutes=workingMinutes(record.startTime,record.endTime);
 else r.workMinutes=0;
 if(has(record.startTime))clockMinutes(record.startTime,'startTime');if(has(record.endTime))clockMinutes(record.endTime,'endTime');
 r.paidMinutes=Math[s.workRounding](r.workMinutes/s.workUnitMinutes)*s.workUnitMinutes;r.hourlyRate=s.hourlyRate;r.basePayManual=record.basePayManual===true;
 r.basePay=r.basePayManual?number(record.basePay,'basePay'):roundMoney(r.paidMinutes*s.hourlyRate/60,s,'basePay');
 r.transport=number(record.transport,'transport',{optional:true});
 if(record.backs!==undefined&&!Array.isArray(record.backs))fail('バックの入力を確認してください。','backs');
 if((record.backs||[]).length>100)fail('バックは1明細につき100行以内で入力してください。','backs');
 r.backs=(record.backs||[]).map(function(b,i){try{return calculateBackWithRules(b,s);}catch(e){e.field='backs.'+i+'.'+e.field;throw e;}});
 r.backTotal=r.backs.reduce(function(sum,b){return sum+b.amount;},0);number(r.backTotal,'backTotal');r.total=number(r.basePay+r.backTotal+r.transport,'total');r.ruleSnapshot=clone(s);return r;
}
function summarize(records,period){
 if(!Array.isArray(records))fail('明細を読み込めません。','records');period=period||{};
 var year=has(period.year)?number(period.year,'year',{min:1900,max:9999}):null,month=has(period.month)?number(period.month,'month',{min:1,max:12}):null;
 var r={days:0,recordCount:0,workMinutes:0,paidMinutes:0,basePay:0,backTotal:0,transport:0,total:0,byType:{}},days={};
 TYPES.forEach(function(t){r.byType[t]={amount:0,quantity:0,minutes:0,entryCount:0,unknownQuantityCount:0,unknownMinutesCount:0};});
 records.forEach(function(record){
  if(record.deleted||record.deletedAt)return;var date=dateOnly(record.businessDate);
  if((year&&Number(date.substring(0,4))!==year)||(month&&Number(date.substring(5,7))!==month))return;
  days[date]=true;r.recordCount++;['workMinutes','paidMinutes','basePay','transport','total'].forEach(function(k){r[k]+=number(record[k],k,{optional:true});});
  if(record.backs!==undefined&&!Array.isArray(record.backs))fail('バックの入力を確認してください。','backs');
  (record.backs||[]).forEach(function(b){if(TYPES.indexOf(b.type)<0)fail('バックの種類を確認してください。','type');
   var amount=number(b.amount,'amount',{optional:true}),g=r.byType[b.type];r.backTotal+=amount;g.amount+=amount;g.entryCount++;if(b.type==='companion'&&!has(b.minutes))g.unknownMinutesCount++;
   if(!has(b.quantity)&&b.type!=='companion')g.unknownQuantityCount++;g.quantity+=number(b.quantity,'quantity',{optional:true,max:QUANTITY_LIMIT});g.minutes+=number(b.minutes,'minutes',{optional:true,max:1440});});
 });
 r.days=Object.keys(days).length;r.averagePerDay=r.days?r.total/r.days:null;r.averageBackPerDay=r.days?r.backTotal/r.days:null;return r;
}
function parseOcr(text){
 var fields={},warnings=['写真の文字から作成した入力候補です。写真と照らし合わせて確認・修正してください。'];
 if(typeof text!=='string'||!text.trim())return {fields:fields,warnings:['文字を読み取れませんでした。写真を確認して手入力してください。']};
 var normalized=text.replace(/[！-～]/g,function(c){return String.fromCharCode(c.charCodeAt(0)-65248);}).replace(/　/g,' ');
 var lines=normalized.split(/\r?\n/).map(function(l){return l.trim();}).filter(Boolean);
 var dm=normalized.match(/(?:令和|R)\s*(\d{1,2})\s*年?\s*(\d{1,2})\s*月(?:分)?\s*[)）]?\s*(\d{1,2})\s*日?/i),year;
 if(dm)year=Number(dm[1])+2018;else{dm=normalized.match(/(20\d{2})\s*[年\/-]\s*(\d{1,2})\s*[月\/-]\s*(\d{1,2})\s*日?/);if(dm)year=Number(dm[1]);}
 if(dm){try{fields.businessDate=dateOnly(year+'-'+('0'+Number(dm[2])).slice(-2)+'-'+('0'+Number(dm[3])).slice(-2));}catch(ignore){warnings.push('日付を読み取れませんでした。営業日を選択してください。');}}
 function digits(v){if(!/^\d{1,3}(?:,\d{3})+$|^\d+$/.test(v))return null;var n=Number(v.replace(/,/g,''));return isFinite(n)&&n<=MONEY_LIMIT?n:null;}
 function moneyAfter(label,line){var rem=line.replace(label,'').trim(),m=rem.match(/^[:：]?\s*[¥￥]?\s*(\d[\d,]*)\s*円?\s*$/);return m?digits(m[1]):null;}
 var backs=[],candidates={},duplicate={};
 function candidate(k,v){if(v===null)return;if(Object.prototype.hasOwnProperty.call(candidates,k)&&candidates[k]!==v){duplicate[k]=true;return;}candidates[k]=v;}
 lines.forEach(function(line,i){
  var work=line.match(/(?:労働時間|勤務時間)\s*[:：]?\s*(\d{1,2})\s*時(?:間)?(?:\s*(\d{1,2})\s*分)?\s*$/);
  if(work){var min=Number(work[1])*60+Number(work[2]||0);if(Number(work[2]||0)<60&&min<1440)candidate('workMinutes',min);}
  [['basePay',/基\s*本\s*給/],['transport',/交\s*通\s*費/],['statementTotal',/差\s*引\s*支\s*給\s*額/]].forEach(function(pair){
   if(!pair[1].test(line))return;var v=moneyAfter(pair[1],line);
   if(v===null&&!line.replace(pair[1],'').trim()&&lines[i+1]){var next=lines[i+1].match(/^[¥￥]?\s*(\d[\d,]*)\s*円?$/);if(next)v=digits(next[1]);}candidate(pair[0],v);
  });
  var bm=line.match(/^(ドリンク(?:バック)?|ドリ|ド|アポ(?:バック)?|同伴(?:バック)?|キャッチ(?:バック)?|キープ(?:ボトル)?(?:バック)?|オリシャン(?:バック)?|シャンパン(?:バック)?|シャン)\s*(.*)$/);
  if(!bm)return;var label=bm[1],tail=bm[2].replace(/[()（）]/g,' ').trim(),tokens=tail.match(/^[:：]?\s*(\d[\d,]*)\s*(杯|人|本|件|時間|分)?(?:\s+(?:[¥￥]\s*)?(\d[\d,]*)\s*円?)?\s*$/);
  if(!tokens)return;var first=digits(tokens[1]),amount=tokens[3]?digits(tokens[3]):first;if(amount===null||first===null)return;
  var type=/^ド/.test(label)?'drink':/^アポ/.test(label)?'apo':/^同伴/.test(label)?'companion':/^キャッチ/.test(label)?'catch':/^キープ/.test(label)?'keep':/^オリ/.test(label)?'original':'champagne';
  var b={type:type,manual:true,amount:amount};if(tokens[3]&&first<=QUANTITY_LIMIT){if(type==='companion'&&(tokens[2]==='時間'||tokens[2]==='分'))b.minutes=first*(tokens[2]==='時間'?60:1);else b.quantity=first;}
  if(b.minutes>1440||(!tokens[3]&&tokens[2]))return;backs.push(b);
 });
 Object.keys(candidates).forEach(function(k){if(!duplicate[k])fields[k]=candidates[k];else warnings.push('同じ項目で異なる数値が読まれました。写真を確認してください。');});
 if(backs.length)fields.backs=backs;if(!Object.keys(fields).length)warnings.push('確実に読み取れた項目がありません。写真を確認して手入力してください。');
 return {fields:fields,warnings:warnings};
}
return {defaultSettings:defaultSettings,validateSettings:validateSettings,businessDate:businessDate,pickSettings:pickSettings,workingMinutes:workingMinutes,calculateBack:calculateBack,calculateRecord:calculateRecord,summarize:summarize,parseOcr:parseOcr};
})();
if(typeof module!=='undefined'&&module.exports)module.exports=SalaryCore;
