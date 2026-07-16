// ============================================================
//  茂原みのり歯科クリニック - 算定管理GAS v13（一宮喜楽園専用）
//  修正: 患者名でrow検索（行挿入ずれ対策）
//  修正: formatDateCell - instanceof Date失敗時のフォールバック追加
//  書き込み形式: 「佐藤Dr 口腔ケア 10:00〜10:30 スポンジブラシ・歯ブラシ」
// ============================================================

const THIS_FACILITY = 'ichimiya';

const FACILITY_CONFIG = {
  izumi:    { name:'いすみ苑',    patientCol:3, dataStartCol:4, limitKea:2, limitRiha:4, headerRows:2 },
  nagaiki:  { name:'長生き邑',    patientCol:2, dataStartCol:3, limitKea:4, limitRiha:4, headerRows:2 },
  ichimiya: { name:'一宮喜楽園',  patientCol:2, dataStartCol:3, limitKea:2, limitRiha:4, headerRows:2 },
};

const MEASURE_SHEET_NAME = '測定記録';
const MEASURE_COLS = { facility:0, id:1, name:2, tongueDate:3, tongueVal:4, dryDate:5, dryVal:6 };

function getSheetName(date) { const d=date||new Date(); return 'R'+(d.getFullYear()-2018)+'.'+( d.getMonth()+1); }
function toMD(d) { return (d.getMonth()+1)+'/'+d.getDate(); }
function toYMD(d) { return d.getFullYear()+'/'+(d.getMonth()+1)+'/'+d.getDate(); }
function formatDateCell(v) {
  if (!v && v !== 0) return '';
  if (v instanceof Date) return toYMD(v);
  // GASでinstanceof Dateが失敗する場合の対応（V8ランタイムの既知の挙動）
  const d = new Date(v);
  if (!isNaN(d.getTime())) return toYMD(d);
  return String(v||'').trim();
}
function getSheetNameYM(year, month) { return 'R'+(year-2018)+'.'+month; }
function parseMeasureDate(s) {
  if(!s) return null; s=String(s).trim();
  let m=s.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
  if(m) return new Date(parseInt(m[1]), parseInt(m[2])-1, parseInt(m[3]));
  m=s.match(/^(\d{1,2})\/(\d{1,2})$/);
  if(m){
    const now=new Date(); let year=now.getFullYear(); const mm=parseInt(m[1]);
    if(mm > now.getMonth()+2) year-=1; // 大幅未来月なら前年と判断（年情報なし旧データ対応）
    return new Date(year, mm-1, parseInt(m[2]));
  }
  return null;
}
function normDateStr(v) {
  if(v instanceof Date){return (v.getMonth()+1)+'/'+v.getDate();}
  if(!v&&v!==0)return''; const s=String(v).trim();
  const m1=s.match(/^\d{4}[\/-](\d{1,2})[\/-](\d{1,2})$/);
  if(m1)return parseInt(m1[1])+'/'+parseInt(m1[2]);
  const m2=s.match(/^(\d{1,2})\/(\d{1,2})$/);
  if(m2)return parseInt(m2[1])+'/'+parseInt(m2[2]);
  return s;
}
function normDisp(s) {
  if(!s)return''; s=String(s).trim();
  const m=s.match(/^(\d{1,2})\/(\d{1,2})$/);
  if(m)return parseInt(m[1])+'/'+parseInt(m[2]);
  return '';
}
// 患者名の正規化（NFKC変換＋スペース除去）: 全角・半角スペース差異を吸収して照合に使用
function normalizePatientName(value) {
  return String(value || '').normalize('NFKC').replace(/\s+/g, '').trim();
}

// ★ contentセルをパースして構造化データを返す
// 形式1（新）: 「佐藤Dr 口腔ケア 10:00〜10:30 スポンジブラシ・歯ブラシ」
// 形式2（旧）: 「口腔ケア 10:00〜10:30 スポンジブラシ・歯ブラシ」
function parseContent(raw) {
  if(!raw) return null;
  const STAFF_LIST = ['佐藤Dr','南雲DH','吉岡DH','末吉DH','高梨DH','古阪DH','石井DH'];
  let staff = '';
  let rest = raw.trim();

  // 先頭に担当者があれば抽出
  for(const s of STAFF_LIST) {
    if(rest.startsWith(s + ' ') || rest === s) {
      staff = s;
      rest = rest.slice(s.length).trim();
      break;
    }
  }

  // 種別判定
  let type = '';
  if(rest.startsWith('口腔ケア')) { type = '口腔ケア'; rest = rest.slice(4).trim(); }
  else if(rest.startsWith('口腔リハ')) { type = '口腔リハ'; rest = rest.slice(4).trim(); }

  // 時間抽出: 10:00〜10:30
  let timeStr = '';
  const timeMatch = rest.match(/^(\d{1,2}:\d{2}[〜~]\d{1,2}:\d{2})/);
  if(timeMatch) { timeStr = timeMatch[1]; rest = rest.slice(timeStr.length).trim(); }

  // 残りが実施内容（／備考: 以降は分離）
  let memo = '';
  const MEMO_SEP = '／備考:';
  const memoSep = rest.indexOf(MEMO_SEP);
  if (memoSep !== -1) {
    memo = rest.slice(memoSep + MEMO_SEP.length).trim();
    rest = rest.slice(0, memoSep).trim();
  }
  const contents = rest ? rest.split('・').filter(Boolean) : [];

  return { staff, type, timeStr, contents, memo, raw };
}

function getPatientList() {
  const cfg=FACILITY_CONFIG[THIS_FACILITY]; const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheetName=getSheetName(); const sheet=ss.getSheetByName(sheetName);
  if(!sheet)return{error:`シート「${sheetName}」が見つかりません`};
  const lastRow=sheet.getLastRow(); if(lastRow<=cfg.headerRows)return{patients:[]};
  const colValues=sheet.getRange(1,cfg.patientCol,lastRow,1).getValues(); const patients=[];
  for(let r=cfg.headerRows;r<colValues.length;r++){
    const raw=String(colValues[r][0]).trim(); if(!raw)continue;
    if(/^(患者氏名|患者名|氏名|名前|患者|本館|分館|c|C)$/.test(raw))continue;
    const normRaw = raw.normalize('NFKC'); // 全角スペースを半角に統一してから番号抽出
    const match = normRaw.match(/^(\d+)\s*(.+)$/); // \s* でスペースなし形式も対応
    patients.push({id:match?match[1]:String(r+1),name:raw,row:r+1});
  }
  return{patients,sheetName};
}

function countThisMonth() {
  const cfg=FACILITY_CONFIG[THIS_FACILITY]; const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheetName=getSheetName(); const sheet=ss.getSheetByName(sheetName);
  if(!sheet)return{error:`シート「${sheetName}」が見つかりません`,facility:cfg.name};
  const lastRow=sheet.getLastRow(); const lastCol=sheet.getLastColumn();
  if(lastRow<=cfg.headerRows||lastCol<cfg.dataStartCol)return{facility:cfg.name,sheetName,patients:[]};
  const allData=sheet.getRange(1,1,lastRow,lastCol).getValues();
  const headerRow=allData[cfg.headerRows-1]; const dateCols=[];
  for(let c=cfg.dataStartCol-1;c<headerRow.length;c++){const label=normDateStr(headerRow[c]);if(label)dateCols.push({col:c,label});}
  const results=[];
  for(let r=cfg.headerRows;r<allData.length;r++){
    const raw=String(allData[r][cfg.patientCol-1]).trim();
    if(!raw||/^(患者氏名|患者名|氏名|名前|患者|本館|分館|c|C)$/.test(raw))continue;
    let keaCount=0,rihaCount=0;
    for(const{col}of dateCols){const cell=String(allData[r][col]).trim();if(cell.includes('口腔ケア'))keaCount++;if(cell.includes('口腔リハ'))rihaCount++;}
    results.push({name:raw,keaCount,rihaCount,keaLimit:cfg.limitKea,rihaLimit:cfg.limitRiha,keaOver:keaCount>cfg.limitKea,rihaOver:rihaCount>cfg.limitRiha});
  }
  return{facility:cfg.name,sheetName,patients:results};
}


function countByMonth(year, month) {
  const cfg=FACILITY_CONFIG[THIS_FACILITY]; const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheetName=getSheetNameYM(year, month); const sheet=ss.getSheetByName(sheetName);
  if(!sheet)return{error:`シート「${sheetName}」が見つかりません`,facility:cfg.name,sheetName,year,month,patients:[]};
  const lastRow=sheet.getLastRow(); const lastCol=sheet.getLastColumn();
  if(lastRow<=cfg.headerRows||lastCol<cfg.dataStartCol)return{facility:cfg.name,sheetName,year,month,patients:[]};
  const allData=sheet.getRange(1,1,lastRow,lastCol).getValues();
  const headerRow=allData[cfg.headerRows-1]; const dateCols=[];
  for(let c=cfg.dataStartCol-1;c<headerRow.length;c++){const label=normDateStr(headerRow[c]);if(label)dateCols.push({col:c,label});}
  const results=[];
  for(let r=cfg.headerRows;r<allData.length;r++){
    const raw=String(allData[r][cfg.patientCol-1]).trim();
    if(!raw||/^(患者氏名|患者名|氏名|名前|患者|本館|分館|c|C)$/.test(raw))continue;
    let keaCount=0,rihaCount=0;
    for(const{col}of dateCols){const cell=String(allData[r][col]).trim();if(cell.includes('口腔ケア'))keaCount++;if(cell.includes('口腔リハ'))rihaCount++;}
    results.push({name:raw,keaCount,rihaCount,keaLimit:cfg.limitKea,rihaLimit:cfg.limitRiha,keaOver:keaCount>cfg.limitKea,rihaOver:rihaCount>cfg.limitRiha});
  }
  return{facility:cfg.name,sheetName,year,month,patients:results};
}

function getMeasureTargets(year, month) {
  const cfg=FACILITY_CONFIG[THIS_FACILITY];
  const plist=getPatientList(); if(plist.error) return {error:plist.error};
  const patients=plist.patients||[];
  const measures=getMeasures();
  const measureMapById={}; const measureMapByName={};
  measures.forEach(m=>{
    measureMapById[String(m.id).trim()]=m;
    if(m.name) measureMapByName[normalizePatientName(m.name)]=m; // 正規化した名前でキー登録
  });
  const targetDate=new Date(year, month-1, 1);
  function monthsGap(lastStr){
    const d=parseMeasureDate(lastStr); if(!d) return null;
    return (targetDate.getFullYear()-d.getFullYear())*12+(targetDate.getMonth()-d.getMonth());
  }
  const tongueTargets=[]; const dryTargets=[];
  patients.forEach(p=>{
    // 1.ID検索 → 2.患者名一致確認 → 3.不一致なら名前検索 → 4.見つからなければnull
    const normPName = normalizePatientName(p.name);
    const mById = measureMapById[String(p.id).trim()];
    let m;
    if (mById && normalizePatientName(mById.name) === normPName) {
      m = mById; // IDも患者名も一致
    } else {
      m = measureMapByName[normPName] || null; // 名前で再検索、見つからなければnull
    }
    const tGap = m ? monthsGap(m.tongueDate) : null;
    const dGap = m ? monthsGap(m.dryDate)    : null;
    tongueTargets.push({id:p.id, name:p.name, lastDate:(m&&m.tongueDate)||'', gap: tGap===null?'未測定':tGap+'ヶ月経過', gapNum: tGap});
    dryTargets.push({id:p.id, name:p.name, lastDate:(m&&m.dryDate)||'', gap: dGap===null?'未測定':dGap+'ヶ月経過', gapNum: dGap});
  });
  return{facility:cfg.name,year,month,tongueTargets,dryTargets};
}


function getMonthCalendar(year, month) {
  const cfg=FACILITY_CONFIG[THIS_FACILITY]; const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheetName=getSheetNameYM(year, month); const sheet=ss.getSheetByName(sheetName);
  const daysInMonth = new Date(year, month, 0).getDate();
  const dayMap = {};
  for(let d=1; d<=daysInMonth; d++) dayMap[d] = { day:d, hasVisit:false, hasDr:false, count:0, countDr:0, countDH:0, staffList:[] };
  if(!sheet) return { facility:cfg.name, year, month, days:Object.values(dayMap), sheetExists:false };
  const lastRow=sheet.getLastRow(); const lastCol=sheet.getLastColumn();
  if(lastRow>cfg.headerRows && lastCol>=cfg.dataStartCol){
    const allData=sheet.getRange(1,1,lastRow,lastCol).getValues();
    const headerRow=allData[cfg.headerRows-1];
    const STAFF_LIST=['佐藤Dr','南雲DH','吉岡DH','末吉DH','高梨DH','古阪DH','石井DH'];
    for(let c=cfg.dataStartCol-1;c<headerRow.length;c++){
      const label=normDateStr(headerRow[c]); if(!label) continue;
      const dm=label.match(/^(\d{1,2})\/(\d{1,2})$/); if(!dm) continue;
      const mm=parseInt(dm[1]), dd=parseInt(dm[2]);
      if(mm!==month || !dayMap[dd]) continue;
      const dayInfo = dayMap[dd];
      const staffSet = new Set();
      for(let r=cfg.headerRows;r<allData.length;r++){
        const cell=String(allData[r][c]).trim();
        if(!cell) continue;
        const lines=cell.split('\n').map(l=>l.trim()).filter(l=>l);
        for(const line of lines){
          const isRecord = STAFF_LIST.some(s=>line.startsWith(s)) || line.startsWith('口腔ケア')||line.startsWith('口腔リハ');
          if(!isRecord) continue;
          dayInfo.hasVisit = true; dayInfo.count++;
          for(const s of STAFF_LIST){ if(line.startsWith(s)){ staffSet.add(s); if(s.endsWith('Dr')){ dayInfo.hasDr=true; dayInfo.countDr++; } else { dayInfo.countDH++; } break; } }
        }
      }
      dayInfo.staffList = Array.from(staffSet);
    }
  }
  return { facility:cfg.name, year, month, days:Object.values(dayMap), sheetExists:true };
}

function getTodayRecords(dateStr) {
  const cfg=FACILITY_CONFIG[THIS_FACILITY]; const ss=SpreadsheetApp.getActiveSpreadsheet();
  const _d=dateStr?new Date(dateStr):new Date();
  const sheetName=getSheetName(_d); const sheet=ss.getSheetByName(sheetName);
  if(!sheet)return{error:`シート「${sheetName}」が見つかりません`};
  const targetDate=dateStr?normDateStr(dateStr):toMD(new Date());
  const lastRow=sheet.getLastRow(); const lastCol=sheet.getLastColumn();
  const allData=sheet.getRange(1,1,lastRow,lastCol).getValues();
  const headerRow=allData[cfg.headerRows-1]; let targetCol=-1;
  for(let c=cfg.dataStartCol-1;c<headerRow.length;c++){if(normDateStr(headerRow[c])===targetDate){targetCol=c;break;}}
  if(targetCol===-1)return{facility:cfg.name,date:targetDate,records:[],message:'該当日の記録なし'};
  const records=[];
  for(let r=cfg.headerRows;r<allData.length;r++){
    const name=String(allData[r][cfg.patientCol-1]).trim();
    if(!name||/^(患者氏名|患者名|氏名|名前|患者|本館|分館|c|C)$/.test(name))continue;
    const cell=String(allData[r][targetCol]).trim();
    if(!cell) continue;
    // 改行で複数件入っている場合は分割して処理
    const lines = cell.split('\n').map(l => l.trim()).filter(l => l);
    const STAFF_LIST = ['佐藤Dr','南雲DH','吉岡DH','末吉DH','高梨DH','古阪DH','石井DH'];
    for(const line of lines){
      // 先頭が担当者名または口腔ケア/リハで始まる行のみ新レコードとして処理
      const isRecord = STAFF_LIST.some(s => line.startsWith(s + ' ') || line === s)
        || line.startsWith('口腔ケア') || line.startsWith('口腔リハ');
      if(!isRecord){
        if(records.length > 0){
          const last = records[records.length-1];
          last.memo = last.memo ? last.memo + ' ' + line : line;
          last.raw  = last.raw + '\n' + line;
        }
        continue;
      }
      const parsed = parseContent(line);
      records.push({
        name,
        staff:    parsed ? parsed.staff    : '',
        type:     parsed ? parsed.type     : '',
        timeStr:  parsed ? parsed.timeStr  : '',
        contents: parsed ? parsed.contents : [],
        memo:     parsed ? parsed.memo     : '',
        raw:      line
      });
    }
  }
  return{facility:cfg.name,date:targetDate,records};
}

function writeRecord(params) {
  const cfg=FACILITY_CONFIG[THIS_FACILITY]; const ss=SpreadsheetApp.getActiveSpreadsheet();
  const _wd=params.visitDate?new Date(params.visitDate):new Date();
  const sheetName=getSheetName(_wd); const sheet=ss.getSheetByName(sheetName);
  if(!sheet)return{error:`シート「${sheetName}」が見つかりません`};
  let visitMD; try{const d=new Date(params.visitDate);visitMD=toMD(d);}catch(e){visitMD=normDateStr(params.visitDate);}

  // ★ 患者名でrow番号を動的に検索（行挿入によるずれを防ぐ）
  const lastRow=sheet.getLastRow();
  const patColValues=sheet.getRange(1,cfg.patientCol,lastRow,1).getValues();
  let targetRow=-1;
  const searchName=String(params.patientName||'').trim();
  for(let r=cfg.headerRows;r<patColValues.length;r++){
    if(String(patColValues[r][0]).trim()===searchName){targetRow=r+1;break;}
  }
  // 見つからない場合はpatientRowをフォールバックとして使用
  if(targetRow===-1){
    if(params.patientRow) targetRow=Number(params.patientRow);
    else return{error:`患者「${searchName}」が見つかりません`};
  }

  // getDisplayValues() でヘッダー比較
  const currentLastCol=sheet.getLastColumn();
  const headerDisp=sheet.getRange(cfg.headerRows,1,1,currentLastCol).getDisplayValues()[0];
  let dateCol=-1;
  for(let c=cfg.dataStartCol-1;c<headerDisp.length;c++){
    if(normDisp(headerDisp[c])===visitMD){dateCol=c;break;}
  }
  if(dateCol===-1){
    dateCol=currentLastCol;
    sheet.getRange(cfg.headerRows,currentLastCol+1).setValue(visitMD);
  }

  const staffStr  = params.staff ? params.staff + ' ' : '';
  const typeName  = params.type==='kea'?'口腔ケア':params.type==='riha'?'口腔リハ':'口腔ケア';
  const timeStr   = params.timeStr   ? ' '+params.timeStr   : '';
  const contentStr= params.contentStr? ' '+params.contentStr: '';
  // 備考は改行で追記（元のテキストをそのまま保持）
  const memoStr   = params.memo ? ' ／備考:' + params.memo : '';
  const content   = staffStr + typeName + timeStr + contentStr + memoStr;

  // 既存セルに内容があれば改行して追記
  const cell = sheet.getRange(targetRow, dateCol+1);
  const existing = String(cell.getValue()).trim();
  const newValue = existing ? existing + '\n' + content : content;
  cell.setValue(newValue);
  return{ok:true,wrote:content,row:targetRow,col:dateCol+1,date:visitMD,found:dateCol!==-1};
}

function getMeasures() {
  const ss=SpreadsheetApp.getActiveSpreadsheet(); const sheet=ss.getSheetByName(MEASURE_SHEET_NAME);
  if(!sheet)return[]; const data=sheet.getDataRange().getValues(); const cfg=FACILITY_CONFIG[THIS_FACILITY];
  return data.slice(1).filter(row=>String(row[MEASURE_COLS.facility]).trim()===cfg.name)
    .map(row=>({id:String(row[MEASURE_COLS.id]).trim(),name:String(row[MEASURE_COLS.name]).trim(),tongueDate:formatDateCell(row[MEASURE_COLS.tongueDate]),tongueVal:String(row[MEASURE_COLS.tongueVal]).trim(),dryDate:formatDateCell(row[MEASURE_COLS.dryDate]),dryVal:String(row[MEASURE_COLS.dryVal]).trim()}));
}

function writeMeasure(params) {
  const ss=SpreadsheetApp.getActiveSpreadsheet(); let sheet=ss.getSheetByName(MEASURE_SHEET_NAME);
  if(!sheet){sheet=ss.insertSheet(MEASURE_SHEET_NAME);sheet.appendRow(['施設','患者ID','患者名','舌圧日','舌圧値','口腔乾燥日','口腔乾燥値']);}
  const cfg=FACILITY_CONFIG[THIS_FACILITY]; const data=sheet.getDataRange().getValues(); let targetRow=-1;
  for(let r=1;r<data.length;r++){if(String(data[r][MEASURE_COLS.facility]).trim()===cfg.name&&String(data[r][MEASURE_COLS.id]).trim()===String(params.id).trim()){targetRow=r+1;break;}}
  const now=toYMD(new Date()); if(targetRow===-1){sheet.appendRow([cfg.name,params.id,params.name,'','','','']);targetRow=sheet.getLastRow();}
  if(params.type==='tongue'){sheet.getRange(targetRow,MEASURE_COLS.tongueDate+1).setValue(now);sheet.getRange(targetRow,MEASURE_COLS.tongueVal+1).setValue(params.value);}
  else if(params.type==='dry'){sheet.getRange(targetRow,MEASURE_COLS.dryDate+1).setValue(now);sheet.getRange(targetRow,MEASURE_COLS.dryVal+1).setValue(params.value);}
  return{ok:true};
}

const MEMO_SHEET_NAME = '備考記録';

function writeMemoRecord(params) {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sheet=ss.getSheetByName(MEMO_SHEET_NAME);
  if(!sheet){
    sheet=ss.insertSheet(MEMO_SHEET_NAME);
    sheet.appendRow(['施設','患者名','日付','担当者','備考内容']);
  }
  const cfg=FACILITY_CONFIG[THIS_FACILITY];
  const now=params.visitDate?normDateStr(params.visitDate):toMD(new Date());
  sheet.appendRow([cfg.name, params.patientName||'', now, params.staff||'', params.memo||'']);
  return{ok:true};
}

function updateRecord(params) {
  const cfg=FACILITY_CONFIG[THIS_FACILITY]; const ss=SpreadsheetApp.getActiveSpreadsheet();
  const _d=params.visitDate?(/^\d{4}[\/\-]/.test(params.visitDate)?new Date(params.visitDate):new Date(new Date().getFullYear()+'/'+params.visitDate)):new Date();
  const sheetName=getSheetName(_d); const sheet=ss.getSheetByName(sheetName);
  if(!sheet)return{error:`シート「${sheetName}」が見つかりません`};
  let visitMD; try{visitMD=toMD(new Date(params.visitDate));}catch(e){visitMD=normDateStr(params.visitDate);}
  const lastRow=sheet.getLastRow();
  const patColValues=sheet.getRange(1,cfg.patientCol,lastRow,1).getValues();
  let targetRow=-1; const searchName=String(params.patientName||'').trim();
  for(let r=cfg.headerRows;r<patColValues.length;r++){if(String(patColValues[r][0]).trim()===searchName){targetRow=r+1;break;}}
  if(targetRow===-1)return{error:`患者「${searchName}」が見つかりません`};
  const currentLastCol=sheet.getLastColumn();
  const headerDisp=sheet.getRange(cfg.headerRows,1,1,currentLastCol).getDisplayValues()[0];
  let dateCol=-1;
  for(let c=cfg.dataStartCol-1;c<headerDisp.length;c++){if(normDisp(headerDisp[c])===visitMD){dateCol=c;break;}}
  if(dateCol===-1)return{error:'日付列が見つかりません'};
  const cell=sheet.getRange(targetRow,dateCol+1);
  const existing=String(cell.getValue()).trim();
  const oldRaw=String(params.oldRaw||'').trim();
  const staffStr=params.staff?params.staff+' ':'';
  const typeName=params.type==='kea'?'口腔ケア':params.type==='riha'?'口腔リハ':'口腔ケア';
  const timeStr=params.timeStr?' '+params.timeStr:'';
  const contentStr=params.contentStr?' '+params.contentStr:'';
  const memoStr=params.memo?' ／備考:'+params.memo:'';
  const newContent=staffStr+typeName+timeStr+contentStr+memoStr;
  const lines=existing.split('\n').map(l=>l.trim()).filter(l=>l);
  let found=false;
  const newLines=lines.map(l=>{if(l.trim()===oldRaw){found=true;return newContent;}return l;});
  if(!found)return{error:'元の記録が見つかりません',oldRaw,existing};
  cell.setValue(newLines.join('\n'));
  return{ok:true,wrote:newContent};
}

function deleteRecord(params) {
  const cfg=FACILITY_CONFIG[THIS_FACILITY]; const ss=SpreadsheetApp.getActiveSpreadsheet();
  const _d=params.visitDate?(/^\d{4}[\/\-]/.test(params.visitDate)?new Date(params.visitDate):new Date(new Date().getFullYear()+'/'+params.visitDate)):new Date();
  const sheetName=getSheetName(_d); const sheet=ss.getSheetByName(sheetName);
  if(!sheet)return{error:`シート「${sheetName}」が見つかりません`};
  let visitMD; try{visitMD=toMD(new Date(params.visitDate));}catch(e){visitMD=normDateStr(params.visitDate);}
  const lastRow=sheet.getLastRow();
  const patColValues=sheet.getRange(1,cfg.patientCol,lastRow,1).getValues();
  let targetRow=-1; const searchName=String(params.patientName||'').trim();
  for(let r=cfg.headerRows;r<patColValues.length;r++){if(String(patColValues[r][0]).trim()===searchName){targetRow=r+1;break;}}
  if(targetRow===-1)return{error:`患者「${searchName}」が見つかりません`};
  const currentLastCol=sheet.getLastColumn();
  const headerDisp=sheet.getRange(cfg.headerRows,1,1,currentLastCol).getDisplayValues()[0];
  let dateCol=-1;
  for(let c=cfg.dataStartCol-1;c<headerDisp.length;c++){if(normDisp(headerDisp[c])===visitMD){dateCol=c;break;}}
  if(dateCol===-1)return{error:'日付列が見つかりません'};
  const cell=sheet.getRange(targetRow,dateCol+1);
  const existing=String(cell.getValue()).trim();
  const rawToDelete=String(params.rawText||'').trim();
  const lines=existing.split('\n').map(l=>l.trim()).filter(l=>l);
  const newLines=lines.filter(l=>l.trim()!==rawToDelete);
  if(newLines.length===lines.length)return{error:'該当記録が見つかりません',rawToDelete};
  cell.setValue(newLines.join('\n'));
  return{ok:true,deleted:rawToDelete};
}

function doGet(e) {
  const action=e&&e.parameter&&e.parameter.action?e.parameter.action:''; let result;
  try{
    if(action==='patients')result=getPatientList();
    else if(action==='measures')result={measures:getMeasures()};
    else if(action==='today')result=getTodayRecords(e&&e.parameter&&e.parameter.date?e.parameter.date:null);
    else if(action==='write')result=writeRecord(e.parameter);
    else if(action==='updateRecord')result=updateRecord(e.parameter);
    else if(action==='deleteRecord')result=deleteRecord(e.parameter);
    else if(action==='monthCount')result=countByMonth(parseInt(e.parameter.year), parseInt(e.parameter.month));
    else if(action==='measureTargets')result=getMeasureTargets(parseInt(e.parameter.year), parseInt(e.parameter.month));
    else if(action==='calendar')result=getMonthCalendar(parseInt(e.parameter.year), parseInt(e.parameter.month));
    else if(action==='writeMeasure')result=writeMeasure(e.parameter);
    else if(action==='writeMemo')result=writeMemoRecord(e.parameter);
    else result=countThisMonth();
  }catch(err){result={error:err.toString()};}
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}
