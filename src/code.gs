function getProducts() {
  const ws = SpreadsheetApp.getActive().getSheetByName("Products");
  const data = ws.getDataRange().getValues();
  if (data.length < 2) return {};

  // อ่าน header row — รองรับทุกลำดับคอลัมน์
  let col = {};
  data[0].forEach((cell, idx) => { col[String(cell).toLowerCase().trim()] = idx; });
  let iBarcode = col['barcode'] !== undefined ? col['barcode'] : 0;
  let iName    = col['name']    !== undefined ? col['name']    : 1;
  let iPrice   = col['price']   !== undefined ? col['price']   : 2;

  let obj = {};
  for (let i = 1; i < data.length; i++) {
    if (!data[i][iBarcode]) continue;
    let raw   = data[i][iBarcode];
    let name  = String(data[i][iName]);
    let price = Number(data[i][iPrice]) || 0;
    let item  = { name, price };

    let asStr = String(raw).trim();
    let asNum = String(Number(raw));
    obj[asStr] = item;
    if (asNum !== asStr) obj[asNum] = item;
  }
  return obj;
}

// ===== helper: โหลด payment channels =====
function _getPayChannels() {
  let chs = [];
  try {
    let raw = PropertiesService.getScriptProperties().getProperty('PAYMENT_CHANNELS');
    if (raw) chs = JSON.parse(raw);
  } catch(e) {}
  if (!chs.length) chs = [
    { key:'cash',     label:'เงินสด',              color:'#27ae60' },
    { key:'transfer', label:'เงินโอน',              color:'#2980b9' },
    { key:'welfare',  label:'บัตรสวัสดิการแห่งรัฐ', color:'#8e44ad' }
  ];
  return chs;
}

// ===== helper: col map จาก header =====
function _getProductColMap(headerRow) {
  let map = {};
  headerRow.forEach((cell, idx) => { map[String(cell).toLowerCase().trim()] = idx; });
  return map;
}

// ===== helper: ดึง/อัปเดต Sales_Header sheet ให้มีคอลัมน์ channel แยก =====
function _ensureSalesHeaderCols(activeChannelKeys) {
  const ss = SpreadsheetApp.getActive();
  let ws   = ss.getSheetByName("Sales_Header");
  let payChannels = _getPayChannels();

  let fixedHeaders   = ['billNo','date','total','type','memberId'];
  let channelHeaders = payChannels.map(c => 'ch_' + c.label);
  // เพิ่ม dynamic channel ที่ไม่อยู่ใน settings
  if (activeChannelKeys) {
    activeChannelKeys.forEach(k => {
      let exists = payChannels.find(c => c.key === k);
      if (!exists) channelHeaders.push('ch_' + k);
    });
  }
  let allHeaders = [...fixedHeaders, ...channelHeaders];

  if (!ws) {
    ws = ss.insertSheet("Sales_Header");
    ws.getRange(1,1,1,allHeaders.length).setValues([allHeaders]);
    ws.setFrozenRows(1);
    ws.getRange(1,1,1,allHeaders.length)
      .setFontWeight("bold").setBackground("#2c3e50").setFontColor("white").setHorizontalAlignment("center");
    ws.autoResizeColumns(1, allHeaders.length);
  } else {
    let existing = ws.getRange(1,1,1,ws.getLastColumn()).getValues()[0].map(String);
    let needUpdate = allHeaders.some(h => !existing.includes(h));
    if (needUpdate) {
      let clearCount = Math.max(existing.length, allHeaders.length);
      ws.getRange(1,1,1,clearCount).clearContent().clearFormat();
      ws.getRange(1,1,1,allHeaders.length).setValues([allHeaders]);
      ws.getRange(1,1,1,allHeaders.length)
        .setFontWeight("bold").setBackground("#2c3e50").setFontColor("white").setHorizontalAlignment("center");
      ws.setFrozenRows(1);
      ws.autoResizeColumns(1, allHeaders.length);
    }
  }
  return ws;
}

function saveSale(data) {
  const ss  = SpreadsheetApp.getActive();
  const wsD = ss.getSheetByName("Sales_Detail");
  const wsP = ss.getSheetByName("Products");
  const wsS = ss.getSheetByName("Settings");

  let billNo = wsS.getRange("B2").getValue();
  wsS.getRange("B2").setValue(billNo + 1);

  // parse channels → channelMap { key: amount }
  let channels   = data.channels || [];
  let channelMap = {};
  channels.forEach(ch => {
    if (ch.amount) channelMap[ch.key] = (channelMap[ch.key]||0) + Number(ch.amount);
  });
  // legacy fallback ถ้าไม่มี channels
  if (!channels.length) {
    if (data.type === 'Cash')       channelMap['cash']     = Number(data.cash)||0;
    else if (data.type === 'QR')    channelMap['transfer'] = Number(data.total)||0;
    else if (data.type === 'Mix') {
      if (data.cash)     channelMap['cash']     = Number(data.cash)||0;
      if (data.transfer) channelMap['transfer'] = Number(data.transfer)||0;
      if (data.welfare)  channelMap['welfare']  = Number(data.welfare)||0;
    }
  }

  // สร้าง/อัปเดต header แล้วเขียนแถว
  let activeKeys = Object.keys(channelMap);
  let ws         = _ensureSalesHeaderCols(activeKeys);
  let headerRow  = ws.getRange(1,1,1,ws.getLastColumn()).getValues()[0].map(String);
  let payChannels = _getPayChannels();
  let tz          = Session.getScriptTimeZone();

  let row = new Array(headerRow.length).fill('');
  let idx_billNo   = headerRow.indexOf('billNo');
  let idx_date     = headerRow.indexOf('date');
  let idx_total    = headerRow.indexOf('total');
  let idx_type     = headerRow.indexOf('type');
  let idx_memberId = headerRow.indexOf('memberId');
  if (idx_billNo   >= 0) row[idx_billNo]   = billNo;
  if (idx_date     >= 0) row[idx_date]     = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm:ss");
  if (idx_total    >= 0) row[idx_total]    = Number(data.total)||0;
  if (idx_type     >= 0) row[idx_type]     = String(data.type||'');
  if (idx_memberId >= 0) row[idx_memberId] = String(data.memberId||'');

  // เติมยอดแต่ละ channel
  payChannels.forEach(ch => {
    let ci = headerRow.indexOf('ch_' + ch.label);
    if (ci >= 0) row[ci] = channelMap[ch.key] || 0;
  });
  activeKeys.forEach(k => {
    let ch    = payChannels.find(c => c.key === k);
    let label = ch ? ch.label : k;
    let ci    = headerRow.indexOf('ch_' + label);
    if (ci >= 0 && row[ci] === '') row[ci] = channelMap[k] || 0;
  });
  // เติม 0 แทน '' ในคอลัมน์ตัวเลข
  row = row.map((v,i) => {
    let h = headerRow[i]||'';
    if (h==='billNo'||h==='total'||h.startsWith('ch_')) return v===''?0:v;
    return v;
  });

  ws.appendRow(row);
  let dataRow = ws.getLastRow();
  ws.getRange(dataRow,1,1,headerRow.length).setBackground(dataRow%2===0?'#f8f9fa':'white');

  // ===== Sales_Detail =====
  let prodData = wsP.getDataRange().getValues();
  let pCol     = _getProductColMap(prodData[0]);
  let iBarcode = pCol['barcode'] !== undefined ? pCol['barcode'] : 0;
  let iCost    = pCol['cost']    !== undefined ? pCol['cost']    : 3;
  let iStock   = pCol['stock']   !== undefined ? pCol['stock']   : 4;

  let detail = [];
  data.items.forEach(item => {
    let cost = 0;
    for (let r = 1; r < prodData.length; r++) {
      if (String(prodData[r][iBarcode]) === String(item.barcode)) {
        cost = Number(prodData[r][iCost]) || 0;
        prodData[r][iStock] = (Number(prodData[r][iStock])||0) - item.qty;
        break;
      }
    }
    detail.push([billNo, item.barcode, item.name, item.price, cost,
                 item.qty, item.price*item.qty, (item.price-cost)*item.qty]);
  });
  if (detail.length > 0)
    wsD.getRange(wsD.getLastRow()+1, 1, detail.length, 8).setValues(detail);
  wsP.getRange(2,1,prodData.length-1,prodData[0].length).setValues(prodData.slice(1));
  return billNo;
}

function doGet(e) {
  if (e.parameter.page === "receipt")       return HtmlService.createHtmlOutputFromFile('receipt');
  if (e.parameter.page === "display")       return HtmlService.createHtmlOutputFromFile('display');
  if (e.parameter.page === "local_display") return HtmlService.createHtmlOutputFromFile('local_display');
  return HtmlService.createHtmlOutputFromFile('index');
}

// ===== helper: อ่าน channels จาก header row ของ Sales_Header =====
function _readChannelsFromRow(h, headerRow, payChannels) {
  let channels = [];
  headerRow.forEach((colName, idx) => {
    if (!colName.startsWith('ch_')) return;
    let amt = Number(h[idx]) || 0;
    if (!amt) return;
    let label = colName.replace(/^ch_/, '');
    let ch    = payChannels.find(c => c.label === label);
    channels.push({ key: ch?ch.key:label, label, amount: amt, color: ch?ch.color:'#7f8c8d' });
  });
  return channels;
}

function getBill(billNo) {
  const ss  = SpreadsheetApp.getActive();
  const wsH = ss.getSheetByName("Sales_Header");
  const wsD = ss.getSheetByName("Sales_Detail");
  const tz  = Session.getScriptTimeZone();
  let headers     = wsH.getDataRange().getValues();
  let details     = wsD.getDataRange().getValues();
  let headerRow   = headers[0].map(String);
  let payChannels = _getPayChannels();
  let result      = { header:null, items:[] };

  for (let i = 1; i < headers.length; i++) {
    let h         = headers[i];
    let rowBillNo = h[headerRow.indexOf('billNo')] !== undefined ? h[headerRow.indexOf('billNo')] : h[0];
    if (rowBillNo == billNo) {
      let channels    = _readChannelsFromRow(h, headerRow, payChannels);
      let cashAmt     = channels.find(c=>c.key==='cash')?.amount     || h[headerRow.indexOf('cash')]     || 0;
      let transferAmt = channels.find(c=>c.key==='transfer')?.amount || h[headerRow.indexOf('transfer')] || 0;
      let welfareAmt  = channels.find(c=>c.key==='welfare')?.amount  || h[headerRow.indexOf('welfare')]  || 0;
      let memberId    = h[headerRow.indexOf('memberId')] !== undefined ? String(h[headerRow.indexOf('memberId')]||'') : String(h[6]||'');
      let dateVal     = h[headerRow.indexOf('date')]     !== undefined ? h[headerRow.indexOf('date')]     : h[1];
      let totalVal    = h[headerRow.indexOf('total')]    !== undefined ? h[headerRow.indexOf('total')]    : h[2];
      let typeVal     = h[headerRow.indexOf('type')]     !== undefined ? h[headerRow.indexOf('type')]     : h[3];
      let dateStr = '';
      try { dateStr = Utilities.formatDate(new Date(dateVal), tz, "dd/MM/yyyy HH:mm"); } catch(e) { dateStr = String(dateVal); }
      result.header = {
        billNo: rowBillNo, date: dateStr,
        total: Number(totalVal)||0, type: String(typeVal||''),
        cash: Number(cashAmt)||0, transfer: Number(transferAmt)||0,
        memberId, welfare: Number(welfareAmt)||0, channels
      };
      break;
    }
  }
  for (let i = 1; i < details.length; i++) {
    if (details[i][0] == billNo)
      result.items.push({ barcode:String(details[i][1]), name:String(details[i][2]),
                          price:Number(details[i][3])||0, cost:Number(details[i][4])||0,
                          qty:Number(details[i][5])||0, total:Number(details[i][6])||0 });
  }
  return result;
}

function getHistory() {
  const ss  = SpreadsheetApp.getActive();
  const wsH = ss.getSheetByName("Sales_Header");
  const wsD = ss.getSheetByName("Sales_Detail");
  const tz  = Session.getScriptTimeZone();
  let headers     = wsH.getDataRange().getValues();
  let details     = wsD.getDataRange().getValues();
  let headerRow   = headers[0].map(String);
  let payChannels = _getPayChannels();
  let bills       = [];

  for (let i = headers.length-1; i >= 1 && bills.length < 50; i--) {
    let h         = headers[i];
    let rowBillNo = h[headerRow.indexOf('billNo')] !== undefined ? h[headerRow.indexOf('billNo')] : h[0];
    if (!rowBillNo) continue;
    let channels    = _readChannelsFromRow(h, headerRow, payChannels);
    let cashAmt     = channels.find(c=>c.key==='cash')?.amount     || Number(h[headerRow.indexOf('cash')])||0;
    let transferAmt = channels.find(c=>c.key==='transfer')?.amount || Number(h[headerRow.indexOf('transfer')])||0;
    let welfareAmt  = channels.find(c=>c.key==='welfare')?.amount  || Number(h[headerRow.indexOf('welfare')])||0;
    let memberId    = h[headerRow.indexOf('memberId')] !== undefined ? String(h[headerRow.indexOf('memberId')]||'') : String(h[6]||'');
    let dateVal     = h[headerRow.indexOf('date')]  !== undefined ? h[headerRow.indexOf('date')]  : h[1];
    let totalVal    = h[headerRow.indexOf('total')] !== undefined ? h[headerRow.indexOf('total')] : h[2];
    let typeVal     = h[headerRow.indexOf('type')]  !== undefined ? h[headerRow.indexOf('type')]  : h[3];
    let dateStr = '';
    try { dateStr = Utilities.formatDate(new Date(dateVal), tz, "dd/MM/yyyy HH:mm"); } catch(e) { dateStr = String(dateVal); }
    let items = [];
    for (let j = 1; j < details.length; j++) {
      if (details[j][0] == rowBillNo)
        items.push({ barcode:String(details[j][1]), name:String(details[j][2]),
                     price:Number(details[j][3])||0, qty:Number(details[j][5])||0, total:Number(details[j][6])||0 });
    }
    bills.push({ billNo:rowBillNo, date:dateStr, total:Number(totalVal)||0, type:String(typeVal||''),
                 cash:Number(cashAmt)||0, transfer:Number(transferAmt)||0,
                 memberId, welfare:Number(welfareAmt)||0, channels, itemCount:items.length, items });
  }
  return bills;
}

function getDashboard() {
  const ss  = SpreadsheetApp.getActive();
  const wsH = ss.getSheetByName("Sales_Header");
  const wsD = ss.getSheetByName("Sales_Detail");
  const tz  = Session.getScriptTimeZone();
  let headers     = wsH.getDataRange().getValues();
  let details     = wsD.getDataRange().getValues();
  let headerRow   = headers[0].map(String);
  let payChannels = _getPayChannels();
  let now         = new Date();
  let todayStr    = Utilities.formatDate(now,tz,"yyyy-MM-dd");
  let monthStr    = Utilities.formatDate(now,tz,"yyyy-MM");
  let profitMap={}, productMap={};
  for (let i = 1; i < details.length; i++) {
    let d=details[i]; if(!d[0]) continue;
    let bn=d[0],name=String(d[2]),qty=Number(d[5])||0,total=Number(d[6])||0,profit=Number(d[7])||0;
    profitMap[bn]=(profitMap[bn]||0)+profit;
    if(!productMap[name])productMap[name]={qty:0,total:0};
    productMap[name].qty+=qty;productMap[name].total+=total;
  }
  let allBills=[];
  for (let i = 1; i < headers.length; i++) {
    let h         = headers[i];
    let rowBillNo = h[headerRow.indexOf('billNo')] !== undefined ? h[headerRow.indexOf('billNo')] : h[0];
    if (!rowBillNo) continue;
    let channels    = _readChannelsFromRow(h, headerRow, payChannels);
    let cashAmt     = channels.find(c=>c.key==='cash')?.amount     || Number(h[headerRow.indexOf('cash')])||0;
    let transferAmt = channels.find(c=>c.key==='transfer')?.amount || Number(h[headerRow.indexOf('transfer')])||0;
    let welfareAmt  = channels.find(c=>c.key==='welfare')?.amount  || Number(h[headerRow.indexOf('welfare')])||0;
    let memberId    = h[headerRow.indexOf('memberId')] !== undefined ? String(h[headerRow.indexOf('memberId')]||'') : String(h[6]||'');
    let dateVal     = h[headerRow.indexOf('date')]  !== undefined ? h[headerRow.indexOf('date')]  : h[1];
    let totalVal    = h[headerRow.indexOf('total')] !== undefined ? h[headerRow.indexOf('total')] : h[2];
    let typeVal     = h[headerRow.indexOf('type')]  !== undefined ? h[headerRow.indexOf('type')]  : h[3];
    let date = new Date(dateVal);
    let dStr = Utilities.formatDate(date,tz,"yyyy-MM-dd");
    let mStr = Utilities.formatDate(date,tz,"yyyy-MM");
    allBills.push({ billNo:rowBillNo, date:Utilities.formatDate(date,tz,"dd/MM/yyyy HH:mm"),
                    dateStr:dStr, monthStr:mStr, total:Number(totalVal)||0, type:String(typeVal||''),
                    cash:Number(cashAmt)||0, transfer:Number(transferAmt)||0,
                    memberId, welfare:Number(welfareAmt)||0, channels,
                    profit:Number(profitMap[rowBillNo])||0 });
  }
  allBills.reverse();
  let summary={today:{sales:0,profit:0,count:0},month:{sales:0,profit:0,count:0},all:{sales:0,profit:0,count:0}};
  let dailyMap={};
  allBills.forEach(b=>{
    summary.all.sales+=b.total;summary.all.profit+=b.profit;summary.all.count++;
    if(b.dateStr===todayStr){summary.today.sales+=b.total;summary.today.profit+=b.profit;summary.today.count++;}
    if(b.monthStr===monthStr){
      summary.month.sales+=b.total;summary.month.profit+=b.profit;summary.month.count++;
      if(!dailyMap[b.dateStr])dailyMap[b.dateStr]={sales:0,profit:0,count:0};
      dailyMap[b.dateStr].sales+=b.total;dailyMap[b.dateStr].profit+=b.profit;dailyMap[b.dateStr].count++;
    }
  });
  let daily=Object.entries(dailyMap).sort((a,b)=>b[0].localeCompare(a[0]))
             .map(([date,v])=>({date,sales:v.sales,profit:v.profit,count:v.count}));
  let top=Object.entries(productMap).sort((a,b)=>b[1].qty-a[1].qty).slice(0,10)
           .map(([name,v])=>({name,qty:v.qty,total:v.total}));
  return { summary, bills:allBills.slice(0,200), daily, top };
}

// ✅ รีเซ็ตระบบทั้งหมด
function resetSystem() {
  const ss  = SpreadsheetApp.getActive();
  const wsH = ss.getSheetByName("Sales_Header");
  const wsD = ss.getSheetByName("Sales_Detail");
  const wsS = ss.getSheetByName("Settings");
  let hHead = wsH.getRange(1,1,1,wsH.getLastColumn()).getValues();
  let dHead = wsD.getRange(1,1,1,wsD.getLastColumn()).getValues();
  wsH.clearContents();
  wsD.clearContents();
  wsH.getRange(1,1,1,hHead[0].length).setValues(hHead);
  wsD.getRange(1,1,1,dHead[0].length).setValues(dHead);
  wsS.getRange("B2").setValue(1);
  return true;
}

// ✅ ตั้งค่าใบเสร็จ
function getReceiptSettings() {
  const ws = SpreadsheetApp.getActive().getSheetByName("Settings");
  return {
    shopName:    String(ws.getRange("B5").getValue()||"ร้านค้า"),
    shopAddress: String(ws.getRange("B6").getValue()||""),
    shopTel:     String(ws.getRange("B7").getValue()||""),
    footer:      String(ws.getRange("B8").getValue()||"ขอบคุณที่ใช้บริการ"),
    logoUrl:     String(ws.getRange("B9").getValue()||"")
  };
}

function saveReceiptSettings(cfg) {
  const ws = SpreadsheetApp.getActive().getSheetByName("Settings");
  ws.getRange("B5").setValue(cfg.shopName    ||"");
  ws.getRange("B6").setValue(cfg.shopAddress ||"");
  ws.getRange("B7").setValue(cfg.shopTel     ||"");
  ws.getRange("B8").setValue(cfg.footer      ||"");
  ws.getRange("B9").setValue(cfg.logoUrl     ||"");
  return true;
}

// ===== UPDATE BILL =====
function updateBill(billNo, newItems, newTotal, newType, newCash, newTransfer) {
  const ss  = SpreadsheetApp.getActive();
  const wsH = ss.getSheetByName("Sales_Header");
  const wsD = ss.getSheetByName("Sales_Detail");
  const wsP = ss.getSheetByName("Products");

  // อัปเดต Header — หา column index ของ total/type/cash/transfer
  let hData     = wsH.getDataRange().getValues();
  let headerRow = hData[0].map(String);
  let iTotal    = headerRow.indexOf('total');
  let iType     = headerRow.indexOf('type');
  // หา billNo column
  let iBillNo   = headerRow.indexOf('billNo');
  if (iBillNo < 0) iBillNo = 0;

  for (let i = 1; i < hData.length; i++) {
    if (hData[i][iBillNo] == billNo) {
      // อัปเดต total และ type
      if (iTotal >= 0) wsH.getRange(i+1, iTotal+1).setValue(newTotal);
      if (iType  >= 0) wsH.getRange(i+1, iType+1).setValue(newType);
      // อัปเดต channel ตามชื่อคอลัมน์
      let payChannels = _getPayChannels();
      let chMap = { cash: newCash||0, transfer: newTransfer||0 };
      payChannels.forEach(ch => {
        let ci = headerRow.indexOf('ch_' + ch.label);
        if (ci >= 0 && chMap[ch.key] !== undefined) {
          wsH.getRange(i+1, ci+1).setValue(chMap[ch.key]);
        }
      });
      break;
    }
  }

  // ลบ detail เดิม
  let dData = wsD.getDataRange().getValues();
  let delRows = [];
  for (let i = 1; i < dData.length; i++) {
    if (dData[i][0] == billNo) delRows.push(i+1);
  }
  for (let i = delRows.length-1; i >= 0; i--) wsD.deleteRow(delRows[i]);

  // เขียน detail ใหม่
  let prodData = wsP.getDataRange().getValues();
  let pCol     = _getProductColMap(prodData[0]);
  let iBarcode = pCol['barcode'] !== undefined ? pCol['barcode'] : 0;
  let iCost    = pCol['cost']    !== undefined ? pCol['cost']    : 3;

  let detail = [];
  newItems.forEach(item => {
    let cost = 0;
    for (let r = 1; r < prodData.length; r++) {
      if (String(prodData[r][iBarcode]) === String(item.barcode)) { cost = Number(prodData[r][iCost])||0; break; }
    }
    detail.push([billNo, item.barcode||'', item.name, Number(item.price)||0, cost,
                 Number(item.qty)||0, Number(item.price)*Number(item.qty),
                 (Number(item.price)-cost)*Number(item.qty)]);
  });
  if (detail.length > 0) wsD.getRange(wsD.getLastRow()+1, 1, detail.length, 8).setValues(detail);
  return true;
}

// ===== SEARCH PRODUCTS =====
function searchProducts(query) {
  const ws   = SpreadsheetApp.getActive().getSheetByName("Products");
  const data = ws.getDataRange().getValues();
  if (data.length < 2) return [];
  let col      = _getProductColMap(data[0]);
  let iBarcode = col['barcode'] !== undefined ? col['barcode'] : 0;
  let iName    = col['name']    !== undefined ? col['name']    : 1;
  let iPrice   = col['price']   !== undefined ? col['price']   : 2;
  let q        = String(query).toLowerCase().trim();
  let results  = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][iBarcode]) continue;
    let barcode = String(data[i][iBarcode]);
    let name    = String(data[i][iName]);
    if (barcode.includes(q) || name.toLowerCase().includes(q)) {
      results.push({ barcode, name, price: Number(data[i][iPrice])||0 });
      if (results.length >= 10) break;
    }
  }
  return results;
}

// ===== AUDIT LOG =====
function logHoldAudit(entry) {
  const ss = SpreadsheetApp.getActive();
  let ws = ss.getSheetByName("AuditLog");
  if (!ws) {
    ws = ss.insertSheet("AuditLog");
    ws.getRange(1,1,1,7).setValues([["วันเวลา","ประเภท","บิลพักที่","ยอดรวม","จำนวนรายการ","รายการสินค้า","IP/หมายเหตุ"]]);
    ws.setFrozenRows(1);
  }
  let itemsStr = (entry.items||[]).map(i=>
    i.name + " x" + i.qty + " [" + (i.barcode||'-') + "] = " + (i.price*i.qty) + "฿"
  ).join(" | ");
  ws.appendRow([
    entry.timeLocal || new Date().toLocaleString(),
    entry.type || "DELETE_HOLD",
    entry.billIndex || "-",
    entry.total || 0,
    entry.qty || 0,
    itemsStr,
    ""
  ]);
  return true;
}

// ===== CUSTOMER DISPLAY =====
function setDisplayData(data) {
  PropertiesService.getScriptProperties().setProperty('DISPLAY_DATA', JSON.stringify(data));
  return true;
}
function getDisplayData() {
  let raw = PropertiesService.getScriptProperties().getProperty('DISPLAY_DATA');
  if (!raw) return null;
  try { return JSON.parse(raw); } catch(e) { return null; }
}

// ===== DELETE BILL =====
function deleteBill(billNo) {
  const ss  = SpreadsheetApp.getActive();
  const wsH = ss.getSheetByName("Sales_Header");
  const wsD = ss.getSheetByName("Sales_Detail");
  let hData     = wsH.getDataRange().getValues();
  let headerRow = hData[0].map(String);
  let iBillNo   = headerRow.indexOf('billNo'); if (iBillNo < 0) iBillNo = 0;
  for (let i = hData.length-1; i >= 1; i--) {
    if (hData[i][iBillNo] == billNo) { wsH.deleteRow(i+1); break; }
  }
  let dData = wsD.getDataRange().getValues();
  for (let i = dData.length-1; i >= 1; i--) {
    if (dData[i][0] == billNo) wsD.deleteRow(i+1);
  }
  let ws = ss.getSheetByName("AuditLog");
  if (!ws) {
    ws = ss.insertSheet("AuditLog");
    ws.getRange(1,1,1,7).setValues([["วันเวลา","ประเภท","บิล#","ยอดรวม","จำนวนรายการ","รายการสินค้า","หมายเหตุ"]]);
    ws.setFrozenRows(1);
  }
  ws.appendRow([new Date().toLocaleString(), "DELETE_BILL", billNo, "", "", "", "ลบโดยเจ้าของร้าน"]);
  return true;
}

// ===== SHIFT MANAGEMENT =====
function openShift(startCash) {
  const ss  = SpreadsheetApp.getActive();
  const tz  = Session.getScriptTimeZone();
  const now = new Date();
  let ws = ss.getSheetByName("Shifts");
  if (!ws) {
    ws = ss.insertSheet("Shifts");
    ws.getRange(1,1,1,12).setValues([[
      "กะที่","เปิดเวลา","ปิดเวลา","ระยะเวลา(ชม.)",
      "เงินทอนเริ่มกะ","ยอดขายรวม","ยอดสด","ยอดโอน","ยอดผสม(สด)","ยอดผสม(โอน)",
      "เงินในเก๊ะที่ต้องมี","จำนวนบิล"
    ]]);
    ws.setFrozenRows(1);
    ws.getRange(1,1,1,12).setFontWeight("bold").setBackground("#2c3e50").setFontColor("white");
  }
  const wsS = ss.getSheetByName("Settings");
  let currentBillNo = wsS ? (wsS.getRange("B2").getValue()-1) : 0;
  let shiftNo = ws.getLastRow();
  let shiftData = {
    shiftNo, startCash: Number(startCash)||0,
    openTime:    Utilities.formatDate(now, tz, "dd/MM/yyyy HH:mm:ss"),
    openTimeISO: now.toISOString(),
    startBillNo: currentBillNo
  };
  PropertiesService.getScriptProperties().setProperty('CURRENT_SHIFT', JSON.stringify(shiftData));
  return shiftData;
}

function closeShift() {
  const ss  = SpreadsheetApp.getActive();
  const tz  = Session.getScriptTimeZone();
  const now = new Date();
  let raw = PropertiesService.getScriptProperties().getProperty('CURRENT_SHIFT');
  if (!raw) return { error: 'ไม่มีกะที่เปิดอยู่' };
  let shift = JSON.parse(raw);

  const wsH    = ss.getSheetByName("Sales_Header");
  let headers  = wsH.getDataRange().getValues();
  let hRow     = headers[0].map(String);
  let openTime = new Date(shift.openTimeISO);
  let payChannels = _getPayChannels();

  let channelTotals = {};
  payChannels.forEach(ch => { channelTotals[ch.key] = 0; });
  let totalSales=0, billCount=0;

  for (let i = 1; i < headers.length; i++) {
    let h = headers[i];
    let rowBillNo = h[hRow.indexOf('billNo')] !== undefined ? h[hRow.indexOf('billNo')] : h[0];
    if (!rowBillNo) continue;
    let dateVal = h[hRow.indexOf('date')] !== undefined ? h[hRow.indexOf('date')] : h[1];
    let saleTime = new Date(dateVal);
    if (saleTime < openTime) continue;
    let total = Number(h[hRow.indexOf('total')] !== undefined ? h[hRow.indexOf('total')] : h[2])||0;
    let type  = String(h[hRow.indexOf('type')]  !== undefined ? h[hRow.indexOf('type')]  : h[3]||'');
    totalSales += total; billCount++;

    // อ่าน channels จากคอลัมน์ ch_
    let hasChCols = false;
    hRow.forEach((colName, idx) => {
      if (!colName.startsWith('ch_')) return;
      let amt = Number(h[idx])||0; if (!amt) return;
      hasChCols = true;
      let label = colName.replace(/^ch_/,'');
      let ch = payChannels.find(c=>c.label===label);
      let key = ch ? ch.key : label;
      channelTotals[key] = (channelTotals[key]||0) + amt;
    });

    // fallback legacy ถ้าไม่มีคอลัมน์ ch_
    if (!hasChCols) {
      let cash     = Number(h[hRow.indexOf('cash')]     !== undefined ? h[hRow.indexOf('cash')]     : h[4])||0;
      let transfer = Number(h[hRow.indexOf('transfer')] !== undefined ? h[hRow.indexOf('transfer')] : h[5])||0;
      let welfare  = Number(h[hRow.indexOf('welfare')]  !== undefined ? h[hRow.indexOf('welfare')]  : h[7])||0;
      // fallback channels JSON (ระบบเก่า)
      let chRaw = h[8] ? String(h[8]) : '';
      if (chRaw && chRaw.startsWith('[')) {
        try {
          JSON.parse(chRaw).forEach(ch => {
            let amt = Number(ch.amount)||0; if (!amt) return;
            channelTotals[ch.key] = (channelTotals[ch.key]||0) + amt;
          });
        } catch(e) {
          if (type==='Cash')    channelTotals['cash']=(channelTotals['cash']||0)+total;
          else if(type==='QR')  channelTotals['transfer']=(channelTotals['transfer']||0)+total;
          else if(type==='Mix'){channelTotals['cash']=(channelTotals['cash']||0)+cash;channelTotals['transfer']=(channelTotals['transfer']||0)+transfer;channelTotals['welfare']=(channelTotals['welfare']||0)+welfare;}
        }
      } else {
        if (type==='Cash')    channelTotals['cash']=(channelTotals['cash']||0)+total;
        else if(type==='QR')  channelTotals['transfer']=(channelTotals['transfer']||0)+total;
        else if(type==='Mix'){channelTotals['cash']=(channelTotals['cash']||0)+cash;channelTotals['transfer']=(channelTotals['transfer']||0)+transfer;channelTotals['welfare']=(channelTotals['welfare']||0)+welfare;}
      }
    }
  }

  let totalCashReceived = channelTotals['cash']||0;
  let totalCashInDrawer = shift.startCash + totalCashReceived;
  let totalTransfer     = channelTotals['transfer']||0;
  let cashSales=channelTotals['cash']||0, qrSales=channelTotals['transfer']||0;
  let diffHrs   = ((now-openTime)/3600000).toFixed(2);
  let closeTimeStr = Utilities.formatDate(now, tz, "dd/MM/yyyy HH:mm:ss");

  let channelKeys = [...payChannels.map(c=>c.key)];
  Object.keys(channelTotals).forEach(k=>{ if(!channelKeys.includes(k)) channelKeys.push(k); });
  let activeChannels = channelKeys.filter(k=>(channelTotals[k]||0)>0).map(k=>{
    let ch=payChannels.find(c=>c.key===k);
    return {key:k,label:ch?ch.label:k,color:ch?ch.color:'#7f8c8d',amount:channelTotals[k]};
  });

  let fixedHeaders = ["กะที่","เปิดเวลา","ปิดเวลา","ระยะเวลา(ชม.)","เงินทอนเริ่มกะ","ยอดขายรวม","จำนวนบิล","เงินในเก๊ะที่ต้องมี","ยอดโอน(transfer)"];
  let allHeaders   = [...fixedHeaders, ...activeChannels.map(c=>c.label)];

  let ws = ss.getSheetByName("Shifts");
  if (!ws) {
    ws = ss.insertSheet("Shifts");
    ws.getRange(1,1,1,allHeaders.length).setValues([allHeaders]);
    ws.setFrozenRows(1);
    ws.getRange(1,1,1,allHeaders.length).setFontWeight("bold").setBackground("#2c3e50").setFontColor("white").setHorizontalAlignment("center");
  } else {
    let ex = ws.getRange(1,1,1,ws.getLastColumn()).getValues()[0].map(String);
    let needUpdate = allHeaders.some(h=>!ex.includes(h))||ex.length!==allHeaders.length;
    if (needUpdate) {
      ws.getRange(1,1,1,Math.max(ex.length,allHeaders.length)).clearContent().clearFormat();
      ws.getRange(1,1,1,allHeaders.length).setValues([allHeaders]);
      ws.getRange(1,1,1,allHeaders.length).setFontWeight("bold").setBackground("#2c3e50").setFontColor("white").setHorizontalAlignment("center");
      ws.setFrozenRows(1);
    }
  }
  let fullRow = [shift.shiftNo,shift.openTime,closeTimeStr,Number(diffHrs),shift.startCash,totalSales,billCount,totalCashInDrawer,totalTransfer,...activeChannels.map(c=>c.amount)];
  let dataRow = ws.getLastRow()+1;
  ws.getRange(dataRow,1,1,fullRow.length).setValues([fullRow]);
  ws.getRange(dataRow,1,1,fullRow.length).setBackground(dataRow%2===0?'#f8f9fa':'white');
  ws.getRange(dataRow,4,1,fullRow.length-3).setNumberFormat("#,##0.00");
  for (let c2=1;c2<=allHeaders.length;c2++) ws.autoResizeColumn(c2);

  PropertiesService.getScriptProperties().deleteProperty('CURRENT_SHIFT');
  let channelSummaryOut={};
  activeChannels.forEach(c=>{channelSummaryOut[c.key]=c;});
  return {
    shiftNo:shift.shiftNo,openTime:shift.openTime,closeTime:closeTimeStr,
    durationHrs:diffHrs,startCash:shift.startCash,
    totalSales,cashSales,qrSales,mixCash:0,mixTf:0,
    totalCashInDrawer,totalTransfer,billCount,channelSummary:channelSummaryOut
  };
}

function resetShiftsSheet() {
  const ss = SpreadsheetApp.getActive();
  let ws = ss.getSheetByName("Shifts");
  if (!ws) { Browser.msgBox("ไม่พบ sheet Shifts"); return; }
  let payChannels = _getPayChannels();
  let fixedHeaders = ["กะที่","เปิดเวลา","ปิดเวลา","ระยะเวลา(ชม.)","เงินทอนเริ่มกะ","ยอดขายรวม","จำนวนบิล","เงินในเก๊ะที่ต้องมี","ยอดโอน(transfer)"];
  let allHeaders   = [...fixedHeaders,...payChannels.map(c=>c.label)];
  let lastCol      = Math.max(ws.getLastColumn(),allHeaders.length);
  ws.getRange(1,1,1,lastCol).clearContent().clearFormat();
  ws.getRange(1,1,1,allHeaders.length).setValues([allHeaders]);
  ws.getRange(1,1,1,allHeaders.length).setFontWeight("bold").setBackground("#2c3e50").setFontColor("white").setHorizontalAlignment("center");
  ws.setFrozenRows(1);
  for (let c=1;c<=allHeaders.length;c++) ws.autoResizeColumn(c);
  Browser.msgBox("✅ อัปเดต header Shifts เรียบร้อย ("+allHeaders.length+" คอลัมน์)");
}

function getShiftStatus() {
  let raw = PropertiesService.getScriptProperties().getProperty('CURRENT_SHIFT');
  if (!raw) return null;
  try { return JSON.parse(raw); } catch(e) { return null; }
}

// ===== MEMBER SYSTEM =====
function _getMembersSheet() {
  const ss = SpreadsheetApp.getActive();
  let ws = ss.getSheetByName("Members");
  if (!ws) {
    ws = ss.insertSheet("Members");
    ws.getRange(1,1,1,7).setValues([["รหัสสมาชิก","ชื่อ","เบอร์โทร","แต้ม","วันที่สมัคร","หมายเหตุ","สถานะ"]]);
    ws.setFrozenRows(1);
    ws.getRange(1,1,1,7).setFontWeight("bold").setBackground("#8e44ad").setFontColor("white");
  }
  return ws;
}
function getMembers() {
  const ws=_getMembersSheet(),data=ws.getDataRange().getValues(),members=[];
  for(let i=1;i<data.length;i++){
    if(!data[i][0])continue;
    members.push({id:String(data[i][0]),name:String(data[i][1]),phone:String(data[i][2]||''),
      points:Number(data[i][3])||0,joinDate:data[i][4]?String(data[i][4]):'',
      note:String(data[i][5]||''),active:data[i][6]!==false&&data[i][6]!=='ไม่ใช้งาน'});
  }
  return members;
}
function searchMember(query) {
  let q=String(query).toLowerCase().trim();
  return getMembers().filter(m=>m.id.toLowerCase().includes(q)||m.name.toLowerCase().includes(q)||m.phone.includes(q)).slice(0,8);
}
function addMember(data) {
  const ws=_getMembersSheet(),all=ws.getDataRange().getValues();
  for(let i=1;i<all.length;i++){if(String(all[i][0])===String(data.id))return{error:'รหัสสมาชิกซ้ำ'};}
  const tz=Session.getScriptTimeZone();
  ws.appendRow([String(data.id),String(data.name),String(data.phone||''),0,
    Utilities.formatDate(new Date(),tz,"dd/MM/yyyy"),String(data.note||''),'ใช้งาน']);
  return{success:true};
}
function updateMember(id,updates) {
  const ws=_getMembersSheet(),data=ws.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(String(data[i][0])===String(id)){
      if(updates.name!==undefined)ws.getRange(i+1,2).setValue(updates.name);
      if(updates.phone!==undefined)ws.getRange(i+1,3).setValue(updates.phone);
      if(updates.note!==undefined)ws.getRange(i+1,6).setValue(updates.note);
      if(updates.active!==undefined)ws.getRange(i+1,7).setValue(updates.active?'ใช้งาน':'ไม่ใช้งาน');
      return{success:true};
    }
  }
  return{error:'ไม่พบสมาชิก'};
}
function getMemberHistory(memberId) {
  const ss=SpreadsheetApp.getActive();
  const wsH=ss.getSheetByName("Sales_Header"),wsD=ss.getSheetByName("Sales_Detail");
  const tz=Session.getScriptTimeZone();
  let headers=wsH.getDataRange().getValues(),details=wsD.getDataRange().getValues();
  let headerRow=headers[0].map(String),payChannels=_getPayChannels();
  let iBillNo=headerRow.indexOf('billNo');if(iBillNo<0)iBillNo=0;
  let iMember=headerRow.indexOf('memberId');if(iMember<0)iMember=6;
  let bills=[];
  for(let i=headers.length-1;i>=1;i--){
    let h=headers[i];
    if(!h[iBillNo])continue;
    if(String(h[iMember]||'')!==String(memberId))continue;
    let channels=_readChannelsFromRow(h,headerRow,payChannels);
    let cashAmt=channels.find(c=>c.key==='cash')?.amount||Number(h[headerRow.indexOf('cash')])||0;
    let transferAmt=channels.find(c=>c.key==='transfer')?.amount||Number(h[headerRow.indexOf('transfer')])||0;
    let welfareAmt=channels.find(c=>c.key==='welfare')?.amount||Number(h[headerRow.indexOf('welfare')])||0;
    let dateVal=h[headerRow.indexOf('date')]!==undefined?h[headerRow.indexOf('date')]:h[1];
    let items=[];
    for(let j=1;j<details.length;j++){
      if(details[j][0]==h[iBillNo])items.push({name:String(details[j][2]),qty:Number(details[j][5])||0,total:Number(details[j][6])||0});
    }
    bills.push({billNo:h[iBillNo],date:Utilities.formatDate(new Date(dateVal),tz,"dd/MM/yyyy HH:mm"),
      total:Number(h[headerRow.indexOf('total')])||0,type:String(h[headerRow.indexOf('type')]||''),
      cash:Number(cashAmt)||0,transfer:Number(transferAmt)||0,welfare:Number(welfareAmt)||0,channels,items});
    if(bills.length>=50)break;
  }
  return bills;
}

// ===== PAYMENT CHANNELS SETTINGS =====
function getPaymentChannels() {
  try{let raw=PropertiesService.getScriptProperties().getProperty('PAYMENT_CHANNELS');if(raw)return JSON.parse(raw);}catch(e){}
  return[{key:'cash',label:'💵 เงินสด',color:'#27ae60'},{key:'transfer',label:'📱 เงินโอน',color:'#2980b9'},{key:'welfare',label:'💳 บัตรสวัสดิการแห่งรัฐ',color:'#8e44ad'}];
}
function savePaymentChannels(channels) {
  PropertiesService.getScriptProperties().setProperty('PAYMENT_CHANNELS',JSON.stringify(channels));
  return true;
}
function getWebAppUrl(){
  return ScriptApp.getService().getUrl();
}
function clearDisplayData() {
  try{PropertiesService.getScriptProperties().deleteProperty('DISPLAY_DATA');}catch(e){}
  return true;
}

// ===== resetSalesHeaderSheet() — Migrate ข้อมูลเก่า JSON → คอลัมน์แยก =====
// รัน 1 ครั้งใน Apps Script เพื่อแปลงข้อมูล Sales_Header เดิม
function resetSalesHeaderSheet() {
  const ss  = SpreadsheetApp.getActive();
  const wsH = ss.getSheetByName("Sales_Header");
  if (!wsH) { Browser.msgBox("ไม่พบ sheet Sales_Header"); return; }

  let payChannels  = _getPayChannels();
  let allData      = wsH.getDataRange().getValues();
  let oldHeader    = allData[0].map(String);

  let fixedHeaders   = ['billNo','date','total','type','memberId'];
  let channelHeaders = payChannels.map(c => 'ch_' + c.label);
  let newHeader      = [...fixedHeaders, ...channelHeaders];

  // index เดิม
  let iBillNo   = oldHeader.indexOf('billNo')   >= 0 ? oldHeader.indexOf('billNo')   : 0;
  let iDate     = oldHeader.indexOf('date')     >= 0 ? oldHeader.indexOf('date')     : 1;
  let iTotal    = oldHeader.indexOf('total')    >= 0 ? oldHeader.indexOf('total')    : 2;
  let iType     = oldHeader.indexOf('type')     >= 0 ? oldHeader.indexOf('type')     : 3;
  let iCash     = oldHeader.indexOf('cash')     >= 0 ? oldHeader.indexOf('cash')     : 4;
  let iTransfer = oldHeader.indexOf('transfer') >= 0 ? oldHeader.indexOf('transfer') : 5;
  let iMember   = oldHeader.indexOf('memberId') >= 0 ? oldHeader.indexOf('memberId') : 6;
  let iWelfare  = oldHeader.indexOf('welfare')  >= 0 ? oldHeader.indexOf('welfare')  : 7;

  let newRows = [newHeader];
  for (let i = 1; i < allData.length; i++) {
    let h = allData[i];
    if (!h[iBillNo]) continue;

    let channelMap = {};
    // ลอง parse JSON column เก่า (index 8)
    let jsonRaw = h[8] ? String(h[8]) : '';
    if (jsonRaw && jsonRaw.startsWith('[')) {
      try {
        JSON.parse(jsonRaw).forEach(ch => {
          let amt = Number(ch.amount)||0; if (!amt) return;
          channelMap[ch.key] = (channelMap[ch.key]||0) + amt;
        });
      } catch(e) {}
    }
    // ถ้าไม่มี JSON ใช้ legacy columns
    if (!Object.keys(channelMap).length) {
      let type=String(h[iType]||''),cash=Number(h[iCash])||0,transfer=Number(h[iTransfer])||0,welfare=Number(h[iWelfare])||0,total=Number(h[iTotal])||0;
      if(type==='Cash')    channelMap['cash']    =cash||total;
      else if(type==='QR') channelMap['transfer']=total;
      else if(type==='Mix'){if(cash)channelMap['cash']=cash;if(transfer)channelMap['transfer']=transfer;if(welfare)channelMap['welfare']=welfare;}
    }
    // ถ้ามี ch_ columns อยู่แล้ว อ่านจากนั้น
    oldHeader.forEach((colName,idx)=>{
      if(!colName.startsWith('ch_'))return;
      let amt=Number(h[idx])||0;if(!amt)return;
      let label=colName.replace(/^ch_/,'');
      let ch=payChannels.find(c=>c.label===label);
      let key=ch?ch.key:label;
      channelMap[key]=(channelMap[key]||0)+amt;
    });

    let row = new Array(newHeader.length).fill('');
    row[newHeader.indexOf('billNo')]   = h[iBillNo];
    row[newHeader.indexOf('date')]     = h[iDate];
    row[newHeader.indexOf('total')]    = Number(h[iTotal])||0;
    row[newHeader.indexOf('type')]     = String(h[iType]||'');
    row[newHeader.indexOf('memberId')] = String(h[iMember]||'');
    payChannels.forEach(ch=>{
      let ci=newHeader.indexOf('ch_'+ch.label);
      if(ci>=0) row[ci]=channelMap[ch.key]||0;
    });
    newRows.push(row);
  }

  wsH.clearContents();
  wsH.getRange(1,1,newRows.length,newHeader.length).setValues(newRows);
  let hRange=wsH.getRange(1,1,1,newHeader.length);
  hRange.setFontWeight("bold").setBackground("#2c3e50").setFontColor("white").setHorizontalAlignment("center");
  wsH.setFrozenRows(1);
  wsH.autoResizeColumns(1,newHeader.length);
  if(newRows.length>1){
    wsH.getRange(2,newHeader.indexOf('total')+1,newRows.length-1,1).setNumberFormat('#,##0.00');
    payChannels.forEach(ch=>{let idx=newHeader.indexOf('ch_'+ch.label);if(idx>=0)wsH.getRange(2,idx+1,newRows.length-1,1).setNumberFormat('#,##0.00');});
  }
  Browser.msgBox('✅ แปลง Sales_Header เรียบร้อย\n'+(newRows.length-1)+' แถว | '+newHeader.length+' คอลัมน์\n\nHeader: '+newHeader.join(' | '));
}
