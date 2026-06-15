function getProducts() {
  const ws = SpreadsheetApp.getActive().getSheetByName("Products");
  const data = ws.getDataRange().getValues();
  let obj = {};
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    let raw = data[i][0];
    let name  = String(data[i][1]);
    let price = Number(data[i][2]) || 0;
    let item  = { name, price };

    // เก็บทั้งแบบ string ตรงๆ และแบบตัวเลข
    // เพื่อรองรับ barcode ที่นำหน้าด้วย 0 เช่น 052010
    let asStr = String(raw).trim();          // "052010"
    let asNum = String(Number(raw));         // "52010"

    obj[asStr] = item;
    if (asNum !== asStr) obj[asNum] = item;  // เก็บทั้ง 2 key
  }
  return obj;
}

function saveSale(data) {
  const ss  = SpreadsheetApp.getActive();
  const wsH = ss.getSheetByName("Sales_Header");
  const wsD = ss.getSheetByName("Sales_Detail");
  const wsP = ss.getSheetByName("Products");
  const wsS = ss.getSheetByName("Settings");

  let billNo = wsS.getRange("B2").getValue();
  wsS.getRange("B2").setValue(billNo + 1);

  // รองรับทั้งแบบเดิม (cash/transfer) และแบบใหม่ (channels array)
  let cashAmt     = data.cash     || 0;
  let transferAmt = data.transfer || 0;
  let welfareAmt  = data.welfare  || 0;
  let extraJSON   = data.channels ? JSON.stringify(data.channels) : "";

  // ถ้ามี channels ให้คำนวณใหม่
  if (data.channels && data.channels.length) {
    cashAmt = transferAmt = welfareAmt = 0;
    data.channels.forEach(ch => {
      let amt = Number(ch.amount) || 0;
      if (ch.key === 'cash')    cashAmt     += amt;
      else if (ch.key === 'transfer') transferAmt += amt;
      else if (ch.key === 'welfare')  welfareAmt  += amt;
    });
  }

  wsH.appendRow([billNo, new Date(), data.total, data.type,
                 cashAmt, transferAmt, data.memberId||"", welfareAmt, extraJSON]);

  let prodData = wsP.getDataRange().getValues();
  let detail   = [];

  data.items.forEach(item => {
    let cost = 0;
    for (let r = 1; r < prodData.length; r++) {
      if (String(prodData[r][0]) === String(item.barcode)) {
        cost = Number(prodData[r][3])||0;
        prodData[r][4] = (Number(prodData[r][4])||0) - item.qty;
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

function getBill(billNo) {
  const ss  = SpreadsheetApp.getActive();
  const wsH = ss.getSheetByName("Sales_Header");
  const wsD = ss.getSheetByName("Sales_Detail");
  const tz  = Session.getScriptTimeZone();
  let headers = wsH.getDataRange().getValues();
  let details = wsD.getDataRange().getValues();
  let result  = { header:null, items:[] };
  for (let i=1;i<headers.length;i++) {
    if (headers[i][0]==billNo) {
      let h = headers[i];
      let channels = [];
      try { if(h[8]) channels = JSON.parse(h[8]); } catch(e){}
      result.header = { billNo:h[0], date:Utilities.formatDate(new Date(h[1]),tz,"dd/MM/yyyy HH:mm"),
                        total:Number(h[2])||0, type:String(h[3]), cash:Number(h[4])||0,
                        transfer:Number(h[5])||0, memberId:String(h[6]||''),
                        welfare:Number(h[7])||0, channels };
      break;
    }
  }
  for (let i=1;i<details.length;i++) {
    if (details[i][0]==billNo)
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
  let headers = wsH.getDataRange().getValues();
  let details = wsD.getDataRange().getValues();
  let bills   = [];
  for (let i=headers.length-1; i>=1 && bills.length<50; i--) {
    let h = headers[i]; if (!h[0]) continue;
    let items = [];
    for (let j=1;j<details.length;j++) {
      if (details[j][0]==h[0])
        items.push({ barcode:String(details[j][1]), name:String(details[j][2]),
                     price:Number(details[j][3])||0, qty:Number(details[j][5])||0,
                     total:Number(details[j][6])||0 });
    }
    let chs2 = [];
    try { if(h[8]) chs2 = JSON.parse(h[8]); } catch(e){}
    bills.push({ billNo:h[0], date:Utilities.formatDate(new Date(h[1]),tz,"dd/MM/yyyy HH:mm"),
                 total:Number(h[2])||0, type:String(h[3]), cash:Number(h[4])||0,
                 transfer:Number(h[5])||0, memberId:String(h[6]||''),
                 welfare:Number(h[7])||0, channels:chs2,
                 itemCount:items.length, items });
  }
  return bills;
}

function getDashboard() {
  const ss  = SpreadsheetApp.getActive();
  const wsH = ss.getSheetByName("Sales_Header");
  const wsD = ss.getSheetByName("Sales_Detail");
  const tz  = Session.getScriptTimeZone();
  let headers = wsH.getDataRange().getValues();
  let details = wsD.getDataRange().getValues();
  let now = new Date();
  let todayStr = Utilities.formatDate(now,tz,"yyyy-MM-dd");
  let monthStr = Utilities.formatDate(now,tz,"yyyy-MM");
  let profitMap={}, productMap={};
  for (let i=1;i<details.length;i++) {
    let d=details[i]; if(!d[0]) continue;
    let bn=d[0], name=String(d[2]), qty=Number(d[5])||0, total=Number(d[6])||0, profit=Number(d[7])||0;
    profitMap[bn]=(profitMap[bn]||0)+profit;
    if(!productMap[name]) productMap[name]={qty:0,total:0};
    productMap[name].qty+=qty; productMap[name].total+=total;
  }
  let allBills=[];
  for (let i=1;i<headers.length;i++) {
    let h=headers[i]; if(!h[0]) continue;
    let date=new Date(h[1]);
    let dStr=Utilities.formatDate(date,tz,"yyyy-MM-dd");
    let mStr=Utilities.formatDate(date,tz,"yyyy-MM");
    let chs3 = [];
    try { if(h[8]) chs3 = JSON.parse(h[8]); } catch(e){}
    allBills.push({ billNo:h[0], date:Utilities.formatDate(date,tz,"dd/MM/yyyy HH:mm"),
                    dateStr:dStr, monthStr:mStr, total:Number(h[2])||0, type:String(h[3]),
                    cash:Number(h[4])||0, transfer:Number(h[5])||0,
                    memberId:String(h[6]||''), welfare:Number(h[7])||0, channels:chs3,
                    profit:Number(profitMap[h[0]])||0 });
  }
  allBills.reverse();
  let summary={ today:{sales:0,profit:0,count:0}, month:{sales:0,profit:0,count:0}, all:{sales:0,profit:0,count:0} };
  let dailyMap={};
  allBills.forEach(b=>{
    summary.all.sales+=b.total; summary.all.profit+=b.profit; summary.all.count++;
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

  // เก็บ header row ไว้
  let hHead = wsH.getRange(1,1,1,wsH.getLastColumn()).getValues();
  let dHead = wsD.getRange(1,1,1,wsD.getLastColumn()).getValues();

  wsH.clearContents();
  wsD.clearContents();

  wsH.getRange(1,1,1,hHead[0].length).setValues(hHead);
  wsD.getRange(1,1,1,dHead[0].length).setValues(dHead);

  wsS.getRange("B2").setValue(1); // รีเซ็ตบิลเลข → 1
  return true;
}

// ✅ ตั้งค่าใบเสร็จ (เก็บใน Settings sheet)
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

// ===== UPDATE BILL (แก้ไขบิล) =====
function updateBill(billNo, newItems, newTotal, newType, newCash, newTransfer) {
  const ss  = SpreadsheetApp.getActive();
  const wsH = ss.getSheetByName("Sales_Header");
  const wsD = ss.getSheetByName("Sales_Detail");
  const wsP = ss.getSheetByName("Products");

  // อัปเดต Header
  let hData = wsH.getDataRange().getValues();
  for (let i = 1; i < hData.length; i++) {
    if (hData[i][0] == billNo) {
      wsH.getRange(i+1, 3, 1, 4).setValues([[newTotal, newType, newCash||0, newTransfer||0]]);
      break;
    }
  }

  // ลบ detail เดิมทิ้ง แล้วเขียนใหม่
  let dData = wsD.getDataRange().getValues();
  let prodData = wsP.getDataRange().getValues();

  // หา row ที่ต้องลบ (รวบแถวที่เป็นบิลนี้)
  let delRows = [];
  for (let i = 1; i < dData.length; i++) {
    if (dData[i][0] == billNo) delRows.push(i+1);
  }
  // ลบจากล่างขึ้นบนเพื่อไม่ให้ index เลื่อน
  for (let i = delRows.length - 1; i >= 0; i--) {
    wsD.deleteRow(delRows[i]);
  }

  // เขียน detail ใหม่
  let detail = [];
  newItems.forEach(item => {
    let cost = 0;
    for (let r = 1; r < prodData.length; r++) {
      if (String(prodData[r][0]) === String(item.barcode)) { cost = Number(prodData[r][3])||0; break; }
    }
    detail.push([billNo, item.barcode||'', item.name, Number(item.price)||0, cost,
                 Number(item.qty)||0, Number(item.price)*Number(item.qty),
                 (Number(item.price)-cost)*Number(item.qty)]);
  });

  if (detail.length > 0) {
    wsD.getRange(wsD.getLastRow()+1, 1, detail.length, 8).setValues(detail);
  }
  return true;
}

// ===== SEARCH PRODUCTS (ค้นหาสินค้าสำหรับแก้ไขบิล) =====
function searchProducts(query) {
  const ws = SpreadsheetApp.getActive().getSheetByName("Products");
  const data = ws.getDataRange().getValues();
  let q = String(query).toLowerCase().trim();
  let results = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    let barcode = String(data[i][0]);
    let name    = String(data[i][1]);
    if (barcode.includes(q) || name.toLowerCase().includes(q)) {
      results.push({ barcode, name, price: Number(data[i][2])||0 });
      if (results.length >= 10) break;
    }
  }
  return results;
}

// ===== AUDIT LOG (บันทึกการลบบิลพัก) =====
function logHoldAudit(entry) {
  const ss = SpreadsheetApp.getActive();
  let ws = ss.getSheetByName("AuditLog");

  // สร้าง sheet ถ้ายังไม่มี
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

// ===== CUSTOMER DISPLAY (cross-device via Script Properties) =====
function setDisplayData(data) {
  PropertiesService.getScriptProperties()
    .setProperty('DISPLAY_DATA', JSON.stringify(data));
  return true;
}

function getDisplayData() {
  let raw = PropertiesService.getScriptProperties()
    .getProperty('DISPLAY_DATA');
  if (!raw) return null;
  try { return JSON.parse(raw); } catch(e) { return null; }
}

// ===== DELETE BILL =====
function deleteBill(billNo) {
  const ss  = SpreadsheetApp.getActive();
  const wsH = ss.getSheetByName("Sales_Header");
  const wsD = ss.getSheetByName("Sales_Detail");

  // ลบ Header row
  let hData = wsH.getDataRange().getValues();
  for (let i = hData.length - 1; i >= 1; i--) {
    if (hData[i][0] == billNo) { wsH.deleteRow(i + 1); break; }
  }

  // ลบ Detail rows (จากล่างขึ้นบน)
  let dData = wsD.getDataRange().getValues();
  for (let i = dData.length - 1; i >= 1; i--) {
    if (dData[i][0] == billNo) wsD.deleteRow(i + 1);
  }

  // บันทึก AuditLog
  let ws = ss.getSheetByName("AuditLog");
  if (!ws) {
    ws = ss.insertSheet("AuditLog");
    ws.getRange(1,1,1,7).setValues([["วันเวลา","ประเภท","บิล#","ยอดรวม","จำนวนรายการ","รายการสินค้า","หมายเหตุ"]]);
    ws.setFrozenRows(1);
  }
  ws.appendRow([
    new Date().toLocaleString(), "DELETE_BILL", billNo, "", "", "", "ลบโดยเจ้าของร้าน"
  ]);

  return true;
}

// ===== SHIFT MANAGEMENT =====

function openShift(startCash) {
  const ss  = SpreadsheetApp.getActive();
  const tz  = Session.getScriptTimeZone();
  const now = new Date();

  // สร้าง sheet Shifts ถ้าไม่มี
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

  // หา billNo ปัจจุบัน (เพื่อ mark จุดเริ่มต้น)
  const wsS = ss.getSheetByName("Settings");
  let currentBillNo = wsS ? (wsS.getRange("B2").getValue() - 1) : 0;

  let shiftNo = ws.getLastRow(); // นับจำนวนกะ (row 1 = header)

  let shiftData = {
    shiftNo:      shiftNo,
    openTime:     Utilities.formatDate(now, tz, "dd/MM/yyyy HH:mm:ss"),
    openTimeISO:  now.toISOString(),
    startCash:    Number(startCash) || 0,
    startBillNo:  currentBillNo
  };

  // เก็บ shift ที่กำลังเปิดใน Script Properties
  PropertiesService.getScriptProperties()
    .setProperty('CURRENT_SHIFT', JSON.stringify(shiftData));

  return shiftData;
}

// ===== แทนที่ฟังก์ชัน closeShift() เดิมทั้งหมด =====
// แก้ bug: ยอดโอนรวมเกินเพราะนับซ้ำ
// Logic ที่ถูก:
//   - totalCashInDrawer = startCash + ยอด ch.key='cash' ทุกบิล
//   - totalTransfer     = ยอด ch.key='transfer' ทุกบิล (ไม่รวม welfare หรือช่องอื่น)
//   - แต่ละ channel นับอิสระจากกัน ไม่ขึ้นกับ type ของ bill

function closeShift() {
  const ss  = SpreadsheetApp.getActive();
  const tz  = Session.getScriptTimeZone();
  const now = new Date();

  let raw = PropertiesService.getScriptProperties().getProperty('CURRENT_SHIFT');
  if (!raw) return { error: 'ไม่มีกะที่เปิดอยู่' };
  let shift = JSON.parse(raw);

  const wsH    = ss.getSheetByName("Sales_Header");
  let headers  = wsH.getDataRange().getValues();
  let openTime = new Date(shift.openTimeISO);

  // โหลด payment channels
  let payChannels = [];
  try {
    let chRaw = PropertiesService.getScriptProperties().getProperty('PAYMENT_CHANNELS');
    if (chRaw) payChannels = JSON.parse(chRaw);
  } catch(e) {}
  if (!payChannels.length) {
    payChannels = [
      { key:'cash',     label:'เงินสด',              color:'#27ae60' },
      { key:'transfer', label:'เงินโอน',              color:'#2980b9' },
      { key:'welfare',  label:'บัตรสวัสดิการแห่งรัฐ', color:'#8e44ad' }
    ];
  }

  // channelTotals: สะสมยอดแต่ละ key อย่างง่าย ไม่แยก Cash/QR/Mix
  // { cash:0, transfer:0, welfare:0, thaihelp:0, ... }
  let channelTotals = {};
  payChannels.forEach(ch => { channelTotals[ch.key] = 0; });

  let totalSales = 0;
  let billCount  = 0;

  for (let i = 1; i < headers.length; i++) {
    let h = headers[i];
    if (!h[0]) continue;
    let saleTime = new Date(h[1]);
    if (saleTime < openTime) continue;

    let total       = Number(h[2]) || 0;
    let type        = String(h[3]);
    let cash        = Number(h[4]) || 0;   // legacy
    let transfer    = Number(h[5]) || 0;   // legacy
    let welfare     = Number(h[7]) || 0;   // legacy
    let channelsRaw = h[8] ? String(h[8]) : '';

    totalSales += total;
    billCount++;

    if (channelsRaw) {
      // ระบบใหม่ — channels array: นับตาม key ตรงๆ
      try {
        let chs = JSON.parse(channelsRaw);
        chs.forEach(ch => {
          let amt = Number(ch.amount) || 0;
          if (!amt) return;
          if (channelTotals[ch.key] === undefined) {
            channelTotals[ch.key] = 0; // channel ใหม่ที่ไม่ได้ตั้งค่า
          }
          channelTotals[ch.key] += amt;
        });
      } catch(e) {
        // parse ล้มเหลว → fallback legacy
        if (type === 'Cash') {
          channelTotals['cash'] = (channelTotals['cash']||0) + total;
        } else if (type === 'QR') {
          channelTotals['transfer'] = (channelTotals['transfer']||0) + total;
        } else if (type === 'Mix') {
          channelTotals['cash']     = (channelTotals['cash']||0)     + cash;
          channelTotals['transfer'] = (channelTotals['transfer']||0) + transfer;
          channelTotals['welfare']  = (channelTotals['welfare']||0)  + welfare;
        }
      }
    } else {
      // legacy (ไม่มี channels array)
      if (type === 'Cash') {
        channelTotals['cash'] = (channelTotals['cash']||0) + total;
      } else if (type === 'QR') {
        channelTotals['transfer'] = (channelTotals['transfer']||0) + total;
      } else if (type === 'Mix') {
        channelTotals['cash']     = (channelTotals['cash']||0)     + cash;
        channelTotals['transfer'] = (channelTotals['transfer']||0) + transfer;
        channelTotals['welfare']  = (channelTotals['welfare']||0)  + welfare;
      }
    }
  }

  // ===== คำนวณยอดสรุป =====
  // เงินในเก๊ะ = startCash + ยอดที่ key='cash' เท่านั้น
  let totalCashReceived = channelTotals['cash'] || 0;
  let totalCashInDrawer = shift.startCash + totalCashReceived;

  // ยอดโอน = key='transfer' เท่านั้น (ไม่รวม welfare หรือช่องอื่น)
  let totalTransfer = channelTotals['transfer'] || 0;

  // legacy fields สำหรับ backward compat
  let cashSales = channelTotals['cash']     || 0;
  let qrSales   = channelTotals['transfer'] || 0;
  let mixCash   = 0;
  let mixTf     = 0;

  let diffMs  = now - openTime;
  let diffHrs = (diffMs / 3600000).toFixed(2);
  let closeTimeStr = Utilities.formatDate(now, tz, "dd/MM/yyyy HH:mm:ss");

  // ===== เขียนลง Shifts Sheet =====
  let ws = ss.getSheetByName("Shifts");

  // หา channels ที่มียอด (เอาทุก key รวมถึงที่สร้างเอง)
  // ใช้ payChannels เป็นลำดับ แต่ถ้ามี key ใหม่ต่อท้าย
  let channelKeys = [...payChannels.map(c=>c.key)];
  Object.keys(channelTotals).forEach(k=>{ if(!channelKeys.includes(k)) channelKeys.push(k); });
  // กรองเฉพาะที่มียอด
  let activeChannels = channelKeys
    .filter(k => (channelTotals[k]||0) > 0)
    .map(k => {
      let ch = payChannels.find(c=>c.key===k);
      return { key:k, label: ch?ch.label:k, color: ch?ch.color:'#7f8c8d', amount: channelTotals[k] };
    });

  let fixedHeaders = [
    "กะที่","เปิดเวลา","ปิดเวลา","ระยะเวลา(ชม.)",
    "เงินทอนเริ่มกะ","ยอดขายรวม","จำนวนบิล",
    "เงินในเก๊ะที่ต้องมี","ยอดโอน(transfer)"
  ];
  let channelLabels = activeChannels.map(c => c.label);
  let allHeaders    = [...fixedHeaders, ...channelLabels];

  if (!ws) {
    ws = ss.insertSheet("Shifts");
    ws.getRange(1,1,1,allHeaders.length).setValues([allHeaders]);
    ws.setFrozenRows(1);
    let hRange = ws.getRange(1,1,1,allHeaders.length);
    hRange.setFontWeight("bold").setBackground("#2c3e50").setFontColor("white").setHorizontalAlignment("center");
  } else {
    // เช็ค/อัปเดต header ถ้าจำเป็น
    let existingHeaders = ws.getRange(1,1,1,ws.getLastColumn()).getValues()[0].map(String);
    let needUpdate = allHeaders.some(h=>!existingHeaders.includes(h)) || existingHeaders.length !== allHeaders.length;
    if (needUpdate) {
      let clearCols = Math.max(existingHeaders.length, allHeaders.length);
      ws.getRange(1,1,1,clearCols).clearContent().clearFormat();
      ws.getRange(1,1,1,allHeaders.length).setValues([allHeaders]);
      ws.getRange(1,1,1,allHeaders.length)
        .setFontWeight("bold").setBackground("#2c3e50").setFontColor("white").setHorizontalAlignment("center");
      ws.setFrozenRows(1);
    }
  }

  // สร้างแถวข้อมูล
  let fixedRow   = [shift.shiftNo, shift.openTime, closeTimeStr, Number(diffHrs),
                    shift.startCash, totalSales, billCount, totalCashInDrawer, totalTransfer];
  let channelRow = activeChannels.map(c => c.amount);
  let fullRow    = [...fixedRow, ...channelRow];

  let dataRow = ws.getLastRow() + 1;
  ws.getRange(dataRow, 1, 1, fullRow.length).setValues([fullRow]);
  ws.getRange(dataRow, 1, 1, fullRow.length).setBackground(dataRow % 2 === 0 ? "#f8f9fa" : "white");

  // format ตัวเลข
  ws.getRange(dataRow, 4, 1, fullRow.length - 3).setNumberFormat("#,##0.00");

  // auto-resize
  for (let c2 = 1; c2 <= allHeaders.length; c2++) ws.autoResizeColumn(c2);

  PropertiesService.getScriptProperties().deleteProperty('CURRENT_SHIFT');

  // build channelSummary ส่งกลับ frontend
  let channelSummaryOut = {};
  activeChannels.forEach(c => { channelSummaryOut[c.key] = c; });

  return {
    shiftNo: shift.shiftNo, openTime: shift.openTime, closeTime: closeTimeStr,
    durationHrs: diffHrs, startCash: shift.startCash,
    totalSales, cashSales, qrSales, mixCash, mixTf,
    totalCashInDrawer, totalTransfer, billCount,
    channelSummary: channelSummaryOut
  };
}

// ===== resetShiftsSheet() — รัน 1 ครั้งเพื่อ rebuild header =====
function resetShiftsSheet() {
  const ss = SpreadsheetApp.getActive();
  let ws = ss.getSheetByName("Shifts");
  if (!ws) { Browser.msgBox("ไม่พบ sheet Shifts"); return; }

  let payChannels = [];
  try {
    let chRaw = PropertiesService.getScriptProperties().getProperty('PAYMENT_CHANNELS');
    if (chRaw) payChannels = JSON.parse(chRaw);
  } catch(e) {}
  if (!payChannels.length) {
    payChannels = [
      { key:'cash',     label:'เงินสด' },
      { key:'transfer', label:'เงินโอน' },
      { key:'welfare',  label:'บัตรสวัสดิการแห่งรัฐ' }
    ];
  }

  let fixedHeaders = [
    "กะที่","เปิดเวลา","ปิดเวลา","ระยะเวลา(ชม.)",
    "เงินทอนเริ่มกะ","ยอดขายรวม","จำนวนบิล",
    "เงินในเก๊ะที่ต้องมี","ยอดโอน(transfer)"
  ];
  let allHeaders = [...fixedHeaders, ...payChannels.map(c=>c.label)];

  let lastCol = Math.max(ws.getLastColumn(), allHeaders.length);
  ws.getRange(1,1,1,lastCol).clearContent().clearFormat();
  ws.getRange(1,1,1,allHeaders.length).setValues([allHeaders]);
  ws.getRange(1,1,1,allHeaders.length)
    .setFontWeight("bold").setBackground("#2c3e50").setFontColor("white").setHorizontalAlignment("center");
  ws.setFrozenRows(1);
  for (let c=1; c<=allHeaders.length; c++) ws.autoResizeColumn(c);

  Browser.msgBox("✅ อัปเดต header Shifts เรียบร้อย (" + allHeaders.length + " คอลัมน์)");
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
    ws.getRange(1,1,1,7).setValues([[
      "รหัสสมาชิก","ชื่อ","เบอร์โทร","แต้ม","วันที่สมัคร","หมายเหตุ","สถานะ"
    ]]);
    ws.setFrozenRows(1);
    ws.getRange(1,1,1,7).setFontWeight("bold").setBackground("#8e44ad").setFontColor("white");
  }
  return ws;
}

function getMembers() {
  const ws   = _getMembersSheet();
  const data = ws.getDataRange().getValues();
  let members = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    members.push({
      id:       String(data[i][0]),
      name:     String(data[i][1]),
      phone:    String(data[i][2] || ''),
      points:   Number(data[i][3]) || 0,
      joinDate: data[i][4] ? String(data[i][4]) : '',
      note:     String(data[i][5] || ''),
      active:   data[i][6] !== false && data[i][6] !== 'ไม่ใช้งาน'
    });
  }
  return members;
}

function searchMember(query) {
  let q = String(query).toLowerCase().trim();
  let all = getMembers();
  return all.filter(m =>
    m.id.toLowerCase().includes(q) ||
    m.name.toLowerCase().includes(q) ||
    m.phone.includes(q)
  ).slice(0, 8);
}

function addMember(data) {
  const ws  = _getMembersSheet();
  const all = ws.getDataRange().getValues();

  // ตรวจ id ซ้ำ
  for (let i = 1; i < all.length; i++) {
    if (String(all[i][0]) === String(data.id)) return { error: 'รหัสสมาชิกซ้ำ' };
  }

  const tz  = Session.getScriptTimeZone();
  ws.appendRow([
    String(data.id),
    String(data.name),
    String(data.phone || ''),
    0,
    Utilities.formatDate(new Date(), tz, "dd/MM/yyyy"),
    String(data.note || ''),
    'ใช้งาน'
  ]);
  return { success: true };
}

function updateMember(id, updates) {
  const ws   = _getMembersSheet();
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      if (updates.name  !== undefined) ws.getRange(i+1, 2).setValue(updates.name);
      if (updates.phone !== undefined) ws.getRange(i+1, 3).setValue(updates.phone);
      if (updates.note  !== undefined) ws.getRange(i+1, 6).setValue(updates.note);
      if (updates.active !== undefined) ws.getRange(i+1, 7).setValue(updates.active ? 'ใช้งาน' : 'ไม่ใช้งาน');
      return { success: true };
    }
  }
  return { error: 'ไม่พบสมาชิก' };
}

function getMemberHistory(memberId) {
  const ss  = SpreadsheetApp.getActive();
  const wsH = ss.getSheetByName("Sales_Header");
  const wsD = ss.getSheetByName("Sales_Detail");
  const tz  = Session.getScriptTimeZone();

  let headers = wsH.getDataRange().getValues();
  let details = wsD.getDataRange().getValues();

  // Sales_Header อาจมี column memberId (column 7, index 6)
  let bills = [];
  for (let i = headers.length - 1; i >= 1; i--) {
    let h = headers[i];
    if (!h[0]) continue;
    if (String(h[6] || '') !== String(memberId)) continue;

    let items = [];
    for (let j = 1; j < details.length; j++) {
      if (details[j][0] == h[0]) {
        items.push({
          name:  String(details[j][2]),
          qty:   Number(details[j][5]) || 0,
          total: Number(details[j][6]) || 0
        });
      }
    }
    let chs4 = [];
    try { if(h[8]) chs4 = JSON.parse(h[8]); } catch(e){}
    bills.push({
      billNo:   h[0],
      date:     Utilities.formatDate(new Date(h[1]), tz, "dd/MM/yyyy HH:mm"),
      total:    Number(h[2]) || 0,
      type:     String(h[3]),
      cash:     Number(h[4]) || 0,
      transfer: Number(h[5]) || 0,
      welfare:  Number(h[7]) || 0,
      channels: chs4,
      items
    });
    if (bills.length >= 50) break;
  }
  return bills;
}

// ===== PAYMENT CHANNELS SETTINGS =====
function getPaymentChannels() {
  try {
    let raw = PropertiesService.getScriptProperties().getProperty('PAYMENT_CHANNELS');
    if (raw) return JSON.parse(raw);
  } catch(e) {}
  // default channels
  return [
    { key:'cash',     label:'💵 เงินสด',                color:'#27ae60' },
    { key:'transfer', label:'📱 เงินโอน',               color:'#2980b9' },
    { key:'welfare',  label:'💳 บัตรสวัสดิการแห่งรัฐ', color:'#8e44ad' }
  ];
}

function savePaymentChannels(channels) {
  PropertiesService.getScriptProperties()
    .setProperty('PAYMENT_CHANNELS', JSON.stringify(channels));
  return true;
}

function getWebAppUrl(){
  return ScriptApp.getService().getUrl();
}

// ===== CLEAR DISPLAY DATA (เรียกหลังแสดง paid แล้ว) =====
function clearDisplayData() {
  try {
    PropertiesService.getScriptProperties().deleteProperty('DISPLAY_DATA');
  } catch(e) {}
  return true;
}
