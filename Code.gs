// ============================================================
//  DAILY FLEET MASTER — Code.gs
//  เชื่อมต่อกับ Sheet "ผังที่นั่ง" (กรกฎาคม 2569)
// ============================================================

// ════════════════════════════════════════════════════════════
//  onEdit TRIGGER — Auto สีเขียวเมื่อพิมพ์ชื่อลงในที่นั่ง
// ════════════════════════════════════════════════════════════
const THAI_MONTHS = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];

function getSeatSheetName(month, year) {
  let yr = parseInt(year);
  if (yr < 2400) yr += 543; // ถ้าเป็น ค.ศ. (เช่น 2026) ให้บวกเป็น พ.ศ. (2569)
  const mIndex = parseInt(month) - 1;
  return "ผังที่นั่ง " + THAI_MONTHS[mIndex] + " " + yr;
}

function onEdit(e) {
  if (!e || !e.range) return; 
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  if (!sheetName.startsWith("ผังที่นั่ง")) return;

  const cell = e.range;
  const row = cell.getRow();
  const col = cell.getColumn();
  const value = cell.getValue();

  // 1. ตรวจสอบว่าเป็นส่วนของผังที่นั่งหรือไม่ (แถว 5-13, 17-25, ...)
  const relativeRow = (row - PLAN_HEADER_ROWS - OFF_SEATS_START - 1) % ROWS_PER_DAY;
  if (relativeRow < 0 || relativeRow >= 9) return; 

  // 2. ตรวจสอบคอลัมน์ (ต้องเป็น A, B, C, D ของ Slot ใด Slot หนึ่ง)
  let isSeatCol = false;
  // เช็ค Hat Yai (สูงสุด 10 รอบ)
  for (let s = 0; s < 10; s++) {
    const hStart = HATYAI_START_COL + (s * SLOT_WIDTH);
    const validH = [hStart, hStart + 1, hStart + 3, hStart + 4];
    if (validH.includes(col)) { isSeatCol = true; break; }
  }
  // เช็ค Phuket (สูงสุด 6 รอบ)
  if (!isSeatCol) {
    for (let s = 0; s < 6; s++) {
      const pStart = getSlotStartCol('phuket', s, sheet);
      const validP = [pStart, pStart + 1, pStart + 3, pStart + 4];
      if (validP.includes(col)) { isSeatCol = true; break; }
    }
  }
  if (!isSeatCol) return;

  // 3. เปลี่ยนสีตามค่าที่กรอก
  if (value && value.toString().trim() !== "") {
    cell.setBackground(COLOR_PREBOOK)
        .setFontColor('#ffffff')
        .setFontWeight('bold');
  } else {
    cell.setBackground(null)
        .setFontColor('#94a3b8')
        .setFontWeight('normal');
  }
}

// ── Sheet Names ──────────────────────────────────────────────
const SHEET_SEAT_PLAN = "ผังที่นั่ง";
const SHEET_LOG       = "Log_ตั๋วสากล";
const SHEET_CONFIG    = "Config_ราคา";
const SHEET_TRIPS     = "รายการเที่ยว";
const SHEET_CUST      = "ข้อมูลลูกค้า";
const ADMIN_PIN       = "1234"; // รหัสผ่านสำหรับแก้ไขหรือยกเลิก (เปลี่ยนได้ที่นี่)

// ── Constants ──────────────────────────────────────────────
const PLAN_HEADER_ROWS = 2;         // แถว Header (Row 1-2) ก่อนเริ่มข้อมูลวัน
const ROWS_PER_DAY     = 12;        // 1 วัน = 12 แถว
const OFF_BUS_INFO     = 0;         // อยู่ Row 3 (startRow)
const OFF_PHONE_TIME   = 1;         // อยู่ Row 4 (startRow + 1)
const OFF_SEATS_START  = 2;         // เริ่มที่นั่ง Row 5 (startRow + 2)
const OFF_DAY_SEP      = 11;        // แถวสีน้ำเงินคั่นวัน

// ── สี ──────────────────────────────────────────────────────
const COLOR_PREBOOK  = "#22c55e";   // เขียว = จองแล้ว (กรอกชื่อในชีทเอง)
const COLOR_CONFIRM  = "#3b82f6";   // ฟ้า   = ยืนยันแล้ว (ผ่านเว็บ)
const COLOR_VACANT   = "#ffffff";   // ขาว   = ว่าง
const COLOR_STAIR    = "#2563eb";   // บันไดกลาง

// ── Column Structure ─────────────────────────────────────────
const SLOT_WIDTH = 6;  // แต่ละ Slot กว้าง 6 คอลัมน์ (A,B,gap,C,D,sep)

const HATYAI_START_COL = 2;  // Column B
const HATYAI_TIMES     = ["09:00", "12:40", "14:00", "16:00", "22:00", "รอบเสริม 1", "รอบเสริม 2", "รอบเสริม 3", "รอบเสริม 4", "รอบเสริม 5"];

const PHUKET_START_COL = 38; // กลับมาใช้ 38 เป็นค่าเริ่มต้น
const PHUKET_TIMES     = ["07:45", "09:45", "11:45", "14:00", "22:30", "รอบเสริม"];

// ── ฟังก์ชันช่วย ─────────────────────────────────────────────
function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// หา startRow ของวันใน Sheet (day = 1-31)
function getDayStartRow(day) {
  return PLAN_HEADER_ROWS + 1 + (day - 1) * ROWS_PER_DAY;
}

// หา startCol ของ Slot (slotIndex = 0,1,2,...)
function getSlotStartCol(station, slotIndex, sheet) {
  let baseCol = (station === 'phuket') ? PHUKET_START_COL : HATYAI_START_COL;
  
  // ถ้าเป็นภูเก็ต และใน Sheet มีการขยายคอลัมน์ (Hat Yai มี 10 รอบ) ให้เลื่อนภูเก็ตไปเริ่มที่ 62
  if (station === 'phuket' && sheet && sheet.getMaxColumns() >= 90) {
    baseCol = 62;
  }
  
  return baseCol + slotIndex * SLOT_WIDTH;
}

// ── doGet ────────────────────────────────────────────────────
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'dashboard';
  if (page === 'ticket') {
    const template = HtmlService.createTemplateFromFile('ticket');
    template.payload = (e && e.parameter && e.parameter.d) || '';
    return template.evaluate()
      .setTitle('พิมพ์ตั๋วโดยสาร')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (page === 'manifest') {
    const template = HtmlService.createTemplateFromFile('manifest');
    template.payload = (e && e.parameter && e.parameter.d) || '';
    return template.evaluate()
      .setTitle('พิมพ์ใบผังที่นั่ง (Trip Manifest)')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (page === 'manual' || page === 'manual_ticket') {
    const template = HtmlService.createTemplateFromFile('manual_ticket');
    template.id = (e && e.parameter && e.parameter.id) || '';
    return template.evaluate()
      .setTitle('ออกตั๋วแมนนวล')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (page === 'print') {
    return HtmlService.createTemplateFromFile('print')
      .evaluate()
      .setTitle('ใบบันทึกเที่ยวรถ')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  const template = HtmlService.createTemplateFromFile('index');
  template.page = page;
  return template.evaluate()
    .setTitle('ระบบบันทึกการขนส่ง')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ════════════════════════════════════════════════════════════
//  READ — ดึงข้อมูลจาก Sheet ผังที่นั่ง
// ════════════════════════════════════════════════════════════

function getSeatMap(day, month, year, station, slotIndex) {
  try {
    const ss    = getSS();
    const sheetName = getSeatSheetName(month, year);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'ไม่พบ Sheet: ' + sheetName };

    const startRow = getDayStartRow(day);
    const startCol = getSlotStartCol(station, slotIndex, sheet);

    // ตำแหน่งตามใบงานออนไลน์ (Verified Precisely)
    // BUS_NO: B(startCol) + C(startCol+1)
    const busNoPart1 = sheet.getRange(startRow,  startCol).getDisplayValue();
    const busNoPart2 = sheet.getRange(startRow,  startCol + 1).getDisplayValue();
    const busNo = (busNoPart1 + busNoPart2).trim() || '-';
    
    // พขร (DRIVER/PHONE): อยู่ในเซลล์รวม D:F (startCol+2)
    // สูตรในชีทคือ นายหัวน้อย\n098...
    const driverRaw = sheet.getRange(startRow, startCol + 2).getDisplayValue() || '-';
    
    // เวลาออก (TIME): เลื่อนไปอยู่ที่คอลัมน์ F (startCol+4) เพราะ E คือคำว่า "เวลา"
    const timeValue = sheet.getRange(startRow + 1, startCol + 4).getDisplayValue() || '-';

    // แยกชื่อกับเบอร์โทรที่ดึงมาจากสูตร (ใช้ \n เป็นตัวคั่น)
    let driverName = driverRaw;
    let driverPhone = '-';
    if (driverRaw.includes('\n')) {
      const parts = driverRaw.split('\n');
      driverName = parts[0].trim();
      driverPhone = parts[1].trim();
    } else if (driverRaw.match(/\d{9,10}/)) {
      driverPhone = driverRaw.match(/\d{9,10}/)[0];
      driverName = driverRaw.replace(driverPhone, '').trim();
    }

    const timeLabel = (station === 'phuket' ? PHUKET_TIMES : HATYAI_TIMES)[slotIndex];
    const originLabel = (station === 'phuket' ? 'ภูเก็ต' : 'หาดใหญ่');
    const dateStr = `${day}/${month}/${year}`;

    // --- ดึง Metadata จาก Log_ตั๋วสากล (Smart Lookup 3.0) ---
    const logDataMap = {};
    const logSheet = ss.getSheetByName(SHEET_LOG);
    if (logSheet) {
      const logs = logSheet.getDataRange().getValues();
      const normalize = (val) => val.toString().replace(/^0/, '').trim().toUpperCase();
      const targetTimeNorm   = normalize(timeLabel);
      const targetDateSimple = day + "/" + month; 

      for (let i = logs.length - 1; i >= 1; i--) {
        const row = logs[i];
        if (!row[2] || !row[7]) continue;

        const logDateStr = normalize(row[2]); 
        const logOrigin  = normalize(row[3]);
        const logTime    = normalize(row[6]);
        const logSeat    = normalize(row[7]);

        const isDateMatch = logDateStr.includes(targetDateSimple) || logDateStr.includes(day + "/" + (month < 10 ? '0'+month : month));
        const isOriginMatch = (logOrigin === normalize(originLabel));
        const isTimeMatch = (logTime === targetTimeNorm);

        if (isDateMatch && isOriginMatch && isTimeMatch) {
          if (!logDataMap[logSeat]) {
            logDataMap[logSeat] = {
              ticketId: row[0]  || '',
              dest:     row[4]  || '',
              boarding: row[5]  || '',
              price:    row[10] || '',
              phone:    row[15] || '',
              paymentType: row[11] || 'cash'
            };
          }
        }
      }
    }

    // --- ดึงข้อมูลเสริมจาก 'ข้อมูลลูกค้า' (Source of Truth) ---
    const custSheet = ss.getSheetByName('ข้อมูลลูกค้า');
    if (custSheet) {
      const custData = custSheet.getDataRange().getValues(); 
      const now = new Date().getTime();

      for (let i = custData.length - 1; i >= 1; i--) {
        const row = custData[i];
        if (!row[4] || !row[9]) continue; 

        // ตรวจสอบการหมดเวลาของ LOCKED
        if (row[21] === "LOCKED" && row[22]) {
          const expireTime = new Date(row[22]).getTime();
          if (now > expireTime) {
            // หมดเวลา! คืนที่นั่งในผังรถ
            const startRowLoc = getDayStartRow(day);
            const startColLoc = getSlotStartCol(station, slotIndex);
            const seatIdLoc = row[9].toString().trim();
            const seatPos = _getSeatCellPos(seatIdLoc, startRowLoc, startColLoc);
            if (seatPos) {
              sheet.getRange(seatPos.row, seatPos.col).setBackground(null).setValue('').setFontWeight('normal');
            }
            // มาร์คสถานะในฐานข้อมูล
            custSheet.getRange(i + 1, 22).setValue("EXPIRED");
            continue; // ไม่ต้องดึงข้อมูลนี้เข้า Map
          }
        }

        const normalize = (val) => val.toString().trim().toUpperCase();
        const logDateStr = Utilities.formatDate(new Date(row[4]), "GMT+7", "d/M");
        const targetDateStr = day + "/" + month;
        
        const logTime    = normalize(row[8]);
        const logOrigin  = normalize(row[5]);
        const targetOrigin = normalizeLocation(originLabel).toUpperCase();
        const targetTime   = normalizeTime(timeLabel).toUpperCase();

        if (logDateStr.includes(targetDateStr) && logTime.includes(targetTime) && logOrigin.includes(targetOrigin)) {
          const logSeatsRaw = row[9] ? row[9].toString().split(/[\s,]+/) : [];
          const logSeats = logSeatsRaw.map(s => normalize(s).replace(/^0/, '')).filter(s => s !== '');
          
          logSeats.forEach(logSeat => {
            if (logSeat && !logDataMap[logSeat]) {
              logDataMap[logSeat] = {
                ticketId: row[0]  || '',
                name:     row[1]  || '',
                phone:    row[2]  || '',
                dest:     row[6]  || '',
                boarding: row[7]  || '',
                price:    row[12] || '',
                time:     row[8]  || '',
                isCash:   (row[13] && row[13].toString().toLowerCase() === 'true'),
                isQR:     (row[14] && row[14].toString().toLowerCase() === 'true'),
                isAgent:  (row[15] && row[15].toString().toLowerCase() === 'true'),
                isCard:   (row[16] && row[16].toString().toLowerCase() === 'true'),
                isInspector: (row[17] && row[17].toString().toLowerCase() === 'true'),
                issuedBy: row[18] || '', 
                allSeats: row[9]  || '',
                gender:   row[20] || '', 
                status:   row[21] || 'ACTIVE', 
                remarks:   row[22] || '', 
                expireTime: row[23] || '' 
              };
            }
          });
        }
      }
    }

    const seats = {};
    const seatIds = ['A','B','C','D'];
    const seatCols = [startCol, startCol + 1, startCol + 3, startCol + 4]; // B, C, E, F

    for (let r = 1; r <= 9; r++) {
      let rowNum = startRow + 2 + (r - 1);
      let isStairRow = (r === 5 || r === 6);

      seatCols.forEach((colIdx, i) => {
        const seatId = r + seatIds[i];
        if (isStairRow && (seatIds[i] === 'A' || seatIds[i] === 'B')) {
          seats[seatId] = { status: 'stair', name: '🪜 บันไดกลาง / ทางเดิน' };
        } else {
          const range = sheet.getRange(rowNum, colIdx);
          const bg    = range.getBackground().toLowerCase();
          const value = range.getDisplayValue().trim();
          let status = 'vacant';
          if (bg === '#3b82f6' || bg === '#4285f4' || bg === '#cfe2f3' || bg === '#0000ff') status = 'confirmed';
          else if (bg === '#22c55e' || bg === '#34a853' || bg === '#b7e1cd' || bg === '#00ff00') status = 'prebooked';
          else if (bg === '#f59e0b' || bg === '#ff9900') status = 'locked';
          
          seats[seatId] = { 
            status: status, 
            name: (status !== 'vacant') ? value : '',
            metadata: logDataMap[seatId.toUpperCase()] || null 
          };
        }
      });
    }

    return {
      success: true,
      header: {
        busNo:  busNo,
        driver: driverName,
        phone:  driverPhone,
        time:   timeValue,
        day, month, year, station, slotIndex
      },
      seats: seats,
      isExpanded: sheet.getMaxColumns() >= 90
    };


  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function getSlotSummary(day, month, year, station) {
  try {
    const ss    = getSS();
    const sheetName = getSeatSheetName(month, year);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'ไม่พบ Sheet: ' + sheetName };

    const times   = (station === 'phuket') ? PHUKET_TIMES : HATYAI_TIMES;
    const results = [];

    for (let i = 0; i < times.length; i++) {
      const startRow = getDayStartRow(day);
      const startCol = getSlotStartCol(station, i, sheet);

      let vacant = 0, prebooked = 0, confirmed = 0;
      for (let r = 1; r <= 9; r++) {
        const rowNum = startRow + OFF_SEATS_START + (r - 1);
        const isStair = (r === 5 || r === 6);
        const cols = isStair ? [startCol + 3, startCol + 4] : [startCol, startCol + 1, startCol + 3, startCol + 4];
        cols.forEach(col => {
          const bg = sheet.getRange(rowNum, col).getBackground().toLowerCase();
          if (bg === '#3b82f6' || bg === '#4285f4' || bg === '#cfe2f3' || bg === '#0000ff') confirmed++;
          else if (bg === '#22c55e' || bg === '#34a853' || bg === '#b7e1cd' || bg === '#00ff00') prebooked++;
          else vacant++;
        });
      }

      const busNo = sheet.getRange(startRow + OFF_BUS_INFO, startCol + 1).getDisplayValue();
      results.push({
        slotIndex: i,
        time:      times[i],
        busNo:     busNo || '-',
        vacant, prebooked, confirmed,
        total:     vacant + prebooked + confirmed
      });
    }

    const isExpanded = sheet.getMaxColumns() >= 90;
    return { success: true, slots: results, times, isExpanded: isExpanded };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ════════════════════════════════════════════════════════════
//  TICKET SYSTEM — จัดการเลขตั๋วและ Log พิเศษ
// ════════════════════════════════════════════════════════════

function generateTicketId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_CUST);
  let lastNum = 0;
  if (logSheet) {
    const ids = logSheet.getRange("A:A").getValues();
    for (let i = ids.length - 1; i >= 1; i--) {
      const val = ids[i][0].toString();
      if (val.startsWith("online")) {
        const numPart = val.replace("online", "");
        lastNum = parseInt(numPart) || 0;
        break;
      }
    }
  }
  const nextNum = lastNum + 1;
  return "online" + nextNum.toString().padStart(6, '0');
}

/**
 * ดึง URL ของ Web App
 */
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

function getTicketData(ticketId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    SpreadsheetApp.flush();
    
    let sheet = ss.getSheetByName("ข้อมูลลูกค้า");
    if (!sheet) {
      const sheets = ss.getSheets();
      for (let s of sheets) {
        if (s.getSheetId().toString() === "1305369588") {
          sheet = s;
          break;
        }
      }
    }

    if (!sheet) return { success: false, error: 'ไม่พบ Sheet ข้อมูลลูกค้า (กรุณาเช็คชื่อชีทหรือ GID)' };

    const data = sheet.getDataRange().getDisplayValues();
    const header = data[0];
    const searchId = ticketId.toString().split(',')[0].trim().toLowerCase();

    let targetRow = null;
    for (let i = data.length - 1; i >= 1; i--) {
      const currentId = data[i][0].toString().trim();
      const currentIdLower = currentId.toLowerCase();
      let isMatch = (currentIdLower === searchId);

      if (!isMatch && /^\d+$/.test(searchId)) {
        const numericPart = currentId.match(/\d+$/);
        if (numericPart && Number(numericPart[0]) === Number(searchId)) isMatch = true;
      }

      if (isMatch) {
        const hasName = data[i][1] && data[i][1].toString().trim() !== "";
        const hasSeat = data[i][9] && data[i][9].toString().trim() !== "";

        if (hasName && hasSeat) {
           targetRow = data[i];
           break;
        }
        if (hasName && !targetRow) {
           targetRow = data[i];
        }
        if (!targetRow) targetRow = data[i];
      }
    }

    if (!targetRow) return { success: false, error: 'ไม่พบรหัสตั๋ว: ' + ticketId };

    const allTickets = [];
    for (let i = 1; i < data.length; i++) {
        const currentId = data[i][0].toString().trim().toLowerCase();
        if (currentId === searchId) {
            allTickets.push(extractTicketRow(data[i]));
        }
    }

    if (allTickets.length > 1) {
        return { success: true, isGroup: true, tickets: allTickets, ticketId: targetRow[0] };
    } else {
        return extractTicketRow(targetRow);
    }

  } catch (e) {
    console.error("Error: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

// ฟังก์ชันช่วยสกัดข้อมูลจากแถว (เพื่อลดความซ้ำซ้อน)
function extractTicketRow(row) {
    // อ่าน isAgent จากคอลัมน์ 16 (boolean)
    const isAgentFlag = (row[15] && row[15].toString().toLowerCase() === 'true');
    
    let payType = 'cash';
    if (row[13] && row[13].toString().toLowerCase() === 'true') payType = 'cash';
    if (row[14] && row[14].toString().toLowerCase() === 'true') payType = 'qr';
    if (isAgentFlag) payType = 'agent';
    if (row[16] && row[16].toString().toLowerCase() === 'true') payType = 'card';
    if (row[17] && row[17].toString().toLowerCase() === 'true') payType = 'inspector';

    const dateStrLog = row[3] || "";
    const dateParts = dateStrLog.toString().split('/');
    
    const result = {
        success: true,
        ticketNo: row[0],
        name: row[1],
        phone: row[2],
        date: dateStrLog,
        day: dateParts.length === 3 ? dateParts[0] : "",
        month: dateParts.length === 3 ? dateParts[1] : "",
        year: dateParts.length === 3 ? dateParts[2] : "",
        origin: row[5],
        destination: row[6],
        boarding: row[7],
        time: row[8],
        seatId: row[9],
        busNo: row[11],
        platform: row[10],
        price: row[12],
        isCash: (row[13] && row[13].toString().toLowerCase() === 'true'),
        isQR: (row[14] && row[14].toString().toLowerCase() === 'true'),
        isAgent: isAgentFlag,
        isCard: (row[16] && row[16].toString().toLowerCase() === 'true'),
        isInspector: (row[17] && row[17].toString().toLowerCase() === 'true'),
        priceType: payType,
        issuedBy: row[18],
        gender: row[20], 
        status: row[21], 
        remarks: row[22] || "",
        expireTime: row[23] || "", 
    };

    // พยายามเดาสถานีและรอบรถเพื่ออัปเดตสีผังที่นั่งได้ถูกต้อง
    let station = 'hatyai';
    if (row[5] && row[5].toString().includes('ภูเก็ต')) station = 'phuket';

    let slotIndex = 0;
    const times = (station === 'phuket' ? PHUKET_TIMES : HATYAI_TIMES);
    const seatTime = normalizeTime(row[8] || "");
    for (let j = 0; j < times.length; j++) {
      if (normalizeTime(times[j]) === seatTime) {
        slotIndex = j;
        break;
      }
    }

    // แยกวัน เดือน ปี (เผื่อนำไปใช้ในฟอร์มได้แม่นยำขึ้น)
    let d=1, m=1, y=2024;
    if (row[4]) {
      const parts = row[4].split('/');
      if (parts.length === 3) {
        d = parseInt(parts[0]);
        m = parseInt(parts[1]);
        y = parseInt(parts[2]);
        if (y > 2500) y -= 543; // แปลง พ.ศ. เป็น ค.ศ.
      }
    }

    const seatIdRaw = row[9] ? row[9].toString() : "";
    const seats = seatIdRaw.split(/[\s,]+/).filter(s => s.trim() !== "");

    // Auto-detect old (20-col) vs new (21-col) schema
    const isNewRow = (row[16] === true || row[16] === false ||
                      row[16].toString().toLowerCase() === 'true' ||
                      row[16].toString().toLowerCase() === 'false' ||
                      row[16].toString().trim() === '');

    // ค้นหาสถานะจากคอลัมน์ V (Index 21) หรือคำว่า CANCELLED
    let isCancelled = false;
    if (row.length >= 22 && row[21] && row[21].toString().toUpperCase() === "CANCELLED") isCancelled = true;

    // =============== โค้ดส่วนที่แก้ไขเพื่อส่งค่าวิธีชำระและหมายเหตุ ================
    const data = {
      success: true,
      ticketNo:    row[0], 
      name:        isCancelled ? (row[1] + " (ยกเลิกแล้ว)") : row[1], 
      phone:       row[2], 
      issuedDate:  row[3], 
      date:        row[4],
      day: d, month: m, year: y, 
      origin:      row[5], 
      destination: row[6], 
      boarding:    row[7], 
      time:        row[8], 
      seatId:      seatIdRaw || (isCancelled ? "คืนที่นั่งแล้ว" : ""), 
      platform:    row[10],
      busNo:       row[11],
      price:       row[12],
      priceType:   payType, 
      isCash:      (row[13] && row[13].toString().toLowerCase() === 'true'),
      isQR:        (row[14] && row[14].toString().toLowerCase() === 'true'),
      isAgent:     isAgentFlag,
      isCard:      (row[16] && row[16].toString().toLowerCase() === 'true'),
      isInspector: (row[17] && row[17].toString().toLowerCase() === 'true'),
      agentName:   isAgentFlag ? (isNewRow ? (row[18] || '') : (row[16] || '')) : '',
      issuedBy:    isNewRow ? (row[18] || '') : (row[16] || ''),
      remark:      isCancelled ? "ตั๋วใบนี้ถูกยกเลิกแล้ว" : (row[22] || ""), 
      gender:      isNewRow ? (row[20] || '') : (row[19] || ''),
      status:      isCancelled ? "CANCELLED" : "ACTIVE",
      station:     station,
      slotIndex:   slotIndex
    };
    // =====================================================================

    if (seats.length > 1) {
        // แตกข้อมูลเป็นรายที่นั่งสำหรับกรณีกลุ่ม
        const totalPrice = parseFloat(row[12]) || 0;
        const pricePerSeat = Math.floor(totalPrice / seats.length);
        const ticketList = seats.map(s => ({
            ...data,
            seatId: s,
            price: pricePerSeat
        }));
        return {
            success: true,
            isGroup: true,
            tickets: ticketList,
            ticketNo: row[0]
        };
    }

    return data;
}

/**
 * ยืนยันการจองที่นั่ง — แก้ไขใหม่: ไม่อัปเดตลง Sheet ข้อมูลลูกค้าแล้ว (ใช้สำหรับสถานะบนเว็บและคำนวณยอดเท่านั้น)
 */
function confirmSeatBooking(params) {
  try {
    const ss = getSS();
    // ถ้าส่งมาเป็นกลุ่ม (Array ของ bookings)
    if (params.bookings && Array.isArray(params.bookings)) {
      const results = [];
      params.bookings.forEach(b => {
        // ให้ระบบ _processSingleBooking สร้างรหัสใหม่แยกรายที่นั่ง (ตามที่ผู้ใช้ต้องการ)
        if (!b.ticketId) b.ticketId = generateTicketId();
        const res = _processSingleBooking(ss, b);
        results.push(res);
      });
      const allTicketIds = results.map(r => r.ticketId).join(', ');
      return { 
        success: true, 
        message: 'บันทึกการจองกลุ่มเรียบร้อย',
        ticketId: allTicketIds // คืนรหัสทั้งหมดรวมกัน
      };
    } else {
      // โหมดปกติ (1 รายการ หรือจองแบบเดิม)
      return _processSingleBooking(ss, params);
    }
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ฟังก์ชันภายในสำหรับประมวลผล 1 รายการจอง
function _processSingleBooking(ss, params) {
    const sheetName = getSeatSheetName(params.month, params.year);
    const sheet = ss.getSheetByName(sheetName);
    const { ticketId: existingId, day, month, year, station, slotIndex, seatId, passengerName, phone, origin, destination, boardingPoint, price, platform, busNo, remarks, time, issuedBy, gender, isCash, isQR, isAgent, isCard, isInspector: isInsp } = params;

    const startRow = getDayStartRow(day);
    const startCol = getSlotStartCol(station, slotIndex);
    const seatList = seatId ? seatId.toString().split(/[\s,]+/).filter(s => s.trim() !== "") : [];
    
    // [1] อัปเดตผังที่นั่ง (UI บน Sheet)
    if (!passengerName || passengerName.trim() === "") {
      seatList.forEach(s => {
        const pos = _getSeatCellPos(s, startRow, startCol);
        if (pos) {
          sheet.getRange(pos.row, pos.col).clearDataValidations()
               .setValue(s).setBackground(COLOR_VACANT).setFontColor('#94a3b8').setFontWeight('normal');
        }
      });
    } else {
      seatList.forEach(s => {
        const pos = _getSeatCellPos(s, startRow, startCol);
        if (pos) {
          sheet.getRange(pos.row, pos.col).clearDataValidations()
               .setValue(passengerName).setBackground(COLOR_CONFIRM).setFontColor('#ffffff').setFontWeight('bold');
        }
      });
    }
    SpreadsheetApp.flush(); // เร่งการอัปเดตสีผังที่นั่งให้เห็นผลทันที
    
    // [2] อัปเดตฐานข้อมูลลูกค้า
    const custSheet = ss.getSheetByName(SHEET_CUST);
    let targetRow = -1;
    let finalTicketId = existingId;

    if (existingId) {
      const ids = custSheet.getRange("A:A").getValues();
      for (let i = ids.length - 1; i >= 1; i--) {
        if (ids[i][0].toString() === existingId.toString()) {
          targetRow = i + 1;
          break;
        }
      }
    }

    if (targetRow === -1 && passengerName && passengerName.trim() !== "") {
      const data = custSheet.getDataRange().getValues();
      const travelDate = `${day}/${month}/${year}`;
      const originLabel = (station === 'phuket' ? 'ภูเก็ต (Phuket)' : 'หาดใหญ่ (Hatyai)');
      
      for (let i = data.length - 1; i >= 1; i--) {
        const r = data[i];
        const sameTrip = (r[4] === travelDate && r[5] === originLabel && r[8] === (time || ""));
        const sameSeat = (r[9] === seatId);
        const isNotCancelled = (r[21] !== "CANCELLED");

        if (sameTrip && sameSeat && isNotCancelled) {
          targetRow = i + 1;
          finalTicketId = r[0]; // ใช้รหัสตั๋วเดิม
          break;
        }
      }
    }

    if (!passengerName || passengerName.trim() === "") {
        if (existingId && targetRow > 0) {
            const currentRowData = custSheet.getRange(targetRow, 1, 1, 13).getValues()[0];
            const currentSeatId = currentRowData[9].toString(); 
            const currentPrice = parseFloat(currentRowData[12]) || 0; 
            
            const currentSeats = currentSeatId.split(/[\s,]+/).filter(s => s.trim() !== "");
            const seatsToCancel = seatId ? seatId.toString().split(/[\s,]+/).filter(s => s.trim() !== "") : [];
            
            const remainingSeats = currentSeats.filter(s => !seatsToCancel.includes(s));
            
            if (remainingSeats.length > 0) {
                const pricePerSeat = currentPrice / currentSeats.length;
                const newPrice = pricePerSeat * remainingSeats.length;
                
                custSheet.getRange(targetRow, 10).setValue(remainingSeats.join(', ')); 
                custSheet.getRange(targetRow, 13).setValue(newPrice); 
                
                if (custSheet.getLastColumn() >= 22) {
                    custSheet.getRange(targetRow, 22).setValue("EDITED");
                }
            } else {
                if (custSheet.getRange(1, 22).getValue() !== "สถานะ") {
                    custSheet.getRange(1, 22).setValue("สถานะ");
                }
                custSheet.getRange(targetRow, 22).setValue("CANCELLED");
            }
        }
        return { success: true, message: 'คืนที่นั่งเรียบร้อย' };
    }

    if (!finalTicketId) finalTicketId = generateTicketId();
    
    const now = new Date();
    const dateStr = Utilities.formatDate(now, "GMT+7", "dd/MM/yyyy");
    const timestampStr = Utilities.formatDate(now, "GMT+7", "dd/MM/yyyy HH:mm:ss");
    const travelDate = `${day}/${month}/${year}`;
    
    let timeLabel = time || "";
    if (!timeLabel) {
      const rawTimeLabel = (station === 'phuket' ? PHUKET_TIMES : HATYAI_TIMES)[slotIndex];
      timeLabel = normalizeTime(rawTimeLabel);
    }

    const originLabel = origin || (station === 'phuket' ? 'ภูเก็ต (Phuket)' : 'หาดใหญ่ (Hatyai)');
    const normalizedDest = normalizeLocation(destination);

    let finalStatus = existingId ? "EDITED" : "CONFIRMED";
    let expireTimeStr = "";
    
    if (params.isLock) {
      finalStatus = "LOCKED";
      const expireDate = new Date(now.getTime() + (params.lockMinutes || 1) * 60000);
      expireTimeStr = expireDate.toISOString();
    }

    // มีทั้งหมด 24 คอลัมน์ ไม่รวม AgentName แล้ว
    const rowData = [
      finalTicketId, passengerName, phone, dateStr, travelDate, originLabel, normalizedDest, boardingPoint, timeLabel, seatId,
      platform, busNo, price, isCash, isQR, isAgent, isCard, isInsp, issuedBy, timestampStr, (gender || ""), finalStatus, (params.remarks || ""), expireTimeStr
    ];

    if (targetRow > 0) {
      custSheet.getRange(1, 21).setValue("เพศ");
      custSheet.getRange(1, 22).setValue("สถานะ");
      custSheet.getRange(1, 23).setValue("หมายเหตุ");
      custSheet.getRange(1, 24).setValue("เวลาหมดอายุ");
      
      custSheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
    } else {
      const dataA = custSheet.getRange("A:A").getValues();
      targetRow = dataA.length + 1;
      for (let i = 0; i < dataA.length; i++) {
        if (dataA[i][0] === "") { targetRow = i + 1; break; }
      }
      custSheet.getRange(1, 21).setValue("เพศ");
      custSheet.getRange(1, 22).setValue("สถานะ");
      custSheet.getRange(1, 23).setValue("หมายเหตุ");
      custSheet.getRange(1, 24).setValue("เวลาหมดอายุ");
      custSheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
    }

    if (params.isLock) {
      const startRowLoc = getDayStartRow(day);
      const startColLoc = getSlotStartCol(station, slotIndex);
      const planSheetObj = sheet; 
      
      seatList.forEach(s => {
        const seatPos = _getSeatCellPos(s, startRowLoc, startColLoc);
        if (seatPos) {
          planSheetObj.getRange(seatPos.row, seatPos.col)
            .setBackground('#f59e0b')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setValue(passengerName || 'ล็อคชั่วคราว');
        }
      });
    }

    SpreadsheetApp.flush(); 
    return { success: true, ticketId: finalTicketId };
}

/**
 * อัปเดตชื่อผู้ขายย้อนหลังเมื่อจะพิมพ์ตั๋ว
 */
function updateTicketSeller(ticketId, sellerName) {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName(SHEET_CUST);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === ticketId.toString()) {
        sheet.getRange(i + 1, 19).setValue(sellerName); // Column S = 19
        return { success: true };
      }
    }
    return { success: false, error: 'ไม่พบรหัสตั๋วนี้' };
  } catch (e) { return { success: false, error: e.toString() }; }
}



/**
 * ฟังก์ชันสำหรับตั้งค่าเลขตั๋วรายคนเมื่อกดปุ่ม "พิมพ์ตั๋ว"
 */
function setTicketForPrint(ticketId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const eticketSheet = ss.getSheetByName('KRR E-ticket');
    if (eticketSheet) {
      eticketSheet.getRange("H3").setValue(ticketId);
      SpreadsheetApp.flush();
      return { success: true };
    }
    return { success: false, error: 'ไม่พบ Sheet KRR E-ticket' };
  } catch(e) { return { success: false, error: e.toString() }; }
}

// หา row, col ของที่นั่ง seatId ภายใน slot block
function _getSeatCellPos(seatId, startRow, startCol) {
  const match = seatId.match(/^(\d+)([ABCD])$/);
  if (!match) return null;
  const r    = parseInt(match[1]); // 1-9
  const side = match[2];           // A/B/C/D
  
  // แถวที่นั่ง 1-4: startRow+2 -> startRow+5
  // แถวบันได 5-6: startRow+6 -> startRow+7
  // แถวที่นั่ง 7-9: startRow+8 -> startRow+10
  let row = startRow + 2 + (r - 1);

  let col;
  if (side === 'A') col = startCol;
  else if (side === 'B') col = startCol + 1;
  else if (side === 'C') col = startCol + 3;
  else if (side === 'D') col = startCol + 4;
  else return null;
  
  return { row, col };
}

// ════════════════════════════════════════════════════════════
//  CONFIG — ราคาเส้นทาง
// ════════════════════════════════════════════════════════════

function getPrices() {
  const ss    = getSS();
  const sheet = ss.getSheetByName(SHEET_CONFIG);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  return data.slice(1).map(row => ({
    origin:       row[0],
    destination:  row[1],
    fullPrice:    row[2],
    cardPrice:    row[3],
    agentPrice:   row[4],
    inspectorPrice: row[5]
  }));
}

function savePriceRoute(index, data) {
  try {
    const ss    = getSS();
    const sheet = ss.getSheetByName(SHEET_CONFIG);
    const values = [[data.origin, data.destination, data.fullPrice,
                     data.cardPrice, data.agentPrice, data.inspectorPrice]];
    if (parseInt(index) >= 0) {
      sheet.getRange(parseInt(index) + 2, 1, 1, 6).setValues(values);
    } else {
      sheet.appendRow(values[0]);
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function deletePriceRoute(index) {
  try {
    const ss    = getSS();
    const sheet = ss.getSheetByName(SHEET_CONFIG);
    sheet.deleteRow(parseInt(index) + 2);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ════════════════════════════════════════════════════════════
//  MANUAL TICKET — ออกตั๋วฉุกเฉินและเซฟลงฐานข้อมูล
// ════════════════════════════════════════════════════════════

// ออกตั๋วฉุกเฉินและเซฟลงฐานข้อมูล
function processManualTicket(data) {
  try {
    const ss = getSS();
    const custSheet = ss.getSheetByName('ข้อมูลลูกค้า');
    let ticketId = generateTicketId();

    const now = new Date();
    const dateStr = Utilities.formatDate(now, "GMT+7", "dd/MM/yyyy");
    const timestampStr = Utilities.formatDate(now, "GMT+7", "dd/MM/yyyy HH:mm:ss");
    
    const dateObj = new Date(data.date);
    const mm = (dateObj.getMonth() + 1).toString().padStart(2, '0');
    const dd = dateObj.getDate().toString().padStart(2, '0');
    const yy = dateObj.getFullYear();
    const formattedTravelDate = `${dd}/${mm}/${yy}`;

    const isCash  = !!data.isCash;
    const isQR    = !!data.isQR;
    const isAgent = !!data.isAgent;
    const isCard  = !!data.isCard;
    const isInsp  = !!data.isInspector;

    const normalizedFrom = normalizeLocation(data.from);
    const normalizedTo   = normalizeLocation(data.to);
    const normalizedTime = normalizeTime(data.time);

    // Full 24-column structure 
    const rowData = [
      ticketId, data.name, data.phone, dateStr, formattedTravelDate, normalizedFrom, normalizedTo, data.boarding, normalizedTime, data.seatNo,
      data.platform || '7', data.busNo, data.price, isCash, isQR, isAgent, isCard, isInsp, data.issuer, timestampStr, (data.gender || ""), "ACTIVE", (data.remarks || ""), ""
    ];

    const dataA = custSheet.getRange("A:A").getValues();
    let targetRow = dataA.length + 1;
    for (let i = 0; i < dataA.length; i++) {
      if (dataA[i][0] === "") {
        targetRow = i + 1;
        break;
      }
    }
    custSheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);

    let bYear = yy + 543;
    let travelDateForTicket = `${dd}/${mm}/${bYear}`;

    const printData = {
        ticketNo: ticketId,
        name: data.name,
        phone: data.phone,
        destination: data.to,
        boarding: data.boarding,
        price: data.price,
        isCash, isQR, isAgent, isCard, isInspector: isInsp,
        date: travelDateForTicket,
        time: data.time,
        seatId: data.seatNo,
        busNo: data.busNo,
        platform: data.platform || '7',
        origin: data.from,
        issuedBy: data.issuer
    };

    return { success: true, printData: printData };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// ════════════════════════════════════════════════════════════
//  BOOKING HISTORY & STATUS MANAGEMENT
// ════════════════════════════════════════════════════════════

function getBookingHistory(limit = 100) {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName(SHEET_CUST);
    if (!sheet) return { success: false, error: 'ไม่พบ Sheet ข้อมูลลูกค้า' };
    
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return { success: true, history: [] };

    const header = data[0].map(h => h.toString().trim()); // Trim หัวตารางทั้งหมด
    
    const findIdx = (names) => {
      const lowerNames = names.map(n => n.toLowerCase());
      for (let i = 0; i < header.length; i++) {
        if (lowerNames.includes(header[i].toLowerCase())) return i;
      }
      return -1;
    };

    const idxId     = 0; 
    let idxName   = findIdx(['ชื่อผู้โดยสาร', 'ชื่อ', 'Passenger Name', 'Name']);
    if (idxName === -1) idxName = 1; // Fallback to Col B

    let idxPhone  = findIdx(['เบอร์โทร', 'เบอร์โทรศัพท์', 'เบอร์', 'โทร', 'Phone', 'Tel']);
    if (idxPhone === -1) idxPhone = 2; // Fallback to Col C

    let idxTravel = findIdx(['วันที่เดินทาง', 'วันเดินทาง', 'Travel Date', 'Date']);
    if (idxTravel === -1) idxTravel = 4; // Fallback to Col E

    const idxOrigin = findIdx(['ต้นทาง', 'จาก', 'Origin', 'From']);
    const idxDest   = findIdx(['ปลายทาง', 'ถึง', 'Destination', 'To']);
    
    let idxTime   = findIdx(['เวลา', 'Time']);
    if (idxTime === -1) idxTime = 8; // Fallback to Col I

    const idxSeat   = findIdx(['เลขที่นั่ง', 'ที่นั่ง', 'Seat']);
    const idxPrice  = findIdx(['ราคา', 'Price']);
    
    // ค้นหาช่องคนขาย
    let idxIssuer = findIdx(['คนขาย', 'ผู้ทำรายการ', 'ผู้ขาย', 'Issuer', 'เจ้าหน้าที่', 'Seller']);
    if (idxIssuer === -1 && header.length >= 19) idxIssuer = 18; // Column S

    const findLastIdx = (names) => {
      for (let i = header.length - 1; i >= 0; i--) {
        if (names.includes(header[i])) return i;
      }
      return -1;
    };
    let idxStatus = findLastIdx(['BookingStatus', 'สถานะ']);
    if (idxStatus === -1) idxStatus = 21;
    if (idxIssuer === -1 && header.length >= 19) idxIssuer = 18; // Fallback ไปที่ Column S
    // บังคับหาคอลัมน์สถานะ (ช่องที่ 22 หรือ V)

    const history = data.slice(1)
      .filter(row => row[0] && row[0].toString().trim() !== "") 
      .reverse()
      .slice(0, limit)
      .map(row => {
        // ค้นหาสถานะแบบไล่เช็คทุกช่อง (เผื่อเยื้อง)
        let rawStatus = 'ACTIVE';
        
        // 1. ลองเช็คจาก Index ที่ระบุไว้ก่อน (ถ้าเจอ)
        if (idxStatus !== -1 && row[idxStatus]) {
            const v = row[idxStatus].toString().trim().toUpperCase();
            if (v === 'CANCELLED' || v === 'EDITED' || v === 'ACTIVE') rawStatus = v;
        }
        
        // 2. ถ้ายังไม่เจอ ให้ไล่สแกนทุกช่องในแถวนั้น (วิธีนี้ชัวร์ที่สุด)
        if (rawStatus === 'ACTIVE') {
            for (let cell of row) {
                const v = cell.toString().trim().toUpperCase();
                if (v === 'CANCELLED' || v === 'EDITED') {
                    rawStatus = v;
                    break;
                }
            }
        }

        return {
          ticketId:   row[idxId] || '-',
          name:       row[idxName] || '-',
          phone:      row[idxPhone] || '-',
          travelDate: row[idxTravel] || '-',
          origin:     row[idxOrigin] || '-',
          dest:       row[idxDest] || '-',
          time:       row[idxTime] || '-',
          seatId:     row[idxSeat] || '-',
          price:      row[idxPrice] || '0',
          issuedBy:   (idxIssuer !== -1 ? row[idxIssuer] : '-') || '-',
          status:     rawStatus
        };
      });

    return { success: true, history: history };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function cancelBookingById(ticketId) {
  try {
    const ss = getSS();
    const custSheet = ss.getSheetByName(SHEET_CUST);
    
    const ids = custSheet.getRange("A:A").getValues();
    let targetRow = -1;
    for (let i = ids.length - 1; i >= 1; i--) {
      if (ids[i][0].toString() === ticketId.toString()) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow === -1) return { success: false, error: 'ไม่พบรหัสตั๋ว: ' + ticketId };

    // 1. Mark as CANCELLED in Cust Log
    custSheet.getRange(targetRow, 22).setValue("CANCELLED"); // Column V = 22

    // 2. Clear from Seat Plan
    const rowData = custSheet.getRange(targetRow, 1, 1, 10).getDisplayValues()[0];
    const travelDate = rowData[4]; // E: TravelDate (dd/mm/yyyy)
    const stationRaw = rowData[5]; // F: Origin
    const timeLabel  = rowData[8]; // I: Time
    const seatIdRaw  = rowData[9]; // J: Seat

    if (travelDate && seatIdRaw) {
      const dateParts = travelDate.split('/');
      const day = parseInt(dateParts[0]);
      const month = parseInt(dateParts[1]);
      const year = parseInt(dateParts[2]);
      const station = stationRaw.includes('ภูเก็ต') ? 'phuket' : 'hatyai';
      
      const sheetName = getSeatSheetName(month, year);
      const planSheet = ss.getSheetByName(sheetName);
      
      if (planSheet) {
        // Find Slot Index
        let slotIdx = 0;
        const times = (station === 'phuket' ? PHUKET_TIMES : HATYAI_TIMES);
        for (let j = 0; j < times.length; j++) {
          if (normalizeTime(times[j]) === normalizeTime(timeLabel)) {
            slotIdx = j;
            break;
          }
        }

        const startRow = getDayStartRow(day);
        const startCol = getSlotStartCol(station, slotIdx);
        const seatList = seatIdRaw.split(/[\s,]+/).filter(s => s.trim() !== "");

        seatList.forEach(s => {
          const pos = _getSeatCellPos(s, startRow, startCol);
          if (pos) {
            planSheet.getRange(pos.row, pos.col)
                    .clearDataValidations()
                    .setValue(s)
                    .setBackground(COLOR_VACANT)
                    .setFontColor('#94a3b8')
                    .setFontWeight('normal');
          }
        });
      }
    }

    SpreadsheetApp.flush();
    return { success: true, message: 'ยกเลิกการจองและคืนที่นั่งเรียบร้อยแล้ว' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ════════════════════════════════════════════════════════════
//  HELPER FUNCTIONS
// ════════════════════════════════════════════════════════════

function getTrips() {
  try {
    const sheet = getSS().getSheetByName(SHEET_TRIPS);
    if (!sheet) return [];
    const rows    = sheet.getDataRange().getValues();
    const headers = rows[0];
    return rows.slice(1).reverse().map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
  } catch (e) { return []; }
}

function getTicketLog(limit) {
  try {
    const sheet = getSS().getSheetByName(SHEET_LOG);
    if (!sheet) return [];
    const rows    = sheet.getDataRange().getValues();
    const headers = rows[0];
    const data    = rows.slice(1).reverse();
    const result  = limit ? data.slice(0, limit) : data;
    return result.map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
  } catch (e) { return []; }
}

// ════════════════════════════════════════════════════════════
//  UTILS
// ════════════════════════════════════════════════════════════

function storePreviewData(data) {
  CacheService.getUserCache().put('trip_preview', JSON.stringify(data), 600);
  return true;
}

function getPreviewData() {
  const d = CacheService.getUserCache().get('trip_preview');
  return d ? JSON.parse(d) : null;
}

function getConfig() {
  return {
    hatyaiTimes: HATYAI_TIMES,
    phuketTimes: PHUKET_TIMES
  };
}

/**
 * Smart Time Mapping เพื่อให้ผ่านกฎ Data Validation
 * เช่น 09:00 -> 9:00 (เวลาต้นทาง)
 */
function normalizeTime(time) {
  if (!time) return "";
  let t = time.toString().trim();
  // ตัด 0 นำหน้า (เช่น 09:00 -> 9:00)
  t = t.replace(/^0/, '');
  // เติมสร้อยถ้ายังไม่มี
  if (t && !t.includes("(เวลาต้นทาง)")) {
    t += " (เวลาต้นทาง)";
  }
  return t;
}

/**
 * Smart Mapping เพื่อให้ผ่านกฎ Data Validation ใน Sheet ข้อมูลลูกค้า
 * โดยจะเติมภาษาอังกฤษในวงเล็บให้หัวเมืองใหญ่โดยอัตโนมัติ
 */
function normalizeLocation(loc) {
  if (!loc) return "";
  const s = loc.trim();
  const map = {
    "หาดใหญ่": "หาดใหญ่ (Hatyai)",
    "ภูเก็ต":   "ภูเก็ต (Phuket)",
    "ตรัง":     "ตรัง (Trang)",
    "พัทลุง":   "พัทลุง (Phatthalung)",
    "สงขลา":   "สงขลา (Songkhla)",
    "กระบี่":   "กระบี่ (Krabi)",
    "พังงา":    "พังงา (Phangnga)",
    "โคกกลอย": "โคกกลอย (Khok-Kloi)",
  };
  // ถ้ามีวงเล็บอยู่แล้วไม่ต้องทำอะไร
  if (s.includes("(")) return s;
  return map[s] || s;
}

/**
 * เรียกใช้ครั้งเดียว: เพิ่มคอลัมน์ ราคาบัตร และ นายตรวจ ที่ชีท ข้อมูลลูกค้า
 */
function setupNewPaymentColumns() {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName('ข้อมูลลูกค้า');
    if (!sheet) {
      Logger.log('ไม่พบ Sheet ข้อมูลลูกค้า');
      return;
    }

    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const col17Val = headerRow[16] ? headerRow[16].toString() : '';
    
    if (col17Val !== 'ราคาบัตร') {
      sheet.insertColumnsAfter(16, 2);
      sheet.getRange(1, 17).setValue('ราคาบัตร');
      sheet.getRange(1, 18).setValue('นายตรวจ');
    }

    SpreadsheetApp.flush();
  } catch(e) {
    Logger.log('เกิดข้อผิดพลาด: ' + e.toString());
  }
}

function getAdminPin() { return ADMIN_PIN; }

/**
 * ฟังก์ชันพิเศษสำหรับขยายตารางใน Sheet เป็น 10 Slots ต่อฝั่ง
 * ให้รันฟังก์ชันนี้เพียง "ครั้งเดียว" จากหน้า Apps Script Editor
 */
// ฟังก์ชันช่วยแปลงตัวเลขคอลัมน์เป็นตัวอักษร (เช่น 34 -> AH)
function getColumnLetter(colIndex) {
  let letter = '';
  while (colIndex > 0) {
    let temp = (colIndex - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    colIndex = (colIndex - temp - 1) / 26;
  }
  return letter;
}

function setupExtraSlots(targetSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = targetSheet || ss.getActiveSheet(); 
  const sheetName = sheet.getName();
  
  // ตรวจสอบว่ากำลังเปิดชีตผังที่นั่งอยู่หรือไม่ (ถ้าไม่ได้ส่ง targetSheet มา)
  if (!targetSheet && !sheetName.startsWith("ผังที่นั่ง")) {
    Browser.msgBox("❌ ผิดชีต!", "คำสั่งนี้ใช้ได้กับชีตที่มีชื่อขึ้นต้นด้วย 'ผังที่นั่ง' เท่านั้นครับ", Browser.Buttons.OK);
    return;
  }

  // 1. แทรกคอลัมน์ (ถ้ายังไม่เคยแทรก เพื่อดันภูเก็ต)
  const maxCols = sheet.getMaxColumns();
  if (maxCols < 110) {
    sheet.insertColumnsAfter(37, 24); 
  }

  // 2. บังคับก๊อปปี้ "ความกว้างคอลัมน์" จาก Slot 6 (AF-AK) ทับของเดิมเสมอ
  for (let s = 6; s < 10; s++) {
    const targetBase = 2 + (s * 6);
    for (let i = 0; i < 6; i++) {
      const sourceWidth = sheet.getColumnWidth(32 + i);
      sheet.setColumnWidth(targetBase + i, sourceWidth);
    }
  }

  // 3. บังคับก๊อปปี้ "เนื้อหาและรูปแบบผังที่นั่ง" ทับของเดิมเสมอ
  for (let d = 1; d <= 31; d++) {
    const startRow = getDayStartRow(d);
    const sourceRange = sheet.getRange(startRow, 32, ROWS_PER_DAY, 6); // แม่แบบ AF-AK (Slot 6)
    
    for (let s = 6; s < 10; s++) {
      const targetCol = 2 + (s * 6);
      sourceRange.copyTo(sheet.getRange(startRow, targetCol));
      
      // ล้างช่องทะเบียนรถและพขร. (บรรทัดบนสุดของ Slot)
      sheet.getRange(startRow, targetCol).setValue("");
      sheet.getRange(startRow, targetCol + 2).setValue("");

      // ระบบทำความสะอาดและกู้คืน Seat ID (1A, 1B, ...)
      for (let r = 1; r <= 9; r++) {
        const rowNum = startRow + OFF_SEATS_START + (r - 1);
        const isStair = (r === 5 || r === 6);
        const colOffsets = isStair ? [3, 4] : [0, 1, 3, 4];
        const chars = ["A", "B", "", "C", "D"];
        
        colOffsets.forEach(cOff => {
          const seatId = r + chars[cOff];
          sheet.getRange(rowNum, targetCol + cOff)
               .setValue(seatId)
               .setBackground(COLOR_VACANT)
               .setFontColor('#94a3b8')
               .setFontWeight('normal')
               .clearDataValidations();
        });
      }
    }
  }

  // 4. ระบบซ่อมแซมสูตรและข้อมูล
  for (let d = 1; d <= 31; d++) {
    const startRow = getDayStartRow(d);
    let baseDriverFormula = sheet.getRange(startRow, 35).getFormula(); 
    if (!baseDriverFormula) {
        const fallbackFormula = sheet.getRange(startRow, 5).getFormula();
        if (fallbackFormula) baseDriverFormula = fallbackFormula.replace(/\$D/g, '$AH');
    }
    
    for (let s = 5; s < 10; s++) {
      const targetCol = 2 + (s * 6);
      const driverIdColLetter = getColumnLetter(targetCol + 2);
      sheet.getRange(startRow + 1, targetCol + 1).setFormula(`=$C${startRow + 1}`);
      if (baseDriverFormula) {
          const newFormula = baseDriverFormula.replace(/\$AH/g, '$' + driverIdColLetter);
          sheet.getRange(startRow, targetCol + 3).setFormula(newFormula);
      }
      sheet.getRange(startRow + 1, targetCol + 4).setValue(HATYAI_TIMES[s]);
      const wrongLabel = sheet.getRange(startRow + 1, targetCol);
      if (wrongLabel.getValue() === "รอบเสริม") wrongLabel.clearContent();
    }
  }
}

/**
 * ฟังก์ชันรับคำสั่งจากหน้าเว็บ Dashboard เพื่อรัน Setup
 */
function runSetupFromWeb(pin, dateString) {
  if (pin !== ADMIN_PIN) return { success: false, error: "รหัส PIN ไม่ถูกต้อง" };
  
  try {
    const [y, m, d] = dateString.split('-').map(Number);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = getSeatSheetName(m, y);
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) return { success: false, error: "ไม่พบ Sheet ชื่อ: " + sheetName + " กรุณาสร้างชีตเดือนนี้ก่อนกดเพิ่มรอบครับ" };
    
    setupExtraSlots(sheet);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * สร้างเมนูพิเศษใน Google Sheets เพื่อให้ผู้ใช้กดรันคำสั่งได้ง่ายๆ ไม่ต้องเข้าโค้ด
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚐 ระบบจัดการรอบเสริม')
    .addItem('🛠️ สร้าง/ซ่อมแซมผังที่นั่งรอบเสริม (หาดใหญ่ 10 รอบ)', 'setupExtraSlots')
    .addItem('🗑️ ลบรอบเสริม 2-5 (ย้ายภูเก็ตกลับมาที่เดิม)', 'removeExtraSlots')
    .addToUi();
}

/**
 * ฟังก์ชันลบรอบเสริมและย้ายภูเก็ตกลับ
 */
function removeExtraSlots() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();
  
  if (!sheetName.startsWith("ผังที่นั่ง")) {
    Browser.msgBox("❌ ผิดชีต!", "คำสั่งนี้ใช้ได้กับชีตที่มีชื่อขึ้นต้นด้วย 'ผังที่นั่ง' เท่านั้นครับ", Browser.Buttons.OK);
    return;
  }

  const confirm = Browser.msgBox("⚠️ ยืนยันการลบ", "คุณต้องการลบรอบเสริม 2-5 และย้ายผังภูเก็ตกลับมาที่เดิมใช่หรือไม่?\\n(ข้อมูลในรอบเสริม 2-5 จะหายไป)", Browser.Buttons.YES_NO);
  if (confirm === "yes") {
    // ลบคอลัมน์ 38 ถึง 61 (24 คอลัมน์)
    if (sheet.getMaxColumns() >= 90) {
      sheet.deleteColumns(38, 24);
      Browser.msgBox("✅ สำเร็จ", "ลบรอบเสริมและย้ายผังภูเก็ตกลับเรียบร้อยแล้วครับ", Browser.Buttons.OK);
    } else {
      Browser.msgBox("ℹ️ แจ้งเตือน", "ชีตนี้ยังไม่ได้เพิ่มรอบเสริมครับ", Browser.Buttons.OK);
    }
  }
}

/**
 * ฟังก์ชันรับคำสั่งลบรอบเสริมจากหน้าเว็บ
 */
function runRemoveFromWeb(pin, dateString) {
  if (pin !== ADMIN_PIN) return { success: false, error: "รหัส PIN ไม่ถูกต้อง" };
  
  try {
    const [y, m, d] = dateString.split('-').map(Number);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = getSeatSheetName(m, y);
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) return { success: false, error: "ไม่พบ Sheet ชื่อ: " + sheetName };
    
    if (sheet.getMaxColumns() < 90) {
      return { success: false, error: "ชีตนี้ยังไม่ได้เพิ่มรอบเสริมครับ (คอลัมน์ไม่ถึง 90)" };
    }

    // ลบคอลัมน์ 38 ถึง 61 (24 คอลัมน์)
    sheet.deleteColumns(38, 24);
    
    // บังคับ Flush เพื่อให้ UI รับรู้การเปลี่ยนแปลงโครงสร้าง
    SpreadsheetApp.flush();
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}