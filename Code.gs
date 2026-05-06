// ============================================================
//  DAILY FLEET MASTER — Code.gs
//  เชื่อมต่อกับ Sheet "ผังที่นั่ง" 
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
  if (station === 'hatyai') {
    return HATYAI_START_COL + (slotIndex * SLOT_WIDTH);
  } else {
    // ภูเก็ตจะเริ่มต่อจากหาดใหญ่รอบสุดท้าย
    // จำนวนรอบหาดใหญ่ (hCount) = (จำนวนคอลัมน์ทั้งหมด - คอลัมน์เริ่มต้น - จำนวนคอลัมน์ภูเก็ต 6 รอบ) / ความกว้าง slot
    const maxCols = sheet ? sheet.getMaxColumns() : 74;
    const hCount = Math.max(6, Math.floor((maxCols - 2 - 36) / 6));
    const phuketStart = 2 + (hCount * 6);
    return phuketStart + (slotIndex * SLOT_WIDTH);
  }
}

// ── doGet ────────────────────────────────────────────────────
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'dashboard';
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

    // ─── Smart Column Detection ──────────────────────────────────────────────
    // ตรวจว่าเซลล์แรกของ Slot เป็น "ป้ายคันที่" หรือ "ทะเบียนรถ"
    // ทะเบียนรถจะขึ้นต้นด้วยตัวเลข เช่น "443-"
    // ป้ายคันที่จะเป็นข้อความภาษาไทย เช่น "คันที่1"
    const firstCell = sheet.getRange(startRow, startCol).getDisplayValue().trim();
    const isLabelSlot = firstCell !== '' && !/^\d/.test(firstCell);
    const slotLabel   = isLabelSlot ? firstCell : '';
    const busOffset   = isLabelSlot ? 1 : 0;

    // BUS_NO: prefix + ID
    const busNoPart1 = sheet.getRange(startRow, startCol + busOffset).getDisplayValue();
    const busNoPart2 = sheet.getRange(startRow, startCol + busOffset + 1).getDisplayValue();
    const busNo = (busNoPart1 + busNoPart2).trim() || '-';
    
    // พขร (DRIVER/PHONE): อยู่ที่ +2 จาก bus prefix
    const driverRaw = sheet.getRange(startRow, startCol + busOffset + 2).getDisplayValue() || '-';
    
    // เวลาออก (TIME): แถว 2 ของ slot, คอลัมน์ +4 เสมอ
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
        const isTimeMatch = (logTime === targetTimeNorm);

        // ค้นหาโดยไม่อิง Origin เพื่อให้ดึงข้อมูลผู้โดยสารที่ขึ้นระหว่างทางได้
        if (isDateMatch && isTimeMatch) {
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

        // ตรวจสอบการหมดเวลาของ LOCKED (ข้ามถ้าเป็นล็อคถาวร PERMANENT)
        if (row[21] === "LOCKED" && row[22] && row[22] !== "PERMANENT") {
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
        // ลบเงื่อนไข logOrigin ออกเพื่อให้ดึงข้อมูลคนที่ขึ้นจากจุดอื่นในเที่ยวรถเดียวกันได้
        const targetTime   = normalizeTime(timeLabel).toUpperCase();

        if (logDateStr.includes(targetDateStr) && logTime.includes(targetTime)) {
          const logSeatsRaw = row[9] ? row[9].toString().split(/[\s,]+/) : [];
          const logSeats = logSeatsRaw.map(s => normalize(s).replace(/^0/, '')).filter(s => s !== '');
          
          logSeats.forEach(logSeat => {
            if (logSeat && !logDataMap[logSeat]) {
              logDataMap[logSeat] = {
                ticketId: row[0]  || '',
                name:     row[1]  || '',
                phone:    row[2]  || '',
                origin:   row[5]  || '',
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
        slotLabel: slotLabel,
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

    const maxCols = sheet.getMaxColumns();
    const hCount = Math.max(6, Math.floor((maxCols - 2 - 36) / 6));
    const loopLimit = (station === 'hatyai') ? Math.min(times.length, hCount) : times.length;

    for (let i = 0; i < loopLimit; i++) {
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
      const timeInSheet = sheet.getRange(startRow + 1, startCol + 4).getDisplayValue();
      
      // ตรวจสอบว่ารอบนี้มีการ "วาดผัง" หรือยัง (เช็คจากที่นั่ง 1A)
      const firstSeatVal = sheet.getRange(startRow + OFF_SEATS_START, startCol).getDisplayValue();
      const hasData = (firstSeatVal !== "");

      results.push({
        slotIndex: i,
        time:      timeInSheet || times[i], // ใช้เวลาจาก Sheet ถ้ามีการพิมพ์ทับ
        busNo:     busNo || '-',
        vacant, prebooked, confirmed,
        total:     vacant + prebooked + confirmed,
        hasData:   hasData // ส่งสถานะไปบอกหน้าเว็บว่ามีข้อมูลผังหรือยัง
      });
    }

    // ตรวจสอบว่ามีการเพิ่มรอบเสริมหาดใหญ่หรือไม่ (ถ้าเกิน 6 รอบปกติถือว่าขยาย)
    const isExpanded = hCount > 6;
    
    return { 
      success: true, 
      slots: results, 
      times, 
      isExpanded: isExpanded,
      hCount: hCount // ส่งจำนวนรอบหาดใหญ่ปัจจุบันกลับไปด้วย
    };
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
        // นำ ticketId ออกจากการบังคับสร้างใหม่ เพื่อให้อัปเดตบรรทัดเดิมได้
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
      const searchId = existingId.toString().trim();
      for (let i = ids.length - 1; i >= 1; i--) {
        if (ids[i][0].toString().trim() === searchId) {
          targetRow = i + 1;
          break;
        }
      }
    }

    if (targetRow === -1 && passengerName && passengerName.trim() !== "") {
      const data = custSheet.getDataRange().getValues();
      const travelDate = `${day}/${month}/${year}`;
      
      // ดึงค่า label มาตรฐานสำหรับเปรียบเทียบ
      let timeLabelComp = time || "";
      if (!timeLabelComp) {
        let raw = sheet.getRange(startRow + 1, startCol + 4).getDisplayValue();
        const baseSlotName = (station === 'phuket' ? PHUKET_TIMES : HATYAI_TIMES)[slotIndex];
        if (!raw || raw === '-') {
          timeLabelComp = normalizeTime(baseSlotName);
        } else {
          if (slotIndex >= 5 && !raw.includes("รอบเสริม")) {
            timeLabelComp = `${baseSlotName} (${raw})`;
          } else {
            timeLabelComp = normalizeTime(raw);
          }
        }
      }
      const originComp = (origin || (station === 'phuket' ? 'ภูเก็ต (Phuket)' : 'หาดใหญ่ (Hatyai)')).toString().trim();

      for (let i = data.length - 1; i >= 1; i--) {
        const r = data[i];
        // เปรียบเทียบ วันเดินทาง + ต้นทาง + รอบเวลา + เลขที่นั่ง
        const isSameTrip = (r[4] === travelDate && r[5].toString().trim() === originComp && r[8].toString().trim() === timeLabelComp);
        const isSameSeat = (r[9].toString().trim() === seatId.toString().trim());
        const isNotCancelled = (r[21] !== "CANCELLED" && r[21] !== "EXPIRED");

        if (isSameTrip && isSameSeat && isNotCancelled) {
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
      let rawTimeLabel = sheet.getRange(startRow + 1, startCol + 4).getDisplayValue();
      const baseSlotName = (station === 'phuket' ? PHUKET_TIMES : HATYAI_TIMES)[slotIndex];
      if (!rawTimeLabel || rawTimeLabel === '-') {
        timeLabel = normalizeTime(baseSlotName);
      } else {
        if (slotIndex >= 5 && !rawTimeLabel.includes("รอบเสริม")) {
          timeLabel = `${baseSlotName} (${rawTimeLabel})`;
        } else {
          timeLabel = normalizeTime(rawTimeLabel);
        }
      }
    }

    const finalRemarks = params.remarks || "";

    const originLabel = origin || (station === 'phuket' ? 'ภูเก็ต (Phuket)' : 'หาดใหญ่ (Hatyai)');
    const normalizedDest = normalizeLocation(destination);

    // หากมีการส่ง ticketId มา (เป็นการแก้ไข) สถานะจะเป็น CONFIRMED เพื่อกลับมาสู่สถานะปกติ
    let finalStatus = "CONFIRMED";
    let expireTimeStr = "";
    
    // หากเป็นการล็อค
    if (params.isLock) {
      finalStatus = "LOCKED";
      // ถ้า lockMinutes > 0 คือล็อคชั่วคราว, ถ้า 0 คือล็อคถาวร (ไม่มีหมดอายุ)
      if (params.lockMinutes && params.lockMinutes > 0) {
        const expireDate = new Date(now.getTime() + params.lockMinutes * 60000);
        expireTimeStr = expireDate.toISOString();
      } else {
        // ล็อคถาวร - ไม่ตั้งเวลาหมดอายุ
        expireTimeStr = "PERMANENT";
      }
    }

    // มีทั้งหมด 24 คอลัมน์ ไม่รวม AgentName แล้ว
    const rowData = [
      finalTicketId, passengerName, phone, dateStr, travelDate, originLabel, normalizedDest, boardingPoint, timeLabel, seatId,
      platform, busNo, price, isCash, isQR, isAgent, isCard, isInsp, issuedBy, timestampStr, (gender || ""), finalStatus, finalRemarks, expireTimeStr
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
  // เติมสร้อยถ้ายังไม่มี (ยกเว้นรอบเสริม)
  if (t && !t.includes("(เวลาต้นทาง)") && !t.includes("รอบเสริม")) {
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

function setupExtraSlots(targetSheet, targetDay) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = targetSheet || ss.getActiveSheet(); 
  const sheetName = sheet.getName();
  
  if (!targetSheet && !sheetName.startsWith("ผังที่นั่ง")) {
    Browser.msgBox("❌ ผิดชีต!", "คำสั่งนี้ใช้ได้กับชีตที่มีชื่อขึ้นต้นด้วย 'ผังที่นั่ง' เท่านั้นครับ", Browser.Buttons.OK);
    return;
  }

  // 1. คำนวณจำนวนรอบหาดใหญ่ปัจจุบัน
  const maxCols = sheet.getMaxColumns();
  let hCount = Math.floor((maxCols - 2 - 36) / 6);
  const startRow = targetDay ? getDayStartRow(targetDay) : null;

  // --- LOGIC SMART ADD: ค้นหา Slot ที่ว่างอยู่ในวันนั้นก่อน ---
  let useSlotIdx = -1;
  let insertNewCols = false;

  if (targetDay) {
    // วนลูปรอบเสริมที่มีอยู่แล้ว (รอบที่ 6 เป็นต้นไป คือ Index 5)
    // หมายเหตุ: โค้ดเก่าคุณเริ่มเช็คที่ Index 5 (รอบที่ 6)
    for (let i = 5; i < hCount; i++) {
      const col = getSlotStartCol('hatyai', i, sheet);
      const seatVal = sheet.getRange(startRow + OFF_SEATS_START, col).getValue();
      if (!seatVal || seatVal === "") {
        useSlotIdx = i;
        break;
      }
    }
  }

  // ถ้าไม่มี Slot ว่างเลย ค่อยเพิ่มคอลัมน์ใหม่
  if (useSlotIdx === -1) {
    if (hCount >= 10) {
      if (!targetSheet) Browser.msgBox("ℹ️ แจ้งเตือน", "รอบเสริมหาดใหญ่เต็มแล้ว (สูงสุด 10 รอบ)", Browser.Buttons.OK);
      return { success: false, error: "รอบเสริมหาดใหญ่เต็มแล้ว (สูงสุด 10 รอบ)" };
    }
    insertNewCols = true;
    useSlotIdx = hCount;
  }

  const insertPos = 2 + (useSlotIdx * 6);
  const sourceBase = 32; // รอบเสริม 1 (AF-AK) เป็นแม่แบบ

  // 2. แทรกคอลัมน์ (เฉพาะกรณีจำเป็น)
  if (insertNewCols) {
    sheet.insertColumnsAfter(insertPos - 1, 6); 
    // ก๊อปปี้ความกว้างคอลัมน์
    for (let i = 0; i < 6; i++) {
      const sourceWidth = sheet.getColumnWidth(sourceBase + i);
      sheet.setColumnWidth(insertPos + i, sourceWidth);
    }
  }

  // 3. วาดผังที่นั่ง "เฉพาะวันที่เลือก"
  if (targetDay) {
    const sourceRange = sheet.getRange(startRow, sourceBase, ROWS_PER_DAY, 6);
    
    // ก๊อปปี้มาวาง
    sourceRange.copyTo(sheet.getRange(startRow, insertPos));
    
    // ล้างข้อมูลเก่า (ทะเบียนรถ/พขร)
    sheet.getRange(startRow, insertPos).setValue("");
    sheet.getRange(startRow, insertPos + 2).setValue("");

    // รีเซ็ตที่นั่งให้ว่าง
    for (let r = 1; r <= 9; r++) {
      const rowNum = startRow + OFF_SEATS_START + (r - 1);
      const isStair = (r === 5 || r === 6);
      const colOffsets = isStair ? [3, 4] : [0, 1, 3, 4];
      const chars = ["A", "B", "", "C", "D"];
      
      colOffsets.forEach(cOff => {
        const seatId = r + chars[cOff];
        sheet.getRange(rowNum, insertPos + cOff)
             .setValue(seatId)
             .setBackground(COLOR_VACANT)
             .setFontColor('#94a3b8')
             .setFontWeight('normal')
             .clearDataValidations();
      });
    }

    // --- จัดการสูตร ---
    const idColLetter = getColumnLetter(insertPos + 2); 
    let baseDriverFormula = sheet.getRange(startRow, sourceBase + 3).getFormula();
    
    if (baseDriverFormula) {
        const sourceIdColLetter = getColumnLetter(sourceBase + 2); 
        const regex = new RegExp('\\$' + sourceIdColLetter, 'g');
        const newFormula = baseDriverFormula.replace(regex, '$' + idColLetter);
        sheet.getRange(startRow, insertPos + 3).setFormula(newFormula);
    }
    
    // ใส่เวลาให้ Slot ใหม่
    sheet.getRange(startRow + 1, insertPos + 4).setValue(HATYAI_TIMES[useSlotIdx] || "รอบเสริม " + (useSlotIdx - 4));
    
    // ใส่สูตรวันที่
    sheet.getRange(startRow + 1, insertPos + 1).setFormula(`=$C${startRow + 1}`);
  }
  return { success: true };
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
    
    setupExtraSlots(sheet, d);
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
    .addItem('➕ เพิ่มรอบเสริม (Smart Add - ระบุวัน)', 'menuSetupExtraSlot')
    .addItem('🗑️ ลบ/ล้างผังรอบเสริม (Smart Delete)', 'menuSmartRemove')
    .addSeparator()
    .addItem('🛠️ จัดการโครงสร้างรอบเสริม (Force Setup 10 รอบ)', 'setupExtraSlots')
    .addItem('❌ ลบคอลัมน์รอบเสริมล่าสุด (Force Delete)', 'removeExtraSlots')
    .addToUi();
}

/**
 * ฟังก์ชันเรียกจากเมนูใน Sheet เพื่อถามวันที่ก่อนเพิ่มรอบ
 */
function menuSetupExtraSlot() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('➕ เพิ่มรอบเสริม', 'กรุณาระบุวันที่ต้องการสร้าง (1-31):', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    const day = parseInt(response.getResponseText());
    if (isNaN(day) || day < 1 || day > 31) {
      ui.alert('❌ วันที่ผิดพลาด', 'กรุณาใส่ตัวเลข 1 ถึง 31 ครับ', ui.ButtonSet.OK);
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    setupExtraSlots(sheet, day);
  }
}

/**
 * ฟังก์ชันลบแบบอัจฉริยะจากเมนูใน Sheet
 */
function menuSmartRemove() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('🗑️ Smart Delete', 'ระบุวันที่ต้องการจัดการ (1-31):', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    const day = parseInt(response.getResponseText());
    if (isNaN(day) || day < 1 || day > 31) {
      ui.alert('❌ วันที่ผิดพลาด', 'กรุณาใส่ตัวเลข 1 ถึง 31 ครับ', ui.ButtonSet.OK);
      return;
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const sheetName = sheet.getName();
    
    // ถามรอบที่จะลบ
    const slotResp = ui.prompt('🗑️ ลบ/ล้างผังรอบเสริม', 'กรุณาระบุลำดับ "รอบเสริมที่เพิ่มมา" ที่ต้องการจัดการ\n(ใส่เลข 1 สำหรับรอบเสริมคันแรกที่เพิ่มมา, 2 สำหรับคันถัดไป...):', ui.ButtonSet.OK_CANCEL);
    if (slotResp.getSelectedButton() == ui.Button.OK) {
      const extraNum = parseInt(slotResp.getResponseText());
      if (isNaN(extraNum) || extraNum < 1) {
        ui.alert('❌ ข้อมูลผิดพลาด', 'กรุณาใส่เฉพาะตัวเลขลำดับรอบครับ (เช่น 1, 2, 3...)');
        return;
      }
      
      const slotIndex = 5 + extraNum; 
      // ดึงปี/เดือนปัจจุบันจากชื่อชีต
      let m = new Date().getMonth() + 1;
      let y = new Date().getFullYear();
      const months = ["", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];
      months.forEach((name, idx) => { if (name && sheetName.includes(name)) m = idx; });
      const yearMatch = sheetName.match(/\d{4}/);
      if (yearMatch) y = parseInt(yearMatch[0]) - 543;

      const dateStr = `${y}-${m}-${day}`;
      const res = smartRemoveSlotLogic(ADMIN_PIN, dateStr, 'hatyai', slotIndex);
      ui.alert(res.success ? '✅ สำเร็จ' : '❌ ผิดพลาด', res.msg || res.error, ui.ButtonSet.OK);
    }
  }
}

/**
 * ฟังก์ชันลบรอบเสริมทั้งหมด (Force Delete All Extra Slots)
 * จะลบคอลัมน์รอบที่ 7 เป็นต้นไปทิ้งทั้งหมด เพื่อให้เหลือแค่ 6 รอบมาตรฐาน
 */
function removeExtraSlots() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();
  
  if (!sheetName.startsWith("ผังที่นั่ง")) {
    Browser.msgBox("❌ ผิดชีต!", "คำสั่งนี้ใช้ได้กับชีตที่มีชื่อขึ้นต้นด้วย 'ผังที่นั่ง' เท่านั้นครับ", Browser.Buttons.OK);
    return;
  }

  // 1. คำนวณจำนวนรอบหาดใหญ่ปัจจุบัน
  const maxCols = sheet.getMaxColumns();
  const hCount = Math.floor((maxCols - 2 - 36) / 6);

  // ถ้ารอบมีแค่ 6 หรือน้อยกว่า แปลว่าไม่มีรอบเสริมให้ลบ
  if (hCount <= 6) {
    Browser.msgBox("ℹ️ แจ้งเตือน", "ไม่มียังรอบเสริม (รอบที่ 7 เป็นต้นไป) ให้ลบครับ", Browser.Buttons.OK);
    return;
  }

  const extraSlotsCount = hCount - 6;
  const confirm = Browser.msgBox("⚠️ ยืนยันการลบทิ้งทั้งหมด", 
    "คุณต้องการลบรอบเสริมทั้งหมดจำนวน " + extraSlotsCount + " รอบทิ้งทันทีใช่หรือไม่?\n" +
    "(รอบที่ 7 เป็นต้นไปจะถูกลบออกทั้งหมด ข้อมูลทุกวันในรอบเหล่านี้จะหายไป)", 
    Browser.Buttons.YES_NO);

  if (confirm === "yes") {
    // ตำแหน่งเริ่มต้นของรอบที่ 7 (Index 6)
    const startCol = getSlotStartCol('hatyai', 6, sheet);
    const numColsToDelete = extraSlotsCount * 6;
    
    if (startCol !== -1) {
      sheet.deleteColumns(startCol, numColsToDelete);
      SpreadsheetApp.flush();
      Browser.msgBox("✅ สำเร็จ", "ลบรอบเสริมทั้งหมดเรียบร้อยแล้วครับ ชีตกลับสู่สถานะ 6 รอบมาตรฐาน", Browser.Buttons.OK);
    }
  }
}

/**
 * ฟังก์ชันหลักสำหรับลบแบบอัจฉริยะ (ใช้ร่วมกันทั้งเว็บและชีต)
 */
function smartRemoveSlotLogic(pin, dateString, station, slotIndex) {
  if (pin !== ADMIN_PIN) return { success: false, error: "รหัส PIN ไม่ถูกต้อง" };
  
  try {
    const [y, m, d] = dateString.split('-').map(Number);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = getSeatSheetName(m, y);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: "ไม่พบ Sheet" };

    const startCol = getSlotStartCol(station, slotIndex, sheet);
    const maxCols = sheet.getMaxColumns();
    const hCount = Math.floor((maxCols - 2 - 36) / 6);
    const phuketStartCol = 2 + (hCount * 6); // จุดเริ่มของภูเก็ต

    // 1. ระบบป้องกันฝั่งภูเก็ต (Hard Boundary)
    if (station === 'hatyai') {
      if (startCol >= phuketStartCol) {
         return { success: false, error: "❌ ผิดพลาด: ตำแหน่งที่ระบุอยู่ในเขตของฝั่งภูเก็ต ระบบได้ทำการระงับการลบเพื่อป้องกันข้อมูลเสียหายครับ" };
      }
      if (slotIndex <= 5) {
         return { success: false, error: "❌ รอบที่ 1-6 ของหาดใหญ่เป็นรอบมาตรฐาน/แม่แบบ ไม่สามารถลบได้ครับ" };
      }
    }

    if (startCol === -1 || startCol >= maxCols) {
       return { success: false, error: "ไม่พบตำแหน่งรอบเสริมที่ระบุครับ (รอบที่ " + (slotIndex - 5) + " อาจยังไม่ได้สร้าง)" };
    }

    // 1. สแกนทั้งเดือน (1-31) เพื่อดูว่ามีวันอื่นที่ "มีผังที่นั่ง" อยู่ไหม
    let usedInOtherDays = false;
    for (let day = 1; day <= 31; day++) {
      if (day === d) continue; 
      
      const row = getDayStartRow(day);
      
      // ดึงค่ามาตรวจสอบแบบละเอียด (Trim และเช็คค่าว่าง)
      const driverVal = String(sheet.getRange(row + 1, startCol).getValue()).trim();
      const plateVal = String(sheet.getRange(row + 3, startCol).getValue()).trim();
      
      // รายการคำที่ไม่ถือว่าเป็นการใช้งาน
      const ignoreValues = ["", "null", "undefined", "false", "FALSE", "ชื่อคนขับ", "ทะเบียน", "0", "0.0"];

      if (driverVal && !ignoreValues.includes(driverVal)) {
        usedInOtherDays = true;
        break;
      }
      if (plateVal && !ignoreValues.includes(plateVal)) {
        usedInOtherDays = true;
        break;
      }
      
      // เช็คที่นั่ง 1A ของวันนั้นๆ
      const seat1ARange = sheet.getRange(row + OFF_SEATS_START, startCol);
      const seat1A = String(seat1ARange.getValue()).trim();
      const seat1ABg = seat1ARange.getBackground().toLowerCase();

      // ถ้าที่นั่งมีการระบายสี (จองแล้ว)
      const isColored = (seat1ABg !== "#ffffff" && seat1ABg !== "white" && seat1ABg !== "rgba(0,0,0,0)" && seat1ABg !== "#1155cc");
      
      // ถ้ามีชื่อคน (ไม่ใช่เลขที่นั่งปกติ และไม่ใช่ค่าใน ignoreValues)
      const hasPassenger = (seat1A && !ignoreValues.includes(seat1A) && seat1A !== "1A" && isNaN(seat1A));

      if (isColored || hasPassenger) {
        usedInOtherDays = true;
        break;
      }
    }

    if (usedInOtherDays) {
      // --- กรณีมีวันอื่นใช้งานอยู่: ล้างผังเฉพาะวันนี้ทิ้ง ---
      const startRow = getDayStartRow(d);
      const targetRange = sheet.getRange(startRow, startCol, ROWS_PER_DAY, SLOT_WIDTH);
      
      targetRange.clearContent();
      targetRange.clearFormat();
      targetRange.setBackground("#1155cc"); // สีน้ำเงินเข้มตามธีม
      targetRange.setBorder(false, false, false, false, false, false); 
      
      SpreadsheetApp.flush();
      return { success: true, msg: "ลบแผนผังของวันที่ " + d + " ออกเรียบร้อยครับ (วันอื่นยังคงอยู่)" };
    } else {
      // --- กรณีไม่มีวันอื่นใช้งานเลย (หรือเหลือแค่ผังเดียวที่เรากำลังลบ): ลบคอลัมน์ทิ้งไปเลย ---
      if (slotIndex > 5) {
        // ลบ 6 คอลัมน์ของ Slot นี้ทิ้ง
        sheet.deleteColumns(startCol, 6);
        SpreadsheetApp.flush();
        return { success: true, msg: "รอบเสริมนี้ไม่มีการใช้งานในวันอื่นแล้ว ระบบจึงลบคอลัมน์ทิ้งเพื่อประหยัดพื้นที่ให้ครับ" };
      } else {
        // ถ้ารอบหลัก 1-6 ให้ล้างข้อมูลแทน
        const startRow = getDayStartRow(d);
        const targetRange = sheet.getRange(startRow, startCol, ROWS_PER_DAY, SLOT_WIDTH);
        targetRange.clearContent();
        targetRange.clearFormat();
        targetRange.setBackground("#1155cc");
        targetRange.setBorder(false, false, false, false, false, false);
        return { success: true, msg: "ล้างข้อมูลผังรอบหลักเรียบร้อย" };
      }
    }
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * ฟังก์ชันรับคำสั่งลบรอบเสริมจากหน้าเว็บ (ปรับเป็น Smart Remove)
 */
function runRemoveFromWeb(pin, dateString, station, slotIndex) {
  if (station === 'phuket') return { success: false, error: "ยังไม่รองรับการลบรอบเสริมฝั่งภูเก็ตผ่านหน้าเว็บครับ" };
  
  // ถ้าไม่ได้ระบุ slotIndex มา ให้ลบรอบสุดท้าย
  if (slotIndex === undefined || slotIndex === null) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const [y, m, d] = dateString.split('-').map(Number);
    const sheet = ss.getSheetByName(getSeatSheetName(m, y));
    if (!sheet) return { success: false, error: "ไม่พบ Sheet" };
    const maxCols = sheet.getMaxColumns();
    const hCount = Math.floor((maxCols - 2 - 36) / 6);
    slotIndex = hCount - 1;
  }
  
  return smartRemoveSlotLogic(pin, dateString, station, slotIndex);
}

function getDateString() {
  const now = new Date();
  return now.getFullYear() + "-" + (now.getMonth() + 1) + "-" + now.getDate();
}
