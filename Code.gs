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
const SHEET_CONFIG    = "Config_ราคา";
const SHEET_TRIPS     = "รายการเที่ยว";
const SHEET_CUST      = "ข้อมูลลูกค้า";
const SHEET_CARGO     = "ข้อมูลฝากของ";
const ADMIN_PIN       = "1234"; // รหัสผ่านหลักและสำหรับแผ่นงาน (Bypass)
const PIN_HATYAI      = "1234"; // รหัสผ่านสำหรับระบบหลังบ้านฝั่งหาดใหญ่
const PIN_PHUKET      = "5678"; // รหัสผ่านสำหรับระบบหลังบ้านฝั่งภูเก็ต

// ── Constants ──────────────────────────────────────────────
const PLAN_HEADER_ROWS = 2;         // แถว Header (Row 1-2) ก่อนเริ่มข้อมูลวัน
const ROWS_PER_DAY     = 12;        // 1 วัน = 12 แถว
const OFF_BUS_INFO     = 0;         // อยู่ Row 3 (startRow)
const OFF_PHONE_TIME   = 1;         // อยู่ Row 4 (startRow + 1)
const OFF_SEATS_START  = 2;         // เริ่มที่นั่ง Row 5 (startRow + 2)
const OFF_DAY_SEP      = 11;        // แถวสีน้ำเงินคั่นวัน

// ── สี ──────────────────────────────────────────────────────
const COLOR_PREBOOK  = "#22c55e";   // เขียว = จองแล้ว (กรอกชื่อในชีทเอง)
const COLOR_CONFIRM  = "#3b82f6";   // ฟ้า   = จ่ายเงินแล้ว (หาดใหญ่)
const COLOR_HATYAI_PENDING = "#93c5fd"; // ฟ้าอ่อน = จองแล้ว (หาดใหญ่ - ยังไม่จ่าย)
const COLOR_PHUKET_PENDING = "#9333ea"; // ม่วง   = จองแล้ว (ภูเก็ต - ยังไม่จ่าย)
const COLOR_PHUKET_PAID    = "#6b21a8"; // ม่วงเข้ม = จ่ายเงินแล้ว (ภูเก็ต)
const COLOR_LOCKED   = "#f59e0b";   // ส้ม   = ล็อคชั่วคราว
const COLOR_LOCKED_PHUKET = "#d8b4fe"; // ม่วงอ่อน = ล็อคโดยภูเก็ต
const COLOR_LOCKED_HATYAI = "#a2c4c9"; // ฟ้าอ่อน/เขียวคราม = ล็อคโดยหาดใหญ่
const COLOR_VACANT   = "#ffffff";   // ขาว   = ว่าง
const COLOR_STAIR    = "#2563eb";   // บันไดกลาง

// ── Column Structure ─────────────────────────────────────────
const SLOT_WIDTH = 6;  // ทุกสถานีกว้าง 6 (A, B, Gap, C, D, Sep)
const PHUKET_SLOT_WIDTH = 6; 

const HATYAI_START_COL = 2;  // Column B
const HATYAI_TIMES     = ["09:00", "12:40", "14:00", "16:00", "22:00", "คันที่ 1", "คันที่ 2", "คันที่ 3", "คันที่ 4", "คันที่ 5"];

const PHUKET_START_COL = 38; 
const PHUKET_TIMES     = ["07:45", "09:45", "11:45", "14:00", "22:30", "คันที่ 1", "คันที่ 2", "คันที่ 3"];

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
    // 🔍 ระบบค้นหาจุดเริ่มภูเก็ตอัตโนมัติ (Dynamic Detection)
    if (!sheet) {
      const ss = getSS();
      const now = new Date();
      const sheetName = getSeatSheetName(now.getMonth() + 1, now.getFullYear());
      sheet = ss.getSheetByName(sheetName);
    }
    if (!sheet) return PHUKET_START_COL + (slotIndex * PHUKET_SLOT_WIDTH); // Fallback

    // ค้นหาคำว่า "ภูเก็ต" หรือ "PHUKET" ในแถวที่ 1-2 เพื่อหาจุดแบ่งเขต
    const row1 = sheet.getRange(1, 1, 1, Math.min(sheet.getMaxColumns(), 100)).getValues()[0];
    const row2 = sheet.getRange(2, 1, 1, Math.min(sheet.getMaxColumns(), 100)).getValues()[0];
    let autoPhuketStart = -1;

    // หาจาก Row 1 ก่อน
    for (let c = 0; c < row1.length; c++) {
      const val = row1[c].toString().toUpperCase();
      if (val.includes("PHUKET") || val.includes("ภูเก็ต")) {
        autoPhuketStart = c + 1;
        break;
      }
    }
    // ถ้าไม่เจอ หาจาก Row 2
    if (autoPhuketStart === -1) {
      for (let c = 0; c < row2.length; c++) {
        const val = row2[c].toString().toUpperCase();
        if (val.includes("PHUKET") || val.includes("ภูเก็ต")) {
          autoPhuketStart = c + 1;
          break;
        }
      }
    }

    const finalStart = (autoPhuketStart !== -1) ? autoPhuketStart : PHUKET_START_COL;
    return finalStart + (slotIndex * PHUKET_SLOT_WIDTH);
  }
}

/**
 * แปลงเลขคอลัมน์เป็นตัวอักษร (เช่น 34 -> AH)
 */
function getColumnLetter(colIndex) {
  let letter = '';
  while (colIndex > 0) {
    let temp = (colIndex - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    colIndex = (colIndex - temp - 1) / 26;
  }
  return letter;
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

    // ─── Optimize: Batch Read the entire Slot Area (Header + 9 seat rows, 6 columns) ───
    const fullSlotRange = sheet.getRange(startRow, startCol, 11, 6);
    const fullValues = fullSlotRange.getDisplayValues();
    const fullBgs = fullSlotRange.getBackgrounds();

    // firstCell is at [0,0]
    const firstCell = fullValues[0][0].trim();
    const isLabelSlot = firstCell !== '' && !/^\d/.test(firstCell);
    const slotLabel   = isLabelSlot ? firstCell : '';
    const busOffset   = isLabelSlot ? 1 : 0;

    // headerBg is at [0,0]
    const headerBg = fullBgs[0][0];
    const isCut = (headerBg === '#ff0000');

    // BUS_NO: prefix at [0, busOffset], ID at [0, busOffset + 1]
    const busNoPart1 = fullValues[0][busOffset];
    const busNoPart2 = fullValues[0][busOffset + 1];
    const busNo = (busNoPart1 + busNoPart2).trim() || '-';
    
    // พขร (DRIVER/PHONE): [0, busOffset + 2]
    const driverRaw = fullValues[0][busOffset + 2] || '-';
    
    // เวลาออก (TIME): [1, 4]
    const timeValue = fullValues[1][4] || '-';

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

    // --- ดึงข้อมูลเสริมจาก 'ข้อมูลลูกค้า' (Source of Truth) ---
    const logDataMap = {}; // ใช้เก็บข้อมูล Metadata สำหรับที่นั่ง
    const custSheet = ss.getSheetByName('ข้อมูลลูกค้า');
    if (custSheet) {
      const lastRowCust = custSheet.getLastRow();
      const startCust = Math.max(1, lastRowCust - 1500); // อ่าน 1500 แถวหลังสุด
      const custData = custSheet.getRange(startCust, 1, (lastRowCust - startCust + 1), 24).getValues(); 
      const now = new Date().getTime();

      for (let i = custData.length - 1; i >= 0; i--) {
        const row = custData[i];
        const actualRowIdx = startCust + i; // แถวจริงใน Sheet
        if (!row[4] || !row[9]) continue; 

        // ตรวจสอบการหมดเวลาของ LOCKED
        if (row[21] === "LOCKED" && row[22] && row[22] !== "PERMANENT") {
          const expireTime = new Date(row[22]).getTime();
          if (now > expireTime) {
            // หมดเวลา! คืนที่นั่งในผังรถ
            const seatIdLoc = row[9].toString().trim();
            const seatPos = _getSeatCellPos(seatIdLoc, startRow, startCol);
            if (seatPos) {
              sheet.getRange(seatPos.row, seatPos.col).setBackground(null).setValue('').setFontWeight('normal');
            }
            // มาร์คสถานะในฐานข้อมูล (ใช้ actualRowIdx ให้ตรงกับแถวจริง)
            custSheet.getRange(actualRowIdx, 22).setValue("EXPIRED");
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
    const seatRelCols = [0, 1, 3, 4]; // Relative to startCol: B, C, E, F

    for (let r = 1; r <= 9; r++) {
      let isStairRow = (r === 5 || r === 6);
      const rowIdxInFull = 2 + (r - 1); // Row 1 of seats is index 2 in fullValues (header is 0,1)

      seatIds.forEach((sChar, i) => {
        const seatId = r + sChar;
        const colIdxInFull = seatRelCols[i];

        if (isStairRow && (sChar === 'A' || sChar === 'B')) {
          seats[seatId] = { status: 'stair', name: '🪜 บันไดกลาง / ทางเดิน' };
        } else {
          const bg    = fullBgs[rowIdxInFull][colIdxInFull].toLowerCase();
          const value = fullValues[rowIdxInFull][colIdxInFull].trim();
          
          let status = 'vacant';
          if (bg === '#3b82f6' || bg === '#4285f4' || bg === '#0000ff') status = 'PAID_HATYAI';
          else if (bg === '#93c5fd' || bg === '#cfe2f3') status = 'WAITING_HATYAI';
          else if (bg === '#22c55e' || bg === '#34a853' || bg === '#b7e1cd' || bg === '#00ff00') status = 'prebooked';
          else if (bg === '#f59e0b' || bg === '#ff9900') status = 'locked';
          else if (bg === '#d8b4fe' || bg === '#ea9999' || bg === '#f4cccc') status = 'locked_phuket';
          else if (bg === '#a2c4c9' || bg === '#d0e0e3') status = 'locked_hatyai';
          else if (bg === '#6b21a8' || bg === '#7e22ce') status = 'PAID_PHUKET';
          else if (bg === '#9333ea' || bg === '#ff0000' || bg === 'red' || bg === '#ff4d4d' || bg === '#e06666') status = 'WAITING_PHUKET';
          
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
        day, month, year, station, slotIndex,
        isCut: isCut,
        platform: (station === 'phuket' ? '18' : '7')
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
    const phuketStart = getSlotStartCol('phuket', 0, sheet);
    const hCount = Math.floor((phuketStart - 2) / 6);
    const pCount = Math.floor((maxCols - phuketStart + 6) / 6); 

    const loopLimit = (station === 'hatyai') ? Math.min(times.length, hCount) : Math.min(times.length, pCount);
    const startRow = getDayStartRow(day);

    // ─── Optimize: Batch Read the entire Day range (11 rows x maxCols) ───
    const dayRange = sheet.getRange(startRow, 1, 11, maxCols);
    const dayBgs = dayRange.getBackgrounds();
    const dayValues = dayRange.getDisplayValues();

    for (let i = 0; i < loopLimit; i++) {
      const colStart = getSlotStartCol(station, i, sheet);
      if (colStart + 5 > maxCols) break;

      const relColStart = colStart - 1; // Array index 0-based

      let vacant = 0, prebooked = 0, confirmed = 0;
      for (let r = 1; r <= 9; r++) {
        const rowIdxInDay = 2 + (r - 1);
        const isStair = (r === 5 || r === 6);
        const relCols = isStair ? [relColStart + 3, relColStart + 4] : [relColStart, relColStart + 1, relColStart + 3, relColStart + 4];
        
        relCols.forEach(cIdx => {
          const bg = dayBgs[rowIdxInDay][cIdx].toLowerCase();
          if (bg === '#3b82f6' || bg === '#4285f4' || bg === '#cfe2f3' || bg === '#0000ff') confirmed++;
          else if (bg === '#22c55e' || bg === '#34a853' || bg === '#b7e1cd' || bg === '#00ff00') prebooked++;
          else vacant++;
        });
      }

      // ดึงข้อมูลส่วนหัวจาก Memory แทน getRange
      const busNo = dayValues[0][relColStart + 1]; // OFF_BUS_INFO (row 0), relative col 1
      const timeInSheet = dayValues[1][relColStart + 4]; // row 1, col 4
      
      const firstSeatVal = dayValues[2][relColStart]; // OFF_SEATS_START (row 2), col 0
      const hasData = (firstSeatVal !== "");

      const slotLabel = (i >= 5) ? ("คันที่ " + (i - 4)) : "";
      results.push({
        slotIndex: i,
        slotLabel: slotLabel,
        time:      timeInSheet || times[i], 
        busNo:     busNo || '-',
        vacant, prebooked, confirmed,
        total:     vacant + prebooked + confirmed,
        hasData:   hasData 
      });
    }

    // ตรวจสอบว่ามีการเพิ่มรอบเสริมหรือไม่ (ถ้าเกิน 6 รอบปกติถือว่าขยาย)
    const currentCount = (station === 'hatyai' ? hCount : pCount);
    const isExpanded = currentCount > 6;
    
    return { 
      success: true, 
      slots: results, 
      times, 
      isExpanded: isExpanded,
      hCount: currentCount // ส่งจำนวนรอบของสาขานั้นๆ กลับไป
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
    if (!ticketId) return { success: false, error: 'กรุณาระบุรหัสตั๋ว' };
    const ss = getSS();
    let sheet = ss.getSheetByName("ข้อมูลลูกค้า");
    if (!sheet) {
      const sheets = ss.getSheets();
      for (let s of sheets) {
        if (s.getName().includes("ข้อมูลลูกค้า") || s.getSheetId().toString() === "1305369588") {
          sheet = s;
          break;
        }
      }
    }

    if (!sheet) return { success: false, error: 'ไม่พบ Sheet ข้อมูลลูกค้า (กรุณาเช็คชื่อชีทหรือ GID)' };

    const data = sheet.getDataRange().getDisplayValues();
    const idList = ticketId.toString().split(',').map(s => s.trim().toLowerCase());
    const allFoundTickets = [];

    // ถ้ามีรหัสเดียว และเป็นตัวเลขล้วน ให้ทำ Fuzzy Match
    if (idList.length === 1 && /^\d+$/.test(idList[0])) {
      const searchNum = idList[0];
      for (let i = data.length - 1; i >= 1; i--) {
        const currentId = data[i][0].toString().trim();
        const numericPart = currentId.match(/\d+$/);
        if (numericPart && Number(numericPart[0]) === Number(searchNum)) {
          allFoundTickets.push(extractTicketRow(data[i]));
        }
      }
    } else {
      // กรณีปกติ หรือหลายรหัส ให้ค้นหาแบบ Exact Match
      for (let i = 1; i < data.length; i++) {
        const currentId = data[i][0].toString().trim().toLowerCase();
        if (idList.includes(currentId)) {
          allFoundTickets.push(extractTicketRow(data[i]));
        }
      }
    }

    if (allFoundTickets.length === 0) {
      return { success: false, error: 'ไม่พบรหัสตั๋ว: ' + ticketId };
    }

    if (allFoundTickets.length > 1) {
      // เรียงลำดับตามที่นั่ง (เพื่อให้พิมพ์ตั๋วเรียงลำดับ A1, A2...)
      allFoundTickets.sort((a, b) => (a.seatId || '').localeCompare(b.seatId || '', undefined, {numeric: true, sensitivity: 'base'}));
      return { success: true, isGroup: true, tickets: allFoundTickets, ticketId: ticketId };
    } else {
      return allFoundTickets[0];
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
      bookingTime: row[19] ? row[19].toString() : "", // ✅ แปลงเป็นข้อความตั้งแต่ตรงนี้เพื่อความชัวร์
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
      
      // สร้าง Ticket IDs เตรียมไว้เลยเพื่อป้องกัน ID ซ้ำกันเมื่อรัน loop เร็วๆ
      let baseIdStr = generateTicketId().replace("online", "");
      let startIdNum = parseInt(baseIdStr, 10);

      // ถ้ามีการส่ง oldTicketId มา (เช่น การแก้จากแถวที่เคยรวมกันอยู่มาเป็นแยกแถว)
      if (params.oldTicketId) {
        cancelBookingById(params.oldTicketId); // ยกเลิกของเก่าทิ้งไปก่อน
      }

      params.bookings.forEach((b, index) => {
        // บังคับกำหนดรหัสตั๋วใหม่แยกแต่ละที่นั่งไปเลย (แยกรหัสตั๋ว)
        b.forcedNewTicketId = "online" + (startIdNum + index).toString().padStart(6, '0');
        const res = _processSingleBooking(ss, b);
        results.push(res);
      });
      const allTicketIds = results.map(r => r.ticketId).join(', ');
      const ticketMap = {};
      results.forEach((res, idx) => {
        if (res.success) {
          const sId = params.bookings[idx].seatId;
          ticketMap[sId] = res.ticketId;
        }
      });

      return { 
        success: true, 
        message: 'บันทึกการจองกลุ่มแยกแถวเรียบร้อย',
        ticketId: allTicketIds,
        ticketMap: ticketMap
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

    // --- ตรวจสอบการตัดวิน (Server-side Protection) ---
    const checkBg = sheet.getRange(startRow, startCol).getBackground();
    if (checkBg === '#ff0000') {
      throw new Error("ขออภัย รอบนี้ถูก 'ตัดวิน' เรียบร้อยแล้ว ไม่สามารถดำเนินการใดๆ ได้ครับ");
    }
    // -----------------------------------------------

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
    }
    
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
          if (slotIndex >= 5 && !raw.includes("รอบเสริม") && !raw.includes("คันที่")) {
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

    if (!finalTicketId) {
      finalTicketId = params.forcedNewTicketId ? params.forcedNewTicketId : generateTicketId();
    }
    
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
        if (slotIndex >= 5 && !rawTimeLabel.includes("รอบเสริม") && !rawTimeLabel.includes("คันที่")) {
          timeLabel = `${baseSlotName} (${rawTimeLabel})`;
        } else {
          timeLabel = normalizeTime(rawTimeLabel);
        }
      }
    }

    const finalRemarks = params.remarks || "";

    const originLabel = origin || (station === 'phuket' ? 'ภูเก็ต (Phuket)' : 'หาดใหญ่ (Hatyai)');
    const normalizedDest = normalizeLocation(destination);

    // ✅ กำหนดสถานะตามฝั่งที่ทำรายการ
    let finalStatus = "CONFIRMED";
    const isOtherBranch = (params.isPhuketBooking === true || params.isPhuketBooking === "true");
    const isPaid = (params.isCash || params.isQR || params.isAgent || params.isCard || params.isInspector);

    let bookingBranch = params.activeStation;
    if (isOtherBranch) {
      bookingBranch = (params.activeStation === 'phuket') ? 'hatyai' : 'phuket';
    }

    // ดึงสถานะเดิมมาเตรียมไว้สำหรับตรวจสอบเงื่อนไข
    let oldStatus = "";
    if (targetRow > 0) {
      oldStatus = custSheet.getRange(targetRow, 22).getValue().toString().trim().toUpperCase();
    }

    if (isOtherBranch) {
      // กรณีจองข้ามฝั่ง (เช่น หาดใหญ่จองให้ภูเก็ต หรือ ภูเก็ตจองให้หาดใหญ่)
      if (params.isLock) {
        finalStatus = (bookingBranch === 'phuket') ? "LOCKED_PHUKET" : "LOCKED_HATYAI";
      } else {
        // ⚠️ กฎเหล็ก: จองหรือแก้ไขข้ามฝั่ง ถ้ากดผ่านฟอร์มนี้ บังคับเป็น WAITING ก่อนเสมอ! 
        // (การเปลี่ยนเป็น PAID_... จะเกิดขึ้นเมื่อกดปุ่ม "ยืนยันชำระเงิน" ที่เรียกฟังก์ชัน confirmPhuketPayment เท่านั้น)
        finalStatus = (bookingBranch === 'phuket') ? "WAITING_PHUKET" : "WAITING_HATYAI";
      }
    } else {
      // ✅ จองปกติ (สาขาตัวเอง) - ยืนยันเป็น CONFIRMED เลยโดยไม่ต้องเช็คการจ่ายเงิน!
      finalStatus = params.isLock ? "LOCKED" : "CONFIRMED";
    }

    // 📝 ตรวจสอบการแก้ไขข้อมูล (Edit Case)
    if (targetRow > 0 && !params.isLock) {
      // เก็บสถานะที่ห้ามถูกทับด้วย EDITED (ล็อค หรือ ข้ามสาขาที่รอจ่ายเงิน)
      const PRESERVE_STATUSES = ["LOCKED", "LOCKED_PHUKET", "LOCKED_HATYAI", "WAITING_HATYAI", "WAITING_PHUKET"];
      
      // ถ้าของเดิมไม่ใช่กลุ่มล็อค/รอจ่าย (คือเคยเป็น CONFIRMED หรือ PAID มาก่อน)
      // ให้เปลี่ยนสถานะเป็น EDITED เพื่อให้รู้ว่ามีการแก้ไขข้อมูล
      if (!PRESERVE_STATUSES.includes(oldStatus) && oldStatus !== "" && oldStatus !== "VACANT") {
        finalStatus = "EDITED";
      }
    }

    // มีทั้งหมด 23 คอลัมน์ ไม่รวม AgentName แล้ว
    const rowData = [
      finalTicketId, passengerName, phone, dateStr, travelDate, originLabel, normalizedDest, boardingPoint, timeLabel, seatId,
      platform, busNo, price, isCash, isQR, isAgent, isCard, isInsp, issuedBy, timestampStr, (gender || ""), finalStatus, finalRemarks
    ];

    if (targetRow > 0) {
      custSheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
    } else {
      const dataA = custSheet.getRange("A:A").getValues();
      targetRow = dataA.length + 1;
      for (let i = 0; i < dataA.length; i++) {
        if (dataA[i][0] === "") { targetRow = i + 1; break; }
      }
      custSheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
    }

    // ✅ อัปเดตสีผังที่นั่งตาม finalStatus ตรงๆ เลยเพื่อป้องกันสีเพี้ยน
    if (passengerName && passengerName.trim() !== "") {
      const startRowLoc = getDayStartRow(day);
      const startColLoc = getSlotStartCol(station, slotIndex);
      
      seatList.forEach(s => {
        const seatPos = _getSeatCellPos(s, startRowLoc, startColLoc);
        if (seatPos) {
          let bgColor = COLOR_CONFIRM;
          let fontColor = '#ffffff';
          let displayText = passengerName || 'ล็อคชั่วคราว';

          switch(finalStatus) {
              case "CONFIRMED":
              case "EDITED":
              case "PAID_HATYAI":
                  bgColor = COLOR_CONFIRM;
                  break;
              case "PAID_PHUKET":
                  bgColor = COLOR_PHUKET_PAID;
                  break;
              case "WAITING_HATYAI":
                  bgColor = COLOR_HATYAI_PENDING;
                  break;
              case "WAITING_PHUKET":
                  bgColor = COLOR_PHUKET_PENDING;
                  break;
              case "LOCKED":
                  bgColor = COLOR_LOCKED;
                  break;
              case "LOCKED_HATYAI":
                  bgColor = COLOR_LOCKED_HATYAI;
                  fontColor = '#1e293b'; 
                  break;
              case "LOCKED_PHUKET":
                  bgColor = COLOR_LOCKED_PHUKET;
                  fontColor = '#1e293b'; 
                  break;
          }

          sheet.getRange(seatPos.row, seatPos.col).clearDataValidations()
            .setBackground(bgColor)
            .setFontColor(fontColor)
            .setFontWeight('bold')
            .setValue(displayText);
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

    const defaultPlatform = (normalizedFrom && normalizedFrom.includes('ภูเก็ต')) ? '18' : '7';
    const rowData = [
      ticketId, data.name, data.phone, dateStr, formattedTravelDate, normalizedFrom, normalizedTo, data.boarding, normalizedTime, data.seatNo,
      data.platform || defaultPlatform, data.busNo, data.price, isCash, isQR, isAgent, isCard, isInsp, data.issuer, timestampStr, (data.gender || ""), "ACTIVE", (data.remarks || "")
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

/**
 * ยืนยันการจ่ายเงินจาก Phuket Booking
 */
function confirmPhuketPayment(params) {
  try {
    const ss = getSS();
    const custSheet = ss.getSheetByName(SHEET_CUST);
    const { ticketId, seatId, day, month, year, station, slotIndex } = params;
    
    // 1. อัปเดตในฐานข้อมูลลูกค้า
    const data = custSheet.getRange("A:A").getValues();
    let targetRow = -1;
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0].toString() === ticketId.toString()) {
        targetRow = i + 1;
        break;
      }
    }
    
    if (targetRow === -1) return { success: false, error: 'ไม่พบรหัสตั๋ว' };
    
    // อ่านสถานะเดิมก่อนเปลี่ยน
    const oldStatus = custSheet.getRange(targetRow, 22).getValue().toString().trim().toUpperCase();
    let newStatus = "CONFIRMED";
    if (oldStatus === "WAITING_PHUKET") newStatus = "PAID_PHUKET";
    else if (oldStatus === "WAITING_HATYAI") newStatus = "PAID_HATYAI";
    
    custSheet.getRange(targetRow, 22).setValue(newStatus); // Column V
    
    // 2. อัปเดตสีในผังที่นั่ง (เปลี่ยนจากแดงเป็นฟ้า)
    const sheetName = getSeatSheetName(month, year);
    const sheet = ss.getSheetByName(sheetName);
    const startRow = getDayStartRow(day);
    const startCol = getSlotStartCol(station, slotIndex);
    
    const seatList = seatId.toString().split(/[\s,]+/).filter(s => s.trim() !== "");
    seatList.forEach(s => {
      const pos = _getSeatCellPos(s, startRow, startCol);
      if (pos) {
        let finalColor = COLOR_CONFIRM;
        if (newStatus === "PAID_PHUKET") finalColor = COLOR_PHUKET_PAID;
        else if (newStatus === "PAID_HATYAI") finalColor = COLOR_CONFIRM;
        
        sheet.getRange(pos.row, pos.col).setBackground(finalColor);
      }
    });
    
    return { success: true };
  } catch (e) {
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
        let rawStatus = 'CONFIRMED';
        
        // 1. ลองเช็คจาก Index ที่ระบุไว้ก่อน (ถ้าเจอ)
        if (idxStatus !== -1 && row[idxStatus]) {
            rawStatus = row[idxStatus].toString().trim().toUpperCase();
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
  if (t && !t.includes("(เวลาต้นทาง)") && !t.includes("รอบเสริม") && !t.includes("คันที่")) {
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

function getAdminPin(station) {
  if (station === 'phuket') {
    return PIN_PHUKET;
  }
  return PIN_HATYAI;
}

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

/**
 * สร้างโครงสร้างรอบเสริม (Extra Slot) โดยอิงจากชีท "ผังที่นั่ง"
 */
function setupExtraSlots(targetSheet, targetDay, station = 'hatyai', timeLabel = '') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = targetSheet || ss.getActiveSheet(); 
  
  const isHatyai = (station === 'hatyai');
  const sWidth = 6; 
  
  // 1. คำนวณขอบเขตและจำนวนรอบปัจจุบัน
  const maxCols = sheet.getMaxColumns();
  const phuketStartInSheet = getSlotStartCol('phuket', 0, sheet);
  const hCount = Math.floor((phuketStartInSheet - 2) / 6);
  const pCount = Math.floor((maxCols - phuketStartInSheet + 6) / 6);
  
  const startRow = targetDay ? getDayStartRow(targetDay) : null;

  // --- LOGIC SMART ADD: ค้นหา Slot ที่ว่างอยู่ในวันนั้นก่อน (ตามแบบที่พี่ส่งมา) ---
  let useSlotIdx = -1;
  let insertNewCols = false;
  const currentCount = isHatyai ? hCount : pCount;

  if (targetDay) {
    // หาดใหญ่เริ่มเช็คที่รอบเสริมคันแรก (Index 5 หรือตามความเหมาะสม)
    // ภูเก็ตเริ่มเช็คตั้งแต่ Index 0
    const checkStart = isHatyai ? 5 : 0; 
    for (let i = checkStart; i < currentCount; i++) {
      const col = getSlotStartCol(station, i, sheet);
      const seatVal = sheet.getRange(startRow + OFF_SEATS_START, col).getValue();
      if (!seatVal || seatVal === "") {
        useSlotIdx = i;
        break;
      }
    }
  }

  // ถ้าไม่มี Slot ว่างเลย (หรือเป็นรอบใหม่เอี่ยม) ให้เพิ่มคอลัมน์ใหม่
  if (useSlotIdx === -1) {
    if (isHatyai && hCount >= 12) return { success: false, error: "รอบเสริมหาดใหญ่เต็มแล้ว" };
    if (!isHatyai && pCount >= 10) return { success: false, error: "รอบเสริมภูเก็ตเต็มแล้ว" };
    
    insertNewCols = true;
    useSlotIdx = currentCount;
  }

  // ตำแหน่งที่จะใช้งาน
  let insertPos = getSlotStartCol(station, useSlotIdx, sheet);
  
  // แม่แบบที่จะก๊อปปี้ (หาดใหญ่ใช้ AF-AK = 32, ภูเก็ตใช้ BP-BU = 68)
  const sourceBase = isHatyai ? 32 : 68;

  // 2. แทรกคอลัมน์ใหม่ (เฉพาะกรณีที่ Slot เดิมไม่พอ)
  if (insertNewCols) {
    if (isHatyai) {
       // แทรกก่อนเริ่มภูเก็ต
       sheet.insertColumnsBefore(phuketStartInSheet, sWidth);
       insertPos = phuketStartInSheet;
       // ขยายหัวตารางหาดใหญ่
       const hyHeader = sheet.getRange(1, 2, 1, phuketStartInSheet - 2 + sWidth);
       hyHeader.merge().setHorizontalAlignment("center").setBackground("#1155cc");

       // ขยายหัวตารางหาดใหญ่ แถวที่ 2 (เพิ่มเข้ามาใหม่)
       const hyHeader2 = sheet.getRange(2, 2, 1, phuketStartInSheet - 2 + sWidth);
       hyHeader2.merge().setHorizontalAlignment("center").setBackground("#1155cc");
    } else {
       // ภูเก็ต: แทรกต่อท้าย
       sheet.insertColumnsAfter(maxCols, sWidth);
       insertPos = maxCols + 1;
       // ขยายหัวตารางภูเก็ต (แถว 1 และ 2)
       const pkHeader = sheet.getRange(1, phuketStartInSheet, 1, maxCols - phuketStartInSheet + 1 + sWidth);
       pkHeader.merge().setHorizontalAlignment("center").setBackground("#6aa84f");
       const pkHeader2 = sheet.getRange(2, phuketStartInSheet, 1, maxCols - phuketStartInSheet + 1 + sWidth);
       pkHeader2.merge().setHorizontalAlignment("center").setBackground("#6aa84f");
    }
    // ปรับความกว้างคอลัมน์ตามแม่แบบ
    for (let i = 0; i < sWidth; i++) {
      sheet.setColumnWidth(insertPos + i, sheet.getColumnWidth(sourceBase + i));
    }
  }

  // 3. จัดการแผนผังที่นั่ง
  const sourceRange = sheet.getRange(3, sourceBase, ROWS_PER_DAY, sWidth);
  const extraNo = (useSlotIdx - 4); // ทั้งสองฝั่งนับเหมือนกัน (รอบ 1-5 คือ 0-4) ดังนั้นรอบ 6 คือคันที่ 1

  // เราจะวนลูปเพื่อจัดการพื้นหลังให้เนียน แต่จะวาดผังจริง "เฉพาะวันที่เลือก" (ตามที่พี่ต้องการ)
  for (let d = 1; d <= 31; d++) {
    const dayRow = getDayStartRow(d);
    
    if (d === targetDay) {
      // 🎨 วาดผังจริงเฉพาะวันที่เลือก
      sourceRange.copyTo(sheet.getRange(dayRow, insertPos));
      
      // ล้างข้อมูลทะเบียน/พขร และใส่ข้อมูลใหม่
      sheet.getRange(dayRow, insertPos).setValue(""); 
      sheet.getRange(dayRow, insertPos + 2).setValue(""); 
      
      const finalLabel = timeLabel || ("คันที่ " + extraNo);
      sheet.getRange(dayRow + 1, insertPos + 4).setValue(finalLabel);
      
      // ใส่สูตรวันที่
      sheet.getRange(dayRow + 1, insertPos + 1).setFormula(`=$C${dayRow + 1}`);

      // ปรับสูตร VLOOKUP พขร.
      const idColLetter = getColumnLetter(insertPos + 2); 
      let baseDriverFormula = sheet.getRange(dayRow, sourceBase + 3).getFormula();
      if (baseDriverFormula) {
          const sourceIdColLetter = getColumnLetter(sourceBase + 2); 
          const regex = new RegExp('\\$' + sourceIdColLetter, 'g');
          const newFormula = baseDriverFormula.replace(regex, '$' + idColLetter);
          sheet.getRange(dayRow, insertPos + 3).setFormula(newFormula);
      }

      _resetSeatDataOnly(sheet, dayRow, insertPos);
      
      // 🔹 บังคับให้แถวคั่นวัน (แถวสุดท้ายของแต่ละวัน) เป็นสีตามฝั่ง เพื่อลบสีส่วนเกิน
      const sepColor = isHatyai ? "#1155cc" : "#6aa84f";
      sheet.getRange(dayRow + 11, insertPos, 1, sWidth).setBackground(sepColor);

    } else {
      // 🟦 วันอื่นๆ: ถ้าเป็น Slot ที่เพิ่งแทรกใหม่ ให้ล้างเป็นสีพื้นหลัง (น้ำเงิน/เขียว) ให้เนียน
      if (insertNewCols) {
        const bgRange = sheet.getRange(dayRow, insertPos, ROWS_PER_DAY, sWidth);
        const inactiveColor = isHatyai ? "#1155cc" : "#6aa84f"; 
        bgRange.clearContent().clearFormat().setBackground(inactiveColor).setBorder(false, false, false, false, false, false);
        
        // 🔹 รักษาแถวคั่นวันให้เป็นสีตามฝั่ง
        sheet.getRange(dayRow + 11, insertPos, 1, sWidth).setBackground(inactiveColor);
      }
    }
  }

  SpreadsheetApp.flush();
  return { success: true };
}

/**
 * รีเซ็ตเฉพาะค่าตัวอักษรในที่นั่ง (1A, 1B...) โดยคงค่าสีและ Validation (Dropdown) ไว้
 */
function _resetSeatDataOnly(sheet, startRow, insertPos) {
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
           .setFontColor('#1e293b')
           // ลบ .clearDataValidations() ออกเพื่อให้ "Filter/Dropdown" ยังอยู่
    });
  }
}

/**
 * ฟังก์ชันรับคำสั่งจากหน้าเว็บ Dashboard เพื่อรัน Setup
 */
function runSetupFromWeb(pin, dateString, station = 'hatyai', timeLabel = '') {
  const correctPin = (station === 'phuket') ? PIN_PHUKET : PIN_HATYAI;
  if (pin !== correctPin && pin !== ADMIN_PIN) return { success: false, error: "รหัส PIN ไม่ถูกต้อง" };
  
  try {
    const [y, m, d] = dateString.split('-').map(Number);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = getSeatSheetName(m, y);
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) return { success: false, error: "ไม่พบ Sheet ชื่อ: " + sheetName + " กรุณาสร้างชีตเดือนนี้ก่อนกดเพิ่มรอบครับ" };
    
    setupExtraSlots(sheet, d, station, timeLabel);
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
    .addSubMenu(ui.createMenu('📍 ฝั่งหาดใหญ่ (HY)')
      .addItem('➕ เพิ่มรอบเสริม (HY)', 'menuSetupExtraSlotHatyai')
      .addItem('🗑️ ล้างผังรอบเสริม (Smart Delete HY)', 'menuSmartRemoveHatyai')
      .addSeparator()
      .addItem('❌ ลบคอลัมน์รอบเสริมล่าสุด (Force Delete HY)', 'removeExtraSlotsHatyai'))
    .addSubMenu(ui.createMenu('📍 ฝั่งภูเก็ต (PK)')
      .addItem('➕ เพิ่มรอบเสริม (PK)', 'menuSetupExtraSlotPhuket')
      .addItem('🗑️ ล้างผังรอบเสริม (Smart Delete PK)', 'menuSmartRemovePhuket')
      .addSeparator()
      .addItem('❌ ลบคอลัมน์รอบเสริมล่าสุด (Force Delete PK)', 'removeExtraSlotsPhuket'))
    .addToUi();
}

// ── Wrapper Functions ─────────────────────────────────────────
function menuSetupExtraSlotHatyai() { menuSetupExtraSlot('hatyai'); }
function menuSetupExtraSlotPhuket() { menuSetupExtraSlot('phuket'); }

function menuSmartRemoveHatyai() { menuSmartRemove('hatyai'); }
function menuSmartRemovePhuket() { menuSmartRemove('phuket'); }

function removeExtraSlotsHatyai() { removeExtraSlotsLogic('hatyai'); }
function removeExtraSlotsPhuket() { removeExtraSlotsLogic('phuket'); }
// ─────────────────────────────────────────────────────────────

function menuSetupExtraSlot(station = 'hatyai') {
  const ui = SpreadsheetApp.getUi();
  const stationName = (station === 'hatyai') ? "หาดใหญ่" : "ภูเก็ต";
  const response = ui.prompt('➕ เพิ่มรอบเสริม (' + stationName + ')', 'กรุณาระบุวันที่ต้องการสร้าง (1-31):', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    const day = parseInt(response.getResponseText());
    if (isNaN(day) || day < 1 || day > 31) {
      ui.alert('❌ วันที่ผิดพลาด', 'กรุณาใส่ตัวเลข 1 ถึง 31 ครับ', ui.ButtonSet.OK);
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const res = setupExtraSlots(sheet, day, station);
    if (res && res.success === false) {
       ui.alert('❌ ผิดพลาด', res.error, ui.ButtonSet.OK);
    }
  }
}

/**
 * ฟังก์ชันลบแบบอัจฉริยะจากเมนูใน Sheet
 */
function menuSmartRemove(station = 'hatyai') {
  const ui = SpreadsheetApp.getUi();
  const stationName = (station === 'hatyai') ? "หาดใหญ่" : "ภูเก็ต";
  const response = ui.prompt('🗑️ Smart Delete (' + stationName + ')', 'ระบุวันที่ต้องการจัดการ (1-31):', ui.ButtonSet.OK_CANCEL);
  
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
    const slotResp = ui.prompt('🗑️ ลบ/ล้างผังรอบเสริม', 'กรุณาระบุลำดับ "คันที่เพิ่มมา" ที่ต้องการจัดการ\n(ใส่เลข 1 สำหรับคันที่ 1 ที่เพิ่มมา, 2 สำหรับคันถัดไป...):', ui.ButtonSet.OK_CANCEL);
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
      const res = smartRemoveSlotLogic(ADMIN_PIN, dateStr, station, slotIndex);
      ui.alert(res.success ? '✅ สำเร็จ' : '❌ ผิดพลาด', res.msg || res.error, ui.ButtonSet.OK);
    }
  }
}

/**
 * ฟังก์ชันหลักสำหรับลบแบบคอลัมน์รอบเสริมล่าสุด (Force Delete) แยกฝั่ง
 */
function removeExtraSlotsLogic(station) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();
  const isHatyai = (station === 'hatyai');
  const stationName = isHatyai ? "หาดใหญ่" : "ภูเก็ต";
  
  if (!sheetName.startsWith("ผังที่นั่ง")) {
    Browser.msgBox("❌ ผิดชีต!", "คำสั่งนี้ใช้ได้กับชีตที่มีชื่อขึ้นต้นด้วย 'ผังที่นั่ง' เท่านั้นครับ", Browser.Buttons.OK);
    return;
  }

  // 1. คำนวณจำนวนรอบของฝั่งที่เลือก
  const phuketStart = getSlotStartCol('phuket', 0, sheet);
  const maxCols = sheet.getMaxColumns();
  let currentCount = 0;
  
  if (isHatyai) {
    currentCount = Math.floor((phuketStart - 2) / 6);
  } else {
    currentCount = Math.floor((maxCols - phuketStart + 6) / 6);
  }

  // รอบปกติรวมแม่แบบแล้ว ทั้งหาดใหญ่และภูเก็ตมี 6 คอลัมน์ (Index 0-5)
  const standardLimit = 6;

  if (currentCount <= standardLimit) {
    Browser.msgBox("ℹ️ แจ้งเตือน", "ไม่มียังรอบเสริมของ" + stationName + "ให้ลบครับ", Browser.Buttons.OK);
    return;
  }

  const extraSlotsCount = currentCount - standardLimit;
  const confirm = Browser.msgBox("⚠️ ยืนยันการลบรอบเสริม" + stationName, 
    "คุณต้องการลบรอบเสริม" + stationName + " ล่าสุด 1 รอบทิ้งทันทีใช่หรือไม่?\n" +
    "(ข้อมูลทุกวันในรอบเสริมนี้จะหายไป)", 
    Browser.Buttons.YES_NO);

  if (confirm === "yes") {
    // ลบรอบสุดท้ายของฝั่งนั้น
    const deleteIdx = currentCount - 1;
    const startCol = getSlotStartCol(station, deleteIdx, sheet);
    
    if (startCol !== -1) {
      sheet.deleteColumns(startCol, 6);
      SpreadsheetApp.flush();
      Browser.msgBox("✅ สำเร็จ", "ลบรอบเสริม" + stationName + "ล่าสุดเรียบร้อยแล้วครับ", Browser.Buttons.OK);
    }
  }
}

/**
 * ฟังก์ชันหลักสำหรับลบแบบอัจฉริยะ (ใช้ร่วมกันทั้งเว็บและชีต)
 */
function smartRemoveSlotLogic(pin, dateString, station, slotIndex) {
  const correctPin = (station === 'phuket') ? PIN_PHUKET : PIN_HATYAI;
  if (pin !== correctPin && pin !== ADMIN_PIN) return { success: false, error: "รหัส PIN ไม่ถูกต้อง" };
  
  try {
    const [y, m, d] = dateString.split('-').map(Number);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = getSeatSheetName(m, y);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: "ไม่พบ Sheet" };

    const startCol = getSlotStartCol(station, slotIndex, sheet);
    const maxCols = sheet.getMaxColumns();
    const phuketStartCol = getSlotStartCol('phuket', 0, sheet); 

    // 1. ระบบป้องกัน (Hard Boundary)
    if (station === 'hatyai') {
      if (startCol >= phuketStartCol) {
         return { success: false, error: "❌ ผิดพลาด: ตำแหน่งที่ระบุอยู่ในเขตของฝั่งภูเก็ต" };
      }
      if (slotIndex <= 5) {
         return { success: false, error: "❌ รอบที่ 1-6 ของหาดใหญ่เป็นรอบมาตรฐาน ไม่สามารถลบได้ครับ" };
      }
    } else {
      // ภูเก็ต มีรอบปกติ 5 รอบ (Index 0-4) + แม่แบบ (Index 5)
      if (slotIndex <= 5) {
         return { success: false, error: "❌ รอบที่ 1-5 และแม่แบบของภูเก็ต เป็นคอลัมน์ระบบ ไม่สามารถลบได้ครับ" };
      }
    }

    if (startCol === -1 || startCol >= maxCols) {
       return { success: false, error: "ไม่พบตำแหน่งรอบที่ระบุครับ (คันที่ " + (slotIndex - 5 + 1) + " อาจยังไม่ได้สร้าง)" };
    }

    // 1. สแกนทั้งเดือน (1-31) เพื่อดูว่ามีวันอื่นที่ "มีผังที่นั่ง" อยู่ไหม
    let usedInOtherDays = false;
    for (let day = 1; day <= 31; day++) {
      if (day === d) continue; 
      
      const row = getDayStartRow(day);
      
      let isUsed = false;
      
      // 1. ตรวจสอบ ทะเบียน/คนขับ แยกตามฝั่ง
      let pPlate = "", pDriver = "";
      if (station === 'hatyai') {
          pPlate = String(sheet.getRange(row, startCol + 1).getValue()).trim();
          pDriver = String(sheet.getRange(row + 1, startCol + 1).getValue()).trim();
      } else {
          // ภูเก็ต
          pPlate = String(sheet.getRange(row, startCol).getValue()).trim();
          pDriver = String(sheet.getRange(row, startCol + 1).getValue()).trim();
      }
      
      const ignoreVals = ["", "null", "undefined", "false", "FALSE", "ชื่อคนขับ", "ทะเบียน", "0", "0.0"];
      if (pPlate && !ignoreVals.includes(pPlate)) isUsed = true;
      if (pDriver && !ignoreVals.includes(pDriver)) isUsed = true;

      // 2. ตรวจสอบว่ามีการจองที่นั่งไหม (ดูสีและชื่อผู้โดยสารทั้ง 9 แถว)
      if (!isUsed) {
          const seatBg = sheet.getRange(row + OFF_SEATS_START, startCol, 9, 4).getBackgrounds();
          const seatVals = sheet.getRange(row + OFF_SEATS_START, startCol, 9, 4).getValues();
          for (let r = 0; r < 9; r++) {
              for (let c = 0; c < 4; c++) {
                  if (c === 2) continue; // ข้ามช่องทางเดิน (Gap)
                  const bg = seatBg[r][c].toLowerCase();
                  const val = String(seatVals[r][c]).trim();
                  
                  // ถ้าระบายสี (ไม่ใช่สีขาว หรือ สีพื้นหลัง)
                  if (bg !== "#ffffff" && bg !== "white" && bg !== "rgba(0,0,0,0)" && bg !== "#1155cc" && bg !== "#6aa84f") {
                      isUsed = true;
                      break;
                  }
                  
                  // ถ้าชื่อไม่ใช่หมายเลขที่นั่ง
                  let seatColLetter = "";
                  if (c === 0) seatColLetter = "A";
                  if (c === 1) seatColLetter = "B";
                  if (c === 3) seatColLetter = "D"; // คอลัมน์ที่ 4 (Index 3) คือ D สำหรับภูเก็ต
                  
                  // สำหรับหาดใหญ่ ช่องสุดท้ายคือ C หรือ D ขึ้นอยู่กับเลย์เอาต์ แต่เช็คแค่ว่ามันไม่ใช่เลข+ตัวอักษรก็พอ
                  if (val && !val.match(/^[1-9][A-D]$/) && !ignoreVals.includes(val)) {
                      isUsed = true;
                      break;
                  }
              }
              if (isUsed) break;
          }
      }

      if (isUsed) {
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
      const inactiveColor = (station === 'hatyai') ? "#1155cc" : "#6aa84f";
      targetRange.setBackground(inactiveColor); 
      targetRange.setBorder(false, false, false, false, false, false); 
      // คืนสีแถวคั่นวันให้เนียนไปกับพื้นหลังตามฝั่ง
      sheet.getRange(startRow + 11, startCol, 1, 6).setBackground(inactiveColor);      
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
        const inactiveColor = (station === 'hatyai') ? "#1155cc" : "#6aa84f";
        targetRange.setBackground(inactiveColor);
        targetRange.setBorder(false, false, false, false, false, false);
        sheet.getRange(startRow + 11, startCol, 1, 6).setBackground(inactiveColor);
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
  // เปิดให้รองรับทั้ง หาดใหญ่ และ ภูเก็ต
  
  // ถ้าไม่ได้ระบุ slotIndex มา ให้ลบรอบสุดท้าย
  if (slotIndex === undefined || slotIndex === null) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const [y, m, d] = dateString.split('-').map(Number);
    const phuketStart = getSlotStartCol('phuket', 0, sheet);
    const maxCols = sheet.getMaxColumns();
    const currentCount = (station === 'hatyai') ? Math.floor((phuketStart - 2) / 6) : Math.floor((maxCols - phuketStart + 6) / 6);
    slotIndex = currentCount - 1;
  }
  
  return smartRemoveSlotLogic(pin, dateString, station, slotIndex);
}

function getDateString() {
  const now = new Date();
  return now.getFullYear() + "-" + (now.getMonth() + 1) + "-" + now.getDate();
}
function toggleCutWinStatus(params) {
  try {
    const { station, slotIndex, day, month, year, action } = params;
    const sheetName = getSeatSheetName(month, year);
    const sheet = getSS().getSheetByName(sheetName);
    if (!sheet) throw new Error("ไม่พบแผ่นงานผังที่นั่ง");

    const startRow = getDayStartRow(day);
    const startCol = getSlotStartCol(station, slotIndex, sheet);
    const isHatyai = (station === 'hatyai');
    const sWidth   = 5; // ระบายสีแค่ 5 คอลัมน์ (เว้นช่องคั่นคอลัมน์ที่ 6 ไว้)

    if (action === 'cut') {
      // 🔴 กรณี "ตัดวิน" -> ระบายสีแดงทับแผงหัวตาราง 2 แถว x 5 ช่อง
      sheet.getRange(startRow, startCol, 2, sWidth)
           .setBackground("#ff0000")
           .setFontColor("#ffffff")
           .setFontWeight("bold");
    } else {
      // 🟢 กรณี "ปลดล็อก" -> ก๊อปปี้รูปแบบดั้งเดิมมาจาก "วันที่ 1"
      const templateRow = getDayStartRow(1);
      const sourceRange = sheet.getRange(templateRow, startCol, 2, sWidth);
      const targetRange = sheet.getRange(startRow, startCol, 2, sWidth);
      
      // ก๊อปปี้เฉพาะรูปแบบ (สี, เส้นขอบ, ฟอนต์)
      sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT);
      
      // ตรวจสอบความปลอดภัย: ถ้าก๊อปมาแล้วเป็นสีแดง หรือเป็นสีพื้นหลังว่างๆ (#1155cc)
      const checkBg = targetRange.getBackground().toLowerCase();
      if (checkBg === "#ff0000" || checkBg === "#1155cc" || checkBg === "#ffffff" || checkBg === "white") {
          const isExtra = (slotIndex >= 5);
          if (isExtra) {
              if (isHatyai) {
                  // --- คืนสีรอบเสริมหาดใหญ่ (ดำ-เหลือง) ---
                  targetRange.getCell(1,1).setBackground("#000000").setFontColor("#00ffff");
                  targetRange.getCell(1,2).setBackground("#ffff00").setFontColor("#000000");
                  targetRange.getCell(1,3).setBackground("#ffff00").setFontColor("#000000"); 
                  targetRange.getCell(1,4).setBackground("#c9daf8").setFontColor("#000000");
                  targetRange.getCell(1,5).setBackground("#c9daf8").setFontColor("#000000"); 
                  sheet.getRange(startRow + 1, startCol, 1, 5).setBackground("#6d9eeb").setFontColor("#000000");
                  sheet.getRange(startRow + 1, startCol + 4).setBackground("#ffffff").setFontColor("#000000");
              } else {
                  // --- คืนสีรอบเสริมภูเก็ต (เขียวตามผังปกติ เพื่อความเนียน) ---
                  targetRange.getCell(1,1).setBackground("#6aa84f").setFontColor("#000000");
                  targetRange.getCell(1,2).setBackground("#b6d7a8").setFontColor("#000000");
                  targetRange.getCell(1,3).setBackground("#ffffff").setFontColor("#000000");
                  targetRange.getCell(1,4).setBackground("#d9ead3").setFontColor("#000000");
                  targetRange.getCell(1,5).setBackground("#d9ead3").setFontColor("#000000");
                  let r2 = sheet.getRange(startRow + 1, startCol, 1, 5);
                  r2.getCell(1,1).setBackground("#93c47d").setFontColor("#000000");
                  r2.getCell(1,2).setBackground("#93c47d").setFontColor("#000000");
                  r2.getCell(1,3).setBackground("#ffffff").setFontColor("#000000");
                  r2.getCell(1,4).setBackground("#d9ead3").setFontColor("#000000");
                  r2.getCell(1,5).setBackground("#d9ead3").setFontColor("#000000");
              }
          } else {
              // --- คืนสีรอบปกติ (แยกสาขาและแยกสีรายคอลัมน์) ---
              if (isHatyai) {
                  // --- หาดใหญ่ ---
                  targetRange.getCell(1,1).setBackground("#3c78d8").setFontColor("#000000");
                  targetRange.getCell(1,2).setBackground("#3c78d8").setFontColor("#000000");
                  targetRange.getCell(1,3).setBackground("#c9daf8").setFontColor("#000000"); // พขร
                  targetRange.getCell(1,4).setBackground("#c9daf8").setFontColor("#000000");
                  targetRange.getCell(1,5).setBackground("#c9daf8").setFontColor("#000000");
                  
                  let r2 = sheet.getRange(startRow + 1, startCol, 1, 5);
                  r2.getCell(1,1).setBackground("#6d9eeb").setFontColor("#000000");
                  r2.getCell(1,2).setBackground("#6d9eeb").setFontColor("#000000");
                  r2.getCell(1,3).setBackground("#ffffff").setFontColor("#000000");
                  r2.getCell(1,4).setBackground("#ffffff").setFontColor("#000000");
                  r2.getCell(1,5).setBackground("#ffffff").setFontColor("#000000");
              } else {
                  // --- ภูเก็ต ---
                  targetRange.getCell(1,1).setBackground("#6aa84f").setFontColor("#000000"); // ทะเบียน
                  targetRange.getCell(1,2).setBackground("#b6d7a8").setFontColor("#000000"); // พขร (ภูเก็ตอยู่ช่อง 2)
                  targetRange.getCell(1,3).setBackground("#ffffff").setFontColor("#000000"); // ช่องว่าง (Gap)
                  targetRange.getCell(1,4).setBackground("#d9ead3").setFontColor("#000000"); // ป้าย "เวลา"
                  targetRange.getCell(1,5).setBackground("#d9ead3").setFontColor("#000000");
                  
                  let r2 = sheet.getRange(startRow + 1, startCol, 1, 5);
                  r2.getCell(1,1).setBackground("#93c47d").setFontColor("#000000"); // ป้าย "วันที่"
                  r2.getCell(1,2).setBackground("#93c47d").setFontColor("#000000"); // วันที่จริง
                  r2.getCell(1,3).setBackground("#ffffff").setFontColor("#000000"); // Gap
                  r2.getCell(1,4).setBackground("#d9ead3").setFontColor("#000000"); // ป้าย "เวลา"
                  r2.getCell(1,5).setBackground("#d9ead3").setFontColor("#000000"); // เวลาจริง
              }
          }
      }
    }


    return { success: true, isCut: (action === 'cut') };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * ดึงรายการราคาเส้นทางทั้งหมดจาก Sheet Config_ราคา
 * Frontend เรียกผ่าน google.script.run.getPrices()
 */
function getPrices() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Config_ราคา");
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    const result = [];
    // สมมติ: Row 1 = Header, แต่ละแถว = [origin, destination, fullPrice, discountPrice, ...]
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0] && !row[1]) continue; // ข้ามแถวว่าง
      result.push({
        origin:        row[0] ? row[0].toString().trim() : '',
        destination:   row[1] ? row[1].toString().trim() : '',
        fullPrice:     parseFloat(row[2]) || 0,
        discountPrice: parseFloat(row[3]) || 0,
        notes:         row[4] ? row[4].toString().trim() : ''
      });
    }
    return result;
  } catch (e) {
    console.error('getPrices error: ' + e.toString());
    return [];
  }
}

// ════════════════════════════════════════════════════════════
//  CARGO SYSTEM — ระบบจัดส่งพัสดุและสัตว์เลี้ยง (ฝากของหลังรถ)
// ════════════════════════════════════════════════════════════

function normalizeDate(dateInput) {
  if (!dateInput) return "";
  if (dateInput instanceof Date) {
    try {
      const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || "GMT+7";
      return Utilities.formatDate(dateInput, tz, "dd/MM/yyyy");
    } catch (e) {
      const dd = dateInput.getDate().toString().padStart(2, '0');
      const mm = (dateInput.getMonth() + 1).toString().padStart(2, '0');
      const yyyy = dateInput.getFullYear();
      return `${dd}/${mm}/${yyyy}`;
    }
  }
  
  let dateStr = dateInput.toString().trim();
  
  // If it's in YYYY-MM-DD format (includes '-')
  if (dateStr.includes("-")) {
    const parts = dateStr.split("-");
    if (parts.length === 3) {
      const yyyy = parts[0];
      const mm = parts[1].padStart(2, '0');
      const dd = parts[2].padStart(2, '0');
      return `${dd}/${mm}/${yyyy}`;
    }
  }
  
  // If it's in DD/MM/YYYY format
  if (dateStr.includes("/")) {
    const parts = dateStr.split("/");
    if (parts.length === 3) {
      const dd = parts[0].padStart(2, '0');
      const mm = parts[1].padStart(2, '0');
      const yyyy = parts[2];
      return `${dd}/${mm}/${yyyy}`;
    }
  }
  
  return dateStr;
}

function migrateCargoSheetIfNeeded(sheet) {
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  
  // หากเป็นชีทว่างเปล่า (ไม่มีข้อมูลเลย) ให้สร้าง Header ทันที
  if (lastRow === 0 || (lastRow === 1 && sheet.getRange("A1").getValue().toString().trim() === "")) {
    const newHeaders = [
      "วันที่บันทึก",
      "วันเดินทาง",
      "รอบเวลา",
      "ต้นทาง",
      "รถจักรยานยนต์",
      "แมว/สัตว์เลี้ยง",
      "สัมภาระอื่นๆ",
      "หมายเหตุ",
      "สล็อต"
    ];
    sheet.getRange(1, 1, 1, 9).setValues([newHeaders]).setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
    SpreadsheetApp.flush();
    return;
  }

  const firstCell = sheet.getRange("A1").getValue().toString().trim();
  if (firstCell === "รหัสพัสดุ" || firstCell === "รหัสสัมภาระ") {
    const maxCols = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, maxCols).getValues()[0].map(h => h.toString().trim());
    
    const idxIssuer = headers.indexOf("ผู้บันทึก") + 1;
    const idxPrice = headers.indexOf("ค่าบริการ") + 1;
    const idxDest = headers.indexOf("ปลายทาง") + 1;
    const idxId = (headers.indexOf("รหัสพัสดุ") !== -1 ? headers.indexOf("รหัสพัสดุ") : headers.indexOf("รหัสสัมภาระ")) + 1;
    
    const deleteIndexes = [idxIssuer, idxPrice, idxDest, idxId].filter(idx => idx > 0).sort((a, b) => b - a);
    deleteIndexes.forEach(idx => sheet.deleteColumn(idx));
    
    sheet.getRange(1, 9).setValue("สล็อต");
    
    const newHeaders = [
      "วันที่บันทึก",
      "วันเดินทาง",
      "รอบเวลา",
      "ต้นทาง",
      "รถจักรยานยนต์",
      "แมว/สัตว์เลี้ยง",
      "สัมภาระอื่นๆ",
      "หมายเหตุ",
      "สล็อต"
    ];
    sheet.getRange(1, 1, 1, 9).setValues([newHeaders]).setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold");
    SpreadsheetApp.flush();
  }
}

function getStandardStationName(name) {
  if (!name) return "";
  const clean = name.toString().replace(/\u00a0/g, " ").trim().toLowerCase();
  if (clean.indexOf("hatyai") !== -1 || clean.indexOf("หาดใหญ่") !== -1) {
    return "หาดใหญ่";
  }
  if (clean.indexOf("phuket") !== -1 || clean.indexOf("ภูเก็ต") !== -1) {
    return "ภูเก็ต";
  }
  return name.toString().replace(/\u00a0/g, " ").trim();
}

function getStandardTime(timeInput) {
  if (!timeInput) return "";
  let t = "";
  if (timeInput instanceof Date) {
    try {
      const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || "GMT+7";
      t = Utilities.formatDate(timeInput, tz, "H:mm");
    } catch (e) {
      const hh = timeInput.getHours();
      const mm = timeInput.getMinutes().toString().padStart(2, '0');
      t = `${hh}:${mm}`;
    }
  } else {
    t = timeInput.toString().replace(/\u00a0/g, " ").trim();
  }
  // นำ " (เวลาต้นทาง)" หรือ " (เวลาต้นทาง)" ออก
  t = t.replace(/\s*\(เวลาต้นทาง\)\s*/gi, "");
  // แปลงจุดเป็นทวิภาค (เช่น 09.00 -> 09:00)
  t = t.replace(/\./g, ":");
  // ตัดเลข 0 นำหน้าในหลักชั่วโมง (เช่น 09:00 -> 9:00)
  t = t.replace(/^0/, "");
  return t;
}

function getCargoList(travelDate, origin, timeSlot) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_CARGO);
    if (!sheet) return [];
    
    migrateCargoSheetIfNeeded(sheet);

    const data = sheet.getDataRange().getValues();
    const results = [];
    if (data.length <= 1) return [];

    const targetTravelDate = normalizeDate(travelDate);
    const targetOrigin = getStandardStationName(origin);
    const targetTimeSlot = getStandardTime(timeSlot);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowDate = normalizeDate(row[1]);
      const rowTime = getStandardTime(row[2]);
      const rowOrigin = getStandardStationName(row[3]);

      if (rowDate === targetTravelDate && rowOrigin === targetOrigin && rowTime === targetTimeSlot) {
        results.push({
          timestamp: row[0] ? row[0].toString() : "",
          travelDate: rowDate,
          timeSlot: rowTime,
          origin: rowOrigin,
          motorcycle: row[4] === true || row[4].toString().toLowerCase() === 'true',
          cat: row[5] === true || row[5].toString().toLowerCase() === 'true',
          otherCargo: row[6] || "",
          remarks: row[7] || "",
          slotId: row[8] || ""
        });
      }
    }
    // ลบแผ่นงาน DEBUG_CARGO ออกหากระบบทำงานถูกต้องแล้ว เพื่อความสะอาดตาของชีท
    let debugSheet = ss.getSheetByName("DEBUG_CARGO");
    if (debugSheet) {
      ss.deleteSheet(debugSheet);
    }
    return results;
  } catch (err) {
    console.error("getCargoList error: " + err.toString());
    return [];
  }
}

function saveCargoBooking(params) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_CARGO);
    
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_CARGO);
      sheet.appendRow([
        "วันที่บันทึก",
        "วันเดินทาง",
        "รอบเวลา",
        "ต้นทาง",
        "รถจักรยานยนต์",
        "แมว/สัตว์เลี้ยง",
        "สัมภาระอื่นๆ",
        "หมายเหตุ",
        "สล็อต"
      ]);
      sheet.getRange("A1:I1").setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold");
    } else {
      migrateCargoSheetIfNeeded(sheet);
    }

    const travelDate = normalizeDate(params.travelDate);
    const timeSlot = params.timeSlot;
    const origin = getStandardStationName(params.origin);
    const slotId = params.slotId || "";

    const existingList = getCargoList(travelDate, origin, timeSlot);
    
    let existingRowIndex = -1;
    const dataRangeValues = sheet.getDataRange().getValues();
    for (let i = 1; i < dataRangeValues.length; i++) {
      const row = dataRangeValues[i];
      const rDate = normalizeDate(row[1]);
      const rTime = getStandardTime(row[2]);
      const rOrigin = getStandardStationName(row[3]);
      const rSlot = row[8] ? row[8].toString().trim() : "";
      if (rDate === travelDate && rOrigin === origin && rTime === getStandardTime(timeSlot) && rSlot === slotId) {
        existingRowIndex = i + 1;
        break;
      }
    }

    if (existingRowIndex !== -1) {
      sheet.getRange(existingRowIndex, 5).setValue(params.motorcycle ? true : false);
      sheet.getRange(existingRowIndex, 6).setValue(params.cat ? true : false);
      sheet.getRange(existingRowIndex, 7).setValue(params.otherCargo || "");
      sheet.getRange(existingRowIndex, 8).setValue(params.remarks || "");
      return { success: true, updated: true };
    } else {
      if (existingList.length >= 5) {
        return { success: false, error: "ขออภัยครับ รอบรถนี้เต็มโควตาฝากของ 5 ที่แล้ว!" };
      }
      
      const newRow = [
        new Date(),
        travelDate,
        timeSlot,
        origin,
        params.motorcycle ? true : false,
        params.cat ? true : false,
        params.otherCargo || "",
        params.remarks || "",
        slotId
      ];

      sheet.appendRow(newRow);
      const addedRow = sheet.getLastRow();
      
      sheet.getRange(addedRow, 5).insertCheckboxes();
      sheet.getRange(addedRow, 6).insertCheckboxes();

      return { success: true, updated: false };
    }
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function cancelCargoBooking(travelDate, origin, timeSlot, slotId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_CARGO);
    if (!sheet) return { success: false, error: "ไม่พบฐานข้อมูลฝากของ" };
    
    migrateCargoSheetIfNeeded(sheet);

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, error: "ไม่มีข้อมูลสัมภาระในระบบ" };
    
    const targetTravelDate = normalizeDate(travelDate);
    const targetOrigin = getStandardStationName(origin);
    const targetTimeSlot = getStandardTime(timeSlot);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rDate = normalizeDate(row[1]);
      const rTime = getStandardTime(row[2]);
      const rOrigin = getStandardStationName(row[3]);
      const rSlot = row[8] ? row[8].toString().trim() : "";
      if (rDate === targetTravelDate && rOrigin === targetOrigin && rTime === targetTimeSlot && rSlot === slotId) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: "ไม่พบข้อมูลสัมภาระช่อง: " + slotId };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}
