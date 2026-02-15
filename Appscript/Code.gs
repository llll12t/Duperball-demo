// ================================
// ค่าคงที่สำหรับ Google Sheet, Calendar และ LINE (Flex message)
const SPREADSHEET_ID = '1djH94Ht78Ig2sUoxdXaVzL37fcggxHSKeafHaWggcsY';
const CALENDAR_ID = '312221312e2b6d65b5a00e33db2b825a64477c1d9e8ee4dfbe63b6b4ce50fab2@group.calendar.google.com'; // Default calendar (main branch)
const CALENDAR_ID2 = '5a224d62bb0df1ea998da9969a7da80447b21d90ae73a23184288653b120fb8c@group.calendar.google.com';  
// ================================

// เพิ่ม Map สำหรับแต่ละสาขา
const CALENDAR_IDS = {
  'main': CALENDAR_ID,
  '1': CALENDAR_ID,
  '2': CALENDAR_ID2,
  // เพิ่มสาขาอื่นๆ ตามต้องการ
};
// Helper: เลือก Calendar ID ตามสาขา
function getCalendarIdByBranch(branch) {
  if (!branch || branch === 'main' || branch === '1') return CALENDAR_ID;
  return CALENDAR_IDS[branch] || CALENDAR_ID;
}
const LINE_TOKEN = 'j5YMPEd0AZdVnBum4ZlIEjWv24jYJMiYqdQQwxP7ggoWPFvFL6nrsTtYHOOnX4XeL+1X2HnPHqPCHtgmzHKvUHhAnJdHJbhz7ECK0ZQLid+cS8obWNbZYe/pT7O3UHyWmJIp3sL0m8d9l87StmtvKgdB04t89/1O/w1cDnyilFU=';

const LIFF_ID_CONFIRM = '2008293202-VJQZWvzL';

// ================================
// ค่าคงที่สำหรับ Telegram Bot
const TELEGRAM_BOT_TOKEN = 'XXXXXXX'; // Bot Token ของคุณ
const TELEGRAM_CHAT_ID = 'YYYYYYY';  // Chat ID

const SHEET_DATETIME = 'ตั้งค่าทั่วไป';   // ชีตเวลานัดหมาย
const SHEET_BOOKING = 'บันทึกนัดหมาย'; // ชีตการจอง
const SHEET_MEMBER = 'รายชื่อสมาชิก';   // ชีตสมาชิก
const SHEET_SERVICES = 'บริการ';         // ชีตข้อมูลบริการ
const SHEET_DOCTOR = 'หมอ'; // ชีตข้อมูลหมอ

// ----------------------------------------------------------------
// ส่งแจ้งเตือนผ่าน Telegram
function sendTelegramNotify(message) {
  const url = 'https://api.telegram.org/bot' + TELEGRAM_BOT_TOKEN + '/sendMessage';
  const payload = { chat_id: TELEGRAM_CHAT_ID, text: message };
  const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload) };
  try {
    const response = UrlFetchApp.fetch(url, options);
    Logger.log('Telegram response: ' + response.getContentText());
  } catch (error) {
    Logger.log('Telegram error: ' + error.toString());
  }
}

// ----------------------------------------------------------------
// Utility: เลือกชื่อชีตตามสาขา
function getSheetNameByBranch(baseName, branch) {
  // กรณีสาขาหลัก (1) ใช้ชื่อเดิม
  if (!branch || branch === 'main' || branch === '1') {
    return baseName;
  }
  
  // กรณีสาขาอื่นๆ (2, 3, ...) ให้เพิ่มหมายเลขสาขาต่อท้าย
  // รูปแบบ: ชื่อชีตเดิม_สาขา2, ชื่อชีตเดิม_สาขา3
  return `${baseName}_มีนบุรี${branch}`;
}

// ----------------------------------------------------------------
// ดึงรายการบริการ (รองรับหลายสาขา)
function fetchServices(e) {
  var branch = e && e.parameter && e.parameter.branch ? e.parameter.branch : '';
  var sheetName = getSheetNameByBranch(SHEET_SERVICES, branch);
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
  }
  var data = sheet.getDataRange().getValues().slice(1);
  var services = data.map(function(row) {
    return {
      id: row[0],
      name: row[1],
      details: row[2],
      price: row[3]
    };
  });
  return ContentService.createTextOutput(JSON.stringify(services)).setMimeType(ContentService.MimeType.JSON);
}

// ----------------------------------------------------------------
// ดึงรายละเอียดบริการตาม id
function getServiceById(id) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_SERVICES);
  const data = sheet.getDataRange().getValues().slice(1);
  for (let r of data) {
    if (r[0].toString() === id.toString()) {
      return { id: r[0], name: r[1], details: r[2], price: r[3] };
    }
  }
  return null;
}

// ----------------------------------------------------------------
// ดึง services + price จาก SHEET_DATETIME (เดิม)
function fetchServicesAndprice() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_DATETIME);
  const data = sheet.getDataRange().getValues().slice(1);
  const services = [...new Set(data.map(r => r[4]).filter(Boolean))];
  const price = [...new Set(data.map(r => r[5]).filter(Boolean))];
  const maxBookings = data.reduce((acc, row) => {
    if (row[0] && row[1]) acc[row[0]] = row[1];
    return acc;
  }, {});
  return ContentService.createTextOutput(JSON.stringify({ services, price, maxBookings }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ----------------------------------------------------------------
// Routing: doGet
function doGet(e) {
  const action = e.parameter.action;
  switch (action) {
    case 'fetchServices':
      return fetchServices(e);
    case 'fetchServicesAndprice':
      return fetchServicesAndprice();
    case 'getMonthAvailability':
      return getMonthAvailability(e);
    case 'fetchBookings':
      return fetchBookings(e);
    case 'fetchDoctorData':
      return fetchDoctorData(e);
    case 'getDoctorById':
      return getDoctorById(e.parameter.id);
    case 'getDoctorsByDate':
      return getDoctorsByDate(e.parameter.date, e.parameter.branch);
    case 'fetchDoctorNames':
      return fetchDoctorNames(e);
    // ===== CODE ที่แก้ไขอยู่ตรงนี้ =====
    case 'confirmBooking':
      return confirmBooking(e.parameter);
    case 'cancelBooking':
      return cancelBooking(e.parameter);
    // ================================
    default:
      return checkAvailability(e);
  }
}


// ----------------------------------------------------------------
// Routing: doPost
function doPost(e) {
  const action = e.parameter.action;
  let result = null;
  if (action === 'makeBooking') {
    result = makeBooking(e);
  } else if (action === 'cancelBooking') {
    result = cancelBooking(e.parameter);
  } else if (action === 'confirmBooking') {
    result = confirmBooking(e.parameter);
  } else if (action === 'completeBooking') {
    result = completeBooking(e);
  } else if (action === 'sendReminder') {
    result = sendReminder(e.parameter.id);
  } else if (action === 'sendReminderWithPayment') {
    result = sendReminderWithPayment(e.parameter);
  } else if (action === 'sendReminderWithoutPayment') {
    result = sendReminderWithoutPayment(e.parameter);
  } else if (action === 'saveDoctorData') {
    result = saveDoctorData(e.parameter);
  } else if (action === 'deleteDoctorData') {
    result = deleteDoctorData(e.parameter);
  } else if (action === 'updateAppointmentDoctor') {
    result = updateAppointmentDoctor(e.parameter);
  }
  // ไม่ต้องเช็คหรือเพิ่ม CORS header
  if (result) return result;
  // รองรับ LINE postback JSON
  if (e.postData && e.postData.type === 'application/json') {
    const data = JSON.parse(e.postData.contents);
    if (data.events) {
      data.events.forEach(evt => {
        if (evt.type === 'postback') {
          const params = parseQueryString(evt.postback.data);
          if (params.action === 'confirmBooking') return confirmBooking(params);
          if (params.action === 'cancelBooking') return cancelBooking(params);
        }
      });
    }
  }
  return ContentService.createTextOutput('Invalid action');
}

// แปลง query string เป็น object
function parseQueryString(qs) {
  return qs.split('&').reduce((o, p) => {
    const [k, v] = p.split('=');
    o[k] = decodeURIComponent(v);
    return o;
  }, {});
}

// ----------------------------------------------------------------
// ตรวจสอบความพร้อมจอง & วันหยุด (Legacy - ถูกแทนด้วย Workday Mode ด้านล่าง)
function checkAvailabilityLegacy(e) {
  if (!e.parameter.date && !e.parameter.getHolidays) {
    return ContentService.createTextOutput(JSON.stringify({ error: "Missing params" }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var branch = e && e.parameter && e.parameter.branch ? e.parameter.branch : '';
  var sheetName = getSheetNameByBranch(SHEET_DATETIME, branch);
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  const rows = sheet.getDataRange().getValues();

  // --- ส่วนที่ปรับปรุง: เพิ่มการตรวจสอบว่ามีข้อมูลในแถวที่ 2 หรือไม่ ---
  // ป้องกัน Error หากในชีตมีแค่ Header
  const permanentHolidaysSetting = (rows && rows.length > 1 && rows[1][4]) ? String(rows[1][4]) : '';
  const permanentHolidayDays = permanentHolidaysSetting.split(',').map(d => parseInt(d.trim())).filter(d => !isNaN(d));
  // --- สิ้นสุดส่วนที่ปรับปรุง ---

  const dataRows = rows.slice(1);

  if (e.parameter.getHolidays) {
    const holidays = dataRows.map(r => {
      const v = r[3];
      if (!v) return null;
      if (typeof v === 'string' && /^\d{2}-\d{2}-\d{4}$/.test(v)) {
        const [d, m, y] = v.split('-');
        return `${y}-${m}-${d}`;
      }
      if (v instanceof Date) {
        return Utilities.formatDate(v, "Asia/Bangkok", "yyyy-MM-dd");
      }
      return null;
    }).filter(Boolean);

    return ContentService.createTextOutput(JSON.stringify({
      holidays: Array.from(new Set(holidays)),
      permanentHolidays: permanentHolidayDays
    }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const inputDate = new Date(e.parameter.date);
  if (isNaN(inputDate)) {
    return ContentService.createTextOutput(JSON.stringify({ error: "Invalid date" }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const times = dataRows.map(r => {
    const t = r[1];
    if (!t) return null;
    if (t instanceof Date) return Utilities.formatDate(t, "Asia/Bangkok", "HH:mm");
    if (typeof t === 'string' && /^\d{2}:\d{2}:\d{2}$/.test(t)) return t.substring(0, 5);
    return t.toString();
  }).filter(Boolean);
  const maxBookings = dataRows.map(r => Number(r[2]) || 0);

  const holidays = dataRows.map(r => {
    const v = r[3];
    if (!v) return null;
    if (typeof v === 'string' && /^\d{2}-\d{2}-\d{4}$/.test(v)) {
      const [d, m, y] = v.split('-');
      return new Date(`${y}-${m}-${d}`);
    }
    if (v instanceof Date) return v;
    return null;
  }).filter(Boolean);

  const formattedInputDate = Utilities.formatDate(inputDate, "Asia/Bangkok", "yyyy-MM-dd");

  const isSpecificHoliday = holidays.some(h => Utilities.formatDate(h, "Asia/Bangkok", "yyyy-MM-dd") === formattedInputDate);
  const isPermanentHoliday = permanentHolidayDays.includes(inputDate.getDay());
  const isHoliday = isSpecificHoliday || isPermanentHoliday;

  let availability = {};
  if (inputDate < today) {
    times.forEach(t => availability[t] = { status: 'เต็ม', remaining: 0, total: 0 });
  } else {
  const calendar = CalendarApp.getCalendarById(getCalendarIdByBranch(branch));
    times.forEach((t, i) => {
      let status = 'Available', remaining = maxBookings[i], total = maxBookings[i];
      if (isHoliday) {
        status = 'ไม่ว่าง';
        remaining = 0;
      } else {
        const [hh, mm] = t.split(':').map(Number);
        const start = new Date(inputDate); start.setHours(hh, mm);
        const end = new Date(start.getTime() + 15 * 60 * 1000); // เปลี่ยนจาก 30 เป็น 15 นาที
        const events = calendar.getEvents(start, end).filter(ev => ev.getTitle().indexOf('นัดคิว:') > -1);
        remaining = total - events.length;
        if (remaining <= 0) { remaining = 0; status = 'เต็ม'; }
      }
      availability[t] = { status, remaining, total };
    });
  }

  return ContentService.createTextOutput(JSON.stringify({
    availability,
    isHoliday,
    color: isHoliday ? 'red' : 'default'
  })).setMimeType(ContentService.MimeType.JSON);
}

// ----------------------------------------------------------------
// สร้างการจอง (รองรับหลายสาขา)
function makeBooking(e) {
  var branch = e && e.parameter && e.parameter.branch ? e.parameter.branch : '';
  var sheetNameBooking = getSheetNameByBranch(SHEET_BOOKING, branch);
  var sheetNameDatetime = getSheetNameByBranch(SHEET_DATETIME, branch);
  var sheetNameMember = getSheetNameByBranch(SHEET_MEMBER, branch);
  const cal = CalendarApp.getCalendarById(getCalendarIdByBranch(branch));
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(sheetNameBooking);
  if (!sheet) {
    sheet = ss.insertSheet(sheetNameBooking);
    // สร้าง Header สำหรับชีตการจอง
    sheet.getRange(1, 1, 1, 18).setValues([
      ['ID', 'userlineid', 'firstName', 'lastName', 'phonenumber', 'idCardOrSocial', 'diseaseAllergy', 'note', 'serviceNames', 'totalPrice', 'formattedDate', 'time', 'status', 'timestamp', 'calendarEventId', 'doctor', 'room', 'branch', 'lang']
    ]);
  }
  let dtSheet = ss.getSheetByName(sheetNameDatetime);
  if (!dtSheet) {
    dtSheet = ss.insertSheet(sheetNameDatetime);
    // สร้าง Header สำหรับชีตเวลานัดหมาย
    dtSheet.getRange(1, 1, 1, 6).setValues([
      ['เวลา', 'maxBookings', 'maxBookings', 'specialWorkday', 'weeklyWorkdays', 'doctorName']
    ]);
  }
  let memSheet = ss.getSheetByName(sheetNameMember);
  if (!memSheet) {
    memSheet = ss.insertSheet(sheetNameMember);
    // สร้าง Header สำหรับชีตสมาชิก
    memSheet.getRange(1, 1, 1, 8).setValues([
      ['ID', 'userlineid', 'firstName', 'lastName', 'phonenumber', 'idCardOrSocial', 'diseaseAllergy', 'timestamp']
    ]);
  }

  const userlineid = e.parameter.userlineid || '';
  const firstName = e.parameter.first_name || '';
  const lastName = e.parameter.last_name || '';
  const name = (firstName + ' ' + lastName).trim();
  let phonenumber = e.parameter.phonenumber || '';
  const idCardOrSocial = e.parameter.id_card_or_social || '';
  const diseaseAllergy = e.parameter.disease_allergy || '';
  const note = e.parameter.note || '';
  const lang = e.parameter.lang || 'th'; // รับค่าภาษา
  const timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd-MM-yyyy HH:mm:ss");
  const date = e.parameter.date || ''; // รูปแบบ: yyyy-MM-dd
  const time = e.parameter.time || ''; // รูปแบบ: HH:mm
  
  // ตรวจสอบข้อมูลที่จำเป็น
  if (!date || !time || !firstName || !lastName || !phonenumber || !idCardOrSocial) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'ข้อมูลไม่ครบถ้วน กรุณากรอกข้อมูลให้ครบ'
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  // แปลงวันที่อย่างปลอดภัย
  const dateParts = date.split('-'); // ["yyyy", "MM", "dd"]
  const datetime = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
  
  // เพิ่มเวลาเข้าไป
  const [hh, mm] = time.split(':').map(Number);
  datetime.setHours(hh, mm, 0, 0);

  // ปิดทุกวันเป็นค่าเริ่มต้น: ต้องอยู่ในวันทำงานถึงจะจองได้
  const dtRowsForWorkday = dtSheet.getDataRange().getValues();
  const workdayConfig = extractWorkdayConfigV2(dtRowsForWorkday);
  if (!isWorkingDateV2(datetime, workdayConfig.weeklyWorkdays, workdayConfig.specialWorkdays)) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'วันที่เลือกไม่เปิดให้บริการ'
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  // Debug: Log วันที่เพื่อตรวจสอบ
  Logger.log('Input date string: ' + date);
  Logger.log('Parsed datetime: ' + datetime);
  Logger.log('Will format to: ' + Utilities.formatDate(datetime, "Asia/Bangkok", 'dd-MM-yyyy'));

  // ✅ เพิ่ม ' เพื่อป้องกันการตัด 0 ออก
  phonenumber = "'" + String(phonenumber).trim();

  // ตรวจ capacity
  const times = dtSheet.getRange('B2:B').getValues().flat();
  const idx = times.indexOf(time);
  if (idx >= 0) {
    const maxBookingsData = dtSheet.getDataRange().getValues().slice(1).map(r => Number(r[2]) || 0);
    const maxBk = maxBookingsData[idx];
    const events = cal.getEvents(datetime, new Date(datetime.getTime() + 15 * 60 * 1000)); // เปลี่ยนจาก 30 เป็น 15 นาที
    if (events.length >= maxBk) {
      return ContentService.createTextOutput('ช่วงเวลาที่เลือกเต็มแล้ว');
    }
  }

  const formattedDate = Utilities.formatDate(datetime, "Asia/Bangkok", 'dd-MM-yyyy');

  // ดึงบริการ
  let serviceNames = "", totalPrice = 0;
  try {
    const arr = JSON.parse(e.parameter.selectedServices || '[]');
    const details = arr.map(id => getServiceById(id)).filter(s => s);
    serviceNames = details.map(s => {
      let price = Number(s.price);
      if (isNaN(price)) price = 0;
      // ถ้าราคาเป็น 0 ไม่ต้องแสดง (ราคา 0)
      return price > 0 ? `${s.name} (ราคา ${price})` : `${s.name}`;
    }).join(", ");
    totalPrice = details.reduce((sum, s) => {
      let price = Number(s.price);
      if (isNaN(price)) price = 0;
      return sum + price;
    }, 0);
  } catch {
    serviceNames = e.parameter.selectedServices;
  }

  // สร้าง Event
  const event = cal.createEvent(`นัดคิว: ${serviceNames}`, datetime, new Date(datetime.getTime() + 15 * 60 * 1000), { // เปลี่ยนจาก 30 เป็น 15 นาที
    description: `ลูกค้า: ${name}\nเบอร์: ${phonenumber}\nเลขบัตร/ประกัน: ${idCardOrSocial}\nโรค/แพ้: ${diseaseAllergy}\nหมายเหตุ: ${note}`
  });
  const calendarEventId = event.getId();

  const nextRow = sheet.getLastRow() + 1;
  // Use branch number in ID, fallback to 0 if not a number
  let branchNum = branch;
  if (typeof branchNum === 'string') {
    branchNum = branchNum.replace(/[^0-9]/g, '');
    if (!branchNum) branchNum = '0';
  }
  // Scan all existing IDs for this branch and find max running number
  let maxRunning = 0;
  const allIds = sheet.getRange(2, 1, Math.max(0, sheet.getLastRow()-1), 1).getValues().flat();
  const idPattern = new RegExp('^BK' + branchNum + '-(\\d+)$');
  allIds.forEach(function(existingId) {
    const m = idPattern.exec(existingId);
    if (m && m[1]) {
      const num = parseInt(m[1], 10);
      if (num > maxRunning) maxRunning = num;
    }
  });
  const id = 'BK' + branchNum + '-' + (maxRunning + 1);
  // ปรับคอลัมน์: [id, userlineid, firstName, lastName, phonenumber, idCardOrSocial, diseaseAllergy, note, serviceNames, totalPrice, formattedDate, time, status, timestamp, calendarEventId, "", "", branch]
  // branch ต้องอยู่คอลัมน์ R (column 18)
  const newRow = [
    id,               // A (1)
    userlineid,       // B (2)
    firstName,        // C (3)
    lastName,         // D (4)
    phonenumber,      // E (5)
    idCardOrSocial,   // F (6)
    diseaseAllergy,   // G (7)
    note,             // H (8)
    serviceNames,     // I (9)
    totalPrice,       // J (10)
    formattedDate,    // K (11)
    time,             // L (12)
    'ยืนยันแล้ว',      // M (13) -- เปลี่ยนสถานะเริ่มต้นเป็น 'ยืนยันแล้ว'
    timestamp,        // N (14)
    calendarEventId,  // O (15)
    '',               // P (16) - doctor
    '',               // Q (17) - room
    branch,           // R (18)
    lang              // S (19) - language
  ];
  sheet.getRange(nextRow, 1, 1, newRow.length).setValues([newRow]);

  // กำหนดฟอร์แมทเป็นข้อความ
  sheet.getRange(nextRow, 5).setNumberFormat("@"); // Phone
  sheet.getRange(nextRow, 11).setNumberFormat("@"); // Date
  sheet.getRange(nextRow, 12).setNumberFormat("@"); // Time
  sheet.getRange(nextRow, 14).setNumberFormat("@"); // Timestamp
  sheet.getRange(nextRow, 15).setNumberFormat("@"); // Calendar Event ID

  // อัปเดตหรือเพิ่มสมาชิก (ตัวอย่าง: อัปเดตเฉพาะชื่อ-นามสกุลและเบอร์)
  const dataMem = memSheet.getDataRange().getValues();
  const existIdx = dataMem.findIndex(r => r[1] === userlineid) + 1;
  if (existIdx > 1) {
    memSheet.getRange(existIdx, 3).setValue(firstName);
    memSheet.getRange(existIdx, 4).setValue(lastName);
    memSheet.getRange(existIdx, 5).setNumberFormat("@").setValue(phonenumber);
    memSheet.getRange(existIdx, 6).setValue(idCardOrSocial);
    memSheet.getRange(existIdx, 7).setValue(diseaseAllergy);
    memSheet.getRange(existIdx, 8).setValue(timestamp);
  } else {
    memSheet.appendRow([id, userlineid, firstName, lastName, phonenumber, idCardOrSocial, diseaseAllergy, timestamp]);
    memSheet.getRange(memSheet.getLastRow(), 5).setNumberFormat("@");
  }

  // แจ้ง Telegram
  sendTelegramNotify(
    `📅 นัดหมายใหม่\n👤: ${name}\n📋: ${serviceNames}\n💰: ${totalPrice}\n📅: ${formattedDate}\n⏰: ${time}`
  );

  return ContentService.createTextOutput('Booking successful!')
    .setMimeType(ContentService.MimeType.JSON);
}


// ----------------------------------------------------------------
// แจ้งเตือนพร้อมจ่ายเงิน
function sendReminderWithPayment(params) {
  const id = params.id;
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_BOOKING);
  const data = sheet.getDataRange().getValues();
  const idx = data.findIndex(r => r[0] == id);
  if (idx < 0) return ContentService.createTextOutput('Booking not found');
  const b = data[idx];
  const totalPrice = b[9]; // Column J: totalPrice
  const fullName = (b[2] || '') + ' ' + (b[3] || ''); // firstName + lastName
  // Remove (ราคา ...) from serviceNames for Flex
  const serviceNoPrice = (b[8] || '').split(',').map(s => s.replace(/\(ราคา.*?\)/g, '').trim()).join(', ');
  const qrCodeUrl = `https://promptpay.io/0623733306/${totalPrice}.png`;
  const liffUrl = `https://liff.line.me/2006029649-EbKnbZJ0?id=${id}&userlineid=${b[1]}&name=${encodeURIComponent(fullName)}&contact=${encodeURIComponent(b[4])}&serviceNames=${encodeURIComponent(serviceNoPrice)}&totalPrice=${totalPrice}`;
  const reminderData = {
    idKey: b[0], service: serviceNoPrice,
    date: b[10], time: b[11], name: fullName, // Column K: date, Column L: time
    qrCodeUrl, liffUrl
  };
  const lang = b[18] || 'th'; // Column S: lang (index 18)
  sendLineMessage(b[1], 'reminder_with_payment', reminderData, lang);
  return ContentService.createTextOutput('Reminder with payment sent.');
}

// ----------------------------------------------------------------
// แจ้งเตือน without payment
function sendReminderWithoutPayment(params) {
  const id = params.id;
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_BOOKING);
  const data = sheet.getDataRange().getValues();
  const idx = data.findIndex(r => r[0] == id);
  if (idx < 0) return ContentService.createTextOutput('Booking not found');
  const b = data[idx];
  const fullName = (b[2] || '') + ' ' + (b[3] || ''); // firstName + lastName
  // Remove (ราคา ...) from serviceNames for Flex
  const serviceNoPrice = (b[8] || '').split(',').map(s => s.replace(/\(ราคา.*?\)/g, '').trim()).join(', ');
  const reminderData = {
    idKey: b[0], service: serviceNoPrice,
    date: b[10], time: b[11], name: fullName // Column K: date, Column L: time
  };
  const lang = b[18] || 'th'; // Column S: lang (index 18)
  sendLineMessage(b[1], 'reminder_no_payment', reminderData, lang);
  return ContentService.createTextOutput('Reminder without payment sent.');
}

// ----------------------------------------------------------------
// ยืนยันการจอง
function confirmBooking(params) {
  const idKey = params.id;
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // ค้นหาในทุกสาขา
  const branches = ['1', '2']; // เพิ่มสาขาตามต้องการ
  let sheet = null;
  let idx = -1;
  let data = null;
  
  for (const branch of branches) {
    const sheetName = getSheetNameByBranch(SHEET_BOOKING, branch);
    const testSheet = spreadsheet.getSheetByName(sheetName);
    if (!testSheet) continue;
    
    const testData = testSheet.getDataRange().getValues();
    const testIdx = testData.findIndex(r => r[0] == idKey);
    
    if (testIdx >= 0) {
      sheet = testSheet;
      idx = testIdx;
      data = testData;
      break;
    }
  }
  
  if (idx < 0) return ContentService.createTextOutput('ไม่พบการนัดหมายนี้.');
  const booking = data[idx];
  const current = booking[12]; // Column M: status
  if (current !== 'รอการยืนยัน') {
    const msg = `ไม่สามารถยืนยันได้ สถานะปัจจุบันคือ "${current}"`;
    sendLineMessage(booking[1], 'notification', { message: msg });
    return ContentService.createTextOutput(msg);
  }
  sheet.getRange(idx + 1, 13).setValue('ยืนยันแล้ว'); // Column M: status
  const fullName = (booking[2] || '') + ' ' + (booking[3] || ''); // firstName + lastName
  // Remove (ราคา ...) from serviceNames for Flex
  const serviceNoPrice = (booking[8] || '').split(',').map(s => s.replace(/\(ราคา.*?\)/g, '').trim()).join(', ');
  const lang = booking[18] || 'th'; // Column S: lang (index 18)
  sendLineMessage(booking[1], 'confirmBooking', {
    service: serviceNoPrice,
    date: booking[10], time: booking[11], // Column K: date, Column L: time
    name: fullName, status: 'ยืนยันแล้ว', idKey
  }, lang);
  return ContentService.createTextOutput('Booking confirmed successfully.');
}

// ----------------------------------------------------------------
// เสร็จสิ้นการจอง
function completeBooking(e) {
  var branch = e && e.parameter && e.parameter.branch ? e.parameter.branch : '';
  var sheetName = getSheetNameByBranch(SHEET_BOOKING, branch);
  const idKey = e.parameter.id;
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const idx = data.findIndex(r => r[0] == idKey);
  if (idx < 0) return ContentService.createTextOutput('Booking not found.');
  sheet.getRange(idx + 1, 13).setValue('เสร็จสิ้น'); // Column M: status
  const booking = data[idx];
  const fullName = (booking[2] || '') + ' ' + (booking[3] || ''); // firstName + lastName
  // Remove (ราคา ...) from serviceNames for Flex
  const serviceNoPrice = (booking[8] || '').split(',').map(s => s.replace(/\(ราคา.*?\)/g, '').trim()).join(', ');
  const lang = booking[18] || 'th'; // Column S: lang (index 18)
  sendLineMessage(booking[1], 'completion', {
    service: serviceNoPrice,
    date: booking[10], time: booking[11], // Column K: date, Column L: time
    name: fullName, status: 'เสร็จสิ้น'
  }, lang);
  return ContentService.createTextOutput('Booking completed successfully.');
}

// ----------------------------------------------------------------
// ยกเลิกการจอง
function cancelBooking(params) {
  const idKey = params.id;
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // ค้นหาในทุกสาขา
  const branches = ['1', '2']; // เพิ่มสาขาตามต้องการ
  let sheet = null;
  let idx = -1;
  let data = null;
  let foundBranch = '';
  
  for (const branch of branches) {
    const sheetName = getSheetNameByBranch(SHEET_BOOKING, branch);
    const testSheet = spreadsheet.getSheetByName(sheetName);
    if (!testSheet) continue;
    
    const testData = testSheet.getDataRange().getValues();
    const testIdx = testData.findIndex(r => r[0] == idKey);
    
    if (testIdx >= 0) {
      sheet = testSheet;
      idx = testIdx;
      data = testData;
      foundBranch = branch;
      break;
    }
  }
  
  if (idx < 0) return ContentService.createTextOutput('ไม่พบการนัดหมายนี้.');
  const booking = data[idx];
  const current = booking[12]; // Column M: status
  if (current !== 'รอการยืนยัน') {
    const msg = `ไม่สามารถยกเลิกได้ สถานะปัจจุบันคือ "${current}"`;
    sendLineMessage(booking[1], 'notification', { message: msg });
    return ContentService.createTextOutput(msg);
  }
  sheet.getRange(idx + 1, 13).setValue('ยกเลิก'); // Column M: status

  const cal = CalendarApp.getCalendarById(getCalendarIdByBranch(foundBranch));
  try {
    const calendarEventId = booking[14]; // Column O: calendarEventId
    if (calendarEventId) {
      const ev = cal.getEventById(calendarEventId);
      if (ev) {
        ev.deleteEvent();
        Logger.log(`Event ID ${calendarEventId} deleted successfully.`);
      } else {
        Logger.log(`Calendar event with ID ${calendarEventId} not found.`);
      }
    } else {
      Logger.log(`No Calendar Event ID found for booking ${idKey}.`);
    }
  } catch (e) {
    Logger.log(`Error deleting calendar event for booking ${idKey}: ${e.toString()}`);
  }

  const fullName = (booking[2] || '') + ' ' + (booking[3] || ''); // firstName + lastName
  // Remove (ราคา ...) from serviceNames for Flex
  const serviceNoPrice = (booking[8] || '').split(',').map(s => s.replace(/\(ราคา.*?\)/g, '').trim()).join(', ');
  const lang = booking[18] || 'th'; // Column S: lang (index 18)
  sendLineMessage(booking[1], 'cancellation', {
    service: serviceNoPrice,
    date: booking[10], time: booking[11], // Column K: date, Column L: time
    name: fullName, status: 'ยกเลิก'
  }, lang);
  return ContentService.createTextOutput('Booking cancelled successfully.');
}
// ----------------------------------------------------------------
// ส่ง reminder ทันที
function sendReminder(id) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_BOOKING);
  const data = sheet.getDataRange().getValues();
  const idx = data.findIndex(r => r[0] == id);
  if (idx < 0) return ContentService.createTextOutput('Booking not found.');
  const b = data[idx];
  const fullName = (b[2] || '') + ' ' + (b[3] || ''); // firstName + lastName
  // Remove (ราคา ...) from serviceNames for Flex
  const serviceNoPrice = (b[8] || '').split(',').map(s => s.replace(/\(ราคา.*?\)/g, '').trim()).join(', ');
  const lang = b[18] || 'th'; // Column S: lang (index 18)
  sendLineMessage(b[1], 'reminder', {
    idKey: b[0], service: serviceNoPrice,
    date: b[10], time: b[11], name: fullName, status: b[12] // Column K: date, Column L: time, Column M: status
  }, lang);
  return ContentService.createTextOutput('Reminder sent successfully.');
}

// ----------------------------------------------------------------
// --- วางทับฟังก์ชันนี้ทั้งหมด ---
function sendReminders() {
  const today = new Date();
  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_BOOKING);
  const data = sheet.getDataRange().getValues().slice(1);

  data.forEach(row => {
    try {
      const dateStr = row[10]; // Column K: formattedDate
      const status = row[12]; // Column M: status

      if (!dateStr || (status !== 'รอการยืนยัน' && status !== 'ยืนยันแล้ว')) {
        return; // ข้ามถ้าสถานะไม่ใช่ 'รอการยืนยัน' หรือ 'ยืนยันแล้ว'
      }

      // *** จุดที่แก้ไข: เปลี่ยนวิธีแปลงวันที่ให้ปลอดภัยและแน่นอน ***
      const dateParts = dateStr.split('-'); // ["DD", "MM", "YYYY"]
      if (dateParts.length !== 3) return; // ข้ามถ้า format ผิด

      // new Date(year, monthIndex, day) // เดือนใน JS เริ่มจาก 0 (ม.ค.=0)
      const bookingDate = new Date(parseInt(dateParts[2]), parseInt(dateParts[1]) - 1, parseInt(dateParts[0]));

      // เปรียบเทียบวันที่
      if (bookingDate.getDate() === tomorrow.getDate() &&
        bookingDate.getMonth() === tomorrow.getMonth() &&
        bookingDate.getFullYear() === tomorrow.getFullYear()) {

        const fullName = (row[2] || '') + ' ' + (row[3] || ''); // firstName + lastName
        const lang = row[18] || 'th'; // Column S: lang (index 18)
        sendLineMessage(row[1], 'reminder', {
          idKey: row[0],
          service: row[8], // Column I: serviceNames
          date: row[10], // Column K: formattedDate
          time: row[11], // Column L: time
          name: fullName,
          status: 'รอการยืนยัน'
        }, lang);
      }
    } catch (e) {
      Logger.log(`Error processing row for sendReminders: ${e.toString()}. Row data: ${row}`);
    }
  });
}

// สร้าง Trigger สำหรับ sendReminders
function createReminderTrigger() {
  ScriptApp.newTrigger('sendReminders')
    .timeBased().everyDays(1).atHour(12).create();
}

// ==========================================================
// ========== ฟังก์ชันสำหรับทดสอบ (DEMO) ==========
// ==========================================================
function testReminderFlex() {
  // 1. จำลองข้อมูลการจอง (เหมือนข้อมูลที่ดึงมาจากชีต)
  // คุณสามารถลองเปลี่ยนค่าพวกนี้เพื่อทดสอบได้
  const mockBookingData = {
    idKey: 'BK-DEMO-001',
    name: 'ลูกค้าทดสอบ',
    service: 'ตัดผม, ทำสีผม',
    price: '1200',
    date: '15-08-2025',
    time: '13:00',
    status: 'รอการยืนยัน'
  };

  try {
    // 2. เรียกใช้ฟังก์ชัน createFlexMessage โดยตรงด้วยข้อมูลจำลอง
    // และระบุ type เป็น 'reminder'
    const flexMessageJson = createFlexMessage('reminder', mockBookingData);

    // 3. แสดงผลลัพธ์ใน Log เพื่อให้เรานำไปตรวจสอบ
    Logger.log("========== DEMO START ==========");
    Logger.log("--- INPUT DATA ---");
    Logger.log(JSON.stringify(mockBookingData, null, 2)); // แสดงข้อมูลจำลองที่ใช้
    Logger.log("--- OUTPUT FLEX JSON ---");
    Logger.log(JSON.stringify(flexMessageJson, null, 2)); // แสดงผลลัพธ์ Flex Message ที่สร้างได้
    Logger.log("========== DEMO END ==========");

  } catch (e) {
    Logger.log("!!! AN ERROR OCCURRED DURING FLEX CREATION !!!");
    Logger.log(e.toString());
  }
}
// ----------------------------------------------------------------
// ส่งข้อความผ่าน LINE
function sendLineMessage(userId, type, data, lang = 'th') {
  const url = 'https://api.line.me/v2/bot/message/push';

  try {
    const msg = createFlexMessage(type, data, lang);
    if (!msg) {
      Logger.log(`Failed to create Flex Message for type: ${type} and data: ${JSON.stringify(data)}`);
      return;
    }

    const payload = { to: userId, messages: [msg] };
    const payloadString = JSON.stringify(payload);

    const opts = {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + LINE_TOKEN },
      payload: payloadString,
      muteHttpExceptions: true // *** IMPORTANT: This allows us to get the error response body ***
    };

    Logger.log('--- Sending LINE Message ---');
    Logger.log('User ID: ' + userId);
    Logger.log('Payload: ' + payloadString);

    const resp = UrlFetchApp.fetch(url, opts);
    const responseCode = resp.getResponseCode();
    const responseBody = resp.getContentText();

    Logger.log('LINE Response Code: ' + responseCode);
    Logger.log('LINE Response Body: ' + responseBody);

    if (responseCode !== 200) {
      // Also send a notification to Telegram for easier debugging
      sendTelegramNotify(`LINE API Error!\nCode: ${responseCode}\nUser: ${userId}\nResponse: ${responseBody}`);
    }

  } catch (err) {
    Logger.log('!!! Uncaught Exception in sendLineMessage !!!');
    Logger.log('Error: ' + err.toString());
    Logger.log('Stack: ' + err.stack);
    // Also send a notification to Telegram for easier debugging
    sendTelegramNotify(`FATAL: Uncaught error in sendLineMessage for user ${userId}: ${err.toString()}`);
  }
}
// ================================================================
// Multi-language Translations
// ================================================================
const TRANSLATIONS = {
  th: {
    header: {
      confirmBooking: '✅ ยืนยันการนัดหมาย\nAppointment Confirmed',
      cancellation: '⚠️ ยกเลิกการนัดหมาย\nAppointment Cancelled',
      completion: '✅ นัดหมายเสร็จสิ้น\nAppointment Completed',
      reminder: '🔔 แจ้งเตือนการนัดหมาย\nAppointment Reminder',
      confirmation: '🔔 การนัดหมาย\nAppointment',
      reminder_with_payment: '🔔 แจ้งเตือนและชำระเงิน',
      reminder_no_payment: '🔔 แจ้งเตือนการนัดหมาย'
    },
    status: {
      confirmBooking: 'ยืนยันแล้ว',
      cancellation: 'ยกเลิก',
      completion: 'เสร็จสิ้น',
      reminder: 'รอการยืนยัน',
      confirmation: 'รอการยืนยัน'
    },
    label: {
      name: 'ชื่อผู้จอง',
      service: 'บริการ',
      date: 'วันที่',
      time: 'เวลา',
      status: 'สถานะ',
      subheader_name: 'Name',
      subheader_service: 'Service',
      subheader_date: 'Date',
      subheader_time: 'Time',
      subheader_status: 'Status'
    },
    footer: {
      manage: 'จัดการนัดหมาย | Manage',
      payment: 'ชำระเงิน | Payment',
      cancelled: 'การนัดหมายถูกยกเลิกแล้ว',
      completed: 'ขอบคุณที่ใช้บริการ',
      saved: 'บันทึกนัดหมายแล้ว',
      early: 'กรุณามาก่อน 10-20 นาที',
      cancel_policy: 'หากต้องการยกเลิกหรือเลื่อนนัด กรุณาติดต่อเจ้าหน้าที่'
    }
  },
  en: {
    header: {
      confirmBooking: '✅ Appointment Confirmed',
      cancellation: '⚠️ Appointment Cancelled',
      completion: '✅ Appointment Completed',
      reminder: '🔔 Appointment Reminder',
      confirmation: '🔔 Appointment Pending',
      reminder_with_payment: '🔔 Reminder & Payment',
      reminder_no_payment: '🔔 Appointment Reminder'
    },
    status: {
      confirmBooking: 'Confirmed',
      cancellation: 'Cancelled',
      completion: 'Completed',
      reminder: 'Pending',
      confirmation: 'Pending'
    },
    label: {
      name: 'Name',
      service: 'Service',
      date: 'Date',
      time: 'Time',
      status: 'Status',
      subheader_name: '',
      subheader_service: '',
      subheader_date: '',
      subheader_time: '',
      subheader_status: ''
    },
    footer: {
      manage: 'Manage Appointment',
      payment: 'Payment',
      cancelled: 'Appointment has been cancelled',
      completed: 'Thank you for your service',
      saved: 'Appointment saved successfully',
      early: 'Please arrive 10-20 minutes early',
      cancel_policy: 'If you wish to cancel/reschedule, please contact staff.'
    }
  },
  my: { // Myanmar
    header: {
      confirmBooking: '✅ ကြိုတင်မှာယူမှု အတည်ပြုပြီး',
      cancellation: '⚠️ ကြိုတင်မှာယူမှု ပယ်ဖျက်ပြီး',
      completion: '✅ ပြီးဆုံးသည်',
      reminder: '🔔 သတိပေးချက်',
      confirmation: '🔔 ကြိုတင်မှာယူမှု',
      reminder_with_payment: '🔔 သတိပေးချက်နှင့်ငွေပေးချေမှု',
      reminder_no_payment: '🔔 သတိပေးချက်'
    },
    status: {
      confirmBooking: 'အတည်ပြုပြီး (Confirmed)',
      cancellation: 'ဖျက်သိမ်းပြီး (Cancelled)',
      completion: 'ပြီးဆုံးသည် (Completed)',
      reminder: 'စောင့်ဆိုင်းနေသည် (Pending)',
      confirmation: 'စောင့်ဆိုင်းနေသည် (Pending)'
    },
    label: {
      name: 'အမည်',
      service: 'ဝန်ဆောင်မှု',
      date: 'နေ့ရက်',
      time: 'အချိန်',
      status: 'အခြေအနေ',
      subheader_name: 'Name',
      subheader_service: 'Service',
      subheader_date: 'Date',
      subheader_time: 'Time',
      subheader_status: 'Status'
    },
    footer: {
      manage: 'စီမံခန့်ခွဲပါ | Manage',
      payment: 'ငွေပေးချေမှု | Payment',
      cancelled: 'ကြိုတင်မှာယူမှု ပယ်ဖျက်လိုက်ပါပြီ',
      completed: 'ဝန်ဆောင်မှုကို အသုံးပြုသည့်အတွက် ကျေးဇူးတင်ပါသည်',
      saved: 'ကြိုတင်မှာယူမှု သိမ်းဆည်းပြီး',
      early: 'ကျေးဇူးပြု၍ ၁၀-၂၀ မိနစ် ကြိုရောက်ပေးပါ',
      cancel_policy: 'ပယ်ဖျက်လိုပါက ဝန်ထမ်းများနှင့် ဆက်သွယ်ပါ။'
    }
  },
  km: { // Khmer
    header: {
      confirmBooking: '✅ ការកក់ត្រូវបានបញ្ជាក់',
      cancellation: '⚠️ ការកក់ត្រូវបានបោះបង់',
      completion: '✅ បានបញ្ចប់',
      reminder: '🔔 ការជូនដំណឹង',
      confirmation: '🔔 ការកក់',
      reminder_with_payment: '🔔 ការជូនដំណឹង និងការទូទាត់',
      reminder_no_payment: '🔔 ការជូនដំណឹង'
    },
    status: {
      confirmBooking: 'បានបញ្ជាក់ (Confirmed)',
      cancellation: 'បានបោះបង់ (Cancelled)',
      completion: 'បានបញ្ចប់ (Completed)',
      reminder: 'កំពុងរង់ចាំ (Pending)',
      confirmation: 'កំពុងរង់ចាំ (Pending)'
    },
    label: {
      name: 'ឈ្មោះ',
      service: 'សេវាកម្ម',
      date: 'កាលបរិច្ឆេទ',
      time: 'ពេលវេលា',
      status: 'ស្ថានភាព',
      subheader_name: 'Name',
      subheader_service: 'Service',
      subheader_date: 'Date',
      subheader_time: 'Time',
      subheader_status: 'Status'
    },
    footer: {
      manage: 'គ្រប់គ្រង | Manage',
      payment: 'ការទូទាត់ | Payment',
      cancelled: 'ការកក់ត្រូវបានបោះបង់ហើយ',
      completed: 'សូមអរគុណសម្រាប់ការប្រើប្រាស់សេវាកម្ម',
      saved: 'បានរក្សាទុកការកក់',
      early: 'សូមមកដល់មុន ១០-២០ នាទី',
      cancel_policy: 'ប្រសិនបើអ្នកចង់បោះបង់ សូមទាក់ទងបុគ្គលិក។'
    }
  }
};

// --- วางทับฟังก์ชันนี้ทั้งหมดใน AP V2.js ---
function createFlexMessage(type, data, lang = 'th') {
  if (type === 'notification') {
    return {
      type: 'text',
      text: data.message
    };
  }
  
  // ตรวจสอบภาษา ถ้าไม่มีให้ใช้ th
  const t = TRANSLATIONS[lang] || TRANSLATIONS['th'];

  const serviceText = data.service || 'ไม่มีบริการระบุ';
  const nameText = data.name || 'ไม่มีชื่อ';
  const dateText = data.date || 'ไม่มีวันที่';
  const timeText = data.time || 'ไม่มีเวลา';
  const idKeyText = data.idKey || '';

  const headerText = t.header[type] || t.header['confirmation'] || 'Detail';
  const statusTextValue = t.status[type] || t.status['reminder'];

  const statusColor = {
    confirmBooking: '#06C755',
    cancellation: '#FF0000',
    completion: '#06C755',
    reminder: '#FFA500',
    confirmation: '#1f1f97',
    reminder_with_payment: '#FFA500',
    reminder_no_payment: '#FFA500'
  }[type] || '#2F1A87';

  // Helper to create label row
  const createRow = (labelKey, value, isStatus = false) => {
    const mainLabel = t.label[labelKey];
    const subLabel = t.label[`subheader_${labelKey}`];
    const labelText = subLabel ? `${mainLabel}\n${subLabel}` : mainLabel;
    
    return {
      type: 'box',
      layout: 'horizontal',
      contents: [{
        type: 'text',
        text: labelText,
        size: 'sm',
        color: '#2F1A87',
        weight: 'bold',
        flex: 0,
        wrap: true
      }, {
        type: 'text',
        text: value,
        size: 'sm',
        color: isStatus ? statusColor : '#111111',
        align: 'end',
        wrap: true
      }]
    };
  };

  const bodyContents = [{
    type: 'text',
    text: headerText,
    weight: 'bold',
    color: statusColor,
    size: 'sm',
    wrap: true
  }, {
    type: 'separator',
    margin: 'md'
  }, {
    type: 'box',
    layout: 'vertical',
    margin: 'md',
    spacing: 'sm',
    contents: [
      createRow('name', nameText),
      {
        type: 'box',
        layout: 'horizontal',
        contents: [{
          type: 'text',
          text: t.label['service'] + (t.label['subheader_service'] ? `\n${t.label['subheader_service']}` : ''),
          size: 'sm',
          color: '#2F1A87',
          weight: 'bold'
        }]
      },
      {
        type: 'box',
        layout: 'vertical',
        margin: 'sm',
        contents: serviceText.split(',').map(item => ({
          type: 'text',
          text: item.trim(),
          size: 'sm',
          color: '#111111',
          wrap: true
        }))
      },
      createRow('date', dateText),
      createRow('time', timeText),
      createRow('status', statusTextValue, true)
    ]
  }];

  const footer = [];
  let footerBg = '#FFFFFF';

  if (type === 'reminder' || type === 'reminder_with_payment' || type === 'reminder_no_payment') {
    const params = `id=${encodeURIComponent(idKeyText)}&name=${encodeURIComponent(nameText)}&service=${encodeURIComponent(serviceText)}&date=${encodeURIComponent(dateText)}&time=${encodeURIComponent(timeText)}`;
    const liffUrl = `https://liff.line.me/${LIFF_ID_CONFIRM}?${params}`; // Note: language might need to be passed here too if frontend supports it

    footer.push({
      type: 'box',
      layout: 'vertical',
      spacing: 'sm',
      contents: [{
        type: 'button',
        action: {
          type: 'uri',
          label: t.footer.manage,
          uri: liffUrl
        },
        style: 'primary',
        height: 'sm',
        color: '#1f1f97'
      },
      {
        type: 'text',
        text: t.footer.cancel_policy,
        size: 'xs',
        color: '#8c8c8c',
        align: 'center',
        wrap: true,
        margin: 'md'
      }]
    });
    footerBg = '#F3F2FA';

  } else if (type === 'cancellation') {
    footer.push({
      type: 'box',
      layout: 'vertical',
      contents: [{
        type: 'text',
        text: t.footer.cancelled,
        size: 'xs',
        color: '#ffffff',
        align: 'center',
        wrap: true
      }]
    });
    footerBg = '#EC726E';

  } else if (type === 'completion') {
    footer.push({
      type: 'box',
      layout: 'vertical',
      contents: [{
        type: 'text',
        text: t.footer.completed,
        size: 'xs',
        color: '#ffffff',
        align: 'center',
        wrap: true
      }]
    });
    footerBg = '#03B555';

  } else if (type === 'confirmation') {
    footer.push({
      type: 'box',
      layout: 'vertical',
      contents: [{
        type: 'text',
        text: t.footer.saved,
        size: 'xs',
        color: '#ffffff',
        align: 'center',
        wrap: true
      }]
    });
    footerBg = '#1f1f97';

  } else if (type === 'confirmBooking') {
    footer.push({
      type: 'box',
      layout: 'vertical',
      contents: [{
        type: 'text',
        text: t.footer.early,
        size: 'xs',
        color: '#ffffff',
        align: 'center',
        wrap: true
      }]
    });
    footerBg = '#06c755';
  }

  if (data.qrCodeUrl && data.liffUrl) {
    bodyContents.push({
      type: 'box',
      layout: 'vertical',
      margin: 'md',
      spacing: 'sm',
      contents: [{
        type: 'image',
        url: data.qrCodeUrl,
        size: 'full',
        aspectMode: 'cover',
        aspectRatio: '1:1'
      }, {
        type: 'button',
        action: {
          type: 'uri',
          label: t.footer.payment,
          uri: data.liffUrl
        },
        style: 'primary',
        color: '#2F1A87',
        margin: 'md'
      }]
    });
  }

  return {
    type: 'flex',
    altText: t.header[type] || 'Appointment Details',
    contents: {
      type: 'bubble',
      body: {
        type: 'box',
        layout: 'vertical',
        contents: bodyContents
      },
      footer: {
        type: 'box',
        layout: 'vertical',
        contents: footer,
        backgroundColor: footerBg
      }
    }
  };
}
// ----------------------------------------------------------------
// ดึงรายการจอง
function fetchBookings(e) {
  var branch = e && e.parameter && e.parameter.branch ? e.parameter.branch : '';
  var sheetName = getSheetNameByBranch(SHEET_BOOKING, branch);
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
  }
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    result.push({
      idKey: r[0],
      userlineid: r[1],
      firstName: r[2],
      lastName: r[3],
      phonenumber: r[4],
      idCardOrSocial: r[5],
      diseaseAllergy: r[6],
      note: r[7],
      service: r[8],
      price: r[9],
      date: r[10],
      time: r[11],
      status: r[12],
      timestamp: r[13],
      calendarEventId: r[14],
      doctor: r[15] || '',
      room: r[16] || '',
      branch: r[17] || ''
    });
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ----------------------------------------------------------------
// Test functions
function testConfirmBooking() {
  const resp = confirmBooking({ id: 2 });
  Logger.log(resp.getContent());
}
function testCancelBooking() {
  const resp = cancelBooking({ id: 1 });
  Logger.log(resp.getContent());
}
function testCompleteBooking() {
  const resp = completeBooking({ parameter: { id: 1 } });
  Logger.log(resp.getContent());
}

// ================================================================
// =================== ฟังก์ชันจัดการข้อมูลหมอ ===================
// ================================================================

// ----------------------------------------------------------------
// ดึงข้อมูลหมอทั้งหมด
function fetchDoctorData(e) {
  var branch = e && e.parameter && e.parameter.branch ? e.parameter.branch : '';
  var sheetName = getSheetNameByBranch(SHEET_DOCTOR, branch);
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
    if (!sheet) {
      // ถ้ายังไม่มีชีต ให้สร้างใหม่
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const newSheet = ss.insertSheet(sheetName);
      // สร้าง Header
      newSheet.getRange(1, 1, 1, 7).setValues([['ID', 'วันที่', 'หมอ 1', 'หมอ 2', 'หมอ 3', 'หมอ 4', 'หมอ 5']]);
      return ContentService.createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }
    const data = sheet.getDataRange().getValues().slice(1); // ข้าม header
    const result = data.map(row => {
      let dateValue = row[1];
      if (dateValue instanceof Date) {
        dateValue = Utilities.formatDate(dateValue, "Asia/Bangkok", "dd-MM-yyyy");
      } else if (typeof dateValue === 'string' && dateValue.includes('-') && dateValue.split('-').length === 3) {
        const parts = dateValue.split('-');
        if (parts[0].length === 4) {
          dateValue = `${parts[2].padStart(2, '0')}-${parts[1].padStart(2, '0')}-${parts[0]}`;
        }
      }
      return {
        id: row[0],
        date: dateValue,
        doctor1: row[2] || '',
        doctor2: row[3] || '',
        doctor3: row[4] || '',
        doctor4: row[5] || '',
        doctor5: row[6] || ''
      };
    }).filter(item => item.id);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
  }
}

// ----------------------------------------------------------------
// ดึงข้อมูลหมอตาม ID
function getDoctorById(id) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_DOCTOR);
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({ error: 'ไม่พบชีตข้อมูลหมอ' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    const data = sheet.getDataRange().getValues();
    const idx = data.findIndex(row => row[0] == id);
    
    if (idx < 0) {
      return ContentService.createTextOutput(JSON.stringify({ error: 'ไม่พบข้อมูล' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    const row = data[idx];
    let dateValue = row[1];
    
    // แปลงวันที่ให้เป็น dd-MM-yyyy
    if (dateValue instanceof Date) {
      dateValue = Utilities.formatDate(dateValue, "Asia/Bangkok", "dd-MM-yyyy");
    } else if (typeof dateValue === 'string' && dateValue.includes('-') && dateValue.split('-').length === 3) {
      const parts = dateValue.split('-');
      if (parts[0].length === 4) {
        // ถ้าเป็น yyyy-MM-dd ให้แปลงเป็น dd-MM-yyyy
        dateValue = `${parts[2].padStart(2, '0')}-${parts[1].padStart(2, '0')}-${parts[0]}`;
      }
    }
    
    const result = {
      id: row[0],
      date: dateValue,
      doctor1: row[2] || '',
      doctor2: row[3] || '',
      doctor3: row[4] || '',
      doctor4: row[5] || '',
      doctor5: row[6] || ''
    };
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    Logger.log('getDoctorById error: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({ error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ----------------------------------------------------------------
// ดึงข้อมูลหมอตามวันที่ (สำหรับใช้ใน dropdown ของหน้า appointment)
function getDoctorsByDate(date, branch) {
  try {
    var sheetName = getSheetNameByBranch(SHEET_DOCTOR, branch || '');
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    const data = sheet.getDataRange().getValues().slice(1);
    const doctors = [];
    
    // แปลงวันที่ให้เป็นรูปแบบเดียวกัน (dd-MM-yyyy)
    let searchDate = date;
    if (date.includes('-') && date.split('-').length === 3) {
      const parts = date.split('-');
      if (parts[0].length === 4) {
        // ถ้าเป็น yyyy-MM-dd หรือ yyyy-M-d ให้แปลงเป็น dd-MM-yyyy
        searchDate = `${parts[2].padStart(2, '0')}-${parts[1].padStart(2, '0')}-${parts[0]}`;
      }
    }
    
    Logger.log('Searching for doctors on date: ' + searchDate);
    
    data.forEach(row => {
      let rowDate = row[1];
      
      // แปลงวันที่ในชีตให้เป็น dd-MM-yyyy
      if (rowDate instanceof Date) {
        rowDate = Utilities.formatDate(rowDate, "Asia/Bangkok", "dd-MM-yyyy");
      } else if (typeof rowDate === 'string') {
        if (rowDate.includes('-') && rowDate.split('-').length === 3) {
          const parts = rowDate.split('-');
          if (parts[0].length === 4) {
            // ถ้าเป็น yyyy-MM-dd ให้แปลงเป็น dd-MM-yyyy
            rowDate = `${parts[2].padStart(2, '0')}-${parts[1].padStart(2, '0')}-${parts[0]}`;
          }
        }
      }
      
      Logger.log(`Comparing: ${rowDate} === ${searchDate}`);
      
      if (rowDate === searchDate) {
        // เก็บชื่อหมอที่ไม่ว่าง
        for (let i = 2; i <= 6; i++) {
          if (row[i]) {
            doctors.push(row[i]);
          }
        }
      }
    });
    
    Logger.log('Found doctors: ' + JSON.stringify(doctors));
    
    return ContentService.createTextOutput(JSON.stringify(doctors))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    Logger.log('getDoctorsByDate error: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ----------------------------------------------------------------
// บันทึกข้อมูลหมอ (เพิ่ม/แก้ไข)
function saveDoctorData(params) {
  var branch = params && params.branch ? params.branch : '';
  var sheetName = getSheetNameByBranch(SHEET_DOCTOR, branch);
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
    if (!sheet) {
      return ContentService.createTextOutput('ไม่พบชีตข้อมูลหมอ');
    }
    
    const id = params.id;
    let date = params.date;
    const doctor1 = params.doctor1 || '';
    const doctor2 = params.doctor2 || '';
    const doctor3 = params.doctor3 || '';
    const doctor4 = params.doctor4 || '';
    const doctor5 = params.doctor5 || '';
    
    // แปลงวันที่ให้เป็นรูปแบบ dd-MM-yyyy
    if (date && date.includes('-') && date.split('-').length === 3) {
      const parts = date.split('-');
      if (parts[0].length === 4) {
        // ถ้าเป็น yyyy-MM-dd ให้แปลงเป็น dd-MM-yyyy
        date = `${parts[2].padStart(2, '0')}-${parts[1].padStart(2, '0')}-${parts[0]}`;
      }
    }
    
    if (id) {
      // แก้ไขข้อมูลเดิม
      const data = sheet.getDataRange().getValues();
      const idx = data.findIndex(row => row[0] == id);
      
      if (idx < 0) {
        return ContentService.createTextOutput('ไม่พบข้อมูลที่ต้องการแก้ไข');
      }
      
      sheet.getRange(idx + 1, 2, 1, 6).setValues([[date, doctor1, doctor2, doctor3, doctor4, doctor5]]);
      sheet.getRange(idx + 1, 2).setNumberFormat("@"); // บังคับให้เป็น Text
      return ContentService.createTextOutput('แก้ไขข้อมูลสำเร็จ');
    } else {
      // เพิ่มข้อมูลใหม่
      const nextRow = sheet.getLastRow() + 1;
      const newId = 'DOC' + (nextRow - 1);
      sheet.getRange(nextRow, 1, 1, 7).setValues([[newId, date, doctor1, doctor2, doctor3, doctor4, doctor5]]);
      sheet.getRange(nextRow, 2).setNumberFormat("@"); // บังคับให้เป็น Text
      return ContentService.createTextOutput('เพิ่มข้อมูลสำเร็จ');
    }
  } catch (error) {
    Logger.log('saveDoctorData error: ' + error.toString());
    return ContentService.createTextOutput('เกิดข้อผิดพลาด: ' + error.toString());
  }
}

// ----------------------------------------------------------------
// ลบข้อมูลหมอ
function deleteDoctorData(params) {
  var branch = params && params.branch ? params.branch : '';
  var sheetName = getSheetNameByBranch(SHEET_DOCTOR, branch);
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
    if (!sheet) {
      return ContentService.createTextOutput('ไม่พบชีตข้อมูลหมอ');
    }
    
    const id = params.id;
    const data = sheet.getDataRange().getValues();
    const idx = data.findIndex(row => row[0] == id);
    
    if (idx < 0) {
      return ContentService.createTextOutput('ไม่พบข้อมูลที่ต้องการลบ');
    }
    
    sheet.deleteRow(idx + 1);
    return ContentService.createTextOutput('ลบข้อมูลสำเร็จ');
  } catch (error) {
    Logger.log('deleteDoctorData error: ' + error.toString());
    return ContentService.createTextOutput('เกิดข้อผิดพลาด: ' + error.toString());
  }
}

// ----------------------------------------------------------------
// อัปเดตข้อมูลหมอและห้องในการนัดหมาย
function updateAppointmentDoctor(params) {
  var branch = params && params.branch ? params.branch : '';
  var sheetName = getSheetNameByBranch(SHEET_BOOKING, branch);
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
    const id = params.id;
    const doctor = params.doctor || '';
    const room = params.room || '';
    const data = sheet.getDataRange().getValues();
    const idx = data.findIndex(row => row[0] == id);
    if (idx < 0) {
      return ContentService.createTextOutput('ไม่พบการนัดหมายนี้');
    }
    // Column P (16) = หมอ, Column Q (17) = ห้อง
    sheet.getRange(idx + 1, 16).setValue(doctor);
    sheet.getRange(idx + 1, 17).setValue(room);
    return ContentService.createTextOutput('อัปเดตข้อมูลสำเร็จ');
  } catch (error) {
    Logger.log('updateAppointmentDoctor error: ' + error.toString());
    return ContentService.createTextOutput('เกิดข้อผิดพลาด: ' + error.toString());
  }
}

// ----------------------------------------------------------------
// ดึงรายชื่อหมอจากชีต "ตั้งค่าทั่วไป" คอลัมน์ F
function fetchDoctorNames(e) {
  var branch = e && e.parameter && e.parameter.branch ? e.parameter.branch : '';
  var sheetName = getSheetNameByBranch(SHEET_DATETIME, branch);
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }
    const data = sheet.getDataRange().getValues();
    const doctorNames = [];
    for (let i = 1; i < data.length; i++) {
      const doctorName = data[i][5];
      if (doctorName && doctorName.toString().trim() !== '') {
        doctorNames.push(doctorName.toString().trim());
      }
    }
    const uniqueDoctors = [...new Set(doctorNames)];
    Logger.log('Found doctors: ' + JSON.stringify(uniqueDoctors));
    return ContentService.createTextOutput(JSON.stringify(uniqueDoctors))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    Logger.log('fetchDoctorNames error: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ===================== Workday Mode (Default Closed) =====================
function parseWeeklyWorkdaysV2(rawValue) {
  return String(rawValue || '')
    .split(',')
    .map((d) => parseInt(String(d).trim(), 10))
    .filter((d) => !isNaN(d) && d >= 0 && d <= 6);
}

function toYmdFromSheetDateV2(value) {
  if (!value) return null;
  if (value instanceof Date) {
    return Utilities.formatDate(value, 'Asia/Bangkok', 'yyyy-MM-dd');
  }
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) return null;
    if (/^\d{2}-\d{2}-\d{4}$/.test(trimmed)) {
      const [dd, mm, yyyy] = trimmed.split('-');
      return `${yyyy}-${mm}-${dd}`;
    }
    if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
      return trimmed;
    }
  }
  return null;
}

function extractWorkdayConfigV2(rows) {
  const safeRows = Array.isArray(rows) ? rows : [];
  const weeklyWorkdays = parseWeeklyWorkdaysV2(
    safeRows.length > 1 && safeRows[1] ? safeRows[1][4] : '',
  );
  const specialWorkdays = Array.from(
    new Set(
      safeRows
        .slice(1)
        .map((r) => toYmdFromSheetDateV2(r ? r[3] : null))
        .filter(Boolean),
    ),
  );
  return { weeklyWorkdays, specialWorkdays };
}

function extractTimeCapacityV2(rows) {
  const dataRows = Array.isArray(rows) ? rows.slice(1) : [];
  const times = [];
  const maxBookings = [];
  dataRows.forEach((r) => {
    if (!r) return;
    const t = r[1];
    if (!t) return;
    let timeStr = '';
    if (t instanceof Date) {
      timeStr = Utilities.formatDate(t, 'Asia/Bangkok', 'HH:mm');
    } else if (typeof t === 'string' && /^\d{2}:\d{2}:\d{2}$/.test(t)) {
      timeStr = t.substring(0, 5);
    } else {
      timeStr = String(t);
    }
    if (!timeStr) return;
    times.push(timeStr);
    maxBookings.push(Number(r[2]) || 0);
  });
  return { times, maxBookings };
}

function isWorkingDateV2(dateObj, weeklyWorkdays, specialWorkdays) {
  const dayOfWeek = dateObj.getDay();
  const dateKey = Utilities.formatDate(dateObj, 'Asia/Bangkok', 'yyyy-MM-dd');
  return weeklyWorkdays.includes(dayOfWeek) || specialWorkdays.includes(dateKey);
}

function buildDayAvailabilityV2(dateObj, times, maxBookings, branch) {
  const dayStart = new Date(dateObj);
  dayStart.setHours(0, 0, 0, 0);
  const dayEnd = new Date(dayStart);
  dayEnd.setDate(dayEnd.getDate() + 1);
  const calendar = CalendarApp.getCalendarById(getCalendarIdByBranch(branch));
  const events = calendar
    .getEvents(dayStart, dayEnd)
    .filter((ev) => ev.getTitle().indexOf('นัดคิว:') > -1);

  const bookedByTime = {};
  events.forEach((ev) => {
    const t = Utilities.formatDate(ev.getStartTime(), 'Asia/Bangkok', 'HH:mm');
    bookedByTime[t] = (bookedByTime[t] || 0) + 1;
  });

  const availability = {};
  times.forEach((t, i) => {
    const total = maxBookings[i] || 0;
    const booked = bookedByTime[t] || 0;
    const remaining = Math.max(0, total - booked);
    availability[t] = {
      status: remaining > 0 ? 'ว่าง' : 'เต็ม',
      remaining,
      total,
    };
  });
  return availability;
}

// Override: monthly availability used by calendar day disable logic
function getMonthAvailability(e) {
  const branch = e && e.parameter && e.parameter.branch ? e.parameter.branch : '';
  const year = Number(e && e.parameter ? e.parameter.year : 0);
  const month = Number(e && e.parameter ? e.parameter.month : 0);
  if (!year || !month || month < 1 || month > 12) {
    return ContentService.createTextOutput(JSON.stringify({}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const sheetName = getSheetNameByBranch(SHEET_DATETIME, branch);
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const rows = sheet.getDataRange().getValues();
  const { weeklyWorkdays, specialWorkdays } = extractWorkdayConfigV2(rows);
  const { times, maxBookings } = extractTimeCapacityV2(rows);
  const calendar = CalendarApp.getCalendarById(getCalendarIdByBranch(branch));

  const monthStart = new Date(year, month - 1, 1);
  monthStart.setHours(0, 0, 0, 0);
  const monthEnd = new Date(year, month, 1);
  monthEnd.setHours(0, 0, 0, 0);

  const events = calendar
    .getEvents(monthStart, monthEnd)
    .filter((ev) => ev.getTitle().indexOf('นัดคิว:') > -1);
  const bookedByDateTime = {};
  events.forEach((ev) => {
    const start = ev.getStartTime();
    const dateKey = Utilities.formatDate(start, 'Asia/Bangkok', 'yyyy-MM-dd');
    const timeKey = Utilities.formatDate(start, 'Asia/Bangkok', 'HH:mm');
    if (!bookedByDateTime[dateKey]) bookedByDateTime[dateKey] = {};
    bookedByDateTime[dateKey][timeKey] = (bookedByDateTime[dateKey][timeKey] || 0) + 1;
  });

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const daysInMonth = new Date(year, month, 0).getDate();
  const result = {};

  for (let day = 1; day <= daysInMonth; day++) {
    const currentDate = new Date(year, month - 1, day);
    currentDate.setHours(0, 0, 0, 0);
    const dateKey = Utilities.formatDate(currentDate, 'Asia/Bangkok', 'yyyy-MM-dd');

    if (currentDate <= today) {
      result[dateKey] = 'full';
      continue;
    }

    const isWorkingDay = isWorkingDateV2(currentDate, weeklyWorkdays, specialWorkdays);
    if (!isWorkingDay || !times.length) {
      result[dateKey] = 'full';
      continue;
    }

    const bookedTimeMap = bookedByDateTime[dateKey] || {};
    let hasAvailable = false;
    for (let i = 0; i < times.length; i++) {
      const total = maxBookings[i] || 0;
      const booked = bookedTimeMap[times[i]] || 0;
      if (total > 0 && booked < total) {
        hasAvailable = true;
        break;
      }
    }
    result[dateKey] = hasAvailable ? 'available' : 'full';
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// Override: day availability based on workday model (default closed)
function checkAvailability(e) {
  if (!e.parameter.date && !e.parameter.getHolidays && !e.parameter.getWorkdays) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'Missing params' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const branch = e && e.parameter && e.parameter.branch ? e.parameter.branch : '';
  const sheetName = getSheetNameByBranch(SHEET_DATETIME, branch);
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet) {
    if (e.parameter.getHolidays || e.parameter.getWorkdays) {
      return ContentService.createTextOutput(
        JSON.stringify({
          workdays: [],
          weeklyWorkdays: [],
          holidays: [],
          permanentHolidays: [],
        }),
      ).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({ availability: {} }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const rows = sheet.getDataRange().getValues();
  const { weeklyWorkdays, specialWorkdays } = extractWorkdayConfigV2(rows);
  if (e.parameter.getHolidays || e.parameter.getWorkdays) {
    return ContentService.createTextOutput(
      JSON.stringify({
        workdays: specialWorkdays,
        weeklyWorkdays,
        // Backward-compatible fields
        holidays: specialWorkdays,
        permanentHolidays: weeklyWorkdays,
      }),
    ).setMimeType(ContentService.MimeType.JSON);
  }

  const inputDate = new Date(e.parameter.date);
  if (isNaN(inputDate)) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'Invalid date' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const { times, maxBookings } = extractTimeCapacityV2(rows);
  const isWorkingDay = isWorkingDateV2(inputDate, weeklyWorkdays, specialWorkdays);

  let availability = {};
  if (inputDate <= today) {
    times.forEach((t) => {
      availability[t] = { status: 'เต็ม', remaining: 0, total: 0 };
    });
  } else if (!isWorkingDay) {
    times.forEach((t) => {
      availability[t] = { status: 'ปิด', remaining: 0, total: 0 };
    });
  } else {
    availability = buildDayAvailabilityV2(inputDate, times, maxBookings, branch);
  }

  return ContentService.createTextOutput(
    JSON.stringify({
      availability,
      isWorkingDay,
      isHoliday: !isWorkingDay,
      color: isWorkingDay ? 'default' : 'red',
    }),
  ).setMimeType(ContentService.MimeType.JSON);
}


