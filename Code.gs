// --- 1. การตั้งค่า (CONFIGURATION) ---
const SPREADSHEET_ID = '1Kr1GOn5F8rBNJGA7_Sqp4be7JirrRvVci0AGhtkA5hQ'; // ID ของ Google Sheet
const FOLDER_ID = '1pPtZlI8XYBle02byB5lthAhtLX8012Pa'; // ID ของ Google Drive Folder สำหรับเก็บรูป

const SHEET_NAMES = {
  STUDENTS: ['รายชื่อนักเรียน ม.3', 'รายชื่อนักเรียน ม.4', 'รายชื่อนักเรียน ม.5', 'รายชื่อนักเรียน ม.6'],
  TEACHERS: 'รายชื่อครู',
  ASSETS: 'รายงานทะเบียนทรัพย์สิน',
  DATA_DB: 'ข้อมูล',
  LOGS: 'Log',
  ADMIN: 'แอดมิน',
  ADVISOR: 'ครูที่ปรึกษา'
};

// --- 2. ฟังก์ชันพื้นฐาน (CORE FUNCTIONS) ---

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('ระบบตรวจสอบสถานะ iPad โรงเรียนอรัญประเทศ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) { 
  return HtmlService.createHtmlOutputFromFile(filename).getContent(); 
}

// *** ฟังก์ชันจัดระเบียบชื่อ (แก้ปัญหาชื่อไม่ตรง) ***
// ตัดคำนำหน้าและช่องว่างออกทั้งหมด เพื่อให้เทียบชื่อได้แม่นยำ
function normalizeName(name) {
  if (!name) return "";
  let n = name.toString().trim();
  
  // 1. ลบคำนำหน้า (ไทย/อังกฤษ)
  n = n.replace(/^(เด็กชาย|เด็กหญิง|ด\.ช\.|ด\.ญ\.|นาย|นางสาว|นาง|น\.ส\.|ว่าที่ร้อยตรี|ว่าที่ ร\.ต\.|ครู|อ\.|Miss|Mr\.|Mrs\.|Ms\.)\s*/g, '');
  
  // 2. ลบช่องว่างทั้งหมด (Space)
  n = n.replace(/\s/g, ''); 
  
  return n;
}

// --- 3. ฟังก์ชันดึงข้อมูลทั้งหมด (MAIN DATA RETRIEVAL) ---

function getAllSystemData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // A. ดึงข้อมูลจากชีต "รายงานทะเบียนทรัพย์สิน" (Master Asset)
  const assetSheet = ss.getSheetByName(SHEET_NAMES.ASSETS);
  const assetData = assetSheet ? assetSheet.getDataRange().getDisplayValues() : [];
  let assetMap = {};
  
  // เริ่มวนลูปตั้งแต่แถวที่ 2 (index 1)
  for (let i = 1; i < assetData.length; i++) {
    // *** สำคัญ: ตรวจสอบตำแหน่งคอลัมน์ตรงนี้ ***
    // [0]=A, [1]=B, [2]=C, [3]=D, [4]=E
    let serial = assetData[i][2]; // สมมติ Serial อยู่คอลัมน์ C
    let rawName = assetData[i][4]; // สมมติ ชื่อ-สกุล อยู่คอลัมน์ E
    
    if(rawName) {
      let cName = normalizeName(rawName); // ล้างชื่อให้สะอาด
      if(cName.length > 0) {
        assetMap[cName] = { 
          serial: serial, 
          status: 'ยืมอยู่' // ค่าตั้งต้นถ้ามีชื่อในชีตนี้
        };
      }
    }
  }

  // B. ดึงข้อมูลจากชีต "ข้อมูล" (Transaction Logs)
  const dbSheet = ss.getSheetByName(SHEET_NAMES.DATA_DB);
  const dbData = dbSheet ? dbSheet.getDataRange().getDisplayValues() : [];
  let dbMap = {}; 
  
  for (let i = 1; i < dbData.length; i++) {
    let id = dbData[i][1];
    let rowSerial = dbData[i][5];
    let rowStatus = dbData[i][8]; 
    // เช็คว่ามีไฟล์แนบไหม
    let hasFiles = (dbData[i][9]!="" || dbData[i][10]!="" || dbData[i][11]!="" || dbData[i][12]!="");
    
    if (id) {
      // สร้าง Object ถ้ายังไม่มี
      if (!dbMap[id]) {
        dbMap[id] = { borrowStatus: 'ยังไม่ยืม', docStatus: 'ยังไม่ส่ง', serial: '-', files: {} };
      }

      // อัปเดต Serial ถ้ามี
      if (rowSerial && rowSerial !== '-' && rowSerial !== '') dbMap[id].serial = rowSerial;
      
      // อัปเดตสถานะเอกสาร
      if (hasFiles) {
        dbMap[id].files = { 
          agreement: dbData[i][9], card_std: dbData[i][10], 
          card_parent: dbData[i][11], house: dbData[i][12], phone: dbData[i][13] 
        };
        // ถ้าสถานะไม่ใช่ Admin แก้ไข ให้ถือว่ารอตรวจสอบ
        if (!rowStatus.includes('ADMIN') && !rowStatus.includes('ADVISOR')) dbMap[id].docStatus = 'รอตรวจสอบ';
      }

      if (rowStatus.includes('เอกสารผ่าน')) dbMap[id].docStatus = 'เอกสารผ่าน';
      else if (rowStatus.includes('ไม่ผ่าน')) dbMap[id].docStatus = 'เอกสารไม่ผ่าน';
      else if (rowStatus.includes('รอตรวจสอบ')) dbMap[id].docStatus = 'รอตรวจสอบ';

      // อัปเดตสถานะเครื่อง (เอาบรรทัดล่างสุดเสมอ)
      if (rowStatus.includes('ยืมอยู่') || rowStatus === 'ยืมได้') dbMap[id].borrowStatus = 'ยืมอยู่';
      else if (rowStatus.includes('คืน')) dbMap[id].borrowStatus = 'คืนแล้ว';
      else if (rowStatus.includes('ซ่อม')) dbMap[id].borrowStatus = 'ส่งซ่อม';
      else if (rowStatus.includes('สละ')) dbMap[id].borrowStatus = 'สละสิทธิ์';
      else if (rowStatus === 'ยังไม่ยืม') dbMap[id].borrowStatus = 'ยังไม่ยืม';
    }
  }

  // C. รวมข้อมูลทั้งหมด (นักเรียน + ทรัพย์สิน + Log)
  let allPeople = [];
  
  const processPerson = (type, no, id, name, room, source) => {
    let cleanedName = normalizeName(name); // ล้างชื่อ
    
    let finalBorrow = 'ยังไม่ยืม';
    let finalDoc = 'ยังไม่ส่ง';
    let finalSerial = '-';
    let finalFiles = {};
    let isInAssetSheet = false; // ตัวแปรเช็คว่าตกหล่นไหม

    // 1. เช็คกับชีตทรัพย์สิน (Master Asset)
    if (assetMap[cleanedName]) { 
      finalBorrow = assetMap[cleanedName].status; 
      finalSerial = assetMap[cleanedName].serial;
      isInAssetSheet = true; // เจอชื่อในทะเบียนทรัพย์สิน
    }

    // 2. เช็คกับชีตข้อมูล (Logs) - ข้อมูลล่าสุดจะ Override
    if (dbMap[id]) {
      if(dbMap[id].borrowStatus !== 'ยังไม่ยืม') finalBorrow = dbMap[id].borrowStatus;
      // ถ้า DB บอกยังไม่ยืม แต่ใน Asset มีชื่อ -> ให้เชื่อ Asset
      else if (finalBorrow === 'ยังไม่ยืม' && assetMap[cleanedName]) finalBorrow = assetMap[cleanedName].status;
      
      finalDoc = dbMap[id].docStatus;
      if(dbMap[id].serial !== '-') finalSerial = dbMap[id].serial;
      finalFiles = dbMap[id].files;
    }

    allPeople.push({ 
      type: type, 
      no: no, 
      id: id, 
      name: name, 
      room: room, 
      source_sheet: source, 
      serial: finalSerial, 
      borrowStatus: finalBorrow, 
      docStatus: finalDoc, 
      files: finalFiles,
      inAsset: isInAssetSheet // ส่งค่า true/false ไปหน้าบ้าน
    });
  };

  // วนลูปรายชื่อนักเรียนทุกชีต
  SHEET_NAMES.STUDENTS.forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    if (sheet) { 
      let data = sheet.getDataRange().getDisplayValues(); 
      for (let i = 1; i < data.length; i++) { 
        if(data[i][2]) { // ถ้ามีชื่อ
          processPerson('student', data[i][0], data[i][1], data[i][2], data[i][3], sheetName); 
        }
      } 
    }
  });

  // วนลูปรายชื่อครู
  let teacherSheet = ss.getSheetByName(SHEET_NAMES.TEACHERS);
  if (teacherSheet) { 
    let tData = teacherSheet.getDataRange().getDisplayValues(); 
    for (let i = 1; i < tData.length; i++) { 
      if(tData[i][1]) {
        processPerson('teacher', tData[i][0], 'T-'+tData[i][0], tData[i][1], 'ห้องพักครู', SHEET_NAMES.TEACHERS); 
      }
    } 
  }

  return allPeople;
}

// --- 4. ฟังก์ชันจัดการ Form (FORM HANDLING) ---

function processForm(formObject) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetData = ss.getSheetByName(SHEET_NAMES.DATA_DB);
  const folder = DriveApp.getFolderById(FOLDER_ID);
  
  try {
    const timestamp = new Date();
    
    // Helper Upload File
    const uploadFile = (fileBlob, prefix) => {
      if (!fileBlob || fileBlob.name == "") return "";
      let fileName = prefix + "_" + formObject.userName + "_" + timestamp.getTime();
      return folder.createFile(fileBlob).setName(fileName).getUrl();
    };

    let url_agreement = uploadFile(formObject.file_agreement, "AGREEMENT");
    let url_card_std = "", url_card_parent = "", url_house = "", parent_phone = "";

    if (formObject.userType === 'student') {
      url_card_std = uploadFile(formObject.file_card_std, "CARD_STD");
      url_card_parent = uploadFile(formObject.file_card_parent, "CARD_PARENT");
      url_house = uploadFile(formObject.file_house, "HOUSE");
      parent_phone = "'" + formObject.parent_phone;
    }

    let statusToSave = formObject.statusSelect;
    if(url_agreement !== "") statusToSave = statusToSave + " | รอตรวจสอบเอกสาร"; 

    sheetData.appendRow([
      timestamp, 
      formObject.userId, 
      formObject.userName, 
      formObject.userType, 
      formObject.userRoom, 
      formObject.userSerial, 
      "USER_UPDATE", 
      formObject.note || "", 
      statusToSave, 
      url_agreement, url_card_std, url_card_parent, url_house, parent_phone
    ]);
    
    return { success: true, message: "บันทึกข้อมูลเรียบร้อย" };
  } catch (error) { 
    return { success: false, message: "เกิดข้อผิดพลาด: " + error.toString() }; 
  }
}

// --- 5. ฟังก์ชันตรวจสอบสิทธิ์ (AUTHENTICATION) ---

function verifyAdmin(u, p) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.ADMIN);
  if (!sheet) return { success: false, message: "No Admin Sheet" };
  
  const data = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(u).trim() && String(data[i][1]).trim() === String(p).trim()) {
      return { success: true, role: 'admin' };
    }
  }
  return { success: false, message: "Login Failed" };
}

function verifyAdvisor(u, p) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.ADVISOR);
  if (!sheet) return { success: false, message: "ไม่พบชีตครูที่ปรึกษา" };
  
  const data = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(u).trim() && String(data[i][1]).trim() === String(p).trim()) {
      return { 
        success: true, 
        role: 'advisor', 
        level: data[i][2], // ระดับชั้น
        room: data[i][3],  // ห้อง
        name: data[i][4] || "คุณครูที่ปรึกษา" 
      };
    }
  }
  return { success: false, message: "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง" };
}

// --- 6. ฟังก์ชันจัดการของแอดมิน (ADMIN ACTIONS) ---

function adminUpdateData(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetData = ss.getSheetByName(SHEET_NAMES.DATA_DB);
  
  try {
    const timestamp = new Date();
    // ระบุตัวตนคนแก้ไข
    let editor = data.editorRole === 'advisor' ? ("ADVISOR: " + data.editorName) : "ADMIN_EDIT";

    // กรณีเปลี่ยนสถานะเครื่อง
    if (data.borrowStatusSelect) {
      sheetData.appendRow([
        timestamp, data.userId, data.userName, data.userType, data.userRoom, data.userSerial, 
        editor, data.note, data.borrowStatusSelect, "", "", "", "", ""
      ]);
    }
    
    // กรณีเปลี่ยนสถานะเอกสาร
    if (data.docStatusSelect && data.docStatusSelect !== "") {
       Utilities.sleep(100); 
       sheetData.appendRow([
         new Date(), data.userId, data.userName, data.userType, data.userRoom, data.userSerial, 
         editor, data.note, data.docStatusSelect, "", "", "", "", ""
       ]);
    }
    return { success: true, message: "อัปเดตข้อมูลสำเร็จ" };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function adminDeleteUser(data) {
  // ครูที่ปรึกษา ห้ามลบ
  if(data.editorRole === 'advisor') return { success: false, message: "ครูที่ปรึกษาไม่ได้รับอนุญาตให้ลบข้อมูล" };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(data.source_sheet);
  if (!sheet) return { success: false, message: "ไม่พบแผ่นงาน" };
  
  try {
    const rows = sheet.getDataRange().getDisplayValues();
    let rowToDelete = -1;
    
    for (let i = 0; i < rows.length; i++) {
      if (data.source_sheet === SHEET_NAMES.TEACHERS) { 
        if (rows[i][1] == data.name) { rowToDelete = i + 1; break; } 
      } else { 
        if (rows[i][1] == data.id) { rowToDelete = i + 1; break; } 
      }
    }
    
    if (rowToDelete > -1) { 
      sheet.deleteRow(rowToDelete); 
      return { success: true, message: "ลบข้อมูลเรียบร้อยแล้ว" }; 
    } else { 
      return { success: false, message: "ไม่พบข้อมูล" }; 
    }
  } catch (e) { return { success: false, message: "Error: " + e.toString() }; }
}

function adminAddUser(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(data.targetSheet); 
  if (!sheet) return { success: false, message: "ไม่พบแผ่นงาน" };
  
  try {
    const nextNo = sheet.getLastRow(); // เลขที่แบบง่าย
    
    if (data.targetSheet === SHEET_NAMES.TEACHERS) {
      sheet.appendRow([nextNo, data.name]);
    } else {
      sheet.appendRow([nextNo, data.id, data.name, data.room]);
    }
    return { success: true, message: "เพิ่มรายชื่อเรียบร้อยแล้ว" };
  } catch (e) { return { success: false, message: "Error: " + e.toString() }; }
}
