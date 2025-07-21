// =================================================================
// == การตั้งค่าส่วนกลาง (Global Settings)
// =================================================================

// 🔥 สำคัญ: ใส่ ID ของ Google Sheet ของคุณตรงนี้
const SPREADSHEET_ID = "14lmzOv4-5iCKRTiWw3jaX93OuNDfRaXDoShkHBGOHzE"; 

// ทำการเชื่อมต่อกับ Google Sheet และชีทที่ต้องใช้บ่อยๆ
const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
const usersSheet = ss.getSheetByName("Users");
// (ชีทอื่นๆ จะถูกเรียกใช้ในฟังก์ชันที่เกี่ยวข้อง)


// =================================================================
// == ฟังก์ชันหลักในการแสดงผลหน้าเว็บ (Web Page Routing)
// =================================================================

/**
 * ฟังก์ชันนี้จะทำงานเป็นอันดับแรกและประมวลผลไฟล์ HTML ก่อนส่งไปแสดงผล
 * เพื่อให้โค้ดพิเศษ (Scriptlets) ใน HTML ทำงานได้
 */
function doGet(e) {
  // หน้าเริ่มต้นคือ Login เสมอ เราจะใช้การเปลี่ยนเนื้อหาแทนการ Redirect
  // return HtmlService.createTemplateFromFile('Login')
  //     .evaluate()
  //     .setTitle("ระบบสอบออนไลน์");

  // ใช้ e.parameter.page เพื่อดูว่าผู้ใช้ขอหน้าอะไรมา, ถ้าไม่ระบุให้ไปหน้า Index
  const page = e.parameter.page || 'Login'; 
  let htmlOutput;

  // ตรวจสอบว่ามี case ครบทุกหน้าหรือไม่
  switch (page) {
    case 'Login':
      htmlOutput = HtmlService.createTemplateFromFile('Login').evaluate();
      break;
    case 'Register':
      htmlOutput = HtmlService.createTemplateFromFile('Register').evaluate();
      break;
    case 'AdminDashboard':
      htmlOutput = HtmlService.createTemplateFromFile('AdminDashboard').evaluate();
      break;
    case 'TeacherDashboard':
      htmlOutput = HtmlService.createTemplateFromFile('TeacherDashboard').evaluate();
      break;
    case 'StudentDashboard':
      htmlOutput = HtmlService.createTemplateFromFile('StudentDashboard').evaluate();
      break;
    case 'Quiz':
      htmlOutput = HtmlService.createTemplateFromFile('Quiz').evaluate();
      break;
    case 'Report':
      htmlOutput = HtmlService.createTemplateFromFile('Report').evaluate();
      break;
    default:
      // ถ้าขอหน้าที่ไม่มีในระบบ ให้ส่งกลับไปหน้าแรกเสมอ
      htmlOutput = HtmlService.createTemplateFromFile('Login').evaluate();
  }
  
  // ตั้งค่าหัวข้อเว็บและอนุญาตให้ฝังใน Google Sites
  return htmlOutput.setTitle('ระบบส่งแผนการสอนออนไลน์').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);


}

/**
 * ฟังก์ชันสำหรับดึงเนื้อหา HTML จากไฟล์อื่นมาเป็นข้อความ
 * (สำหรับใช้กับเทคนิค Single-Page Application)
 * @param {string} filename - ชื่อไฟล์ HTML ที่ไม่รวม .html
 * @returns {string} - เนื้อหา HTML ของไฟล์นั้น
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// =================================================================
// == ฟังก์ชันเกี่ยวกับการจัดการผู้ใช้ (User Management)
// =================================================================

/**
 * ตรวจสอบชื่อผู้ใช้และรหัสผ่านเพื่อเข้าสู่ระบบ
 */
function authenticateUser(username, password) {
  try {
    const lastRow = usersSheet.getLastRow();
    if (lastRow < 2) { return { success: false, message: "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง" }; }
    
    const data = usersSheet.getRange(2, 1, lastRow - 1, 7).getValues();
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row[0] === username && row[1] === password) {
        const userStatus = row[6]; // Column G: Status
        if (userStatus === "Active") {
          return { success: true, role: row[4], message: "เข้าสู่ระบบสำเร็จ!" };
        } else if (userStatus === "Pending") {
          return { success: false, message: "บัญชีของคุณยังรอการอนุมัติ" };
        } else {
          return { success: false, message: "การลงทะเบียนของคุณถูกปฏิเสธ" };
        }
      }
    }
    return { success: false, message: "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง" };
  } catch (error) {
    Logger.log("Error in authenticateUser: " + error.message);
    return { success: false, message: "เกิดข้อผิดพลาดฝั่งเซิร์ฟเวอร์" };
  }
}

/**
 * ประมวลผลข้อมูลการลงทะเบียนที่ส่งมาจากฟอร์มหน้าเว็บ
 */
function processRegistration(formData) {
  try {
    const lastRow = usersSheet.getLastRow();
    if (lastRow > 1) {
      const usersData = usersSheet.getRange(2, 1, lastRow - 1, 6).getValues();
      if (usersData.some(row => row[0].toLowerCase() === formData.username.toLowerCase())) {
        return { success: false, message: "ชื่อผู้ใช้นี้ถูกใช้งานแล้ว" };
      }
      if (usersData.some(row => row[5].toLowerCase() === formData.email.toLowerCase())) {
        return { success: false, message: "อีเมลนี้ถูกใช้งานแล้ว" };
      }
    }

    let role = "student";
    let status = "Pending";

    if (lastRow === 1) {
      role = "admin";
      status = "Active";
    }

    usersSheet.appendRow([
      formData.username, formData.password, formData.name,
      formData.class, role, formData.email, status, "" // DriveFolderID
    ]);
    
    // TODO: ส่งอีเมลแจ้ง Admin ว่ามีคนสมัครใหม่

    if (status === "Pending") {
      return { success: true, message: "ลงทะเบียนสำเร็จ! บัญชีของคุณกำลังรอการอนุมัติ" };
    } else {
      return { success: true, message: "ลงทะเบียนสำเร็จ! คุณคือผู้ดูแลระบบคนแรก" };
    }
  } catch (error) {
    Logger.log("Error in processRegistration: " + error.message);
    return { success: false, message: "เกิดข้อผิดพลาดในการลงทะเบียน" };
  }
}


// =================================================================
// == ฟังก์ชันสำหรับ Admin (Admin Functions)
// =================================================================

/**
 * (สำหรับ Admin) ดึงรายชื่อผู้ใช้ที่รอการอนุมัติ
 */
function getPendingUsers() {
  try {
    const data = usersSheet.getDataRange().getValues();
    const headers = data.shift();
    
    const pendingUsers = data
      .filter(row => row[6] === 'Pending') // คอลัมน์ G คือ Status
      .map(row => ({
        username: row[0],
        name: row[2],
        email: row[5],
        class: row[3]
      }));
    return pendingUsers;
  } catch (e) {
    Logger.log(e);
    return [];
  }
}

/**
 * (สำหรับ Admin) อนุมัติผู้ใช้งาน
 */
function approveUser(username) {
  try {
    const data = usersSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === username) {
        usersSheet.getRange(i + 1, 7).setValue('Active');
        // TODO: ส่งอีเมลแจ้งผู้ใช้ว่าบัญชีถูกอนุมัติแล้ว
        return { success: true };
      }
    }
    return { success: false, message: 'User not found' };
  } catch (e) {
    Logger.log(e);
    return { success: false, message: e.message };
  }
}

/**
 * (สำหรับ Admin) ปฏิเสธผู้ใช้งาน
 */
function rejectUser(username) {
  try {
    const data = usersSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === username) {
        usersSheet.getRange(i + 1, 7).setValue('Rejected');
        // TODO: ส่งอีเมลแจ้งผู้ใช้ว่าบัญชีถูกปฏิเสธ
        return { success: true };
      }
    }
    return { success: false, message: 'User not found' };
  } catch (e) {
    Logger.log(e);
    return { success: false, message: e.message };
  }
}

/**
 * (สำหรับ Admin) สร้างไฟล์สำรองข้อมูล
 */
function backupData() {
    // TODO: สร้างฟังก์ชันเพื่อคัดลอกไฟล์ Google Sheet ทั้งหมด
    // ไปเก็บไว้ในโฟลเดอร์สำรองข้อมูลใน Google Drive
    // พร้อมตั้งชื่อไฟล์ด้วยวันที่และเวลา
    return { success: true, message: "สำรองข้อมูลสำเร็จ (ยังไม่สร้างไฟล์จริง)" };
}

/**
 * (สำหรับ Admin) แสดงรายการไฟล์ที่สำรองไว้
 */
function getBackupFiles() {
    // TODO: สร้างฟังก์ชันเพื่ออ่านรายชื่อไฟล์ในโฟลเดอร์สำรองข้อมูล
    // แล้ว return กลับไปเป็น Array
    return [];
}

/**
 * (สำหรับ Admin) กู้คืนข้อมูลจากไฟล์ที่เลือก
 */
function restoreData(fileId) {
    // TODO: สร้างฟังก์ชันที่ซับซ้อนและต้องใช้ความระมัดระวัง
    // เพื่อคัดลอกข้อมูลจากไฟล์สำรองมาทับไฟล์ปัจจุบัน
    return { success: true, message: "กู้คืนข้อมูลสำเร็จ (ยังไม่ทำงานจริง)" };
}
