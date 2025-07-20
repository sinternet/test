// =================================================================
// == การตั้งค่าส่วนกลาง (Global Settings)
// =================================================================

// 🔥 สำคัญ: ใส่ ID ของ Google Sheet ของคุณตรงนี้
const SPREADSHEET_ID = "14lmzOv4-5iCKRTiWw3jaX93OuNDfRaXDoShkHBGOHzE"; 

// ทำการเชื่อมต่อกับ Google Sheet และชีทที่ต้องใช้บ่อยๆ
const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
const usersSheet = ss.getSheetByName("Users");
const subjectsSheet = ss.getSheetByName("Subjects");
const quizzesSheet = ss.getSheetByName("Quizzes");
// (ชีทอื่นๆ จะถูกเรียกใช้ในฟังก์ชันที่เกี่ยวข้องต่อไป)


// =================================================================
// == ฟังก์ชันหลักในการแสดงผลหน้าเว็บ (Web Page Routing)
// =================================================================

/**
 * ฟังก์ชันนี้จะทำงานเป็นอันดับแรกเมื่อมีคนเปิด URL ของเว็บแอป
 * และจะตรวจสอบ URL parameter เพื่อตัดสินใจว่าจะแสดงหน้าเว็บใด
 */
/**
 * ฟังก์ชันนี้จะทำงานเป็นอันดับแรกและประมวลผลไฟล์ HTML ก่อนส่งไปแสดงผล
 */
function doGet(e) {
  let pageTitle = "LMS";
  let template;

  switch (e.parameter.page) {
    case 'register':
      template = HtmlService.createTemplateFromFile('Register');
      pageTitle = "LMS - ลงทะเบียน";
      break;
    case 'admin':
      template = HtmlService.createTemplateFromFile('AdminDashboard');
      pageTitle = "LMS - Admin Dashboard";
      break;
    // --- ในอนาคตเราจะเพิ่ม case สำหรับ 'teacher' และ 'student' ตรงนี้ ---
    default:
      // หน้าเริ่มต้นคือหน้า Login
      template = HtmlService.createTemplateFromFile('Login');
      pageTitle = "LMS - เข้าสู่ระบบ";
      break;
  }

  // .evaluate() คือหัวใจสำคัญที่ทำให้โค้ดพิเศษใน HTML ทำงาน
  return template.evaluate().setTitle(pageTitle);
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
          // *** ในอนาคต: สร้าง session หรือ token สำหรับการ login ***
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
    
    // *** ในอนาคต: ส่งอีเมลแจ้ง Admin ว่ามีคนสมัครใหม่ ***

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

/**
 * (สำหรับ Admin) ดึงรายชื่อผู้ใช้ที่รอการอนุมัติ
 */
function getPendingUsers() {
  try {
    const data = usersSheet.getDataRange().getValues();
    const headers = data.shift(); // ดึงแถวหัวข้อออก

    const pendingUsers = data.filter(row => row[6] === 'Pending') // คอลัมน์ G คือ Status
      .map(row => {
        // แปลง array เป็น object เพื่อให้ใช้ง่ายในฝั่ง Client
        return {
          username: row[0],
          name: row[2],
          email: row[5],
          class: row[3]
        };
      });
    return pendingUsers;
  } catch (e) {
    Logger.log(e);
    return []; // ถ้ามีปัญหา ให้ส่ง array ว่างกลับไป
  }
}

/**
 * (สำหรับ Admin) อนุมัติผู้ใช้งาน
 */
function approveUser(username) {
  try {
    const data = usersSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { // เริ่มจาก 1 เพื่อข้ามหัวข้อ
      if (data[i][0] === username) {
        usersSheet.getRange(i + 1, 7).setValue('Active'); // แถวที่ i+1, คอลัมน์ที่ 7 (G)
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
        usersSheet.getRange(i + 1, 7).setValue('Rejected'); // แถวที่ i+1, คอลัมน์ที่ 7 (G)
        return { success: true };
      }
    }
    return { success: false, message: 'User not found' };
  } catch (e) {
    Logger.log(e);
    return { success: false, message: e.message };
  }
}


// =================================================================
// == ฟังก์ชันอื่นๆ (Placeholder สำหรับการพัฒนาในอนาคต)
// =================================================================

// TODO: สร้างฟังก์ชันสำหรับจัดการคลังข้อสอบ (สร้าง, แก้ไข, ลบ)
// TODO: สร้างฟังก์ชันสำหรับจัดการชุดข้อสอบ (Test Templates)
// TODO: สร้างฟังก์ชันสำหรับมอบหมายงาน (Assignments)
// TODO: สร้างฟังก์ชันสำหรับดึงข้อมูลข้อสอบมาแสดงให้นักเรียน
// TODO: สร้างฟังก์ชันสำหรับตรวจคำตอบและบันทึกผล
// TODO: สร้างฟังก์ชันสำหรับระบบ Gamification, Announcements, Logs ฯลฯ
