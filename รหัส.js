/**
 * ฟังก์ชัน doGet(e)
 * คืนค่าไฟล์ index.html เสมอ (Single‐Page App)
 * เราจะใช้ URL parameter เช่น '?page=a4' มาตรวจฝั่ง Client (JS) ใน index.html
 */
function doGet(e) {
  // 1) สร้าง Template จากไฟล์ index.html
  // 2) กำหนดค่า baseUrl ให้ฝั่ง HTML/JS เอาไปใช้ได้ (optional)
  var template = HtmlService.createTemplateFromFile('index');
  // กำหนดตัวแปร baseUrl ให้เข้าไปใช้ใน index.html ได้ (ถ้าต้องการเอาไปต่อเป็นลิงก์เต็ม)
  template.baseUrl = ScriptApp.getService().getUrl();
  
  return template
    .evaluate()
    .setTitle('ระบบเบิกวัสดุสำนักงาน')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * ฟังก์ชัน search(id)
 * สมมติเรามี Sheet แรก (Sheet1) ที่คอลัมน์ C เก็บรหัส item (id) 
 * หากมี ?id=xxx ก็จะแสดงตารางบนฟอร์ม
 */
function search(id) {
  var sheet = SpreadsheetApp.getActive().getSheets()[0];
  var data = sheet.getDataRange().getDisplayValues();
  data.shift(); // ตัด header
  var ar = [];
  data.forEach(function(f) {
    if (f[2].toString().indexOf(id) > -1) {
      ar.push(f);
    }
  });
  return ar;
}

/**
 * ฟังก์ชัน saveDataToSheet(...)
 * รับค่าจากฝั่ง client แล้วเขียนลงชีต “Form_Data”
 */
function saveDataToSheet(id, fname, department, quantity, purpose, additionalInfo, time) {
  var sheet = SpreadsheetApp.getActive().getSheetByName("Form_Data");
  let newrow = [
    "",         // A: เว้นว่าง (จะถูก Timestamp ใหม่ในชีตแทน)
    new Date(), // B: Timestamp ปัจจุบัน
    id || "",   // C: id (ถ้ามี) ถ้าไม่มีก็ใส่ค่าว่าง
    "",         // D: เว้นว่าง
    fname,      // E: ชื่อผู้ขอเบิก
    department, // F: แผนก
    quantity,   // G: จำนวน
    purpose,    // H: ความจำเป็น
    additionalInfo, // I: ข้อมูลเพิ่มเติม
    "",         // J: เว้นว่าง
    // K: สูตรเช็กจาก Checkbox ในคอลัมน์ K
    `=IF(K${sheet.getLastRow()+1}=TRUE,"อนุมัติ","ไม่อนุมัติ")`
  ];
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, newrow.length).setValues([newrow]);
  sheet.getRange('K' + sheet.getLastRow()).insertCheckboxes();
}