function doGet(e) {
  // สร้าง HTML จากไฟล์ template ของคุณ
  var html = HtmlService.createTemplateFromFile('public/index').evaluate();
  
  // เพิ่ม Viewport Meta Tag ด้วยวิธีของ Apps Script
  html.addMetaTag('viewport', 'content="width=device-width, initial-scale=1.0, viewport-fit=cover, user-scalable=no"');
  
  // ---> ✨ บรรทัดที่ต้องเพิ่มเข้ามาเพื่อแก้ปัญหา X-Frame-Options ✨ <---
  html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  
  return html;
}

// ฟังก์ชันอื่นๆ ที่เหลือไม่ต้องแก้ไขครับ
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * ดึง URL ปัจจุบันของเว็บแอปที่ Deploy ไว้
 * @returns {string} URL ของเว็บแอป
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}
