// ==========================================
// 🚀 BACKEND: SYSTEM CONFIG & ROUTING
// ==========================================
const SPREADSHEET_ID = '1cM0oR6-DHT11cRucOoUIVD6_1Lcp3Av7e7sLWcQEOO0'; // เปลี่ยนเป็น ID ของคุณ

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Enterprise Asset Management')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================
// 🔐 SECURITY & AUTHENTICATION
// ==========================================
function authenticateUser(username, password) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == username && data[i][1] == password) {
      logAction(username, 'LOGIN', 'User logged in successfully');
      return { success: true, user: { username: data[i][0], role: data[i][2], name: data[i][3] } };
    }
  }
  return { success: false, message: 'Invalid credentials' };
}

// ==========================================
// 📦 ASSET MANAGEMENT (CRUD)
// ==========================================
function getDashboardData() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Assets');
  // 💡 เปลี่ยนจาก getValues() เป็น getDisplayValues()
  const data = sheet.getDataRange().getDisplayValues(); 
  
  let total = data.length - 1;
  let normal = 0, broken = 0, lost = 0;
  
  for(let i = 1; i < data.length; i++) {
    let status = data[i][5];
    if(status === 'ปกติ') normal++;
    else if(status === 'ชำรุด') broken++;
    else if(status === 'สูญหาย') lost++;
  }
  
  return {
    total: total,
    normal: normal,
    broken: broken,
    lost: lost,
    assets: data.slice(1, 11)
  };
}

function addAsset(assetData, username) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Assets');
    
    // 1. ตรวจสอบรหัสครุภัณฑ์ (ถ้าไม่ได้กรอกมา ให้ระบบสร้างให้)
    let assetId = assetData.id;
    if (!assetId || assetId.trim() === '') {
      assetId = 'AST-' + new Date().getTime(); // สุ่มรหัสอัตโนมัติ
    } else {
      // ตรวจสอบว่ารหัสที่กรอกมา ซ้ำกับในระบบหรือไม่
      const existingData = sheet.getDataRange().getDisplayValues();
      for (let i = 1; i < existingData.length; i++) {
        if (existingData[i][0] === assetId) {
          return { success: false, message: 'รหัสครุภัณฑ์นี้มีอยู่ในระบบแล้ว กรุณาใช้รหัสอื่น!' };
        }
      }
    }
    
    // 2. สร้าง QR Code
    const qrCode = `https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=${assetId}`;
    
    // 3. บันทึกลง Sheet
    sheet.appendRow([
      assetId, assetData.name, assetData.category, assetData.location, 
      assetData.price, 'ปกติ', assetData.imageUrl, qrCode, new Date()
    ]);
    
    logAction(username, 'ADD_ASSET', `Added asset: ${assetData.name} (${assetId})`);
    return { success: true, message: 'บันทึกครุภัณฑ์สำเร็จ!' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ==========================================
// 🛡️ AUDIT LOG & NOTIFICATIONS
// ==========================================
function logAction(user, action, details) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Logs');
  sheet.appendRow([new Date(), user, action, details]);
}

// 🔔 Hook สำหรับ Telegram Webhook
function sendTelegramNotification(message) {
  const token = "YOUR_TELEGRAM_BOT_TOKEN";
  const chatId = "YOUR_CHAT_ID";
  const url = `https://api.telegram.org/bot${token}/sendMessage?chat_id=${chatId}&text=${encodeURIComponent(message)}`;
  // UrlFetchApp.fetch(url); // Uncomment เมื่อใส่ Token แล้ว
}

function getAllAssets() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Assets');
  // 💡 เปลี่ยนจาก getValues() เป็น getDisplayValues()
  const data = sheet.getDataRange().getDisplayValues(); 
  return data.slice(1);
}

// ==========================================
// 📝 INSPECTION SYSTEM
// ==========================================
function saveInspection(assetId, status, note, username) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const assetSheet = ss.getSheetByName('Assets');
    const inspSheet = ss.getSheetByName('Inspections');
    
    // 1. ตรวจสอบว่ามีรหัสครุภัณฑ์นี้ในระบบหรือไม่ และอัปเดตสถานะ
    const assetData = assetSheet.getDataRange().getDisplayValues();
    let assetFound = false;
    let timestamp = new Date();
    
    for (let i = 1; i < assetData.length; i++) {
      if (assetData[i][0] === assetId) {
        // อัปเดตสถานะใหม่ในคอลัมน์ F (Index 6 ใน GAS) และเวลาอัปเดตในคอลัมน์ I (Index 9)
        assetSheet.getRange(i + 1, 6).setValue(status); 
        assetSheet.getRange(i + 1, 9).setValue(timestamp); 
        assetFound = true;
        break;
      }
    }
    
    if (!assetFound) {
      return { success: false, message: 'ไม่พบรหัสครุภัณฑ์นี้ในระบบ กรุณาตรวจสอบรหัสอีกครั้ง!' };
    }
    
    // 2. บันทึกประวัติการตรวจลงชีต Inspections
    const inspId = 'INSP-' + new Date().getTime();
    inspSheet.appendRow([inspId, assetId, status, note, '', username, timestamp]);
    
    logAction(username, 'INSPECT', `ตรวจสภาพ ${assetId}: ${status}`);
    return { success: true, message: 'บันทึกผลการตรวจสอบและอัปเดตสถานะเรียบร้อยแล้ว!' };
    
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ==========================================
// 📦 BULK IMPORT (นำเข้าข้อมูลหลายรายการ)
// ==========================================
function addMultipleAssets(assetsArray, username) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Assets');
    const existingData = sheet.getDataRange().getDisplayValues();
    const existingIds = existingData.map(row => row[0]); // เก็บรายชื่อ ID ที่มีอยู่แล้วเพื่อกันซ้ำ
    
    let newRows = [];
    let addedCount = 0;
    let timestamp = new Date();

    for (let i = 0; i < assetsArray.length; i++) {
      let asset = assetsArray[i];
      if (!asset.name || asset.name.trim() === '') continue; // ข้ามบรรทัดที่ไม่มีชื่อ
      
      let assetId = asset.id ? asset.id.trim() : '';
      
      // ถ้ารหัสว่าง ให้สร้างใหม่ / ถ้ามีรหัสแล้ว แต่ดันซ้ำกับในระบบ ให้ข้ามรายการนั้นไป
      if (assetId === '') {
        assetId = 'AST-' + new Date().getTime() + '-' + i; 
      } else if (existingIds.includes(assetId)) {
        continue; 
      }
      
      const qrCode = `https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=${assetId}`;
      
      // เรียงคอลัมน์ให้ตรงกับ Google Sheets
      newRows.push([
        assetId, 
        asset.name, 
        asset.category || 'Other', 
        asset.location || '-', 
        asset.price || 0, 
        'ปกติ', 
        '', // imageUrl (ว่างไว้ก่อน)
        qrCode, 
        timestamp
      ]);
      
      existingIds.push(assetId); // อัปเดต ID กันการซ้ำกันเองในชุดข้อมูลที่กำลังนำเข้า
      addedCount++;
    }
    
    // บันทึกลง Sheet ทีเดียวทั้งหมด (รวดเร็วมาก)
    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
      logAction(username, 'BULK_ADD', `Imported ${addedCount} assets`);
      return { success: true, message: `นำเข้าข้อมูลสำเร็จจำนวน ${addedCount} รายการ!` };
    } else {
      return { success: false, message: 'ไม่มีข้อมูลใหม่ถูกบันทึก (อาจมีรหัสซ้ำ หรือข้อมูลไม่ถูกต้อง)' };
    }
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}