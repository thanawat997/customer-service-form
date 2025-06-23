const SHEET_ID = '1W9GRGByHj7qXB0MegfrLpLBqlKBUatoatlUzmRlwxrY';
const SHEET_NAME = 'รถเข้าสาขา';
const FOLDER_ID = '1RsG_0b0aEesEuGgdcVc7oDwlq-BCSKom'; // กำหนด FOLDER_ID ที่นี่

// --- เพิ่มส่วนนี้สำหรับ LINE Messaging API ---
const LINE_CHANNEL_ACCESS_TOKEN = 'ODXARGhis+j5WseDE1c7ARmpdkT9AVJhqzIV3uQeWiFvjUQCZ7ilRgJSn8FZEbPG3b5jYknk+O2Nr/efH+QLRsYwZeXA4Sm7/qMtHrF3FnuiOH+71SoRjJTHnaiYOepmdOhe0Yrc/6hvpWdfqHC9FgdB04t89/1O/w1cDnyilFU='; // ใช้ Token ที่ถูกต้องของคุณ
const LINE_GROUP_ID = 'C39ee429b28cf8210dc80d241d35a25a2'; 
// ------------------------------------------

function doGet() {
  return HtmlService.createHtmlOutputFromFile('form');
}
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function submitForm(data) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const timestamp = new Date();
  const headers = [
    "Timestamp",
    "วันที่รับบริการ",
    "ช่องทางการติดต่อของลูกค้า",
    "สาขา",
    "สถานที่ให้บริการ",
    "ผู้รับเคส",
    "ช่างผู้ตรวจรถ",
    "สถานะลูกค้า",
    "ชื่อลูกค้า",
    "เบอร์โทร",
    "ทะเบียนรถ",
    "ยี่ห้อ",
    "รุ่น",
    "รุ่นย่อย",
    "ปีรถ",
    "เลขไมล์",
    "สถานะเล่มรถ",
    "ชื่อไฟแนนซ์",
    "ยอดไฟแนนซ์",
    "ข้อมูลประกันภัย",
    "ประกันภัย",
    "ชื่อบริษัทประกันภัย",
    "วันหมดอายุภาษี",
    "ราคาคาดหวัง",
    "เหตุผลในการขายรถ",
    "หมายเหตุ",
    "รูปหน้ารถURL" // คอลัมน์สำหรับ URL รูปภาพ
  ];

  // แก้ไข: ย้ายการจัดการรูปภาพมาไว้ตรงนี้ เพื่อให้ imageUrl พร้อมใช้ก่อนสร้าง rowData
  let imageUrl = '';
  if (data.imageData && data.imageFileName) {
    try {
      imageUrl = uploadImage(data.imageData, data.imageFileName);
      // สำคัญ: เพิ่ม imageUrl เข้าไปใน data ที่จะส่งไปยัง Sheet ด้วย
      data['รูปหน้ารถURL'] = imageUrl; 
    } catch (error) {
      Logger.log('Error during image upload in submitForm: ' + error.message);
      throw new Error('ไม่สามารถอัปโหลดรูปภาพได้: ' + error.message);
    }
  } else {
    // กรณีที่ไม่มีการอัปโหลดรูปภาพ หรือไม่ได้รับข้อมูลรูปภาพ
    data['รูปหน้ารถURL'] = ''; // ตั้งค่าให้เป็นค่าว่างใน Sheet
  }


  const rowData = headers.map(header => {
    let value = data[header];
    if (header === "Timestamp") {
      return timestamp;
    } else if (header === "เบอร์โทร") {
      return value ? String(value).replace(/[^0-9]/g, '').substring(0, 10) : '';
    } else if (header === "เลขไมล์" || header === "ยอดไฟแนนซ์" || header === "ราคาคาดหวัง") {
      const cleanedValue = String(value || '').replace(/,/g, '').replace(/[^0-9]/g, '');
      return cleanedValue ? parseFloat(cleanedValue) : '';
    } else {
      return value !== undefined ? value : '';
    }
  });

  const phoneColumnIndex = headers.indexOf("เบอร์โทร");
  if (phoneColumnIndex === -1) {
    throw new Error("Column 'เบอร์โทร' not found in headers.");
  }
  
  sheet.appendRow(rowData);

  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, phoneColumnIndex + 1).setNumberFormat('@'); 


  const serviceDate = data["วันที่รับบริการ"] || "ไม่ระบุ";
  const contactChannel = data["ช่องทางการติดต่อของลูกค้า"] || "ไม่ระบุ";
  const branch = data["สาขา"] || "ไม่ระบุ";
  const place = data["สถานที่ให้บริการ"] || "ไม่ระบุ";
  const servicer = data["ผู้รับเคส"] || "ไม่ระบุ";
  const inspector = data["ช่างผู้ตรวจรถ"] || "ไม่ระบุ";
  const status = data["สถานะลูกค้า"] || "ไม่ระบุ";
  const customerName = data["ชื่อลูกค้า"] || "ไม่ระบุ";
  const licensePlate = data["ทะเบียนรถ"] || "ไม่ระบุ";
  const brand = data["ยี่ห้อ"] || "ไม่ระบุ";
  const series = data["รุ่น"] || "ไม่ระบุ";
  const subSeries = data["รุ่นย่อย"] || "ไม่ระบุ";
  const year = data["ปีรถ"] || "ไม่ระบุ";
  const mile = data["เลขไมล์"] || "ไม่ระบุ";
  const statusCar = data["สถานะเล่มรถ"] || "ไม่ระบุ";
  const insurance = data["ข้อมูลประกันภัย"] || "ไม่ระบุ";
  const dateTax = data["วันหมดอายุภาษี"] || "ไม่ระบุ";
  const hope = data["ราคาคาดหวัง"] || "ไม่ระบุ";
  const reason = data["เหตุผลในการขายรถ"] || "ไม่ระบุ";
  const ps = data["หมายเหตุ"] || "ไม่ระบุ";
  const carImageUrl = data["รูปหน้ารถURL"] || "ไม่มีรูปภาพ";
  

  let lineMessage = `ข้อมูลลูกค้ารับบริการ\n`;
  lineMessage += `วันที่รับบริการ : ${serviceDate}\n`;
  lineMessage += `ช่องทางติดต่อ : ${contactChannel}\n`;
  lineMessage += `สาขา : ${branch}\n`;
  lineMessage += `สถานที่ให้บริการ : ${place}\n`;
  lineMessage += `ผู้รับเคส: ${servicer}\n`;
  lineMessage += `ช่างผู้ตรวจรถ: ${inspector}\n`;
  lineMessage += `สถานะลูกค้า: ${status}\n`;
  lineMessage += `ชื่อลูกค้า: ${customerName}\n`;
  lineMessage += `ทะเบียนรถ: ${licensePlate}\n`;
  lineMessage += `ยี่ห้อ: ${brand}\n`;
  lineMessage += `รุ่น: ${series}\n`;
  lineMessage += `รุ่นย่อย: ${subSeries}\n`;
  lineMessage += `ปีรถ: ${year}\n`;
  lineMessage += `เลขไมล์: ${mile}\n`;
  lineMessage += `สถานะเล่มรถ: ${statusCar}\n`;
  lineMessage += `ข้อมูลประกันภัย: ${insurance}\n`;
  lineMessage += `วันหมดอายุภาษี: ${dateTax}\n`;
  lineMessage += `ราคาคาดหวัง: ${hope}\n`;
  lineMessage += `เหตุผลในการขายรถ: ${reason}\n`;
  lineMessage += `หมายเหตุ : ${ps}\n`;
  lineMessage += `รูปหน้ารถ : ${carImageUrl !== "ไม่มีรูปภาพ" ? carImageUrl : "ไม่มีรูปภาพ"}\n`; 
  
  sendLineNotification(lineMessage);
  // --------------------------------------------------

  return 'บันทึกข้อมูลเรียบร้อยแล้ว';
}

function uploadImage(base64Data, fileName) {
  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const contentTypeMatch = base64Data.match(/^data:(image\/.+);base64,/);
    if (!contentTypeMatch) {
      throw new Error("Invalid base64 data format for image.");
    }
    const contentType = contentTypeMatch[1];
    const bytes = Utilities.base64Decode(base64Data.replace(/^data:image\/\w+;base64,/, ''));
    const blob = Utilities.newBlob(bytes, contentType, fileName);
    const file = folder.createFile(blob);
    Logger.log('Image uploaded: ' + file.getUrl());
    return file.getUrl();
  } catch (error) {
    Logger.log('Error uploading image: ' + error.message);
    throw new Error('ไม่สามารถอัปโหลดรูปภาพได้: ' + error.message);
  }
}

function sendLineNotification(message) {
  const url = "https://api.line.me/v2/bot/message/push"; 
  const headers = {
    "Content-Type": "application/json",
    "Authorization": "Bearer " + LINE_CHANNEL_ACCESS_TOKEN
  };
  const postData = {
    "to": LINE_GROUP_ID,
    "messages": [
      {
        "type": "text",
        "text": message
      }
    ]
  };

  const options = {
    "method": "post",
    "headers": headers,
    "payload": JSON.stringify(postData),
    "muteHttpExceptions": true 
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();
    
    if (responseCode === 200) {
      Logger.log("Line notification sent successfully.");
    } else {
      Logger.log("Failed to send Line notification. Response Code: " + responseCode + ", Body: " + responseBody);
    }
  } catch (e) {
    Logger.log("Error sending Line notification: " + e.message);
  }
}

function doPost(e) {
  try {
    if (e.postData && e.postData.type === 'application/json') {
      const data = JSON.parse(e.postData.contents);
      const result = submitForm(data); 
      return ContentService.createTextOutput(JSON.stringify({ status: 'success', message: result }))
        .setMimeType(ContentService.MimeType.JSON);
    } else {
      throw new Error('Unsupported content type or method.');
    }
  } catch (error) {
    Logger.log('ERROR in doPost: ' + error.message);
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
