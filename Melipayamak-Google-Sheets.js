// تنظیمات پنل پیامکی ملی پیامک
const CONFIG = {
  username: "",
  password: "",
  from: ""
};

// تابع اصلی ارسال انبوه
function sendBulkSMS() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  let sentCount = 0;
  let errorCount = 0;

  Logger.log("شروع ارسال انبوه...");
  Logger.log("تعداد ردیف‌های موجود: " + data.length);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const phoneNumber = row[0];
    const message = row[1];
    const serviceType = row[2] || "";
    const currentStatus = row[3] || "";

    Logger.log(`ردیف ${i+1}: شماره=${phoneNumber}, پیام=${message}, نوع سرویس=${serviceType}, وضعیت=${currentStatus}`);

    if (phoneNumber && message && currentStatus !== "ارسال شد") {
      Logger.log("این ردیف برای ارسال واجد شرایط است");

      try {
        let result;

        // بررسی نوع وب سرویس ارسال پیامک
        if (serviceType === "خدماتی اشتراکی" || serviceType === "base") {
          Logger.log("ارسال از نوع خدماتی اشتراکی (base)");
          result = sendBaseSMS(phoneNumber, message);
        } else {
          // حالت پیش‌فرض: خط اختصاصی
          Logger.log("ارسال از نوع خط اختصاصی (direct)");
          result = sendDirectSMS(phoneNumber, message);
        }

        Logger.log("نتیجه ارسال: " + JSON.stringify(result));

        // بروزرسانی وضعیت
        sheet.getRange(i + 1, 4).setValue("ارسال شد");
        sheet.getRange(i + 1, 5).setValue(JSON.stringify(result));
        sheet.getRange(i + 1, 6).setValue(new Date());
        sheet.getRange(i + 1, 7).setValue("✅");

        sentCount++;

      } catch (error) {
        Logger.log("خطا در ارسال: " + error.toString());
        sheet.getRange(i + 1, 4).setValue("خطا در ارسال");
        sheet.getRange(i + 1, 5).setValue("خطا: " + error.toString());
        sheet.getRange(i + 1, 7).setValue("❌");
        errorCount++;
      }

      Utilities.sleep(800);
    } else {
      Logger.log("این ردیف واجد شرایط نیست - یا شماره/پیام خالی است یا قبلاً ارسال شده");
    }
  }

  Logger.log(`پایان ارسال - موفق: ${sentCount}, خطا: ${errorCount}`);

  // نمایش نتیجه
  const message = `ارسال انبوه تکمیل شد!\n✅ تعداد ارسال موفق: ${sentCount}\n❌ تعداد خطا: ${errorCount}`;
  SpreadsheetApp.getUi().alert(message);
}

// ارسال پیامک از طریق پترن و الگو با وب سرویس خدماتی (خط خدماتی اشتراکی ملی پیامک)
function sendBaseSMS(phoneNumber, message) {
  const url = "https://rest.payamak-panel.com/api/SendSMS/BaseServiceNumber";
  const payload = {
    "username": CONFIG.username,
    "password": CONFIG.password,
    "text": message,
    "to": phoneNumber,
    "bodyId": "" // کد پترن و الگوی وارد‌شده در ملی پیامک را وارد کنید.
  };

  const options = {
    "method": "POST",
    "headers": {"content-type": "application/x-www-form-urlencoded"},
    "payload": payload,
    "muteHttpExceptions": true
  };

  const response = UrlFetchApp.fetch(url, options);
  return JSON.parse(response.getContentText());
}

// ارسال پیامک از طریق خطوط اختصاصی (در صورتی که خط فرستندۀ شما تبلیغاتی باشد، به لیست سیاه مخابرات ارسال نخواهد شد.)
function sendDirectSMS(phoneNumber, message) {
  const url = "https://rest.payamak-panel.com/api/SendSMS/SendSMS";
  const payload = {
    "username": CONFIG.username,
    "password": CONFIG.password,
    "text": message,
    "to": phoneNumber,
    "from": CONFIG.from
  };

  const options = {
    "method": "POST",
    "headers": {"content-type": "application/x-www-form-urlencoded"},
    "payload": payload,
    "muteHttpExceptions": true
  };

  const response = UrlFetchApp.fetch(url, options);
  return JSON.parse(response.getContentText());
}

// تابع برای ایجاد منوی dropdown در ستون نوع سرویس
function setupDataValidation() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();

  // محدوده ستون C (ستون سوم - نوع سرویس)
  const range = sheet.getRange(2, 3, lastRow - 1, 1);

  // ایجاد قانون اعتبارسنجی
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['خط اختصاصی', 'خدماتی اشتراکی'], true)
    .setAllowInvalid(false)
    .setHelpText('لطفاً یکی از گزینه‌ها را انتخاب کنید: خط اختصاصی یا خدماتی اشتراکی')
    .build();

  range.setDataValidation(rule);

  SpreadsheetApp.getUi().alert("منوی انتخابی برای ستون نوع سرویس ایجاد شد!");
}

// منوی سفارسی
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📱 سیستم پیامک')
    .addItem('ارسال انبوه پیام‌ها', 'sendBulkSMS')
    .addItem('ایجاد منوی انتخابی', 'setupDataValidation')
    .addItem('بازنشانی وضعیت‌ها', 'resetAllStatuses')
    .addItem('اضافه کردن ردیف تست', 'addTestRow')
    .addItem('تست اتصال API', 'testConnection')
    .addToUi();
}

// بازنشانی تمام وضعیت‌ها برای ارسال مجدد
function resetAllStatuses() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  let resetCount = 0;

  for (let i = 1; i < data.length; i++) {
    const currentStatus = data[i][3]; // ستون وضعیت

    if (currentStatus === "ارسال شد" || currentStatus === "خطا در ارسال") {
      sheet.getRange(i + 1, 4).setValue(""); // پاک کردن وضعیت
      sheet.getRange(i + 1, 5).setValue(""); // پاک کردن پاسخ
      sheet.getRange(i + 1, 7).setValue(""); // پاک کردن آیکون
      resetCount++;
    }
  }

  SpreadsheetApp.getUi().alert(`✅ ${resetCount} ردیف برای ارسال مجدد آماده شدند`);
}

// اضافه کردن ردیف تست با نوع سرویس پیش‌فرض
function addTestRow() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const testData = [
    "09123456789",
    "این یک پیام تست است",
    "خط اختصاصی", // نوع سرویس پیش‌فرض
    "", "", "", ""
  ];
  sheet.appendRow(testData);
}

// تابع تست اتصال
function testConnection() {
  try {
    const testPayload = {
      "username": CONFIG.username,
      "password": CONFIG.password,
      "text": "ارسال پیامک تست از گوگل شیت لغو11",
      "to": "09123456789",
      "from": CONFIG.from
    };

    const options = {
      "method": "POST",
      "headers": {"content-type": "application/x-www-form-urlencoded"},
      "payload": testPayload,
      "muteHttpExceptions": true
    };

    const response = UrlFetchApp.fetch("https://rest.payamak-panel.com/api/SendSMS/SendSMS", options);
    const result = response.getContentText();

    Logger.log("نتایج تست اتصال: " + result);
    SpreadsheetApp.getUi().alert("تست اتصال انجام شد. نتایج در لاگ قابل مشاهده است.");

  } catch (error) {
    Logger.log("خطا در تست اتصال: " + error.toString());
    SpreadsheetApp.getUi().alert("خطا در اتصال به API");
  }
}
