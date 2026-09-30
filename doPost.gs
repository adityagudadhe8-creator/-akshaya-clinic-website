// ─────────────────────────────────────────────────────────────
//  Akshaya Child Clinic — Google Apps Script
//  Paste this entire file into your Apps Script project.
//
//  ONE-TIME SETUP (run once from the editor):
//    1. Open Extensions > Apps Script in your Google Sheet
//    2. Paste this code, replacing all existing code
//    3. From the function dropdown, select setWhatsAppCredentials → Run
//    4. Deploy → Manage deployments → New version
// ─────────────────────────────────────────────────────────────

function setWhatsAppCredentials() {
  PropertiesService.getScriptProperties().setProperties({
    WHATSAPP_PHONE_NUMBER_ID: 'PASTE_YOUR_PHONE_NUMBER_ID',
    WHATSAPP_ACCESS_TOKEN:    'PASTE_YOUR_PERMANENT_ACCESS_TOKEN'
  });
}
/**
 * ── ONE-TIME SETUP ──
 * Run this once from the Apps Script editor (pick setWhatsAppCredentials
 * from the function dropdown, then click Run) after you have your values
 * from Meta. Keeps the token out of the code body itself. Safe to leave
 * in place afterwards — the values persist in Script Properties either way.
 */
function setWhatsAppCredentials() {
  PropertiesService.getScriptProperties().setProperties({
    WHATSAPP_PHONE_NUMBER_ID: 'PASTE_YOUR_PHONE_NUMBER_ID',
    WHATSAPP_ACCESS_TOKEN: 'PASTE_YOUR_PERMANENT_ACCESS_TOKEN'
  });
}

function doGet(e) {
  const date = String((e.parameter && e.parameter.date) || '').trim();
  const callback = String((e.parameter && e.parameter.callback) || '').replace(/[^a-zA-Z0-9_]/g, '');

  let bookedSlots = [];

  if (date) {
    try {
      const sheet = SpreadsheetApp
        .getActiveSpreadsheet()
        .getSheetByName('Appointments');

      const lastRow = sheet.getLastRow();

      if (lastRow >= 2) {
        const rows = sheet.getRange(2, 1, lastRow - 1, 8).getValues();

        rows.forEach(function(row) {
          const existingDate = formatDate(row[4]);
          const existingTime = formatTime(row[5]);
          const existingStatus = String(row[7]).trim().toLowerCase();

          if (existingDate === date && existingStatus !== 'cancelled') {
            bookedSlots.push(existingTime);
          }
        });
      }
    } catch (error) {
      // return empty on error — doPost still guards against double-booking
    }
  }

  const json = JSON.stringify({ bookedSlots: bookedSlots });

  // JSONP: wrap in callback if provided, so the browser can read it
  // despite the GAS redirect stripping CORS headers
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}


function doPost(e) {
  const lock = LockService.getScriptLock();

  let data, date, time;
  let result;

  try {
    lock.waitLock(10000);

    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Appointments");

    data = JSON.parse(e.postData.contents);

    date = String(data.date || "").trim();
    time = String(data.time || "").trim();

    if (!date || !time) {
      result = { message: "Please select an appointment date and time.", success: false };

    } else if (!isValidIndianMobile_(data.phone)) {
      result = { message: "Please enter a valid 10-digit Indian mobile number.", success: false };

    } else {
      // Check existing appointments
      const lastRow = sheet.getLastRow();
      let conflict = false;

      if (lastRow >= 2) {
        const rows = sheet.getRange(2, 1, lastRow - 1, 8).getValues();

        for (let i = 0; i < rows.length; i++) {
          const existingDate = formatDate(rows[i][4]);
          const existingTime = formatTime(rows[i][5]);
          const existingStatus = String(rows[i][7]).trim().toLowerCase();

          // Cancelled appointments don't block the slot
          if (
            existingDate === date &&
            existingTime === time &&
            existingStatus !== "cancelled"
          ) {
            conflict = true;
            break;
          }
        }
      }

      if (conflict) {
        result = {
          message: "This appointment time is already booked. Please choose another time.",
          success: false
        };

      } else {
        // Save new appointment — auto-confirmed since it passed the
        // conflict check above while holding the lock
        sheet.appendRow([
          new Date(),
          data.parent || "",
          data.child || "",
          data.phone || "",
          date,
          time,
          data.reason || "",
          "Confirmed"
        ]);

        result = { message: "Appointment confirmed.", success: true };
      }
    }

  } catch (error) {
    result = {
      message: "Unable to submit appointment: " + error.toString(),
      success: false
    };

  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }

  // WhatsApp is sent AFTER the lock is released, so a slow API call never
  // makes other parents wait longer to book. If it fails for any reason
  // (token expired, template not approved yet, bad number), the
  // appointment itself is unaffected since it's already saved above —
  // the failure is only logged, for you to check under Executions.
  if (result.success) {
    try {
      sendWhatsAppConfirmation(data.phone, data.parent, data.child, date, time);
    } catch (waError) {
      console.error("WhatsApp send failed: " + waError);
    }
  }

  return response(result.message, result.success);
}


function sendWhatsAppConfirmation(phone, parentName, childName, date, time) {
  const props = PropertiesService.getScriptProperties();
  const PHONE_NUMBER_ID = props.getProperty('WHATSAPP_PHONE_NUMBER_ID');
  const ACCESS_TOKEN = props.getProperty('WHATSAPP_ACCESS_TOKEN');
  const TEMPLATE_NAME = 'appointment_confirmed'; // must exactly match the approved template name
  const TEMPLATE_LANG = 'en';

  if (!PHONE_NUMBER_ID || !ACCESS_TOKEN) {
    console.error('WhatsApp not sent: credentials missing. Run setWhatsAppCredentials() first.');
    return;
  }

  // Normalise to WhatsApp's expected format: country code + number,
  // digits only, no +, spaces, or leading 0.
  let to = String(phone || '').replace(/\D/g, '');
  if (to.length === 10) to = '91' + to;
  else if (to.length === 11 && to.charAt(0) === '0') to = '91' + to.slice(1);

  const payload = {
    messaging_product: 'whatsapp',
    to: to,
    type: 'template',
    template: {
      name: TEMPLATE_NAME,
      language: { code: TEMPLATE_LANG },
      components: [{
        type: 'body',
        parameters: [
          { type: 'text', text: parentName || 'there' },
          { type: 'text', text: childName || 'your child' },
          { type: 'text', text: friendlyDate(date) },
          { type: 'text', text: friendlyTime(time) }
        ]
      }]
    }
  };

  const res = UrlFetchApp.fetch(
    'https://graph.facebook.com/v26.0/' + PHONE_NUMBER_ID + '/messages',
    {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + ACCESS_TOKEN },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    }
  );

  const code = res.getResponseCode();
  if (code >= 300) {
    console.error('WhatsApp API error ' + code + ': ' + res.getContentText());
  }
}


function friendlyDate(dateStr) {
  const d = new Date(dateStr + 'T00:00:00');
  const days = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  const months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  return days[d.getDay()] + ', ' + months[d.getMonth()] + ' ' + d.getDate();
}


function friendlyTime(timeStr) {
  const parts = timeStr.split(':').map(Number);
  const h = parts[0], m = parts[1];
  const hour12 = h % 12 || 12;
  const ampm = h >= 12 ? 'PM' : 'AM';
  return hour12 + ':' + String(m).padStart(2, '0') + ' ' + ampm;
}


function isValidIndianMobile_(raw) {
  const d = String(raw || '').replace(/\D/g, '');
  return /^[6-9]\d{9}$/.test(d) || /^91[6-9]\d{9}$/.test(d);
}


function formatDate(value) {
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );
  }

  return String(value).trim();
}


function formatTime(value) {
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      "HH:mm"
    );
  }

  return String(value).trim();
}


function response(message, success) {
  return ContentService
    .createTextOutput(
      JSON.stringify({
        success: success,
        message: message
      })
    )
    .setMimeType(ContentService.MimeType.JSON);
}
