// ════════════════════════════════════════════════════════════════
//  The XL Academy — Delhi Landing Page Lead Capture & Email Alert
//
//  DEPLOYMENT INSTRUCTIONS:
//  ─────────────────────────────────────────────────────────────
//  1. Open your Delhi Leads Google Sheet (where rows are arriving)
//  2. In top menu, click: Extensions → Apps Script
//  3. Select and DELETE all existing code in Code.gs
//  4. Copy and PASTE this entire file content into Code.gs
//  5. Click Save (Ctrl+S or disk icon)
//  6. Click "Deploy" (blue button at top right) → "New deployment"
//     (Or "Manage deployments" → click pencil icon → version "New version")
//  7. Make sure:
//     - Select type: "Web app"
//     - Execute as: "Me"
//     - Who has access: "Anyone"
//  8. Click "Deploy", click "Authorize access", and grant permissions
// ════════════════════════════════════════════════════════════════

var SHEET_ID = '1UO43gfeqcgif9R5BGTz45ejMq5suCy7HngHYGzklaN8';
var NOTIFY_TO = 'support@thexlacademy.com, Analyticsproschool@gmail.com';

function doPost(e) {
  var p = {};
  try {
    if (e.postData && e.postData.contents) {
      p = JSON.parse(e.postData.contents);
    }
  } catch(_) {
    p = (e && e.parameter) ? e.parameter : {};
  }

  var saved = false, mailed = false, saveErr = '', mailErr = '';
  try { saveRow(p); saved = true; } catch(err) { saveErr = err.message; console.error('Save error: ' + err.message); }
  try { sendMail(p); mailed = true; } catch(err) { mailErr = err.message; console.error('Mail error: ' + err.message); }

  return ContentService.createTextOutput(
    JSON.stringify({ status: saved ? 'success' : 'error', message: saved ? 'Lead recorded successfully' : saveErr, saved: saved, mailed: mailed, mailErr: mailErr })
  ).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  if (!p.name) {
    return ContentService.createTextOutput('OK — Delhi LP Lead Capture & Email Alert Script is active. Connected to Sheet: ' + SHEET_ID).setMimeType(ContentService.MimeType.TEXT);
  }
  var saved = false, mailed = false;
  try { saveRow(p); saved = true; } catch(_) {}
  try { sendMail(p); mailed = true; } catch(_) {}
  return ContentService.createTextOutput(JSON.stringify({ status: saved ? 'success' : 'error', saved: saved, mailed: mailed })).setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  var ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch(e) {}
  if (!ss && SHEET_ID) {
    try {
      ss = SpreadsheetApp.openById(SHEET_ID);
    } catch(err) {
      console.error('Error opening sheet by ID: ' + err.message);
    }
  }
  if (!ss) {
    throw new Error('Spreadsheet not found. Please check SHEET_ID or script container.');
  }
  var sh = ss.getActiveSheet();
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Timestamp', 'Name', 'Phone', 'Email', 'Course', 'City', 'Action Type', 'Source', 'Page URL']);
    var hdr = sh.getRange(1, 1, 1, 9);
    hdr.setFontWeight('bold').setBackground('#E8470A').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function saveRow(d) {
  var ts = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd-MM-yyyy HH:mm:ss');
  var sh = getSheet();
  sh.appendRow([
    ts,
    d.name || '',
    d.phone || '',
    d.email || '',
    d.course || 'Data Analytics',
    d.city || 'Delhi',
    d.action_type || 'Enquiry / Demo',
    d.source || 'Delhi Landing Page',
    d.page_url || ''
  ]);
}

function sendMail(d) {
  var n   = d.name  || 'Student';
  var ph  = d.phone || '—';
  var em  = d.email || '';
  var co  = d.course || 'Data Analytics Course';
  var ci  = d.city || 'Delhi';
  var act = d.action_type || 'Enquiry / Demo';
  var ts  = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd MMM yyyy, hh:mm a');

  var html =
    '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,0.08);background:#ffffff;">' +
    '<div style="background:linear-gradient(135deg, #E8470A, #F5A623);padding:28px 32px;color:#ffffff;">' +
    '<h1 style="margin:0;font-size:22px;color:#ffffff;">📩 New Lead — Delhi Landing Page</h1>' +
    '<p style="margin:6px 0 0;font-size:13px;opacity:0.9;color:#ffffff;">' + ts + ' IST</p>' +
    '</div>' +
    '<div style="padding:28px 32px;">' +
    field('Full Name', n) +
    field('Mobile Number', '+91 ' + ph) +
    field('Email Address', em) +
    field('Course Interested In', co) +
    field('City', ci) +
    field('Action Type', act) +
    (d.page_url ? field('Page URL', '<a href="' + d.page_url + '" style="color:#E8470A;">' + d.page_url + '</a>') : '') +
    '</div>' +
    '<div style="background:#0D1B35;padding:16px 32px;text-align:center;color:rgba(255,255,255,0.6);font-size:12px;">' +
    '© ' + new Date().getFullYear() + ' The XL Academy &nbsp;|&nbsp; Delhi Landing Page Lead Alert' +
    '</div>' +
    '</div>';

  var recipients = d.notify_to || NOTIFY_TO;
  MailApp.sendEmail({
    to: recipients,
    subject: 'New Lead : Delhi LP - ' + n,
    htmlBody: html,
    replyTo: em || 'support@thexlacademy.com'
  });

  // Also send an automated confirmation email to the student
  if (em && em.indexOf('@') > -1) {
    try {
      var autoHtml =
        '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,0.08);background:#ffffff;">' +
        '<div style="background:linear-gradient(135deg, #0D1B35, #1A3A6B);padding:28px 32px;text-align:center;color:#ffffff;">' +
        '<h1 style="margin:0;font-size:22px;color:#ffffff;">🎉 Thank You, ' + n + '!</h1>' +
        '<p style="margin:6px 0 0;font-size:13px;opacity:0.85;color:#ffffff;">Your enquiry has been received successfully.</p>' +
        '</div>' +
        '<div style="padding:28px 32px;color:#333;font-size:15px;line-height:1.6;">' +
        '<p>Hi <strong>' + n + '</strong>,</p>' +
        '<p>Thank you for reaching out to <strong>The XL Academy</strong>. We have received your enquiry and our career counsellor will call you shortly.</p>' +
        '<div style="background:#fff7f4;border-left:4px solid #E8470A;padding:12px 16px;border-radius:6px;margin:20px 0;font-size:14px;color:#1A3A6B;">' +
        '📚 <b>Course:</b> ' + co + '<br>' +
        '🏙️ <b>City:</b> ' + ci + '<br>' +
        '📱 <b>Contact Phone:</b> +91 ' + ph +
        '</div>' +
        '<p>Need immediate assistance? Feel free to call us directly:</p>' +
        '<a href="tel:+917428703467" style="display:inline-block;background:linear-gradient(135deg, #E8470A, #ff6b35);color:#fff;text-decoration:none;padding:12px 28px;border-radius:50px;font-weight:700;font-size:14px;margin:10px 0;">📞 Call: +91 74287 03467</a>' +
        '</div>' +
        '<div style="background:#0D1B35;padding:16px 32px;text-align:center;color:rgba(255,255,255,0.6);font-size:12px;">' +
        '© ' + new Date().getFullYear() + ' The XL Academy &nbsp;|&nbsp; Dwarka Sec-13, New Delhi – 110078' +
        '</div>' +
        '</div>';

      MailApp.sendEmail({
        to: em,
        subject: 'Thank you for your enquiry — The XL Academy',
        htmlBody: autoHtml,
        replyTo: 'support@thexlacademy.com'
      });
    } catch(studentErr) {
      console.warn('Could not send student auto-reply: ' + studentErr.message);
    }
  }
}

function field(label, value) {
  return '<div style="margin-bottom:16px;">' +
    '<div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#E8470A;margin-bottom:4px;">' + label + '</div>' +
    '<div style="font-size:15px;color:#1A3A6B;font-weight:600;background:#f8faff;border-left:3px solid #E8470A;padding:10px 14px;border-radius:6px;">' + value + '</div>' +
    '</div>';
}
