/**
 * Rafin Cyber Cafe & CSC — প্যান আবেদনের ব্যাকএন্ড (Google Apps Script)
 * আবেদন জমা হয় Google Sheet-এ, কাগজপত্র আপনার Google Drive-এ (প্রাইভেট)।
 */
const CFG = {
  sheet: 'Applications',
  folder: 'PAN Application Files',
  notify: 'rafincybercafencsc2025@gmail.com',   // নতুন আবেদনের ইমেইল এখানে আসবে
  maxFileBytes: 2 * 1024 * 1024,
  hourlyLimit: 60                                // ঘণ্টায় সর্বোচ্চ আবেদন (স্প্যাম ঠেকাতে)
};
const STATUS = ['জমা হয়েছে', 'প্রক্রিয়া চলছে', 'সরকারি পোর্টালে জমা হয়েছে', 'সম্পন্ন', 'বাতিল / তথ্য সংশোধন প্রয়োজন'];
const HEAD = ['Ref', 'তারিখ', 'সেবা', 'স্ট্যাটাস', 'গ্রাহককে বার্তা', 'নাম', 'পিতার নাম', 'মাতার নাম',
  'জন্মতারিখ', 'লিঙ্গ', 'আধার', 'মোবাইল', 'ইমেইল', 'ধরন', 'ঠিকানা', 'বিদ্যমান প্যান', 'সংশোধন',
  'সঠিক তথ্য', 'ফাইল', 'নোট (শুধু আপনার জন্য)'];
const FILE_KEYS = { photo: 'ছবি', sign: 'স্বাক্ষর', aadhaarfile: 'আধার', dobproof: 'জন্মপ্রমাণ', fixproof: 'সংশোধনপ্রমাণ' };
const MIME = { 'image/jpeg': 'jpg', 'image/png': 'png', 'application/pdf': 'pdf' };

/** একবার চালান: শিট ও ফোল্ডার তৈরি করে */
function setup() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.sheet) || ss.insertSheet(CFG.sheet);
  sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold')
    .setBackground('#1b7a3d').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.getRange(2, 4, 2000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STATUS, true).build());
  getFolder();
}

function getFolder() {
  const it = DriveApp.getFoldersByName(CFG.folder);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CFG.folder);
}
function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function safe(v, max) {            // শিট ফর্মুলা ইনজেকশন ঠেকায়
  v = String(v == null ? '' : v).trim().slice(0, max || 200);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}
function need(ok, msg) { if (!ok) throw new Error(msg); }

/** আবেদন জমা */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const d = JSON.parse(e.postData.contents);

    const cache = CacheService.getScriptCache();
    const hk = 'h' + Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyMMddHH');
    const n = Number(cache.get(hk) || 0);
    need(n < CFG.hourlyLimit, 'এখন অনেক আবেদন আসছে, কিছুক্ষণ পরে চেষ্টা করুন।');

    const type = d.type === 'correction' ? 'correction' : 'new';
    const g = function (k, m) { return String(d[k] == null ? '' : d[k]).trim().slice(0, m || 200); };
    need(g('name', 100) && g('father', 100) && g('addr1') && g('po') && g('district') && g('state'), 'প্রয়োজনীয় তথ্য বাদ পড়েছে।');
    need(/^\d{12}$/.test(g('aadhaar')), 'আধার নম্বর ১২ অঙ্কের হতে হবে।');
    need(/^[6-9]\d{9}$/.test(g('phone')), 'সঠিক মোবাইল নম্বর দিন।');
    need(/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(g('email')), 'সঠিক ইমেইল দিন।');
    need(/^\d{6}$/.test(g('pin')), 'পিন কোড ৬ অঙ্কের হতে হবে।');
    need(/^\d{4}-\d{2}-\d{2}$/.test(g('dob')) && new Date(g('dob')) < new Date(), 'জন্মতারিখ ঠিক নয়।');
    if (type === 'correction') need(/^[A-Za-z]{5}\d{4}[A-Za-z]$/.test(g('oldpan')), 'প্যান নম্বর ঠিক নয়।');

    // রেফারেন্স
    const sh = SpreadsheetApp.getActive().getSheetByName(CFG.sheet);
    need(sh, 'সেটআপ সম্পূর্ণ হয়নি।');
    let ref;
    do {
      ref = 'RCC' + Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyMMdd') + ('000' + Math.floor(Math.random() * 10000)).slice(-4);
    } while (sh.createTextFinder(ref).matchEntireCell(true).findNext());

    // ফাইল
    const dir = getFolder(), links = [];
    (d.files || []).slice(0, 5).forEach(function (f) {
      if (!FILE_KEYS[f.key]) return;
      need(MIME[f.mime], 'শুধু JPG, PNG বা PDF ফাইল দিন।');
      const bytes = Utilities.base64Decode(String(f.data || ''));
      need(bytes.length <= CFG.maxFileBytes, 'কোনো ফাইল ২ MB এর বেশি হতে পারবে না।');
      const file = dir.createFile(Utilities.newBlob(bytes, f.mime, ref + '_' + f.key + '.' + MIME[f.mime]));
      links.push(FILE_KEYS[f.key] + ': ' + file.getUrl());
    });

    const addr = [g('addr1'), g('po'), g('ps'), g('district'), g('state'), g('pin')].filter(String).join(', ');
    const fixes = (Array.isArray(d.fix) ? d.fix : []).slice(0, 10).map(function (x) { return String(x).slice(0, 40); }).join(', ');
    const row = [ref, new Date(), type === 'new' ? 'নতুন প্যান' : 'সংশোধন / রিপ্রিন্ট', STATUS[0], '',
      safe(g('name', 100)), safe(g('father', 100)), safe(g('mother', 100)), g('dob'), safe(g('gender', 20)),
      "'" + g('aadhaar'), "'" + g('phone'), safe(g('email'), 100), safe(g('category', 40)), safe(addr, 300),
      type === 'correction' ? g('oldpan').toUpperCase() : '', safe(fixes, 200), safe(g('newvalue', 500), 500),
      links.join('\n'), ''];
    sh.appendRow(row);
    cache.put(hk, String(n + 1), 3600);

    try {
      MailApp.sendEmail(CFG.notify, 'নতুন প্যান আবেদন ' + ref,
        'রেফারেন্স: ' + ref + '\nসেবা: ' + row[2] + '\nনাম: ' + g('name', 100) + '\nমোবাইল: ' + g('phone') +
        '\n\nসব তথ্য ও ফাইল Google Sheet-এ দেখুন।');
    } catch (mailErr) { /* ইমেইল না গেলেও আবেদন সংরক্ষিত থাকে */ }

    return out({ ok: true, ref: ref });
  } catch (err) {
    return out({ ok: false, message: String(err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/** স্ট্যাটাস দেখা: ?ref=RCC2610270123 */
function doGet(e) {
  const ref = String((e.parameter && e.parameter.ref) || '').trim().toUpperCase();
  if (!/^RCC\d{10}$/.test(ref)) return out({ ok: false, message: 'রেফারেন্স নম্বর ঠিক নয়।' });
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.sheet);
  if (!sh || sh.getLastRow() < 2) return out({ ok: false, message: 'এই নম্বরে কোনো আবেদন পাওয়া যায়নি।' });
  const hit = sh.getRange(2, 1, sh.getLastRow() - 1, 1).createTextFinder(ref).matchEntireCell(true).findNext();
  if (!hit) return out({ ok: false, message: 'এই নম্বরে কোনো আবেদন পাওয়া যায়নি।' });
  const r = sh.getRange(hit.getRow(), 1, 1, 5).getValues()[0];
  return out({ ok: true, status: r[3], note: r[4] });   // শুধু স্ট্যাটাস ও বার্তা, ব্যক্তিগত তথ্য নয়
}
