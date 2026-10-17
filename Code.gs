const EMAIL_TO = "nussaiba31@gmail.com";
const SPOTS = 10;                                   // عدد الأماكن لكل حصة (السيرفر هو اللي بقرر، مش المتصفح)
const COACHES = ["Sara", "Lina", "Hala", "Dana"];   // أسماء الكوتشات المسموحة للتقييم

/* لازم يطابق جدول الحصص بـ index.html. إذا غيّرتي الجدول هناك، غيّريه هون كمان. */
const SCHEDULE = {
  "السبت":    [["Zumba", "18:00"], ["Pilates", "19:15"], ["TRX", "20:30"]],
  "الأحد":    [["Functional training", "17:30"], ["Barre pilates", "18:45"], ["Oriental dance", "20:00"]],
  "الاثنين":  [["Spinning", "17:30"], ["Body bump", "18:30"], ["Core muscles", "19:45"]],
  "الثلاثاء": [["Bands", "17:30"], ["Steps", "18:30"], ["Zumba", "19:45"]],
  "الأربعاء": [["Gym ball", "17:30"], ["Dumbbell sculpt", "18:30"], ["Gymnastics", "19:45"]],
  "الخميس":   [["Sticks", "17:30"], ["Ankle weights", "18:30"], ["Pilates", "19:30"]],
  "الجمعة":   []
};
const DOW = { "الأحد": 0, "الاثنين": 1, "الثلاثاء": 2, "الأربعاء": 3, "الخميس": 4, "الجمعة": 5, "السبت": 6 };

/* ---------- helpers ---------- */

function tabName_(cls) {
  var n = String(cls).split("").filter(function (ch) {
    return "[]*?:/\\".indexOf(ch) === -1;
  }).join("").trim().slice(0, 90);
  return n || "Class";
}

function sheetFor_(cls) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = tabName_(cls);
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(["Time", "Key", "Class", "Day", "ClassTime", "Name", "Phone"]);
  }
  return sh;
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function digits_(s) {
  return String(s).split("").filter(function (ch) { return ch >= "0" && ch <= "9"; }).join("");
}

function norm_(s) {
  return String(s).split(" ").filter(function (w) { return w; }).join(" ").toLowerCase();
}

/* رقم أردني: 07[789]xxxxxxx (بيقبل كمان 962 / 00962). بيرجّع null إذا مش صحيح */
function phone_(s) {
  var d = digits_(s);
  if (d.indexOf("00962") === 0) d = d.slice(5);
  else if (d.indexOf("962") === 0) d = d.slice(3);
  if (d.charAt(0) === "7") d = "0" + d;
  return /^07[789]\d{7}$/.test(d) ? d : null;
}

/* بيمنع أي نص يبلش بـ = + - @ من إنو يتنفذ كمعادلة بالشيت */
function safe_(s) {
  s = String(s);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

/* حد أقصى للطلبات خلال فترة. بيرجّع false إذا تجاوز الحد */
function throttle_(key, max, seconds) {
  var cache = CacheService.getScriptCache();
  var n = Number(cache.get(key) || 0);
  if (n >= max) return false;
  cache.put(key, String(n + 1), seconds);
  return true;
}

/* التواريخ المسموحة للحجز: من اليوم لـ 7 أيام قدام */
function allowedDates_() {
  var tz = Session.getScriptTimeZone(), list = [];
  for (var i = 0; i <= 7; i++) {
    list.push(Utilities.formatDate(new Date(Date.now() + i * 86400000), tz, "yyyy-MM-dd"));
  }
  return list;
}

/* بيتأكد إنو الحجز لحصة حقيقية بالجدول وبتاريخ صحيح */
function validBooking_(p) {
  var key = String(p.key || ""), parts = key.split("|");
  if (parts.length < 3) return false;
  var date = parts[0], time = parts[1], cls = parts.slice(2).join("|");
  if (cls !== p.cls || time !== p.time) return false;
  var list = SCHEDULE[p.day];
  if (!list) return false;
  var found = list.some(function (c) { return c[0] === cls && c[1] === time; });
  if (!found) return false;
  if (allowedDates_().indexOf(date) < 0) return false;
  var d = new Date(date + "T12:00:00");
  return d.getDay() === DOW[p.day];
}

/* ---------- entry point ---------- */

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === "counts") return out_({ ok: true, counts: counts_() });
  if (p.action === "book") return out_(book_(p));
  if (p.action === "rate") return out_(rate_(p));
  return out_({ ok: false, error: "bad request" });
}

function counts_() {
  var c = {};
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(function (sh) {
    var vals = sh.getDataRange().getValues();
    if (vals.length < 2 || vals[0][0] !== "Time") return;
    vals.slice(1).forEach(function (r) { c[r[1]] = (c[r[1]] || 0) + 1; });
  });
  return c;
}

/* ---------- booking ---------- */

function book_(p) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (!validBooking_(p)) return { ok: false, error: "bad request" };

    var phone = phone_(p.phone);
    var name = String(p.name || "").trim().slice(0, 60);
    if (!phone || name.length < 2) return { ok: false, error: "bad request" };

    // حد أقصى: 5 محاولات لكل رقم كل 10 دقايق، و80 محاولة للكل كل 10 دقايق
    if (!throttle_("bp_" + phone.slice(-9), 5, 600) || !throttle_("b_all", 80, 600)) {
      return { ok: false, error: "busy" };
    }

    var sh = sheetFor_(p.cls);
    var rows = sh.getDataRange().getValues().slice(1);
    var taken = rows.filter(function (r) { return r[1] === p.key; });

    if (taken.length >= SPOTS) return { ok: false, error: "full" };

    var dup = taken.some(function (r) {
      return digits_(r[6]).slice(-9) === phone.slice(-9) || norm_(r[5]) === norm_(name);
    });
    if (dup) return { ok: false, error: "dup" };

    sh.appendRow([new Date(), p.key, p.cls, p.day, p.time, safe_(name), "'" + phone]);
    var left = SPOTS - taken.length - 1;

    try {
      MailApp.sendEmail(
        EMAIL_TO,
        "حجز جديد - " + p.cls,
        "الحصة: " + p.cls + "\nاليوم: " + p.day + " " + p.time +
        "\nالاسم: " + name + "\nالرقم: " + phone +
        "\nالأماكن المتبقية: " + left
      );
    } catch (err) {}

    return { ok: true, left: left };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- coach rating ---------- */

function rate_(p) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var coach = String(p.coach || "").trim();
    var stars = Number(p.stars);
    if (COACHES.indexOf(coach) < 0 || !(stars >= 1 && stars <= 5) || stars % 1 !== 0) {
      return { ok: false, error: "bad request" };
    }

    // حد أقصى 40 تقييم بالساعة للكل
    if (!throttle_("r_all", 40, 3600)) return { ok: false, error: "busy" };

    var comment = String(p.comment || "").trim().slice(0, 300);
    var name = String(p.name || "").trim().slice(0, 60);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName("Ratings");
    if (!sh) {
      sh = ss.insertSheet("Ratings");
      sh.appendRow(["Date", "Coach", "Stars", "Comment", "Name"]);
    }
    sh.appendRow([new Date(), coach, stars, safe_(comment), safe_(name)]);

    try {
      MailApp.sendEmail(
        EMAIL_TO,
        "تقييم جديد - " + coach + " (" + stars + "/5)",
        "الكوتش: " + coach + "\nالتقييم: " + stars + " من 5" +
        (comment ? "\nالتعليق: " + comment : "") +
        (name ? "\nالاسم: " + name : "")
      );
    } catch (err) {}

    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}
