// ==================== ส่วนเสริมระดับ Production (v3) ====================
const VERSION = "3.2.0";
const AUDIT_SHEET = "_audit";
const ERROR_SHEET = "_errors";
const PENDING_DAYS = 7;
const AI_LIMIT_PER_MIN = 6;
const BACKUP_KEEP = 8;
const OPEN_STATUSES = ["รอซ่อม", "เสีย", "รอช่างนัด"];

var ACTOR_ = "";
var CHAT_ = "";

function sysSheet_(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.appendRow(header);
    s.setFrozenRows(1);
    s.hideSheet();
  }
  return s;
}

function audit_(actor, chatId, text) {
  try {
    const s = sysSheet_(AUDIT_SHEET, ["เวลา", "แชท", "ผู้ใช้", "ข้อความ"]);
    s.appendRow([
      new Date(),
      String(chatId),
      actor,
      String(text).substring(0, 200),
    ]);
    if (s.getLastRow() > 5000) s.deleteRows(2, 1000);
  } catch (e) {}
}

function notifyAdmin_(msg) {
  try {
    sendTelegramMessage(
      msg,
      prop_("ADMIN_CHAT_ID") || prop_("TELEGRAM_CHAT_ID"),
    );
  } catch (e) {}
}

function logError_(where, err) {
  try {
    const m = String((err && err.message) || err);
    sysSheet_(ERROR_SHEET, ["เวลา", "จุดที่พัง", "ข้อความ", "stack"]).appendRow(
      [
        new Date(),
        where,
        m,
        String((err && err.stack) || "").substring(0, 500),
      ],
    );
    const c = CacheService.getScriptCache();
    if (!c.get("err_" + where)) {
      c.put("err_" + where, "1", 3600);
      notifyAdmin_(
        "🚨 <b>ระบบแอร์ผิดพลาด</b> (" + esc_(where) + ")\n" + esc_(m),
      );
    }
  } catch (e) {}
}

function safe_(name, fn) {
  try {
    fn();
  } catch (e) {
    logError_(name, e);
  }
}

function isAllowed_(chatId) {
  return (
    (prop_("ALLOWED_CHAT_IDS") || "")
      .split(",")
      .concat([prop_("TELEGRAM_CHAT_ID") || ""])
      .map(function (s) {
        return s.trim();
      })
      .filter(String)
      .indexOf(chatId) !== -1
  );
}

function aiGuard_(who) {
  const c = CacheService.getScriptCache();
  const k = "ai_" + who;
  const n = Number(c.get(k) || 0);
  if (n >= AI_LIMIT_PER_MIN)
    throw new Error("ถาม AI ถี่เกินไป รอสักครู่แล้วลองใหม่");
  c.put(k, String(n + 1), 60);
}

function tg_(method, body) {
  const url =
    "https://api.telegram.org/bot" + prop_("TELEGRAM_TOKEN") + "/" + method;
  let last;
  for (let i = 0; i < 3; i++) {
    try {
      const r = UrlFetchApp.fetch(url, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(body),
        muteHttpExceptions: true,
      });
      const code = r.getResponseCode();
      if (code === 200) return JSON.parse(r.getContentText());
      last = new Error(
        "Telegram " + code + ": " + r.getContentText().substring(0, 200),
      );
      if (code !== 429 && code < 500) break;
    } catch (e) {
      last = e;
    }
    Utilities.sleep(800 * (i + 1));
  }
  throw last;
}

// ชื่อบอท (cache 6 ชม.) ใช้ตรวจว่าข้อความในกลุ่มพูดกับบอทไหม
function botUsername_() {
  const c = CacheService.getScriptCache();
  let u = c.get("bot_username");
  if (!u) {
    const r = tg_("getMe", {});
    u = (r.result && r.result.username) || "";
    if (u) c.put("bot_username", u, 21600);
  }
  return u;
}

function reply_(chatId, r) {
  if (!r) return;
  if (typeof r === "string") sendTelegramMessage(r, chatId);
  else sendTelegramMessage(r.text, chatId, r.markup);
}

function btn_(text, data) {
  return { text: text, callback_data: data };
}
function undoKeyboard_() {
  return { inline_keyboard: [[btn_("↩️ ย้อนกลับ", "u:")]] };
}
function acKeyboard_(id) {
  return {
    inline_keyboard: [
      [btn_("🧼 ล้างแล้ว", "w:" + id), btn_("✅ ซ่อมเสร็จ", "f:" + id)],
    ],
  };
}
function dueKeyboard_() {
  const list = getAirs_()
    .filter(function (a) {
      return a.alert === "เลยกำหนด" || a.alert === "ใกล้ถึงกำหนด";
    })
    .slice(0, 8);
  if (!list.length) return undefined;
  const rows = [];
  list.forEach(function (a, i) {
    if (i % 2 === 0) rows.push([]);
    rows[rows.length - 1].push(btn_("🧼 ล้างแล้ว " + a.id, "w:" + a.id));
  });
  return { inline_keyboard: rows };
}

function handleCallback_(cb, chatId) {
  // ตอบ Telegram ทันทีเพื่อให้ปุ่มหยุดหมุน
  try {
    tg_("answerCallbackQuery", { callback_query_id: cb.id, text: "รับทราบ" });
  } catch (e) {}

  const p = String(cb.data || "").split(":");
  let res;
  let undo = false;
  try {
    if (p[0] === "w") {
      res = cmdWash_([p[1]]);
      undo = true;
    } else if (p[0] === "f") {
      res = cmdFix_([p[1], "ซ่อมเสร็จ (กดปุ่ม)"]);
      undo = true;
    } else if (p[0] === "u") {
      res = cmdUndo_(chatId);
    } else if (p[0] === "a") {
      const id = normId_(p[1]);
      res = { text: cmdAc_(id), markup: acKeyboard_(id) };
    } else if (p[0] === "m") {
      if (p[1] === "status") res = cmdStatus_();
      else if (p[1] === "due")
        res = {
          text: buildAlertMessage_() + pendingNote_() + techNote_(),
          markup: dueKeyboard_(),
        };
      else if (p[1] === "list") res = cmdList_();
      else if (p[1] === "cost") res = cmdCost_("");
      else if (p[1] === "tech") res = cmdTech_();
      else if (p[1] === "power") res = cmdPower_();
      else res = "ไม่พบเมนูนี้";
    } else {
      res = "ไม่รู้จักปุ่มนี้";
    }
    audit_(ACTOR_, chatId, "[ปุ่ม] " + cb.data);
  } catch (e) {
    res = "❌ " + esc_(e.message);
    undo = false;
  }

  if (res && typeof res === "object") {
    sendTelegramMessage(res.text, chatId, res.markup);
  } else {
    sendTelegramMessage(res, chatId, undo ? undoKeyboard_() : undefined);
  }
}

function undoPush_(row, id, type) {
  if (!CHAT_) return;
  CacheService.getScriptCache().put(
    "undo_" + CHAT_,
    JSON.stringify({ row: row, id: id, type: type }),
    21600,
  );
}
function undoStatus_(id, prev) {
  if (!CHAT_) return;
  const c = CacheService.getScriptCache();
  const raw = c.get("undo_" + CHAT_);
  if (!raw) return;
  const r = JSON.parse(raw);
  if (r.id === id && r.prev === undefined) {
    r.prev = prev;
    c.put("undo_" + CHAT_, JSON.stringify(r), 21600);
  }
}

function cmdUndo_(chatId) {
  const c = CacheService.getScriptCache();
  const raw = c.get("undo_" + chatId);
  if (!raw)
    throw new Error(
      "ไม่มีรายการให้ย้อนกลับ (ย้อนได้เฉพาะรายการล่าสุด ภายใน 6 ชั่วโมง)",
    );
  const r = JSON.parse(raw);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_(LOG_SHEET);
    const row = sh.getRange(r.row, 1, 1, 8).getValues()[0];
    if (
      normId_(row[LOG.ID]) !== normId_(r.id) ||
      String(row[LOG.TYPE]) !== r.type
    )
      throw new Error(
        "แถวในชีตถูกแก้ไขไปแล้ว ย้อนกลับอัตโนมัติไม่ได้ กรุณาแก้ในชีตเอง",
      );
    sh.getRange(r.row, 1, 1, 8).clearContent();
    if (r.prev !== undefined) setStatus_(r.id, r.prev);
  } finally {
    lock.releaseLock();
  }
  c.remove("undo_" + chatId);
  return "↩️ ย้อนรายการ <b>" + esc_(r.id) + "</b> (" + esc_(r.type) + ") แล้ว";
}

function cmdList_() {
  return (
    "📋 <b>แอร์ทั้งหมด</b>\n" +
    getAirs_()
      .map(function (a) {
        return (
          (ALERT_ICON[a.alert] || "🟢") +
          " <b>" +
          esc_(a.id) +
          "</b> " +
          esc_(a.loc || "-") +
          " — " +
          esc_(a.status || "-")
        );
      })
      .join("\n")
  );
}

function cmdCost_(arg) {
  const key = /^\d{4}-\d{2}$/.test(arg || "")
    ? arg
    : Utilities.formatDate(new Date(), TZ, "yyyy-MM");
  let total = 0;
  const by = {};
  getLogs_().forEach(function (l) {
    if (Utilities.formatDate(l.date, TZ, "yyyy-MM") !== key) return;
    total += l.cost;
    by[l.id] = (by[l.id] || 0) + l.cost;
  });
  let t = "💰 <b>ค่าใช้จ่ายเดือน " + key + "</b>: " + baht_(total) + " บาท\n";
  Object.keys(by)
    .filter(function (k) {
      return by[k];
    })
    .sort(function (a, b) {
      return by[b] - by[a];
    })
    .forEach(function (id) {
      t += "• " + esc_(id) + ": " + baht_(by[id]) + "\n";
    });
  return t;
}

function cmdHistory_(id) {
  const logs = getLogs_()
    .filter(function (l) {
      return !id || l.id === id;
    })
    .slice(0, 10);
  if (!logs.length) return "ยังไม่มีประวัติ";
  return (
    "🧾 <b>ประวัติล่าสุด" +
    (id ? " " + esc_(id) : "") +
    "</b>\n" +
    logs
      .map(function (l) {
        return (
          "• " +
          fmt_(l.date) +
          " <b>" +
          esc_(l.id) +
          "</b> " +
          esc_(l.type) +
          (l.detail ? " — " + esc_(l.detail) : "") +
          (l.cost ? " (" + baht_(l.cost) + "฿)" : "")
        );
      })
      .join("\n")
  );
}

function pendingList_() {
  const logs = getLogs_();
  const out = [];
  getAirs_().forEach(function (a) {
    if (OPEN_STATUSES.indexOf(a.status) === -1) return;
    const l = logs.filter(function (x) {
      return x.id === a.id && x.type === "แจ้งปัญหา";
    })[0];
    const days = l
      ? Math.floor((Date.now() - l.date.getTime()) / 86400000)
      : null;
    if (days === null || days >= PENDING_DAYS)
      out.push({ id: a.id, status: a.status, days: days });
  });
  return out;
}

function pendingNote_() {
  const p = pendingList_();
  if (!p.length) return "";
  return (
    "\n⏳ <b>ค้างซ่อมนาน</b>\n" +
    p
      .map(function (x) {
        return (
          "• <b>" +
          esc_(x.id) +
          "</b> " +
          esc_(x.status) +
          (x.days === null ? "" : " (" + x.days + " วัน)")
        );
      })
      .join("\n")
  );
}

function weeklyReport() {
  safe_("weekly", function () {
    const airs = getAirs_();
    const bad = airs.filter(function (a) {
      return a.status && a.status !== "ปกติ";
    }).length;
    sendTelegramMessage(
      "📅 <b>สรุปประจำสัปดาห์</b>\nแอร์ " +
        airs.length +
        " เครื่อง | ปกติ " +
        (airs.length - bad) +
        " | มีปัญหา " +
        bad +
        "\n\n" +
        cmdCost_("") +
        "\n" +
        buildAlertMessage_() +
        pendingNote_(),
    );
  });
}

function weeklyBackup() {
  safe_("backup", function () {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const file = DriveApp.getFileById(ss.getId());
    const fid = prop_("BACKUP_FOLDER_ID");
    const prefix = ss.getName() + " - backup ";
    const name = prefix + Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd");
    let folder;
    if (fid) folder = DriveApp.getFolderById(fid);
    else {
      const ps = file.getParents();
      folder = ps.hasNext() ? ps.next() : DriveApp.getRootFolder();
    }
    file.makeCopy(name, folder);
    const olds = [];
    const it = folder.getFiles();
    while (it.hasNext()) {
      const f = it.next();
      if (f.getName().indexOf(prefix) === 0) olds.push(f);
    }
    olds
      .sort(function (a, b) {
        return b.getDateCreated() - a.getDateCreated();
      })
      .slice(BACKUP_KEEP)
      .forEach(function (f) {
        f.setTrashed(true);
      });
  });
}

function healthCheck(quick) {
  const out = [];
  function t(name, fn) {
    try {
      out.push("✅ " + name + (fn() || ""));
    } catch (e) {
      out.push("❌ " + name + ": " + esc_(e.message));
    }
  }
  [
    "TELEGRAM_TOKEN",
    "TELEGRAM_CHAT_ID",
    "DASH_KEY",
    "WEBHOOK_SECRET",
    "WEBAPP_URL",
  ].forEach(function (k) {
    t("ค่า " + k, function () {
      if (!prop_(k)) throw new Error("ยังไม่ได้ตั้ง");
    });
  });
  t("Telegram bot", function () {
    const r = JSON.parse(
      UrlFetchApp.fetch(
        "https://api.telegram.org/bot" + prop_("TELEGRAM_TOKEN") + "/getMe",
        { muteHttpExceptions: true },
      ).getContentText(),
    );
    if (!r.ok) throw new Error(r.description);
    return " (@" + r.result.username + ")";
  });
  t("ชีตเครื่องแอร์", function () {
    return " (" + getAirs_().length + " เครื่อง)";
  });
  t("ชีตประวัติซ่อม", function () {
    return " (" + getLogs_().length + " รายการ)";
  });
  t("Webhook", function () {
    const r = JSON.parse(
      UrlFetchApp.fetch(
        "https://api.telegram.org/bot" +
          prop_("TELEGRAM_TOKEN") +
          "/getWebhookInfo",
      ).getContentText(),
    ).result;
    if (!r.url) throw new Error("ยังไม่ได้ตั้ง webhook (รัน setupAll)");
    const base = prop_("WEBAPP_URL") || "";
    if (base && r.url.indexOf(base) !== 0)
      throw new Error("webhook ชี้ไปลิงก์อื่น (รัน setupAll ใหม่)");
    // Apps Script ตอบ 302 เป็นเรื่องปกติ ไม่นับเป็นข้อผิดพลาด; แจ้งเฉพาะ error ใน 1 ชม.ล่าสุด
    const recent =
      r.last_error_date && Date.now() / 1000 - r.last_error_date < 3600;
    if (r.last_error_message && recent && !/302/.test(r.last_error_message))
      throw new Error(r.last_error_message);
    return " (ค้างส่ง " + r.pending_update_count + ")";
  });
  t("AI key (Groq/Gemini)", function () {
    const g = [];
    if (prop_("GROQ_API_KEY")) g.push("Groq");
    if (prop_("GEMINI_API_KEY")) g.push("Gemini");
    if (!g.length)
      throw new Error("ยังไม่ได้ตั้ง GROQ_API_KEY หรือ GEMINI_API_KEY");
    return " (" + g.join(" + ") + ")";
  });
  t("AI ตอบได้", function () {
    if (quick) return " (ข้าม — พิมพ์ /ask สวัสดี เพื่อทดสอบ)";
    const t0 = Date.now();
    askLLM_("ตอบคำว่า OK คำเดียว");
    const c = CacheService.getScriptCache();
    const p = c.get("ai_provider_ok") || "?";
    return (
      " (" +
      p +
      ": " +
      (c.get(p + "_ok") || "?") +
      ", " +
      (Date.now() - t0) +
      " ms)"
    );
  });
  t("Trigger", function () {
    const n = ScriptApp.getProjectTriggers().length;
    if (!n) throw new Error("ยังไม่มี trigger (รัน setupAll)");
    return " (" + n + ")";
  });
  const msg = "🩺 <b>Health check v" + VERSION + "</b>\n" + out.join("\n");
  Logger.log(msg.replace(/<[^>]+>/g, ""));
  return msg;
}

// รันครั้งเดียวหลัง Deploy: ตั้ง trigger + webhook + เมนูคำสั่งของบอท
function setupAll() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("checkAndSendAirNotification")
    .timeBased()
    .atHour(8)
    .nearMinute(0)
    .everyDays(1)
    .inTimezone(TZ)
    .create();
  ScriptApp.newTrigger("weeklyReport")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .nearMinute(30)
    .inTimezone(TZ)
    .create();
  ScriptApp.newTrigger("weeklyBackup")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(2)
    .inTimezone(TZ)
    .create();
  setWebhook();
  setBotCommands();
  Logger.log(healthCheck().replace(/<[^>]+>/g, ""));
}

// เมนู "/" ในช่องพิมพ์ของ Telegram
function setBotCommands() {
  tg_("setMyCommands", {
    commands: [
      { command: "menu", description: "เมนูด่วน" },
      { command: "status", description: "สรุปภาพรวมแอร์" },
      { command: "due", description: "แอร์ที่ถึงรอบล้าง/ต้องดูแล" },
      { command: "list", description: "รายการแอร์ทั้งหมด" },
      { command: "wash", description: "บันทึกล้างแอร์" },
      { command: "issue", description: "แจ้งปัญหาแอร์" },
      { command: "fix", description: "บันทึกซ่อมเสร็จ" },
      { command: "history", description: "ประวัติล่าสุด" },
      { command: "cost", description: "ค่าใช้จ่ายรายเดือน" },
      { command: "tech", description: "รายชื่อช่าง" },
      { command: "power", description: "สถิติค่าไฟ" },
      { command: "ask", description: "ถาม AI Gemini เรื่องแอร์" },
      { command: "dashboard", description: "ลิงก์ Dashboard" },
      { command: "health", description: "ตรวจสุขภาพระบบ" },
      { command: "help", description: "วิธีใช้ทั้งหมด" },
    ],
  });
}

// เรียกจาก Editor เพื่อทดสอบ AI ตรง ๆ (ดูผลที่ Execution log) — ทดสอบทีละเจ้า + ตัวกลาง
function testAI() {
  ["Groq", "Gemini"].forEach(function (name) {
    if (!prop_(name.toUpperCase() + "_API_KEY")) {
      Logger.log(name + ": ข้าม (ไม่มีคีย์)");
      return;
    }
    const t0 = Date.now();
    try {
      const a =
        name === "Groq"
          ? askGroq_("แอร์เย็นน้อยเกิดจากอะไรได้บ้าง ตอบสั้น ๆ 2 ข้อ")
          : askGemini_("แอร์เย็นน้อยเกิดจากอะไรได้บ้าง ตอบสั้น ๆ 2 ข้อ");
      const c = CacheService.getScriptCache();
      Logger.log(
        name +
          " ✅ " +
          c.get(name.toLowerCase() + "_ok") +
          " (" +
          (Date.now() - t0) +
          " ms)\n" +
          a,
      );
    } catch (e) {
      Logger.log(name + " ❌ " + e.message);
    }
  });
}

function debugWebhook() {
  const r = JSON.parse(
    UrlFetchApp.fetch(
      "https://api.telegram.org/bot" +
        prop_("TELEGRAM_TOKEN") +
        "/getWebhookInfo",
    ).getContentText(),
  ).result;
  r.url = String(r.url || "").replace(/\?s=.*/, "?s=***");
  Logger.log(JSON.stringify(r, null, 2));
  Logger.log("getUrl() = " + ScriptApp.getService().getUrl());
  Logger.log("WEBAPP_URL = " + prop_("WEBAPP_URL"));
  Logger.log(
    "ALLOWED_CHAT_IDS = " +
      prop_("ALLOWED_CHAT_IDS") +
      " | TELEGRAM_CHAT_ID = " +
      prop_("TELEGRAM_CHAT_ID"),
  );
}

function typing_(chatId) {
  try {
    tg_("sendChatAction", { chat_id: chatId, action: "typing" });
  } catch (e) {}
}

function seen_(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = PropertiesService.getScriptProperties();
    const arr = JSON.parse(p.getProperty("SEEN_UPDATES") || "[]");
    if (arr.indexOf(id) !== -1) return true;
    arr.push(id);
    p.setProperty("SEEN_UPDATES", JSON.stringify(arr.slice(-300)));
    return false;
  } finally {
    lock.releaseLock();
  }
}

function mainKeyboard_() {
  return {
    keyboard: [
      [{ text: "📊 /status" }, { text: "🔔 /due" }, { text: "📋 /list" }],
      [{ text: "🧼 /wash" }, { text: "🛠️ /history" }, { text: "💰 /cost" }],
      [{ text: "📞 /tech" }, { text: "⚡ /power" }, { text: "🩺 /health" }],
      [{ text: "🤖 /ask" }, { text: "📈 /dashboard" }, { text: "❓ /help" }],
    ],
    resize_keyboard: true,
    persistent: true,
  };
}

function cmdMenu_() {
  return {
    text: "🤖 <b>เมนูด่วนระบบดูแลแอร์</b>\nเลือกคำสั่งที่ต้องการได้เลยครับ 👇",
    markup: {
      inline_keyboard: [
        [
          { text: "📊 สรุปภาพรวม", callback_data: "m:status" },
          { text: "🔔 เช็คดิวล้าง", callback_data: "m:due" },
        ],
        [
          { text: "📋 รายการแอร์", callback_data: "m:list" },
          { text: "💰 ค่าใช้จ่าย", callback_data: "m:cost" },
        ],
        [
          { text: "📞 รายชื่อช่าง", callback_data: "m:tech" },
          { text: "⚡ สถิติค่าไฟ", callback_data: "m:power" },
        ],
      ],
    },
  };
}

// ==================== จัดการเครื่องแอร์จาก Dashboard (เพิ่ม/แก้ไข/ลบ) ====================
// หาแถวของเครื่องในชีต 'เครื่องแอร์' จากรหัส (รวมแถวที่ใส่รหัสไว้ล่วงหน้าแต่ยังไม่กรอกข้อมูล)
function airRowOf_(sh, id) {
  const n = sh.getLastRow() - FIRST_DATA_ROW + 1;
  if (n < 1) return 0;
  const ids = sh.getRange(FIRST_DATA_ROW, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++)
    if (normId_(ids[i][0]) === id) return FIRST_DATA_ROW + i;
  return 0;
}

// แถวว่างถัดไป (ช่องรหัสเครื่องว่าง) ถ้าไม่มีให้ต่อท้าย
function nextAirRow_(sh) {
  const n = sh.getLastRow() - FIRST_DATA_ROW + 1;
  if (n < 1) return FIRST_DATA_ROW;
  const ids = sh.getRange(FIRST_DATA_ROW, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++)
    if (!String(ids[i][0]).trim()) return FIRST_DATA_ROW + i;
  return FIRST_DATA_ROW + n;
}

// ถ้าแถวนี้ไม่มีสูตร (รูป B + คอลัมน์ I–M) ให้คัดลอกจากแถวบนที่มีสูตรอยู่
function ensureAirFormulas_(sh, row) {
  const colM = COL.ALERT + 1;
  if (sh.getRange(row, colM).getFormula()) return;
  for (let r = row - 1; r >= FIRST_DATA_ROW; r--) {
    if (sh.getRange(r, colM).getFormula()) {
      sh.getRange(r, 2).copyTo(sh.getRange(row, 2));
      sh.getRange(r, COL.LAST_WASH + 1, 1, 5).copyTo(
        sh.getRange(row, COL.LAST_WASH + 1, 1, 5),
      );
      return;
    }
  }
}

function saveAir_(o, actor) {
  const id = normId_(clean_(o.id, 20));
  if (!/^[A-Z0-9][A-Z0-9-]{1,19}$/.test(id))
    throw userErr_("รหัสเครื่องไม่ถูกต้อง (เช่น AC-16)");
  const loc = clean_(o.loc, 100);
  if (!loc) throw userErr_("กรุณาใส่ตำแหน่ง/ห้อง");
  const model = clean_(o.model, 80);
  const btu = cleanNum_(o.btu, 3000, 1000000, "BTU", false);
  const status = clean_(o.status, 30) || "ปกติ";
  if (STATUSES.indexOf(status) === -1)
    throw userErr_("สถานะต้องเป็น: " + STATUSES.join(" / "));
  return withLock_(function () {
    const sh = sheet_(SHEET_NAME);
    let row = airRowOf_(sh, id);
    const isNew = !row;
    if (isNew) {
      row = nextAirRow_(sh);
      sh.getRange(row, COL.ID + 1)
        .setNumberFormat("@")
        .setValue(id);
    }
    sh.getRange(row, COL.LOCATION + 1).setValue(loc);
    sh.getRange(row, COL.MODEL + 1).setValue(model);
    sh.getRange(row, COL.BTU + 1).setValue(btu === null ? "" : btu);
    sh.getRange(row, COL.STATUS + 1).setValue(status);
    ensureAirFormulas_(sh, row);
    SpreadsheetApp.flush();
    audit_(actor, "web", "[แอร์] " + (isNew ? "เพิ่ม " : "แก้ไข ") + id);
    return { id: id, isNew: isNew };
  });
}

// "ลบ" = ล้างข้อมูลเครื่อง (ตำแหน่ง/ยี่ห้อ/รุ่น/BTU/วันติดตั้ง/สถานะ) แต่คงรหัสและประวัติซ่อมไว้
// เครื่องจะหายจากบอทและ Dashboard และเพิ่มกลับด้วยรหัสเดิมได้
function deleteAir_(id, actor) {
  const cid = normId_(clean_(id, 20));
  return withLock_(function () {
    const sh = sheet_(SHEET_NAME);
    const row = airRowOf_(sh, cid);
    if (!row) throw userErr_("ไม่พบเครื่อง " + cid);
    sh.getRange(
      row,
      COL.LOCATION + 1,
      1,
      COL.STATUS - COL.LOCATION + 1,
    ).clearContent();
    SpreadsheetApp.flush();
    audit_(actor, "web", "[แอร์] ลบ " + cid);
    return null;
  });
}

function apiAirSave(key, o) {
  return api_("airSave", key, true, function () {
    saveAir_(o || {}, "dashboard");
    return getDashboardData(key);
  });
}
function apiAirDelete(key, id) {
  return api_("airDelete", key, true, function () {
    deleteAir_(id, "dashboard");
    return getDashboardData(key);
  });
}

// รันครั้งเดียวจาก Editor: เติมสูตรรูป + คอลัมน์ I–M ให้ทุกแถวเครื่องที่มีรหัสแต่ยังไม่มีสูตร (เช่น AC-10 ลงไป)
function fixAirFormulas() {
  const sh = sheet_(SHEET_NAME);
  const n = sh.getLastRow() - FIRST_DATA_ROW + 1;
  if (n < 1) return;
  const ids = sh.getRange(FIRST_DATA_ROW, 1, n, 1).getValues();
  let fixed = 0;
  ids.forEach(function (r, i) {
    const row = FIRST_DATA_ROW + i;
    if (!String(r[0]).trim()) return;
    if (sh.getRange(row, COL.ALERT + 1).getFormula()) return;
    ensureAirFormulas_(sh, row);
    fixed++;
  });
  SpreadsheetApp.flush();
  Logger.log("เติมสูตรให้ " + fixed + " แถว");
}
