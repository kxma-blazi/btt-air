// ==================== ระบบดูแลรักษาแอร์ (Sheets + Telegram + Groq/Gemini + Dashboard) — v3.1 Production ====================
// Script properties ที่ต้องตั้ง (Project Settings > Script properties):
//   TELEGRAM_TOKEN    = โทเคนบอทจาก BotFather
//   TELEGRAM_CHAT_ID  = รหัสแชท/กลุ่มหลัก (ใช้ส่งแจ้งเตือน + เป็นแชทที่อนุญาตสั่งบอท) เช่น -1003743434778
//   DASH_KEY          = รหัสลับอะไรก็ได้ ใช้ต่อท้าย URL ของ Dashboard (?key=...)
//   WEBHOOK_SECRET    = รหัสลับอะไรก็ได้ ใช้ป้องกัน webhook
//   WEBAPP_URL        = ลิงก์ Web app ที่ลงท้าย /exec (จากหน้า Deploy)
//   GROQ_API_KEY      = คีย์จาก console.groq.com (ฟรี) — แนะนำเป็นตัวหลัก
//   GEMINI_API_KEY    = คีย์จาก Google AI Studio (ตัวสำรอง)
//   ALLOWED_CHAT_IDS  = (ไม่บังคับ) รหัสแชทที่อนุญาตเพิ่ม คั่นด้วย ,
//   ADMIN_CHAT_ID     = (ไม่บังคับ) แชทที่รับแจ้งเตือนระบบพัง (ไม่ตั้ง = ใช้ TELEGRAM_CHAT_ID)
// (ห้ามพิมพ์โทเคน/คีย์ลงในโค้ดโดยตรง)

const SHEET_NAME = "เครื่องแอร์";
const LOG_SHEET = "ประวัติซ่อม";
const FIRST_DATA_ROW = 5; // ชีตเครื่องแอร์
const LOG_FIRST_ROW = 4; // ชีตประวัติซ่อม
const TZ = "Asia/Bangkok";
const WASH_INTERVAL_MONTHS = 6; // ล้างแอร์ทุกกี่เดือน (ใช้คำนวณนัดครั้งถัดไป)
const GEMINI_MODEL = "gemini-3.1-flash-lite"; // ถ้าโมเดลถูกปลดระวาง ให้ตั้ง Script property GEMINI_MODEL ทับ

// ลำดับคอลัมน์ชีต "เครื่องแอร์" (เริ่มนับ 0)
const COL = {
  ID: 0,
  LOCATION: 2,
  MODEL: 3,
  BTU: 5,
  STATUS: 7,
  LAST_WASH: 8,
  NEXT_DATE: 9,
  COUNT: 10,
  COST: 11,
  ALERT: 12,
};
// ลำดับคอลัมน์ชีต "ประวัติซ่อม"
const LOG = {
  DATE: 0,
  ID: 1,
  TYPE: 2,
  DETAIL: 3,
  TECH: 4,
  COST: 5,
  NEXT: 6,
  NOTE: 7,
};

const STATUSES = ["ปกติ", "เย็นน้อย", "เสีย", "รอซ่อม", "รอช่างนัด"];
const ALERT_LEVELS = [
  "เลยกำหนด",
  "ใกล้ถึงกำหนด",
  "ตรวจสอบสถานะ",
  "ยังไม่มีนัด",
];
const ALERT_ICON = {
  เลยกำหนด: "🔴",
  ใกล้ถึงกำหนด: "🟠",
  ตรวจสอบสถานะ: "🟡",
  ยังไม่มีนัด: "⚪",
};

// ==================== Helpers ====================
function esc_(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// อ่านค่าลับจาก Script properties เท่านั้น (ไม่ฝังในโค้ด)
function prop_(k) {
  return PropertiesService.getScriptProperties().getProperty(k);
}

function sheet_(name) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!s) throw new Error("ไม่พบชีต " + name);
  return s;
}
function fmt_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, "dd/MM/yyyy");
  return v === null || v === undefined ? "" : String(v);
}
function addMonths_(d, n) {
  const r = new Date(d);
  r.setMonth(r.getMonth() + n);
  return r;
}
function normId_(s) {
  return String(s || "")
    .trim()
    .toUpperCase()
    .replace(/^AC-?(\d+)$/, function (_, n) {
      return "AC-" + ("0" + n).slice(-2);
    });
}
function baht_(n) {
  return Number(n || 0).toLocaleString("th-TH");
}

// ==================== อ่านข้อมูลจากชีต ====================
function getAirs_() {
  const sh = sheet_(SHEET_NAME);
  const n = sh.getLastRow() - FIRST_DATA_ROW + 1;
  if (n < 1) return [];
  const v = sh.getRange(FIRST_DATA_ROW, 1, n, COL.ALERT + 1).getValues();
  const out = [];
  v.forEach(function (r, i) {
    if (!r[COL.ID]) return;
    // ข้ามแถวที่มีแต่รหัสเครื่องล่วงหน้า (เช่น AC-11..AC-15 ที่ยังไม่ได้กรอกตำแหน่ง/รุ่น/BTU)
    if (!r[COL.LOCATION] && !r[COL.MODEL] && !r[COL.BTU]) return;
    out.push({
      row: FIRST_DATA_ROW + i,
      id: String(r[COL.ID]).trim(),
      loc: String(r[COL.LOCATION] || ""),
      model: String(r[COL.MODEL] || ""),
      btu: r[COL.BTU] || "",
      status: String(r[COL.STATUS] || ""),
      lastWash: fmt_(r[COL.LAST_WASH]),
      next: fmt_(r[COL.NEXT_DATE]),
      days:
        r[COL.NEXT_DATE] instanceof Date
          ? Math.ceil((r[COL.NEXT_DATE] - new Date()) / 86400000)
          : null,
      count: Number(r[COL.COUNT]) || 0,
      cost: Number(r[COL.COST]) || 0,
      alert: String(r[COL.ALERT] || ""),
    });
  });
  return out;
}

function getLogs_() {
  const sh = sheet_(LOG_SHEET);
  const n = sh.getLastRow() - LOG_FIRST_ROW + 1;
  if (n < 1) return [];
  const v = sh.getRange(LOG_FIRST_ROW, 1, n, 8).getValues();
  return v
    .filter(function (r) {
      return r[LOG.DATE] instanceof Date && r[LOG.ID];
    })
    .map(function (r) {
      return {
        date: r[LOG.DATE],
        id: String(r[LOG.ID]).trim(),
        type: String(r[LOG.TYPE] || ""),
        detail: String(r[LOG.DETAIL] || ""),
        tech: String(r[LOG.TECH] || ""),
        cost: Number(r[LOG.COST]) || 0,
        next: r[LOG.NEXT],
      };
    })
    .sort(function (a, b) {
      return b.date - a.date;
    });
}

// ==================== เขียนข้อมูลลงชีต ====================
function addLog_(o) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_(LOG_SHEET);
    const last = Math.max(sh.getLastRow(), LOG_FIRST_ROW);
    const col = sh
      .getRange(LOG_FIRST_ROW, 1, last - LOG_FIRST_ROW + 1, 1)
      .getValues();
    let row = last + 1;
    for (let i = 0; i < col.length; i++) {
      if (col[i][0] === "") {
        row = LOG_FIRST_ROW + i;
        break;
      }
    }
    sh.getRange(row, 1, 1, 8).setValues([
      [
        o.date,
        o.id,
        o.type,
        o.detail || "",
        o.tech || "",
        o.cost || "",
        o.next || "",
        o.note || "บันทึกผ่าน Telegram" + (ACTOR_ ? " โดย " + ACTOR_ : ""),
      ],
    ]);
    sh.getRange(row, 1).setNumberFormat("dd/mm/yyyy");
    sh.getRange(row, 7).setNumberFormat("dd/mm/yyyy");
    undoPush_(row, o.id, o.type);
  } finally {
    lock.releaseLock();
  }
}

function setStatus_(id, status) {
  const a = getAirs_().filter(function (x) {
    return x.id === id;
  })[0];
  if (!a) throw new Error("ไม่พบเครื่อง " + id);
  undoStatus_(id, a.status);
  sheet_(SHEET_NAME)
    .getRange(a.row, COL.STATUS + 1)
    .setValue(status);
}

// ==================== Telegram ====================
function sendTelegramMessage(message, chatId, markup) {
  const token = prop_("TELEGRAM_TOKEN");
  const chat = chatId || prop_("TELEGRAM_CHAT_ID");
  if (!token || !chat)
    throw new Error(
      "ยังไม่ได้ตั้งค่า TELEGRAM_TOKEN / TELEGRAM_CHAT_ID ใน Script properties",
    );
  // กันส่งข้อความว่าง/undefined ออกไปในแชท
  if (
    message === undefined ||
    message === null ||
    String(message).trim() === ""
  )
    return null;
  const body = {
    chat_id: chat,
    text: String(message).substring(0, 4000),
    parse_mode: "HTML",
    disable_web_page_preview: true,
  };

  // ถ้ามีการระบุปุ่มเฉพาะกิจ (เช่น Inline Button) ให้ใช้ตัวนั้น
  // ถ้าไม่มี ให้แสดงแผงปุ่มกดหลักด้านล่างอัตโนมัติ
  body.reply_markup = markup ? markup : mainKeyboard_();

  try {
    return tg_("sendMessage", body);
  } catch (e) {
    if (/parse entities/i.test(e.message)) {
      delete body.parse_mode;
      body.text = body.text.replace(/<[^>]+>/g, "");
      return tg_("sendMessage", body);
    }
    throw e;
  }
}

function testTelegram() {
  sendTelegramMessage(
    "🔔 <b>ทดสอบระบบแจ้งเตือนแอร์</b>\nเชื่อมต่อ Telegram เรียบร้อยแล้วครับ",
  );
}

// ==================== แจ้งเตือนประจำวัน ====================
function buildAlertMessage_() {
  const groups = {};
  ALERT_LEVELS.forEach(function (lv) {
    groups[lv] = [];
  });
  getAirs_().forEach(function (a) {
    if (groups[a.alert]) {
      groups[a.alert].push(
        "• <b>" +
          esc_(a.id) +
          "</b> (" +
          esc_(a.loc || "-") +
          ") สถานะ: " +
          esc_(a.status || "-") +
          (a.next ? " | นัด " + esc_(a.next) : ""),
      );
    }
  });
  let msg = "";
  ALERT_LEVELS.forEach(function (lv) {
    if (groups[lv].length)
      msg +=
        "\n" +
        ALERT_ICON[lv] +
        " <b>" +
        lv +
        "</b>\n" +
        groups[lv].join("\n") +
        "\n";
  });
  return msg
    ? "⚠️ <b>รายการแอร์ที่ต้องเข้าดูแล</b>\n" + msg
    : "✅ <b>รายงานประจำวัน:</b> แอร์ทุกตัวปกติ ไม่มีนัดค้างครับ";
}

function checkAndSendAirNotification() {
  safe_("daily", function () {
    sendTelegramMessage(
      buildAlertMessage_() + pendingNote_() + techNote_(),
      undefined,
      dueKeyboard_(),
    );
  });
}

function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "checkAndSendAirNotification")
      ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("checkAndSendAirNotification")
    .timeBased()
    .atHour(8)
    .nearMinute(0)
    .everyDays(1)
    .inTimezone(TZ)
    .create();
}

// ==================== Webhook (รับคำสั่งจาก Telegram) ====================
function setWebhook() {
  const base = prop_("WEBAPP_URL") || ScriptApp.getService().getUrl();
  if (!/\/exec$/.test(base))
    throw new Error(
      "URL ต้องลงท้าย /exec — ให้ตั้ง Script property WEBAPP_URL เป็นลิงก์ Web app จากหน้า Deploy (ตอนนี้ได้: " +
        base +
        ")",
    );
  const secret = prop_("WEBHOOK_SECRET");
  if (!secret) throw new Error("ตั้ง WEBHOOK_SECRET ใน Script properties ก่อน");
  const me = JSON.parse(
    UrlFetchApp.fetch(
      "https://api.telegram.org/bot" + prop_("TELEGRAM_TOKEN") + "/getMe",
      { muteHttpExceptions: true },
    ).getContentText(),
  );
  if (!me.ok)
    throw new Error(
      "TELEGRAM_TOKEN ใช้ไม่ได้ (" +
        me.description +
        ") — เอาโทเคนปัจจุบันจาก BotFather มาใส่ใน Script properties",
    );
  const url = base + "?s=" + encodeURIComponent(secret);
  const res = UrlFetchApp.fetch(
    "https://api.telegram.org/bot" +
      prop_("TELEGRAM_TOKEN") +
      "/setWebhook?url=" +
      encodeURIComponent(url) +
      "&drop_pending_updates=true&max_connections=1&allowed_updates=" +
      encodeURIComponent('["message","callback_query"]'),
  );
  Logger.log(base + "\n" + res.getContentText());
}

function ok_() {
  return ContentService.createTextOutput("ok");
}

function doPost(e) {
  let chatId = null;
  try {
    const secret = prop_("WEBHOOK_SECRET");
    if (!secret || !e.parameter || e.parameter.s !== secret) return ok_();
    const u = JSON.parse(e.postData.contents);
    if (seen_(u.update_id)) return ok_();

    const cb = u.callback_query;
    const msg = cb ? cb.message : u.message;
    if (!msg) return ok_();
    chatId = String(msg.chat.id);

    const from = cb ? cb.from : msg.from;
    ACTOR_ = from
      ? from.username
        ? "@" + from.username
        : [from.first_name, from.last_name].filter(String).join(" ")
      : "";

    if (!isAllowed_(chatId)) {
      audit_(ACTOR_, chatId, "[ถูกบล็อก] " + (cb ? cb.data : msg.text || ""));
      // แชทส่วนตัวที่ยังไม่ได้อนุญาต: บอก chat id เพื่อให้ผู้ดูแลนำไปใส่ ALLOWED_CHAT_IDS
      if (Number(chatId) > 0 && !cb)
        sendTelegramMessage(
          "⛔ แชทนี้ยังไม่ได้รับอนุญาตให้ใช้บอท\nChat ID: <code>" +
            esc_(chatId) +
            "</code>\nแจ้งผู้ดูแลเพื่อเพิ่มใน ALLOWED_CHAT_IDS",
          chatId,
          { remove_keyboard: true },
        );
      return ok_();
    }

    CHAT_ = chatId;
    if (cb) {
      handleCallback_(cb, chatId);
      return ok_();
    }
    if (!msg.text) return ok_();
    audit_(ACTOR_, chatId, msg.text);
    reply_(chatId, handleText_(msg.text.trim(), chatId, msg));
  } catch (err) {
    if (
      /Telegram|Gemini|Exception|Service|quota|limit|timed out/i.test(
        err.message,
      )
    )
      logError_("doPost", err);
    if (chatId) {
      try {
        sendTelegramMessage("❌ " + esc_(err.message), chatId);
      } catch (e2) {}
    }
  }
  return ok_();
}

const HELP_ =
  "🤖 <b>ผู้ช่วยดูแลแอร์</b>\n\n" +
  "/status — สรุปภาพรวม\n" +
  "/list — รายการแอร์ทั้งหมด\n" +
  "/due — รายการที่ต้องดูแล (มีปุ่มกดล้างแล้ว)\n" +
  "/ac [AC-01] — รายละเอียด + ประวัติ (ไม่ใส่รหัส = เลือกจากปุ่ม)\n" +
  "/wash [AC-01] [ค่าใช้จ่าย] [ช่าง] — บันทึกล้างแอร์ (นัดถัดไป +" +
  WASH_INTERVAL_MONTHS +
  " เดือน) ไม่ใส่รหัส = เลือกจากปุ่ม\n" +
  "/issue AC-01 อาการ — แจ้งปัญหา (สถานะ=รอซ่อม)\n" +
  "/fix AC-01 [ค่าใช้จ่าย] รายละเอียด — ซ่อมเสร็จ (สถานะ=ปกติ)\n" +
  "/set AC-01 เสีย — เปลี่ยนสถานะ (" +
  STATUSES.join("/") +
  ")\n" +
  "/undo — ย้อนรายการบันทึกล่าสุด\n" +
  "/history [AC-01] — ประวัติ 10 รายการล่าสุด\n" +
  "/cost [2026-10] — ค่าใช้จ่ายรายเดือน\n" +
  "/ask คำถาม — ถาม AI (Groq/Gemini)\n" +
  "   • แชทส่วนตัว: พิมพ์ข้อความธรรมดาได้เลย\n" +
  "   • ในกลุ่ม: ตอบกลับข้อความของบอท หรือพิมพ์ @ชื่อบอท ตามด้วยคำถาม\n" +
  "/dashboard — ลิงก์ Dashboard\n" +
  "/health — ตรวจสุขภาพระบบ";

// ปุ่มคีย์บอร์ดส่งข้อความมาเป็น "📊 /status" — ตัดอีโมจิ/สัญลักษณ์หน้าคำสั่งออกให้เป็น "/status"
function normalizeCmd_(text) {
  return String(text || "")
    .trim()
    .replace(/^[^\/\p{L}\p{N}]{1,8}(?=\/\w)/u, "");
}

function addressedToBot_(msg, text) {
  if (!msg) return false;
  let me = "";
  try {
    me = botUsername_().toLowerCase();
  } catch (e) {}
  const rf = msg.reply_to_message;
  if (
    rf &&
    rf.from &&
    rf.from.is_bot &&
    (!me || String(rf.from.username || "").toLowerCase() === me)
  )
    return true;
  return (
    !!me &&
    String(text)
      .toLowerCase()
      .indexOf("@" + me) !== -1
  );
}

function stripMention_(text) {
  let me = "";
  try {
    me = botUsername_();
  } catch (e) {}
  return me
    ? String(text)
        .replace(new RegExp("@" + me, "ig"), " ")
        .trim()
    : String(text).trim();
}

// ปุ่มเลือกเครื่อง (3 ปุ่มต่อแถว) prefix: w=ล้างแล้ว, a=ดูรายละเอียด
function pickKeyboard_(prefix, icon, onlyDue) {
  let list = getAirs_();
  if (onlyDue) {
    const d = list.filter(function (a) {
      return a.alert === "เลยกำหนด" || a.alert === "ใกล้ถึงกำหนด";
    });
    if (d.length) list = d;
  }
  list = list.slice(0, 30);
  const rows = [];
  list.forEach(function (a, i) {
    if (i % 3 === 0) rows.push([]);
    rows[rows.length - 1].push(btn_(icon + " " + a.id, prefix + ":" + a.id));
  });
  return rows.length ? { inline_keyboard: rows } : undefined;
}

function handleText_(raw, chatId, msg) {
  const text = normalizeCmd_(raw);
  const m = text.match(/^\/(\w+)(?:@\w+)?\s*([\s\S]*)$/);
  if (!m) {
    // ข้อความธรรมดา → ส่งให้ Gemini
    // แชทส่วนตัว: ตอบเสมอ | กลุ่ม: ตอบเมื่อ reply บอท หรือ @mention เท่านั้น
    if (Number(chatId) < 0 && !addressedToBot_(msg, text)) return null;
    const q = stripMention_(text);
    if (!q) return "พิมพ์คำถามต่อท้าย เช่น แอร์เย็นน้อยเกิดจากอะไรได้บ้าง";
    aiGuard_(chatId);
    typing_(chatId);
    return askLLM_(q.substring(0, 500));
  }
  const cmd = m[1].toLowerCase();
  const args = m[2].trim().split(/\s+/).filter(String);
  const rest = m[2].trim();
  const undoable = function (s) {
    return { text: s, markup: undoKeyboard_() };
  };

  const extra = extraCmd_(cmd, args, rest);
  if (extra !== null) return extra;

  switch (cmd) {
    case "start":
    case "help":
      return HELP_ + HELP_EXTRA_;
    case "menu":
      return cmdMenu_();
    case "status":
      return cmdStatus_();
    case "list":
      return cmdList_();
    case "due":
      return {
        text: buildAlertMessage_() + pendingNote_() + techNote_(),
        markup: dueKeyboard_(),
      };
    case "ac": {
      if (!args.length) {
        const kb = pickKeyboard_("a", "❄️", false);
        return kb
          ? { text: "❄️ เลือกเครื่องที่ต้องการดู", markup: kb }
          : "ยังไม่มีเครื่องแอร์ในชีต";
      }
      const id = normId_(args[0]);
      return { text: cmdAc_(id), markup: acKeyboard_(id) };
    }
    case "wash": {
      if (!args.length) {
        const kb = pickKeyboard_("w", "🧼", true);
        return kb
          ? {
              text: "🧼 เลือกเครื่องที่ล้างแล้ว (กดแล้วบันทึกทันที)",
              markup: kb,
            }
          : "ยังไม่มีเครื่องแอร์ในชีต";
      }
      return undoable(cmdWash_(args));
    }
    case "issue":
      return undoable(cmdIssue_(args));
    case "fix":
      return undoable(cmdFix_(args));
    case "set":
      return cmdSet_(args);
    case "undo":
      return cmdUndo_(chatId);
    case "history":
      return cmdHistory_(args[0] ? normId_(args[0]) : "");
    case "cost":
      return cmdCost_(args[0]);
    case "health":
      return healthCheck(true);
    case "ask":
      if (!rest)
        return "พิมพ์คำถามต่อท้าย เช่น /ask แอร์เย็นน้อยเกิดจากอะไรได้บ้าง";
      aiGuard_(chatId);
      typing_(chatId);
      return askLLM_(rest.substring(0, 500));
    case "dashboard":
      return (
        "📊 " +
        (prop_("WEBAPP_URL") || ScriptApp.getService().getUrl()) +
        "?key=" +
        prop_("DASH_KEY")
      );
    default:
      return "ไม่รู้จักคำสั่งนี้ พิมพ์ /help ดูรายการคำสั่ง";
  }
}

function requireAir_(id) {
  if (!id) throw new Error("ระบุรหัสเครื่อง เช่น AC-01");
  const a = getAirs_().filter(function (x) {
    return x.id === id;
  })[0];
  if (!a) throw new Error("ไม่พบเครื่อง " + id);
  return a;
}

function cmdStatus_() {
  const airs = getAirs_();
  const byStatus = {};
  const byAlert = {};
  let cost = 0;
  airs.forEach(function (a) {
    byStatus[a.status || "ไม่ระบุ"] =
      (byStatus[a.status || "ไม่ระบุ"] || 0) + 1;
    byAlert[a.alert] = (byAlert[a.alert] || 0) + 1;
    cost += a.cost;
  });
  let t = "📊 <b>สรุปแอร์ " + airs.length + " เครื่อง</b>\n\n<b>สถานะ</b>\n";
  Object.keys(byStatus).forEach(function (k) {
    t += "• " + esc_(k) + ": " + byStatus[k] + "\n";
  });
  t += "\n<b>แจ้งเตือน</b>\n";
  ALERT_LEVELS.forEach(function (lv) {
    if (byAlert[lv]) t += ALERT_ICON[lv] + " " + lv + ": " + byAlert[lv] + "\n";
  });
  return t + "\n💰 ค่าใช้จ่ายรวม " + baht_(cost) + " บาท";
}

function cmdAc_(id) {
  const a = requireAir_(id);
  let t =
    "❄️ <b>" +
    esc_(a.id) +
    "</b> — " +
    esc_(a.loc) +
    "\n" +
    esc_(a.model) +
    (a.btu ? " | " + baht_(a.btu) + " BTU" : "") +
    "\n" +
    "สถานะ: <b>" +
    esc_(a.status || "-") +
    "</b> | แจ้งเตือน: " +
    esc_(a.alert || "-") +
    "\n" +
    "ล้างล่าสุด: " +
    esc_(a.lastWash || "-") +
    " | นัดถัดไป: " +
    esc_(a.next || "-") +
    "\n" +
    "ซ่อม/บำรุง " +
    a.count +
    " ครั้ง | รวม " +
    baht_(a.cost) +
    " บาท\n";
  const logs = getLogs_()
    .filter(function (l) {
      return l.id === id;
    })
    .slice(0, 5);
  if (logs.length) {
    t += "\n<b>ประวัติล่าสุด</b>\n";
    logs.forEach(function (l) {
      t +=
        "• " +
        fmt_(l.date) +
        " " +
        esc_(l.type) +
        (l.detail ? " — " + esc_(l.detail) : "") +
        (l.cost ? " (" + baht_(l.cost) + "฿)" : "") +
        "\n";
    });
  }
  return t;
}

function takeCost_(args) {
  if (args.length && /^\d+(\.\d+)?$/.test(args[0])) return Number(args.shift());
  return "";
}

function cmdWash_(args) {
  const id = normId_(args.shift());
  requireAir_(id);
  const cost = takeCost_(args);
  const today = new Date();
  const next = addMonths_(today, WASH_INTERVAL_MONTHS);
  addLog_({
    date: today,
    id: id,
    type: "ล้างแอร์",
    detail: "",
    tech: args.join(" "),
    cost: cost,
    next: next,
  });
  setStatus_(id, "ปกติ");
  return (
    "🧼 บันทึกล้างแอร์ <b>" +
    esc_(id) +
    "</b> แล้ว\nนัดครั้งถัดไป: " +
    fmt_(next) +
    (cost ? "\nค่าใช้จ่าย " + baht_(cost) + " บาท" : "")
  );
}

function cmdIssue_(args) {
  const id = normId_(args.shift());
  requireAir_(id);
  if (!args.length)
    throw new Error("ระบุอาการด้วย เช่น /issue AC-01 เย็นน้อย มีน้ำหยด");
  addLog_({
    date: new Date(),
    id: id,
    type: "แจ้งปัญหา",
    detail: args.join(" "),
  });
  setStatus_(id, "รอซ่อม");
  return (
    "🔧 บันทึกปัญหา <b>" +
    esc_(id) +
    "</b> แล้ว สถานะ → รอซ่อม\n" +
    esc_(args.join(" "))
  );
}

function cmdFix_(args) {
  const id = normId_(args.shift());
  requireAir_(id);
  const cost = takeCost_(args);
  addLog_({
    date: new Date(),
    id: id,
    type: "ซ่อม",
    detail: args.join(" "),
    cost: cost,
  });
  setStatus_(id, "ปกติ");
  return (
    "✅ บันทึกซ่อม <b>" +
    esc_(id) +
    "</b> เสร็จแล้ว สถานะ → ปกติ" +
    (cost ? "\nค่าใช้จ่าย " + baht_(cost) + " บาท" : "")
  );
}

function cmdSet_(args) {
  const id = normId_(args.shift());
  const st = args.join(" ");
  requireAir_(id);
  if (STATUSES.indexOf(st) === -1)
    throw new Error("สถานะต้องเป็น: " + STATUSES.join(" / "));
  setStatus_(id, st);
  return "✏️ <b>" + esc_(id) + "</b> สถานะ → " + esc_(st);
}

function buildContext_(maxLogs) {
  const airs = getAirs_();
  const logs = getLogs_().slice(0, maxLogs || 40);
  let t =
    "วันนี้: " +
    fmt_(new Date()) +
    "\n\nเครื่องแอร์ (รหัส | ตำแหน่ง | รุ่น | BTU | สถานะ | ล้างล่าสุด | นัดถัดไป | แจ้งเตือน | ซ่อมกี่ครั้ง | ค่าใช้จ่ายรวม)\n";
  airs.forEach(function (a) {
    t +=
      [
        a.id,
        a.loc,
        a.model,
        a.btu,
        a.status || "-",
        a.lastWash || "-",
        a.next || "-",
        a.alert || "-",
        a.count,
        a.cost,
      ].join(" | ") + "\n";
  });
  t +=
    "\nประวัติล่าสุด (วันที่ | รหัส | ประเภท | รายละเอียด | ช่าง | ค่าใช้จ่าย)\n";
  logs.forEach(function (l) {
    t +=
      [
        fmt_(l.date),
        l.id,
        l.type,
        l.detail || "-",
        l.tech || "-",
        l.cost || 0,
      ].join(" | ") + "\n";
  });
  return t;
}

// ==================== AI: Groq + Gemini (ฟรีทั้งคู่ สลับสำรองให้อัตโนมัติ) ====================
// Script properties:
//   GROQ_API_KEY   = คีย์จาก console.groq.com (ฟรี ไม่ต้องใช้บัตร)
//   GEMINI_API_KEY = คีย์จาก aistudio.google.com
//   AI_PROVIDER    = (ไม่บังคับ) "groq" หรือ "gemini" = ตัวหลัก (ค่าเริ่มต้น: groq ถ้ามี GROQ_API_KEY ไม่งั้น gemini)
// ตัวหลักล้ม/โควต้าเต็ม → ลองตัวสำรองให้เอง

const AI_SYSTEM_ =
  "คุณเป็นผู้ช่วยดูแลรักษาเครื่องปรับอากาศของทีมซ่อมบำรุง ตอบเป็นภาษาไทย กระชับ ใช้ข้อมูลจากชีตที่ให้เป็นหลักและอ้างรหัสเครื่องเมื่อเกี่ยวข้อง " +
  "ห้ามเดาข้อมูลที่ไม่มีในชีต ถ้าถามเรื่องอาการ/สาเหตุ ให้ไล่จากสิ่งที่ตรวจเองได้ง่ายไปยาก (ฟิลเตอร์ เดรน คอยล์ น้ำยา คอมเพรสเซอร์) " +
  "และแนะนำให้เรียกช่างเมื่อเกี่ยวกับไฟฟ้าแรงสูง น้ำยา หรือคอมเพรสเซอร์ ตอบเป็นข้อความธรรมดา ไม่ใช้ markdown";

// ---------- Groq ----------
// ตัวแรกคุณภาพดีสุด โควต้าฟรี/วันของแต่ละตัวแยกกัน เต็มตัวไหนจะข้ามไปตัวถัดไป
const GROQ_MODELS = [
  "openai/gpt-oss-120b",
  "llama-3.3-70b-versatile",
  "openai/gpt-oss-20b",
  "llama-3.1-8b-instant",
];

function groqCall_(model, key, system, userText, noReasoning) {
  const body = {
    model: model,
    messages: [
      { role: "system", content: system },
      { role: "user", content: userText },
    ],
    temperature: 0.4,
    max_completion_tokens: 2048,
  };
  if (!noReasoning && /^openai\/gpt-oss/.test(model))
    body.reasoning_effort = "low";
  const res = UrlFetchApp.fetch(
    "https://api.groq.com/openai/v1/chat/completions",
    {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: "Bearer " + key },
      muteHttpExceptions: true,
      payload: JSON.stringify(body),
    },
  );
  let j = {};
  try {
    j = JSON.parse(res.getContentText());
  } catch (e) {}
  return { code: res.getResponseCode(), j: j };
}

function askGroq_(question) {
  const key = prop_("GROQ_API_KEY");
  if (!key)
    throw new Error("ยังไม่ได้ตั้งค่า GROQ_API_KEY ใน Script properties");
  // โควต้าฟรีของ Groq จำกัดโทเคน/นาที → ส่งประวัติให้น้อยลง
  const userText =
    "ข้อมูลปัจจุบัน:\n" + buildContext_(25) + "\nคำถาม: " + question;

  const cache = CacheService.getScriptCache();
  const lastOk = cache.get("groq_ok");
  const order = [];
  [lastOk, prop_("GROQ_MODEL")].concat(GROQ_MODELS).forEach(function (m) {
    if (m && order.indexOf(m) === -1) order.push(m);
  });

  let lastErr = "";
  for (let i = 0; i < order.length; i++) {
    let r = groqCall_(order[i], key, AI_SYSTEM_, userText, false);
    if (r.code === 400 && /^openai\/gpt-oss/.test(order[i]))
      r = groqCall_(order[i], key, AI_SYSTEM_, userText, true);

    if (r.code === 200) {
      const ch = r.j.choices && r.j.choices[0];
      const text = String((ch && ch.message && ch.message.content) || "")
        .replace(/\*\*/g, "")
        .trim();
      if (text) {
        cache.put("groq_ok", order[i], 21600);
        return esc_(text);
      }
      lastErr = "ไม่มีคำตอบ (" + ((ch && ch.finish_reason) || "?") + ")";
      continue;
    }
    lastErr = (r.j.error && r.j.error.message) || String(r.code);
    if (order[i] === lastOk) cache.remove("groq_ok");
    // 404 ไม่มีโมเดล | 413 คำขอใหญ่เกินโควต้า/นาที | 429 โควต้าเต็ม | 5xx ล่ม → ลองตัวถัดไป
    if ([404, 413, 429, 500, 502, 503].indexOf(r.code) === -1) break;
  }
  throw new Error("Groq: " + lastErr);
}

// ---------- Gemini ----------
// (ตระกูล 2.5 ตั้งปลดระวาง 16 ต.ค. 2026 — เก็บไว้ท้ายสุด; 3.7/3.5 ถูกยุบไปชี้ 3.8/3.6 แล้ว)
const GEMINI_FALLBACKS = [
  "gemini-3.1-flash-lite",
  "gemini-3.8-flash",
  "gemini-3.6-flash",
  "gemini-3-flash-preview",
  "gemini-2.5-flash",
  "gemini-2.5-flash-lite",
];

function geminiThinking_(model) {
  if (/^gemini-3/.test(model)) return { thinkingLevel: "low" };
  if (/^gemini-2\.5-flash/.test(model)) return { thinkingBudget: 0 };
  return null; // รุ่นอื่นไม่ส่ง thinkingConfig กัน 400
}

function geminiCall_(model, key, system, userText, skipThinking) {
  const gc = { maxOutputTokens: 2048 };
  const th = skipThinking ? null : geminiThinking_(model);
  if (th) gc.thinkingConfig = th;
  const res = UrlFetchApp.fetch(
    "https://generativelanguage.googleapis.com/v1beta/models/" +
      encodeURIComponent(model) +
      ":generateContent",
    {
      method: "post",
      contentType: "application/json",
      headers: { "x-goog-api-key": key },
      muteHttpExceptions: true,
      payload: JSON.stringify({
        systemInstruction: { parts: [{ text: system }] },
        generationConfig: gc,
        contents: [{ role: "user", parts: [{ text: userText }] }],
      }),
    },
  );
  let j = {};
  try {
    j = JSON.parse(res.getContentText());
  } catch (e) {}
  return { code: res.getResponseCode(), j: j };
}

function askGemini_(question) {
  const key = prop_("GEMINI_API_KEY");
  if (!key)
    throw new Error("ยังไม่ได้ตั้งค่า GEMINI_API_KEY ใน Script properties");
  const userText =
    "ข้อมูลปัจจุบัน:\n" + buildContext_(40) + "\nคำถาม: " + question;

  const cache = CacheService.getScriptCache();
  const lastOk = cache.get("gemini_ok");
  const order = [];
  [lastOk, prop_("GEMINI_MODEL"), GEMINI_MODEL]
    .concat(GEMINI_FALLBACKS)
    .forEach(function (m) {
      if (m && order.indexOf(m) === -1) order.push(m);
    });

  let lastErr = "";
  for (let i = 0; i < order.length; i++) {
    let r = geminiCall_(order[i], key, AI_SYSTEM_, userText, false);
    if (r.code === 400 && geminiThinking_(order[i]))
      r = geminiCall_(order[i], key, AI_SYSTEM_, userText, true);

    if (r.code === 200) {
      const cand = r.j.candidates && r.j.candidates[0];
      const parts = (cand && cand.content && cand.content.parts) || [];
      const text = parts
        .filter(function (p) {
          return !p.thought;
        })
        .map(function (p) {
          return p.text || "";
        })
        .join("")
        .replace(/\*\*/g, "")
        .trim();
      if (text) {
        cache.put("gemini_ok", order[i], 21600);
        return esc_(text);
      }
      const why =
        (r.j.promptFeedback && r.j.promptFeedback.blockReason) ||
        (cand && cand.finishReason) ||
        "";
      lastErr = "ไม่มีคำตอบ" + (why ? " (" + why + ")" : "");
      if (order[i] === lastOk) cache.remove("gemini_ok");
      continue;
    }
    lastErr = (r.j.error && r.j.error.message) || String(r.code);
    if (order[i] === lastOk) cache.remove("gemini_ok");
    if ([404, 429, 500, 503].indexOf(r.code) === -1) break;
  }
  throw new Error("Gemini: " + lastErr);
}

// ---------- ตัวกลาง: เลือกผู้ให้บริการ + สำรอง ----------
function askLLM_(question) {
  const hasGroq = !!prop_("GROQ_API_KEY");
  const hasGem = !!prop_("GEMINI_API_KEY");
  if (!hasGroq && !hasGem)
    throw new Error(
      "ยังไม่ได้ตั้งค่า GROQ_API_KEY หรือ GEMINI_API_KEY ใน Script properties",
    );
  let first = String(prop_("AI_PROVIDER") || "").toLowerCase();
  if (first !== "groq" && first !== "gemini")
    first = hasGroq ? "groq" : "gemini";
  const order = first === "groq" ? ["groq", "gemini"] : ["gemini", "groq"];
  let lastErr = null;
  for (let i = 0; i < order.length; i++) {
    const p = order[i];
    if ((p === "groq" && !hasGroq) || (p === "gemini" && !hasGem)) continue;
    try {
      const a = p === "groq" ? askGroq_(question) : askGemini_(question);
      CacheService.getScriptCache().put("ai_provider_ok", p, 21600);
      return a;
    } catch (e) {
      lastErr = e;
    }
  }
  throw lastErr;
}

function askAI(key, question) {
  requireKey_(key);
  aiGuard_("dash");
  return askLLM_(String(question).substring(0, 500))
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

// ==================== Dashboard ====================
function doGet(e) {
  const t = HtmlService.createTemplateFromFile("Dashboard");
  t.key = prop_("DASH_KEY") || "free";
  return t
    .evaluate()
    .setTitle("Dashboard ดูแลแอร์")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// ฟังก์ชันดึงอุณหภูมิกรุงเทพฯ แบบเรียลไทม์จาก Open-Meteo API ฟรี (ไม่ต้องใช้ API Key)
function getBangkokWeather_() {
  try {
    const url =
      "https://api.open-meteo.com/v1/forecast?latitude=13.7563&longitude=100.5018&current=temperature_2m,relative_humidity_2m,weather_code&timezone=Asia%2FBangkok";
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const data = JSON.parse(res.getContentText());
    if (data && data.current) {
      return {
        temp: data.current.temperature_2m,
        humidity: data.current.relative_humidity_2m,
      };
    }
  } catch (e) {}
  return { temp: null, humidity: null };
}

// ปรับปรุงฟังก์ชัน getDashboardData ให้แนบข้อมูล weather ไปด้วย
function getDashboardData(key) {
  requireKey_(key);
  const airs = getAirs_();
  const logs = getLogs_();
  const byStatus = {};
  const byAlert = {};
  let cost = 0;
  airs.forEach(function (a) {
    byStatus[a.status || "ไม่ระบุ"] =
      (byStatus[a.status || "ไม่ระบุ"] || 0) + 1;
    if (a.alert && a.alert !== "-")
      byAlert[a.alert] = (byAlert[a.alert] || 0) + 1;
    cost += a.cost;
  });

  const monthly = [];
  const now = new Date();
  for (let i = 5; i >= 0; i--) {
    const d = addMonths_(new Date(now.getFullYear(), now.getMonth(), 1), -i);
    monthly.push({
      key: Utilities.formatDate(d, TZ, "yyyy-MM"),
      label: Utilities.formatDate(d, TZ, "MM/yy"),
      cost: 0,
    });
  }
  logs.forEach(function (l) {
    const k = Utilities.formatDate(l.date, TZ, "yyyy-MM");
    monthly.forEach(function (m) {
      if (m.key === k) m.cost += l.cost;
    });
  });

  return {
    updated: Utilities.formatDate(new Date(), TZ, "dd/MM/yyyy HH:mm"),
    weather: getBangkokWeather_(), // <--- ดึงอุณหภูมิกรุงเทพฯ ใส่ตรงนี้
    total: airs.length,
    cost: cost,
    byStatus: byStatus,
    byAlert: byAlert,
    airs: airs,
    monthly: monthly,
    pending: pendingList_(),
    recent: logs.slice(0, 12).map(function (l) {
      return {
        date: fmt_(l.date),
        id: l.id,
        type: l.type,
        detail: l.detail,
        tech: l.tech,
        cost: l.cost,
      };
    }),
  };
}

// ฟังก์ชันบันทึกงานด่วนจากหน้า Modal ของ Dashboard
function apiQuickLog(key, o) {
  return api_("quickLog", key, true, function () {
    const id = normId_(clean_(o.id, 20));
    const air = getAirs_().filter(function (x) {
      return x.id === id;
    })[0];
    if (!air) throw userErr_("ไม่พบเครื่อง " + id);

    const type = clean_(o.type, 30) || "ล้างแอร์";
    const detail = clean_(o.detail, 200);
    const cost = cleanNum_(o.cost, 0, 1000000, "ค่าใช้จ่าย", false) || 0;
    const today = new Date();
    let next = null;

    // ถ้าเป็นการล้างแอร์ ให้คำนวณนัดถัดไป +6 เดือนอัตโนมัติ
    if (type === "ล้างแอร์") {
      next = addMonths_(today, WASH_INTERVAL_MONTHS);
      setStatus_(id, "ปกติ");
    } else if (type === "ซ่อมแอร์" || type === "ตรวจเช็ค") {
      setStatus_(id, "ปกติ");
    }

    addLog_({
      date: today,
      id: id,
      type: type,
      detail: detail,
      tech: "Dashboard (ด่วน)",
      cost: cost,
      next: next,
      note: "บันทึกด่วนจากหน้า Modal",
    });

    audit_("dashboard", "web", "[ด่วน] " + type + " " + id);
    return getDashboardData(key);
  });
}
