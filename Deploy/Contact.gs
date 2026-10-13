// ==================== ช่างล้างแอร์ + สถิติค่าไฟ + ประวัติ (v2.1) ====================
const TECH_SHEET = "ช่างล้างแอร์";
const POWER_SHEET = "ค่าไฟ";
const TECH_HEADER = [
  "รหัส",
  "ชื่อ/บริษัท",
  "เบอร์โทร",
  "LINE ID",
  "ความชำนาญ/พื้นที่",
  "ราคาอ้างอิง (บาท/เครื่อง)",
  "หมายเหตุ",
  "แนะนำ",
  "สถานะ",
  "เพิ่มเมื่อ",
  "เพิ่มโดย",
];
const TECH = {
  ID: 0,
  NAME: 1,
  PHONE: 2,
  LINE: 3,
  AREA: 4,
  PRICE: 5,
  NOTE: 6,
  FAV: 7,
  STATE: 8,
  CREATED: 9,
  BY: 10,
};
const TECH_FMT = [
  "@",
  "@",
  "@",
  "@",
  "@",
  "#,##0",
  "@",
  "@",
  "@",
  "dd/mm/yyyy hh:mm",
  "@",
];
const POWER_HEADER = [
  "เดือน (yyyy-MM)",
  "ค่าไฟ (บาท)",
  "หน่วยไฟ (kWh)",
  "หมายเหตุ",
  "อัปเดตล่าสุด",
  "โดย",
];
const PW = { MONTH: 0, BAHT: 1, KWH: 2, NOTE: 3, UPDATED: 4, BY: 5 };
const PW_FMT = ["@", "#,##0.00", "#,##0.##", "@", "dd/mm/yyyy hh:mm", "@"];
const CFG_DEFAULT = { rate: 4.5, hours: 8, days: 26, eer: 11 };
const WRITE_LIMIT_PER_MIN = 30;

const HELP_EXTRA_ =
  "\n\n<b>ช่างและค่าไฟ</b>\n" +
  "/tech — รายชื่อช่าง + เบอร์โทร\n" +
  "/addtech ชื่อช่าง เบอร์โทร [LINE ID] — เพิ่มช่าง\n" +
  "/bill [2026-09] ค่าไฟ(บาท) [หน่วย] — บันทึกบิลค่าไฟ\n" +
  "/power — สถิติค่าไฟ";

function requireKey_(key) {
  const k = prop_("DASH_KEY");
  if (!k || typeof key !== "string" || key.length !== k.length)
    throw new Error("ไม่มีสิทธิ์เข้าถึง (key ไม่ถูกต้อง)");
  let d = 0;
  for (let i = 0; i < k.length; i++) d |= key.charCodeAt(i) ^ k.charCodeAt(i);
  if (d) throw new Error("ไม่มีสิทธิ์เข้าถึง (key ไม่ถูกต้อง)");
}

function userErr_(msg) {
  const e = new Error(msg);
  e.user = true;
  return e;
}

function writeGuard_() {
  const c = CacheService.getScriptCache();
  const n = Number(c.get("wr_dash") || 0);
  if (n >= WRITE_LIMIT_PER_MIN)
    throw userErr_("บันทึกถี่เกินไป รอสักครู่แล้วลองใหม่");
  c.put("wr_dash", String(n + 1), 60);
}

function api_(name, key, write, fn) {
  requireKey_(key);
  try {
    if (write) writeGuard_();
    return fn();
  } catch (e) {
    if (e && e.user) throw e;
    logError_(name, e);
    throw new Error("ระบบขัดข้อง ลองใหม่อีกครั้ง (ถ้ายังไม่หาย แจ้งผู้ดูแล)");
  }
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw userErr_("ระบบกำลังยุ่ง ลองใหม่อีกครั้ง");
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function clean_(v, max) {
  return String(v == null ? "" : v)
    .replace(/[\u0000-\u001f\u007f]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .substring(0, max);
}
function num_(v) {
  if (v === "" || v === null || v === undefined) return null;
  const n = Number(String(v).replace(/,/g, ""));
  return isFinite(n) ? n : null;
}
function cleanNum_(v, min, max, label, required) {
  if (v === "" || v === null || v === undefined) {
    if (required) throw userErr_("กรุณาใส่" + label);
    return null;
  }
  const n = Number(String(v).replace(/,/g, ""));
  if (!isFinite(n) || n < min || n > max)
    throw userErr_(label + " ต้องเป็นตัวเลขระหว่าง " + min + " – " + max);
  return n;
}
function cleanPhone_(v) {
  const s = clean_(v, 30).replace(/[\s\-().]/g, "");
  if (!/^\+?\d{8,15}$/.test(s))
    throw userErr_("เบอร์โทรไม่ถูกต้อง (ใส่ตัวเลข 8–15 หลัก)");
  return s;
}
function cleanLine_(v) {
  const s = clean_(v, 60).replace(/^https?:\/\/line\.me\/R\/ti\/p\/~?/i, "");
  if (!s) return "";
  if (!/^@?[A-Za-z0-9._\-]{2,38}$/.test(s))
    throw userErr_("LINE ID ใช้ได้เฉพาะ a-z 0-9 . _ - (ขึ้นต้นด้วย @ ได้)");
  return s;
}
function phoneRaw_(v) {
  if (typeof v === "number") {
    const s = String(Math.round(v));
    return s.length === 9 ? "0" + s : s;
  }
  return String(v || "").trim();
}
function monthKey_(d) {
  return Utilities.formatDate(d, TZ, "yyyy-MM");
}
function dt_(v, withTime) {
  return v instanceof Date
    ? Utilities.formatDate(v, TZ, withTime ? "dd/MM/yyyy HH:mm" : "dd/MM/yyyy")
    : String(v || "");
}

function ensureSheet_(name, header, textRanges, numFmts) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.getRange(1, 1, 1, header.length)
      .setValues([header])
      .setFontWeight("bold")
      .setBackground("#e8eef7");
    s.setFrozenRows(1);
    textRanges.forEach(function (r) {
      s.getRange(r).setNumberFormat("@");
    });
    Object.keys(numFmts || {}).forEach(function (r) {
      s.getRange(r).setNumberFormat(numFmts[r]);
    });
  }
  return s;
}
function techSheet_() {
  return ensureSheet_(TECH_SHEET, TECH_HEADER, ["A2:E", "G2:I", "K2:K"], {
    "F2:F": "#,##0",
    "J2:J": "dd/mm/yyyy hh:mm",
  });
}
function powerSheet_() {
  return ensureSheet_(POWER_SHEET, POWER_HEADER, ["A2:A", "D2:D", "F2:F"], {
    "B2:C": "#,##0.00",
    "E2:E": "dd/mm/yyyy hh:mm",
  });
}

function readTechRows_(sh) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh
    .getRange(2, 1, n, TECH_HEADER.length)
    .getValues()
    .map(function (r, i) {
      return {
        row: i + 2,
        id: String(r[TECH.ID] || "").trim(),
        name: String(r[TECH.NAME] || "").trim(),
        phone: phoneRaw_(r[TECH.PHONE]),
        line: String(r[TECH.LINE] || "").trim(),
        area: String(r[TECH.AREA] || "").trim(),
        price: num_(r[TECH.PRICE]),
        note: String(r[TECH.NOTE] || "").trim(),
        fav: String(r[TECH.FAV] || "").trim() !== "",
        state: String(r[TECH.STATE] || "").trim(),
        created: dt_(r[TECH.CREATED], false),
        by: String(r[TECH.BY] || "").trim(),
      };
    })
    .filter(function (t) {
      return t.name;
    });
}

function backfillTechIds_(sh) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const v = sh.getRange(2, 1, n, 2).getValues();
  let changed = false;
  v.forEach(function (r) {
    if (!String(r[0]).trim() && String(r[1]).trim()) {
      r[0] = "T" + Utilities.getUuid().replace(/-/g, "").substring(0, 8);
      changed = true;
    }
  });
  if (changed)
    sh.getRange(2, 1, n, 1)
      .setNumberFormat("@")
      .setValues(
        v.map(function (r) {
          return [r[0]];
        }),
      );
}

function sameTech_(logTech, name) {
  const norm = function (s) {
    return String(s || "")
      .toLowerCase()
      .replace(/\s+/g, " ")
      .trim();
  };
  const a = norm(logTech),
    b = norm(name);
  if (!a || !b) return false;
  if (a === b) return true;
  return (
    (b.length >= 3 && a.indexOf(b) !== -1) ||
    (a.length >= 3 && b.indexOf(a) !== -1)
  );
}

function techListData_() {
  const sh = techSheet_();
  const logs = getLogs_();
  const techs = readTechRows_(sh)
    .filter(function (t) {
      return t.state !== "ลบแล้ว";
    })
    .map(function (t) {
      let jobs = 0,
        total = 0,
        last = "";
      const recent = [];
      logs.forEach(function (l) {
        if (!sameTech_(l.tech, t.name)) return;
        jobs++;
        total += l.cost;
        if (!last) last = fmt_(l.date);
        if (recent.length < 5)
          recent.push({
            date: fmt_(l.date),
            id: l.id,
            type: l.type,
            detail: l.detail,
            cost: l.cost,
          });
      });
      return {
        id: t.id,
        name: t.name,
        phone: t.phone,
        line: t.line,
        area: t.area,
        price: t.price,
        note: t.note,
        fav: t.fav,
        created: t.created,
        by: t.by,
        jobs: jobs,
        total: total,
        last: last,
        recent: recent,
      };
    });
  techs.sort(function (a, b) {
    if (a.fav !== b.fav) return a.fav ? -1 : 1;
    return a.name < b.name ? -1 : a.name > b.name ? 1 : 0;
  });
  return { techs: techs };
}

function saveTech_(o, actor) {
  const name = clean_(o.name, 80);
  if (!name) throw userErr_("กรุณาใส่ชื่อช่าง/บริษัท");
  const phone = cleanPhone_(o.phone);
  const line = cleanLine_(o.line);
  const area = clean_(o.area, 120);
  const note = clean_(o.note, 300);
  const price = cleanNum_(o.price, 0, 1000000, "ราคา", false);
  const fav = o.fav ? "ใช่" : "";
  const id = clean_(o.id, 20);
  return withLock_(function () {
    const sh = techSheet_();
    const rows = readTechRows_(sh).filter(function (t) {
      return t.state !== "ลบแล้ว";
    });
    const digits = function (s) {
      return String(s).replace(/\D/g, "");
    };
    const dup = rows.filter(function (t) {
      return t.id !== id && digits(t.phone) === digits(phone);
    })[0];
    if (dup) throw userErr_('เบอร์นี้มีอยู่แล้วในชื่อ "' + dup.name + '"');
    if (id) {
      const ex = rows.filter(function (t) {
        return t.id === id;
      })[0];
      if (!ex) throw userErr_("ไม่พบช่างคนนี้ (อาจถูกลบไปแล้ว) กรุณารีเฟรช");
      sh.getRange(ex.row, 2, 1, 8)
        .setNumberFormats([TECH_FMT.slice(1, 9)])
        .setValues([
          [
            name,
            phone,
            line,
            area,
            price === null ? "" : price,
            note,
            fav,
            "ใช้งาน",
          ],
        ]);
      audit_(actor, "web", "[ช่าง] แก้ไข " + name);
    } else {
      const row = sh.getLastRow() + 1;
      sh.getRange(row, 1, 1, TECH_HEADER.length)
        .setNumberFormats([TECH_FMT])
        .setValues([
          [
            "T" + Utilities.getUuid().replace(/-/g, "").substring(0, 8),
            name,
            phone,
            line,
            area,
            price === null ? "" : price,
            note,
            fav,
            "ใช้งาน",
            new Date(),
            actor || "",
          ],
        ]);
      audit_(actor, "web", "[ช่าง] เพิ่ม " + name);
    }
    return null;
  });
}

function setTechField_(id, fn, auditText, actor) {
  return withLock_(function () {
    const sh = techSheet_();
    const t = readTechRows_(sh).filter(function (x) {
      return x.id === id && x.state !== "ลบแล้ว";
    })[0];
    if (!t) throw userErr_("ไม่พบช่างคนนี้ (อาจถูกลบไปแล้ว) กรุณารีเฟรช");
    fn(sh, t);
    audit_(actor, "web", auditText + " " + t.name);
  });
}

function apiTechList(key) {
  return api_("techList", key, false, function () {
    withLock_(function () {
      backfillTechIds_(techSheet_());
    });
    return techListData_();
  });
}
function apiTechSave(key, o) {
  return api_("techSave", key, true, function () {
    saveTech_(o || {}, "dashboard");
    return techListData_();
  });
}
function apiTechFav(key, id, on) {
  return api_("techFav", key, true, function () {
    setTechField_(
      id,
      function (sh, t) {
        sh.getRange(t.row, TECH.FAV + 1)
          .setNumberFormat("@")
          .setValue(on ? "ใช่" : "");
      },
      on ? "[ช่าง] ตั้งเป็นช่างประจำ" : "[ช่าง] เลิกเป็นช่างประจำ",
      "dashboard",
    );
    return techListData_();
  });
}
function apiTechDelete(key, id) {
  return api_("techDelete", key, true, function () {
    setTechField_(
      id,
      function (sh, t) {
        sh.getRange(t.row, TECH.STATE + 1)
          .setNumberFormat("@")
          .setValue("ลบแล้ว");
      },
      "[ช่าง] ลบ",
      "dashboard",
    );
    return techListData_();
  });
}

function getCfg_() {
  const p = PropertiesService.getScriptProperties();
  const c = {};
  Object.keys(CFG_DEFAULT).forEach(function (k) {
    const v = Number(p.getProperty("CFG_" + k));
    c[k] = v > 0 ? v : CFG_DEFAULT[k];
  });
  return c;
}

function readPowerRows_(sh) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh
    .getRange(2, 1, n, POWER_HEADER.length)
    .getValues()
    .map(function (r, i) {
      const m =
        r[PW.MONTH] instanceof Date
          ? monthKey_(r[PW.MONTH])
          : String(r[PW.MONTH] || "").trim();
      return {
        row: i + 2,
        month: m,
        baht: num_(r[PW.BAHT]),
        kwh: num_(r[PW.KWH]),
        note: String(r[PW.NOTE] || "").trim(),
        updated: dt_(r[PW.UPDATED], true),
        by: String(r[PW.BY] || "").trim(),
      };
    })
    .filter(function (b) {
      return /^\d{4}-(0[1-9]|1[0-2])$/.test(b.month) && b.baht !== null;
    });
}

function powerStats_(bills, cfg, airs, curKey) {
  const r1 = function (n) {
    return Math.round(n * 10) / 10;
  };
  const keyOf = function (y, m) {
    const d = new Date(Date.UTC(y, m - 1, 1));
    return d.getUTCFullYear() + "-" + ("0" + (d.getUTCMonth() + 1)).slice(-2);
  };
  const map = {};
  bills.forEach(function (b) {
    map[b.month] = b;
  });
  const months = Object.keys(map).sort();
  const cy = Number(curKey.slice(0, 4)),
    cm = Number(curKey.slice(5, 7));

  const series = [];
  for (let i = 11; i >= 0; i--) {
    const k = keyOf(cy, cm - i),
      b = map[k];
    series.push({
      key: k,
      label: k.slice(5) + "/" + k.slice(2, 4),
      baht: b ? b.baht : null,
      kwh: b && b.kwh !== null ? b.kwh : null,
    });
  }

  const momOf = function (k) {
    const b = map[k];
    const p = map[keyOf(Number(k.slice(0, 4)), Number(k.slice(5, 7)) - 1)];
    return b && p && p.baht > 0 ? r1(((b.baht - p.baht) / p.baht) * 100) : null;
  };

  const latestKey = months.length ? months[months.length - 1] : null;
  const latest = latestKey ? map[latestKey] : null;
  const withData = series.filter(function (s) {
    return s.baht !== null;
  });
  const total = withData.reduce(function (s, x) {
    return s + x.baht;
  }, 0);
  let peak = null,
    low = null;
  withData.forEach(function (s) {
    if (!peak || s.baht > peak.baht) peak = { month: s.key, baht: s.baht };
    if (!low || s.baht < low.baht) low = { month: s.key, baht: s.baht };
  });
  let sb = 0,
    sk = 0;
  series.forEach(function (s) {
    if (s.baht !== null && s.kwh) {
      sb += s.baht;
      sk += s.kwh;
    }
  });
  let yoy = null;
  if (latest) {
    const ly =
      map[
        keyOf(Number(latestKey.slice(0, 4)) - 1, Number(latestKey.slice(5, 7)))
      ];
    if (ly && ly.baht > 0) yoy = r1(((latest.baht - ly.baht) / ly.baht) * 100);
  }
  const kpi = {
    latest: latest
      ? { month: latestKey, baht: latest.baht, kwh: latest.kwh }
      : null,
    mom: latestKey ? momOf(latestKey) : null,
    yoy: yoy,
    avg: withData.length ? Math.round(total / withData.length) : 0,
    total: Math.round(total),
    n: withData.length,
    peak: peak,
    low: low,
    rate: sk > 0 ? Math.round((sb / sk) * 100) / 100 : null,
  };

  const list = months
    .slice()
    .reverse()
    .map(function (k) {
      const b = map[k];
      return {
        month: k,
        baht: b.baht,
        kwh: b.kwh,
        note: b.note || "",
        rate: b.kwh > 0 ? Math.round((b.baht / b.kwh) * 100) / 100 : null,
        mom: momOf(k),
      };
    });

  const rows = airs.map(function (a) {
    const btu = Number(a.btu) || 0;
    if (!btu)
      return { id: a.id, loc: a.loc, btu: 0, kwh: null, baht: null, share: 0 };
    const kwh = (btu / cfg.eer / 1000) * cfg.hours * cfg.days;
    return {
      id: a.id,
      loc: a.loc,
      btu: btu,
      kwh: r1(kwh),
      baht: Math.round(kwh * cfg.rate),
      share: 0,
    };
  });
  const eb = rows.reduce(function (s, x) {
    return s + (x.baht || 0);
  }, 0);
  const ek = rows.reduce(function (s, x) {
    return s + (x.kwh || 0);
  }, 0);
  rows.forEach(function (x) {
    x.share = eb > 0 && x.baht !== null ? Math.round((x.baht / eb) * 100) : 0;
  });
  rows.sort(function (a, b) {
    return (b.baht === null ? -1 : b.baht) - (a.baht === null ? -1 : a.baht);
  });

  return {
    cur: curKey,
    cfg: cfg,
    series: series,
    kpi: kpi,
    bills: list,
    est: {
      rows: rows,
      kwh: r1(ek),
      baht: eb,
      missing: rows.filter(function (x) {
        return x.baht === null;
      }).length,
    },
  };
}

function powerData_() {
  return powerStats_(
    readPowerRows_(powerSheet_()),
    getCfg_(),
    getAirs_(),
    monthKey_(new Date()),
  );
}

function savePower_(o, actor) {
  const month = clean_(o.month, 7);
  if (
    !/^\d{4}-(0[1-9]|1[0-2])$/.test(month) ||
    Number(month.slice(0, 4)) < 2000
  )
    throw userErr_("เดือนไม่ถูกต้อง (รูปแบบ 2026-09)");
  const baht = cleanNum_(o.baht, 0, 100000000, "ค่าไฟ (บาท)", true);
  const kwh = cleanNum_(o.kwh, 0, 10000000, "หน่วยไฟ (kWh)", false);
  const note = clean_(o.note, 200);
  return withLock_(function () {
    const sh = powerSheet_();
    const ex = readPowerRows_(sh)
      .filter(function (r) {
        return r.month === month;
      })
      .pop();
    const row = ex ? ex.row : sh.getLastRow() + 1;
    sh.getRange(row, 1, 1, POWER_HEADER.length)
      .setNumberFormats([PW_FMT])
      .setValues([
        [month, baht, kwh === null ? "" : kwh, note, new Date(), actor || ""],
      ]);
    audit_(
      actor,
      "web",
      "[ค่าไฟ] " + (ex ? "แก้ไข " : "บันทึก ") + month + " = " + baht,
    );
    return { updated: !!ex, month: month, baht: baht, kwh: kwh };
  });
}

function apiPowerData(key) {
  return api_("powerData", key, false, powerData_);
}
function apiPowerSave(key, o) {
  return api_("powerSave", key, true, function () {
    const r = savePower_(o || {}, "dashboard");
    const d = powerData_();
    d.saved = r;
    return d;
  });
}
function apiPowerDelete(key, month) {
  return api_("powerDelete", key, true, function () {
    const m = clean_(month, 7);
    withLock_(function () {
      const sh = powerSheet_();
      const hit = readPowerRows_(sh).filter(function (r) {
        return r.month === m;
      });
      if (!hit.length) throw userErr_("ไม่พบข้อมูลเดือน " + m);
      hit
        .map(function (r) {
          return r.row;
        })
        .sort(function (a, b) {
          return b - a;
        })
        .forEach(function (row) {
          sh.deleteRow(row);
        });
      audit_("dashboard", "web", "[ค่าไฟ] ลบ " + m);
    });
    return powerData_();
  });
}
function apiCfgSave(key, o) {
  return api_("cfgSave", key, true, function () {
    o = o || {};
    const c = {
      rate: cleanNum_(o.rate, 0.5, 50, "ค่าไฟต่อหน่วย", true),
      hours: cleanNum_(o.hours, 0.5, 24, "ชั่วโมงใช้งานต่อวัน", true),
      days: cleanNum_(o.days, 1, 31, "วันใช้งานต่อเดือน", true),
      eer: cleanNum_(o.eer, 3, 40, "EER", true),
    };
    PropertiesService.getScriptProperties().setProperties({
      CFG_rate: String(c.rate),
      CFG_hours: String(c.hours),
      CFG_days: String(c.days),
      CFG_eer: String(c.eer),
    });
    audit_("dashboard", "web", "[ค่าไฟ] ปรับสมมติฐาน " + JSON.stringify(c));
    return powerData_();
  });
}

function apiHistory(key) {
  return api_("history", key, false, function () {
    return {
      logs: getLogs_()
        .slice(0, 1500)
        .map(function (l) {
          return {
            date: fmt_(l.date),
            ym: monthKey_(l.date),
            id: l.id,
            type: l.type,
            detail: l.detail,
            tech: l.tech,
            cost: l.cost,
          };
        }),
    };
  });
}

function extraCmd_(cmd, args, rest) {
  switch (cmd) {
    case "tech":
      return cmdTech_();
    case "addtech":
      return cmdAddTech_(rest);
    case "bill":
      return cmdBill_(args);
    case "power":
      return cmdPower_();
  }
  return null;
}

function cmdTech_() {
  const ts = techListData_().techs;
  if (!ts.length)
    return 'ยังไม่มีช่างในระบบ เพิ่มด้วย /addtech ชื่อช่าง 081-234-5678 หรือที่ Dashboard แท็บ "ช่าง"';
  return (
    "📞 <b>ช่างล้างแอร์</b>\n" +
    ts
      .map(function (t) {
        return (
          (t.fav ? "⭐ " : "• ") +
          "<b>" +
          esc_(t.name) +
          "</b>\n   " +
          esc_(t.phone) +
          (t.line ? " | LINE: " + esc_(t.line) : "") +
          (t.area ? "\n   " + esc_(t.area) : "")
        );
      })
      .join("\n")
  );
}

function cmdAddTech_(rest) {
  const m = String(rest || "").match(
    /^(.+?)\s+(\+?\d[\d\s\-]{7,16}\d)\s*(\S*)\s*([\s\S]*)$/,
  );
  if (!m)
    throw new Error(
      "รูปแบบ: /addtech ชื่อช่าง เบอร์โทร [LINE ID] [หมายเหตุ]\nเช่น /addtech ช่างสมชาย 081-234-5678 somchai.air",
    );
  saveTech_(
    { name: m[1], phone: m[2], line: m[3], note: m[4] },
    ACTOR_ || "telegram",
  );
  return (
    "✅ เพิ่มช่าง <b>" +
    esc_(clean_(m[1], 80)) +
    "</b> แล้ว\n" +
    esc_(cleanPhone_(m[2]))
  );
}

function cmdBill_(args) {
  args = args.slice();
  let month = monthKey_(new Date());
  if (args.length && /^\d{4}-\d{2}$/.test(args[0])) month = args.shift();
  if (!args.length)
    throw new Error(
      "รูปแบบ: /bill [2026-09] ค่าไฟ(บาท) [หน่วย kWh]\nเช่น /bill 2026-09 5600 1240",
    );
  const r = savePower_(
    { month: month, baht: args[0], kwh: args[1] || "" },
    ACTOR_ || "telegram",
  );
  return (
    "⚡ " +
    (r.updated ? "แก้ไข" : "บันทึก") +
    "ค่าไฟเดือน <b>" +
    esc_(month) +
    "</b> แล้ว: " +
    baht_(r.baht) +
    " บาท" +
    (r.kwh !== null ? " (" + baht_(r.kwh) + " หน่วย)" : "") +
    "\nดูสถิติ: /power"
  );
}

function cmdPower_() {
  const s = powerData_();
  const k = s.kpi;
  if (!k.latest)
    return "⚡ ยังไม่มีข้อมูลค่าไฟ บันทึกด้วย /bill 2026-09 5600 1240 (เดือน ค่าไฟ หน่วย)";
  let t =
    "⚡ <b>สถิติค่าไฟ</b>\nเดือน " +
    esc_(k.latest.month) +
    ": <b>" +
    baht_(k.latest.baht) +
    "</b> บาท" +
    (k.latest.kwh !== null ? " (" + baht_(k.latest.kwh) + " หน่วย)" : "") +
    "\n";
  if (k.mom !== null)
    t += (k.mom > 0 ? "📈 +" : "📉 ") + k.mom + "% เทียบเดือนก่อน\n";
  if (k.yoy !== null)
    t += (k.yoy > 0 ? "📈 +" : "📉 ") + k.yoy + "% เทียบปีก่อน\n";
  t +=
    "เฉลี่ย " +
    k.n +
    " เดือนล่าสุด: " +
    baht_(k.avg) +
    " บาท/เดือน | รวม " +
    baht_(k.total) +
    " บาท";
  if (k.peak)
    t +=
      "\nสูงสุด: " + esc_(k.peak.month) + " (" + baht_(k.peak.baht) + " บาท)";
  return t;
}

function techNote_() {
  try {
    const need =
      getAirs_().some(function (a) {
        return (
          a.alert === "เลยกำหนด" ||
          a.alert === "ใกล้ถึงกำหนด" ||
          (a.status && a.status !== "ปกติ")
        );
      }) || pendingList_().length > 0;
    if (!need) return "";
    const ts = techListData_().techs.slice(0, 3);
    if (!ts.length) return "";
    return (
      "\n📞 <b>ติดต่อช่าง</b>\n" +
      ts
        .map(function (t) {
          return (
            "• " +
            (t.fav ? "⭐ " : "") +
            "<b>" +
            esc_(t.name) +
            "</b> " +
            esc_(t.phone) +
            (t.line ? " | LINE: " + esc_(t.line) : "")
          );
        })
        .join("\n")
    );
  } catch (e) {
    return "";
  }
}
