/**
 * ============================================================
 *  مولّد بيانات الهيكل الإداري والموظفين — Google Apps Script
 * ============================================================
 *  - يقرأ الإعدادات من ورقة «الإعدادات»
 *  - يقرأ القوائم الخام من ورقة «البيانات_المصدرية»
 *  - يختار الدوائر عشوائياً، ويوزّع الأقسام عليها (2–5 لكل دائرة) حسب المجال
 *  - يولّد أسماء ثلاثية عُمانية فريدة ويوزّعها على الأقسام
 *  - آمن لإعادة التشغيل: يمسح أوراق المخرجات ويعيد بناءها في كل مرة
 * ============================================================
 */

const CONFIG = {
  // ضع هنا معرّف ملف الشيت إذا كان السكربت منفصلاً (أُنشئ من script.google.com)
  // المعرّف هو الجزء بين /d/ و /edit في رابط الملف. اتركه فارغاً إذا كان السكربت داخل الشيت.
  SPREADSHEET_ID: '1OTemh_z5x67M0bu2glfIDzl8tQu7tbahTogD4gSEJH4',
  SHEETS: {
    SETTINGS:  'الإعدادات',
    SOURCE:    'البيانات_المصدرية',
    STRUCTURE: 'الهيكل_الإداري',
    EMPLOYEES: 'الموظفون',
    SUMMARY:   'ملخص_الدوائر'
  },
  SETTINGS_CELLS: {
    circles:          'B2',
    departments:      'B3',
    employees:        'B4',
    minDeptPerCircle: 'B5',
    maxDeptPerCircle: 'B6',
    minEmpPerDept:    'B7',
    femaleRatio:      'B8',
    assignDirectors:  'B9'
  },
  // أرقام الأعمدة في ورقة المصدر (1 = A)
  SOURCE_COLS: {
    circleName: 1,  circleDomain: 2,   // A, B
    deptName:   4,  deptDomain:   5,   // D, E
    firstName:  7,  gender:       8,   // G, H
    secondName: 10,                    // J  الاسم الثاني (اسم الأب)
    thirdName:  12,                    // L  الاسم الرابع (القبيلة)
    grandName:  14                     // N  اسم الجد (إن كان العمود فارغاً يُؤخذ من قائمة الآباء)
  },
  LABELS: {
    male: 'ذكر', female: 'أنثى', yes: 'نعم',
    bin: 'بن', bint: 'بنت',
    director: 'مدير الدائرة', head: 'رئيس القسم', staff: 'موظف',
    directorOffice: 'مكتب مدير الدائرة'
  },
  EMP_ID_PREFIX:   'EMP-',
  CIRCLE_PREFIX:   'C',
  DEPT_SIZE_WEIGHT: { min: 0.5, max: 2.5 },  // تفاوت أحجام الأقسام
  HEADER_BG: '#1F4E78',
  HEADER_FG: '#FFFFFF',
  TOTAL_BG:  '#DDEBF7'
};

/* ======================= القائمة ======================= */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('مولّد البيانات')
    .addItem('توليد البيانات', 'generateAll')
    .addItem('التحقق من الإعدادات (بدون تنفيذ)', 'checkSettings')
    .addSeparator()
    .addItem('مسح المخرجات', 'clearOutputs')
    .addToUi();
}

/* ======================= نقاط الدخول ======================= */

function generateAll() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) { notify_('يوجد تنفيذ آخر قيد التشغيل، حاول لاحقاً.'); return; }
  try {
    const s = readSettings_();
    const src = readSource_();
    const errors = validate_(s, src);
    if (errors.length) { notify_('لا يمكن التوليد:\n\n• ' + errors.join('\n• ')); return; }

    const structure = buildStructure_(s, src);       // الدوائر + الأقسام
    const people    = generatePeople_(s, src);       // أسماء ثلاثية فريدة
    const employees = assignEmployees_(s, structure, people);

    writeStructure_(structure);
    writeEmployees_(employees);
    writeSummary_(structure);

    try { getSS_().getSheetByName(CONFIG.SHEETS.STRUCTURE).activate(); } catch (ignore) {}
    notify_('✅ تم التوليد بنجاح\n\n' +
      'الدوائر: ' + structure.length + '\n' +
      'الأقسام: ' + structure.reduce((n, c) => n + c.depts.length, 0) + '\n' +
      'الموظفون: ' + employees.length);
  } catch (e) {
    notify_('حدث خطأ: ' + e.message);
    throw e;
  } finally {
    lock.releaseLock();
  }
}

/** فحص الإعدادات فقط (dry run) دون كتابة أي شيء */
function checkSettings() {
  const s = readSettings_();
  const src = readSource_();
  const errors = validate_(s, src);
  const cap = nameCapacity_(src);
  const info =
    'المتاح في المصدر:\n' +
    '• دوائر: ' + src.circles.length + ' | أقسام: ' + src.depts.length + '\n' +
    '• أسماء أولى: ' + src.maleFirst.length + ' ذكور / ' + src.femaleFirst.length + ' إناث\n' +
    '• أسماء آباء: ' + src.second.length + ' | أجداد: ' + src.grand.length + ' | قبائل: ' + src.third.length + '\n' +
    '• أقصى أسماء فريدة: ' + cap.male + ' ذكور / ' + cap.female + ' إناث\n\n' +
    'نطاق الأقسام المسموح: ' + (s.circles * s.minDept) + ' – ' + (s.circles * s.maxDept) + '\n' +
    'أقل عدد موظفين: ' + minEmployees_(s) + '\n\n';
  notify_(errors.length
    ? info + '⚠️ مشاكل:\n• ' + errors.join('\n• ')
    : info + '✅ الإعدادات صالحة ويمكن التوليد.');
}

function clearOutputs() {
  const ui = getUi_();
  if (ui && ui.alert('سيتم مسح أوراق المخرجات الثلاث. متابعة؟', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const ss = getSS_();
  [CONFIG.SHEETS.STRUCTURE, CONFIG.SHEETS.EMPLOYEES, CONFIG.SHEETS.SUMMARY].forEach(n => {
    const sh = ss.getSheetByName(n);
    if (sh) { if (sh.getFilter()) sh.getFilter().remove(); sh.clearContents(); }
  });
}

/* ======================= القراءة ======================= */

function readSettings_() {
  const sh = getSheetOrThrow_(CONFIG.SHEETS.SETTINGS);
  const v = k => sh.getRange(CONFIG.SETTINGS_CELLS[k]).getValue();
  let ratio = Number(v('femaleRatio')) || 0;
  if (ratio > 1) ratio = ratio / 100;              // لو كُتبت 35 بدل 35%
  return {
    circles:   toInt_(v('circles')),
    depts:     toInt_(v('departments')),
    employees: toInt_(v('employees')),
    minDept:   toInt_(v('minDeptPerCircle')),
    maxDept:   toInt_(v('maxDeptPerCircle')),
    minEmp:    toInt_(v('minEmpPerDept')),
    femaleRatio: Math.min(Math.max(ratio, 0), 1),
    directors: clean_(v('assignDirectors')) === CONFIG.LABELS.yes
  };
}

function readSource_() {
  const sh = getSheetOrThrow_(CONFIG.SHEETS.SOURCE);
  const lastRow = sh.getLastRow();
  const C = CONFIG.SOURCE_COLS;
  const width = Math.max.apply(null, Object.values(C));
  const data = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, width).getValues() : [];
  const col = i => data.map(r => r[i - 1]);

  // أزواج (اسم + مجال) بدون تكرار
  const pairs = (nameCol, domCol) => {
    const map = new Map();
    data.forEach(r => {
      const name = clean_(r[nameCol - 1]);
      if (name && !map.has(name)) map.set(name, { name: name, domain: clean_(r[domCol - 1]).toUpperCase() });
    });
    return Array.from(map.values());
  };

  const maleFirst = [], femaleFirst = [], seenFirst = new Set();
  data.forEach(r => {
    const n = clean_(r[C.firstName - 1]);
    if (!n || seenFirst.has(n)) return;
    seenFirst.add(n);
    const g = normalizeArabic_(clean_(r[C.gender - 1]));
    (g === normalizeArabic_(CONFIG.LABELS.female) ? femaleFirst : maleFirst).push(n);
  });

  return {
    circles: pairs(C.circleName, C.circleDomain),
    depts:   pairs(C.deptName, C.deptDomain),
    maleFirst: maleFirst,
    femaleFirst: femaleFirst,
    second: uniqueClean_(col(C.secondName)),
    third:  uniqueClean_(col(C.thirdName)),
    grand:  (function () {
      const g = uniqueClean_(col(C.grandName));
      return g.length ? g : uniqueClean_(col(C.secondName));
    })()
  };
}

/* ======================= التحقق ======================= */

function minEmployees_(s) {
  return s.depts * s.minEmp + (s.directors ? s.circles : 0);
}

function nameCapacity_(src) {
  const secondSet = new Set(src.second);
  // نتجنّب أن يكون الاسم الأول = اسم الأب (مثل: أحمد أحمد)
  const pairs = list => list.reduce((n, f) => n + src.second.length - (secondSet.has(f) ? 1 : 0), 0);
  return {
    // اسم الجد يختلف عن اسم الأب
    male:   pairs(src.maleFirst)   * Math.max(src.grand.length - 1, 1) * src.third.length,
    female: pairs(src.femaleFirst) * Math.max(src.grand.length - 1, 1) * src.third.length
  };
}

function validate_(s, src) {
  const e = [];
  ['circles', 'depts', 'employees', 'minDept', 'maxDept', 'minEmp'].forEach(k => {
    if (!(s[k] >= 1)) e.push('قيمة غير صالحة في الإعدادات: ' + k);
  });
  if (e.length) return e;

  if (s.maxDept < s.minDept) e.push('أكبر عدد أقسام أقل من أقل عدد.');
  if (s.circles > src.circles.length)
    e.push('عدد الدوائر (' + s.circles + ') أكبر من المتاح في المصدر (' + src.circles.length + ').');
  if (s.depts < s.circles * s.minDept || s.depts > s.circles * s.maxDept)
    e.push('عدد الأقسام يجب أن يكون بين ' + (s.circles * s.minDept) + ' و ' + (s.circles * s.maxDept) + '.');
  if (s.depts > src.depts.length)
    e.push('عدد الأقسام (' + s.depts + ') أكبر من المتاح في المصدر (' + src.depts.length + ').');
  if (s.employees < minEmployees_(s))
    e.push('عدد الموظفين أقل من الحد الأدنى (' + minEmployees_(s) + ').');

  if (!src.second.length || !src.third.length || !src.grand.length) e.push('قائمة أسماء الآباء أو الأجداد أو القبائل فارغة.');
  const females = Math.round(s.employees * s.femaleRatio);
  const males = s.employees - females;
  const cap = nameCapacity_(src);
  // نترك هامشاً 20% حتى لا يطول البحث العشوائي عن أسماء غير مكررة
  if (males   > cap.male * 0.8)   e.push('أسماء الذكور غير كافية لتوليد ' + males + ' اسماً فريداً.');
  if (females > cap.female * 0.8) e.push('أسماء الإناث غير كافية لتوليد ' + females + ' اسماً فريداً.');
  return e;
}

/* ======================= بناء الهيكل ======================= */

/** توزيع عدد الأقسام على الدوائر: كل دائرة بين min و max والمجموع = depts */
function splitDeptCounts_(circles, total, min, max) {
  const counts = new Array(circles).fill(min);
  let remaining = total - circles * min;
  while (remaining > 0) {
    const open = [];
    counts.forEach((c, i) => { if (c < max) open.push(i); });
    counts[open[randInt_(0, open.length - 1)]]++;
    remaining--;
  }
  return counts;
}

function buildStructure_(s, src) {
  const chosen = shuffle_(src.circles).slice(0, s.circles);
  const counts = splitDeptCounts_(s.circles, s.depts, s.minDept, s.maxDept);

  // تجميع الأقسام حسب المجال (بترتيب عشوائي)
  const byDomain = {};
  shuffle_(src.depts).forEach(d => { (byDomain[d.domain] = byDomain[d.domain] || []).push(d); });

  // المرحلة 1: كل دائرة تأخذ من أقسام مجالها
  const result = chosen.map((c, i) => {
    const pool = byDomain[c.domain] || [];
    const take = c.domain ? pool.splice(0, counts[i]) : [];
    return { name: c.name, domain: c.domain, need: counts[i], depts: take };
  });

  // المرحلة 2: أي نقص يُكمَّل من الأقسام المتبقية (مجالات أخرى أو بلا مجال)
  let leftovers = shuffle_([].concat.apply([], Object.values(byDomain)));
  result.forEach(c => {
    while (c.depts.length < c.need) {
      if (!leftovers.length) throw new Error('الأقسام المتاحة لا تكفي لإكمال التوزيع.');
      c.depts.push(leftovers.shift());
    }
  });

  // الرموز
  return result.map((c, i) => {
    const code = CONFIG.CIRCLE_PREFIX + pad_(i + 1, 2);
    return {
      code: code,
      name: c.name,
      depts: c.depts.map((d, j) => ({ code: code + '-' + pad_(j + 1, 2), name: d.name }))
    };
  });
}

/* ======================= توليد الأسماء ======================= */

function generatePeople_(s, src) {
  const females = Math.round(s.employees * s.femaleRatio);
  const males = s.employees - females;
  const used = new Set();
  const list = []
    .concat(uniqueNames_(males,   src.maleFirst,   src, used, CONFIG.LABELS.male))
    .concat(uniqueNames_(females, src.femaleFirst, src, used, CONFIG.LABELS.female));
  return shuffle_(list);
}

function uniqueNames_(count, firstList, src, used, gender) {
  const out = [];
  if (!count) return out;
  if (!firstList.length) throw new Error('لا توجد أسماء أولى للجنس: ' + gender);
  const maxAttempts = count * 200;
  let attempts = 0;
  while (out.length < count) {
    if (++attempts > maxAttempts) throw new Error('تعذّر توليد أسماء فريدة كافية (' + gender + ').');
    const f = pick_(firstList), m = pick_(src.second), g = pick_(src.grand), t = pick_(src.third);
    if (f === m || m === g) continue;          // تجنّب: أحمد أحمد / سالم سعيد سعيد
    const key = f + '|' + m + '|' + g + '|' + t;
    if (used.has(key)) continue;
    used.add(key);
    out.push({ first: f, second: m, grand: g, third: t, gender: gender });
  }
  return out;
}

/* ======================= توزيع الموظفين ======================= */

function assignEmployees_(s, structure, people) {
  const allDepts = [];
  structure.forEach(c => c.depts.forEach(d => allDepts.push(d)));

  const staffTotal = s.employees - (s.directors ? structure.length : 0);
  const sizes = allocateSizes_(allDepts.length, staffTotal, s.minEmp);

  const pool = people.slice();
  const rows = [];
  let di = 0;
  structure.forEach(c => {
    if (s.directors) rows.push(empRow_(pool.shift(), c, null, CONFIG.LABELS.director));
    c.depts.forEach(d => {
      const n = sizes[di++];
      for (let k = 0; k < n; k++) {
        rows.push(empRow_(pool.shift(), c, d, k === 0 ? CONFIG.LABELS.head : CONFIG.LABELS.staff));
      }
    });
  });

  rows.forEach((r, i) => { r.id = CONFIG.EMP_ID_PREFIX + pad_(i + 1, 4); });
  return rows;
}

/** أحجام أقسام متفاوتة بشكل واقعي: حد أدنى ثابت + توزيع الباقي بأوزان عشوائية */
function allocateSizes_(nDepts, total, minEach) {
  const sizes = new Array(nDepts).fill(minEach);
  const extra = total - nDepts * minEach;
  if (extra <= 0) return sizes;
  const W = CONFIG.DEPT_SIZE_WEIGHT;
  const w = sizes.map(() => W.min + Math.random() * (W.max - W.min));
  const sumW = w.reduce((a, b) => a + b, 0);
  const raw = w.map(x => extra * x / sumW);
  let given = 0;
  raw.forEach((x, i) => { const f = Math.floor(x); sizes[i] += f; given += f; });
  raw.map((x, i) => ({ i: i, frac: x - Math.floor(x) }))
     .sort((a, b) => b.frac - a.frac)
     .slice(0, extra - given)
     .forEach(o => sizes[o.i]++);
  return sizes;
}

function empRow_(p, circle, dept, role) {
  return {
    id: '', first: p.first, second: p.second, grand: p.grand, third: p.third,
    full: fullName_(p),
    gender: p.gender,
    circleCode: circle.code, circleName: circle.name,
    deptCode: dept ? dept.code : circle.code + '-00',
    deptName: dept ? dept.name : CONFIG.LABELS.directorOffice,
    role: role
  };
}

/**
 * الاسم الكامل بالصيغة العُمانية المعتادة:
 *   ذكر:  أحمد بن محمد بن سعود الحارثي
 *   أنثى: سلمى بنت أحمد بن سعود الحارثي   (الجد دائماً يُسبق بـ «بن»)
 */
function fullName_(p) {
  const link = p.gender === CONFIG.LABELS.female ? CONFIG.LABELS.bint : CONFIG.LABELS.bin;
  return [p.first, link, p.second, CONFIG.LABELS.bin, p.grand, p.third].join(' ');
}

/* ======================= الكتابة ======================= */

function writeStructure_(structure) {
  const empSheet = "'" + CONFIG.SHEETS.EMPLOYEES + "'";
  const rows = [];
  structure.forEach(c => c.depts.forEach(d => {
    const r = rows.length + 2;
    rows.push([c.code, c.name, d.code, d.name, '=COUNTIF(' + empSheet + '!J:J,C' + r + ')']);
  }));
  writeTable_(CONFIG.SHEETS.STRUCTURE,
    ['رمز الدائرة', 'الدائرة', 'رمز القسم', 'القسم', 'عدد الموظفين'], rows);
}

function writeEmployees_(employees) {
  writeTable_(CONFIG.SHEETS.EMPLOYEES,
    ['الرقم الوظيفي', 'الاسم الأول', 'اسم الأب', 'اسم الجد', 'القبيلة', 'الاسم الكامل',
     'الجنس', 'رمز الدائرة', 'الدائرة', 'رمز القسم', 'القسم', 'المسمى'],
    employees.map(e => [e.id, e.first, e.second, e.grand, e.third, e.full, e.gender,
                        e.circleCode, e.circleName, e.deptCode, e.deptName, e.role]));
}

function writeSummary_(structure) {
  const st = "'" + CONFIG.SHEETS.STRUCTURE + "'", em = "'" + CONFIG.SHEETS.EMPLOYEES + "'";
  const L = CONFIG.LABELS;
  const rows = structure.map((c, i) => {
    const r = i + 2;
    return [c.code, c.name,
      '=COUNTIF(' + st + '!A:A,A' + r + ')',
      '=COUNTIF(' + em + '!H:H,A' + r + ')',
      '=COUNTIFS(' + em + '!H:H,A' + r + ',' + em + '!G:G,"' + L.male + '")',
      '=COUNTIFS(' + em + '!H:H,A' + r + ',' + em + '!G:G,"' + L.female + '")'];
  });
  const last = rows.length + 1;
  rows.push(['', 'الإجمالي',
    '=SUM(C2:C' + last + ')', '=SUM(D2:D' + last + ')',
    '=SUM(E2:E' + last + ')', '=SUM(F2:F' + last + ')']);
  const sh = writeTable_(CONFIG.SHEETS.SUMMARY,
    ['رمز الدائرة', 'الدائرة', 'عدد الأقسام', 'عدد الموظفين', 'ذكور', 'إناث'], rows);
  // صف الإجمالي يتغير موضعه مع عدد الدوائر، لذلك يُميَّز في كل تشغيل
  sh.getRange(rows.length + 1, 1, 1, 6).setFontWeight('bold').setBackground(CONFIG.TOTAL_BG);
}

/**
 * كتابة جدول مع الحفاظ على تنسيق المستخدم:
 * - الورقة الجديدة فقط تأخذ التنسيق الافتراضي (ترويسة، تجميد، عرض الأعمدة)
 * - الورقة الموجودة: تُمسح القيم فقط، ويبقى التنسيق كما هو
 * - الصفوف الجديدة تأخذ تنسيق الصف الثاني (أول صف بيانات) ليبقى الشكل موحّداً
 */
function writeTable_(name, headers, rows) {
  const ss = getSS_();
  let sh = ss.getSheetByName(name);
  const isNew = !sh;
  if (isNew) sh = ss.insertSheet(name);

  const width = headers.length;
  const filter = sh.getFilter();
  if (filter) filter.remove();                     // الفلتر فقط، لا يمس التنسيق
  sh.clearContents();                              // القيم فقط

  sh.getRange(1, 1, 1, width).setValues([headers]);
  if (rows.length) sh.getRange(2, 1, rows.length, width).setValues(rows);

  if (isNew) {
    sh.setRightToLeft(true);
    sh.getRange(1, 1, 1, width)
      .setFontWeight('bold').setBackground(CONFIG.HEADER_BG).setFontColor(CONFIG.HEADER_FG)
      .setHorizontalAlignment('center');
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, width);
  } else if (rows.length > 1) {
    sh.getRange(2, 1, 1, width).copyFormatToRange(sh, 1, width, 3, rows.length + 1);
  }

  sh.getRange(1, 1, rows.length + 1, width).createFilter();
  return sh;
}

/* ======================= أدوات مساعدة ======================= */

/** الملف المستهدف: الملف المرتبط بالسكربت، أو الملف المحدد بالمعرّف في CONFIG */
function getSS_() {
  if (CONFIG.SPREADSHEET_ID) return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('السكربت غير مرتبط بملف شيت. ضع معرّف الملف في CONFIG.SPREADSHEET_ID أعلى السكربت.');
  return ss;
}

/** واجهة المستخدم متاحة فقط إذا كان السكربت مرتبطاً بالشيت */
function getUi_() {
  try { return SpreadsheetApp.getUi(); } catch (e) { return null; }
}

/** رسالة منبثقة في الشيت إن أمكن، وإلا تُكتب في سجل التنفيذ */
function notify_(msg) {
  const ui = getUi_();
  if (ui) ui.alert(msg); else Logger.log(msg);
}

function getSheetOrThrow_(name) {
  const sh = getSS_().getSheetByName(name);
  if (!sh) throw new Error('الورقة غير موجودة: ' + name);
  return sh;
}

/** إزالة المسافات المخفية وعلامات الاتجاه الشائعة في النص العربي */
function clean_(v) {
  return String(v == null ? '' : v)
    .replace(/[\u200B-\u200F\u202A-\u202E\u2066-\u2069\u00A0\uFEFF]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/** توحيد الهمزات لمقارنة الجنس (أنثى/انثى) */
function normalizeArabic_(s) {
  return s.replace(/[أإآ]/g, 'ا').replace(/ى/g, 'ي');
}

function uniqueClean_(arr) {
  return Array.from(new Set(arr.map(clean_).filter(Boolean)));
}

function toInt_(v) { const n = parseInt(v, 10); return isNaN(n) ? 0 : n; }
function pad_(n, len) { return String(n).padStart(len, '0'); }
function randInt_(a, b) { return a + Math.floor(Math.random() * (b - a + 1)); }
function pick_(arr) { return arr[Math.floor(Math.random() * arr.length)]; }

function shuffle_(arr) {
  const a = arr.slice();
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const t = a[i]; a[i] = a[j]; a[j] = t;
  }
  return a;
}
