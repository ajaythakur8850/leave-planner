/**
 * Management Leave Planner - Google Apps Script backend
 * -----------------------------------------------------
 * 1. Check the values below (AM E-Code, name, first PIN).
 * 2. Run the function "setup" once (Run button at the top). It creates the sheets and
 *    loads all 20 stores and 83 managers (SM/ASM/DM) with first PIN 1234.
 * 3. Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone.
 * 4. Copy the web app URL into API_URL in index.html.
 */
var AM_EMP_CODE  = '6025872';
var AM_NAME      = 'Ajay Kumar';
var AM_FIRST_PIN = '1234';      // you will be asked to set your own PIN at first sign-in
var TIMEZONE     = 'Asia/Kolkata';
var SHEET_ID     = '';   // the "Leave Planner" Google Sheet
function getSS() { return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); }

var SHEETS = {
  Users:  ['emp', 'name', 'role', 'store', 'pinHash', 'mustChange', 'active', 'dept'],
  Stores: ['code', 'name'],
  Leaves: ['id', 'emp', 'name', 'role', 'store', 'from', 'to', 'days', 'type', 'reason', 'status',
           'smBy', 'smAt', 'smRemark', 'amBy', 'amAt', 'amRemark', 'createdAt']
};

function setup() {
  var ss = getSS();
  Object.keys(SHEETS).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.getRange(1, 1, sh.getMaxRows(), SHEETS[name].length).setNumberFormat('@');
    sh.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SECRET')) props.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  var db = sheetDb();
  var core = LeaveCore.create(db, makeEnv());
  var exists = db.all('Users').some(function (u) { return String(u.emp).toUpperCase() === AM_EMP_CODE.toUpperCase(); });
  var rec = { name: AM_NAME, role: 'AM', store: '', pinHash: core.hashPin(AM_EMP_CODE, AM_FIRST_PIN), mustChange: 'Y', active: 'Y' };
  if (exists) db.update('Users', 'emp', AM_EMP_CODE, rec);
  else { rec.emp = AM_EMP_CODE.toUpperCase(); db.insert('Users', rec); }
  // Load stores and managers (existing people are skipped, so their PINs are kept)
  var haveStores = {}, haveUsers = {}, added = 0;
  db.all('Stores').forEach(function (s) { haveStores[String(s.code).toUpperCase()] = 1; });
  db.all('Users').forEach(function (u) { haveUsers[String(u.emp).toUpperCase()] = 1; });
  MGT_DATA.stores.forEach(function (s) { if (!haveStores[s[0]]) db.insert('Stores', { code: s[0], name: s[1] }); });
  MGT_DATA.staff.forEach(function (u) {
    if (haveUsers[u.emp]) return;
    db.insert('Users', { emp: u.emp, name: u.name, role: u.role, dept: u.dept, store: u.store,
      pinHash: core.hashPin(u.emp, AM_FIRST_PIN), mustChange: 'Y', active: 'Y' });
    added++;
  });
  Logger.log('Setup done. AM login: ' + AM_EMP_CODE + ' / PIN ' + AM_FIRST_PIN + '. Managers added: ' + added);
}

function makeEnv() {
  return {
    secret: PropertiesService.getScriptProperties().getProperty('SECRET') || 'not-set-run-setup',
    today: function () { return Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd'); },
    now: function () { return Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd HH:mm"); }
  };
}

function sheetDb() {
  var ss = getSS(), cache = {};
  function fmt(v) {
    if (v instanceof Date) return Utilities.formatDate(v, TIMEZONE, 'yyyy-MM-dd');
    return v == null ? '' : String(v);
  }
  function read(t) {
    if (cache[t]) return cache[t];
    var sh = ss.getSheetByName(t), vals = sh.getDataRange().getValues(), head = vals[0];
    var rows = vals.slice(1).map(function (r, i) {
      var o = { _row: i + 2 };
      head.forEach(function (k, j) { o[k] = fmt(r[j]); });
      return o;
    }).filter(function (o) { return head.some(function (k) { return o[k] !== ''; }); });
    cache[t] = rows;
    return rows;
  }
  return {
    all: function (t) { return read(t).map(function (r) { return Object.assign({}, r); }); },
    insert: function (t, obj) {
      var sh = ss.getSheetByName(t), head = SHEETS[t];
      var r = sh.getLastRow() + 1;
      if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 50);
      sh.getRange(r, 1, 1, head.length).setNumberFormat('@')
        .setValues([head.map(function (k) { return obj[k] == null ? '' : String(obj[k]); })]);
      delete cache[t];
    },
    update: function (t, key, val, patch) {
      var head = SHEETS[t];
      var row = read(t).filter(function (r) { return String(r[key]).toUpperCase() === String(val).toUpperCase(); })[0];
      if (!row) throw new Error('Record not found');
      var merged = head.map(function (k) { return patch.hasOwnProperty(k) ? String(patch[k]) : row[k]; });
      ss.getSheetByName(t).getRange(row._row, 1, 1, head.length).setValues([merged]);
      delete cache[t];
    }
  };
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var req = JSON.parse(e.postData.contents);
    var core = LeaveCore.create(sheetDb(), makeEnv());
    return json({ ok: true, data: core.handle(req) });
  } catch (err) {
    return json({ ok: false, error: err.message || String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function doGet() {
  return ContentService.createTextOutput('Leave Planner API is running.');
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ================= staff list from MGT_Detail.xlsb ================= */
var MGT_DATA = {"am":["6025872","Ajay Kumar"],"stores":[],"staff":[]}; // staff list is added only inside the private Apps Script

/* ================= shared logic below - do not edit ================= */
// ===== Leave Planner core logic (shared by the web page demo and Google Apps Script) =====
var LeaveCore = (function () {
  function sha256(str) {
    var ascii = unescape(encodeURIComponent(String(str)));
    function rr(v, a) { return (v >>> a) | (v << (32 - a)); }
    var maxWord = Math.pow(2, 32), i, j, result = '';
    var words = [], bitLen = ascii.length * 8;
    var hash = [], k = [], pc = 0, isComp = {};
    for (var c = 2; pc < 64; c++) {
      if (!isComp[c]) {
        for (i = 0; i < 313; i += c) isComp[i] = c;
        hash[pc] = (Math.pow(c, .5) * maxWord) | 0;
        k[pc++] = (Math.pow(c, 1 / 3) * maxWord) | 0;
      }
    }
    hash = hash.slice(0, 8);
    ascii += '\x80';
    while (ascii.length % 64 - 56) ascii += '\x00';
    for (i = 0; i < ascii.length; i++) {
      j = ascii.charCodeAt(i);
      words[i >> 2] |= j << ((3 - i) % 4) * 8;
    }
    words[words.length] = ((bitLen / maxWord) | 0);
    words[words.length] = (bitLen);
    for (j = 0; j < words.length;) {
      var w = words.slice(j, j += 16), old = hash;
      hash = hash.slice(0, 8);
      for (i = 0; i < 64; i++) {
        var w15 = w[i - 15], w2 = w[i - 2], a = hash[0], e = hash[4];
        var t1 = hash[7] + (rr(e, 6) ^ rr(e, 11) ^ rr(e, 25)) + ((e & hash[5]) ^ ((~e) & hash[6])) + k[i] +
          (w[i] = (i < 16) ? w[i] : (w[i - 16] + (rr(w15, 7) ^ rr(w15, 18) ^ (w15 >>> 3)) + w[i - 7] + (rr(w2, 17) ^ rr(w2, 19) ^ (w2 >>> 10))) | 0);
        var t2 = (rr(a, 2) ^ rr(a, 13) ^ rr(a, 22)) + ((a & hash[1]) ^ (a & hash[2]) ^ (hash[1] & hash[2]));
        hash = [(t1 + t2) | 0].concat(hash);
        hash[4] = (hash[4] + t1) | 0;
      }
      for (i = 0; i < 8; i++) hash[i] = (hash[i] + old[i]) | 0;
    }
    for (i = 0; i < 8; i++) for (j = 3; j + 1; j--) { var b = (hash[i] >> (j * 8)) & 255; result += ((b < 16) ? 0 : '') + b.toString(16); }
    return result;
  }

  var STAFF_ROLES = ['SM', 'ASM', 'DM'];
  var LEAVE_TYPES = ['PL', 'CL', 'SL', 'Comp Off', 'Other'];
  var TOKEN_HOURS = 12;
  var DEFAULT_PIN = '1234';

  function create(db, env) {
    function fail(msg) { throw new Error(msg); }
    function up(s) { return String(s == null ? '' : s).trim().toUpperCase(); }
    function isDate(s) { return /^\d{4}-\d{2}-\d{2}$/.test(String(s)); }
    function dayNum(s) { var p = s.split('-'); return Date.UTC(+p[0], +p[1] - 1, +p[2]) / 86400000; }
    function hashPin(emp, pin) { return sha256(env.secret + '|pin|' + up(emp) + '|' + String(pin)); }
    function users() { return db.all('Users'); }
    function findUser(emp) { emp = up(emp); return users().filter(function (u) { return up(u.emp) === emp; })[0]; }
    function isActive(u) { return String(u.active || 'Y').toUpperCase() !== 'N'; }
    function pub(u) { return { emp: up(u.emp), name: u.name, role: u.role, dept: u.dept || '', store: up(u.store), active: isActive(u) ? 'Y' : 'N' }; }
    function makeToken(u) {
      var exp = Date.now() + TOKEN_HOURS * 3600000;
      return up(u.emp) + '.' + exp + '.' + sha256(env.secret + '|tok|' + up(u.emp) + '|' + exp + '|' + u.pinHash);
    }
    function auth(token) {
      var p = String(token || '').split('.');
      if (p.length !== 3) fail('Please sign in again.');
      var u = findUser(p[0]);
      if (!u || !isActive(u)) fail('Please sign in again.');
      if (+p[1] < Date.now()) fail('Session expired. Please sign in again.');
      if (sha256(env.secret + '|tok|' + up(u.emp) + '|' + p[1] + '|' + u.pinHash) !== p[2]) fail('Please sign in again.');
      return u;
    }
    function leaves() { return db.all('Leaves'); }
    function canSee(u, store) { return u.role === 'AM' || up(u.store) === up(store); }
    function cleanLeave(l) {
      var o = {}; for (var key in l) if (key !== '_row') o[key] = l[key];
      o.days = +o.days || 0; return o;
    }
    function monthRange(m) {
      if (!/^\d{4}-\d{2}$/.test(String(m))) fail('Bad month');
      var y = +m.slice(0, 4), mo = +m.slice(5, 7);
      var last = new Date(Date.UTC(y, mo, 0)).getUTCDate();
      return [m + '-01', m + '-' + (last < 10 ? '0' : '') + last];
    }

    var actions = {
      login: function (r) {
        var u = findUser(r.emp);
        if (!u || !isActive(u)) fail('Emp Code not found or inactive.');
        if (hashPin(u.emp, r.pin) !== u.pinHash) fail('Wrong PIN.');
        return { token: makeToken(u), user: pub(u), mustChange: String(u.mustChange).toUpperCase() === 'Y' };
      },
      changePin: function (r) {
        var u = auth(r.token);
        if (hashPin(u.emp, r.oldPin) !== u.pinHash) fail('Current PIN is wrong.');
        if (!/^\d{4,6}$/.test(String(r.newPin))) fail('New PIN must be 4 to 6 digits.');
        if (String(r.newPin) === DEFAULT_PIN) fail('Please choose a PIN other than ' + DEFAULT_PIN + '.');
        var h = hashPin(u.emp, r.newPin);
        db.update('Users', 'emp', u.emp, { pinHash: h, mustChange: 'N' });
        u.pinHash = h;
        return { token: makeToken(u) };
      },
      load: function (r) {
        var u = auth(r.token);
        var range = monthRange(r.month);
        var staff = users().filter(function (x) { return x.role !== 'AM' && canSee(u, x.store); }).map(pub);
        var all = leaves();
        var month = all.filter(function (l) {
          return canSee(u, l.store) && l.status !== 'CANCELLED' && l.from <= range[1] && l.to >= range[0];
        }).map(cleanLeave);
        var pending = all.filter(function (l) {
          if (u.role === 'AM') return l.status === 'PENDING_AM' || l.status === 'PENDING_SM';
          if (u.role === 'SM') return l.status === 'PENDING_SM' && up(l.store) === up(u.store) && up(l.emp) !== up(u.emp);
          return false;
        }).map(cleanLeave);
        var mine = all.filter(function (l) { return up(l.emp) === up(u.emp); }).map(cleanLeave)
          .sort(function (a, b) { return a.from < b.from ? 1 : -1; }).slice(0, 30);
        var stores = db.all('Stores').filter(function (s) { return canSee(u, s.code); })
          .map(function (s) { return { code: up(s.code), name: s.name }; });
        return { user: pub(u), staff: staff, leaves: month, pending: pending, mine: mine, stores: stores, today: env.today(), types: LEAVE_TYPES };
      },
      apply: function (r) {
        var u = auth(r.token);
        if (STAFF_ROLES.indexOf(u.role) < 0) fail('Only SM, ASM and DM can apply for leave here.');
        if (!isDate(r.from) || !isDate(r.to)) fail('Please choose both dates.');
        if (r.from > r.to) fail('"From" date must be before "To" date.');
        if (r.from < env.today()) fail('You cannot plan leave for a past date.');
        var days = dayNum(r.to) - dayNum(r.from) + 1;
        if (days > 31) fail('Maximum 31 days in one request.');
        if (LEAVE_TYPES.indexOf(r.type) < 0) fail('Choose a leave type.');
        var clash = leaves().filter(function (l) {
          return up(l.emp) === up(u.emp) && l.status !== 'REJECTED' && l.status !== 'CANCELLED' && l.from <= r.to && l.to >= r.from;
        });
        if (clash.length) fail('You already have leave on these dates (' + clash[0].from + ' to ' + clash[0].to + ').');
        var hasSM = users().some(function (x) { return x.role === 'SM' && isActive(x) && up(x.store) === up(u.store); });
        var rec = {
          id: 'L' + Date.now().toString(36) + Math.floor(Math.random() * 1e6).toString(36),
          emp: up(u.emp), name: u.name, role: u.role, store: up(u.store),
          from: r.from, to: r.to, days: days, type: r.type, reason: String(r.reason || '').slice(0, 200),
          status: (u.role === 'SM' || !hasSM) ? 'PENDING_AM' : 'PENDING_SM',
          smBy: '', smAt: '', smRemark: '', amBy: '', amAt: '', amRemark: '', createdAt: env.now()
        };
        db.insert('Leaves', rec);
        return { id: rec.id, status: rec.status };
      },
      decide: function (r) {
        var u = auth(r.token);
        var l = leaves().filter(function (x) { return x.id === r.id; })[0];
        if (!l) fail('Leave request not found.');
        var ok = r.decision === 'approve', remark = String(r.remark || '').slice(0, 200), patch;
        if (u.role === 'AM') {
          if (l.status !== 'PENDING_AM' && l.status !== 'PENDING_SM') fail('This request is already ' + l.status + '.');
          patch = { status: ok ? 'APPROVED' : 'REJECTED', amBy: u.name + ' (' + up(u.emp) + ')', amAt: env.now(), amRemark: remark };
        } else if (u.role === 'SM') {
          if (up(l.store) !== up(u.store)) fail('Not your store.');
          if (up(l.emp) === up(u.emp)) fail('You cannot approve your own leave.');
          if (l.status !== 'PENDING_SM') fail('This request is already ' + l.status + '.');
          patch = { status: ok ? 'PENDING_AM' : 'REJECTED', smBy: u.name + ' (' + up(u.emp) + ')', smAt: env.now(), smRemark: remark };
        } else fail('Only SM or AM can approve.');
        db.update('Leaves', 'id', l.id, patch);
        return { status: patch.status };
      },
      cancel: function (r) {
        var u = auth(r.token);
        var l = leaves().filter(function (x) { return x.id === r.id; })[0];
        if (!l || up(l.emp) !== up(u.emp)) fail('Leave request not found.');
        if (l.status === 'REJECTED' || l.status === 'CANCELLED') fail('Already ' + l.status + '.');
        if (l.status === 'APPROVED' && l.from <= env.today()) fail('Leave has already started; ask your AM.');
        db.update('Leaves', 'id', l.id, { status: 'CANCELLED' });
        return { status: 'CANCELLED' };
      },
      saveStore: function (r) {
        var u = auth(r.token);
        if (u.role !== 'AM') fail('Only AM can manage stores.');
        var code = up(r.code);
        if (!code) fail('Store code required.');
        var ex = db.all('Stores').filter(function (s) { return up(s.code) === code; })[0];
        if (ex) db.update('Stores', 'code', ex.code, { name: r.name || '' });
        else db.insert('Stores', { code: code, name: r.name || '' });
        return { code: code };
      },
      saveUser: function (r) {
        var u = auth(r.token);
        if (u.role !== 'AM') fail('Only AM can manage staff.');
        var emp = up(r.emp);
        if (!emp) fail('Emp Code required.');
        if (STAFF_ROLES.indexOf(r.role) < 0) fail('Role must be SM, ASM or DM.');
        if (!up(r.store)) fail('Store required.');
        if (!String(r.name || '').trim()) fail('Name required.');
        var ex = findUser(emp);
        if (ex && ex.role === 'AM') fail('That Emp Code belongs to an AM.');
        var patch = { name: String(r.name).trim(), role: r.role, dept: r.role === 'DM' ? String(r.dept || '') : '', store: up(r.store), active: r.active === 'N' ? 'N' : 'Y' };
        if (!ex || r.resetPin) { patch.pinHash = hashPin(emp, DEFAULT_PIN); patch.mustChange = 'Y'; }
        if (ex) db.update('Users', 'emp', ex.emp, patch);
        else { patch.emp = emp; db.insert('Users', patch); }
        return { emp: emp };
      }
    };

    return {
      hashPin: hashPin,
      handle: function (req) {
        var fn = actions[req && req.action];
        if (!fn) fail('Unknown action');
        return fn(req);
      }
    };
  }

  return { create: create, DEFAULT_PIN: DEFAULT_PIN, sha256: sha256, STAFF_ROLES: STAFF_ROLES, LEAVE_TYPES: LEAVE_TYPES };
})();

