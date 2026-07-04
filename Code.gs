/***************************************************************
 * BYD CRM - Apps Script API đúng theo file Excel của bạn
 * Bản v8: Fix lưu giờ ghi nhận + giữ số 0 đầu SĐT
 * File dữ liệu chính:
 *  - 00_Users: tài khoản đăng nhập / phân quyền
 *  - 03_Data_Theo_Doi: data khách hàng BYD
 *  - 07_Danh_Muc: danh mục dropdown
 *  - CRM_Sessions: phiên đăng nhập, tự tạo nếu chưa có
 *  - CRM_Logs: nhật ký thao tác, tự tạo nếu chưa có
 ***************************************************************/

const CONFIG = {
  APP_NAME: 'BYD CRM Ngoc Anh',
  USERS_SHEETS: ['00_Users', 'CRM_Users'],
  LEADS_SHEET: '03_Data_Theo_Doi',
  META_SHEET: '07_Danh_Muc',
  SESSIONS_SHEET: 'CRM_Sessions',
  LOGS_SHEET: 'CRM_Logs',
  LEAD_HEADER_ROW: 3,
  LEAD_START_ROW: 4,
  LEAD_COL_COUNT: 25,
  USER_HEADER_ROW: 1,
  SESSION_DAYS: 7,
};

// Cột trong sheet 03_Data_Theo_Doi theo đúng file Excel bạn gửi
const LEAD_COL = {
  stt: 1,
  createdDate: 2,
  department: 3,
  saleName: 4,
  customerName: 5,
  phone: 6,
  area: 7,
  model: 8,
  source: 9,
  interest: 10,
  status: 11,
  careType: 12,
  nextDate: 13,
  note: 14,
  phoneCheck: 15,
  duplicateCheck: 16,
  qualityScore: 17,
  week: 18,
  month: 19,
  year: 20,
  saleId: 21,
  userId: 22,
  createdBy: 23,
  updatedBy: 24,
  createdTime: 25,
};

// Cột trong sheet 00_Users theo đúng file Excel bạn gửi
const USER_COL = {
  userId: 1,
  name: 2,
  username: 3,
  password: 4,
  email: 5,
  role: 6,
  department: 7,
  saleId: 8,
  position: 9,
  phone: 10,
  status: 11,
  createdAt: 12,
  lastLogin: 13,
  note: 14,
};

/*********************** API ENTRY ******************************/

function doGet(e) {
  return json_({
    ok: true,
    app: CONFIG.APP_NAME,
    message: 'BYD CRM API đang hoạt động',
    sheets: {
      users: getUsersSheetName_(),
      leads: CONFIG.LEADS_SHEET,
      meta: CONFIG.META_SHEET,
    },
    time: new Date(),
  });
}

function doPost(e) {
  try {
    const req = parseRequest_(e);
    const action = String(req.action || '').trim();
    if (!action) throw new Error('Thiếu action');

    let result;
    if (action === 'login') {
      result = login_(req);
    } else {
      const user = requireAuth_(req.token);
      if (action === 'bootstrap') result = bootstrap_(user);
      else if (action === 'saveLead') result = saveLead_(user, req.lead || {});
      else if (action === 'deleteLead') result = deleteLead_(user, req.id);
      else if (action === 'saveUser') result = saveUser_(user, req.user || {});
      else if (action === 'logout') result = logout_(user, req.token);
      else result = { ok: false, message: 'Action không hợp lệ: ' + action };
    }
    return json_(result);
  } catch (err) {
    return json_({ ok: false, message: err && err.message ? err.message : String(err) });
  }
}

function parseRequest_(e) {
  if (!e || !e.postData || !e.postData.contents) return {};
  try {
    return JSON.parse(e.postData.contents || '{}');
  } catch (err) {
    throw new Error('Dữ liệu gửi lên không phải JSON hợp lệ');
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj, null, 0))
    .setMimeType(ContentService.MimeType.JSON);
}

/*********************** SETUP ******************************/

function setupOnce() {
  const ss = ss_();
  const usersName = getUsersSheetName_();

  if (!usersName) {
    const users = ss.insertSheet('00_Users');
    users.getRange(1, 1, 1, 14).setValues([[ 
      'UserID','Họ tên','Username','Password','Email','Vai trò','Phòng','Sale_ID','Chức vụ','Số điện thoại','Trạng thái','Ngày tạo','Đăng nhập cuối','Ghi chú'
    ]]);
    users.getRange(2, 1, 1, 14).setValues([[ 
      'U000','Quản trị Marketing','admin','123456','admin@ngocanh.local','Admin','ALL','','Admin/MKT','','Active',new Date(),'','Tài khoản quản trị toàn quyền'
    ]]);
    users.setFrozenRows(1);
  }

  const sessions = getOrCreateSheet_(CONFIG.SESSIONS_SHEET);
  if (sessions.getLastRow() === 0) {
    sessions.getRange(1, 1, 1, 5).setValues([['token','userId','createdAt','expiredAt','status']]);
    sessions.setFrozenRows(1);
  }

  const logs = getOrCreateSheet_(CONFIG.LOGS_SHEET);
  if (logs.getLastRow() === 0) {
    logs.getRange(1, 1, 1, 7).setValues([['time','userId','userName','action','targetId','before','after']]);
    logs.setFrozenRows(1);
  }

  // Kiểm tra 2 sheet bắt buộc của file CRM
  ensureLeadSheetFormat_(sheet_(CONFIG.LEADS_SHEET));
  sheet_(CONFIG.META_SHEET);

  formatSystemSheets_();
  return 'Đã khởi tạo BYD CRM. Đã kiểm tra cột Giờ ghi nhận và định dạng SĐT dạng Text.';
}

function formatSystemSheets_() {
  [getUsersSheetName_(), CONFIG.SESSIONS_SHEET, CONFIG.LOGS_SHEET].forEach(name => {
    if (!name) return;
    const sh = ss_().getSheetByName(name);
    if (!sh || sh.getLastRow() < 1) return;
    sh.getRange(1, 1, 1, sh.getLastColumn())
      .setBackground('#0b6b63')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sh.autoResizeColumns(1, sh.getLastColumn());
  });
}

/*********************** AUTH ******************************/

function login_(req) {
  const username = norm_(req.username);
  const password = String(req.password || '');
  if (!username || !password) throw new Error('Vui lòng nhập tài khoản và mật khẩu');

  const users = getUsers_();
  const u = users.find(x =>
    (norm_(x.username) === username || norm_(x.email) === username) &&
    String(x.password) === password &&
    x.status !== 'locked'
  );

  if (!u) throw new Error('Sai tài khoản, mật khẩu hoặc tài khoản đã bị khóa');

  const token = Utilities.getUuid();
  const createdAt = new Date();
  const expiredAt = new Date(Date.now() + 1000 * 60 * 60 * 24 * CONFIG.SESSION_DAYS);
  sheet_(CONFIG.SESSIONS_SHEET).appendRow([token, u.id, createdAt, expiredAt, 'active']);
  updateLastLogin_(u.id, createdAt);

  return { ok: true, token, user: safeUser_(u) };
}

function requireAuth_(token) {
  if (!token) throw new Error('Phiên đăng nhập không hợp lệ');
  const sh = sheet_(CONFIG.SESSIONS_SHEET);
  const rows = sh.getDataRange().getValues();

  const session = rows.slice(1).find(r =>
    String(r[0]) === String(token) &&
    String(r[4]).toLowerCase() === 'active' &&
    new Date(r[3]).getTime() > Date.now()
  );

  if (!session) throw new Error('Phiên đăng nhập hết hạn, vui lòng đăng nhập lại');

  const user = getUsers_().find(x => String(x.id) === String(session[1]) && x.status !== 'locked');
  if (!user) throw new Error('Tài khoản không còn hoạt động');
  return user;
}

function logout_(user, token) {
  const sh = sheet_(CONFIG.SESSIONS_SHEET);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(token)) {
      sh.getRange(i + 1, 5).setValue('inactive');
      break;
    }
  }
  return { ok: true };
}

function updateLastLogin_(userId, date) {
  try {
    const sh = usersSheet_();
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][USER_COL.userId - 1]) === String(userId)) {
        sh.getRange(i + 1, USER_COL.lastLogin).setValue(date);
        return;
      }
    }
  } catch (err) {}
}

/*********************** USERS ******************************/

function getUsers_() {
  const sh = usersSheet_();
  const values = sh.getDataRange().getValues();
  return values.slice(1)
    .filter(r => r[USER_COL.userId - 1] || r[USER_COL.username - 1])
    .map(r => ({
      id: String(r[USER_COL.userId - 1] || '').trim(),
      name: String(r[USER_COL.name - 1] || '').trim(),
      username: String(r[USER_COL.username - 1] || '').trim(),
      password: String(r[USER_COL.password - 1] || ''),
      email: String(r[USER_COL.email - 1] || '').trim(),
      role: roleNorm_(r[USER_COL.role - 1]),
      department: String(r[USER_COL.department - 1] || '').trim(),
      saleId: String(r[USER_COL.saleId - 1] || '').trim(),
      position: String(r[USER_COL.position - 1] || '').trim(),
      phone: String(r[USER_COL.phone - 1] || '').trim(),
      status: statusNorm_(r[USER_COL.status - 1]),
      createdAt: r[USER_COL.createdAt - 1],
      lastLogin: r[USER_COL.lastLogin - 1],
      note: String(r[USER_COL.note - 1] || '').trim(),
    }));
}

function safeUser_(u) {
  return {
    id: u.id,
    name: u.name,
    username: u.username,
    email: u.email,
    role: u.role,
    department: u.department,
    saleId: u.saleId,
    position: u.position,
    phone: u.phone,
    status: u.status,
  };
}

function saveUser_(actor, u) {
  if (actor.role !== 'admin') throw new Error('Chỉ Admin được quản lý tài khoản');

  const sh = usersSheet_();
  const values = sh.getDataRange().getValues();
  const id = String(u.id || '').trim() || ('U-' + Utilities.getUuid().slice(0, 8).toUpperCase());

  let row = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][USER_COL.userId - 1]) === id) {
      row = i + 1;
      break;
    }
  }

  const old = row > 0 ? values[row - 1] : null;
  const oldPassword = old ? old[USER_COL.password - 1] : '';
  const oldCreatedAt = old ? old[USER_COL.createdAt - 1] : new Date();
  const oldLastLogin = old ? old[USER_COL.lastLogin - 1] : '';
  const password = String(u.password || '').trim() || oldPassword || '123456';
  const role = roleLabel_(u.role || 'sale');
  const department = String(u.department || '').trim();
  const saleId = String(u.saleId || '').trim() || (roleNorm_(role) === 'sale' ? autoSaleId_() : '');

  const rowData = [[
    id,
    String(u.name || '').trim(),
    String(u.username || '').trim(),
    password,
    String(u.email || '').trim(),
    role,
    department,
    saleId,
    String(u.position || '').trim(),
    String(u.phone || '').trim(),
    statusLabel_(u.status || 'active'),
    oldCreatedAt,
    oldLastLogin,
    String(u.note || '').trim(),
  ]];

  if (row > 0) sh.getRange(row, 1, 1, 14).setValues(rowData);
  else sh.appendRow(rowData[0]);

  log_(actor, old ? 'updateUser' : 'createUser', id, old, u);
  return { ok: true, id };
}

/*********************** BOOTSTRAP ******************************/

function bootstrap_(user) {
  let leads = getLeads_();
  leads = filterLeadsByRole_(user, leads);
  const users = user.role === 'admin' ? getUsers_().map(safeUser_) : [];
  const meta = getMeta_();
  return { ok: true, user: safeUser_(user), leads, users, meta };
}

function filterLeadsByRole_(user, leads) {
  if (user.role === 'admin' || user.role === 'viewer') return leads;

  if (user.role === 'manager') {
    return leads.filter(l => norm_(l.department) === norm_(user.department));
  }

  return leads.filter(l =>
    norm_(l.saleId) === norm_(user.saleId) ||
    norm_(l.userId) === norm_(user.id) ||
    norm_(l.saleName) === norm_(user.name) ||
    norm_(l.saleEmail) === norm_(user.email)
  );
}

/*********************** LEADS ******************************/

function getLeads_() {
  const sh = sheet_(CONFIG.LEADS_SHEET);
  ensureLeadSheetFormat_(sh);
  const last = sh.getLastRow();
  if (last < CONFIG.LEAD_START_ROW) return [];

  const rows = sh.getRange(CONFIG.LEAD_START_ROW, 1, last - CONFIG.LEAD_START_ROW + 1, CONFIG.LEAD_COL_COUNT).getValues();
  return rows
    .map((r, idx) => mapLeadRow_(r, CONFIG.LEAD_START_ROW + idx))
    .filter(l => l.createdDate || l.customerName || l.phone || l.saleName);
}

function mapLeadRow_(r, rowNumber) {
  return {
    row: rowNumber,
    id: 'R' + rowNumber,
    stt: r[LEAD_COL.stt - 1],
    createdDate: dateOut_(r[LEAD_COL.createdDate - 1]),
    department: String(r[LEAD_COL.department - 1] || '').trim(),
    saleName: String(r[LEAD_COL.saleName - 1] || '').trim(),
    customerName: String(r[LEAD_COL.customerName - 1] || '').trim(),
    phone: phoneOut_(r[LEAD_COL.phone - 1]),
    area: String(r[LEAD_COL.area - 1] || '').trim(),
    model: String(r[LEAD_COL.model - 1] || '').trim(),
    source: String(r[LEAD_COL.source - 1] || '').trim(),
    interest: String(r[LEAD_COL.interest - 1] || '').trim(),
    status: String(r[LEAD_COL.status - 1] || '').trim(),
    careType: String(r[LEAD_COL.careType - 1] || '').trim(),
    nextDate: dateOut_(r[LEAD_COL.nextDate - 1]),
    note: String(r[LEAD_COL.note - 1] || '').trim(),
    phoneCheck: String(r[LEAD_COL.phoneCheck - 1] || '').trim(),
    duplicateCheck: String(r[LEAD_COL.duplicateCheck - 1] || '').trim(),
    qualityScore: r[LEAD_COL.qualityScore - 1],
    week: r[LEAD_COL.week - 1],
    month: r[LEAD_COL.month - 1],
    year: r[LEAD_COL.year - 1],
    saleId: String(r[LEAD_COL.saleId - 1] || '').trim(),
    userId: String(r[LEAD_COL.userId - 1] || '').trim(),
    createdBy: String(r[LEAD_COL.createdBy - 1] || '').trim(),
    updatedBy: String(r[LEAD_COL.updatedBy - 1] || '').trim(),
    createdTime: timeOut_(r[LEAD_COL.createdTime - 1]),
  };
}

function saveLead_(user, lead) {
  if (user.role === 'viewer') throw new Error('Tài khoản chỉ xem không được thêm/sửa khách');

  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);

  try {
    const sh = sheet_(CONFIG.LEADS_SHEET);
    ensureLeadSheetFormat_(sh);
    const before = lead.id ? getLeadById_(lead.id) : null;
    if (before && !canEditLead_(user, before)) throw new Error('Bạn không có quyền sửa khách này');

    let row = getRowFromId_(lead.id);
    const isNew = !row;
    if (!row) row = Math.max(sh.getLastRow() + 1, CONFIG.LEAD_START_ROW);

    const assigned = resolveAssignment_(user, lead, before);
    const createdDate = parseDate_(lead.createdDate || (before && before.createdDate) || new Date());
    const nextDate = parseDate_(lead.nextDate || '');
    const phone = phoneClean_(lead.phone || (before && before.phone) || '');
    let createdTime = normalizeTime_(lead.createdTime || lead.time || lead.gio || lead.gioPhatSinh || lead.recordedTime || (before && before.createdTime) || '');
    if (isNew && !createdTime) {
      createdTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss');
    }

    const rowData = [[
      row - CONFIG.LEAD_START_ROW,
      createdDate,
      assigned.department,
      assigned.saleName,
      String(lead.customerName || '').trim(),
      phone,
      String(lead.area || '').trim(),
      String(lead.model || '').trim(),
      String(lead.source || '').trim(),
      String(lead.interest || '').trim(),
      String(lead.status || '').trim(),
      String(lead.careType || '').trim(),
      nextDate,
      String(lead.note || '').trim(),
      phoneCheck_(phone),
      duplicateCheck_(phone, row),
      qualityScore_(lead, phone),
      weekNum_(createdDate),
      monthNum_(createdDate),
      yearNum_(createdDate),
      assigned.saleId,
      assigned.userId,
      before && before.createdBy ? before.createdBy : user.id,
      user.id,
      createdTime,
    ]];

    // BẮT BUỘC đặt cột SĐT và cột giờ về dạng text trước/sau khi ghi,
    // nếu không Google Sheets có thể tự xóa số 0 đầu hoặc hiểu giờ sai định dạng.
    sh.getRange(row, LEAD_COL.phone).setNumberFormat('@');
    sh.getRange(row, LEAD_COL.createdTime).setNumberFormat('@');

    sh.getRange(row, 1, 1, CONFIG.LEAD_COL_COUNT).setValues(rowData);

    sh.getRange(row, LEAD_COL.createdDate).setNumberFormat('dd/mm/yyyy');
    sh.getRange(row, LEAD_COL.nextDate).setNumberFormat('dd/mm/yyyy');
    sh.getRange(row, LEAD_COL.phone).setNumberFormat('@').setValue(String(phone || ''));
    sh.getRange(row, LEAD_COL.createdTime).setNumberFormat('@').setValue(String(createdTime || ''));

    log_(user, isNew ? 'createLead' : 'updateLead', 'R' + row, before, lead);
    return { ok: true, id: 'R' + row, createdTime: createdTime };
  } finally {
    lock.releaseLock();
  }
}

function resolveAssignment_(user, lead, before) {
  if (user.role === 'sale') {
    return {
      department: user.department,
      saleName: user.name,
      saleId: user.saleId,
      userId: user.id,
    };
  }

  let department = String(lead.department || (before && before.department) || user.department || '').trim();
  let saleName = String(lead.saleName || (before && before.saleName) || '').trim();

  if (user.role === 'manager') {
    department = user.department;
  }

  const saleUser = findSaleUser_(saleName, department);
  return {
    department,
    saleName,
    saleId: saleUser ? saleUser.saleId : String(lead.saleId || (before && before.saleId) || '').trim(),
    userId: saleUser ? saleUser.id : String(lead.userId || (before && before.userId) || '').trim(),
  };
}

function canEditLead_(user, lead) {
  if (user.role === 'admin') return true;
  if (user.role === 'viewer') return false;
  if (user.role === 'manager') return norm_(lead.department) === norm_(user.department);
  if (user.role === 'sale') {
    return norm_(lead.saleId) === norm_(user.saleId) ||
      norm_(lead.userId) === norm_(user.id) ||
      norm_(lead.saleName) === norm_(user.name);
  }
  return false;
}

function deleteLead_(user, id) {
  if (user.role !== 'admin') throw new Error('Chỉ Admin được xóa dữ liệu');
  const row = getRowFromId_(id);
  if (!row) throw new Error('ID khách không hợp lệ');

  const sh = sheet_(CONFIG.LEADS_SHEET);
  const before = getLeadById_(id);
  sh.deleteRow(row);
  log_(user, 'deleteLead', id, before, null);
  return { ok: true };
}

function getLeadById_(id) {
  const row = getRowFromId_(id);
  if (!row) return null;
  const sh = sheet_(CONFIG.LEADS_SHEET);
  ensureLeadTimeColumn_(sh);
  if (row < CONFIG.LEAD_START_ROW || row > sh.getLastRow()) return null;
  const r = sh.getRange(row, 1, 1, CONFIG.LEAD_COL_COUNT).getValues()[0];
  return mapLeadRow_(r, row);
}

function getRowFromId_(id) {
  if (!id) return null;
  const s = String(id).trim();
  if (!/^R\d+$/.test(s)) return null;
  const row = Number(s.replace('R', ''));
  if (!row || row < CONFIG.LEAD_START_ROW) return null;
  return row;
}

/*********************** META / DANH MUC ******************************/

function getMeta_() {
  const sh = sheet_(CONFIG.META_SHEET);
  const values = sh.getDataRange().getValues();

  // Sheet 07_Danh_Muc của bạn: dòng 3 là tiêu đề, dữ liệu từ dòng 4
  const data = values.slice(3);
  const col = idx => unique_(data.map(r => String(r[idx] || '').trim()).filter(Boolean));

  const departments = col(0);
  const sales = col(1);
  const models = col(2);
  const areas = col(3);
  const sources = col(4);
  const statuses = col(5);
  const interests = col(6);
  const careTypes = col(7);

  // Bảng phụ bên phải: L=Phòng, M=Sale, N=Chức danh
  const saleMap = data
    .map(r => ({ department: String(r[11] || '').trim(), saleName: String(r[12] || '').trim(), title: String(r[13] || '').trim() }))
    .filter(x => x.department && x.saleName);

  return { departments, sales, models, areas, sources, statuses, interests, careTypes, saleMap };
}

/*********************** HELPERS ******************************/

function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Không tìm thấy sheet: ' + name);
  return sh;
}

function getOrCreateSheet_(name) {
  return ss_().getSheetByName(name) || ss_().insertSheet(name);
}

function getUsersSheetName_() {
  const ss = ss_();
  for (const name of CONFIG.USERS_SHEETS) {
    if (ss.getSheetByName(name)) return name;
  }
  return '';
}

function usersSheet_() {
  const name = getUsersSheetName_();
  if (!name) throw new Error('Không tìm thấy sheet tài khoản. Cần có sheet 00_Users hoặc CRM_Users');
  return sheet_(name);
}

function roleNorm_(v) {
  const s = norm_(v);
  if (s === 'admin' || s.includes('quản trị') || s.includes('quan tri')) return 'admin';
  if (s === 'manager' || s.includes('trưởng') || s.includes('truong')) return 'manager';
  if (s === 'viewer' || s.includes('lãnh đạo') || s.includes('lanh dao') || s.includes('xem')) return 'viewer';
  return 'sale';
}

function roleLabel_(v) {
  const r = roleNorm_(v);
  if (r === 'admin') return 'Admin';
  if (r === 'manager') return 'Manager';
  if (r === 'viewer') return 'Viewer';
  return 'Sale';
}

function statusNorm_(v) {
  const s = norm_(v);
  if (s === 'locked' || s === 'inactive' || s.includes('khóa') || s.includes('khoa')) return 'locked';
  return 'active';
}

function statusLabel_(v) {
  return statusNorm_(v) === 'locked' ? 'Locked' : 'Active';
}

function norm_(v) {
  return String(v || '').trim().toLowerCase();
}

function unique_(arr) {
  const out = [];
  const seen = {};
  arr.forEach(v => {
    const k = norm_(v);
    if (!k || seen[k]) return;
    seen[k] = true;
    out.push(v);
  });
  return out;
}

function phoneClean_(v) {
  let p = String(v || '').trim();
  p = p.replace(/[^0-9]/g, '');

  // Google Sheets có thể làm mất số 0 đầu khi ô SĐT bị định dạng Number.
  // Với số điện thoại Việt Nam bị còn 9 số, tự bù lại số 0 ở đầu.
  if (/^\d{9}$/.test(p)) p = '0' + p;

  // Nếu nhập dạng 84xxxxxxxxx thì đổi về 0xxxxxxxxx.
  if (/^84\d{9}$/.test(p)) p = '0' + p.slice(2);

  return p;
}

function phoneOut_(v) {
  if (v === null || v === undefined) return '';
  return phoneClean_(v);
}

function phoneCheck_(p) {
  const phone = phoneClean_(p);
  return /^0\d{9}$/.test(phone) ? 'Đúng' : 'Sai';
}

function duplicateCheck_(phone, currentRow) {
  if (!phone) return 'Không trùng';
  const sh = sheet_(CONFIG.LEADS_SHEET);
  const last = sh.getLastRow();
  if (last < CONFIG.LEAD_START_ROW) return 'Không trùng';

  const phones = sh.getRange(CONFIG.LEAD_START_ROW, LEAD_COL.phone, last - CONFIG.LEAD_START_ROW + 1, 1).getValues().flat();
  const count = phones
    .map((p, i) => ({ phone: phoneClean_(p), row: CONFIG.LEAD_START_ROW + i }))
    .filter(x => x.phone === phone && x.row !== currentRow).length;

  return count > 0 ? 'Trùng' : 'Không trùng';
}

function qualityScore_(lead, phone) {
  let s = 0;
  if (String(lead.customerName || '').trim()) s++;
  if (/^0\d{9}$/.test(String(phone || ''))) s++;
  if (String(lead.model || '').trim()) s++;
  if (String(lead.source || '').trim()) s++;
  if (norm_(lead.interest).includes('nóng') || norm_(lead.interest) === 'hot') s++;
  return s;
}



function ensureLeadSheetFormat_(sh) {
  if (!sh) return;
  ensureLeadTimeColumn_(sh);

  const rows = Math.max(1, sh.getMaxRows() - CONFIG.LEAD_START_ROW + 1);
  sh.getRange(CONFIG.LEAD_START_ROW, LEAD_COL.phone, rows, 1).setNumberFormat('@');
  sh.getRange(CONFIG.LEAD_START_ROW, LEAD_COL.createdTime, rows, 1).setNumberFormat('@');
}

// Chạy hàm này 1 lần nếu các số cũ trong sheet đang bị mất số 0 đầu.
// Hàm sẽ sửa các SĐT 9 số thành 10 số bằng cách thêm 0 ở đầu và định dạng cột SĐT là Text.
function fixPhoneAndTimeOnce() {
  const sh = sheet_(CONFIG.LEADS_SHEET);
  ensureLeadSheetFormat_(sh);

  const last = sh.getLastRow();
  if (last < CONFIG.LEAD_START_ROW) return 'Chưa có dữ liệu khách hàng để xử lý';

  const numRows = last - CONFIG.LEAD_START_ROW + 1;
  const phoneRange = sh.getRange(CONFIG.LEAD_START_ROW, LEAD_COL.phone, numRows, 1);
  const values = phoneRange.getValues();
  const fixed = values.map(row => [phoneClean_(row[0])]);

  phoneRange.setNumberFormat('@');
  phoneRange.setValues(fixed);
  phoneRange.setNumberFormat('@');

  sh.getRange(CONFIG.LEAD_START_ROW, LEAD_COL.createdTime, numRows, 1).setNumberFormat('@');

  return 'Đã định dạng lại cột SĐT dạng Text, tự bù số 0 cho SĐT 9 số và kiểm tra cột Giờ ghi nhận.';
}

function ensureLeadTimeColumn_(sh) {
  if (!sh) return;
  if (sh.getMaxColumns() < LEAD_COL.createdTime) {
    sh.insertColumnsAfter(sh.getMaxColumns(), LEAD_COL.createdTime - sh.getMaxColumns());
  }
  const header = sh.getRange(CONFIG.LEAD_HEADER_ROW, LEAD_COL.createdTime);
  if (!String(header.getValue() || '').trim()) {
    header.setValue('Giờ ghi nhận');
    header.setBackground('#0b6b63').setFontColor('#ffffff').setFontWeight('bold');
  }
  const rows = Math.max(1, sh.getMaxRows() - CONFIG.LEAD_START_ROW + 1);
  sh.getRange(CONFIG.LEAD_START_ROW, LEAD_COL.createdTime, rows, 1).setNumberFormat('@');
}

function normalizeTime_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm:ss');
  }
  const s = String(v || '').trim();
  const m = s.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (!m) return '';
  const hh = ('0' + Math.min(23, Number(m[1]))).slice(-2);
  const mm = ('0' + Math.min(59, Number(m[2]))).slice(-2);
  const ss = ('0' + Math.min(59, Number(m[3] || 0))).slice(-2);
  return hh + ':' + mm + ':' + ss;
}

function timeOut_(v) {
  return normalizeTime_(v);
}

function parseDate_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return v;
  const s = String(v || '').trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  const d = new Date(v);
  return isNaN(d.getTime()) ? v : d;
}

function dateOut_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v);
}

function weekNum_(d) {
  try { return Number(Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'w')); }
  catch (e) { return ''; }
}

function monthNum_(d) {
  try { return new Date(d).getMonth() + 1; }
  catch (e) { return ''; }
}

function yearNum_(d) {
  try { return new Date(d).getFullYear(); }
  catch (e) { return ''; }
}

function findSaleUser_(saleName, department) {
  const users = getUsers_();
  return users.find(u =>
    u.role === 'sale' &&
    norm_(u.name) === norm_(saleName) &&
    (!department || norm_(u.department) === norm_(department))
  ) || null;
}

function autoSaleId_() {
  const users = getUsers_();
  let max = 0;
  users.forEach(u => {
    const m = String(u.saleId || '').match(/S(\d+)/i);
    if (m) max = Math.max(max, Number(m[1]));
  });
  return 'S' + String(max + 1).padStart(3, '0');
}

function log_(user, action, targetId, before, after) {
  try {
    sheet_(CONFIG.LOGS_SHEET).appendRow([
      new Date(),
      user ? user.id : '',
      user ? user.name : '',
      action,
      targetId,
      JSON.stringify(before || ''),
      JSON.stringify(after || ''),
    ]);
  } catch (err) {}
}
