const API_URL = (window.BYD_CRM_CONFIG && window.BYD_CRM_CONFIG.API_URL) || '';

const state = {
  token: localStorage.getItem('byd_crm_token') || '',
  user: null,
  meta: { models: [], areas: [], sources: [], statuses: [], interests: [], cares: [], departments: [], sales: [] },
  leads: [],
  dashboard: null,
  users: [],
  view: 'dashboard',
  filters: { q: '', status: '', model: '', department: '' },
  editingLeadId: null,
  loading: false
};

const root = document.getElementById('root');
const leadDialog = document.getElementById('leadDialog');
const leadForm = document.getElementById('leadForm');
const userDialog = document.getElementById('userDialog');
const userForm = document.getElementById('userForm');

document.addEventListener('click', (e) => {
  if (e.target.matches('[data-close-dialog]')) {
    e.target.closest('dialog').close();
  }
});

leadForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  const lead = formToObject(leadForm);
  const sale = state.meta.sales.find(s => s.login === lead.saleLogin);
  if (sale) {
    lead.sale = sale.name;
    lead.department = sale.department;
  }
  await withLoading(async () => {
    if (state.editingLeadId) {
      await api('updateLead', { leadId: state.editingLeadId, patch: lead });
      toast('Đã cập nhật khách hàng.');
    } else {
      await api('createLead', { lead });
      toast('Đã thêm khách hàng mới.');
    }
    leadDialog.close();
    await loadData();
    render();
  });
});

userForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  await withLoading(async () => {
    await api('saveUser', { user: formToObject(userForm) });
    userDialog.close();
    await loadUsers();
    renderUsers();
    toast('Đã lưu tài khoản.');
  });
});

boot();

async function boot() {
  if (!API_URL || API_URL.includes('PASTE_APPS_SCRIPT')) {
    renderLogin('Bạn cần mở file web/config.js và dán Apps Script Web App URL trước khi dùng.');
    return;
  }
  if (!state.token) {
    renderLogin();
    return;
  }
  try {
    const me = await api('me');
    state.user = me.user;
    await loadData();
    render();
  } catch (err) {
    localStorage.removeItem('byd_crm_token');
    state.token = '';
    renderLogin(err.message || 'Phiên đăng nhập hết hạn.');
  }
}

function renderLogin(message = '') {
  root.innerHTML = `
    <div class="login-wrap">
      <form class="login-card" id="loginForm">
        <div class="brand-row">
          <div class="logo-mark">BYD</div>
          <div>
            <h1>CRM Sale BYD</h1>
            <p>Đăng nhập để cập nhật khách hàng</p>
          </div>
        </div>
        <label>Tài khoản
          <input name="login" autocomplete="username" placeholder="VD: pbh01_sale01" required>
        </label>
        <label>Mật khẩu
          <input name="password" type="password" autocomplete="current-password" placeholder="Nhập mật khẩu" required>
        </label>
        <button class="primary" type="submit">Đăng nhập</button>
        <div class="error" id="loginError">${escapeHtml(message)}</div>
        <p class="help">Tài khoản mặc định sau khi chạy setupOnce(): <b>admin</b> / <b>123456</b>. Hãy đổi mật khẩu sau khi chạy thật.</p>
      </form>
    </div>`;
  document.getElementById('loginForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const data = formToObject(e.currentTarget);
    const errorEl = document.getElementById('loginError');
    errorEl.textContent = '';
    try {
      const res = await api('login', data, false);
      state.token = res.token;
      state.user = res.user;
      state.meta = res.meta || state.meta;
      localStorage.setItem('byd_crm_token', state.token);
      await loadData();
      render();
    } catch (err) {
      errorEl.textContent = err.message || 'Đăng nhập thất bại.';
    }
  });
}

function render() {
  root.innerHTML = `
    <div class="app">
      <aside class="sidebar">
        <div class="brand-row">
          <div class="logo-mark">BYD</div>
          <div>
            <h1>CRM BYD</h1>
            <p>Sale tracking system</p>
          </div>
        </div>
        <nav class="nav">
          ${navButton('dashboard', 'Dashboard')}
          ${navButton('leads', 'Khách hàng BYD')}
          ${isAdmin() ? navButton('users', 'Tài khoản') : ''}
        </nav>
        <div class="user-box">
          <strong>${escapeHtml(state.user.HoTen || state.user.Login)}</strong>
          <span>${escapeHtml(state.user.VaiTro)} · ${escapeHtml(state.user.Phong || '')}</span>
          <button class="ghost" id="logoutBtn">Đăng xuất</button>
        </div>
      </aside>
      <main class="main" id="main"></main>
    </div>`;

  document.querySelectorAll('[data-view]').forEach(btn => {
    btn.addEventListener('click', () => {
      state.view = btn.dataset.view;
      render();
    });
  });
  document.getElementById('logoutBtn').addEventListener('click', logout);

  if (state.view === 'dashboard') renderDashboard();
  if (state.view === 'leads') renderLeads();
  if (state.view === 'users' && isAdmin()) renderUsersPage();
}

function navButton(view, label) {
  return `<button class="${state.view === view ? 'active' : ''}" data-view="${view}">${label}</button>`;
}

function renderDashboard() {
  const d = state.dashboard || {};
  const main = document.getElementById('main');
  main.innerHTML = `
    <div class="topbar">
      <div>
        <h1>Dashboard theo dõi BYD</h1>
        <p>${dashboardSubTitle()}</p>
      </div>
      <button class="primary" id="refreshBtn">Làm mới dữ liệu</button>
    </div>
    <div class="kpi-grid">
      ${kpi('Tổng lead', d.total || 0)}
      ${kpi('Lead nóng', d.hot || 0)}
      ${kpi('Hẹn/Lái thử', d.appointment || 0)}
      ${kpi('Chốt mua/Hồ sơ', d.closed || 0)}
      ${kpi('SĐT sai', d.wrongPhone || 0)}
      ${kpi('Lead trùng', d.duplicate || 0)}
    </div>
    <div class="grid-2">
      ${miniTable('Theo dòng xe BYD', d.byModel || [])}
      ${miniTable('Theo trạng thái khách', d.byStatus || [])}
      ${miniTable('Theo nguồn/kênh', d.bySource || [])}
      ${miniTable(isSale() ? 'Theo Sale' : 'Theo phòng/Sale', isSale() ? (d.bySale || []) : (d.byDepartment || []))}
    </div>`;
  document.getElementById('refreshBtn').addEventListener('click', async () => {
    await withLoading(loadData);
    render();
  });
}

function dashboardSubTitle() {
  if (isSale()) return 'Bạn chỉ thấy khách hàng thuộc tài khoản của mình.';
  if (isManager()) return `Bạn đang xem dữ liệu phòng ${state.user.Phong}.`;
  return 'Admin đang xem toàn bộ dữ liệu 4 phòng bán hàng.';
}

function kpi(label, value) {
  return `<div class="kpi"><span>${label}</span><strong>${value}</strong></div>`;
}

function miniTable(title, rows) {
  const body = rows.length ? rows.slice(0, 8).map(r => `<tr><td>${escapeHtml(r.name)}</td><td><b>${r.count}</b></td></tr>`).join('') : `<tr><td colspan="2" class="empty">Chưa có dữ liệu</td></tr>`;
  return `<section class="card"><h2>${title}</h2><div class="table-wrap"><table><thead><tr><th>Nội dung</th><th>Số lượng</th></tr></thead><tbody>${body}</tbody></table></div></section>`;
}

function renderLeads() {
  const main = document.getElementById('main');
  main.innerHTML = `
    <div class="topbar">
      <div>
        <h1>Khách hàng BYD</h1>
        <p>Sale cập nhật trạng thái, lịch hẹn và ghi chú chăm sóc tại đây.</p>
      </div>
      <button class="primary" id="addLeadBtn">+ Thêm khách</button>
    </div>
    ${isSale() ? '<div class="notice">Bạn chỉ được xem và sửa khách hàng thuộc tài khoản của mình. Các cột phòng/sale được hệ thống tự gán.</div>' : ''}
    <section class="card">
      <div class="toolbar">
        <input id="filterQ" placeholder="Tìm tên khách, SĐT, ghi chú..." value="${escapeAttr(state.filters.q)}">
        <select id="filterStatus">${options([''], state.meta.statuses, state.filters.status, 'Tất cả trạng thái')}</select>
        <select id="filterModel">${options([''], state.meta.models, state.filters.model, 'Tất cả dòng xe')}</select>
        <select id="filterDept">${options([''], state.meta.departments, state.filters.department, 'Tất cả phòng')}</select>
        <button class="ghost" id="clearFilterBtn">Xóa lọc</button>
      </div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Ngày</th><th>Khách hàng</th><th>SĐT</th><th>Dòng xe</th><th>Nguồn</th><th>Mức độ</th><th>Trạng thái</th><th>Hẹn tiếp</th><th>Sale</th><th>Thao tác</th>
            </tr>
          </thead>
          <tbody>${leadRows()}</tbody>
        </table>
      </div>
    </section>`;

  document.getElementById('addLeadBtn').addEventListener('click', () => openLeadDialog());
  ['filterQ', 'filterStatus', 'filterModel', 'filterDept'].forEach(id => {
    const el = document.getElementById(id);
    el.addEventListener('input', () => {
      state.filters.q = document.getElementById('filterQ').value;
      state.filters.status = document.getElementById('filterStatus').value;
      state.filters.model = document.getElementById('filterModel').value;
      state.filters.department = document.getElementById('filterDept').value;
      renderLeads();
    });
  });
  document.getElementById('clearFilterBtn').addEventListener('click', () => {
    state.filters = { q: '', status: '', model: '', department: '' };
    renderLeads();
  });
  document.querySelectorAll('[data-edit-lead]').forEach(btn => {
    btn.addEventListener('click', () => {
      const lead = state.leads.find(l => l.leadId === btn.dataset.editLead);
      openLeadDialog(lead);
    });
  });
  document.querySelectorAll('[data-delete-lead]').forEach(btn => {
    btn.addEventListener('click', async () => {
      if (!confirm('Xóa khách hàng này? Chỉ Admin nên dùng chức năng này.')) return;
      await withLoading(async () => {
        await api('deleteLead', { leadId: btn.dataset.deleteLead });
        await loadData();
        renderLeads();
      });
    });
  });
}

function leadRows() {
  const rows = filteredLeads();
  if (!rows.length) return `<tr><td colspan="10" class="empty">Chưa có khách hàng phù hợp bộ lọc.</td></tr>`;
  return rows.map(l => `
    <tr>
      <td>${fmt(l.date)}</td>
      <td><b>${escapeHtml(l.customer)}</b><br><small>${escapeHtml(l.area || '')}</small></td>
      <td>${escapeHtml(l.phone)}<br><small>${escapeHtml(l.phoneCheck || '')} · ${escapeHtml(l.duplicateCheck || '')}</small></td>
      <td>${escapeHtml(l.model)}</td>
      <td>${escapeHtml(l.source)}</td>
      <td>${interestBadge(l.interest)}</td>
      <td>${statusBadge(l.status)}</td>
      <td>${fmt(l.nextDate)}</td>
      <td>${escapeHtml(l.department)}<br><small>${escapeHtml(l.sale)}</small></td>
      <td class="actions">
        <button class="link-btn" data-edit-lead="${escapeAttr(l.leadId)}">Sửa</button>
        ${isAdmin() ? `<button class="link-btn" data-delete-lead="${escapeAttr(l.leadId)}">Xóa</button>` : ''}
      </td>
    </tr>`).join('');
}

function filteredLeads() {
  const q = normalizeText(state.filters.q);
  return state.leads.filter(l => {
    const hay = normalizeText([l.customer, l.phone, l.area, l.model, l.source, l.status, l.sale, l.note].join(' '));
    if (q && !hay.includes(q)) return false;
    if (state.filters.status && l.status !== state.filters.status) return false;
    if (state.filters.model && l.model !== state.filters.model) return false;
    if (state.filters.department && l.department !== state.filters.department) return false;
    return true;
  }).sort((a, b) => (b.date || '').localeCompare(a.date || ''));
}

function openLeadDialog(lead = null) {
  state.editingLeadId = lead ? lead.leadId : null;
  document.getElementById('leadDialogTitle').textContent = lead ? 'Cập nhật khách BYD' : 'Thêm khách BYD';
  leadForm.reset();
  fillLeadSelects();

  const today = new Date().toISOString().slice(0, 10);
  setFormValue(leadForm, 'date', lead ? lead.date : today);
  setFormValue(leadForm, 'customer', lead ? lead.customer : '');
  setFormValue(leadForm, 'phone', lead ? lead.phone : '');
  setFormValue(leadForm, 'area', lead ? lead.area : '');
  setFormValue(leadForm, 'model', lead ? lead.model : '');
  setFormValue(leadForm, 'source', lead ? lead.source : '');
  setFormValue(leadForm, 'interest', lead ? lead.interest : '');
  setFormValue(leadForm, 'status', lead ? lead.status : '');
  setFormValue(leadForm, 'care', lead ? lead.care : '');
  setFormValue(leadForm, 'nextDate', lead ? lead.nextDate : '');
  setFormValue(leadForm, 'note', lead ? lead.note : '');
  setFormValue(leadForm, 'department', lead ? lead.department : state.user.Phong);
  setFormValue(leadForm, 'saleLogin', lead ? lead.saleLogin : state.user.Login);

  document.querySelectorAll('.admin-field').forEach(el => {
    el.style.display = isSale() ? 'none' : 'grid';
  });
  leadDialog.showModal();
}

function fillLeadSelects() {
  fillSelect(leadForm.elements.department, state.meta.departments, true);
  fillSelect(leadForm.elements.area, state.meta.areas, true);
  fillSelect(leadForm.elements.model, state.meta.models, true);
  fillSelect(leadForm.elements.source, state.meta.sources, true);
  fillSelect(leadForm.elements.interest, state.meta.interests, true);
  fillSelect(leadForm.elements.status, state.meta.statuses, true);
  fillSelect(leadForm.elements.care, state.meta.cares, true);

  const saleOptions = state.meta.sales.map(s => ({ value: s.login, label: `${s.department} - ${s.name}` }));
  fillSelect(leadForm.elements.saleLogin, saleOptions, true);
}

function renderUsersPage() {
  const main = document.getElementById('main');
  main.innerHTML = `
    <div class="topbar">
      <div>
        <h1>Quản lý tài khoản</h1>
        <p>Admin tạo/sửa tài khoản Sale, Trưởng phòng, Viewer tại đây.</p>
      </div>
      <button class="primary" id="addUserBtn">+ Tạo tài khoản</button>
    </div>
    <section class="card" id="usersCard"><div class="empty">Đang tải tài khoản...</div></section>`;
  document.getElementById('addUserBtn').addEventListener('click', () => openUserDialog());
  loadUsers().then(renderUsers).catch(err => {
    document.getElementById('usersCard').innerHTML = `<div class="empty">${escapeHtml(err.message)}</div>`;
  });
}

function renderUsers() {
  const card = document.getElementById('usersCard');
  if (!card) return;
  const rows = state.users.map(u => `
    <tr>
      <td><b>${escapeHtml(u.Login)}</b><br><small>${escapeHtml(u.TrangThai)}</small></td>
      <td>${escapeHtml(u.HoTen)}</td>
      <td>${escapeHtml(u.VaiTro)}</td>
      <td>${escapeHtml(u.Phong)}</td>
      <td>${escapeHtml(u.SaleName)}</td>
      <td>${escapeHtml(u.LastLogin || '')}</td>
      <td class="actions"><button class="link-btn" data-edit-user="${escapeAttr(u.Login)}">Sửa</button><button class="link-btn" data-lock-user="${escapeAttr(u.Login)}">Khóa</button></td>
    </tr>`).join('');
  card.innerHTML = `<div class="table-wrap"><table><thead><tr><th>Login</th><th>Họ tên</th><th>Vai trò</th><th>Phòng</th><th>Tên Sale</th><th>Đăng nhập cuối</th><th>Thao tác</th></tr></thead><tbody>${rows}</tbody></table></div>`;
  document.querySelectorAll('[data-edit-user]').forEach(btn => {
    btn.addEventListener('click', () => openUserDialog(state.users.find(u => u.Login === btn.dataset.editUser)));
  });
  document.querySelectorAll('[data-lock-user]').forEach(btn => {
    btn.addEventListener('click', async () => {
      if (!confirm('Khóa tài khoản này?')) return;
      await withLoading(async () => {
        await api('deleteUser', { login: btn.dataset.lockUser });
        await loadUsers();
        renderUsers();
      });
    });
  });
}

function openUserDialog(user = null) {
  userForm.reset();
  fillSelect(userForm.elements.Phong, ['All', ...state.meta.departments], true);
  if (user) {
    ['Login', 'HoTen', 'VaiTro', 'Phong', 'SaleName', 'TrangThai', 'GhiChu'].forEach(k => setFormValue(userForm, k, user[k] || ''));
    userForm.elements.Password.value = '';
  } else {
    userForm.elements.TrangThai.value = 'Active';
  }
  userDialog.showModal();
}

async function loadData() {
  const res = await api('listLeads');
  state.leads = res.leads || [];
  state.dashboard = res.dashboard || null;
  const meta = await api('meta');
  state.meta = meta.meta || state.meta;
}

async function loadUsers() {
  const res = await api('listUsers');
  state.users = res.users || [];
}

async function api(action, payload = {}, useToken = true) {
  const body = Object.assign({ action }, payload);
  if (useToken) body.token = state.token;
  const res = await fetch(API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(body)
  });
  const text = await res.text();
  let data;
  try { data = JSON.parse(text); } catch (err) { throw new Error('API không trả JSON hợp lệ. Kiểm tra URL Apps Script.'); }
  if (!data.ok) throw new Error(data.message || 'API báo lỗi.');
  return data;
}

async function logout() {
  try { await api('logout'); } catch (err) { /* ignore */ }
  localStorage.removeItem('byd_crm_token');
  state.token = '';
  state.user = null;
  renderLogin();
}

async function withLoading(fn) {
  if (state.loading) return;
  state.loading = true;
  try { await fn(); }
  catch (err) { alert(err.message || String(err)); }
  finally { state.loading = false; }
}

function fillSelect(select, values, blank = false) {
  if (!select) return;
  let html = blank ? '<option value="">-- Chọn --</option>' : '';
  values.forEach(v => {
    if (typeof v === 'object') html += `<option value="${escapeAttr(v.value)}">${escapeHtml(v.label)}</option>`;
    else html += `<option value="${escapeAttr(v)}">${escapeHtml(v)}</option>`;
  });
  select.innerHTML = html;
}

function options(prefix, values, selected, placeholder) {
  let html = `<option value="">${placeholder}</option>`;
  values.forEach(v => { html += `<option value="${escapeAttr(v)}" ${v === selected ? 'selected' : ''}>${escapeHtml(v)}</option>`; });
  return html;
}

function formToObject(form) {
  return Object.fromEntries(new FormData(form).entries());
}

function setFormValue(form, name, value) {
  if (form.elements[name]) form.elements[name].value = value || '';
}

function interestBadge(v) {
  const t = normalizeText(v);
  let cls = 'gray';
  if (t.includes('nong')) cls = 'hot';
  else if (t.includes('am')) cls = 'warm';
  else if (t.includes('lanh')) cls = 'cold';
  return `<span class="badge ${cls}">${escapeHtml(v || 'Chưa rõ')}</span>`;
}

function statusBadge(v) {
  const t = normalizeText(v);
  const cls = (t.includes('chot') || t.includes('ho so') || t.includes('lai thu') || t.includes('showroom')) ? 'ok' : 'gray';
  return `<span class="badge ${cls}">${escapeHtml(v || 'Chưa cập nhật')}</span>`;
}

function isAdmin() { return state.user && normalizeText(state.user.VaiTro) === 'admin'; }
function isManager() { return state.user && ['manager', 'truongphong'].includes(normalizeText(state.user.VaiTro)); }
function isSale() { return state.user && normalizeText(state.user.VaiTro) === 'sale'; }
function fmt(v) { return v ? escapeHtml(v) : '<span style="color:#9ca3af">—</span>'; }
function escapeHtml(v) { return String(v ?? '').replace(/[&<>'"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[c])); }
function escapeAttr(v) { return escapeHtml(v).replace(/`/g, '&#96;'); }
function normalizeText(v) { return String(v || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim(); }
function toast(message) { console.log(message); }
