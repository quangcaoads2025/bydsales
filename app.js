const CFG = window.BYD_CRM_CONFIG || {};
const API_URL = CFG.API_URL || "";
const LS_KEY = "byd_crm_session_v1";

const state = {
  session: JSON.parse(localStorage.getItem(LS_KEY) || "null"),
  page: "dashboard",
  leads: [], users: [], meta: {}, loading: false,
  filters: { q:"", status:"", model:"", department:"" },
  editingLead: null, editingUser: null
};

const $ = (sel) => document.querySelector(sel);
const app = () => $("#app");
const esc = (v) => String(v ?? "").replace(/[&<>'"]/g, s => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#039;','"':'&quot;'}[s]));
const today = () => new Date().toISOString().slice(0,10);
const fmt = (d) => !d ? "" : String(d).slice(0,10);
const isAdmin = () => state.session?.user?.role === "admin";
const isManager = () => state.session?.user?.role === "manager";
const canManageUsers = () => isAdmin();

async function api(action, payload={}){
  if(!API_URL || API_URL.includes("PASTE_")) throw new Error("Chưa cấu hình API_URL trong file config.js");
  const body = JSON.stringify({ action, token: state.session?.token, ...payload });
  const res = await fetch(API_URL, { method:"POST", body, headers:{"Content-Type":"text/plain;charset=utf-8"} });
  const data = await res.json();
  if(!data.ok) throw new Error(data.message || "Có lỗi xảy ra");
  return data;
}
function toast(msg){
  const el = document.createElement("div"); el.className="toast"; el.textContent=msg; document.body.appendChild(el);
  setTimeout(()=> el.remove(), 2600);
}
function badge(text, type="gray"){
  const t = (text||"").toLowerCase();
  let cls = "b-gray";
  if(t.includes("nóng")||t.includes("hot")) cls="b-hot";
  if(t.includes("ấm")) cls="b-warm";
  if(t.includes("lạnh")) cls="b-cold";
  if(t.includes("chốt")||t.includes("lái thử")||t.includes("hẹn showroom")) cls="b-ok";
  if(type==="brand") cls="b-brand";
  return `<span class="badge ${cls}">${esc(text||"-")}</span>`;
}

function render(){
  if(!state.session) return renderLogin();
  app().innerHTML = `<div class="app-shell">${sidebar()}<main class="main">${pageHtml()}</main></div>${modalHtml()}`;
  bindEvents();
}
function renderLogin(err=""){
  app().innerHTML = `<div class="login-wrap">
    <section class="login-hero">
      <div class="brand-mark"><div class="logo">BYD</div><div>${esc(CFG.BRAND_NAME||"Ngọc Anh BYD")}</div></div>
      <div class="hero-title"><h1>CRM quản lý khách hàng BYD chuyên nghiệp</h1><p>Sale cập nhật khách của mình, trưởng phòng theo dõi đội nhóm, Admin quản trị toàn bộ dữ liệu. Dữ liệu lưu trực tiếp về Google Sheets theo file báo cáo của bạn.</p></div>
      <div class="hero-cards"><div class="hero-card"><b>Phân quyền</b><br>Sale / Trưởng phòng / Admin</div><div class="hero-card"><b>Dashboard</b><br>Lead, xe, nguồn, trạng thái</div><div class="hero-card"><b>Google Sheets</b><br>Dữ liệu dễ kiểm soát</div></div>
    </section>
    <section class="login-panel"><form class="login-card" id="loginForm">
      <h2>Đăng nhập hệ thống</h2><p>Nhập tài khoản được Admin cấp để sử dụng CRM.</p>${err?`<div class="error">${esc(err)}</div>`:""}
      <div class="field"><label>Tài khoản hoặc email</label><input name="username" autocomplete="username" placeholder="admin hoặc email sale" required /></div>
      <div class="field"><label>Mật khẩu</label><input name="password" type="password" autocomplete="current-password" placeholder="••••••••" required /></div>
      <button class="btn btn-primary btn-block" type="submit">Đăng nhập</button>
      <div class="hint">Mặc định sau khi cài API: <b>admin / 123456</b>. Vào quản lý tài khoản để đổi ngay.</div>
    </form></section>
  </div>`;
  $("#loginForm").addEventListener("submit", async e=>{
    e.preventDefault(); const fd = new FormData(e.target);
    try{ const data = await api("login", { username:fd.get("username"), password:fd.get("password") }); state.session={token:data.token,user:data.user}; localStorage.setItem(LS_KEY, JSON.stringify(state.session)); await loadAll(); }
    catch(ex){ renderLogin(ex.message); }
  });
}
function sidebar(){
  const u=state.session.user; const items=[
    ["dashboard","📊","Dashboard"],["leads","🚗","Khách hàng"],["calendar","📅","Lịch hẹn"],["reports","📈","Báo cáo"],
    ...(canManageUsers()?[["users","👥","Tài khoản"]]:[]),["settings","⚙️","Cài đặt"]
  ];
  return `<aside class="sidebar"><div class="side-brand"><div class="logo">BYD</div><div><b>BYD CRM</b><span>Ngọc Anh Sales System</span></div></div><nav class="nav">${items.map(i=>`<button data-page="${i[0]}" class="${state.page===i[0]?"active":""}"><b>${i[1]}</b><span>${i[2]}</span></button>`).join("")}</nav><div class="side-user"><b>${esc(u.name)}</b><span>${esc(u.department||"Toàn hệ thống")}</span><div class="role-pill">${roleLabel(u.role)}</div><button class="btn btn-light btn-block" id="logoutBtn" style="margin-top:12px">Đăng xuất</button></div></aside>`;
}
function roleLabel(r){return ({admin:"Admin/MKT",manager:"Trưởng phòng",sale:"Sale",viewer:"Lãnh đạo"})[r]||r}
function topbar(title,desc,button="") { return `<div class="topbar"><div><h1>${title}</h1><p>${desc}</p></div><div class="actions">${button}</div></div>`; }
function pageHtml(){
  if(state.page==="dashboard") return dashboardPage();
  if(state.page==="leads") return leadsPage();
  if(state.page==="calendar") return calendarPage();
  if(state.page==="reports") return reportsPage();
  if(state.page==="users") return usersPage();
  if(state.page==="settings") return settingsPage();
  return dashboardPage();
}
function visibleLeads(){
  let rows=[...state.leads]; const f=state.filters;
  if(f.q){ const q=f.q.toLowerCase(); rows=rows.filter(x=>[x.customerName,x.phone,x.saleName,x.model,x.area].join(" ").toLowerCase().includes(q)); }
  if(f.status) rows=rows.filter(x=>x.status===f.status);
  if(f.model) rows=rows.filter(x=>x.model===f.model);
  if(f.department) rows=rows.filter(x=>x.department===f.department);
  return rows;
}
function stats(rows=state.leads){
  const hot=rows.filter(x=>x.interest==="Nóng").length;
  const todayLead=rows.filter(x=>fmt(x.createdDate)===today()).length;
  const test=rows.filter(x=>String(x.status).includes("lái thử")).length;
  const appt=rows.filter(x=>String(x.status).includes("Hẹn showroom")).length;
  const closed=rows.filter(x=>String(x.status).includes("Chốt")||String(x.status).includes("Đã mua")).length;
  const bad=rows.filter(x=>String(x.phoneCheck).includes("Sai")||String(x.status).includes("Không nghe")||String(x.status).includes("Không nhu cầu")).length;
  return {total:rows.length,todayLead,hot,test,appt,closed,bad};
}
function dashboardPage(){
  const s=stats();
  return `${topbar("Dashboard BYD","Tổng quan hiệu quả lead theo 4 phòng bán hàng.",`<button class="btn btn-soft" id="refreshBtn">↻ Tải lại</button><button class="btn btn-primary" id="addLeadBtn">+ Thêm khách</button>`)}
  <div class="grid kpis">
    ${kpi("Tổng lead",s.total,"Toàn bộ dữ liệu")}${kpi("Lead hôm nay",s.todayLead,"Phát sinh mới")}${kpi("Lead nóng",s.hot,"Nhu cầu cao")}${kpi("Hẹn showroom",s.appt,"Đã có lịch hẹn")}${kpi("Đã lái thử",s.test,"Khách trải nghiệm")}${kpi("Chốt mua",s.closed,"Kết quả bán hàng")}
  </div>
  <div class="chart-grid" style="margin-top:16px">
    ${chartCard("Theo dòng xe BYD", groupCount(state.leads,"model"))}
    ${chartCard("Theo nguồn/kênh", groupCount(state.leads,"source"))}
    ${chartCard("Theo trạng thái khách", groupCount(state.leads,"status"))}
    ${chartCard("Theo phòng bán hàng", groupCount(state.leads,"department"))}
  </div>
  <div class="card" style="margin-top:16px"><div class="panel-title"><h3>Khách cần chăm sóc gần nhất</h3><button data-page="leads" class="btn btn-light goto">Xem tất cả</button></div>${leadTable(state.leads.slice(0,8), true)}</div>`;
}
function kpi(label,val,mini){ return `<div class="card kpi"><div class="label">${label}</div><div class="value">${val}</div><div class="mini">${mini}</div></div>`; }
function groupCount(rows,key){ const m={}; rows.forEach(r=>{const k=r[key]||"Chưa rõ"; m[k]=(m[k]||0)+1}); return Object.entries(m).sort((a,b)=>b[1]-a[1]).slice(0,8); }
function chartCard(title,data){ const max=Math.max(1,...data.map(x=>x[1])); return `<div class="card"><div class="panel-title"><h3>${title}</h3></div>${data.length?data.map(([k,v])=>`<div class="bar-row"><div>${esc(k)}</div><div class="bar"><span style="width:${Math.max(4,v/max*100)}%"></span></div><b>${v}</b></div>`).join(""):`<div class="empty">Chưa có dữ liệu</div>`}</div>`; }
function leadsPage(){
  return `${topbar("Khách hàng BYD","Quản lý khách hàng, lịch hẹn, trạng thái và ghi chú chăm sóc.",`<button class="btn btn-soft" id="refreshBtn">↻ Tải lại</button><button class="btn btn-primary" id="addLeadBtn">+ Thêm khách</button>`)}${filtersHtml()}<div class="card">${leadTable(visibleLeads())}</div>`;
}
function filtersHtml(){
  const statuses=state.meta.statuses||[], models=state.meta.models||[], deps=state.meta.departments||[];
  return `<div class="toolbar"><input id="qFilter" placeholder="Tìm tên, SĐT, sale, dòng xe..." value="${esc(state.filters.q)}" />${select("statusFilter",statuses,"Trạng thái",state.filters.status)}${select("modelFilter",models,"Dòng xe",state.filters.model)}${select("depFilter",deps,"Phòng",state.filters.department)}</div>`;
}
function select(id, arr, label, val){ return `<select id="${id}"><option value="">${label}</option>${arr.map(x=>`<option ${x===val?"selected":""}>${esc(x)}</option>`).join("")}</select>`; }
function leadTable(rows, compact=false){
  if(!rows.length) return `<div class="empty">Chưa có dữ liệu phù hợp</div>`;
  return `<div class="table-wrap"><table><thead><tr><th>Ngày</th><th>Khách hàng</th><th>SĐT</th><th>Dòng xe</th><th>Nguồn</th><th>Mức độ</th><th>Trạng thái</th><th>Sale</th><th>Hẹn tiếp theo</th><th>Ghi chú</th><th></th></tr></thead><tbody>${rows.map(r=>`<tr><td>${fmt(r.createdDate)}</td><td><b>${esc(r.customerName)}</b><br><span class="hint">${esc(r.area||"")}</span></td><td>${esc(r.phone)}</td><td>${badge(r.model,"brand")}</td><td>${esc(r.source)}</td><td>${badge(r.interest)}</td><td>${badge(r.status)}</td><td>${esc(r.saleName)}<br><span class="hint">${esc(r.department)}</span></td><td>${fmt(r.nextDate)}</td><td>${esc(r.note||"")}</td><td><div class="row-actions"><button class="iconbtn editLead" data-id="${esc(r.id)}">✎</button>${isAdmin()?`<button class="iconbtn deleteLead" data-id="${esc(r.id)}">×</button>`:""}</div></td></tr>`).join("")}</tbody></table></div>`;
}
function calendarPage(){
  const rows=state.leads.filter(x=>x.nextDate).sort((a,b)=>String(a.nextDate).localeCompare(String(b.nextDate)));
  return `${topbar("Lịch hẹn chăm sóc","Theo dõi khách cần gọi lại, hẹn showroom hoặc lái thử.")}<div class="card">${leadTable(rows)}</div>`;
}
function reportsPage(){
  return `${topbar("Báo cáo","Báo cáo nhanh theo phòng, sale, nguồn và dòng xe.",`<button class="btn btn-soft" onclick="exportCsv()">⬇ Xuất CSV</button>`)}<div class="chart-grid">${chartCard("Top Sale",groupCount(state.leads,"saleName"))}${chartCard("Top phòng",groupCount(state.leads,"department"))}${chartCard("Top dòng xe",groupCount(state.leads,"model"))}${chartCard("Top nguồn",groupCount(state.leads,"source"))}</div>`;
}
function usersPage(){
  return `${topbar("Quản lý tài khoản","Admin tạo tài khoản, phân quyền và gán phòng/sale.",`<button class="btn btn-primary" id="addUserBtn">+ Thêm tài khoản</button>`)}<div class="card"><div class="table-wrap"><table><thead><tr><th>Tên</th><th>Tài khoản</th><th>Email</th><th>Vai trò</th><th>Phòng</th><th>Trạng thái</th><th></th></tr></thead><tbody>${state.users.map(u=>`<tr><td><b>${esc(u.name)}</b></td><td>${esc(u.username)}</td><td>${esc(u.email)}</td><td>${badge(roleLabel(u.role),"brand")}</td><td>${esc(u.department||"All")}</td><td>${badge(u.status||"active")}</td><td><button class="iconbtn editUser" data-id="${esc(u.id)}">✎</button></td></tr>`).join("")}</tbody></table></div></div>`;
}
function settingsPage(){
  return `${topbar("Cài đặt triển khai","Thông tin kết nối và hướng dẫn vận hành.")}<div class="card"><h3>Thông tin hệ thống</h3><p><b>API URL:</b> ${API_URL.includes("PASTE_")?"Chưa cấu hình":esc(API_URL)}</p><p><b>File dữ liệu:</b> Google Sheets theo mẫu BYD của bạn. Sheet chính: <code>03_Data_Theo_Doi</code>.</p><p><b>Quyền:</b> Sale chỉ xem/sửa lead của mình; Trưởng phòng xem phòng mình; Admin toàn quyền.</p></div>`;
}
function modalHtml(){
  if(state.editingLead!==null) return leadModal(state.editingLead);
  if(state.editingUser!==null) return userModal(state.editingUser);
  return "";
}
function leadModal(lead){
  const isNew=!lead?.id; const meta=state.meta; const current=lead||{};
  return `<div class="modal-backdrop"><form class="modal" id="leadForm"><div class="modal-head"><h3>${isNew?"Thêm khách BYD":"Cập nhật khách BYD"}</h3><button type="button" class="iconbtn closeModal">×</button></div><div class="modal-body"><div class="form-grid">
  <input type="hidden" name="id" value="${esc(current.id||"")}" />
  ${field("Ngày phát sinh","createdDate","date",fmt(current.createdDate)||today())}
  ${field("Tên khách hàng","customerName","text",current.customerName||"",true)}
  ${field("Số điện thoại","phone","text",current.phone||"",true)}
  ${selectField("Khu vực","area",meta.areas||[],current.area)}
  ${selectField("Dòng xe BYD","model",meta.models||[],current.model)}
  ${selectField("Nguồn/Kênh","source",meta.sources||[],current.source)}
  ${selectField("Mức độ quan tâm","interest",meta.interests||[],current.interest)}
  ${selectField("Trạng thái khách","status",meta.statuses||[],current.status)}
  ${selectField("Hình thức chăm sóc","careType",meta.careTypes||[],current.careType)}
  ${field("Ngày hẹn tiếp theo","nextDate","date",fmt(current.nextDate)||"")}
  ${isAdmin()||isManager()?selectField("Phòng bán hàng","department",meta.departments||[],current.department||state.session.user.department):`<input type="hidden" name="department" value="${esc(current.department||state.session.user.department||"")}" />`}
  ${isAdmin()||isManager()?selectField("Sale phụ trách","saleName",meta.sales||[],current.saleName||state.session.user.name):`<input type="hidden" name="saleName" value="${esc(current.saleName||state.session.user.name)}" />`}
  <div class="field full"><label>Ghi chú chăm sóc</label><textarea name="note">${esc(current.note||"")}</textarea></div>
  </div><div class="actions" style="justify-content:flex-end;margin-top:10px"><button type="button" class="btn btn-light closeModal">Hủy</button><button class="btn btn-primary">Lưu dữ liệu</button></div></div></form></div>`;
}
function field(label,name,type,value,required=false){return `<div class="field"><label>${label}</label><input name="${name}" type="${type}" value="${esc(value)}" ${required?"required":""}/></div>`}
function selectField(label,name,arr,value){return `<div class="field"><label>${label}</label><select name="${name}"><option value="">Chọn</option>${arr.map(x=>`<option ${x===value?"selected":""}>${esc(x)}</option>`).join("")}</select></div>`}
function userModal(user){ const u=user||{}; return `<div class="modal-backdrop"><form class="modal" id="userForm"><div class="modal-head"><h3>${u.id?"Sửa tài khoản":"Thêm tài khoản"}</h3><button type="button" class="iconbtn closeModal">×</button></div><div class="modal-body"><div class="form-grid"><input type="hidden" name="id" value="${esc(u.id||"")}" />${field("Họ tên","name","text",u.name||"",true)}${field("Tài khoản","username","text",u.username||"",true)}${field("Email","email","email",u.email||"")}${field("Mật khẩu mới","password","text","",!u.id)}${selectField("Vai trò","role",["admin","manager","sale","viewer"],u.role||"sale")}${selectField("Phòng","department",state.meta.departments||[],u.department)}${selectField("Trạng thái","status",["active","locked"],u.status||"active")}</div><div class="actions" style="justify-content:flex-end;margin-top:10px"><button type="button" class="btn btn-light closeModal">Hủy</button><button class="btn btn-primary">Lưu tài khoản</button></div></div></form></div>`; }
function bindEvents(){
  document.querySelectorAll(".nav button,.goto").forEach(b=>b.onclick=()=>{state.page=b.dataset.page;render()});
  $("#logoutBtn")?.addEventListener("click",()=>{localStorage.removeItem(LS_KEY);state.session=null;render()});
  $("#refreshBtn")?.addEventListener("click",loadAll);
  $("#addLeadBtn")?.addEventListener("click",()=>{state.editingLead={};render()});
  $("#addUserBtn")?.addEventListener("click",()=>{state.editingUser={};render()});
  document.querySelectorAll(".closeModal").forEach(b=>b.onclick=()=>{state.editingLead=null;state.editingUser=null;render()});
  document.querySelectorAll(".editLead").forEach(b=>b.onclick=()=>{state.editingLead=state.leads.find(x=>x.id===b.dataset.id);render()});
  document.querySelectorAll(".deleteLead").forEach(b=>b.onclick=async()=>{if(confirm("Xóa khách này?")){await api("deleteLead",{id:b.dataset.id});toast("Đã xóa");await loadAll();}});
  document.querySelectorAll(".editUser").forEach(b=>b.onclick=()=>{state.editingUser=state.users.find(x=>x.id===b.dataset.id);render()});
  ["qFilter","statusFilter","modelFilter","depFilter"].forEach(id=>$("#"+id)?.addEventListener("input",e=>{const map={qFilter:"q",statusFilter:"status",modelFilter:"model",depFilter:"department"};state.filters[map[id]]=e.target.value;render()}));
  $("#leadForm")?.addEventListener("submit", async e=>{e.preventDefault(); const o=Object.fromEntries(new FormData(e.target).entries()); await api("saveLead",{lead:o}); toast("Đã lưu khách hàng"); state.editingLead=null; await loadAll();});
  $("#userForm")?.addEventListener("submit", async e=>{e.preventDefault(); const o=Object.fromEntries(new FormData(e.target).entries()); await api("saveUser",{user:o}); toast("Đã lưu tài khoản"); state.editingUser=null; await loadAll();});
}
async function loadAll(){
  try{ const data=await api("bootstrap"); state.leads=data.leads||[]; state.users=data.users||[]; state.meta=data.meta||{}; render(); }
  catch(ex){ toast(ex.message); render(); }
}
function exportCsv(){
  const rows=visibleLeads(); const head=["Ngày","Phòng","Sale","Tên khách","SĐT","Khu vực","Dòng xe","Nguồn","Mức độ","Trạng thái","Ngày hẹn","Ghi chú"];
  const body=rows.map(r=>[fmt(r.createdDate),r.department,r.saleName,r.customerName,r.phone,r.area,r.model,r.source,r.interest,r.status,fmt(r.nextDate),r.note]);
  const csv=[head,...body].map(row=>row.map(v=>`"${String(v??"").replace(/"/g,'""')}"`).join(",")).join("\n");
  const a=document.createElement("a"); a.href=URL.createObjectURL(new Blob(["\ufeff"+csv],{type:"text/csv;charset=utf-8"})); a.download="byd-crm-export.csv"; a.click();
}
window.exportCsv=exportCsv;
if(state.session) loadAll(); else render();
