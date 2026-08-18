/* ============================================================
   ระบบนิเทศภายในโรงเรียนเฉลิมพระเกียรติ กรมหลวงนราธิวาสราชนครินทร์ บางปอประชารักษ์ - Core JavaScript
   รองรับทั้ง GitHub Pages + Apps Script API
   ============================================================ */

const APP = {
  user: null,
  currentPage: 'dashboard',
  toastTimeout: null,
  isGas: typeof google !== 'undefined' && typeof google.script !== 'undefined'
};

// ============================================================
// API Layer - รองรับทั้ง Google Apps Script และ GitHub Pages
// ============================================================

function apiPost(action, data) {
  return new Promise((resolve, reject) => {
    if (APP.isGas) {
      google.script.run
        .withSuccessHandler(resolve)
        .withFailureHandler(reject)
        .doPost(JSON.stringify({ action, ...data }));
    } else {
      fetch(CONFIG.APPS_SCRIPT_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ action, ...data })
      })
      .then(function(r) {
        if (!r.ok) throw new Error('HTTP ' + r.status);
        return r.json();
      })
      .then(resolve)
      .catch(reject);
    }
  });
}

// ============================================================
// Authentication
// ============================================================

function checkAuth() {
  const saved = localStorage.getItem('nited_user');
  if (saved) {
    APP.user = JSON.parse(saved);
    updateUserUI();
    return true;
  }
  return false;
}

function login(username, password) {
  return apiPost('login', { username, password });
}

function handleLogin() {
  const username = document.getElementById('loginUsername').value.trim();
  const password = document.getElementById('loginPassword').value;
  const errorEl = document.getElementById('loginError');
  const btn = document.querySelector('.login-box .btn-primary');

  if (!username || !password) {
    errorEl.textContent = 'กรุณากรอกชื่อผู้ใช้และรหัสผ่าน';
    errorEl.style.display = 'block';
    return;
  }

  btn.disabled = true;
  btn.innerHTML = '<span class="spinner" style="width:20px;height:20px;border-width:2px;"></span> กำลังเข้าสู่ระบบ...';

  login(username, password)
    .then(result => {
      if (result.success) {
        APP.user = result.user;
        localStorage.setItem('nited_user', JSON.stringify(result.user));
        var root = window.location.pathname.indexOf('/pages/') >= 0 ? '../' : '';
        window.location.href = root + 'index.html';
      } else {
        errorEl.textContent = result.message;
        errorEl.style.display = 'block';
      }
    })
    .catch(err => {
      errorEl.textContent = 'เกิดข้อผิดพลาด: ' + (err.message || err);
      errorEl.style.display = 'block';
    })
    .finally(() => {
      btn.disabled = false;
      btn.innerHTML = 'เข้าสู่ระบบ';
    });
}

function logout() {
  APP.user = null;
  localStorage.removeItem('nited_user');
  var root = window.location.pathname.indexOf('/pages/') >= 0 ? '../' : '';
  window.location.href = root + 'index.html';
}

function isAdmin() {
  return APP.user && (APP.user.role === 'admin' || APP.user.role === 'supervisor');
}

function updateUserUI() {
  const nameEl = document.getElementById('userName');
  const roleEl = document.getElementById('userRole');
  if (nameEl && APP.user) nameEl.textContent = APP.user.fullName;
  if (roleEl && APP.user) roleEl.textContent = isAdmin() ? 'ผู้ดูแลระบบ' : APP.user.role;

  document.querySelectorAll('.admin-only').forEach(el => {
    el.style.display = isAdmin() ? '' : 'none';
  });
}

// ============================================================
// Toast Notifications
// ============================================================

function showToast(message, type = 'success') {
  let container = document.querySelector('.toast-container');
  if (!container) {
    container = document.createElement('div');
    container.className = 'toast-container';
    document.body.appendChild(container);
  }

  const icons = { success: '&#10003;', error: '&#10007;', warning: '&#9888;', info: '&#8505;' };
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.innerHTML = `<span>${icons[type] || ''}</span> ${message}`;
  container.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateX(100%)';
    setTimeout(() => toast.remove(), 300);
  }, 3500);
}

// ============================================================
// Modal Helpers
// ============================================================

function openModal(id) {
  const modal = document.getElementById(id);
  if (modal) { modal.classList.add('active'); document.body.style.overflow = 'hidden'; }
}

function closeModal(id) {
  const modal = document.getElementById(id);
  if (modal) { modal.classList.remove('active'); document.body.style.overflow = ''; }
}

function closeAllModals() {
  document.querySelectorAll('.modal-overlay').forEach(m => m.classList.remove('active'));
  document.body.style.overflow = '';
}

// ============================================================
// Tab Helpers
// ============================================================

function switchTab(tabGroup, tabName) {
  const group = document.querySelector(`[data-tab-group="${tabGroup}"]`);
  if (!group) return;
  group.querySelectorAll('.tab-btn').forEach(btn => btn.classList.toggle('active', btn.dataset.tab === tabName));
  group.parentElement.querySelectorAll('.tab-content').forEach(c => c.classList.toggle('active', c.dataset.tabContent === tabName));
}

// ============================================================
// Calendar
// ============================================================

let calendarDate = new Date();
let calendarEvents = [];

function renderCalendar(containerId) {
  const container = document.getElementById(containerId);
  if (!container) return;

  const year = calendarDate.getFullYear();
  const month = calendarDate.getMonth();
  const today = new Date();
  const monthNames = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
  const dayNames = ['อา', 'จ', 'อ', 'พ', 'พฤ', 'ศ', 'ส'];
  const firstDay = new Date(year, month, 1).getDay();
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const daysInPrevMonth = new Date(year, month, 0).getDate();

  let html = `<div class="calendar"><div class="calendar-header">
    <button class="calendar-nav" onclick="prevMonth('${containerId}')">&#8249;</button>
    <h3>${monthNames[month]} ${year + 543}</h3>
    <button class="calendar-nav" onclick="nextMonth('${containerId}')">&#8250;</button>
  </div><div class="calendar-grid">`;

  dayNames.forEach(d => { html += `<div class="calendar-day-header">${d}</div>`; });

  const totalCells = Math.ceil((firstDay + daysInMonth) / 7) * 7;
  for (let i = 0; i < totalCells; i++) {
    let dayNum, dateStr, isOther = false;
    if (i < firstDay) {
      dayNum = daysInPrevMonth - firstDay + i + 1;
      const pm = month === 0 ? 11 : month - 1, py = month === 0 ? year - 1 : year;
      dateStr = `${py}-${String(pm + 1).padStart(2, '0')}-${String(dayNum).padStart(2, '0')}`;
      isOther = true;
    } else if (i >= firstDay + daysInMonth) {
      dayNum = i - firstDay - daysInMonth + 1;
      const nm = month === 11 ? 0 : month + 1, ny = month === 11 ? year + 1 : year;
      dateStr = `${ny}-${String(nm + 1).padStart(2, '0')}-${String(dayNum).padStart(2, '0')}`;
      isOther = true;
    } else {
      dayNum = i - firstDay + 1;
      dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(dayNum).padStart(2, '0')}`;
    }

    const isToday = !isOther && today.getDate() === dayNum && today.getMonth() === month && today.getFullYear() === year;
    const dayEvents = calendarEvents.filter(e => e.date === dateStr);
    let cls = 'calendar-day' + (isOther ? ' other-month' : '') + (isToday ? ' today' : '');

    html += `<div class="${cls}"><div class="day-number">${dayNum}</div>`;
    dayEvents.slice(0, 2).forEach(evt => {
      const sc = evt.status === 'ยืนยันแล้ว' ? 'confirmed' : evt.status === 'นิเทศแล้ว' ? 'completed' : 'pending';
      html += `<div class="calendar-event ${sc}" title="${escapeHtml(evt.title)}">${escapeHtml(evt.title)}</div>`;
    });
    if (dayEvents.length > 2) html += `<div style="font-size:10px;color:#999;">+${dayEvents.length - 2}</div>`;
    html += '</div>';
  }
  html += '</div></div>';
  container.innerHTML = html;
}

function prevMonth(id) { calendarDate.setMonth(calendarDate.getMonth() - 1); loadCalendarEvents(id); }
function nextMonth(id) { calendarDate.setMonth(calendarDate.getMonth() + 1); loadCalendarEvents(id); }

function loadCalendarEvents(containerId) {
  apiPost('getCalendarData', { month: calendarDate.getMonth(), year: calendarDate.getFullYear() })
    .then(result => { if (result.success) { calendarEvents = result.events; renderCalendar(containerId); } })
    .catch(err => console.error('Calendar error:', err));
}

// ============================================================
// Utility Functions
// ============================================================

function formatDate(dateStr) {
  if (!dateStr) return '-';
  const d = new Date(dateStr);
  const mn = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
  return `${d.getDate()} ${mn[d.getMonth()]} ${d.getFullYear() + 543}`;
}

function formatDateTime(dateStr) {
  if (!dateStr) return '-';
  const d = new Date(dateStr);
  return formatDate(dateStr) + ' ' + String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0') + ' น.';
}

function getStatusBadge(status) {
  const map = {
    'รอดำเนินการ': 'badge-pending', 'ยืนยันแล้ว': 'badge-confirmed',
    'นิเทศแล้ว': 'badge-completed', 'ปฏิเสธ': 'badge-rejected',
    'รอตรวจสอบ': 'badge-pending', 'ผ่าน': 'badge-approved',
    'ปรับปรุง': 'badge-revise', 'ดีมาก': 'badge-excellent',
    'ดี': 'badge-good', 'พอใช้': 'badge-fair'
  };
  return `<span class="badge ${map[status] || 'badge-pending'}">${status}</span>`;
}

function getFileTypeIcon(type) {
  const icons = { 'แผนการสอน': '&#128203;', 'สื่อการสอน': '&#127909;', 'ภาพกิจกรรม': '&#128247;', 'คลิปวิดีโอ': '&#127909;' };
  return icons[type] || '&#128196;';
}

function escapeHtml(text) {
  if (!text) return '';
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

// ============================================================
// Print Report
// ============================================================

function printReport(title, content) {
  const pw = window.open('', '_blank');
  pw.document.write(`<!DOCTYPE html><html><head><title>${title}</title>
    <style>
      body{font-family:'Sarabun','Segoe UI',sans-serif;padding:20px;font-size:13px;}
      h2{text-align:center;color:#1a5276;margin-bottom:4px;}
      h3{text-align:center;color:#333;font-size:14px;margin-bottom:20px;}
      table{width:100%;border-collapse:collapse;margin-top:10px;}
      th{background:#1a5276!important;color:white;padding:8px;text-align:left;font-size:12px;}
      td{padding:6px 8px;border:1px solid #ddd;font-size:12px;}
      .badge{padding:2px 8px;border-radius:10px;font-size:11px;}
      .badge-pending{background:#fff3cd;color:#856404;}
      .badge-confirmed{background:#d1ecf1;color:#0c5460;}
      .badge-completed{background:#d4edda;color:#155724;}
      .badge-rejected{background:#f8d7da;color:#721c24;}
      .footer{margin-top:30px;text-align:right;font-size:11px;color:#666;}
      @media print{.no-print{display:none;}}
    </style></head><body>
    <h2>โรงเรียนเฉลิมพระเกียรติ กรมหลวงนราธิวาสราชนครินทร์ บางปอประชารักษ์</h2><h3>${title}</h3>
    <button class="no-print" onclick="window.print()" style="position:fixed;top:10px;right:10px;padding:8px 16px;background:#1a5276;color:white;border:none;border-radius:6px;cursor:pointer;font-size:13px;">&#128424; พิมพ์</button>
    ${content}
    <div class="footer">พิมพ์วันที่: ${new Date().toLocaleDateString('th-TH')} เวลา: ${new Date().toLocaleTimeString('th-TH')}</div>
  </body></html>`);
  pw.document.close();
}

// ============================================================
// Init
// ============================================================

document.addEventListener('DOMContentLoaded', () => {
  document.addEventListener('keydown', e => { if (e.key === 'Escape') closeAllModals(); });
  document.querySelectorAll('.modal-overlay').forEach(o => {
    o.addEventListener('click', e => { if (e.target === o) closeAllModals(); });
  });
});
