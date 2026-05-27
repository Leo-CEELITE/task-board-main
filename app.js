(function() {
  'use strict';

  const STATUS_ORDER = ['Pending', 'To Do', 'Doing', 'Waiting', 'Ongoing', 'Delay', 'Done', 'On Hold'];
  const STATUS_CLASS = {
    'Pending': 'status-pending',
    'To Do': 'status-todo',
    'Doing': 'status-doing',
    'Waiting': 'status-waiting',
    'Ongoing': 'status-ongoing',
    'Delay': 'status-delay',
    'Done': 'status-done',
    'On Hold': 'status-cancel'
  };

  const AUTH_KEY = 'ce_auth_token';
  const AUTH_EXPIRY_DAYS = 7;

  const state = {
    tasks: [],
    summary: null,
    view: 'board',
    filters: { owner: '', type: '', priority: '', search: '', dueRange: '' },
    sortField: 'dueDay',
    sortDir: 'asc'
  };

  const $ = (id) => document.getElementById(id);

  function hashPassword(str) {
    let hash = 0;
    for (let i = 0; i < str.length; i++) {
      hash = ((hash << 5) - hash) + str.charCodeAt(i);
      hash = hash & hash;
    }
    return hash.toString(36);
  }

  function checkAuth() {
    const stored = localStorage.getItem(AUTH_KEY);
    if (!stored) return false;
    try {
      const data = JSON.parse(stored);
      if (data.expiry < Date.now()) {
        localStorage.removeItem(AUTH_KEY);
        return false;
      }
      return data.hash === hashPassword(window.CE_CONFIG.PASSWORD);
    } catch (e) {
      return false;
    }
  }

  function saveAuth() {
    const data = {
      hash: hashPassword(window.CE_CONFIG.PASSWORD),
      expiry: Date.now() + AUTH_EXPIRY_DAYS * 24 * 60 * 60 * 1000
    };
    localStorage.setItem(AUTH_KEY, JSON.stringify(data));
  }

  function logout() {
    localStorage.removeItem(AUTH_KEY);
    location.reload();
  }

  function unlock() {
    $('lock-screen').style.display = 'none';
    $('app').style.display = 'block';
    bindEvents();
    loadAndRender();
  }

  function tryPassword() {
    const input = $('lock-input').value;
    const err = $('lock-error');

    if (!window.CE_CONFIG || !window.CE_CONFIG.PASSWORD) {
      err.textContent = 'config.js 未設定 PASSWORD';
      return;
    }

    if (input === window.CE_CONFIG.PASSWORD) {
      saveAuth();
      err.textContent = '';
      unlock();
    } else {
      err.textContent = '密碼錯誤';
      $('lock-input').value = '';
      $('lock-input').focus();
    }
  }

  function initLock() {
    if (checkAuth()) {
      unlock();
      return;
    }

    $('lock-submit').addEventListener('click', tryPassword);
    $('lock-input').addEventListener('keydown', e => {
      if (e.key === 'Enter') tryPassword();
    });
    setTimeout(() => $('lock-input').focus(), 100);
  }

  function getShortStatus(fullStatus) {
    if (!fullStatus) return 'Unknown';
    for (const key of STATUS_ORDER) {
      if (fullStatus.indexOf(key) !== -1) return key;
    }
    return fullStatus.split(' ')[0];
  }

  function getTagClass(type) {
    if (!type) return 'tag-default';
    const t = type.toLowerCase();
    if (t === 'social media') return 'tag-social-media';
    if (t === 'social') return 'tag-social';
    if (t.indexOf('external') !== -1) return 'tag-ad-external';
    if (t.indexOf('internal') !== -1) return 'tag-ad-internal';
    if (t.indexOf('incentive') !== -1) return 'tag-ad-incentive';
    return 'tag-default';
  }

  function getPriorityClass(p) {
    if (p === '高') return 'priority-high';
    if (p === '中') return 'priority-med';
    return 'priority-low';
  }

  function getInitials(name) {
    if (!name) return '?';
    if (/^[A-Za-z]/.test(name)) return name.slice(0, 2).toUpperCase();
    return name.slice(0, 1);
  }

  function getDueInfo(dueDay, today, status) {
    if (!dueDay) return { text: '-', cls: 'due-normal', isRelative: false };
    const isDone = status && status.indexOf('Done') !== -1;
    const isCancel = status && status.indexOf('Canceled') !== -1;

    const d1 = new Date(dueDay);
    const d2 = new Date(today);
    const diffDays = Math.round((d1 - d2) / (1000 * 60 * 60 * 24));

    if (isDone || isCancel) {
      return { text: dueDay.slice(5), cls: 'due-normal', isRelative: false };
    }
    if (diffDays < 0) return { text: '逾期 ' + Math.abs(diffDays) + ' 天', cls: 'due-overdue', isRelative: true };
    if (diffDays === 0) return { text: '今日到期', cls: 'due-today', isRelative: true };
    if (diffDays <= 3) return { text: diffDays + ' 天後', cls: 'due-soon', isRelative: true };
    return { text: dueDay.slice(5), cls: 'due-normal', isRelative: false };
  }

  // 以本機時區格式化日期為 yyyy-mm-dd，避免 toISOString() 的 UTC 偏移
  function fmtDate(d) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${dd}`;
  }

  function escapeHtml(str) {
    if (str == null) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  async function fetchData() {
    if (!window.CE_CONFIG || !window.CE_CONFIG.API_URL || !window.CE_CONFIG.API_TOKEN) {
      throw new Error('config.js 未正確設定 API_URL 或 API_TOKEN');
    }
    const url = window.CE_CONFIG.API_URL + '?token=' + encodeURIComponent(window.CE_CONFIG.API_TOKEN) + '&action=all';
    const res = await fetch(url);
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const data = await res.json();
    if (data.error) throw new Error(data.error);
    return data;
  }

  function renderKPIs() {
    const s = state.summary;
    if (!s) return;

    const filtered = applyFilters(state.tasks);
    const todayStr = s.today;

    const dueTodayCount = filtered.filter(t => {
      if (!t.dueDay) return false;
      if (t.status && t.status.indexOf('Done') !== -1) return false;
      if (t.status && t.status.indexOf('Canceled') !== -1) return false;
      return t.dueDay === todayStr;
    }).length;
    $('kpi-today').textContent = dueTodayCount;

    const overdueCount = filtered.filter(t => {
      if (!t.dueDay) return false;
      if (t.status && t.status.indexOf('Done') !== -1) return false;
      if (t.status && t.status.indexOf('Canceled') !== -1) return false;
      return t.dueDay < todayStr;
    }).length;
    $('kpi-overdue').textContent = overdueCount;

    const doneCount = filtered.filter(t => t.status && t.status.indexOf('Done') !== -1).length;
    const rate = filtered.length > 0 ? (doneCount / filtered.length * 100) : 0;
    $('kpi-completion').textContent = rate.toFixed(2) + '%';

    const [weekStartStr, weekEndStr] = getDateRange('thisWeek', s.today);
    const thisWeek = filtered.filter(t => {
      if (!t.dueDay) return false;
      if (t.status && t.status.indexOf('Done') !== -1) return false;
      return t.dueDay >= weekStartStr && t.dueDay <= weekEndStr;
    });
    $('kpi-this-week').textContent = thisWeek.length;
  }

  function renderFilters() {
    const owners = Array.from(new Set(state.tasks.map(t => t.owner).filter(Boolean))).sort();
    const types = Array.from(new Set(state.tasks.map(t => t.type).filter(Boolean))).sort();

    const ownerSel = $('filter-owner');
    const typeSel = $('filter-type');

    const currentOwner = ownerSel.value;
    const currentType = typeSel.value;

    ownerSel.innerHTML = '<option value="">所有負責人</option>' +
      owners.map(o => `<option value="${escapeHtml(o)}"${o === currentOwner ? ' selected' : ''}>${escapeHtml(o)}</option>`).join('');
    typeSel.innerHTML = '<option value="">所有類型</option>' +
      types.map(t => `<option value="${escapeHtml(t)}"${t === currentType ? ' selected' : ''}>${escapeHtml(t)}</option>`).join('');
  }

  function getDateRange(value, todayStr) {
    if (!value) return null;
    const [ty, tm, td] = todayStr.split('-').map(Number);
    const today = new Date(ty, tm - 1, td);
    const day = today.getDay();
    const daysFromMonday = day === 0 ? 6 : day - 1;

    if (value === 'past7') {
      const start = new Date(today);
      start.setDate(today.getDate() - 7);
      const yesterday = new Date(today);
      yesterday.setDate(today.getDate() - 1);
      return [fmtDate(start), fmtDate(yesterday)];
    }

    if (value === 'lastWeek') {
      const lastMonday = new Date(today);
      lastMonday.setDate(today.getDate() - daysFromMonday - 7);
      const lastSunday = new Date(lastMonday);
      lastSunday.setDate(lastMonday.getDate() + 6);
      return [fmtDate(lastMonday), fmtDate(lastSunday)];
    }

    if (value === 'today') {
      return [todayStr, todayStr];
    }
    if (value === 'thisWeek' || value === 'nextWeek') {
      const monday = new Date(today);
      monday.setDate(today.getDate() - daysFromMonday + (value === 'nextWeek' ? 7 : 0));
      const sunday = new Date(monday);
      sunday.setDate(monday.getDate() + 6);
      return [fmtDate(monday), fmtDate(sunday)];
    }
    if (value === 'next7' || value === 'next30') {
      const end = new Date(today);
      end.setDate(today.getDate() + (value === 'next7' ? 7 : 30));
      return [todayStr, fmtDate(end)];
    }
    if (value === 'thisMonth' || value === 'nextMonth') {
      const offset = value === 'nextMonth' ? 1 : 0;
      const start = new Date(today.getFullYear(), today.getMonth() + offset, 1);
      const end = new Date(today.getFullYear(), today.getMonth() + offset + 1, 0);
      return [fmtDate(start), fmtDate(end)];
    }
    return null;
  }

  function applyFilters(tasks) {
    const todayStr = state.summary ? state.summary.today : fmtDate(new Date());
    const dateRange = getDateRange(state.filters.dueRange, todayStr);

    return tasks.filter(t => {
      if (state.filters.owner && t.owner !== state.filters.owner) return false;
      if (state.filters.type && t.type !== state.filters.type) return false;
      if (state.filters.priority && t.priority !== state.filters.priority) return false;
      if (state.filters.search) {
        const q = state.filters.search.toLowerCase();
        if (!(t.name || '').toLowerCase().includes(q)) return false;
      }
      if (state.filters.dueRange === 'overdue') {
        if (!t.dueDay) return false;
        if (t.status && t.status.indexOf('Done') !== -1) return false;
        if (t.status && t.status.indexOf('Canceled') !== -1) return false;
        if (t.dueDay >= todayStr) return false;
      } else if (dateRange) {
        if (!t.dueDay) return false;
        if (t.dueDay < dateRange[0] || t.dueDay > dateRange[1]) return false;
      }
      return true;
    });
  }

  function renderBoard() {
    const filtered = applyFilters(state.tasks);
    const today = state.summary ? state.summary.today : fmtDate(new Date());

    const cols = ['Delay', 'To Do', 'Doing', 'Waiting', 'Done'];
    const byCol = {};
    cols.forEach(c => byCol[c] = []);

    filtered.forEach(t => {
      const isDone = t.status && t.status.indexOf('Done') !== -1;
      const isCancel = t.status && t.status.indexOf('Canceled') !== -1;
      const isOverdue = t.dueDay && !isDone && !isCancel && t.dueDay < today;

      if (isOverdue) {
        byCol['Delay'].push(t);
      } else {
        const s = getShortStatus(t.status);
        if (byCol[s]) byCol[s].push(t);
      }
    });

    Object.keys(byCol).forEach(col => {
      byCol[col].sort((a, b) => {
        const da = a.dueDay || '9999-12-31';
        const db = b.dueDay || '9999-12-31';
        return da.localeCompare(db);
      });
    });

    const html = cols.map(col => {
      const items = byCol[col];
      return `<div class="board-column">
        <div class="board-column-header">
          <span class="board-column-title">${col}</span>
          <span class="board-column-count">${items.length}</span>
        </div>
        ${items.map(t => renderCard(t, today)).join('')}
      </div>`;
    }).join('');

    $('main-content').innerHTML = '<div class="board">' + html + '</div>';
  }

  function renderCard(t, today) {
    const due = getDueInfo(t.dueDay, today, t.status);
    const tagCls = getTagClass(t.type);
    const priCls = getPriorityClass(t.priority);

    let dateWeekLabel;
    if (due.isRelative && t.dueDay) {
      dateWeekLabel = t.dueDay.slice(5) + (t.week ? '・' + t.week : '');
    } else {
      dateWeekLabel = t.week || '';
    }

    return `<div class="task-card">
      <div class="task-card-title">
        <span class="priority-bar ${priCls}"></span>${escapeHtml(t.name)}
      </div>
      <div class="task-card-meta">
        <span class="task-tag ${tagCls}">${escapeHtml(t.type || '其他')}</span>
        <span class="due-pill ${due.cls}">${due.text}</span>
      </div>
      <div class="task-card-meta">
        <span class="owner-chip">
          <span class="avatar">${escapeHtml(getInitials(t.owner))}</span>${escapeHtml(t.owner || '未指派')}
        </span>
        <span style="font-size:11px; color:#9ca3af;">${escapeHtml(dateWeekLabel)}</span>
      </div>
    </div>`;
  }

  function renderList() {
    const filtered = applyFilters(state.tasks);
    const today = state.summary ? state.summary.today : fmtDate(new Date());

    const sorted = filtered.slice().sort((a, b) => {
      const field = state.sortField;
      const dir = state.sortDir === 'asc' ? 1 : -1;
      const va = (a[field] || '').toString();
      const vb = (b[field] || '').toString();
      return va < vb ? -1 * dir : va > vb ? 1 * dir : 0;
    });

    if (sorted.length === 0) {
      $('main-content').innerHTML = '<div class="list-view"><div class="empty">沒有符合條件的任務</div></div>';
      return;
    }

    const rows = sorted.map(t => {
      const due = getDueInfo(t.dueDay, today, t.status);
      const tagCls = getTagClass(t.type);
      const shortStatus = getShortStatus(t.status);
      const statusCls = STATUS_CLASS[shortStatus] || 'status-waiting';
      const priCls = getPriorityClass(t.priority);

      return `<tr>
        <td>${escapeHtml(t.week || '')}</td>
        <td><span class="priority-bar ${priCls}"></span>${escapeHtml(t.name)}</td>
        <td><span class="owner-chip"><span class="avatar">${escapeHtml(getInitials(t.owner))}</span>${escapeHtml(t.owner || '-')}</span></td>
        <td><span class="task-tag ${tagCls}">${escapeHtml(t.type || '-')}</span></td>
        <td><span class="status-pill ${statusCls}">${escapeHtml(shortStatus)}</span></td>
        <td><span class="due-pill ${due.cls}">${escapeHtml(t.dueDay || '-')}</span></td>
        <td style="font-size:11px; color:#9ca3af;">${escapeHtml(due.text)}</td>
      </tr>`;
    }).join('');

    const html = `<div class="list-view">
      <table class="list-table">
        <thead>
          <tr>
            <th class="sortable" data-sort="week">週次</th>
            <th class="sortable" data-sort="name">任務名稱</th>
            <th class="sortable" data-sort="owner">負責人</th>
            <th class="sortable" data-sort="type">類型</th>
            <th class="sortable" data-sort="status">狀態</th>
            <th class="sortable" data-sort="dueDay">Due Day</th>
            <th>剩餘</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>`;

    $('main-content').innerHTML = html;

    document.querySelectorAll('.list-table th.sortable').forEach(th => {
      th.addEventListener('click', () => {
        const field = th.dataset.sort;
        if (state.sortField === field) {
          state.sortDir = state.sortDir === 'asc' ? 'desc' : 'asc';
        } else {
          state.sortField = field;
          state.sortDir = 'asc';
        }
        renderList();
      });
    });
  }

  function render() {
    renderKPIs();
    if (state.view === 'board') renderBoard();
    else renderList();
  }

  async function loadAndRender() {
    const btn = $('refresh-btn');
    btn.classList.add('spinning');
    try {
      const data = await fetchData();
      state.tasks = data.tasks || [];
      state.summary = data.summary || null;
      const now = new Date();
      $('last-updated').textContent = '最後更新：' + now.toLocaleString('zh-TW', { hour12: false });
      renderFilters();
      render();
    } catch (err) {
      console.error(err);
      $('main-content').innerHTML = `<div class="error">載入失敗：${escapeHtml(err.message)}<br><br>請檢查 config.js 的 API_URL 和 API_TOKEN 是否正確。</div>`;
    } finally {
      btn.classList.remove('spinning');
    }
  }

  function bindEvents() {
    document.querySelectorAll('#view-toggle button').forEach(b => {
      b.addEventListener('click', () => {
        document.querySelectorAll('#view-toggle button').forEach(x => x.classList.remove('active'));
        b.classList.add('active');
        state.view = b.dataset.view;
        render();
      });
    });

    $('filter-owner').addEventListener('change', e => { state.filters.owner = e.target.value; render(); });
    $('filter-type').addEventListener('change', e => { state.filters.type = e.target.value; render(); });
    $('filter-priority').addEventListener('change', e => { state.filters.priority = e.target.value; render(); });
    $('filter-daterange').addEventListener('change', e => { state.filters.dueRange = e.target.value; render(); });
    $('filter-search').addEventListener('input', e => { state.filters.search = e.target.value; render(); });

    $('filter-reset').addEventListener('click', () => {
      state.filters = { owner: '', type: '', priority: '', search: '', dueRange: '' };
      $('filter-owner').value = '';
      $('filter-type').value = '';
      $('filter-priority').value = '';
      $('filter-daterange').value = '';
      $('filter-search').value = '';
      render();
    });

    $('refresh-btn').addEventListener('click', loadAndRender);

    if (window.CE_CONFIG && window.CE_CONFIG.SHEET_URL) {
      $('sheet-link').href = window.CE_CONFIG.SHEET_URL;
    }
  }

  initLock();
})();
