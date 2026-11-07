let currentConfig = null;
let currentRecords = [];
let editingRecordId = null;

// ソート済みビュー用データ保持
let sortedViewsData = null;
let activeViewTab = 'view-master';
let selectedDateKey = null;

const CUSTOM_OPTION_VALUE = '__CUSTOM__';
const ADD_ACTOR_OPTION_VALUE = '__ADD_ACTOR__';
const TEST_ACTOR_NAME = 'テスト用（動作確認）';

const SELECT_WITH_CUSTOM_FIELDS = [
  'field_status',
  'field_staff',
  'field_form_type',
  'field_request_content',
  'field_same_day_application',
];

function showToast(message, type = 'info') {
  const toast = document.getElementById('toast');
  toast.textContent = message;
  toast.className = `toast show ${type}`;
  setTimeout(() => {
    toast.className = 'toast';
  }, 3500);
}

async function checkSessionInfo() {
  try {
    const savedActor = localStorage.getItem('sort_excel_actor') || '';
    const params = new URLSearchParams({ log_access: 'true' });
    if (savedActor) params.set('actor', savedActor);

    const res = await fetch(`/api/session?${params.toString()}`);
    const json = await res.json();
    if (json.success && json.data && json.data.is_admin_pc) {
      const adminLink = document.getElementById('btnAdminLink');
      if (adminLink) {
        adminLink.style.display = 'inline-flex';
      }
    }
  } catch (err) {
    console.error('セッション確認失敗:', err);
  }
}

function getCurrentActorName() {
  return localStorage.getItem('sort_excel_actor') || '';
}

async function setCurrentActorName(name, logAccess = false) {
  const trimmed = (name || '').trim();
  if (!trimmed) return;
  localStorage.setItem('sort_excel_actor', trimmed);

  const actorSelect = document.getElementById('currentActorSelect');
  if (actorSelect) {
    actorSelect.value = trimmed;
  }

  document.getElementById('actorSetupModal').classList.remove('active');

  if (logAccess) {
    try {
      const params = new URLSearchParams({ log_access: 'true', actor: trimmed });
      await fetch(`/api/session?${params.toString()}`);
    } catch (e) {
      console.error(e);
    }
  }
}

function openActorSetupModal() {
  const grid = document.getElementById('actorButtonsGrid');
  if (grid && currentConfig) {
    grid.innerHTML = '';
    currentConfig.staff_order.forEach(s => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'btn btn-outline';
      btn.style.padding = '8px 16px';
      btn.style.fontSize = '14px';
      btn.textContent = s;
      btn.addEventListener('click', async () => {
        await setCurrentActorName(s, true);
        showToast(`操作者を ${s} に設定しました`, 'success');
      });
      grid.appendChild(btn);
    });
  }
  document.getElementById('newActorNameInput').value = '';
  document.getElementById('actorSetupModal').classList.add('active');
}

async function addAndSelectNewActor() {
  const input = document.getElementById('newActorNameInput');
  const name = (input.value || '').trim();
  if (!name) {
    showToast('名前を入力してください', 'error');
    input.focus();
    return;
  }

  try {
    const res = await fetch('/api/actors', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name }),
    });
    const json = await res.json();
    if (json.success && json.data) {
      currentConfig = json.data;
      initSelectOptions();
      await setCurrentActorName(name, true);
      showToast(`${name} をリストに追加し、操作者に設定しました`, 'success');
    } else {
      showToast(json.error || '追加に失敗しました', 'error');
    }
  } catch (err) {
    showToast('通信エラーが発生しました', 'error');
  }
}

async function loadConfig() {
  try {
    const res = await fetch('/api/config');
    const json = await res.json();
    if (json.success) {
      currentConfig = json.data;
      initSelectOptions();
    }
  } catch (err) {
    console.error('設定取得失敗:', err);
  }
}

function initSelectOptions() {
  if (!currentConfig) return;

  // 操作者セレクター（localStorageに保存 + テスト用アカウント + リストになければ追加可能）
  const actorSelect = document.getElementById('currentActorSelect');
  if (actorSelect) {
    let savedActor = localStorage.getItem('sort_excel_actor') || '';
    // 以前テスト入力された Altair がブラウザに残っていた場合は自動的にテスト用へ切り替え
    if (savedActor.toLowerCase() === 'altair') {
      savedActor = TEST_ACTOR_NAME;
      localStorage.setItem('sort_excel_actor', savedActor);
    }

    actorSelect.innerHTML = '<option value="">(未設定 - クリックして選択)</option>';
    currentConfig.staff_order.forEach(s => {
      const opt = document.createElement('option');
      opt.value = s;
      opt.textContent = s;
      actorSelect.appendChild(opt);
    });

    const testOpt = document.createElement('option');
    testOpt.value = TEST_ACTOR_NAME;
    testOpt.textContent = `[テスト] ${TEST_ACTOR_NAME}`;
    actorSelect.appendChild(testOpt);

    if (savedActor && savedActor !== TEST_ACTOR_NAME && !currentConfig.staff_order.includes(savedActor)) {
      const opt = document.createElement('option');
      opt.value = savedActor;
      opt.textContent = savedActor;
      actorSelect.appendChild(opt);
    }
    const addOpt = document.createElement('option');
    addOpt.value = ADD_ACTOR_OPTION_VALUE;
    addOpt.textContent = '+ リストにない名前を追加...';
    actorSelect.appendChild(addOpt);

    actorSelect.value = savedActor;
    if (!actorSelect.dataset.listenerAttached) {
      actorSelect.addEventListener('change', () => {
        if (actorSelect.value === ADD_ACTOR_OPTION_VALUE || actorSelect.value === '') {
          actorSelect.value = localStorage.getItem('sort_excel_actor') || '';
          openActorSetupModal();
        } else {
          setCurrentActorName(actorSelect.value, false);
          showToast(`操作者を ${actorSelect.value} に切り替えました`, 'info');
        }
      });
      actorSelect.dataset.listenerAttached = 'true';
    }
  }

  const staffFilter = document.getElementById('staffFilter');
  const currentStaffVal = staffFilter.value;
  staffFilter.innerHTML = '<option value="">全員</option>';
  currentConfig.staff_order.forEach(s => {
    const opt = document.createElement('option');
    opt.value = s;
    opt.textContent = s;
    staffFilter.appendChild(opt);
  });
  staffFilter.value = currentStaffVal;

  const statusFilter = document.getElementById('statusFilter');
  const currentStatusVal = statusFilter.value;
  statusFilter.innerHTML = '<option value="">すべて</option>';
  currentConfig.status_options.forEach(st => {
    const opt = document.createElement('option');
    opt.value = st;
    opt.textContent = st;
    statusFilter.appendChild(opt);
  });
  statusFilter.value = currentStatusVal;

  populateSelectWithCustom('field_staff', currentConfig.staff_order, true);
  populateSelectWithCustom('field_form_type', currentConfig.form_type_options, true);
  populateSelectWithCustom('field_request_content', currentConfig.request_content_options, true);
  populateSelectWithCustom('field_same_day_application', currentConfig.same_day_options, true);
  populateSelectWithCustom('field_status', currentConfig.status_options, false);
}

function populateSelectWithCustom(id, options, allowEmpty = true) {
  const el = document.getElementById(id);
  const customInput = document.getElementById(`${id}_custom`);
  if (!el) return;

  el.innerHTML = '';
  if (allowEmpty) {
    const emptyOpt = document.createElement('option');
    emptyOpt.value = '';
    emptyOpt.textContent = '(未選択)';
    el.appendChild(emptyOpt);
  }

  options.forEach(opt => {
    const o = document.createElement('option');
    o.value = opt;
    o.textContent = opt;
    el.appendChild(o);
  });

  const customOpt = document.createElement('option');
  customOpt.value = CUSTOM_OPTION_VALUE;
  customOpt.textContent = 'その他（自由入力...）';
  el.appendChild(customOpt);

  if (!el.dataset.listenerAttached) {
    el.addEventListener('change', () => {
      if (!customInput) return;
      if (el.value === CUSTOM_OPTION_VALUE) {
        customInput.style.display = 'block';
        customInput.focus();
      } else {
        customInput.style.display = 'none';
        customInput.value = '';
      }
    });
    el.dataset.listenerAttached = 'true';
  }
}

function setSelectOrCustomValue(id, value) {
  const el = document.getElementById(id);
  const customInput = document.getElementById(`${id}_custom`);
  if (!el) return;

  const val = value || '';
  const hasOption = Array.from(el.options).some(
    opt => opt.value === val && opt.value !== CUSTOM_OPTION_VALUE
  );

  if (hasOption) {
    el.value = val;
    if (customInput) {
      customInput.style.display = 'none';
      customInput.value = '';
    }
  } else if (val !== '') {
    el.value = CUSTOM_OPTION_VALUE;
    if (customInput) {
      customInput.style.display = 'block';
      customInput.value = val;
    }
  } else {
    el.value = '';
    if (customInput) {
      customInput.style.display = 'none';
      customInput.value = '';
    }
  }
}

function getSelectOrCustomValue(id) {
  const el = document.getElementById(id);
  const customInput = document.getElementById(`${id}_custom`);
  if (!el) return '';
  if (el.value === CUSTOM_OPTION_VALUE) {
    return customInput ? customInput.value.trim() : '';
  }
  return el.value;
}

async function loadStats() {
  try {
    const res = await fetch('/api/stats');
    const json = await res.json();
    if (json.success) {
      const data = json.data;
      document.getElementById('statTotal').textContent = data.total_count;
      document.getElementById('statInProgress').textContent = data.in_progress_count;
      document.getElementById('statPending').textContent = data.pending_count;
      document.getElementById('statCompleted').textContent = data.completed_count;
    }
  } catch (err) {
    console.error('統計取得失敗:', err);
  }
}

async function loadRecords() {
  const q = document.getElementById('searchInput').value;
  const staff = document.getElementById('staffFilter').value;
  const status = document.getElementById('statusFilter').value;

  const params = new URLSearchParams();
  if (q) params.set('q', q);
  if (staff) params.set('staff', staff);
  if (status) params.set('status', status);

  try {
    const res = await fetch(`/api/records?${params.toString()}`);
    const json = await res.json();
    if (json.success) {
      currentRecords = json.data;
      renderTable(currentRecords);
    }
  } catch (err) {
    showToast('案件一覧の取得に失敗しました', 'error');
  }
}

function renderTable(records) {
  const tbody = document.getElementById('recordsTbody');
  tbody.innerHTML = '';
  document.getElementById('recordCountLabel').textContent = `表示件数: ${records.length}件`;

  if (records.length === 0) {
    const tr = document.createElement('tr');
    const td = document.createElement('td');
    td.colSpan = 12;
    td.style.textAlign = 'center';
    td.style.padding = '24px';
    td.style.color = 'var(--text-muted)';
    td.textContent = '該当する案件がありません';
    tr.appendChild(td);
    tbody.appendChild(tr);
    return;
  }

  records.forEach(r => {
    const tr = document.createElement('tr');
    tr.dataset.id = r.id;

    let badgeClass = 'badge-active';
    if (r.status === '完了') badgeClass = 'badge-done';
    if (r.status === '日程未定' || r.status === 'リスケ') badgeClass = 'badge-pending';

    tr.innerHTML = `
      <td>${escapeHtml(String(r.id))}</td>
      <td><span class="stat-badge ${badgeClass}">${escapeHtml(r.status)}</span></td>
      <td>${escapeHtml(r.loan_date || r.loan_date_note || '-')}</td>
      <td>${escapeHtml(r.settlement_date || r.settlement_date_note || '-')}</td>
      <td><b>${escapeHtml(r.client_name)}</b></td>
      <td>${escapeHtml(r.property_name)}</td>
      <td>${escapeHtml(r.form_type)}</td>
      <td>${escapeHtml(r.request_content)}</td>
      <td>${escapeHtml(r.staff)}</td>
      <td>${escapeHtml(r.attendance_time || '-')}</td>
      <td>${escapeHtml(r.same_day_application || '-')}</td>
      <td>
        <button class="btn btn-sm btn-outline edit-row-btn" data-id="${escapeHtml(String(r.id))}">編集</button>
      </td>
    `;

    tr.addEventListener('click', () => {
      openEditPanel(r.id);
    });

    tbody.appendChild(tr);
  });
}

// ==========================================================
// ソート済みビュー（日付別スケジュール / 融資実行日順 / 金消日順 / 統合データ）
// ==========================================================
async function loadSortedViews() {
  try {
    const res = await fetch('/api/sorted-views');
    const json = await res.json();
    if (json.success && json.data) {
      sortedViewsData = json.data;
      renderActiveSortedView();
    }
  } catch (err) {
    console.error('ソート済みデータの取得失敗:', err);
  }
}

function formatDateWithWeekday(dateStr) {
  if (!dateStr) return '-';
  const parts = dateStr.split('-');
  if (parts.length !== 3) return dateStr;
  const d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
  if (isNaN(d.getTime())) return dateStr;
  const weekdays = ['日', '月', '火', '水', '木', '金', '土'];
  return `${dateStr} (${weekdays[d.getDay()]})`;
}

function getAllSortedDates() {
  if (!sortedViewsData || !sortedViewsData.date_groups) return [];
  return Object.keys(sortedViewsData.date_groups).sort();
}

function findNearestDateKey(dates) {
  if (!dates || dates.length === 0) return null;
  const now = new Date();
  const yyyy = now.getFullYear();
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const dd = String(now.getDate()).padStart(2, '0');
  const todayStr = `${yyyy}-${mm}-${dd}`;

  // 今日以降で最も近い日付を探す。なければ最後の日付
  for (const d of dates) {
    if (d >= todayStr) return d;
  }
  return dates[dates.length - 1];
}

function renderActiveSortedView() {
  if (!sortedViewsData) return;

  if (activeViewTab === 'view-daily') {
    renderDateSidebar();
    renderDailyScheduleTable();
  } else if (activeViewTab === 'view-loan') {
    renderSubsheetTable('sorted_by_融資実行日（融資実行日順シート）', sortedViewsData.loan_rows || []);
  } else if (activeViewTab === 'view-cancel') {
    renderSubsheetTable('sorted_by_金消日・面談日（金消日 / 面談日順シート）', sortedViewsData.cancel_rows || []);
  } else if (activeViewTab === 'view-combined') {
    renderSubsheetTable('統合データシート（全日程統合）', sortedViewsData.combined_rows || []);
  }
}

function renderDateSidebar() {
  const ul = document.getElementById('dateListUl');
  const filterText = (document.getElementById('dateSearchInput').value || '').trim();
  ul.innerHTML = '';

  const allDates = getAllSortedDates();
  if (!selectedDateKey || !sortedViewsData.date_groups[selectedDateKey]) {
    selectedDateKey = findNearestDateKey(allDates);
  }

  const filteredDates = allDates.filter(d => !filterText || d.includes(filterText));

  if (filteredDates.length === 0) {
    const li = document.createElement('li');
    li.style.padding = '16px';
    li.style.textAlign = 'center';
    li.style.color = 'var(--text-muted)';
    li.textContent = '該当する日付がありません';
    ul.appendChild(li);
    return;
  }

  filteredDates.forEach(dateKey => {
    const group = sortedViewsData.date_groups[dateKey] || [];
    const loanCount = group.filter(r => r.category === '融資実行日').length;
    const cancelCount = group.filter(r => r.category !== '融資実行日').length;

    const li = document.createElement('li');
    li.className = `date-item ${dateKey === selectedDateKey ? 'active' : ''}`;
    li.dataset.date = dateKey;

    li.innerHTML = `
      <span>${escapeHtml(formatDateWithWeekday(dateKey))}</span>
      <div class="date-counts">
        ${loanCount > 0 ? `<span class="cat-badge-loan" title="融資実行日">実 ${loanCount}</span>` : ''}
        ${cancelCount > 0 ? `<span class="cat-badge-cancel" title="金消日・面談日">金 ${cancelCount}</span>` : ''}
      </div>
    `;

    li.addEventListener('click', () => {
      selectedDateKey = dateKey;
      renderDateSidebar();
      renderDailyScheduleTable();
    });

    ul.appendChild(li);
  });
}

function renderDailyScheduleTable() {
  const titleEl = document.getElementById('selectedDateTitle');
  const summaryEl = document.getElementById('selectedDateSummary');
  const tbody = document.getElementById('dailyScheduleTbody');
  tbody.innerHTML = '';

  if (!selectedDateKey || !sortedViewsData || !sortedViewsData.date_groups[selectedDateKey]) {
    titleEl.textContent = '日付データがありません';
    summaryEl.textContent = '';
    const tr = document.createElement('tr');
    const td = document.createElement('td');
    td.colSpan = 12;
    td.style.textAlign = 'center';
    td.style.padding = '24px';
    td.style.color = 'var(--text-muted)';
    td.textContent = '表示するスケジュールがありません';
    tr.appendChild(td);
    tbody.appendChild(tr);
    return;
  }

  const rows = sortedViewsData.date_groups[selectedDateKey] || [];
  const loanCount = rows.filter(r => r.category === '融資実行日').length;
  const cancelCount = rows.filter(r => r.category !== '融資実行日').length;

  titleEl.textContent = `${formatDateWithWeekday(selectedDateKey)} のスケジュール`;
  summaryEl.textContent = `合計 ${rows.length}件 （融資実行日: ${loanCount}件 / 金消日・面談日: ${cancelCount}件） - 行をクリックすると編集できます`;

  rows.forEach(item => {
    const tr = document.createElement('tr');
    const isLoan = item.category === '融資実行日';
    const catBadge = isLoan
      ? `<span class="cat-badge-loan">融資実行日</span>`
      : `<span class="cat-badge-cancel">金消日・面談日</span>`;

    tr.innerHTML = `
      <td>${catBadge}</td>
      <td><b>${escapeHtml(item.staff || '-')}</b></td>
      <td><b>${escapeHtml(item.time || '-')}</b></td>
      <td><b>${escapeHtml(item.client_name)}</b></td>
      <td>${escapeHtml(item.property_name)}</td>
      <td>${escapeHtml(item.form_type)}</td>
      <td>${escapeHtml(item.request_content)}</td>
      <td>${escapeHtml(item.jurisdiction || '-')}</td>
      <td>${escapeHtml(item.place || '-')}</td>
      <td>${escapeHtml(item.confirmer || '-')}</td>
      <td>${escapeHtml(item.application || '-')}</td>
      <td>
        <button class="btn btn-sm btn-outline">編集</button>
      </td>
    `;

    tr.addEventListener('click', () => {
      openEditPanel(item.record_id);
    });

    tbody.appendChild(tr);
  });
}

function renderSubsheetTable(title, rows) {
  document.getElementById('subsheetTitleLabel').textContent = title;
  const q = (document.getElementById('subsheetSearchInput').value || '').trim().toLowerCase();
  const tbody = document.getElementById('subsheetTbody');
  tbody.innerHTML = '';

  const filtered = rows.filter(r => {
    if (!q) return true;
    return (
      (r.client_name && r.client_name.toLowerCase().includes(q)) ||
      (r.property_name && r.property_name.toLowerCase().includes(q)) ||
      (r.staff && r.staff.toLowerCase().includes(q)) ||
      (r.date && r.date.toLowerCase().includes(q)) ||
      (r.place && r.place.toLowerCase().includes(q)) ||
      (r.jurisdiction && r.jurisdiction.toLowerCase().includes(q))
    );
  });

  document.getElementById('subsheetCountLabel').textContent = `表示件数: ${filtered.length}件`;

  if (filtered.length === 0) {
    const tr = document.createElement('tr');
    const td = document.createElement('td');
    td.colSpan = 13;
    td.style.textAlign = 'center';
    td.style.padding = '24px';
    td.style.color = 'var(--text-muted)';
    td.textContent = '該当するデータがありません';
    tr.appendChild(td);
    tbody.appendChild(tr);
    return;
  }

  filtered.forEach(item => {
    const tr = document.createElement('tr');
    const isLoan = item.category === '融資実行日';
    const catBadge = isLoan
      ? `<span class="cat-badge-loan">融資実行日</span>`
      : `<span class="cat-badge-cancel">金消日・面談日</span>`;

    tr.innerHTML = `
      <td><b>${escapeHtml(formatDateWithWeekday(item.date))}</b></td>
      <td>${catBadge}</td>
      <td><b>${escapeHtml(item.staff || '-')}</b></td>
      <td>${escapeHtml(item.time || '-')}</td>
      <td><b>${escapeHtml(item.client_name)}</b></td>
      <td>${escapeHtml(item.property_name)}</td>
      <td>${escapeHtml(item.form_type)}</td>
      <td>${escapeHtml(item.request_content)}</td>
      <td>${escapeHtml(item.jurisdiction || '-')}</td>
      <td>${escapeHtml(item.place || '-')}</td>
      <td>${escapeHtml(item.confirmer || '-')}</td>
      <td>${escapeHtml(item.application || '-')}</td>
      <td>
        <button class="btn btn-sm btn-outline">編集</button>
      </td>
    `;

    tr.addEventListener('click', () => {
      openEditPanel(item.record_id);
    });

    tbody.appendChild(tr);
  });
}

// ビュータブの切り替えイベント
document.querySelectorAll('.view-tab-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.view-tab-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');

    activeViewTab = btn.dataset.view;
    document.querySelectorAll('.view-pane').forEach(p => p.classList.remove('active'));

    if (activeViewTab === 'view-master') {
      document.getElementById('view-master').classList.add('active');
    } else if (activeViewTab === 'view-daily') {
      document.getElementById('view-daily').classList.add('active');
      renderActiveSortedView();
    } else {
      document.getElementById('view-subsheet').classList.add('active');
      renderActiveSortedView();
    }
  });
});

// 日付検索と前後移動ボタン
document.getElementById('dateSearchInput').addEventListener('input', () => {
  renderDateSidebar();
});

document.getElementById('btnJumpNearestDate').addEventListener('click', () => {
  const allDates = getAllSortedDates();
  selectedDateKey = findNearestDateKey(allDates);
  document.getElementById('dateSearchInput').value = '';
  renderDateSidebar();
  renderDailyScheduleTable();
  const activeEl = document.querySelector('#dateListUl .date-item.active');
  if (activeEl) activeEl.scrollIntoView({ block: 'center', behavior: 'smooth' });
});

document.getElementById('btnPrevDate').addEventListener('click', () => {
  const allDates = getAllSortedDates();
  const idx = allDates.indexOf(selectedDateKey);
  if (idx > 0) {
    selectedDateKey = allDates[idx - 1];
    renderDateSidebar();
    renderDailyScheduleTable();
  }
});

document.getElementById('btnNextDate').addEventListener('click', () => {
  const allDates = getAllSortedDates();
  const idx = allDates.indexOf(selectedDateKey);
  if (idx >= 0 && idx < allDates.length - 1) {
    selectedDateKey = allDates[idx + 1];
    renderDateSidebar();
    renderDailyScheduleTable();
  }
});

document.getElementById('subsheetSearchInput').addEventListener('input', () => {
  renderActiveSortedView();
});

function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// 編集パネル開閉（ソート済み画面から開いた場合も単一取得APIで確実に開けるよう対応）
async function openEditPanel(recordId = null) {
  editingRecordId = recordId;
  const panel = document.getElementById('editPanel');
  const backdrop = document.getElementById('panelBackdrop');
  const title = document.getElementById('panelTitle');
  const form = document.getElementById('recordForm');
  const btnDelete = document.getElementById('btnDeleteRecord');

  form.reset();
  SELECT_WITH_CUSTOM_FIELDS.forEach(id => {
    const customInput = document.getElementById(`${id}_custom`);
    if (customInput) {
      customInput.style.display = 'none';
      customInput.value = '';
    }
  });

  if (recordId) {
    title.textContent = `案件の編集 (ID: ${recordId})`;
    if (btnDelete) btnDelete.style.display = 'inline-flex';
    let rec = currentRecords.find(r => r.id === recordId);
    if (!rec) {
      try {
        const res = await fetch(`/api/records/${recordId}`);
        const json = await res.json();
        if (json.success) rec = json.data;
      } catch (e) {
        console.error(e);
      }
    }
    if (rec) {
      document.getElementById('field_id').value = rec.id;
      document.getElementById('field_version').value = rec.version;
      setSelectOrCustomValue('field_status', rec.status);
      setSelectOrCustomValue('field_staff', rec.staff);
      document.getElementById('field_acceptance_date').value = rec.acceptance_date;
      setSelectOrCustomValue('field_form_type', rec.form_type);
      document.getElementById('field_client_name').value = rec.client_name;
      document.getElementById('field_property_name').value = rec.property_name;
      setSelectOrCustomValue('field_request_content', rec.request_content);
      document.getElementById('field_jurisdiction').value = rec.jurisdiction;
      document.getElementById('field_loan_date').value = rec.loan_date;
      document.getElementById('field_loan_date_note').value = rec.loan_date_note;
      document.getElementById('field_attendance_time').value = rec.attendance_time;
      document.getElementById('field_attendee').value = rec.attendee;
      document.getElementById('field_attendance_place').value = rec.attendance_place;
      setSelectOrCustomValue('field_same_day_application', rec.same_day_application);
      document.getElementById('field_execution_time').value = rec.execution_time;
      document.getElementById('field_settlement_date').value = rec.settlement_date;
      document.getElementById('field_settlement_time').value = rec.settlement_time;
      document.getElementById('field_settlement_date_note').value = rec.settlement_date_note;
      document.getElementById('field_settlement_place').value = rec.settlement_place;
      document.getElementById('field_confirmation').value = rec.confirmation;
      document.getElementById('field_doc_check').value = rec.doc_check;
      document.getElementById('field_broker').value = rec.broker;
      document.getElementById('field_agency').value = rec.agency;
      document.getElementById('field_opponent_scrivener').value = rec.opponent_scrivener;
      document.getElementById('field_remarks').value = rec.remarks;
      document.getElementById('field_progress_application').value = rec.progress_application || '';
      document.getElementById('field_progress_required_docs').value = rec.progress_required_docs || '';
      document.getElementById('field_progress_transfer_docs').value = rec.progress_transfer_docs || '';
      document.getElementById('field_progress_complete_docs').value = rec.progress_complete_docs || '';
      document.getElementById('field_progress_receive_date').value = rec.progress_receive_date || '';
      document.getElementById('field_progress_tokyo_complete').value = rec.progress_tokyo_complete || '';
      document.getElementById('field_progress_payment_date').value = rec.progress_payment_date || '';
    }
  } else {
    title.textContent = '新規案件の登録';
    if (btnDelete) btnDelete.style.display = 'none';
    document.getElementById('field_id').value = '';
    document.getElementById('field_version').value = '';
    setSelectOrCustomValue('field_status', '進行中');
    setSelectOrCustomValue('field_same_day_application', 'オンライン');
  }

  backdrop.classList.add('active');
  panel.classList.add('active');
}

function closeEditPanel() {
  document.getElementById('panelBackdrop').classList.remove('active');
  document.getElementById('editPanel').classList.remove('active');
  editingRecordId = null;
}

// フォーム送信（保存）
document.getElementById('recordForm').addEventListener('submit', async (e) => {
  e.preventDefault();

  const id = document.getElementById('field_id').value;
  const version = document.getElementById('field_version').value;

  const data = {
    status: getSelectOrCustomValue('field_status') || '進行中',
    staff: getSelectOrCustomValue('field_staff'),
    acceptance_date: document.getElementById('field_acceptance_date').value,
    form_type: getSelectOrCustomValue('field_form_type'),
    client_name: document.getElementById('field_client_name').value,
    property_name: document.getElementById('field_property_name').value,
    request_content: getSelectOrCustomValue('field_request_content'),
    jurisdiction: document.getElementById('field_jurisdiction').value,
    loan_date: document.getElementById('field_loan_date').value,
    loan_date_note: document.getElementById('field_loan_date_note').value,
    attendance_time: document.getElementById('field_attendance_time').value,
    attendee: document.getElementById('field_attendee').value,
    attendance_place: document.getElementById('field_attendance_place').value,
    same_day_application: getSelectOrCustomValue('field_same_day_application'),
    execution_time: document.getElementById('field_execution_time').value,
    settlement_date: document.getElementById('field_settlement_date').value,
    settlement_time: document.getElementById('field_settlement_time').value,
    settlement_date_note: document.getElementById('field_settlement_date_note').value,
    settlement_place: document.getElementById('field_settlement_place').value,
    confirmation: document.getElementById('field_confirmation').value,
    doc_check: document.getElementById('field_doc_check').value,
    broker: document.getElementById('field_broker').value,
    agency: document.getElementById('field_agency').value,
    opponent_scrivener: document.getElementById('field_opponent_scrivener').value,
    remarks: document.getElementById('field_remarks').value,
    progress_application: document.getElementById('field_progress_application').value,
    progress_required_docs: document.getElementById('field_progress_required_docs').value,
    progress_transfer_docs: document.getElementById('field_progress_transfer_docs').value,
    progress_complete_docs: document.getElementById('field_progress_complete_docs').value,
    progress_receive_date: document.getElementById('field_progress_receive_date').value,
    progress_tokyo_complete: document.getElementById('field_progress_tokyo_complete').value,
    progress_payment_date: document.getElementById('field_progress_payment_date').value,
    version: version ? parseInt(version, 10) : undefined,
    actor_name: getCurrentActorName(),
  };

  try {
    let res;
    if (id) {
      res = await fetch(`/api/records/${id}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });
    } else {
      res = await fetch('/api/records', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });
    }

    const json = await res.json();
    if (json.success) {
      showToast(id ? '案件を更新し、ソート結果とExcelを最新化しました' : '新規登録し、ソート結果とExcelを最新化しました', 'success');
      closeEditPanel();
      await Promise.all([loadRecords(), loadStats(), loadSortedViews()]);
    } else {
      showToast(json.error || '保存に失敗しました', 'error');
    }
  } catch (err) {
    showToast('通信エラーが発生しました', 'error');
  }
});

// 案件削除（ゴミ箱へ移動）
document.getElementById('btnDeleteRecord').addEventListener('click', async () => {
  const id = document.getElementById('field_id').value;
  if (!id) return;
  if (!confirm(`案件 ID: ${id} を削除しますか？（管理者のゴミ箱から復元可能です）`)) {
    return;
  }

  try {
    const actor = encodeURIComponent(getCurrentActorName());
    const res = await fetch(`/api/records/${id}?actor=${actor}`, { method: 'DELETE' });
    const json = await res.json();
    if (json.success) {
      showToast('案件を削除しました', 'success');
      closeEditPanel();
      await Promise.all([loadRecords(), loadStats(), loadSortedViews()]);
    } else {
      showToast(json.error || '削除に失敗しました', 'error');
    }
  } catch (err) {
    showToast('通信エラーが発生しました', 'error');
  }
});

// 手動Excel出力
document.getElementById('btnExport').addEventListener('click', async () => {
  showToast('Excelファイルを出力中...', 'info');
  try {
    const actor = encodeURIComponent(getCurrentActorName());
    const res = await fetch(`/api/export?actor=${actor}`, { method: 'POST' });
    const json = await res.json();
    if (json.success) {
      showToast(json.data || 'Excelを出力しました', 'success');
    } else {
      showToast(json.error || 'Excel出力に失敗しました', 'error');
    }
  } catch (err) {
    showToast('通信エラーが発生しました', 'error');
  }
});

// 使い方ガイドモーダル
document.getElementById('btnHelp').addEventListener('click', () => {
  document.getElementById('helpModal').classList.add('active');
});
document.getElementById('btnCloseHelp').addEventListener('click', () => {
  document.getElementById('helpModal').classList.remove('active');
});
document.getElementById('btnOkHelp').addEventListener('click', () => {
  document.getElementById('helpModal').classList.remove('active');
});

// 全画面表示トグル
const btnFullscreen = document.getElementById('btnFullscreen');
function updateFullscreenUI() {
  const isFs = !!document.fullscreenElement || document.body.classList.contains('wide-fullscreen');
  if (btnFullscreen) {
    btnFullscreen.textContent = isFs ? '全画面を解除' : '全画面表示';
  }
}

if (btnFullscreen) {
  btnFullscreen.addEventListener('click', async () => {
    try {
      if (!document.fullscreenElement) {
        await document.documentElement.requestFullscreen();
        document.body.classList.add('wide-fullscreen');
      } else {
        await document.exitFullscreen();
        document.body.classList.remove('wide-fullscreen');
      }
    } catch (e) {
      // フルスクリーンAPI非対応環境向けフォールバック（余白最小化ワイドモード）
      document.body.classList.toggle('wide-fullscreen');
    }
    updateFullscreenUI();
  });
}

document.addEventListener('fullscreenchange', () => {
  if (!document.fullscreenElement) {
    document.body.classList.remove('wide-fullscreen');
  } else {
    document.body.classList.add('wide-fullscreen');
  }
  updateFullscreenUI();
});

// 検索とフィルターのイベント
let debounceTimer;
document.getElementById('searchInput').addEventListener('input', () => {
  clearTimeout(debounceTimer);
  debounceTimer = setTimeout(loadRecords, 200);
});
document.getElementById('staffFilter').addEventListener('change', loadRecords);
document.getElementById('statusFilter').addEventListener('change', loadRecords);
document.getElementById('btnResetFilter').addEventListener('click', () => {
  document.getElementById('searchInput').value = '';
  document.getElementById('staffFilter').value = '';
  document.getElementById('statusFilter').value = '';
  loadRecords();
});

// 新規作成ボタン
document.getElementById('btnNewRecord').addEventListener('click', () => openEditPanel(null));

// パネル閉じる
document.getElementById('btnClosePanel').addEventListener('click', closeEditPanel);
document.getElementById('btnCancelEdit').addEventListener('click', closeEditPanel);
document.getElementById('panelBackdrop').addEventListener('click', closeEditPanel);

// 初期化実行
document.getElementById('btnSelectTestActor').addEventListener('click', async () => {
  await setCurrentActorName(TEST_ACTOR_NAME, true);
  showToast(`操作者を ${TEST_ACTOR_NAME} に設定しました`, 'info');
});
document.getElementById('btnAddAndSelectActor').addEventListener('click', addAndSelectNewActor);
document.getElementById('newActorNameInput').addEventListener('keydown', (e) => {
  if (e.key === 'Enter') {
    e.preventDefault();
    addAndSelectNewActor();
  }
});

window.addEventListener('DOMContentLoaded', async () => {
  await loadConfig();
  await checkSessionInfo();
  await Promise.all([loadStats(), loadRecords(), loadSortedViews()]);

  // 操作者が未設定の場合は最初に選択・追加を促すモーダルを表示
  if (!getCurrentActorName()) {
    openActorSetupModal();
  }
});
