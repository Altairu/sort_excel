let currentConfig = null;
let currentRecords = [];
let editingRecordId = null;

// トースト表示
function showToast(message, type = 'info') {
  const toast = document.getElementById('toast');
  toast.textContent = message;
  toast.className = `toast show ${type}`;
  setTimeout(() => {
    toast.className = 'toast';
  }, 3500);
}

// 設定の読み込みとセレクトボックス初期化
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

  // 担当フィルター
  const staffFilter = document.getElementById('staffFilter');
  const currentVal = staffFilter.value;
  staffFilter.innerHTML = '<option value="">全員</option>';
  currentConfig.staff_order.forEach(s => {
    const opt = document.createElement('option');
    opt.value = s;
    opt.textContent = s;
    staffFilter.appendChild(opt);
  });
  staffFilter.value = currentVal;

  // フォーム用セレクトボックス
  populateSelect('field_staff', currentConfig.staff_order);
  populateSelect('field_form_type', currentConfig.form_type_options);
  populateSelect('field_request_content', currentConfig.request_content_options);
  populateSelect('field_same_day_application', currentConfig.same_day_options);
  populateSelect('field_status', currentConfig.status_options);
}

function populateSelect(id, options) {
  const el = document.getElementById(id);
  if (!el) return;
  el.innerHTML = '<option value="">(未選択)</option>';
  options.forEach(opt => {
    const o = document.createElement('option');
    o.value = opt;
    o.textContent = opt;
    el.appendChild(o);
  });
}

// 統計情報取得
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

// 案件一覧の取得と描画
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
    tr.innerHTML = `<td colspan="12" style="text-align:center; padding: 24px; color: var(--text-muted);">該当する案件がありません</td>`;
    tbody.appendChild(tr);
    return;
  }

  records.forEach(r => {
    const tr = document.createElement('tr');
    tr.dataset.id = r.id;

    // ステータスバッジ
    let badgeClass = 'badge-active';
    if (r.status === '完了') badgeClass = 'badge-done';
    if (r.status === '日程未定' || r.status === 'リスケ') badgeClass = 'badge-pending';

    tr.innerHTML = `
      <td>${r.id}</td>
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
        <button class="btn btn-sm btn-outline edit-row-btn" data-id="${r.id}">編集</button>
      </td>
    `;

    // 行クリックで編集パネルを開く
    tr.addEventListener('click', (e) => {
      openEditPanel(r.id);
    });

    tbody.appendChild(tr);
  });
}

function escapeHtml(str) {
  if (!str) return '';
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// 編集パネル開閉
function openEditPanel(recordId = null) {
  editingRecordId = recordId;
  const panel = document.getElementById('editPanel');
  const backdrop = document.getElementById('panelBackdrop');
  const title = document.getElementById('panelTitle');
  const form = document.getElementById('recordForm');

  form.reset();

  if (recordId) {
    title.textContent = `案件の編集 (ID: ${recordId})`;
    const rec = currentRecords.find(r => r.id === recordId);
    if (rec) {
      document.getElementById('field_id').value = rec.id;
      document.getElementById('field_version').value = rec.version;
      document.getElementById('field_status').value = rec.status;
      document.getElementById('field_staff').value = rec.staff;
      document.getElementById('field_acceptance_date').value = rec.acceptance_date;
      document.getElementById('field_form_type').value = rec.form_type;
      document.getElementById('field_client_name').value = rec.client_name;
      document.getElementById('field_property_name').value = rec.property_name;
      document.getElementById('field_request_content').value = rec.request_content;
      document.getElementById('field_jurisdiction').value = rec.jurisdiction;
      document.getElementById('field_loan_date').value = rec.loan_date;
      document.getElementById('field_loan_date_note').value = rec.loan_date_note;
      document.getElementById('field_attendance_time').value = rec.attendance_time;
      document.getElementById('field_attendee').value = rec.attendee;
      document.getElementById('field_attendance_place').value = rec.attendance_place;
      document.getElementById('field_same_day_application').value = rec.same_day_application;
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
    }
  } else {
    title.textContent = '新規案件の登録';
    document.getElementById('field_id').value = '';
    document.getElementById('field_version').value = '';
    document.getElementById('field_status').value = '進行中';
    document.getElementById('field_same_day_application').value = 'オンライン';
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
    status: document.getElementById('field_status').value,
    staff: document.getElementById('field_staff').value,
    acceptance_date: document.getElementById('field_acceptance_date').value,
    form_type: document.getElementById('field_form_type').value,
    client_name: document.getElementById('field_client_name').value,
    property_name: document.getElementById('field_property_name').value,
    request_content: document.getElementById('field_request_content').value,
    jurisdiction: document.getElementById('field_jurisdiction').value,
    loan_date: document.getElementById('field_loan_date').value,
    loan_date_note: document.getElementById('field_loan_date_note').value,
    attendance_time: document.getElementById('field_attendance_time').value,
    attendee: document.getElementById('field_attendee').value,
    attendance_place: document.getElementById('field_attendance_place').value,
    same_day_application: document.getElementById('field_same_day_application').value,
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
    version: version ? parseInt(version, 10) : undefined,
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
      showToast(id ? '案件を更新し、Excelを最新化しました' : '新規登録し、Excelを最新化しました', 'success');
      closeEditPanel();
      loadRecords();
      loadStats();
    } else {
      showToast(json.error || '保存に失敗しました', 'error');
    }
  } catch (err) {
    showToast('通信エラーが発生しました', 'error');
  }
});

// 手動Excel出力
document.getElementById('btnExport').addEventListener('click', async () => {
  showToast('Excelファイルを出力中...', 'info');
  try {
    const res = await fetch('/api/export', { method: 'POST' });
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

// 管理設定モーダル
document.getElementById('btnSettings').addEventListener('click', () => {
  openSettingsModal();
});
document.getElementById('btnCloseSettings').addEventListener('click', () => {
  document.getElementById('settingsModal').classList.remove('active');
});
document.getElementById('btnCancelSettings').addEventListener('click', () => {
  document.getElementById('settingsModal').classList.remove('active');
});

function openSettingsModal() {
  if (!currentConfig) return;

  renderStaffOrderList(currentConfig.staff_order);
  document.getElementById('settingFormTypes').value = currentConfig.form_type_options.join('\n');
  document.getElementById('settingRequestContents').value = currentConfig.request_content_options.join('\n');

  loadBackupList();
  document.getElementById('settingsModal').classList.add('active');
}

function renderStaffOrderList(staffList) {
  const ul = document.getElementById('staffOrderList');
  ul.innerHTML = '';
  staffList.forEach((s, idx) => {
    const li = document.createElement('li');
    li.className = 'reorder-item';
    li.innerHTML = `
      <span>${idx + 1}. ${escapeHtml(s)}</span>
      <div class="reorder-actions">
        <button class="btn btn-sm btn-outline move-up-btn" data-idx="${idx}">↑</button>
        <button class="btn btn-sm btn-outline move-down-btn" data-idx="${idx}">↓</button>
        <button class="btn btn-sm btn-outline delete-staff-btn" data-idx="${idx}">削除</button>
      </div>
    `;
    ul.appendChild(li);
  });

  // 上下移動イベント
  ul.querySelectorAll('.move-up-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const idx = parseInt(btn.dataset.idx, 10);
      if (idx > 0) {
        const tmp = staffList[idx];
        staffList[idx] = staffList[idx - 1];
        staffList[idx - 1] = tmp;
        renderStaffOrderList(staffList);
      }
    });
  });

  ul.querySelectorAll('.move-down-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const idx = parseInt(btn.dataset.idx, 10);
      if (idx < staffList.length - 1) {
        const tmp = staffList[idx];
        staffList[idx] = staffList[idx + 1];
        staffList[idx + 1] = tmp;
        renderStaffOrderList(staffList);
      }
    });
  });

  ul.querySelectorAll('.delete-staff-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const idx = parseInt(btn.dataset.idx, 10);
      staffList.splice(idx, 1);
      renderStaffOrderList(staffList);
    });
  });
}

document.getElementById('btnAddStaff').addEventListener('click', () => {
  const input = document.getElementById('newStaffInput');
  const val = input.value.trim();
  if (val && currentConfig) {
    currentConfig.staff_order.push(val);
    renderStaffOrderList(currentConfig.staff_order);
    input.value = '';
  }
});

// バックアップ一覧取得
async function loadBackupList() {
  try {
    const res = await fetch('/api/backups');
    const json = await res.json();
    const tbody = document.getElementById('backupListTbody');
    tbody.innerHTML = '';
    if (json.success && json.data.length > 0) {
      json.data.forEach(b => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td>${b.date_str}</td>
          <td>${b.created_at}</td>
          <td>${b.excel_file || '-'}</td>
          <td>${b.db_file || '-'}</td>
        `;
        tbody.appendChild(tr);
      });
    } else {
      tbody.innerHTML = '<tr><td colspan="4" style="text-align:center; color: var(--text-muted);">バックアップ履歴はまだありません</td></tr>';
    }
  } catch (err) {
    console.error('バックアップ一覧取得失敗:', err);
  }
}

// 設定保存
document.getElementById('btnSaveSettings').addEventListener('click', async () => {
  if (!currentConfig) return;

  currentConfig.form_type_options = document.getElementById('settingFormTypes').value
    .split('\n')
    .map(s => s.trim())
    .filter(Boolean);

  currentConfig.request_content_options = document.getElementById('settingRequestContents').value
    .split('\n')
    .map(s => s.trim())
    .filter(Boolean);

  try {
    const res = await fetch('/api/config', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(currentConfig),
    });
    const json = await res.json();
    if (json.success) {
      showToast('設定を保存し、Excelを再ソートしました', 'success');
      document.getElementById('settingsModal').classList.remove('active');
      initSelectOptions();
      loadRecords();
    } else {
      showToast(json.error || '設定の保存に失敗しました', 'error');
    }
  } catch (err) {
    showToast('通信エラーが発生しました', 'error');
  }
});

// インポート実行
document.getElementById('btnExecuteImport').addEventListener('click', async () => {
  const filePath = document.getElementById('importFilePath').value.trim();
  const resDiv = document.getElementById('importResult');
  resDiv.innerHTML = '<span style="color:var(--text-muted)">取り込みを実行中...</span>';

  try {
    const res = await fetch('/api/import', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ file_path: filePath }),
    });
    const json = await res.json();
    if (json.success) {
      resDiv.innerHTML = `<span style="color:var(--secondary)">${json.data}</span>`;
      showToast(json.data, 'success');
      loadRecords();
      loadStats();
    } else {
      resDiv.innerHTML = `<span style="color:var(--danger)">エラー: ${json.error}</span>`;
    }
  } catch (err) {
    resDiv.innerHTML = `<span style="color:var(--danger)">通信エラー</span>`;
  }
});

// 設定タブ切り替え
document.querySelectorAll('.settings-tabs .tab-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.settings-tabs .tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById(btn.dataset.tab).classList.add('active');
  });
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
window.addEventListener('DOMContentLoaded', async () => {
  await loadConfig();
  await loadStats();
  await loadRecords();
});
