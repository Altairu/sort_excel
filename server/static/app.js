let currentConfig = null;
let currentRecords = [];
let editingRecordId = null;

const CUSTOM_OPTION_VALUE = '__CUSTOM__';

// コンボボックス化対象のフィールド一覧
const SELECT_WITH_CUSTOM_FIELDS = [
  'field_status',
  'field_staff',
  'field_form_type',
  'field_request_content',
  'field_same_day_application',
];

// トースト表示
function showToast(message, type = 'info') {
  const toast = document.getElementById('toast');
  toast.textContent = message;
  toast.className = `toast show ${type}`;
  setTimeout(() => {
    toast.className = 'toast';
  }, 3500);
}

// 接続元PCがサーバー（管理者PC）かどうかを確認し、管理者画面へのリンクを表示
async function checkSessionInfo() {
  try {
    const res = await fetch('/api/session');
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
  const currentStaffVal = staffFilter.value;
  staffFilter.innerHTML = '<option value="">全員</option>';
  currentConfig.staff_order.forEach(s => {
    const opt = document.createElement('option');
    opt.value = s;
    opt.textContent = s;
    staffFilter.appendChild(opt);
  });
  staffFilter.value = currentStaffVal;

  // ステータスフィルター
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

  // フォーム用セレクトボックス（自由入力フォールバック付き）
  populateSelectWithCustom('field_staff', currentConfig.staff_order, true);
  populateSelectWithCustom('field_form_type', currentConfig.form_type_options, true);
  populateSelectWithCustom('field_request_content', currentConfig.request_content_options, true);
  populateSelectWithCustom('field_same_day_application', currentConfig.same_day_options, true);
  populateSelectWithCustom('field_status', currentConfig.status_options, false);
}

// セレクトボックスに選択肢と「その他（自由入力...）」を設定
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

// セレクトボックスに値をセット（選択肢にない値の場合は自動的に自由入力欄へ展開）
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

// セレクトボックスまたは自由入力欄から最終的な送信値を取得
function getSelectOrCustomValue(id) {
  const el = document.getElementById(id);
  const customInput = document.getElementById(`${id}_custom`);
  if (!el) return '';
  if (el.value === CUSTOM_OPTION_VALUE) {
    return customInput ? customInput.value.trim() : '';
  }
  return el.value;
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

function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// 編集パネル開閉
function openEditPanel(recordId = null) {
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
    const rec = currentRecords.find(r => r.id === recordId);
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

// 案件削除（ゴミ箱へ移動）
document.getElementById('btnDeleteRecord').addEventListener('click', async () => {
  const id = document.getElementById('field_id').value;
  if (!id) return;
  if (!confirm(`案件 ID: ${id} を削除しますか？（管理者のゴミ箱から復元可能です）`)) {
    return;
  }

  try {
    const res = await fetch(`/api/records/${id}`, { method: 'DELETE' });
    const json = await res.json();
    if (json.success) {
      showToast('案件を削除しました', 'success');
      closeEditPanel();
      loadRecords();
      loadStats();
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
  await checkSessionInfo();
  await loadConfig();
  await loadStats();
  await loadRecords();
});
