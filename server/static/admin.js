let adminConfig = null;

function showToast(message, type = 'info') {
  const toast = document.getElementById('toast');
  toast.textContent = message;
  toast.className = `toast show ${type}`;
  setTimeout(() => {
    toast.className = 'toast';
  }, 3500);
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

// 統計情報の読み込み
async function loadAdminStats() {
  try {
    const res = await fetch('/api/stats');
    const json = await res.json();
    if (json.success && json.data) {
      document.getElementById('adminStatTotal').textContent = json.data.total_count;
      document.getElementById('adminStatTrash').textContent = json.data.trash_count;
      document.getElementById('navTrashCount').textContent = json.data.trash_count;
    }
  } catch (err) {
    console.error('統計取得失敗:', err);
  }
}

// 社員共有用アクセスURLの読み込みとコピー機能
async function loadShareUrls() {
  const container = document.getElementById('shareUrlContainer');
  if (!container) return;

  try {
    const res = await fetch('/api/session');
    const json = await res.json();
    if (json.success && json.data && Array.isArray(json.data.share_urls)) {
      container.innerHTML = '';
      json.data.share_urls.forEach((url, idx) => {
        const label = idx === 0 ? 'IPアドレス版 (推奨)' : 'PC名版';
        const box = document.createElement('div');
        box.className = 'share-url-box';
        box.innerHTML = `
          <span style="font-size:11px; color:#64748b; font-weight:600;">${escapeHtml(label)}:</span>
          <input type="text" readonly value="${escapeHtml(url)}" id="shareUrlInput_${idx}">
          <button type="button" class="btn btn-sm btn-primary copy-url-btn" data-url="${escapeHtml(url)}">URLをコピー</button>
        `;
        container.appendChild(box);
      });

      container.querySelectorAll('.share-url-box input').forEach(inp => {
        inp.addEventListener('click', () => inp.select());
      });

      container.querySelectorAll('.copy-url-btn').forEach(btn => {
        btn.addEventListener('click', async () => {
          const url = btn.dataset.url;
          try {
            await navigator.clipboard.writeText(url);
            showToast(`コピーしました: ${url}`, 'success');
          } catch (e) {
            // フォールバック
            const inp = btn.parentElement.querySelector('input');
            if (inp) {
              inp.select();
              document.execCommand('copy');
              showToast(`コピーしました: ${url}`, 'success');
            }
          }
        });
      });
    }
  } catch (err) {
    console.error('共有URL取得失敗:', err);
  }
}

// 設定の読み込みと画面への反映
async function loadAdminConfig() {
  try {
    const res = await fetch('/api/config');
    const json = await res.json();
    if (json.success) {
      adminConfig = json.data;
      populateConfigFields();
    }
  } catch (err) {
    showToast('設定の読み込みに失敗しました', 'error');
  }
}

function populateConfigFields() {
  if (!adminConfig) return;

  // 1. 担当順リスト
  renderStaffOrderList(adminConfig.staff_order);

  // 2. ソートキー
  const keys = adminConfig.sort_keys || [];
  if (keys[0]) {
    document.getElementById('sortKey1').value = keys[0].key;
    document.getElementById('sortOrder1').value = keys[0].ascending ? 'asc' : 'desc';
  }
  if (keys[1]) {
    document.getElementById('sortKey2').value = keys[1].key;
    document.getElementById('sortOrder2').value = keys[1].ascending ? 'asc' : 'desc';
  }
  if (keys[2]) {
    document.getElementById('sortKey3').value = keys[2].key;
    document.getElementById('sortOrder3').value = keys[2].ascending ? 'asc' : 'desc';
  }

  // 3. 選択肢マスター
  document.getElementById('settingFormTypes').value = (adminConfig.form_type_options || []).join('\n');
  document.getElementById('settingRequestContents').value = (adminConfig.request_content_options || []).join('\n');
  document.getElementById('settingSameDayOptions').value = (adminConfig.same_day_options || []).join('\n');
  document.getElementById('settingStatusOptions').value = (adminConfig.status_options || []).join('\n');

  // 4. 出力先と自動化設定
  document.getElementById('settingExportDir').value = adminConfig.export_dir || './exports';
  document.getElementById('settingExportFileName').value = adminConfig.export_file_name || '受託確認票.xlsx';
  document.getElementById('settingRetentionDays').value = adminConfig.backup_retention_days || 7;
  document.getElementById('settingAutoExport').checked = !!adminConfig.auto_export_on_change;

  // 5. 遮断IPリストの描画
  renderBlockedIps();
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
        <button type="button" class="btn btn-sm btn-outline move-up-btn" data-idx="${idx}">↑ 上へ</button>
        <button type="button" class="btn btn-sm btn-outline move-down-btn" data-idx="${idx}">↓ 下へ</button>
        <button type="button" class="btn btn-sm btn-outline delete-staff-btn" data-idx="${idx}" style="color:var(--danger);">削除</button>
      </div>
    `;
    ul.appendChild(li);
  });

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
  if (val && adminConfig) {
    adminConfig.staff_order.push(val);
    renderStaffOrderList(adminConfig.staff_order);
    input.value = '';
  }
});

// 設定の保存（全設定ボタン共通）
async function saveConfig() {
  if (!adminConfig) return;

  adminConfig.sort_keys = [
    {
      key: document.getElementById('sortKey1').value,
      ascending: document.getElementById('sortOrder1').value === 'asc',
    },
    {
      key: document.getElementById('sortKey2').value,
      ascending: document.getElementById('sortOrder2').value === 'asc',
    },
    {
      key: document.getElementById('sortKey3').value,
      ascending: document.getElementById('sortOrder3').value === 'asc',
    },
  ];

  adminConfig.form_type_options = document.getElementById('settingFormTypes').value
    .split('\n')
    .map(s => s.trim())
    .filter(Boolean);

  adminConfig.request_content_options = document.getElementById('settingRequestContents').value
    .split('\n')
    .map(s => s.trim())
    .filter(Boolean);

  adminConfig.same_day_options = document.getElementById('settingSameDayOptions').value
    .split('\n')
    .map(s => s.trim())
    .filter(Boolean);

  adminConfig.status_options = document.getElementById('settingStatusOptions').value
    .split('\n')
    .map(s => s.trim())
    .filter(Boolean);

  adminConfig.export_dir = document.getElementById('settingExportDir').value.trim();
  adminConfig.export_file_name = document.getElementById('settingExportFileName').value.trim();
  adminConfig.backup_retention_days = parseInt(document.getElementById('settingRetentionDays').value, 10) || 7;
  adminConfig.auto_export_on_change = document.getElementById('settingAutoExport').checked;

  try {
    const res = await fetch('/api/config', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(adminConfig),
    });
    const json = await res.json();
    if (json.success) {
      showToast('管理者設定を保存し、Excelへ反映しました', 'success');
    } else {
      showToast(json.error || '設定保存に失敗しました', 'error');
    }
  } catch (err) {
    showToast('通信エラーが発生しました', 'error');
  }
}

document.querySelectorAll('.btn-save-config').forEach(btn => {
  btn.addEventListener('click', saveConfig);
});

// 全画面表示トグル
const btnAdminFullscreen = document.getElementById('btnAdminFullscreen');
if (btnAdminFullscreen) {
  btnAdminFullscreen.addEventListener('click', async () => {
    try {
      if (!document.fullscreenElement) {
        await document.documentElement.requestFullscreen();
        document.body.classList.add('wide-fullscreen');
      } else {
        await document.exitFullscreen();
        document.body.classList.remove('wide-fullscreen');
      }
    } catch (e) {
      document.body.classList.toggle('wide-fullscreen');
    }
    btnAdminFullscreen.textContent =
      (document.fullscreenElement || document.body.classList.contains('wide-fullscreen'))
        ? '全画面を解除'
        : '全画面表示';
  });
}

document.addEventListener('fullscreenchange', () => {
  if (btnAdminFullscreen) {
    btnAdminFullscreen.textContent = document.fullscreenElement ? '全画面を解除' : '全画面表示';
  }
});

// 手動Excel出力
document.getElementById('btnManualExport').addEventListener('click', async () => {
  showToast('Excelを再出力しています...', 'info');
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

// Excel一括インポート
document.getElementById('btnExecuteImport').addEventListener('click', async () => {
  const filePath = document.getElementById('importFilePath').value.trim();
  const mode = document.getElementById('importMode').value;
  const resDiv = document.getElementById('importResult');

  if (mode === 'replace') {
    if (!confirm('現在の全案件データをクリアして、指定したExcelファイルの内容で総入れ替えします。よろしいですか？')) {
      return;
    }
  }

  resDiv.style.color = 'var(--text-muted)';
  resDiv.textContent = '取り込みを実行中...';

  try {
    const res = await fetch('/api/import', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ file_path: filePath, mode }),
    });
    const json = await res.json();
    if (json.success) {
      resDiv.style.color = 'var(--secondary)';
      resDiv.textContent = json.data;
      showToast(json.data, 'success');
      loadAdminStats();
    } else {
      resDiv.style.color = 'var(--danger)';
      resDiv.textContent = `エラー: ${json.error}`;
    }
  } catch (err) {
    resDiv.style.color = 'var(--danger)';
    resDiv.textContent = '通信エラーが発生しました';
  }
});

// バックアップ一覧と操作
async function loadBackupList() {
  const tbody = document.getElementById('backupListTbody');
  tbody.innerHTML = '';

  try {
    const res = await fetch('/api/backups');
    const json = await res.json();
    if (json.success && json.data && json.data.length > 0) {
      json.data.forEach(b => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td><b>${escapeHtml(b.date_str)}</b></td>
          <td>${escapeHtml(b.created_at)}</td>
          <td>${escapeHtml(b.excel_file || '-')}</td>
          <td>${escapeHtml(b.db_file || '-')}</td>
          <td>
            ${
              b.db_file
                ? `<button class="btn btn-sm btn-outline restore-backup-btn" data-date="${escapeHtml(b.date_str)}">この時点に復元</button>`
                : '-'
            }
          </td>
        `;
        tbody.appendChild(tr);
      });

      tbody.querySelectorAll('.restore-backup-btn').forEach(btn => {
        btn.addEventListener('click', async () => {
          const dateStr = btn.dataset.date;
          if (!confirm(`${dateStr} のバックアップ時点へ全案件データを復元しますか？現在のデータは上書きされます。`)) {
            return;
          }
          try {
            const r = await fetch('/api/backups/restore', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ date_str: dateStr }),
            });
            const j = await r.json();
            if (j.success) {
              showToast(j.data, 'success');
              loadAdminStats();
            } else {
              showToast(j.error || '復元に失敗しました', 'error');
            }
          } catch (e) {
            showToast('通信エラーが発生しました', 'error');
          }
        });
      });
    } else {
      const tr = document.createElement('tr');
      const td = document.createElement('td');
      td.colSpan = 5;
      td.style.textAlign = 'center';
      td.style.color = 'var(--text-muted)';
      td.textContent = 'バックアップ履歴はまだありません';
      tr.appendChild(td);
      tbody.appendChild(tr);
    }
  } catch (err) {
    console.error('バックアップ一覧取得失敗:', err);
  }
}

document.getElementById('btnCreateBackupNow').addEventListener('click', async () => {
  try {
    const res = await fetch('/api/backups', { method: 'POST' });
    const json = await res.json();
    if (json.success) {
      showToast(json.data || 'バックアップを作成しました', 'success');
      loadBackupList();
    } else {
      showToast(json.error || 'バックアップ作成に失敗しました', 'error');
    }
  } catch (err) {
    showToast('通信エラーが発生しました', 'error');
  }
});

// ゴミ箱（削除済み案件）一覧と操作
async function loadTrashList() {
  const tbody = document.getElementById('trashListTbody');
  tbody.innerHTML = '';

  try {
    const res = await fetch('/api/trash');
    const json = await res.json();
    if (json.success && json.data && json.data.length > 0) {
      json.data.forEach(r => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td>${escapeHtml(String(r.id))}</td>
          <td>${escapeHtml(r.updated_at)}</td>
          <td>${escapeHtml(r.status)}</td>
          <td><b>${escapeHtml(r.client_name)}</b></td>
          <td>${escapeHtml(r.property_name)}</td>
          <td>${escapeHtml(r.staff)}</td>
          <td>${escapeHtml(r.loan_date || r.loan_date_note || '-')}</td>
          <td style="display:flex; gap:6px;">
            <button class="btn btn-sm btn-secondary restore-trash-btn" data-id="${escapeHtml(String(r.id))}">復元</button>
            <button class="btn btn-sm btn-outline purge-trash-btn" data-id="${escapeHtml(String(r.id))}" style="color:var(--danger); border-color:var(--danger);">完全削除</button>
          </td>
        `;
        tbody.appendChild(tr);
      });

      tbody.querySelectorAll('.restore-trash-btn').forEach(btn => {
        btn.addEventListener('click', async () => {
          const id = btn.dataset.id;
          try {
            const r = await fetch(`/api/trash/${id}/restore`, { method: 'POST' });
            const j = await r.json();
            if (j.success) {
              showToast('案件を復元し、Excelを更新しました', 'success');
              loadTrashList();
              loadAdminStats();
            } else {
              showToast(j.error || '復元に失敗しました', 'error');
            }
          } catch (e) {
            showToast('通信エラーが発生しました', 'error');
          }
        });
      });

      tbody.querySelectorAll('.purge-trash-btn').forEach(btn => {
        btn.addEventListener('click', async () => {
          const id = btn.dataset.id;
          if (!confirm(`案件 ID: ${id} をデータベースから完全に消去しますか？（元に戻せません）`)) {
            return;
          }
          try {
            const r = await fetch(`/api/trash/${id}`, { method: 'DELETE' });
            const j = await r.json();
            if (j.success) {
              showToast('案件を完全削除しました', 'success');
              loadTrashList();
              loadAdminStats();
            } else {
              showToast(j.error || '完全削除に失敗しました', 'error');
            }
          } catch (e) {
            showToast('通信エラーが発生しました', 'error');
          }
        });
      });
    } else {
      const tr = document.createElement('tr');
      const td = document.createElement('td');
      td.colSpan = 8;
      td.style.textAlign = 'center';
      td.style.color = 'var(--text-muted)';
      td.textContent = '削除された案件はありません';
      tr.appendChild(td);
      tbody.appendChild(tr);
    }
  } catch (err) {
    console.error('ゴミ箱取得失敗:', err);
  }
}

// IPアドレス遮断（出禁）/ 解除の切り替え処理
function extractRawIp(ipDisplay) {
  if (!ipDisplay) return '';
  return String(ipDisplay).split(/\s+/)[0].trim();
}

function isLoopbackIp(rawIp) {
  return !rawIp || rawIp === '127.0.0.1' || rawIp === '::1' || rawIp.toLowerCase() === 'localhost';
}

function renderBlockedIps() {
  const container = document.getElementById('blockedIpBadges');
  if (!container || !adminConfig) return;
  const list = adminConfig.blocked_ips || [];
  container.innerHTML = '';

  if (list.length === 0) {
    container.innerHTML = '<span style="font-size: 12px; color: #64748b;">現在遮断されているIPアドレスはありません</span>';
    return;
  }

  list.forEach(ip => {
    const span = document.createElement('span');
    span.className = 'stat-badge';
    span.style.backgroundColor = '#ffe4e6';
    span.style.color = '#9f1239';
    span.style.border = '1px solid #fda4af';
    span.style.padding = '4px 10px';
    span.style.fontFamily = 'monospace';
    span.innerHTML = `
      <b>${escapeHtml(ip)}</b>
      <button type="button" class="btn btn-sm btn-outline unblock-ip-btn" data-ip="${escapeHtml(ip)}" style="padding: 1px 6px; font-size: 11px; margin-left: 6px;">解除</button>
    `;
    container.appendChild(span);
  });

  container.querySelectorAll('.unblock-ip-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      toggleIpBlockStatus(btn.dataset.ip, false);
    });
  });
}

async function toggleIpBlockStatus(ipStr, shouldBlock) {
  const rawIp = extractRawIp(ipStr);
  if (!rawIp) return;

  if (shouldBlock) {
    if (!confirm(`IPアドレス「${rawIp}」からのアクセスを遮断（出禁）にしますか？\n（いつでも解除できます）`)) {
      return;
    }
  }

  try {
    const res = await fetch('/api/ip-block', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ ip: rawIp, block: shouldBlock }),
    });
    const json = await res.json();
    if (json.success && Array.isArray(json.data)) {
      if (adminConfig) {
        adminConfig.blocked_ips = json.data;
      }
      renderBlockedIps();
      await loadAuditLogs();
      showToast(
        shouldBlock ? `IP: ${rawIp} をアクセス遮断（出禁）に設定しました` : `IP: ${rawIp} のアクセス遮断を解除しました`,
        shouldBlock ? 'info' : 'success'
      );
    } else {
      showToast(json.error || 'IP制限の更新に失敗しました', 'error');
    }
  } catch (err) {
    showToast('通信エラーが発生しました', 'error');
  }
}

document.getElementById('btnManualBlockIp').addEventListener('click', () => {
  const input = document.getElementById('manualBlockIpInput');
  const val = (input.value || '').trim();
  if (!val) {
    showToast('遮断するIPアドレスを入力してください', 'error');
    return;
  }
  toggleIpBlockStatus(val, true);
  input.value = '';
});

// アクセス・操作履歴ログ一覧と検索
function getActionBadge(actionType) {
  const map = {
    CREATE: { label: '新規登録', cls: 'badge-done' },
    UPDATE: { label: '案件更新', cls: 'badge-active' },
    DELETE: { label: '削除', cls: 'badge-pending' },
    RESTORE: { label: '復元', cls: 'badge-done' },
    PURGE: { label: '完全削除', cls: 'badge-pending' },
    EXPORT: { label: 'Excel出力', cls: 'stat-badge' },
    CONFIG: { label: '設定変更', cls: 'stat-badge' },
    IMPORT: { label: '一括取込', cls: 'badge-active' },
    BACKUP: { label: 'バックアップ', cls: 'stat-badge' },
    ACCESS: { label: '画面アクセス', cls: 'stat-badge' },
  };
  const item = map[actionType] || { label: actionType, cls: 'stat-badge' };
  return `<span class="stat-badge ${item.cls}">${escapeHtml(item.label)}</span>`;
}

async function loadAuditLogs() {
  const tbody = document.getElementById('auditLogsTbody');
  if (!tbody) return;
  tbody.innerHTML = '';

  const q = (document.getElementById('logSearchInput').value || '').trim();
  const actionType = document.getElementById('logActionFilter').value || '';

  const params = new URLSearchParams();
  if (q) params.set('q', q);
  if (actionType) params.set('action_type', actionType);
  params.set('limit', '500');

  try {
    const res = await fetch(`/api/audit-logs?${params.toString()}`);
    const json = await res.json();
    if (json.success && json.data && json.data.length > 0) {
      const blockedList = (adminConfig && adminConfig.blocked_ips) ? adminConfig.blocked_ips : [];

      json.data.forEach(log => {
        const rawIp = extractRawIp(log.client_ip);
        const isLocal = isLoopbackIp(rawIp);
        const isBlocked = blockedList.includes(rawIp);

        let ipActionBtn = '';
        if (!isLocal) {
          if (isBlocked) {
            ipActionBtn = `<button type="button" class="btn btn-sm btn-outline log-ip-toggle-btn" data-ip="${escapeHtml(rawIp)}" data-block="false" style="margin-left:8px; color:var(--secondary); border-color:var(--secondary); padding:2px 6px; font-size:11px;">遮断中 (解除)</button>`;
          } else {
            ipActionBtn = `<button type="button" class="btn btn-sm btn-outline log-ip-toggle-btn" data-ip="${escapeHtml(rawIp)}" data-block="true" style="margin-left:8px; color:var(--danger); border-color:var(--danger); padding:2px 6px; font-size:11px;">出禁にする</button>`;
          }
        }

        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td>${escapeHtml(log.created_at)}</td>
          <td>${getActionBadge(log.action_type)}</td>
          <td><b>${escapeHtml(log.actor_name || '(未選択)')}</b></td>
          <td style="font-family: monospace;">
            <span>${escapeHtml(log.client_ip)}</span>
            ${ipActionBtn}
          </td>
          <td><b>${escapeHtml(log.summary)}</b></td>
          <td style="white-space: normal; max-width: 480px; color: #334155;">${escapeHtml(log.details)}</td>
        `;
        tbody.appendChild(tr);
      });

      tbody.querySelectorAll('.log-ip-toggle-btn').forEach(btn => {
        btn.addEventListener('click', () => {
          const ip = btn.dataset.ip;
          const shouldBlock = btn.dataset.block === 'true';
          toggleIpBlockStatus(ip, shouldBlock);
        });
      });
    } else {
      const tr = document.createElement('tr');
      const td = document.createElement('td');
      td.colSpan = 6;
      td.style.textAlign = 'center';
      td.style.color = 'var(--text-muted)';
      td.textContent = '該当する履歴ログはありません';
      tr.appendChild(td);
      tbody.appendChild(tr);
    }
  } catch (err) {
    console.error('監査ログ取得失敗:', err);
  }
}

let logDebounceTimer;
document.getElementById('logSearchInput').addEventListener('input', () => {
  clearTimeout(logDebounceTimer);
  logDebounceTimer = setTimeout(loadAuditLogs, 200);
});
document.getElementById('logActionFilter').addEventListener('change', loadAuditLogs);
document.getElementById('btnRefreshLogs').addEventListener('click', loadAuditLogs);

// ナビゲーション切り替え
document.querySelectorAll('.admin-nav-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.admin-nav-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.admin-card').forEach(c => c.classList.remove('active'));

    btn.classList.add('active');
    const targetId = btn.dataset.target;
    document.getElementById(targetId).classList.add('active');

    if (targetId === 'pane-backup') {
      loadBackupList();
    } else if (targetId === 'pane-trash') {
      loadTrashList();
    } else if (targetId === 'pane-logs') {
      loadAuditLogs();
    }
  });
});

window.addEventListener('DOMContentLoaded', async () => {
  await Promise.all([
    loadShareUrls(),
    loadAdminStats(),
    loadAdminConfig(),
    loadBackupList(),
    loadTrashList(),
    loadAuditLogs(),
  ]);
});
