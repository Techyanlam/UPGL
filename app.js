/* UPG Pallet 倉存管理 — 前端
 * 所有數據經 Cloudflare Worker API（config.js 嘅 window.UPG_WORKER_URL）。
 * 前端唔寫死任何板數據。
 */
(function () {
  'use strict';

  var WORKER_URL = (window.UPG_WORKER_URL || '').replace(/\/+$/, '');
  var state = { inStock: [], departed: [], pending: [], activity: [] };
  var selectedDepart = null;

  // 每日記錄（Excel 版面改做每日記錄）：讀 Lark Pallet record 嘅
  // 每日快照（daily-record.json，每晚由 GitHub Actions 匯出），
  // 唔再用 Worker 嘅即時倉存。null = 未載入，false = 載入失敗。
  var dailyRecord = null;
  var dailyRecordLoading = false;
  function loadDailyRecord() {
    if (dailyRecord !== null || dailyRecordLoading) return;
    dailyRecordLoading = true;
    fetch('daily-record.json').then(function (r) {
      if (!r.ok) throw new Error('HTTP ' + r.status);
      return r.json();
    }).then(function (j) {
      dailyRecord = j;
      dailyRecordLoading = false;
      // 今日未有快照（今晚 20:00 同步先寫）就顯示最新有數嗰日，
      // 唔好留低一個 0 行嘅今日。
      var latest = '';
      (dailyRecord.rows || []).forEach(function (r) {
        if (r.dmd && r.dmd > latest) latest = r.dmd;
      });
      var cur = $('excelDate').value;
      var hasCur = (dailyRecord.rows || []).some(function (r) { return r.dmd === cur; });
      if (latest && !hasCur) $('excelDate').value = latest;
      renderExcel();
    }).catch(function () {
      dailyRecord = false;
      dailyRecordLoading = false;
      renderExcel();
    });
  }

  var $ = function (id) { return document.getElementById(id); };

  function todayStr() {
    var d = new Date();
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }

  function showError(msg) {
    var box = $('errorBox');
    box.textContent = msg;
    box.style.display = 'block';
  }
  function hideError() { $('errorBox').style.display = 'none'; }

  // 體積：長×闊×高 cm → CBM，ROUNDUP 到 1 位小數
  function calcVolume(lengthCm, widthCm, heightCm) {
    if (!lengthCm || !widthCm || !heightCm) return null;
    var v = (lengthCm * widthCm * heightCm) / 1000000;
    return Math.ceil(v * 10 - 1e-9) / 10;
  }
  function fmtVolume(v) {
    if (v === null || v === undefined || v === '') return '–';
    return (Math.round(v * 10) / 10).toFixed(1);
  }
  function fmtDate(v) { return v ? String(v) : '–'; }

  // 模糊搜尋：唔分大細楷、唔理連字號同空格
  function norm(s) { return String(s || '').toLowerCase().replace(/[-\s]/g, ''); }
  function matches(board, q) {
    if (!q) return true;
    var nq = norm(q);
    return norm(board.palletId).indexOf(nq) !== -1 || norm(board.location).indexOf(nq) !== -1;
  }

  function apiGetState() {
    return fetch(WORKER_URL + '/state').then(function (r) {
      if (!r.ok) throw new Error('Worker 回應 HTTP ' + r.status);
      return r.json();
    });
  }

  function apiPost(body) {
    return fetch(WORKER_URL + '/', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    }).then(function (r) { return r.json(); }).then(function (j) {
      if (!j || j.ok !== true) throw new Error((j && j.error) || 'Worker 回應失敗');
      return j;
    });
  }

  function loadState() {
    if (!WORKER_URL || WORKER_URL.indexOf('%%WORKER_URL%%') !== -1) {
      $('loadingBox').style.display = 'none';
      showError('未設定 Worker 連線：config.js 入面嘅 Worker URL 仲係佔位符，請填入真正嘅 Cloudflare Worker URL 之後再用。');
      return;
    }
    hideError();
    $('loadingBox').style.display = 'block';
    $('app').style.display = 'none';
    apiGetState().then(function (s) {
      state.inStock = s.inStock || [];
      state.departed = s.departed || [];
      state.pending = s.pending || [];
      state.activity = s.activity || [];
      state._inStockCount = s.inStockCount;
      state._inStockVolumeCbm = s.inStockVolumeCbm;
      state._todayDepartures = s.todayDepartures;
      $('loadingBox').style.display = 'none';
      $('app').style.display = 'block';
      renderAll();
    }).catch(function (e) {
      $('loadingBox').style.display = 'none';
      showError('連唔到倉存服務（Cloudflare Worker），請檢查網絡或稍後再試。錯誤：' + e.message);
    });
  }

  function withBusy(btn, fn) {
    btn.disabled = true;
    var old = btn.textContent;
    btn.textContent = '處理中…';
    fn().then(function () { loadState(); })
      .catch(function (e) { showError('操作失敗：' + e.message); })
      .finally(function () { btn.disabled = false; btn.textContent = old; });
  }

  /* ---------- 統計 ---------- */
  function renderStats() {
    var inStock = state.inStock;
    var totalVol = state._inStockVolumeCbm;
    if (totalVol === undefined || totalVol === null) {
      totalVol = inStock.reduce(function (sum, b) {
        var v = (b.volumeCbm !== undefined && b.volumeCbm !== null && b.volumeCbm !== '')
          ? Number(b.volumeCbm) : calcVolume(b.lengthCm, b.widthCm, b.heightCm);
        return sum + (v || 0);
      }, 0);
    }
    var today = state._todayDepartures;
    if (today === undefined || today === null) {
      var t = todayStr();
      today = state.departed.filter(function (b) { return b.departedDate === t; }).length;
    }
    $('statInStock').textContent = (state._inStockCount !== undefined && state._inStockCount !== null)
      ? state._inStockCount : inStock.length;
    $('statVolume').textContent = fmtVolume(totalVol);
    $('statToday').textContent = today;
    $('statPending').textContent = state.pending.length;
  }

  /* ---------- 出倉登記 ---------- */
  function renderDepartDropdown() {
    var q = $('departSearch').value.trim();
    var dd = $('departDropdown');
    if (!q) { dd.style.display = 'none'; return; }
    var hits = state.inStock.filter(function (b) { return matches(b, q); }).slice(0, 20);
    if (!hits.length) {
      dd.innerHTML = '<div>搵唔到在倉板（請檢查板號）</div>';
      dd.style.display = 'block';
      return;
    }
    dd.innerHTML = '';
    hits.forEach(function (b) {
      var div = document.createElement('div');
      div.textContent = b.palletId + '　' + (b.location || '') + '　' + b.lengthCm + '×' + b.widthCm + '×' + b.heightCm + ' cm';
      div.addEventListener('click', function () { selectDepart(b); });
      dd.appendChild(div);
    });
    dd.style.display = 'block';
  }

  function selectDepart(b) {
    selectedDepart = b;
    $('departSearch').value = b.palletId;
    $('departDropdown').style.display = 'none';
    var box = $('departSelected');
    box.textContent = '已揀：' + b.palletId + '（' + (b.location || '–') + '，' +
      b.lengthCm + '×' + b.widthCm + '×' + b.heightCm + ' cm，' + fmtVolume(b.volumeCbm) + ' CBM）';
    box.style.display = 'block';
    $('departBtn').disabled = false;
  }

  /* ---------- 待確認清單 ---------- */
  function renderPending() {
    var list = $('pendingList');
    list.innerHTML = '';
    if (!state.pending.length) {
      var p = document.createElement('p');
      p.className = 'empty';
      p.textContent = '暫無待確認新貨。';
      list.appendChild(p);
      return;
    }
    state.pending.forEach(function (b) {
      var item = document.createElement('div');
      item.className = 'pending-item';
      var vol = (b.volumeCbm !== undefined && b.volumeCbm !== null && b.volumeCbm !== '')
        ? Number(b.volumeCbm) : calcVolume(b.lengthCm, b.widthCm, b.heightCm);
      var received = b.receivedAt || b.addedDate || b.createdAt || '';
      var title = document.createElement('div');
      title.className = 'title';
      title.textContent = b.palletId;
      var meta = document.createElement('div');
      meta.className = 'meta';
      meta.textContent = '尺寸：' + b.lengthCm + '×' + b.widthCm + '×' + b.heightCm + ' cm　體積：' + fmtVolume(vol) +
        ' CBM　收貨日期：' + fmtDate(b.dockReceiptDate) + '　BU：' + (b.businessUnit || '–') +
        '　位置：' + (b.location || '–') + '　來源：Lark 同步' + (received ? '　收到時間：' + received : '');
      var actions = document.createElement('div');
      actions.className = 'actions';
      var reason = document.createElement('input');
      reason.type = 'text';
      reason.placeholder = '唔加入原因（可選填）';
      var locInput = document.createElement('input');
      locInput.type = 'text';
      locInput.placeholder = '新位置'; locInput.style.width = '110px';
      var locBtn = document.createElement('button');
      locBtn.className = 'btn small'; locBtn.type = 'button'; locBtn.textContent = '位置更新';
      locBtn.addEventListener('click', function () {
        var loc = locInput.value.trim();
        if (!loc) { showError('請輸入新位置。'); return; }
        hideError();
        withBusy(locBtn, function () {
          return apiPost({ action: 'updateLocation', palletId: b.palletId, newLocation: loc, date: todayStr() });
        });
      });
      var okBtn = document.createElement('button');
      okBtn.className = 'btn small';
      okBtn.type = 'button';
      okBtn.textContent = '確認加入';
      okBtn.disabled = !b.location || !String(b.location).trim();
      okBtn.addEventListener('click', function () {
        if (!b.location || !String(b.location).trim()) { showError('請先設定位置。'); return; }
        hideError();
        withBusy(okBtn, function () {
          return apiPost({ action: 'confirm', palletId: b.palletId, date: todayStr() });
        });
      });
      var noBtn = document.createElement('button');
      noBtn.className = 'btn small danger';
      noBtn.type = 'button';
      noBtn.textContent = '唔加入';
      noBtn.addEventListener('click', function () {
        withBusy(noBtn, function () {
          var body = { action: 'reject', palletId: b.palletId, date: todayStr() };
          if (reason.value.trim()) body.reason = reason.value.trim();
          return apiPost(body);
        });
      });
      actions.appendChild(reason);
      actions.appendChild(locInput);
      actions.appendChild(locBtn);
      actions.appendChild(okBtn);
      actions.appendChild(noBtn);
      item.appendChild(title);
      item.appendChild(meta);
      item.appendChild(actions);
      list.appendChild(item);
    });
  }

  /* ---------- 在倉表 ---------- */
  function renderStock() {
    var q = $('stockSearch').value.trim();
    var body = $('stockBody');
    body.innerHTML = '';
    var rows = state.inStock.filter(function (b) { return matches(b, q); });
    if (!rows.length) {
      body.innerHTML = '<tr><td colspan="7">暫無在倉板。</td></tr>';
      return;
    }
    rows.forEach(function (b) {
      var tr = document.createElement('tr');
      var vol = (b.volumeCbm !== undefined && b.volumeCbm !== null && b.volumeCbm !== '')
        ? Number(b.volumeCbm) : calcVolume(b.lengthCm, b.widthCm, b.heightCm);
      tr.innerHTML = '<td>' + esc(b.palletId) + '</td><td>' + esc(b.location) + '</td><td>' + esc(b.dockReceiptDate) +
        '</td><td>' + esc(b.businessUnit) + '</td><td>' + esc(b.lengthCm) + '×' + esc(b.widthCm) + '×' + esc(b.heightCm) +
        '</td><td>' + fmtVolume(vol) + '</td>';
      var td = document.createElement('td');
      var input = document.createElement('input');
      input.type = 'number'; input.min = '1'; input.step = '1';
      input.placeholder = '新高度 cm'; input.style.width = '110px';
      var btn = document.createElement('button');
      btn.className = 'btn small'; btn.type = 'button'; btn.textContent = '改高度';
      btn.addEventListener('click', function () {
        var h = Number(input.value);
        if (!h || h <= 0) { showError('請輸入有效嘅新高度。'); return; }
        hideError();
        withBusy(btn, function () {
          return apiPost({ action: 'updateHeight', palletId: b.palletId, heightCm: h, date: todayStr() });
        });
      });
      td.appendChild(input); td.appendChild(btn);
      var locInput = document.createElement('input');
      locInput.type = 'text';
      locInput.placeholder = '新位置'; locInput.style.width = '110px';
      var locBtn = document.createElement('button');
      locBtn.className = 'btn small'; locBtn.type = 'button'; locBtn.textContent = '位置更新';
      locBtn.addEventListener('click', function () {
        var loc = locInput.value.trim();
        if (!loc) { showError('請輸入新位置。'); return; }
        hideError();
        withBusy(locBtn, function () {
          return apiPost({ action: 'updateLocation', palletId: b.palletId, newLocation: loc, date: todayStr() });
        });
      });
      td.appendChild(locInput); td.appendChild(locBtn);
      tr.appendChild(td);
      body.appendChild(tr);
    });
  }

  /* ---------- 已出倉清單 ---------- */
  function renderDeparted() {
    var body = $('departedBody');
    body.innerHTML = '';
    if (!state.departed.length) {
      body.innerHTML = '<tr><td colspan="6">暫無已出倉板。</td></tr>';
      return;
    }
    state.departed.forEach(function (b) {
      var tr = document.createElement('tr');
      var vol = (b.volumeCbm !== undefined && b.volumeCbm !== null && b.volumeCbm !== '')
        ? Number(b.volumeCbm) : calcVolume(b.lengthCm, b.widthCm, b.heightCm);
      tr.innerHTML = '<td>' + esc(b.palletId) + '</td><td>' + esc(b.location) + '</td><td>' + esc(b.dockReceiptDate) +
        '</td><td>' + esc(b.departedDate) + '</td><td>' + fmtVolume(vol) + '</td>';
      var td = document.createElement('td');
      var btn = document.createElement('button');
      btn.className = 'btn small secondary'; btn.type = 'button'; btn.textContent = '還原';
      btn.addEventListener('click', function () {
        withBusy(btn, function () {
          return apiPost({ action: 'restore', palletId: b.palletId, date: todayStr() });
        });
      });
      td.appendChild(btn);
      tr.appendChild(td);
      body.appendChild(tr);
    });
  }

  /* ---------- 活動記錄 ---------- */
  var ACT_LABEL = {
    depart: '出倉', updateHeight: '改高度', updateLocation: '位置更新', add: '入倉', restore: '還原',
    pending: '待確認', confirm: '確認加入', reject: '唔加入'
  };
  function renderActivity() {
    var ul = $('activityList');
    ul.innerHTML = '';
    if (!state.activity.length) {
      ul.innerHTML = '<li>暫無活動記錄。</li>';
      return;
    }
    state.activity.forEach(function (a) {
      var li = document.createElement('li');
      var when = a.timestamp ? String(a.timestamp).slice(0, 16).replace('T', ' ') : (a.datetime || a.time || ((a.date || '') + (a.time ? ' ' + a.time : '')));
      var label = ACT_LABEL[a.type] || a.type || a.action || '活動';
      var text = (when ? when + '　' : '') + label +
        (a.palletId ? '　' + a.palletId : '') +
        (a.detail ? '　' + a.detail : '') +
        (a.message ? '　' + a.message : '') +
        (a.reason ? '（原因：' + a.reason + '）' : '');
      li.textContent = text;
      ul.appendChild(li);
    });
  }

  /* ---------- Excel 版面 ---------- */
  function excelRows() {
    var filter = $('excelFilter').value;
    var q = $('excelSearch').value.trim();
    var date = $('excelDate').value;
    var rows = [];
    if (!dailyRecord || !dailyRecord.rows) return rows;
    dailyRecord.rows.forEach(function (r) {
      var b = {
        palletId: r.palletId, location: r.location,
        dockReceiptDate: r.dockReceiptDate, businessUnit: r.businessUnit,
        reference: r.reference, lengthCm: r.lengthCm, widthCm: r.widthCm,
        heightCm: r.heightCm, volumeCbm: r.volumeCbm,
        addedDate: r.dmd, departedDate: r.shipDate || ''
      };
      // 有 Ship Date 嘅板嗰日已經出倉，唔計入在倉數。
      rows.push({ b: b, inStock: !r.shipDate });
    });
    if (filter === 'in') rows = rows.filter(function (r) { return r.inStock; });
    if (filter === 'out') rows = rows.filter(function (r) { return !r.inStock; });
    if (q) {
      var nq = norm(q);
      rows = rows.filter(function (r) { return norm(r.b.palletId).indexOf(nq) !== -1; });
    }
    if (date) {
      rows = rows.filter(function (r) {
        return String(r.b.addedDate || '').slice(0, 10) === date;
      });
    }
    return rows;
  }

  function renderExcel() {
    var body = $('excelBody');
    body.innerHTML = '';
    if (dailyRecord === null) {
      loadDailyRecord();
      body.innerHTML = '<tr><td colspan="13">載入每日記錄…</td></tr>';
      $('excelTotals').textContent = '';
      return;
    }
    if (dailyRecord === false) {
      body.innerHTML = '<tr><td colspan="13">每日記錄未能載入（daily-record.json 未生成）。</td></tr>';
      $('excelTotals').textContent = '';
      return;
    }
    var rows = excelRows();
    var totalVol = 0;
    rows.forEach(function (r) {
      var b = r.b;
      var vol = (b.volumeCbm !== undefined && b.volumeCbm !== null && b.volumeCbm !== '')
        ? Number(b.volumeCbm) : calcVolume(b.lengthCm, b.widthCm, b.heightCm);
      totalVol += vol || 0;
      var tr = document.createElement('tr');
      var cells = [
        b.addedDate || b.dockReceiptDate || '',   // Daily Measurement Date
        b.location || '',                          // Timestamp（位置）
        b.dockReceiptDate || '',                   // Dock Receipt Date
        b.businessUnit || '',                      // Business Unit
        b.reference || '',                         // Reference
        b.palletId || '',                          // Batch#
        b.addedDate || b.dockReceiptDate || '',   // Date of Measurement
        b.lengthCm || '', b.widthCm || '', b.heightCm || '',
        fmtVolume(vol),
        r.inStock ? '' : (b.departedDate || ''),   // Ship Date
        r.inStock ? '' : '最後'                    // Remark
      ];
      cells.forEach(function (c) {
        var td = document.createElement('td');
        td.textContent = c;
        tr.appendChild(td);
      });
      body.appendChild(tr);
    });
    if (!rows.length) body.innerHTML = '<tr><td colspan="13">暫無資料。</td></tr>';
    $('excelTotals').textContent = '總板數：' + rows.length + '　總體積：' + fmtVolume(totalVol) + ' CBM';
  }

  function exportCsv() {
    var rows = excelRows();
    var header = ['Daily Measurement Date', 'Timestamp', 'Dock Receipt Date', 'Business Unit', 'Reference',
      'Batch#', 'Date of Measurement', 'Length', 'Width', 'Height', 'Volume', 'Ship Date', 'Remark'];
    var lines = [header.join(',')];
    rows.forEach(function (r) {
      var b = r.b;
      var vol = (b.volumeCbm !== undefined && b.volumeCbm !== null && b.volumeCbm !== '')
        ? Number(b.volumeCbm) : calcVolume(b.lengthCm, b.widthCm, b.heightCm);
      var cells = [
        b.addedDate || b.dockReceiptDate || '', b.location || '', b.dockReceiptDate || '',
        b.businessUnit || '', b.reference || '', b.palletId || '',
        b.addedDate || b.dockReceiptDate || '', b.lengthCm || '', b.widthCm || '', b.heightCm || '',
        fmtVolume(vol), r.inStock ? '' : (b.departedDate || ''), r.inStock ? '' : '最後'
      ];
      lines.push(cells.map(csvCell).join(','));
    });
    var blob = new Blob(['\ufeff' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'UPGL-' + ($('excelDate').value || todayStr()) + '.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
  }
  function csvCell(v) {
    var s = String(v === null || v === undefined ? '' : v);
    return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }

  function esc(v) {
    return String(v === null || v === undefined ? '–' : v)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  function renderAll() {
    renderStats();
    renderPending();
    renderStock();
    renderDeparted();
    renderActivity();
    renderExcel();
  }

  /* ---------- 事件綁定 ---------- */
  $('tabDaily').addEventListener('click', function () {
    $('tabDaily').classList.add('active'); $('tabExcel').classList.remove('active');
    $('pageDaily').style.display = 'block'; $('pageExcel').style.display = 'none';
  });
  $('tabExcel').addEventListener('click', function () {
    $('tabExcel').classList.add('active'); $('tabDaily').classList.remove('active');
    $('pageExcel').style.display = 'block'; $('pageDaily').style.display = 'none';
    renderExcel();
  });

  $('departSearch').addEventListener('input', function () { renderDepartDropdown(); });
  $('departToggle').addEventListener('click', function () {
    var body = $('departBody');
    var open = body.style.display === 'none';
    body.style.display = open ? 'block' : 'none';
    $('departToggle').textContent = open ? '－' : '＋';
  });
  $('addToggle').addEventListener('click', function () {
    var body = $('addBody');
    var open = body.style.display === 'none';
    body.style.display = open ? 'block' : 'none';
    $('addToggle').textContent = open ? '－' : '＋';
  });
  document.addEventListener('click', function (e) {
    if (!$('departSearch').contains(e.target) && !$('departDropdown').contains(e.target)) {
      $('departDropdown').style.display = 'none';
    }
  });

  $('departBtn').addEventListener('click', function () {
    if (!selectedDepart) return;
    hideError();
    var body = { action: 'depart', palletId: selectedDepart.palletId, date: $('departDate').value || todayStr() };
    var h = Number($('departHeight').value);
    if (h > 0) body.newHeightCm = h;
    var btn = $('departBtn');
    withBusy(btn, function () { return apiPost(body); });
    selectedDepart = null;
    $('departSelected').style.display = 'none';
    $('departSearch').value = '';
    $('departHeight').value = '';
  });

  $('stockSearch').addEventListener('input', renderStock);
  $('excelFilter').addEventListener('change', renderExcel);
  $('excelSearch').addEventListener('input', renderExcel);
  $('excelDate').addEventListener('change', renderExcel);
  $('csvBtn').addEventListener('click', exportCsv);

  $('addBtn').addEventListener('click', function () {
    hideError();
    var palletId = $('addPalletId').value.trim();
    var height = Number($('addHeight').value);
    var date = $('addDate').value || todayStr();
    if (!palletId) { showError('請輸入 Batch#。'); return; }
    if (!height || height <= 0) { showError('請輸入有效嘅高度。'); return; }
    var body = {
      action: 'add', palletId: palletId, heightCm: height, date: date,
      dockReceiptDate: $('addDockDate').value || undefined,
      location: $('addLocation').value.trim() || undefined,
      businessUnit: $('addBU').value.trim() || undefined,
      reference: $('addReference').value.trim() || undefined,
      lengthCm: Number($('addLength').value) || undefined,
      widthCm: Number($('addWidth').value) || undefined
    };
    withBusy($('addBtn'), function () { return apiPost(body); });
  });

  // 預設值
  $('departDate').value = todayStr();
  $('addDate').value = todayStr();
  $('excelDate').value = todayStr();

  loadState();
})();
