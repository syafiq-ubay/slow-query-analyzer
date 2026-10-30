(() => {
  const $ = (id) => document.getElementById(id);
  const state = { rows: [], filtered: [], displayRows: [], page: 0, pageSize: 10, selectedUser: '', uniqueMode: false };
  const fields = {
    time: 'mysql_@timestamp', user: 'mysql_user', schema: 'mysql_schema', queryTime: 'mysql_query_time',
    status: 'http_status_code', responseTime: 'http_response_time', path: 'http_path', referer: 'http_referer', query: 'mysql_query'
  };
  const number = (value) => {
    if (typeof value === 'number') return Number.isFinite(value) ? value : null;
    if (value == null || String(value).trim() === '') return null;
    const parsed = Number(String(value).replace(',', '.'));
    return Number.isFinite(parsed) ? parsed : null;
  };
  const display = (value) => value == null || value === '' ? '—' : String(value);
  const average = (values) => values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : null;
  const formatSeconds = (value) => value == null ? '—' : `${value < 0.01 && value > 0 ? value.toFixed(4) : value.toFixed(2)}s`;
  const escapeHtml = (value) => String(value ?? '').replace(/[&<>"']/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char]));

  $('excel-file').addEventListener('change', (event) => {
    const file = event.target.files?.[0];
    if (file) loadWorkbook(file);
  });
  $('clear-file').addEventListener('click', () => { $('excel-file').value = ''; $('excel-file').click(); });
  $('search').addEventListener('input', (event) => { filterRows(event.target.value); });
  $('user-filter').addEventListener('change', (event) => { state.selectedUser = event.target.value; filterRows($('search').value); });
  $('unique-mode').addEventListener('change', (event) => { state.uniqueMode = event.target.value === 'unique'; state.page = 0; renderTable(); });
  $('page-size').addEventListener('change', (event) => { state.pageSize = Number(event.target.value); state.page = 0; renderTable(); });
  $('export-excel').addEventListener('click', exportFilteredExcel);
  document.querySelectorAll('[data-download-chart]').forEach((button) => {
    button.addEventListener('click', () => downloadChart(button.dataset.downloadChart));
  });
  $('previous-page').addEventListener('click', () => { state.page--; renderTable(); });
  $('next-page').addEventListener('click', () => { state.page++; renderTable(); });
  document.addEventListener('click', (event) => {
    const exportButton = event.target.closest('[data-export-path-user]');
    if (exportButton) {
      exportPathExcel(exportButton.dataset.exportPathUser);
      return;
    }
    const userCanvas = $('user-chart');
    if (userCanvas.pinnedPoint && !event.target.closest('.chart-tooltip') && event.target !== userCanvas) {
      userCanvas.pinnedPoint = null;
      userCanvas.chartTooltip.hidden = true;
    }
  });
  window.addEventListener('resize', debounce(() => { if (state.rows.length) drawCharts(); }, 120));

  // ─── Data Source Mode (auto / manual) ────────────────────────────────────────
  (function initDataMode() {
    const cfg = (window.APP_CONFIG) || {};
    const mode = (cfg.DATA_MODE || 'manual').toLowerCase();
    const filePath = (cfg.AUTO_FILE_PATH || 'output/merge_logs.xlsx').trim();

    const badge = $('mode-badge');
    const uploadLabel = $('upload-label');
    const autoFetchBtn = $('auto-fetch-btn');

    function setModeBadge(activeMode, isFallback) {
      if (!badge) return;
      badge.hidden = false;
      badge.className = 'mode-badge ' + (activeMode === 'auto' ? 'mode-auto' : 'mode-manual');
      if (activeMode === 'auto') {
        badge.textContent = 'Mode Auto';
      } else if (isFallback) {
        badge.textContent = 'Mode Manual (fallback)';
        badge.classList.add('mode-manual');
      } else {
        badge.textContent = 'Mode Manual';
      }
    }

    async function fetchAutoFile() {
      $('loading-title').textContent = `Memuat otomatis: ${filePath}`;
      $('loading-overlay').hidden = false;
      try {
        const response = await fetch(filePath, { cache: 'no-store' });
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const arrayBuffer = await response.arrayBuffer();
        const file = new File([arrayBuffer], filePath.split('/').pop() || 'merge_logs.xlsx', {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
        await loadWorkbook(file);
        setModeBadge('auto', false);
        // Show refresh button, hide manual upload in intro
        if (uploadLabel) uploadLabel.hidden = true;
        if (autoFetchBtn) autoFetchBtn.hidden = false;
      } catch (err) {
        // Fallback to manual mode
        console.warn('[DataMode] Auto fetch gagal, fallback ke manual:', err.message);
        setModeBadge('manual', true);
        showNotice(
          `Mode auto: file "${filePath}" tidak dapat diakses (${err.message}). ` +
          'Beralih ke mode manual — silakan upload file Excel secara manual.'
        );
        if (uploadLabel) {
          uploadLabel.hidden = false;
          uploadLabel.classList.add('auto-fallback');
        }
        if (autoFetchBtn) autoFetchBtn.hidden = true;
      } finally {
        $('loading-overlay').hidden = true;
      }
    }

    if (mode === 'auto') {
      // Trigger auto fetch on load
      fetchAutoFile();
      // Wire refresh button
      if (autoFetchBtn) autoFetchBtn.addEventListener('click', fetchAutoFile);
    } else {
      // Manual mode: normal UI
      setModeBadge('manual', false);
      if (uploadLabel) uploadLabel.hidden = false;
      if (autoFetchBtn) autoFetchBtn.hidden = true;
    }
  })();


  async function loadWorkbook(file) {
    hideNotice();
    if (!window.XLSX) return showNotice('Library pembaca Excel belum termuat. Periksa koneksi internet lalu muat ulang halaman.');
    $('loading-title').textContent = `Memproses ${file.name}`;
    $('loading-overlay').hidden = false;
    try {
      await new Promise((resolve) => requestAnimationFrame(resolve));
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: 'array', cellDates: true });
      const sheet = workbook.Sheets['Combined Logs'] || workbook.Sheets[workbook.SheetNames[0]];
      if (!sheet) throw new Error('Workbook tidak memiliki sheet yang dapat dibaca.');
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
      if (!rows.length) throw new Error('Sheet Excel tidak memiliki baris data.');
      const headers = Object.keys(rows[0]);
      const missing = Object.values(fields).filter((field) => !headers.includes(field));
      if (missing.length) throw new Error(`Kolom merge tidak lengkap. Kolom yang tidak ditemukan: ${missing.join(', ')}`);
      state.rows = rows;
      $('file-name').textContent = file.name;
      $('file-detail').textContent = `${rows.length.toLocaleString('id-ID')} baris · ${workbook.SheetNames.length} sheet`;
      $('file-strip').hidden = false;
      $('empty-state').hidden = true;
      $('dashboard').hidden = false;
      $('notice').hidden = true;
      populateUsers();
      renderMetrics();
      renderCharts();
      filterRows($('search').value);
    } catch (error) {
      showNotice(error.message || 'File Excel tidak dapat dibaca.');
    } finally {
      $('loading-overlay').hidden = true;
    }
  }

  function renderMetrics() {
    const queryTimes = state.rows.map((row) => number(row[fields.queryTime])).filter((value) => value != null);
    const responseTimes = state.rows.map((row) => number(row[fields.responseTime])).filter((value) => value != null);
    const statuses = state.rows.map((row) => number(row[fields.status])).filter((value) => value != null);
    const errors = statuses.filter((status) => status >= 400 && status < 600).length;
    $('metric-total').textContent = state.rows.length.toLocaleString('id-ID');
    $('metric-query-time').textContent = formatSeconds(average(queryTimes));
    $('metric-response-time').textContent = formatSeconds(average(responseTimes));
    $('metric-errors').textContent = errors.toLocaleString('id-ID');
    $('metric-error-rate').textContent = `${statuses.length ? (errors / statuses.length * 100).toFixed(1) : '0'}% dari request`;
  }

  function renderCharts() {
    drawUserChart(userData());
    drawChart('schema-chart', 'schema-empty', schemaData(), '#6b5ce7', true, 'detik');
    drawChart('status-chart', 'status-empty', statusData(), '#4b9ce8', false, 'request');
    drawTrendChart(trendData());
    drawScatterChart(scatterData());
  }
  function trendData() {
    const offset = 7 * 60 * 60 * 1000, buckets = new Map();
    state.rows.forEach((row) => {
      const time = parseLogDate(row[fields.time]), queryTime = number(row[fields.queryTime]);
      if (!time || queryTime == null) return;
      const hour = Math.floor((time.getTime() + offset) / 3600000) * 3600000 - offset;
      if (!buckets.has(hour)) buckets.set(hour, []);
      buckets.get(hour).push({ queryTime, schema: String(row[fields.schema] || '(schema kosong)'), user: String(row[fields.user] || '(user kosong)'), path: String(row[fields.path] || '') });
    });
    const label = (time) => new Intl.DateTimeFormat('id-ID', { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit', hour12: false, timeZone: 'Asia/Jakarta' }).format(new Date(time));
    return [...buckets.entries()].sort((a, b) => a[0] - b[0]).map(([time, records]) => {
      const values = records.map((record) => record.queryTime), schemaCounts = new Map();
      records.forEach((record) => schemaCounts.set(record.schema, (schemaCounts.get(record.schema) || 0) + 1));
      const peak = records.reduce((best, record) => record.queryTime > best.queryTime ? record : best, records[0]);
      const topSchema = [...schemaCounts.entries()].sort((a, b) => b[1] - a[1])[0];
      const min = values.reduce((best, value) => Math.min(best, value), Infinity), max = values.reduce((best, value) => Math.max(best, value), -Infinity);
      return { time, label: label(time), value: average(values), count: values.length, min, max, peak, topSchema };
    });
  }
  function drawTrendChart(points) {
    const canvas = $('trend-chart'), empty = $('trend-empty'), rect = canvas.getBoundingClientRect();
    if (!points.length || rect.width < 2) { canvas.hidden = true; empty.hidden = false; return; }
    canvas.hidden = false; empty.hidden = true;
    const ratio = window.devicePixelRatio || 1; canvas.width = Math.round(rect.width * ratio); canvas.height = Math.round(rect.height * ratio);
    const ctx = canvas.getContext('2d'); ctx.scale(ratio, ratio);
    const width = rect.width, height = rect.height, margin = { left: 48, right: 14, top: 15, bottom: 34 };
    const plotW = width - margin.left - margin.right, plotH = height - margin.top - margin.bottom;
    const max = Math.max(...points.map(point => point.value), 0.0001), hits = [];
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, width, height); ctx.font = '9px DM Sans, sans-serif'; ctx.textBaseline = 'middle';
    for (let i = 0; i <= 4; i++) {
      const y = margin.top + plotH * i / 4, value = max * (1 - i / 4);
      ctx.strokeStyle = '#edf0f6'; ctx.beginPath(); ctx.moveTo(margin.left, y); ctx.lineTo(width - margin.right, y); ctx.stroke();
      ctx.fillStyle = '#98a3b5'; ctx.textAlign = 'right'; ctx.fillText(compact(value), margin.left - 7, y);
    }
    const slot = points.length > 1 ? plotW / (points.length - 1) : 0;
    const coordinates = points.map((point, index) => ({ x: points.length > 1 ? margin.left + slot * index : margin.left + plotW / 2, y: margin.top + plotH - (point.value / max) * plotH, point }));
    if (coordinates.length > 1) { ctx.beginPath(); coordinates.forEach((item, index) => index ? ctx.lineTo(item.x, item.y) : ctx.moveTo(item.x, item.y)); ctx.strokeStyle = '#6658df'; ctx.lineWidth = 2; ctx.stroke(); }
    coordinates.forEach((item) => {
      ctx.beginPath(); ctx.arc(item.x, item.y, 3.5, 0, Math.PI * 2); ctx.fillStyle = '#6658df'; ctx.fill(); ctx.strokeStyle = '#fff'; ctx.lineWidth = 1.5; ctx.stroke();
      hits.push({ x: item.x - 9, y: item.y - 9, w: 18, h: 18, point: item.point });
    });
    const labelCount = Math.min(points.length, 7);
    for (let i = 0; i < labelCount; i++) {
      const index = labelCount === 1 ? 0 : Math.round(i * (points.length - 1) / (labelCount - 1)), item = coordinates[index];
      ctx.fillStyle = '#8290a5'; ctx.textAlign = index === 0 ? 'left' : index === points.length - 1 ? 'right' : 'center';
      ctx.fillText(item.point.label, item.x, height - 13);
    }
    bindTooltip(canvas, getTooltip(canvas), hits, (point) => `<strong>${escapeHtml(point.label)} WIB</strong><span>Rata-rata query: ${point.value.toFixed(4)} detik</span><span>${point.count.toLocaleString('id-ID')} slow query</span><span>Tercepat: ${point.min.toFixed(4)} detik</span><span>Terlama: ${point.max.toFixed(4)} detik · ${escapeHtml(point.peak.user)}</span><span>Schema teratas: ${escapeHtml(point.topSchema[0])} · ${point.topSchema[1].toLocaleString('id-ID')} query</span>${point.peak.path ? `<span>Path query terlama: ${escapeHtml(point.peak.path)}</span>` : ''}`);
  }
  function scatterData() {
    const points = [];
    state.rows.forEach((row) => {
      const query = number(row[fields.queryTime]), response = number(row[fields.responseTime]);
      if (query == null || response == null) return;
      points.push({ query, response, user: String(row[fields.user] || '(user kosong)'), path: String(row[fields.path] || ''), time: String(row[fields.time] || '') });
    });
    const step = Math.max(1, Math.ceil(points.length / 1800));
    return { points: points.filter((_, index) => index % step === 0), total: points.length };
  }
  function drawScatterChart(dataset) {
    const canvas = $('scatter-chart'), empty = $('scatter-empty'), rect = canvas.getBoundingClientRect(), points = dataset.points;
    if (!points.length || rect.width < 2) { canvas.hidden = true; empty.hidden = false; return; }
    canvas.hidden = false; empty.hidden = true;
    const ratio = window.devicePixelRatio || 1; canvas.width = Math.round(rect.width * ratio); canvas.height = Math.round(rect.height * ratio);
    const ctx = canvas.getContext('2d'); ctx.scale(ratio, ratio);
    const width = rect.width, height = rect.height, margin = { left: 47, right: 14, top: 14, bottom: 36 };
    const plotW = width - margin.left - margin.right, plotH = height - margin.top - margin.bottom;
    const maxX = Math.max(...points.map(point => point.query), 0.0001), maxY = Math.max(...points.map(point => point.response), 0.0001), hits = [];
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, width, height); ctx.font = '9px DM Sans, sans-serif'; ctx.textBaseline = 'middle';
    for (let i = 0; i <= 4; i++) {
      const x = margin.left + plotW * i / 4, y = margin.top + plotH * (1 - i / 4);
      ctx.strokeStyle = '#edf0f6'; ctx.beginPath(); ctx.moveTo(x, margin.top); ctx.lineTo(x, height - margin.bottom); ctx.stroke();
      ctx.beginPath(); ctx.moveTo(margin.left, y); ctx.lineTo(width - margin.right, y); ctx.stroke();
      ctx.fillStyle = '#98a3b5'; ctx.textAlign = 'center'; ctx.fillText(compact(maxX * i / 4), x, height - 24);
      ctx.textAlign = 'right'; ctx.fillText(compact(maxY * i / 4), margin.left - 7, y);
    }
    points.forEach((point) => {
      const x = margin.left + point.query / maxX * plotW, y = margin.top + plotH - point.response / maxY * plotH;
      ctx.beginPath(); ctx.arc(x, y, 3, 0, Math.PI * 2); ctx.fillStyle = '#4b9ce899'; ctx.fill();
      hits.push({ x: x - 5, y: y - 5, w: 10, h: 10, point });
    });
    ctx.fillStyle = '#758197'; ctx.textAlign = 'center'; ctx.fillText('Query time (detik)', margin.left + plotW / 2, height - 5);
    ctx.save(); ctx.translate(10, margin.top + plotH / 2); ctx.rotate(-Math.PI / 2); ctx.fillText('Response time (detik)', 0, 0); ctx.restore();
    bindTooltip(canvas, getTooltip(canvas), hits, (point) => `<strong>${escapeHtml(point.user)}</strong><span>Query time: ${point.query.toFixed(4)} detik</span><span>Response time: ${point.response.toFixed(4)} detik</span><span>${escapeHtml(point.path || 'Path kosong')}</span>`);
  }
  function userData() {
    const users = new Map();
    state.rows.forEach((row) => {
      const path = String(row[fields.path] || '').trim();
      if (!path) return;
      const user = String(row[fields.user] || '(user kosong)');
      if (!users.has(user)) users.set(user, new Map());
      const paths = users.get(user);
      paths.set(path, (paths.get(path) || 0) + 1);
    });
    return [...users.entries()].map(([label, paths]) => ({
      label, value: [...paths.values()].reduce((sum, count) => sum + count, 0),
      uniquePaths: paths.size,
      tooltipLines: [...paths.entries()].sort((a, b) => b[1] - a[1]).map(([path, count]) => `${path} · ${count} kemunculan`)
    })).sort((a, b) => b.value - a.value);
  }
  function schemaData() {
    const groups = new Map();
    state.rows.forEach((row) => {
      const value = number(row[fields.queryTime]);
      const name = String(row[fields.schema] || '(kosong)');
      if (value == null) return;
      if (!groups.has(name)) groups.set(name, []);
      groups.get(name).push(value);
    });
    return [...groups.entries()].map(([label, values]) => ({ label, value: average(values) })).sort((a, b) => b.value - a.value).slice(0, 7);
  }
  function statusData() {
    const counts = new Map();
    state.rows.forEach((row) => {
      const status = number(row[fields.status]);
      if (status != null) counts.set(String(Math.round(status)), (counts.get(String(Math.round(status))) || 0) + 1);
    });
    return [...counts.entries()].map(([label, value]) => ({ label, value })).sort((a, b) => a.label.localeCompare(b.label, undefined, { numeric: true })).slice(0, 8);
  }

  function drawChart(canvasId, emptyId, points, color, decimal, unit) {
    const canvas = $(canvasId), empty = $(emptyId), rect = canvas.getBoundingClientRect();
    if (!points.length || rect.width < 2) { canvas.hidden = true; empty.hidden = false; return; }
    canvas.hidden = false; empty.hidden = true;
    const tooltip = getTooltip(canvas);
    const ratio = window.devicePixelRatio || 1;
    canvas.width = Math.round(rect.width * ratio); canvas.height = Math.round(rect.height * ratio);
    const ctx = canvas.getContext('2d'); ctx.scale(ratio, ratio);
    const width = rect.width, height = rect.height;
    const margin = { top: 14, right: 9, bottom: 38, left: 42 };
    const plotW = width - margin.left - margin.right, plotH = height - margin.top - margin.bottom;
    const max = Math.max(...points.map((point) => point.value), 0.0001);
    const hits = [];
    ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, width, height);
    ctx.font = '10px DM Sans, sans-serif'; ctx.textBaseline = 'middle';
    for (let line = 0; line <= 4; line++) {
      const y = margin.top + plotH * line / 4;
      ctx.strokeStyle = '#edf0f6'; ctx.lineWidth = 1; ctx.beginPath(); ctx.moveTo(margin.left, y); ctx.lineTo(width - margin.right, y); ctx.stroke();
      const value = max * (1 - line / 4);
      ctx.fillStyle = '#98a3b5'; ctx.textAlign = 'right'; ctx.fillText(decimal ? compact(value) : Math.round(value).toLocaleString('id-ID'), margin.left - 8, y);
    }
    const slot = plotW / points.length, barW = Math.min(38, slot * .54);
    points.forEach((point, index) => {
      const barH = Math.max(2, point.value / max * plotH);
      const x = margin.left + index * slot + (slot - barW) / 2, y = margin.top + plotH - barH;
      roundedRect(ctx, x, y, barW, barH, Math.min(5, barW / 2)); ctx.fillStyle = color; ctx.fill();
      hits.push({ x, y, w: barW, h: barH, point });
      ctx.fillStyle = '#758197'; ctx.textAlign = 'center';
      const label = point.label.length > 10 ? `${point.label.slice(0, 9)}…` : point.label;
      ctx.fillText(label, x + barW / 2, height - 17);
    });
    bindTooltip(canvas, tooltip, hits, (point) => `<strong>${escapeHtml(point.label)}</strong><span>${decimal ? Number(point.value).toFixed(4) : point.value.toLocaleString('id-ID')} ${unit}</span>`);
  }

  function drawUserChart(points) {
    const canvas = $('user-chart'), empty = $('user-empty'), area = canvas.parentElement;
    canvas.pinnedPoint = null;
    if (canvas.chartTooltip) canvas.chartTooltip.hidden = true;
    if (!points.length) { canvas.hidden = true; canvas.style.height = '100%'; area.style.height = ''; empty.hidden = false; return; }
    canvas.hidden = false; empty.hidden = true;
    const rowHeight = 27;
    const height = Math.max(170, points.length * rowHeight + 28);
    canvas.style.height = `${height}px`;
    area.style.height = `${height}px`;
    const rect = canvas.getBoundingClientRect(), ratio = window.devicePixelRatio || 1;
    canvas.width = Math.round(rect.width * ratio); canvas.height = Math.round(rect.height * ratio);
    const ctx = canvas.getContext('2d'); ctx.setTransform(ratio, 0, 0, ratio, 0, 0);
    const width = rect.width, h = rect.height, margin = { left: 130, right: 44, top: 12, bottom: 21 };
    const rowH = (h - margin.top - margin.bottom) / points.length;
    const max = Math.max(...points.map((point) => point.value), 1), tooltip = getTooltip(canvas), hits = [];
    ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, width, h); ctx.font = '10px DM Sans, sans-serif'; ctx.textBaseline = 'middle';
    margin.right = Math.max(margin.right, ctx.measureText(max.toLocaleString('id-ID')).width + 18);
    const chartW = Math.max(40, width - margin.left - margin.right);
    for (let line = 0; line <= 4; line++) {
      const x = margin.left + chartW * line / 4;
      ctx.strokeStyle = '#edf0f6'; ctx.beginPath(); ctx.moveTo(x, margin.top); ctx.lineTo(x, h - margin.bottom); ctx.stroke();
      ctx.fillStyle = '#98a3b5'; ctx.textAlign = 'center'; ctx.fillText(Math.round(max * line / 4).toLocaleString('id-ID'), x, h - 9);
    }
    points.forEach((point, index) => {
      const y = margin.top + index * rowH, barH = Math.min(15, rowH * .62), barW = Math.max(2, point.value / max * chartW);
      ctx.fillStyle = '#667085'; ctx.textAlign = 'right'; ctx.fillText(point.label.length > 21 ? `${point.label.slice(0, 20)}…` : point.label, margin.left - 10, y + rowH / 2);
      roundedRect(ctx, margin.left, y + (rowH - barH) / 2, barW, barH, 5); ctx.fillStyle = '#6b5ce7'; ctx.fill();
      ctx.fillStyle = '#59667a'; ctx.textAlign = 'left'; ctx.fillText(point.value.toLocaleString('id-ID'), margin.left + barW + 7, y + rowH / 2);
      hits.push({ x: margin.left, y, w: chartW, h: rowH, point });
    });
    bindTooltip(canvas, tooltip, hits, (point, pinned) => `<strong>${escapeHtml(point.label)}</strong><span>${point.value.toLocaleString('id-ID')} total path · ${point.uniquePaths.toLocaleString('id-ID')} path unik</span><div class="tooltip-paths">${point.tooltipLines.slice(0, 8).map((line) => `<span>${escapeHtml(line)}</span>`).join('')}${point.tooltipLines.length > 8 ? `<span>+ ${point.tooltipLines.length - 8} path lainnya</span>` : ''}</div>${pinned ? `<button class="tooltip-export" type="button" data-export-path-user="${escapeHtml(point.label)}">Unduh Excel path</button>` : ''}`, true);
  }
  function getTooltip(canvas) {
    let tooltip = canvas.chartTooltip;
    if (!tooltip) { tooltip = document.createElement('div'); tooltip.className = 'chart-tooltip'; tooltip.hidden = true; document.body.append(tooltip); canvas.chartTooltip = tooltip; }
    return tooltip;
  }
  function bindTooltip(canvas, tooltip, hits, content, pinOnClick = false) {
    const findHit = (event) => {
      const rect = canvas.getBoundingClientRect(), x = event.clientX - rect.left, y = event.clientY - rect.top;
      return hits.find((item) => x >= item.x && x <= item.x + item.w && y >= item.y && y <= item.y + item.h);
    };
    const show = (point, event, pinned) => {
      tooltip.innerHTML = content(point, pinned); tooltip.hidden = false;
      tooltip.style.left = `${Math.max(8, Math.min(event.clientX + 14, window.innerWidth - tooltip.offsetWidth - 8))}px`;
      tooltip.style.top = `${Math.max(8, Math.min(event.clientY - 18, window.innerHeight - tooltip.offsetHeight - 8))}px`;
    };
    canvas.onmousemove = (event) => {
      if (canvas.pinnedPoint) return;
      const hit = findHit(event);
      if (!hit) { tooltip.hidden = true; canvas.style.cursor = 'default'; return; }
      show(hit.point, event, false); canvas.style.cursor = 'pointer';
    };
    canvas.onclick = pinOnClick ? (event) => {
      const hit = findHit(event);
      if (!hit) { canvas.pinnedPoint = null; tooltip.hidden = true; return; }
      if (canvas.pinnedPoint?.label === hit.point.label) { canvas.pinnedPoint = null; tooltip.hidden = true; return; }
      canvas.pinnedPoint = hit.point; show(hit.point, event, true);
    } : null;
    canvas.onmouseleave = () => { if (!canvas.pinnedPoint) tooltip.hidden = true; canvas.style.cursor = 'default'; };
  }
  function compact(value) { return value === 0 ? '0' : value >= 10 ? value.toFixed(0) : value >= 1 ? value.toFixed(1) : value >= 0.01 ? value.toFixed(2) : value.toFixed(4); }
  function roundedRect(ctx, x, y, w, h, r) { ctx.beginPath(); ctx.roundRect(x, y, w, h, [r, r, 0, 0]); }

  function filterRows(query) {
    const term = query.trim().toLocaleLowerCase('id-ID');
    state.filtered = state.rows.filter((row) => {
      const matchesUser = !state.selectedUser || String(row[fields.user] || '(user kosong)') === state.selectedUser;
      const matchesTerm = !term || [fields.time, fields.user, fields.schema, fields.queryTime, fields.status, fields.responseTime, fields.path, fields.referer, fields.query].some((key) => String(row[key] ?? '').toLocaleLowerCase('id-ID').includes(term));
      return matchesUser && matchesTerm;
    });
    state.page = 0; renderTable();
  }
  function populateUsers() {
    const select = $('user-filter'), selected = state.selectedUser;
    select.innerHTML = '<option value="">Semua user</option>';
    [...new Set(state.rows.map((row) => String(row[fields.user] || '(user kosong)')))].sort((a, b) => a.localeCompare(b)).forEach((user) => {
      const option = document.createElement('option'); option.value = user; option.textContent = user; select.append(option);
    });
    if ([...select.options].some((option) => option.value === selected)) { select.value = selected; } else { state.selectedUser = ''; }
  }
  function renderTable() {
    state.displayRows = state.uniqueMode ? uniqueRows(state.filtered) : state.filtered.map((row) => ({ ...row, __count: 1 }));
    const total = state.displayRows.length, pages = Math.max(1, Math.ceil(total / state.pageSize));
    state.page = Math.min(Math.max(state.page, 0), pages - 1);
    const start = state.page * state.pageSize;
    $('log-rows').innerHTML = state.displayRows.slice(start, start + state.pageSize).map((row) => {
      const status = number(row[fields.status]);
      const statusClass = status == null ? '' : status >= 500 ? 'status-error' : status >= 400 ? 'status-warn' : 'status-ok';
      return `<tr><td>${escapeHtml(display(row[fields.time]))}</td><td>${escapeHtml(display(row[fields.user]))}</td><td>${escapeHtml(display(row[fields.schema]))}</td><td>${Number(row.__count || 1).toLocaleString('id-ID')}</td><td>${escapeHtml(formatSeconds(number(row[fields.queryTime])))}</td><td>${status == null ? '—' : `<span class="status-pill ${statusClass}">${escapeHtml(Math.round(status))}</span>`}</td><td>${escapeHtml(formatSeconds(number(row[fields.responseTime])))}</td><td title="${escapeHtml(row[fields.path])}">${escapeHtml(display(row[fields.path]))}</td><td title="${escapeHtml(row[fields.referer])}">${escapeHtml(display(row[fields.referer]))}</td><td title="${escapeHtml(row[fields.query])}">${escapeHtml(display(row[fields.query]))}</td></tr>`;
    }).join('');
    const range = dateRangeText(state.filtered);
    const summary = state.uniqueMode
      ? `${total.toLocaleString('id-ID')} baris unik dari ${state.filtered.length.toLocaleString('id-ID')} baris`
      : `${total.toLocaleString('id-ID')} baris`;
    $('table-count').textContent = `${summary}${range ? ` - ${range}` : ''}`;
    $('page-info').textContent = total ? `${(start + 1).toLocaleString('id-ID')}–${Math.min(start + state.pageSize, total).toLocaleString('id-ID')} dari ${total.toLocaleString('id-ID')}` : 'Tidak ada baris yang cocok';
    $('previous-page').disabled = state.page === 0; $('next-page').disabled = state.page >= pages - 1;
  }
  function uniqueRows(rows) {
    const groups = new Map();
    rows.forEach((row) => {
      const identity = row[fields.path]
        ? [fields.user, fields.path]
        : [fields.user, fields.path, fields.time, fields.schema, fields.query, fields.queryTime, fields.status, fields.responseTime, fields.referer];
      const key = JSON.stringify(identity.map((field) => row[field] ?? ''));
      if (!groups.has(key)) groups.set(key, { row: { ...row, __count: 0 }, queryTotal: 0, queryCount: 0, responseTotal: 0, responseCount: 0, latest: parseLogDate(row[fields.time]) });
      const group = groups.get(key);
      group.row.__count++;
      const queryTime = number(row[fields.queryTime]), responseTime = number(row[fields.responseTime]), date = parseLogDate(row[fields.time]);
      if (queryTime != null) { group.queryTotal += queryTime; group.queryCount++; }
      if (responseTime != null) { group.responseTotal += responseTime; group.responseCount++; }
      if (date && (!group.latest || date > group.latest)) { group.latest = date; group.row[fields.time] = row[fields.time]; }
    });
    return [...groups.values()].map((group) => {
      if (group.queryCount) group.row[fields.queryTime] = group.queryTotal / group.queryCount;
      if (group.responseCount) group.row[fields.responseTime] = group.responseTotal / group.responseCount;
      return group.row;
    });
  }
  function dateRangeText(rows) {
    const dates = rows.map((row) => parseLogDate(row[fields.time])).filter(Boolean).sort((a, b) => a - b);
    if (!dates.length) return '';
    const format = (date) => new Intl.DateTimeFormat('id-ID', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Asia/Jakarta' }).format(date);
    const first = format(dates[0]), last = format(dates[dates.length - 1]);
    return first === last ? first : `${first} s.d. ${last}`;
  }
  function parseLogDate(value) {
    if (value instanceof Date) return Number.isNaN(value.getTime()) ? null : value;
    if (value == null || value === '') return null;
    const text = String(value).trim();
    const kibana = text.match(/^([A-Za-z]{3})\s+(\d{1,2}),\s*(\d{4})\s*@\s*(\d{2}:\d{2}:\d{2}(?:\.\d+)?)$/);
    const parsed = kibana ? new Date(`${kibana[1]} ${kibana[2]}, ${kibana[3]} ${kibana[4]} GMT+0700`) : new Date(text);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }
  function showNotice(message) { $('notice').textContent = message; $('notice').hidden = false; }
  function hideNotice() { $('notice').hidden = true; }
  function downloadChart(canvasId) {
    const canvas = $(canvasId);
    if (canvas.hidden || !canvas.width) return showNotice('Grafik belum memiliki data untuk diunduh.');
    const names = { 'user-chart': 'path-per-user', 'schema-chart': 'query-time-per-schema', 'status-chart': 'status-http' };
    const link = document.createElement('a'); link.href = canvas.toDataURL('image/png'); link.download = `${names[canvasId] || 'grafik'}.png`; link.click();
  }
  // ─── Query Classification ────────────────────────────────────────────────────
  function cleanQuery(queryText) {
    if (!queryText) return '';
    let q = String(queryText)
      .replace(/\/\*[\s\S]*?\*\//g, ' ')
      .replace(/--[^\r\n]*/g, ' ')
      .replace(/#[^\r\n]*/g, ' ')
      .trim();
    let prev;
    do {
      prev = q;
      q = q.replace(/^\s*(?:SET\s+[^;]+;|USE\s+[^;]+;)\s*/i, '').trim();
    } while (q !== prev);
    return q;
  }

  function parseMainTable(q) {
    const cleaned = cleanQuery(q);
    const fromMatch = cleaned.match(/\bFROM\s+(?!SELECT\b|\()\s*`?([a-zA-Z0-9_$]+)`?/i)
      || cleaned.match(/\bFROM\s+`?([a-zA-Z0-9_$]+)`?/i);
    const intoMatch = cleaned.match(/\bINTO\s+`?([a-zA-Z0-9_$]+)`?/i);
    const updateMatch = cleaned.match(/\bUPDATE\s+`?([a-zA-Z0-9_$]+)`?/i);
    const ddlMatch = cleaned.match(/^(?:CREATE|ALTER|DROP|TRUNCATE)\s+(?:TEMPORARY\s+)?(?:TABLE|VIEW)\s+(?:IF\s+(?:NOT\s+)?EXISTS\s+)?`?([a-zA-Z0-9_$]+)`?/i);
    return fromMatch?.[1] || intoMatch?.[1] || updateMatch?.[1] || ddlMatch?.[1] || null;
  }

  function getSelectPurpose(q, qUp) {
    if (/\bCOUNT\s*\(/i.test(q)) return 'Hitung jumlah data';
    if (/\b(SUM|AVG|MIN|MAX)\s*\(/i.test(q)) return 'Kalkulasi agregasi data';
    if (/\b(LIKE|REGEXP|RLIKE|MATCH)\b/i.test(qUp)) return 'Pencarian data';
    if (/\bLIMIT\s+(?:0\s*,\s*)?1\b/i.test(qUp)) return 'Ambil detail data';
    if (/\bGROUP\s+BY\b/i.test(qUp)) return 'Rekapitulasi data';
    if (/\bJOIN\b/i.test(qUp)) return 'Ambil data relasi';
    if (/\bLIMIT\b/i.test(qUp) || /\bORDER\s+BY\b/i.test(qUp)) return 'Ambil daftar data';
    if (/\bWHERE\b/i.test(qUp)) return 'Filter data';
    if (/^SELECT\s+(?:DISTINCT\s+)?\*\s+FROM\b/i.test(qUp)) return 'Ambil seluruh data';
    return 'Ambil data spesifik';
  }

  function classifyQuery(queryText) {
    if (!queryText || String(queryText).trim() === '') return { k1: 'Tidak Diketahui', k2: '-' };
    const cleaned = cleanQuery(queryText);
    const qUp = cleaned.toUpperCase().replace(/\s+/g, ' ');
    const table = parseMainTable(cleaned) || 'tabel tidak diketahui';
    const tableSuffix = ` '${table}'`;
    const clauses = ['JOIN', 'WHERE', 'LIKE', 'GROUP BY', 'HAVING', 'ORDER BY', 'LIMIT', 'DISTINCT'];
    const tags = clauses.filter((clause) => new RegExp(`\\b${clause.replace(' ', '\\s+')}\\b`).test(qUp));
    const describe = (purpose) => `${purpose}${tags.length ? ` [${tags.join(', ')}]` : ''}${tableSuffix}`;

    if (/^SELECT\b/.test(qUp)) {
      return { k1: 'SELECT', k2: describe(getSelectPurpose(cleaned, qUp)) };
    }
    if (/^INSERT\b/.test(qUp)) {
      const purpose = /\bON\s+DUPLICATE\s+KEY\b/.test(qUp) ? 'Tambah/perbarui data' : /\bINSERT\s+INTO\b.*\bSELECT\b/is.test(cleaned) ? 'Salin data' : 'Tambah data';
      return { k1: 'INSERT', k2: describe(purpose) };
    }
    if (/^UPDATE\b/.test(qUp)) return { k1: 'UPDATE', k2: describe('Perbarui data') };
    if (/^DELETE\b/.test(qUp)) return { k1: 'DELETE', k2: describe('Hapus data') };
    if (/^CALL\b/.test(qUp)) {
      const procMatch = cleaned.match(/CALL\s+`?(\w+)`?/i);
      return { k1: 'CALL', k2: `Jalankan prosedur (${procMatch?.[1] || '?'})` };
    }
    if (/^CREATE\b/.test(qUp)) {
      const type = qUp.includes('TABLE') ? 'tabel' : qUp.includes('INDEX') ? 'indeks' : qUp.includes('VIEW') ? 'view' : 'objek';
      return { k1: 'DDL', k2: describe(`Buat ${type}`) };
    }
    if (/^ALTER\b/.test(qUp)) return { k1: 'DDL', k2: describe('Ubah struktur tabel') };
    if (/^DROP\b/.test(qUp)) return { k1: 'DDL', k2: describe('Hapus struktur/tabel') };
    if (/^TRUNCATE\b/.test(qUp)) return { k1: 'DDL', k2: describe('Kosongkan data') };
    if (/^(SHOW|DESCRIBE|EXPLAIN)\b/.test(qUp)) return { k1: 'SHOW', k2: 'Tampilkan informasi database' };
    if (/^(SET|COMMIT|ROLLBACK|BEGIN|START\s+TRANSACTION)\b/.test(qUp)) return { k1: 'SESSION', k2: 'Perintah sesi/transaksi' };
    return { k1: 'Lainnya', k2: describe('Perintah lainnya') };
  }

  // ─── Excel Styling & Export Helpers ──────────────────────────────────────────
  const EXCEL_STYLES = {
    title: {
      fill: { fgColor: { rgb: '1E3A8A' } },
      font: { name: 'Calibri', sz: 14, bold: true, color: { rgb: 'FFFFFF' } },
      alignment: { horizontal: 'center', vertical: 'center' }
    },
    subtitle: {
      fill: { fgColor: { rgb: 'F1F5F9' } },
      font: { name: 'Calibri', sz: 10, italic: true, color: { rgb: '475569' } },
      alignment: { horizontal: 'center', vertical: 'center' }
    },
    section: {
      fill: { fgColor: { rgb: '1E293B' } },
      font: { name: 'Calibri', sz: 11, bold: true, color: { rgb: 'FFFFFF' } },
      alignment: { horizontal: 'left', vertical: 'center', indent: 1 }
    },
    th: {
      fill: { fgColor: { rgb: '2563EB' } },
      font: { name: 'Calibri', sz: 10, bold: true, color: { rgb: 'FFFFFF' } },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
      border: {
        top: { style: 'thin', color: { rgb: '1D4ED8' } },
        bottom: { style: 'medium', color: { rgb: '1E40AF' } },
        left: { style: 'thin', color: { rgb: '3B82F6' } },
        right: { style: 'thin', color: { rgb: '3B82F6' } }
      }
    },
    thDark: {
      fill: { fgColor: { rgb: '1E293B' } },
      font: { name: 'Calibri', sz: 10, bold: true, color: { rgb: 'FFFFFF' } },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
      border: {
        top: { style: 'thin', color: { rgb: '0F172A' } },
        bottom: { style: 'medium', color: { rgb: '0F172A' } },
        left: { style: 'thin', color: { rgb: '334155' } },
        right: { style: 'thin', color: { rgb: '334155' } }
      }
    },
    thPlain: {
      font: { name: 'Calibri', sz: 10, bold: true, color: { rgb: '000000' } },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
      border: {
        top: { style: 'thin', color: { rgb: 'CBD5E1' } },
        bottom: { style: 'medium', color: { rgb: '94A3B8' } },
        left: { style: 'thin', color: { rgb: 'CBD5E1' } },
        right: { style: 'thin', color: { rgb: 'CBD5E1' } }
      }
    },
    cellLeft: {
      font: { name: 'Calibri', sz: 10, color: { rgb: '1E293B' } },
      alignment: { horizontal: 'left', vertical: 'center' },
      border: {
        top: { style: 'thin', color: { rgb: 'E2E8F0' } },
        bottom: { style: 'thin', color: { rgb: 'E2E8F0' } },
        left: { style: 'thin', color: { rgb: 'E2E8F0' } },
        right: { style: 'thin', color: { rgb: 'E2E8F0' } }
      }
    },
    cellRight: {
      font: { name: 'Calibri', sz: 10, color: { rgb: '1E293B' } },
      alignment: { horizontal: 'right', vertical: 'center' },
      border: {
        top: { style: 'thin', color: { rgb: 'E2E8F0' } },
        bottom: { style: 'thin', color: { rgb: 'E2E8F0' } },
        left: { style: 'thin', color: { rgb: 'E2E8F0' } },
        right: { style: 'thin', color: { rgb: 'E2E8F0' } }
      }
    },
    cellCenter: {
      font: { name: 'Calibri', sz: 10, color: { rgb: '1E293B' } },
      alignment: { horizontal: 'center', vertical: 'center' },
      border: {
        top: { style: 'thin', color: { rgb: 'E2E8F0' } },
        bottom: { style: 'thin', color: { rgb: 'E2E8F0' } },
        left: { style: 'thin', color: { rgb: 'E2E8F0' } },
        right: { style: 'thin', color: { rgb: 'E2E8F0' } }
      }
    },
    badgeNormal: {
      fill: { fgColor: { rgb: 'DCFCE7' } },
      font: { name: 'Calibri', sz: 10, bold: true, color: { rgb: '166534' } },
      alignment: { horizontal: 'center', vertical: 'center' },
      border: {
        top: { style: 'thin', color: { rgb: 'BBF7D0' } },
        bottom: { style: 'thin', color: { rgb: 'BBF7D0' } },
        left: { style: 'thin', color: { rgb: 'BBF7D0' } },
        right: { style: 'thin', color: { rgb: 'BBF7D0' } }
      }
    },
    badgeSlow: {
      fill: { fgColor: { rgb: 'FEF3C7' } },
      font: { name: 'Calibri', sz: 10, bold: true, color: { rgb: '92400E' } },
      alignment: { horizontal: 'center', vertical: 'center' },
      border: {
        top: { style: 'thin', color: { rgb: 'FDE68A' } },
        bottom: { style: 'thin', color: { rgb: 'FDE68A' } },
        left: { style: 'thin', color: { rgb: 'FDE68A' } },
        right: { style: 'thin', color: { rgb: 'FDE68A' } }
      }
    },
    badgeVerySlow: {
      fill: { fgColor: { rgb: 'FEE2E2' } },
      font: { name: 'Calibri', sz: 10, bold: true, color: { rgb: '991B1B' } },
      alignment: { horizontal: 'center', vertical: 'center' },
      border: {
        top: { style: 'thin', color: { rgb: 'FECACA' } },
        bottom: { style: 'thin', color: { rgb: 'FECACA' } },
        left: { style: 'thin', color: { rgb: 'FECACA' } },
        right: { style: 'thin', color: { rgb: 'FECACA' } }
      }
    },
    recWarn: {
      fill: { fgColor: { rgb: 'FEF2F2' } },
      font: { name: 'Calibri', sz: 10, bold: true, color: { rgb: '991B1B' } },
      alignment: { horizontal: 'left', vertical: 'center' },
      border: {
        top: { style: 'thin', color: { rgb: 'FCA5A5' } },
        bottom: { style: 'thin', color: { rgb: 'FCA5A5' } },
        left: { style: 'medium', color: { rgb: 'DC2626' } },
        right: { style: 'thin', color: { rgb: 'FCA5A5' } }
      }
    },
    recInfo: {
      fill: { fgColor: { rgb: 'EFF6FF' } },
      font: { name: 'Calibri', sz: 10, color: { rgb: '1E40AF' } },
      alignment: { horizontal: 'left', vertical: 'center' },
      border: {
        top: { style: 'thin', color: { rgb: 'BFDBFE' } },
        bottom: { style: 'thin', color: { rgb: 'BFDBFE' } },
        left: { style: 'medium', color: { rgb: '2563EB' } },
        right: { style: 'thin', color: { rgb: 'BFDBFE' } }
      }
    }
  };

  function writeCell(ws, r, c, val, baseStyle, isEvenRow, numFmt) {
    const ref = XLSX.utils.encode_cell({ r, c });
    const cell = { v: val ?? '' };
    if (typeof val === 'number') {
      cell.t = 'n';
      if (numFmt) cell.z = numFmt;
    } else if (typeof val === 'boolean') {
      cell.t = 'b';
    } else {
      cell.t = 's';
      cell.v = String(val ?? '');
    }
    let s = baseStyle ? { ...baseStyle } : {};
    if (isEvenRow && !s.fill) {
      s.fill = { fgColor: { rgb: 'F8FAFC' } };
    }
    cell.s = s;
    ws[ref] = cell;
    return cell;
  }

  function addMerge(ws, sR, sC, eR, eC, style) {
    if (!ws['!merges']) ws['!merges'] = [];
    ws['!merges'].push({ s: { r: sR, c: sC }, e: { r: eR, c: eC } });
    if (style) {
      for (let r = sR; r <= eR; r++) {
        for (let c = sC; c <= eC; c++) {
          const ref = XLSX.utils.encode_cell({ r, c });
          if (!ws[ref]) ws[ref] = { t: 's', v: '' };
          ws[ref].s = { ...(ws[ref].s || {}), ...style };
        }
      }
    }
  }

  // ─── Analysis & Export ────────────────────────────────────────────────────────
  function buildSummarySheet(rows) {
    const queryTimes = rows.map((r) => number(r[fields.queryTime])).filter((v) => v != null);
    const responseTimes = rows.map((r) => number(r[fields.responseTime])).filter((v) => v != null);
    const statuses = rows.map((r) => number(r[fields.status])).filter((v) => v != null);
    const avg = (arr) => arr.length ? arr.reduce((s, v) => s + v, 0) / arr.length : null;
    const pct = (n, d) => d ? `${(n / d * 100).toFixed(1)}%` : '0%';
    const f4 = (v) => v != null ? Number(v.toFixed(4)) : null;

    const total = rows.length;
    const errors = statuses.filter((s) => s >= 400).length;
    const clientErr = statuses.filter((s) => s >= 400 && s < 500).length;
    const serverErr = statuses.filter((s) => s >= 500).length;
    const avgQt = avg(queryTimes), maxQt = queryTimes.length ? Math.max(...queryTimes) : null, minQt = queryTimes.length ? Math.min(...queryTimes) : null;
    const avgRt = avg(responseTimes), maxRt = responseTimes.length ? Math.max(...responseTimes) : null;
    const normal = queryTimes.filter((t) => t < 1).length;
    const slow = queryTimes.filter((t) => t >= 1 && t < 5).length;
    const verySlow = queryTimes.filter((t) => t >= 5).length;

    // Klasifikasi 1 distribution
    const k1Map = new Map();
    rows.forEach((r) => {
      const { k1 } = classifyQuery(r[fields.query]);
      k1Map.set(k1, (k1Map.get(k1) || 0) + 1);
    });
    const k1Ranked = [...k1Map.entries()].sort((a, b) => b[1] - a[1]);

    // Klasifikasi 2 distribution (top occurrences)
    const k2Map = new Map();
    rows.forEach((r) => {
      const { k2 } = classifyQuery(r[fields.query]);
      k2Map.set(k2, (k2Map.get(k2) || 0) + 1);
    });
    const k2Ranked = [...k2Map.entries()].sort((a, b) => b[1] - a[1]).slice(0, 15);

    // Top 10 user
    const userMap = new Map();
    rows.forEach((r) => {
      const user = String(r[fields.user] || '(user kosong)');
      const qt = number(r[fields.queryTime]);
      if (!userMap.has(user)) userMap.set(user, { count: 0, total: 0, max: 0 });
      const u = userMap.get(user); u.count++;
      if (qt != null) { u.total += qt; u.max = Math.max(u.max, qt); }
    });
    const userRanked = [...userMap.entries()].sort((a, b) => b[1].total - a[1].total).slice(0, 10);

    // Top 10 schema
    const schMap = new Map();
    rows.forEach((r) => {
      const sch = String(r[fields.schema] || '(kosong)');
      const qt = number(r[fields.queryTime]);
      if (!schMap.has(sch)) schMap.set(sch, { count: 0, total: 0, max: 0 });
      const s = schMap.get(sch); s.count++;
      if (qt != null) { s.total += qt; s.max = Math.max(s.max, qt); }
    });
    const schRanked = [...schMap.entries()].sort((a, b) => (b[1].total / (b[1].count || 1)) - (a[1].total / (a[1].count || 1))).slice(0, 10);

    // Status dist
    const stMap = new Map();
    statuses.forEach((s) => { const k = String(Math.round(s)); stMap.set(k, (stMap.get(k) || 0) + 1); });
    const stRanked = [...stMap.entries()].sort((a, b) => b[1] - a[1]);

    // Top 5 slowest
    const top5 = [...rows].filter((r) => number(r[fields.queryTime]) != null)
      .sort((a, b) => (number(b[fields.queryTime]) || 0) - (number(a[fields.queryTime]) || 0)).slice(0, 5)
      .map((r) => {
        const { k1, k2 } = classifyQuery(r[fields.query]);
        return [
          String(r[fields.time] || '-'), String(r[fields.user] || '-'), String(r[fields.schema] || '-'),
          f4(number(r[fields.queryTime])), k1, k2, String(r[fields.path] || '-')
        ];
      });

    const ws = {};
    const rowHeights = [];
    let curRow = 0;
    const ts = new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' });

    // Title banner
    writeCell(ws, curRow, 0, 'SLOW QUERY & ACCESS LOG ANALYZER — EXECUTIVE SUMMARY', EXCEL_STYLES.title, false);
    addMerge(ws, curRow, 0, curRow, 6, EXCEL_STYLES.title);
    rowHeights[curRow] = { hpt: 30 };
    curRow++;

    // Subtitle banner
    writeCell(ws, curRow, 0, `Dibuat: ${ts} WIB   |   Total Data Dianalisa: ${total.toLocaleString('id-ID')} baris`, EXCEL_STYLES.subtitle, false);
    addMerge(ws, curRow, 0, curRow, 6, EXCEL_STYLES.subtitle);
    rowHeights[curRow] = { hpt: 20 };
    curRow++;

    // Spacer
    rowHeights[curRow] = { hpt: 12 };
    curRow++;

    function renderSection(title, headers, rowsData, alignments, numFormats, badgeColIdx, badgeStyles) {
      writeCell(ws, curRow, 0, title, EXCEL_STYLES.section, false);
      addMerge(ws, curRow, 0, curRow, headers.length - 1, EXCEL_STYLES.section);
      rowHeights[curRow] = { hpt: 24 };
      curRow++;

      headers.forEach((h, c) => {
        writeCell(ws, curRow, c, h, EXCEL_STYLES.th, false);
      });
      rowHeights[curRow] = { hpt: 22 };
      curRow++;

      rowsData.forEach((row, rIdx) => {
        const isEven = rIdx % 2 === 1;
        rowHeights[curRow] = { hpt: 20 };
        row.forEach((val, c) => {
          let style = alignments?.[c] === 'right' ? EXCEL_STYLES.cellRight
            : alignments?.[c] === 'center' ? EXCEL_STYLES.cellCenter
              : EXCEL_STYLES.cellLeft;
          if (badgeColIdx === c && badgeStyles?.[rIdx]) {
            style = badgeStyles[rIdx];
          }
          const numFmt = numFormats?.[c];
          writeCell(ws, curRow, c, val, style, isEven, numFmt);
        });
        curRow++;
      });

      rowHeights[curRow] = { hpt: 12 };
      curRow++;
    }

    // 1. KPI
    renderSection(
      '1. RINGKASAN KINERJA SISTEM (KPI)',
      ['Indikator Metrik', 'Nilai', 'Satuan / Keterangan'],
      [
        ['Total Log Dianalisa', total, 'baris log'],
        ['Query Time — Rata-rata', avgQt, 'detik (MySQL)'],
        ['Query Time — Tercepat (Min)', minQt, 'detik (MySQL)'],
        ['Query Time — Terlama (Max)', maxQt, 'detik (MySQL)'],
        ['Response Time — Rata-rata', avgRt, 'detik (Nginx)'],
        ['Response Time — Tertinggi (Max)', maxRt, 'detik (Nginx)'],
        ['Total HTTP Error (4xx & 5xx)', errors, `${pct(errors, statuses.length)} dari total request`],
        ['HTTP Sukses (2xx)', statuses.filter((s) => s >= 200 && s < 300).length, `${pct(statuses.filter((s) => s >= 200 && s < 300).length, statuses.length)} dari total request`]
      ],
      ['left', 'right', 'left'],
      [null, '0.0000', null]
    );

    // 2. Kecepatan Query
    renderSection(
      '2. DISTRIBUSI KECEPATAN QUERY MYSQL',
      ['Kategori Kecepatan', 'Kriteria Ambang Batas', 'Jumlah Log', 'Persentase'],
      [
        ['Normal', '< 1 detik', normal, pct(normal, total)],
        ['Lambat', '1 – 5 detik', slow, pct(slow, total)],
        ['Sangat Lambat', '>= 5 detik', verySlow, pct(verySlow, total)]
      ],
      ['center', 'center', 'right', 'right'],
      [null, null, '#,##0', null],
      0,
      [EXCEL_STYLES.badgeNormal, EXCEL_STYLES.badgeSlow, EXCEL_STYLES.badgeVerySlow]
    );

    // 3. Klasifikasi 1
    renderSection(
      '3. DISTRIBUSI TIPE QUERY (KLASIFIKASI 1)',
      ['Tipe Query (SQL)', 'Jumlah Log', 'Persentase'],
      k1Ranked.map(([k, n]) => [k, n, pct(n, total)]),
      ['left', 'right', 'right'],
      [null, '#,##0', null]
    );

    // 4. Klasifikasi 2
    renderSection(
      '4. TOP 15 TUJUAN & PENGELOMPOKAN QUERY (KLASIFIKASI 2)',
      ['Pengelompokan & Tujuan Query', 'Jumlah Log', 'Persentase'],
      k2Ranked.map(([k, n]) => [k, n, pct(n, total)]),
      ['left', 'right', 'right'],
      [null, '#,##0', null]
    );

    // 5. Top 10 User
    renderSection(
      '5. TOP 10 USER MYSQL — TOTAL QUERY TIME',
      ['User MySQL', 'Jumlah Log', 'Total Qt (detik)', 'Rata-rata Qt (detik)', 'Maks Qt (detik)'],
      userRanked.map(([u, d]) => [u, d.count, f4(d.total), f4(d.count ? d.total / d.count : 0), f4(d.max)]),
      ['left', 'right', 'right', 'right', 'right'],
      [null, '#,##0', '0.0000', '0.0000', '0.0000']
    );

    // 6. Top 10 Schema
    renderSection(
      '6. TOP 10 SCHEMA DATABASE — RATA-RATA QUERY TIME',
      ['Schema Database', 'Jumlah Log', 'Total Qt (detik)', 'Rata-rata Qt (detik)', 'Maks Qt (detik)'],
      schRanked.map(([s, d]) => [s, d.count, f4(d.total), f4(d.count ? d.total / d.count : 0), f4(d.max)]),
      ['left', 'right', 'right', 'right', 'right'],
      [null, '#,##0', '0.0000', '0.0000', '0.0000']
    );

    // 7. HTTP Status
    const statusRows = [
      ...stRanked.map(([s, n]) => {
        const c = Number(s);
        const cat = c >= 500 ? 'Server Error (5xx)' : c >= 400 ? 'Client Error (4xx)' : c >= 300 ? 'Redirect (3xx)' : 'Sukses (2xx)';
        return [s, cat, n, pct(n, statuses.length)];
      }),
      ['4xx Total', 'Subtotal Client Error', clientErr, pct(clientErr, statuses.length)],
      ['5xx Total', 'Subtotal Server Error', serverErr, pct(serverErr, statuses.length)]
    ];
    renderSection(
      '7. DISTRIBUSI STATUS HTTP (NGINX)',
      ['HTTP Status', 'Kategori Respons', 'Jumlah Request', 'Persentase'],
      statusRows,
      ['center', 'left', 'right', 'right'],
      [null, null, '#,##0', null]
    );

    // 8. Top 5 Slowest
    renderSection(
      '8. 5 QUERY PALING LAMBAT (TOP BOTTLENECK)',
      ['Waktu Eksekusi', 'User MySQL', 'Schema', 'Query Time (det)', 'Tipe (K1)', 'Tujuan Query (K2)', 'Path Nginx'],
      top5,
      ['center', 'left', 'left', 'right', 'center', 'left', 'left'],
      [null, null, null, '0.0000', null, null, null]
    );

    // 9. Rekomendasi
    writeCell(ws, curRow, 0, '9. REKOMENDASI DAN TINDAK LANJUT', EXCEL_STYLES.section, false);
    addMerge(ws, curRow, 0, curRow, 6, EXCEL_STYLES.section);
    rowHeights[curRow] = { hpt: 24 };
    curRow++;

    const recs = [];
    if (verySlow > 0) {
      recs.push({ text: `[!] PERHATIAN: Ditemukan ${verySlow} query dengan durasi >= 5 detik. Prioritaskan optimasi query dan evaluasi pembuatan indeks komposit pada tabel terkait.`, warn: true });
    }
    if (errors > 0) {
      recs.push({ text: `[!] PERHATIAN: Terjadi ${errors} HTTP error (${pct(errors, statuses.length)}). Periksa korelasi waktu query lambat terhadap timeout atau crash server Nginx/PHP.`, warn: true });
    }
    if (k1Ranked[0]) {
      recs.push({ text: `[i] Operasi SQL dominan adalah ${k1Ranked[0][0]} (${k1Ranked[0][1]} log — ${pct(k1Ranked[0][1], total)}). Pastikan pola caching diterapkan untuk mengurangi beban pembacaan berulang.`, warn: false });
    }
    if (k2Ranked[0]) {
      recs.push({ text: `[i] Tujuan query terbanyak: ${k2Ranked[0][0]} (${k2Ranked[0][1]} log).`, warn: false });
    }
    if (userRanked[0]) {
      recs.push({ text: `[i] User database paling membebani: ${userRanked[0][0]} (total akumulasi query time: ${userRanked[0][1].total.toFixed(2)} detik).`, warn: false });
    }
    recs.push({ text: `[i] Lihat sheet "Detail Query" untuk melihat daftar lengkap waktu, path, referer, dan original_query.`, warn: false });

    recs.forEach((rec) => {
      rowHeights[curRow] = { hpt: 22 };
      const st = rec.warn ? EXCEL_STYLES.recWarn : EXCEL_STYLES.recInfo;
      writeCell(ws, curRow, 0, rec.text, st, false);
      addMerge(ws, curRow, 0, curRow, 6, st);
      curRow++;
    });

    ws['!rows'] = rowHeights;
    ws['!cols'] = [
      { wch: 48 },
      { wch: 22 },
      { wch: 28 },
      { wch: 22 },
      { wch: 20 },
      { wch: 55 },
      { wch: 42 }
    ];

    const range = { s: { r: 0, c: 0 }, e: { r: curRow - 1, c: 6 } };
    ws['!ref'] = XLSX.utils.encode_range(range);
    return ws;
  }

  function buildQueryDetailSheet(rows) {
    const ws = {};
    const headers = ['waktu', 'path', 'referer', 'original_query'];
    ws['!cols'] = [{ wch: 26 }, { wch: 45 }, { wch: 45 }, { wch: 95 }];
    const rowHeights = [{ hpt: 26 }];

    // 1. Header row (tanpa warna / no fill)
    headers.forEach((h, c) => {
      writeCell(ws, 0, c, h, EXCEL_STYLES.thPlain, false);
    });

    // 2. Data rows
    rows.forEach((r, idx) => {
      const rowIdx = idx + 1;
      rowHeights[rowIdx] = { hpt: 20 };

      writeCell(ws, rowIdx, 0, r[fields.time] ?? '', EXCEL_STYLES.cellCenter, false);
      writeCell(ws, rowIdx, 1, r[fields.path] ?? '', EXCEL_STYLES.cellLeft, false);
      writeCell(ws, rowIdx, 2, r[fields.referer] ?? '', EXCEL_STYLES.cellLeft, false);
      writeCell(ws, rowIdx, 3, r[fields.query] ?? '', {
        ...EXCEL_STYLES.cellLeft,
        alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
      }, false);
    });

    ws['!rows'] = rowHeights;
    const range = { s: { r: 0, c: 0 }, e: { r: rows.length, c: headers.length - 1 } };
    ws['!ref'] = XLSX.utils.encode_range(range);
    ws['!autofilter'] = { ref: ws['!ref'] };
    return ws;
  }

  function exportFilteredExcel() {
    if (!window.XLSX) return showNotice('Library Excel belum termuat. Periksa koneksi internet lalu muat ulang halaman.');
    if (!state.displayRows.length) return showNotice('Tidak ada baris hasil filter untuk diunduh.');
    const columns = [
      ['Waktu', fields.time, 24], ['User MySQL', fields.user, 22], ['Schema', fields.schema, 20],
      ['Jumlah', '__count', 12], ['Query Time (detik)', fields.queryTime, 18],
      ['HTTP Status', fields.status, 14], ['Response Time (detik)', fields.responseTime, 22],
      ['Path', fields.path, 42], ['Referer', fields.referer, 48],
      ['Klasifikasi 1', '__k1', 16], ['Klasifikasi 2 ', '__k2', 52]
    ];
    const ws1 = {};
    ws1['!cols'] = columns.map(([, , wch]) => ({ wch }));
    const rowHeights = [{ hpt: 24 }];

    columns.forEach(([title], c) => {
      writeCell(ws1, 0, c, title, EXCEL_STYLES.thPlain, false);
    });

    state.displayRows.forEach((row, idx) => {
      const rowIdx = idx + 1;
      rowHeights[rowIdx] = { hpt: 20 };

      columns.forEach(([, field], c) => {
        let val;
        let numFmt = null;
        let align = EXCEL_STYLES.cellLeft;

        if (field === '__count') {
          val = row.__count || 1;
          numFmt = '#,##0';
          align = EXCEL_STYLES.cellRight;
        } else if (field === '__k1') {
          val = classifyQuery(row[fields.query]).k1;
          align = EXCEL_STYLES.cellCenter;
        } else if (field === '__k2') {
          val = classifyQuery(row[fields.query]).k2;
          align = EXCEL_STYLES.cellLeft;
        } else if (field === fields.queryTime || field === fields.responseTime) {
          val = number(row[field]);
          numFmt = '0.0000';
          align = EXCEL_STYLES.cellRight;
        } else if (field === fields.status) {
          val = number(row[field]);
          numFmt = '0';
          align = EXCEL_STYLES.cellCenter;
        } else if (field === fields.time) {
          val = row[field] ?? '';
          align = EXCEL_STYLES.cellCenter;
        } else {
          val = row[field] ?? '';
          align = EXCEL_STYLES.cellLeft;
        }

        writeCell(ws1, rowIdx, c, val, align, false, numFmt);
      });
    });

    ws1['!rows'] = rowHeights;
    ws1['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: state.displayRows.length, c: columns.length - 1 } });
    ws1['!autofilter'] = { ref: ws1['!ref'] };

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, ws1, 'Merged Logs');
    XLSX.utils.book_append_sheet(workbook, buildSummarySheet(state.displayRows), 'Summary Report');
    XLSX.utils.book_append_sheet(workbook, buildQueryDetailSheet(state.displayRows), 'Detail Query');
    const date = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(workbook, `merged-logs-filtered-${date}.xlsx`);
    hideNotice();
  }
  function exportPathExcel(user) {
    if (!window.XLSX) return showNotice('Library Excel belum termuat. Periksa koneksi internet lalu muat ulang halaman.');
    const counts = new Map();
    state.rows.forEach((row) => {
      const rowUser = String(row[fields.user] || '(user kosong)');
      const path = String(row[fields.path] || '').trim();
      if (rowUser === user && path) counts.set(path, (counts.get(path) || 0) + 1);
    });
    const rows = [...counts.entries()].sort((a, b) => b[1] - a[1]);
    if (!rows.length) return showNotice('Tidak ada path untuk user tersebut.');

    const ws = {};
    ws['!cols'] = [{ wch: 58 }, { wch: 38 }];
    const rowHeights = [{ hpt: 24 }, { hpt: 12 }, { hpt: 22 }];

    writeCell(ws, 0, 0, 'User MySQL', EXCEL_STYLES.thDark, false);
    writeCell(ws, 0, 1, user, EXCEL_STYLES.cellLeft, false);

    writeCell(ws, 2, 0, 'Path Nginx', EXCEL_STYLES.th, false);
    writeCell(ws, 2, 1, 'Jumlah kemunculan pada slow query', EXCEL_STYLES.th, false);

    rows.forEach(([path, count], idx) => {
      const rowIdx = idx + 3;
      const isEven = idx % 2 === 1;
      rowHeights[rowIdx] = { hpt: 20 };
      writeCell(ws, rowIdx, 0, path, EXCEL_STYLES.cellLeft, isEven);
      writeCell(ws, rowIdx, 1, count, EXCEL_STYLES.cellRight, isEven, '#,##0');
    });

    ws['!rows'] = rowHeights;
    ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rows.length + 2, c: 1 } });
    ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 2, c: 0 }, e: { r: rows.length + 2, c: 1 } }) };

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, ws, 'Path per User');
    const safeUser = user.replace(/[^a-z0-9_-]+/gi, '-').replace(/^-|-$/g, '') || 'user';
    XLSX.writeFile(workbook, `paths-slow-query-${safeUser}.xlsx`);
  }
  function debounce(fn, delay) { let timer; return (...args) => { clearTimeout(timer); timer = setTimeout(() => fn(...args), delay); }; }
})();
