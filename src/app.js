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
  const escapeHtml = (value) => String(value ?? '').replace(/[&<>"']/g, (char) => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));

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
      tooltipLines: [...paths.entries()].sort((a,b) => b[1]-a[1]).map(([path,count]) => `${path} · ${count} kemunculan`)
    })).sort((a,b) => b.value-a.value);
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
    return [...groups.entries()].map(([label, values]) => ({ label, value: average(values) })).sort((a,b) => b.value-a.value).slice(0, 7);
  }
  function statusData() {
    const counts = new Map();
    state.rows.forEach((row) => {
      const status = number(row[fields.status]);
      if (status != null) counts.set(String(Math.round(status)), (counts.get(String(Math.round(status))) || 0) + 1);
    });
    return [...counts.entries()].map(([label, value]) => ({ label, value })).sort((a,b) => a.label.localeCompare(b.label, undefined, {numeric:true})).slice(0, 8);
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
    const margin = {top: 14, right: 9, bottom: 38, left: 42};
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
      hits.push({x, y, w:barW, h:barH, point});
      ctx.fillStyle = '#758197'; ctx.textAlign = 'center';
      const label = point.label.length > 10 ? `${point.label.slice(0,9)}…` : point.label;
      ctx.fillText(label, x + barW / 2, height - 17);
    });
    bindTooltip(canvas, tooltip, hits, (point) => `<strong>${escapeHtml(point.label)}</strong><span>${decimal ? Number(point.value).toFixed(4) : point.value.toLocaleString('id-ID')} ${unit}</span>`);
  }

  function drawUserChart(points) {
    const canvas = $('user-chart'), empty = $('user-empty'), area = canvas.parentElement;
    canvas.pinnedPoint = null;
    if(canvas.chartTooltip) canvas.chartTooltip.hidden = true;
    if (!points.length) { canvas.hidden = true; canvas.style.height = '100%'; area.style.height = ''; empty.hidden = false; return; }
    canvas.hidden = false; empty.hidden = true;
    const rowHeight = 27;
    const height = Math.max(170, points.length * rowHeight + 28);
    canvas.style.height = `${height}px`;
    area.style.height = `${height}px`;
    const rect = canvas.getBoundingClientRect(), ratio = window.devicePixelRatio || 1;
    canvas.width = Math.round(rect.width * ratio); canvas.height = Math.round(rect.height * ratio);
    const ctx = canvas.getContext('2d'); ctx.setTransform(ratio,0,0,ratio,0,0);
    const width = rect.width, h = rect.height, margin = {left:130,right:44,top:12,bottom:21};
    const rowH = (h-margin.top-margin.bottom)/points.length;
    const max = Math.max(...points.map((point) => point.value), 1), tooltip = getTooltip(canvas), hits = [];
    ctx.fillStyle='#ffffff';ctx.fillRect(0,0,width,h);ctx.font = '10px DM Sans, sans-serif'; ctx.textBaseline = 'middle';
    margin.right=Math.max(margin.right,ctx.measureText(max.toLocaleString('id-ID')).width+18);
    const chartW=Math.max(40,width-margin.left-margin.right);
    for (let line=0; line<=4; line++) {
      const x = margin.left + chartW*line/4;
      ctx.strokeStyle='#edf0f6';ctx.beginPath();ctx.moveTo(x,margin.top);ctx.lineTo(x,h-margin.bottom);ctx.stroke();
      ctx.fillStyle='#98a3b5';ctx.textAlign='center';ctx.fillText(Math.round(max*line/4).toLocaleString('id-ID'),x,h-9);
    }
    points.forEach((point,index) => {
      const y=margin.top+index*rowH, barH=Math.min(15,rowH*.62), barW=Math.max(2,point.value/max*chartW);
      ctx.fillStyle='#667085';ctx.textAlign='right';ctx.fillText(point.label.length>21?`${point.label.slice(0,20)}…`:point.label,margin.left-10,y+rowH/2);
      roundedRect(ctx,margin.left,y+(rowH-barH)/2,barW,barH,5);ctx.fillStyle='#6b5ce7';ctx.fill();
      ctx.fillStyle='#59667a';ctx.textAlign='left';ctx.fillText(point.value.toLocaleString('id-ID'),margin.left+barW+7,y+rowH/2);
      hits.push({x:margin.left,y,w:chartW,h:rowH,point});
    });
    bindTooltip(canvas,tooltip,hits,(point,pinned) => `<strong>${escapeHtml(point.label)}</strong><span>${point.value.toLocaleString('id-ID')} total path · ${point.uniquePaths.toLocaleString('id-ID')} path unik</span><div class="tooltip-paths">${point.tooltipLines.slice(0,8).map((line)=>`<span>${escapeHtml(line)}</span>`).join('')}${point.tooltipLines.length>8?`<span>+ ${point.tooltipLines.length-8} path lainnya</span>`:''}</div>${pinned?`<button class="tooltip-export" type="button" data-export-path-user="${escapeHtml(point.label)}">Unduh Excel path</button>`:''}`, true);
  }
  function getTooltip(canvas) {
    let tooltip=canvas.chartTooltip;
    if(!tooltip){tooltip=document.createElement('div');tooltip.className='chart-tooltip';tooltip.hidden=true;document.body.append(tooltip);canvas.chartTooltip=tooltip;}
    return tooltip;
  }
  function bindTooltip(canvas,tooltip,hits,content,pinOnClick=false) {
    const findHit=(event)=>{
      const rect=canvas.getBoundingClientRect(),x=event.clientX-rect.left,y=event.clientY-rect.top;
      return hits.find((item)=>x>=item.x&&x<=item.x+item.w&&y>=item.y&&y<=item.y+item.h);
    };
    const show=(point,event,pinned)=>{
      tooltip.innerHTML=content(point,pinned);tooltip.hidden=false;
      tooltip.style.left=`${Math.max(8,Math.min(event.clientX+14,window.innerWidth-tooltip.offsetWidth-8))}px`;
      tooltip.style.top=`${Math.max(8,Math.min(event.clientY-18,window.innerHeight-tooltip.offsetHeight-8))}px`;
    };
    canvas.onmousemove=(event)=>{
      if(canvas.pinnedPoint) return;
      const hit=findHit(event);
      if(!hit){tooltip.hidden=true;canvas.style.cursor='default';return;}
      show(hit.point,event,false);canvas.style.cursor='pointer';
    };
    canvas.onclick=pinOnClick?(event)=>{
      const hit=findHit(event);
      if(!hit){canvas.pinnedPoint=null;tooltip.hidden=true;return;}
      if(canvas.pinnedPoint?.label===hit.point.label){canvas.pinnedPoint=null;tooltip.hidden=true;return;}
      canvas.pinnedPoint=hit.point;show(hit.point,event,true);
    }:null;
    canvas.onmouseleave=()=>{if(!canvas.pinnedPoint)tooltip.hidden=true;canvas.style.cursor='default';};
  }
  function compact(value) { return value >= 10 ? value.toFixed(0) : value >= 1 ? value.toFixed(1) : value.toFixed(2); }
  function roundedRect(ctx,x,y,w,h,r) { ctx.beginPath(); ctx.roundRect(x,y,w,h,[r,r,0,0]); }

  function filterRows(query) {
    const term = query.trim().toLocaleLowerCase('id-ID');
    state.filtered = state.rows.filter((row) => {
      const matchesUser = !state.selectedUser || String(row[fields.user] || '(user kosong)') === state.selectedUser;
      const matchesTerm = !term || [fields.time,fields.user,fields.schema,fields.queryTime,fields.status,fields.responseTime,fields.path,fields.referer,fields.query].some((key) => String(row[key] ?? '').toLocaleLowerCase('id-ID').includes(term));
      return matchesUser && matchesTerm;
    });
    state.page = 0; renderTable();
  }
  function populateUsers() {
    const select=$('user-filter'), selected=state.selectedUser;
    select.innerHTML='<option value="">Semua user</option>';
    [...new Set(state.rows.map((row)=>String(row[fields.user]||'(user kosong)')))].sort((a,b)=>a.localeCompare(b)).forEach((user)=>{
      const option=document.createElement('option');option.value=user;option.textContent=user;select.append(option);
    });
    if([...select.options].some((option)=>option.value===selected)){select.value=selected;}else{state.selectedUser='';}
  }
  function renderTable() {
    state.displayRows = state.uniqueMode ? uniqueRows(state.filtered) : state.filtered.map((row) => ({...row,__count:1}));
    const total = state.displayRows.length, pages = Math.max(1, Math.ceil(total / state.pageSize));
    state.page = Math.min(Math.max(state.page, 0), pages - 1);
    const start = state.page * state.pageSize;
    $('log-rows').innerHTML = state.displayRows.slice(start, start + state.pageSize).map((row) => {
      const status = number(row[fields.status]);
      const statusClass = status == null ? '' : status >= 500 ? 'status-error' : status >= 400 ? 'status-warn' : 'status-ok';
      return `<tr><td>${escapeHtml(display(row[fields.time]))}</td><td>${escapeHtml(display(row[fields.user]))}</td><td>${escapeHtml(display(row[fields.schema]))}</td><td>${Number(row.__count||1).toLocaleString('id-ID')}</td><td>${escapeHtml(formatSeconds(number(row[fields.queryTime])))}</td><td>${status == null ? '—' : `<span class="status-pill ${statusClass}">${escapeHtml(Math.round(status))}</span>`}</td><td>${escapeHtml(formatSeconds(number(row[fields.responseTime])))}</td><td title="${escapeHtml(row[fields.path])}">${escapeHtml(display(row[fields.path]))}</td><td title="${escapeHtml(row[fields.referer])}">${escapeHtml(display(row[fields.referer]))}</td><td title="${escapeHtml(row[fields.query])}">${escapeHtml(display(row[fields.query]))}</td></tr>`;
    }).join('');
    const range = dateRangeText(state.filtered);
    const summary = state.uniqueMode
      ? `${total.toLocaleString('id-ID')} baris unik dari ${state.filtered.length.toLocaleString('id-ID')} baris`
      : `${total.toLocaleString('id-ID')} baris`;
    $('table-count').textContent = `${summary}${range ? ` - ${range}` : ''}`;
    $('page-info').textContent = total ? `${(start+1).toLocaleString('id-ID')}–${Math.min(start+state.pageSize,total).toLocaleString('id-ID')} dari ${total.toLocaleString('id-ID')}` : 'Tidak ada baris yang cocok';
    $('previous-page').disabled = state.page === 0; $('next-page').disabled = state.page >= pages - 1;
  }
  function uniqueRows(rows) {
    const groups = new Map();
    rows.forEach((row) => {
      const identity = row[fields.path]
        ? [fields.user,fields.path]
        : [fields.user,fields.path,fields.time,fields.schema,fields.query,fields.queryTime,fields.status,fields.responseTime,fields.referer];
      const key = JSON.stringify(identity.map((field) => row[field] ?? ''));
      if (!groups.has(key)) groups.set(key,{row:{...row,__count:0},queryTotal:0,queryCount:0,responseTotal:0,responseCount:0,latest:parseLogDate(row[fields.time])});
      const group=groups.get(key);
      group.row.__count++;
      const queryTime=number(row[fields.queryTime]),responseTime=number(row[fields.responseTime]),date=parseLogDate(row[fields.time]);
      if(queryTime!=null){group.queryTotal+=queryTime;group.queryCount++;}
      if(responseTime!=null){group.responseTotal+=responseTime;group.responseCount++;}
      if(date&&(!group.latest||date>group.latest)){group.latest=date;group.row[fields.time]=row[fields.time];}
    });
    return [...groups.values()].map((group)=>{
      if(group.queryCount) group.row[fields.queryTime]=group.queryTotal/group.queryCount;
      if(group.responseCount) group.row[fields.responseTime]=group.responseTotal/group.responseCount;
      return group.row;
    });
  }
  function dateRangeText(rows) {
    const dates=rows.map((row)=>parseLogDate(row[fields.time])).filter(Boolean).sort((a,b)=>a-b);
    if(!dates.length) return '';
    const format=(date)=>new Intl.DateTimeFormat('id-ID',{day:'numeric',month:'long',year:'numeric',timeZone:'Asia/Jakarta'}).format(date);
    const first=format(dates[0]),last=format(dates[dates.length-1]);
    return first===last?first:`${first} s.d. ${last}`;
  }
  function parseLogDate(value) {
    if(value instanceof Date) return Number.isNaN(value.getTime())?null:value;
    if(value==null||value==='') return null;
    const text=String(value).trim();
    const kibana=text.match(/^([A-Za-z]{3})\s+(\d{1,2}),\s*(\d{4})\s*@\s*(\d{2}:\d{2}:\d{2}(?:\.\d+)?)$/);
    const parsed=kibana?new Date(`${kibana[1]} ${kibana[2]}, ${kibana[3]} ${kibana[4]} GMT+0700`):new Date(text);
    return Number.isNaN(parsed.getTime())?null:parsed;
  }
  function showNotice(message) { $('notice').textContent = message; $('notice').hidden = false; }
  function hideNotice() { $('notice').hidden = true; }
  function downloadChart(canvasId) {
    const canvas=$(canvasId);
    if(canvas.hidden||!canvas.width) return showNotice('Grafik belum memiliki data untuk diunduh.');
    const names={'user-chart':'path-per-user','schema-chart':'query-time-per-schema','status-chart':'status-http'};
    const link=document.createElement('a');link.href=canvas.toDataURL('image/png');link.download=`${names[canvasId]||'grafik'}.png`;link.click();
  }
  function exportFilteredExcel() {
    if(!window.XLSX) return showNotice('Library Excel belum termuat. Periksa koneksi internet lalu muat ulang halaman.');
    if(!state.displayRows.length) return showNotice('Tidak ada baris hasil filter untuk diunduh.');
    const columns=[
      ['Waktu',fields.time,24],['User MySQL',fields.user,22],['Schema',fields.schema,20],['Jumlah','__count',12],['Query time (detik)',fields.queryTime,18],
      ['HTTP status',fields.status,14],['Response time (detik)',fields.responseTime,22],['Path',fields.path,42],['Referer',fields.referer,48],['Query',fields.query,60]
    ];
    const values=state.displayRows.map((row)=>columns.map(([,field])=>{
      if(field==='__count') return row.__count||1;
      if(field===fields.queryTime||field===fields.responseTime||field===fields.status) return number(row[field]);
      return row[field]??'';
    }));
    const worksheet=XLSX.utils.aoa_to_sheet([columns.map(([title])=>title),...values]);
    worksheet['!cols']=columns.map(([, ,wch])=>({wch}));
    worksheet['!rows']=[{hpt:24}];
    worksheet['!autofilter']={ref:XLSX.utils.encode_range({s:{r:0,c:0},e:{r:state.displayRows.length,c:columns.length-1}})};
    for(let row=1;row<=state.displayRows.length;row++){
      for(const column of [4,6]){const cell=worksheet[XLSX.utils.encode_cell({r:row,c:column})];if(cell)cell.z='0.0000';}
      for(const column of [3,5]){const cell=worksheet[XLSX.utils.encode_cell({r:row,c:column})];if(cell)cell.z='0';}
    }
    const workbook=XLSX.utils.book_new();XLSX.utils.book_append_sheet(workbook,worksheet,'Merged Logs');
    const date=new Date().toISOString().slice(0,10);
    XLSX.writeFile(workbook,`merged-logs-filtered-${date}.xlsx`);
    hideNotice();
  }
  function exportPathExcel(user) {
    if(!window.XLSX) return showNotice('Library Excel belum termuat. Periksa koneksi internet lalu muat ulang halaman.');
    const counts=new Map();
    state.rows.forEach((row)=>{
      const rowUser=String(row[fields.user]||'(user kosong)');
      const path=String(row[fields.path]||'').trim();
      if(rowUser===user&&path) counts.set(path,(counts.get(path)||0)+1);
    });
    const rows=[...counts.entries()].sort((a,b)=>b[1]-a[1]);
    if(!rows.length) return showNotice('Tidak ada path untuk user tersebut.');
    const worksheet=XLSX.utils.aoa_to_sheet([
      ['User MySQL',user],[],['Path Nginx','Jumlah kemunculan pada slow query'],...rows
    ]);
    worksheet['!cols']=[{wch:58},{wch:38}];
    worksheet['!rows']=[{hpt:22}, {hpt:8}, {hpt:22}];
    worksheet['!autofilter']={ref:XLSX.utils.encode_range({s:{r:2,c:0},e:{r:rows.length+2,c:1}})};
    const workbook=XLSX.utils.book_new();XLSX.utils.book_append_sheet(workbook,worksheet,'Path per User');
    const safeUser=user.replace(/[^a-z0-9_-]+/gi,'-').replace(/^-|-$/g,'')||'user';
    XLSX.writeFile(workbook,`paths-slow-query-${safeUser}.xlsx`);
  }
  function debounce(fn, delay) { let timer; return (...args) => { clearTimeout(timer); timer = setTimeout(() => fn(...args), delay); }; }
})();
