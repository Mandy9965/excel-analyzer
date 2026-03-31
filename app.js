/* ============================================================
   看图说话器 — app.js
   ============================================================ */

/* ---------- 常量 ---------- */
const FRIDAY_BASE = 'https://aigc.sankuai.com/v1';
const FRIDAY_MODEL = 'claude-3-7-sonnet-20250219';
const MAX_ROWS_PREVIEW = 200;   // 表格最多显示行数
const MAX_ROWS_AI      = 400;   // 发给 AI 的最大行数

/* ---------- 状态 ---------- */
let workbook   = null;
let sheetNames = [];
let curSheet   = '';
let allData    = [];   // 当前 sheet 全量数据（含表头）
let headers    = [];
let rows       = [];
let chartInst  = null;

// 透视表状态
let pivotConfig = { rows: [], cols: [], vals: [] };
let pivotSaved  = false;   // 是否已保存透视配置用于 AI

// AI tab 状态
let aiTab = 'raw';   // 'raw' | 'pivot'

/* ---------- DOM 快捷 ---------- */
const $ = id => document.getElementById(id);

/* ============================================================
   初始化
   ============================================================ */
document.addEventListener('DOMContentLoaded', () => {
  initUpload();
  initModal();
  initChartTypeBtns();
  initPivot();
  initAiTabs();
  $('btnAnalyze').addEventListener('click', runAnalysis);
  $('btnReset').addEventListener('click', resetAll);
  $('btnSettings').addEventListener('click', () => openModal());
});

/* ============================================================
   上传 & 解析
   ============================================================ */
function initUpload() {
  const zone = $('uploadZone');
  const inp  = $('fileInput');

  zone.addEventListener('click', () => inp.click());
  inp.addEventListener('change', e => handleFile(e.target.files[0]));

  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
  zone.addEventListener('drop', e => {
    e.preventDefault(); zone.classList.remove('drag-over');
    handleFile(e.dataTransfer.files[0]);
  });
}

function handleFile(file) {
  if (!file) return;
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx','xls','csv'].includes(ext)) {
    showToast('仅支持 .xlsx / .xls / .csv 文件', 'err'); return;
  }
  const reader = new FileReader();
  reader.onload = e => {
    try {
      if (ext === 'csv') {
        workbook = XLSX.read(e.target.result, { type: 'string' });
      } else {
        workbook = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      }
      sheetNames = workbook.SheetNames;
      $('fileName').textContent = file.name;
      $('fileSize').textContent = formatBytes(file.size);
      renderSheetTabs();
      switchSheet(sheetNames[0]);
      $('uploadSection').classList.add('hidden');
      $('resultSection').classList.remove('hidden');
    } catch (err) {
      showToast('文件解析失败：' + err.message, 'err');
    }
  };
  if (ext === 'csv') reader.readAsText(file, 'UTF-8');
  else reader.readAsArrayBuffer(file);
}

function renderSheetTabs() {
  const wrap = $('sheetTabs');
  wrap.innerHTML = '';
  sheetNames.forEach(name => {
    const btn = document.createElement('button');
    btn.className = 'sheet-tab';
    btn.textContent = name;
    btn.addEventListener('click', () => switchSheet(name));
    wrap.appendChild(btn);
  });
}

function switchSheet(name) {
  curSheet = name;
  document.querySelectorAll('.sheet-tab').forEach(b => {
    b.classList.toggle('active', b.textContent === name);
  });
  const ws = workbook.Sheets[name];
  allData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  if (!allData.length) { showToast('该 Sheet 为空', 'warn'); return; }
  headers = allData[0].map(String);
  rows    = allData.slice(1).filter(r => r.some(c => c !== ''));
  updateStats();
  renderTable();
  resetPivot();
  renderDefaultChart();
}

/* ============================================================
   统计卡片（只保留行数、列数）
   ============================================================ */
function updateStats() {
  $('statRows').textContent = rows.length.toLocaleString();
  $('statCols').textContent = headers.length.toLocaleString();
}

/* ============================================================
   数据表格
   ============================================================ */
function renderTable() {
  const displayRows = rows.slice(0, MAX_ROWS_PREVIEW);
  const numCols = detectNumericCols(headers, rows);

  let th = '<tr><th class="idx">#</th>' +
    headers.map(h => `<th>${escHtml(h)}</th>`).join('') + '</tr>';

  let tb = displayRows.map((r, i) =>
    '<tr><td class="idx">' + (i + 1) + '</td>' +
    headers.map((_, ci) => {
      const v = r[ci] ?? '';
      const cls = numCols.has(ci) ? ' class="num"' : '';
      return `<td${cls} title="${escHtml(String(v))}">${escHtml(String(v))}</td>`;
    }).join('') + '</tr>'
  ).join('');

  $('tableHead').innerHTML = th;
  $('tableBody').innerHTML = tb;

  const badge = $('tableBadge');
  if (badge) badge.textContent = `前 ${displayRows.length} / ${rows.length} 行`;
}

/* ============================================================
   图表
   ============================================================ */
function setPivotMode(on) {
  const chartPanel = document.querySelector('.panel-chart');
  const tablePanel = document.querySelector('.panel-table');
  if (on) {
    chartPanel.classList.add('pivot-mode');
    tablePanel.classList.add('pivot-mode');
    $('pivotArea').style.display = '';
    renderPivotPool();
  } else {
    chartPanel.classList.remove('pivot-mode');
    tablePanel.classList.remove('pivot-mode');
    $('pivotArea').style.display = 'none';
  }
}

function initChartTypeBtns() {
  document.querySelectorAll('.chart-type-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.chart-type-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      const type = btn.dataset.type;
      if (type === 'pivot') {
        setPivotMode(true);
      } else {
        setPivotMode(false);
        renderDefaultChart(type);
      }
    });
  });
}

function renderDefaultChart(type) {
  const activeBtn = document.querySelector('.chart-type-btn.active');
  const rawType   = type || (activeBtn ? activeBtn.dataset.type : 'bar');
  const chartType = rawType === 'pivot' ? 'bar' : rawType;

  const numColsSet = detectNumericCols(headers, rows);
  const numCols = [...numColsSet];
  if (!numCols.length) { showToast('未检测到数值列，无法绘图', 'warn'); return; }

  const labelCol = headers.findIndex((_, i) => !numColsSet.has(i));
  const labels   = rows.slice(0, 30).map(r => String(r[labelCol >= 0 ? labelCol : 0] ?? ''));
  const valCol   = numCols[0];
  const data     = rows.slice(0, 30).map(r => parseFloat(r[valCol]) || 0);

  drawChart(chartType, labels, [{ label: headers[valCol], data }]);
}

function drawChart(type, labels, datasets) {
  const ctx = $('myChart').getContext('2d');
  if (chartInst) chartInst.destroy();

  const palette = [
    'rgba(79,110,247,.75)', 'rgba(34,197,94,.75)', 'rgba(245,158,11,.75)',
    'rgba(239,68,68,.75)',  'rgba(168,85,247,.75)', 'rgba(20,184,166,.75)',
  ];

  const coloredDatasets = datasets.map((ds, i) => ({
    ...ds,
    backgroundColor: type === 'line'
      ? palette[i % palette.length]
      : labels.map((_, j) => palette[j % palette.length]),
    borderColor: palette[i % palette.length],
    borderWidth: type === 'line' ? 2 : 1,
    fill: type === 'line' ? false : undefined,
    tension: .35,
    pointRadius: type === 'line' ? 3 : undefined,
  }));

  chartInst = new Chart(ctx, {
    type,
    data: { labels, datasets: coloredDatasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: datasets.length > 1, position: 'top' },
        tooltip: { mode: 'index', intersect: false },
      },
      scales: type === 'pie' || type === 'doughnut' ? {} : {
        x: { ticks: { maxRotation: 45, font: { size: 11 } } },
        y: { beginAtZero: true, ticks: { font: { size: 11 } } },
      },
    },
  });
}

/* ============================================================
   透视表
   ============================================================ */
function initPivot() {
  // 拖拽区域
  ['dropRow','dropCol','dropVal'].forEach(id => {
    const el = $(id);
    el.addEventListener('dragover', e => { e.preventDefault(); el.classList.add('drag-over'); });
    el.addEventListener('dragleave', () => el.classList.remove('drag-over'));
    el.addEventListener('drop', e => {
      e.preventDefault(); el.classList.remove('drag-over');
      const field = e.dataTransfer.getData('text/plain');
      const zone  = el.dataset.zone;
      addToPivotZone(field, zone);
    });
  });

  $('btnApplyPivot').addEventListener('click', applyPivotToChart);
  $('btnSavePivotForAI').addEventListener('click', savePivotForAI);
  $('btnResetPivot').addEventListener('click', () => { resetPivot(); renderPivotPool(); showToast('透视配置已重置'); });
}

function resetPivot() {
  pivotConfig = { rows: [], cols: [], vals: [] };
  pivotSaved  = false;
  updatePivotSavedBadge();
  renderPivotPool();
  renderPivotZones();
  $('pivotResult').innerHTML = '';
}

function renderPivotPool() {
  const pool = $('pivotPool');
  pool.innerHTML = '';
  headers.forEach(h => {
    pool.appendChild(makeChip(h, () => {
      // 点击字段池中的 chip 不做操作（拖拽才分配）
    }, false));
  });
  // 让 chip 可拖拽
  pool.querySelectorAll('.pivot-chip').forEach(chip => {
    chip.setAttribute('draggable', 'true');
    chip.addEventListener('dragstart', e => {
      e.dataTransfer.setData('text/plain', chip.dataset.field);
    });
  });
}

function renderPivotZones() {
  ['rows','cols','vals'].forEach(zone => {
    const el = $('drop' + capitalize(zone === 'rows' ? 'Row' : zone === 'cols' ? 'Col' : 'Val'));
    // 保留 drop 区域本身，只清除 chip
    el.querySelectorAll('.pivot-chip').forEach(c => c.remove());
    pivotConfig[zone].forEach(field => {
      const chip = makeChip(field, () => removeFromPivotZone(field, zone), true);
      el.appendChild(chip);
    });
  });
}

function makeChip(field, onRemove, showRemove) {
  const chip = document.createElement('span');
  chip.className = 'pivot-chip';
  chip.dataset.field = field;
  chip.setAttribute('draggable', 'true');
  chip.addEventListener('dragstart', e => {
    e.dataTransfer.setData('text/plain', field);
  });
  chip.textContent = field;
  if (showRemove) {
    const x = document.createElement('span');
    x.className = 'chip-remove'; x.textContent = '×';
    x.addEventListener('click', e => { e.stopPropagation(); onRemove(); });
    chip.appendChild(x);
  }
  return chip;
}

function addToPivotZone(field, zone) {
  const key = zone === 'row' ? 'rows' : zone === 'col' ? 'cols' : 'vals';
  if (!pivotConfig[key].includes(field)) {
    pivotConfig[key].push(field);
    renderPivotZones();
    // 有行字段和值字段时自动汇总并刷新图表
    if (pivotConfig.rows.length > 0 && pivotConfig.vals.length > 0) {
      applyPivotToChart();
    }
  }
}

function removeFromPivotZone(field, zone) {
  pivotConfig[zone] = pivotConfig[zone].filter(f => f !== field);
  renderPivotZones();
}

function capitalize(s) { return s.charAt(0).toUpperCase() + s.slice(1); }

/* 计算透视表 */
function computePivot() {
  const { rows: rFields, cols: cFields, vals: vFields } = pivotConfig;
  if (!rFields.length || !vFields.length) return null;

  const rIdxs = rFields.map(f => headers.indexOf(f));
  const cIdxs = cFields.map(f => headers.indexOf(f));
  const vIdxs = vFields.map(f => headers.indexOf(f));

  const map = new Map();
  const colKeys = new Set();

  rows.forEach(row => {
    const rKey = rIdxs.map(i => row[i] ?? '').join(' | ');
    const cKey = cIdxs.length ? cIdxs.map(i => row[i] ?? '').join(' | ') : '__total__';
    colKeys.add(cKey);
    if (!map.has(rKey)) map.set(rKey, new Map());
    const inner = map.get(rKey);
    vIdxs.forEach((vi, idx) => {
      const fullKey = cKey + '::' + vFields[idx];
      inner.set(fullKey, (inner.get(fullKey) || 0) + (parseFloat(row[vi]) || 0));
    });
  });

  const colArr = [...colKeys];
  const tableHeaders = [rFields.join(' + ')];
  colArr.forEach(ck => {
    vFields.forEach(vf => {
      tableHeaders.push(ck === '__total__' ? vf : `${ck} / ${vf}`);
    });
  });

  const tableRows = [];
  map.forEach((inner, rKey) => {
    const tr = [rKey];
    colArr.forEach(ck => {
      vFields.forEach((vf, idx) => {
        const fullKey = ck + '::' + vf;
        tr.push(inner.get(fullKey) || 0);
      });
    });
    tableRows.push(tr);
  });

  return { headers: tableHeaders, rows: tableRows };
}

function applyPivotToChart() {
  const result = computePivot();
  if (!result) { showToast('请至少配置「行」和「值」字段', 'warn'); return; }

  renderPivotResultTable(result);

  // 取前 30 行绘图，透视按钮本身不是合法 chart type，回退到 bar
  const labels = result.rows.slice(0, 30).map(r => String(r[0]));
  const datasets = result.headers.slice(1).map((h, i) => ({
    label: h,
    data: result.rows.slice(0, 30).map(r => parseFloat(r[i + 1]) || 0),
  }));
  const activeBtn = document.querySelector('.chart-type-btn.active');
  const rawType = activeBtn ? activeBtn.dataset.type : 'bar';
  const chartType = rawType === 'pivot' ? 'bar' : rawType;
  drawChart(chartType, labels, datasets);
  showToast('透视图表已更新', 'ok');
}

function renderPivotResultTable(result) {
  const wrap = $('pivotResult');
  let html = '<table class="pivot-table"><thead><tr>';
  result.headers.forEach(h => { html += `<th>${escHtml(h)}</th>`; });
  html += '</tr></thead><tbody>';
  result.rows.forEach(row => {
    html += '<tr>';
    row.forEach((cell, i) => {
      const isNum = i > 0 && typeof cell === 'number';
      html += `<td${isNum ? ' class="num"' : ''}>${isNum ? fmtNum(cell) : escHtml(String(cell))}</td>`;
    });
    html += '</tr>';
  });
  html += '</tbody></table>';
  wrap.innerHTML = html;
}

function savePivotForAI() {
  const result = computePivot();
  if (!result) { showToast('请先配置透视表字段', 'warn'); return; }
  pivotSaved = true;
  // 自动切换 AI tab 到「透视表分析」，确保点击 AI 分析时使用透视数据
  aiTab = 'pivot';
  document.querySelectorAll('.ai-tab').forEach(b => {
    b.classList.toggle('active', b.dataset.tab === 'pivot');
  });
  updatePivotSavedBadge();
  showToast('透视配置已保存，已切换到「透视表分析」模式', 'ok');
}

function updatePivotSavedBadge() {
  const badge = $('pivotSavedBadge');
  if (!badge) return;
  badge.style.display = pivotSaved ? 'inline-flex' : 'none';
}

/* ============================================================
   AI 双 Tab
   ============================================================ */
function initAiTabs() {
  document.querySelectorAll('.ai-tab').forEach(btn => {
    btn.addEventListener('click', () => {
      aiTab = btn.dataset.tab;
      document.querySelectorAll('.ai-tab').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
    });
  });
}

/* ============================================================
   AI 分析
   ============================================================ */
async function runAnalysis() {
  if (!rows.length) { showToast('请先上传数据文件', 'warn'); return; }

  const cfg = loadConfig();
  const FRIDAY_BASE_TRIM = FRIDAY_BASE.replace(/\/$/, '');
  const base  = cfg.base?.trim() || FRIDAY_BASE_TRIM;
  const key   = cfg.key?.trim();
  const model = cfg.model?.trim() || FRIDAY_MODEL;

  if (!key) { openModal(); showToast('请先配置 API Key', 'warn'); return; }

  const btn = $('btnAnalyze');
  btn.disabled = true; btn.textContent = '分析中…';

  const aiBody = $('aiBody');
  aiBody.innerHTML = `
    <div class="ai-loading">
      <div class="dots"><span></span><span></span><span></span></div>
      <span>AI 正在分析数据，请稍候…</span>
    </div>`;

  try {
    const prompt = buildPrompt();
    const url    = base.replace(/\/$/, '') + '/chat/completions';

    const resp = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${key}` },
      body: JSON.stringify({
        model,
        stream: true,
        messages: [
          { role: 'system', content: '你是一位专业的数据分析师，擅长从数据中发现业务洞察，给出具体、可落地的建议。' },
          { role: 'user',   content: prompt },
        ],
      }),
    });

    if (!resp.ok) {
      const errText = await resp.text();
      throw new Error(`HTTP ${resp.status}: ${errText}`);
    }

    let fullText = '';
    const reader  = resp.body.getReader();
    const decoder = new TextDecoder();
    aiBody.innerHTML = '<div class="ai-result" id="aiResult"></div>';
    const resultEl = $('aiResult');

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      const chunk = decoder.decode(value, { stream: true });
      chunk.split('\n').forEach(line => {
        if (!line.startsWith('data: ')) return;
        const raw = line.slice(6).trim();
        if (raw === '[DONE]') return;
        try {
          const parsed = JSON.parse(raw);
          // 兼容标准 OpenAI 格式和 FRIDAY 格式
          const delta = parsed.choices?.[0]?.delta?.content
                     ?? parsed.choices?.[0]?.message?.content
                     ?? parsed.content
                     ?? null;
          if (delta) { fullText += delta; resultEl.innerHTML = marked.parse(fullText); }
        } catch {}
      });
    }

    // 流结束后如果仍无内容，说明响应格式不匹配
    if (!fullText.trim()) {
      aiBody.innerHTML = `<div class="ai-empty">
        <div class="ai-empty-icon">⚠️</div>
        <p>收到响应但内容为空，可能是模型返回格式不匹配。</p>
        <p class="ai-empty-tip">请检查模型名称是否正确，或尝试切换其他模型</p>
      </div>`;
    }
  } catch (err) {
    aiBody.innerHTML = `<div class="ai-empty">
      <div class="ai-empty-icon">⚠️</div>
      <p>分析失败：${escHtml(err.message)}</p>
      <p class="ai-empty-tip">请检查 API Key 和网络连接</p>
    </div>`;
  } finally {
    btn.disabled = false; btn.textContent = '✨ AI 智能分析';
  }
}

/* ============================================================
   构建 Prompt
   ============================================================ */
function buildPrompt() {
  const usePivot = aiTab === 'pivot' && pivotSaved;

  /* 1. 识别文件主题 */
  const theme = guessTheme();

  /* 2. 数据摘要 */
  const summary = buildDataSummary();

  /* 3. 核心数据（原始 or 透视） */
  let dataSection = '';
  if (usePivot) {
    const result = computePivot();
    if (result && result.rows.length > 0) {
      dataSection = `\n## 透视表数据（已按配置聚合）\n${pivotToText(result)}\n`;
    } else {
      // 透视配置无效时 fallback 到原始数据
      dataSection = `\n## 原始数据样本（前 ${Math.min(rows.length, MAX_ROWS_AI)} 行）\n${rawDataToText()}\n`;
    }
  } else {
    dataSection = `\n## 原始数据样本（前 ${Math.min(rows.length, MAX_ROWS_AI)} 行）\n${rawDataToText()}\n`;
  }

  /* 4. 组装 Prompt */
  return `你是一位专业数据分析师。请对以下${theme}数据进行深度分析，输出结构化报告。

## 数据概况
- 文件主题：${theme}
- 数据行数：${rows.length.toLocaleString()} 行
- 数据列数：${headers.length} 列
- 字段列表：${headers.join('、')}
${summary}
${dataSection}

---

请按以下结构输出分析报告（使用 Markdown 格式）：

## 一、核心发现
（**必须具体**：列出 Top 3-5 个关键发现，每条需包含具体数字。例如：「商品 A 贡献 GMV 120 万元，占总 GMV 的 34%」「品类 B 的转化率为 8.2%，高于均值 3.1 个百分点」。禁止泛泛而谈。）

## 二、数据结论
（**以数据为主**：用数字说话，罗列关键指标的实际值、排名、占比、趋势。每条结论必须有具体数值支撑。）

## 三、行动建议
（**发散思考**：基于数据结论，给出 3-5 条可落地的业务建议。建议要有创意、有逻辑，不要只是重复数据，而是要说明"因此应该做什么、怎么做、预期效果是什么"。）

## 四、风险提示
（数据质量问题、异常值、需要进一步验证的假设。）`;
}

/* 猜测文件主题 */
function guessTheme() {
  const text = (curSheet + ' ' + headers.join(' ')).toLowerCase();
  if (/gmv|销售额|营业额|revenue|sales/.test(text)) return '销售/GMV';
  if (/订单|order/.test(text)) return '订单';
  if (/商品|sku|product|item/.test(text)) return '商品';
  if (/用户|user|customer|会员/.test(text)) return '用户';
  if (/曝光|点击|转化|ctr|cvr|uv|pv/.test(text)) return '流量/转化';
  if (/库存|inventory|stock/.test(text)) return '库存';
  if (/财务|利润|成本|profit|cost/.test(text)) return '财务';
  if (/活动|campaign|promotion/.test(text)) return '营销活动';
  return curSheet || '业务';
}

/* 数据摘要（数值列统计） */
function buildDataSummary() {
  const numCols = detectNumericCols(headers, rows);
  if (!numCols.size) return '';
  const lines = [];
  numCols.forEach(ci => {
    const vals = rows.map(r => parseFloat(r[ci])).filter(v => !isNaN(v));
    if (!vals.length) return;
    const sum = vals.reduce((a, b) => a + b, 0);
    const avg = sum / vals.length;
    const max = Math.max(...vals);
    const min = Math.min(...vals);
    lines.push(`- ${headers[ci]}：合计 ${fmtNum(sum)}，均值 ${fmtNum(avg)}，最大 ${fmtNum(max)}，最小 ${fmtNum(min)}`);
  });
  return lines.length ? '\n## 数值字段统计\n' + lines.join('\n') : '';
}

/* 原始数据转文本 */
function rawDataToText() {
  const sample = rows.slice(0, MAX_ROWS_AI);
  const lines  = [headers.join('\t')];
  sample.forEach(r => lines.push(headers.map((_, i) => r[i] ?? '').join('\t')));
  return '```\n' + lines.join('\n') + '\n```';
}

/* 透视表转文本 */
function pivotToText(result) {
  const lines = [result.headers.join('\t')];
  result.rows.forEach(r => lines.push(r.map(v => typeof v === 'number' ? fmtNum(v) : v).join('\t')));
  return '```\n' + lines.join('\n') + '\n```';
}

/* ============================================================
   配置 Modal
   ============================================================ */
function initModal() {
  $('modalClose').addEventListener('click', closeModal);
  $('modalMask').addEventListener('click', e => { if (e.target === $('modalMask')) closeModal(); });
  $('btnSaveConfig').addEventListener('click', saveConfig);
  $('btnCancelConfig').addEventListener('click', closeModal);

  // 回填已保存配置
  const cfg = loadConfig();
  if (cfg.key)   $('inputKey').value   = cfg.key;
  if (cfg.base)  $('inputBase').value  = cfg.base;
  if (cfg.model) $('inputModel').value = cfg.model;
}

function openModal() {
  $('modalMask').classList.add('open');
}
function closeModal() {
  $('modalMask').classList.remove('open');
}

function saveConfig() {
  const key   = $('inputKey').value.trim();
  const base  = $('inputBase').value.trim();
  const model = $('inputModel').value.trim();
  if (!key) { showToast('API Key 不能为空', 'err'); return; }
  localStorage.setItem('analyzer_cfg', JSON.stringify({ key, base, model }));
  closeModal();
  showToast('配置已保存', 'ok');
}

function loadConfig() {
  try { return JSON.parse(localStorage.getItem('analyzer_cfg') || '{}'); }
  catch { return {}; }
}

/* ============================================================
   重置
   ============================================================ */
function resetAll() {
  workbook = null; sheetNames = []; curSheet = '';
  allData = []; headers = []; rows = [];
  if (chartInst) { chartInst.destroy(); chartInst = null; }
  pivotConfig = { rows: [], cols: [], vals: [] };
  pivotSaved  = false;
  aiTab = 'raw';
  $('resultSection').classList.add('hidden');
  $('uploadSection').classList.remove('hidden');
  $('fileInput').value = '';
}

/* ============================================================
   工具函数
   ============================================================ */
function detectNumericCols(hdrs, dataRows) {
  const set = new Set();
  hdrs.forEach((_, ci) => {
    const sample = dataRows.slice(0, 50).map(r => r[ci]);
    const numCount = sample.filter(v => v !== '' && !isNaN(parseFloat(v))).length;
    if (numCount / Math.max(sample.length, 1) > 0.6) set.add(ci);
  });
  return set;
}

function fmtNum(n) {
  if (Math.abs(n) >= 1e8) return (n / 1e8).toFixed(2) + '亿';
  if (Math.abs(n) >= 1e4) return (n / 1e4).toFixed(2) + '万';
  return n.toLocaleString('zh-CN', { maximumFractionDigits: 2 });
}

function formatBytes(b) {
  if (b < 1024) return b + ' B';
  if (b < 1024 * 1024) return (b / 1024).toFixed(1) + ' KB';
  return (b / 1024 / 1024).toFixed(1) + ' MB';
}

function escHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

let toastTimer = null;
function showToast(msg, type = '') {
  const t = $('toast');
  t.textContent = msg;
  t.className = 'toast show' + (type ? ' ' + type : '');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { t.className = 'toast'; }, 2800);
}
