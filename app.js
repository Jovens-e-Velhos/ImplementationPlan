/* ===========================================================================
   Farol PMO — client-side engine
   A faithful JavaScript port of farol_pmo.py. Runs entirely in the browser:
   the user drops in an "Implementation Plan" .xlsx and this file parses it
   (via SheetJS) and renders the same traffic-light dashboard the Python
   script used to produce, with no server involved.
   =========================================================================== */

// ─── Stage weights (mirrors STAGE_WEIGHTS in farol_pmo.py) ──────────────────
const STAGE_WEIGHTS = [
  { label: "Initiation",                    weight: 0.05, min: 201, max: 299 },
  { label: "Design Mapeamento operacional", weight: 0.30, min: 401, max: 499 },
  { label: "Contrato",                      weight: 0.10, min: 301, max: 399 },
  { label: "Systems (execução)",            weight: 0.30, min: 503, max: 503 },
  { label: "Cadastros",                     weight: 0.10, min: 502, max: 502 },
  { label: "Go Live",                       weight: 0.10, min: 601, max: 699 },
  { label: "Evaluation",                    weight: 0.05, min: 701, max: 799 },
];

// Order matters: first matching substring wins (mirrors STATUS_SCORE dict walk)
const STATUS_SCORE = [
  ["finished", 1.0],
  ["on track", 0.5],
  ["not started", 0.0],
  ["postponed", 0.3],
  ["delayed", 0.3],
  ["overdue", 0.1],
  ["on hold", 0.0],
];

const COLOR_HEX = {
  green: "#22c55e",
  yellow: "#eab308",
  red: "#ef4444",
  blue: "#3b82f6",
  gray: "#94a3b8",
};

const COLOR_LABEL = {
  green: "On Track",
  yellow: "Atenção",
  red: "Atrasado",
  blue: "Concluído",
  gray: "Não Iniciado",
};

const SHEET_NAME = "Implementation Plan";
// Fixed column layout of the template (letters, not positional array indices,
// so the parser doesn't break if a sheet's used-range happens to start at
// column A vs column B in different exports).
const COLS = {
  task_id: "B",
  subtask_id: "C",
  task: "D",
  responsible: "E",
  start_date: "F",
  end_date: "G",
  days: "H",
  status: "I",
  remarks: "J",
};

function cellValue(ws, col, row) {
  const cell = ws[col + row];
  if (!cell) return null;
  return cell.v === undefined ? null : cell.v;
}

function asStr(v) {
  if (v === null || v === undefined) return "";
  return String(v).trim();
}

function asDate(v) {
  // Use duck-typing instead of `instanceof Date` so this still works if the
  // Date object ever crosses a realm boundary (e.g. testing harnesses).
  return v && typeof v.getTime === "function" && !isNaN(v.getTime()) ? v : null;
}

// ─── load_data() ─────────────────────────────────────────────────────────
function loadData(workbook) {
  const ws = workbook.Sheets[SHEET_NAME];
  if (!ws) {
    throw new Error(`A planilha não contém uma aba chamada "${SHEET_NAME}".`);
  }
  const range = XLSX.utils.decode_range(ws["!ref"]);
  const rows = [];
  // Data (and the header row, which gets filtered out below) starts at
  // Excel row 7 — matches pandas' data.iloc[6:] on a header=None read.
  for (let r = 7; r <= range.e.r + 1; r++) {
    const taskStr = asStr(cellValue(ws, COLS.task, r));
    if (taskStr === "") continue;
    if (/Task ID|Sub Task ID/.test(taskStr)) continue;

    const status_raw = asStr(cellValue(ws, COLS.status, r));
    rows.push({
      task_id: asStr(cellValue(ws, COLS.task_id, r)),
      subtask_id: asStr(cellValue(ws, COLS.subtask_id, r)),
      task: taskStr,
      responsible: asStr(cellValue(ws, COLS.responsible, r)),
      start_date: asDate(cellValue(ws, COLS.start_date, r)),
      end_date: asDate(cellValue(ws, COLS.end_date, r)),
      status_raw,
      status_norm: status_raw.toLowerCase().trim(),
      remarks: asStr(cellValue(ws, COLS.remarks, r)),
    });
  }
  return rows;
}

// ─── get_project_metadata() ─────────────────────────────────────────────
function getProjectMetadata(workbook) {
  let latest_update = "Não informado";
  let responsible = "Não informado";
  const ws = workbook.Sheets[SHEET_NAME];
  if (!ws) return { latest_update, responsible };

  const range = XLSX.utils.decode_range(ws["!ref"]);
  const maxCol = Math.min(range.e.c, XLSX.utils.decode_col("Z"));

  for (let r = 1; r <= 6; r++) {
    const rowVals = [];
    for (let c = 0; c <= maxCol; c++) {
      const col = XLSX.utils.encode_col(c);
      rowVals.push(cellValue(ws, col, r));
    }
    for (let i = 0; i < rowVals.length; i++) {
      const cell = rowVals[i];
      const cellStr = asStr(cell).toLowerCase();

      if (cellStr.startsWith("latest update")) {
        const raw = asStr(cell);
        if (raw.includes(":")) {
          const parts = raw.split(":");
          const after = parts.slice(1).join(":").trim();
          if (after !== "") {
            latest_update = tryFormatDate(after);
            continue;
          }
        }
        for (let k = i + 1; k < Math.min(i + 6, rowVals.length); k++) {
          const val = rowVals[k];
          if (val !== null && val !== undefined && asStr(val) !== "") {
            latest_update = asDate(val) ? fmtDate(val) : tryFormatDate(asStr(val));
            break;
          }
        }
      } else if (cellStr.startsWith("responsible")) {
        const raw = asStr(cell);
        if (raw.includes(":")) {
          const parts = raw.split(":");
          const after = parts.slice(1).join(":").trim();
          if (after !== "") {
            responsible = after;
            continue;
          }
        }
        for (let k = i + 1; k < Math.min(i + 8, rowVals.length); k++) {
          const val = rowVals[k];
          const valStr = asStr(val).toLowerCase();
          if (val !== null && val !== undefined && valStr !== "" && !valStr.includes("overdue")) {
            responsible = asStr(val);
            break;
          }
        }
      }
    }
  }
  return { latest_update, responsible };
}

function tryFormatDate(str) {
  const d = new Date(str);
  if (!isNaN(d.getTime())) return fmtDate(d);
  return str;
}

// ─── get_milestone_date() ───────────────────────────────────────────────
function getMilestoneDate(data, taskId) {
  const row = data.find((r) => r.task_id === String(taskId));
  if (row) return fmtDate(row.end_date);
  return "Não definido";
}

// ─── task_num() ──────────────────────────────────────────────────────────
function taskNum(row) {
  if (row.task_id) {
    const n = parseInt(row.task_id, 10);
    if (!isNaN(n)) return n;
  }
  if (row.subtask_id) {
    const n = parseInt(row.subtask_id.split(".")[0], 10);
    if (!isNaN(n)) return n;
  }
  return null;
}

// ─── status_color() ──────────────────────────────────────────────────────
function statusColor(statusNorm, endDate) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const daysDelta = (d) => {
    const dd = new Date(d);
    dd.setHours(0, 0, 0, 0);
    return Math.round((dd - today) / 86400000);
  };

  if (statusNorm.includes("finished")) return "blue";
  if (statusNorm.includes("on track")) {
    if (endDate) {
      const delta = daysDelta(endDate);
      if (delta < 0) return "red";
      if (delta <= 2) return "yellow";
    }
    return "green";
  }
  if (statusNorm.includes("overdue")) return "red";
  if (statusNorm.includes("delayed") || statusNorm.includes("postponed")) return "yellow";
  if (statusNorm.includes("not started") || statusNorm.includes("on hold")) return "gray";

  if (endDate) {
    const delta = daysDelta(endDate);
    if (delta < 0) return "red";
    if (delta <= 2) return "yellow";
  }
  return "gray";
}

// ─── pct_for_tasks() ─────────────────────────────────────────────────────
function pctForTasks(rows) {
  if (rows.length === 0) return 0.0;
  let sum = 0;
  for (const r of rows) {
    let score = 0.0;
    for (const [k, v] of STATUS_SCORE) {
      if (r.status_norm.includes(k)) {
        score = v;
        break;
      }
    }
    sum += score;
  }
  return Math.round((sum / rows.length) * 1000) / 10;
}

// ─── stage_dominant_color() ──────────────────────────────────────────────
function stageDominantColor(rows) {
  if (rows.length === 0) return "gray";
  const colors = rows.map((r) => statusColor(r.status_norm, r.end_date));
  const priority = ["red", "yellow", "green", "blue", "gray"];
  for (const c of priority) {
    if (colors.includes(c)) return c;
  }
  return "gray";
}

// ─── build_stages() ──────────────────────────────────────────────────────
function buildStages(data) {
  const withNum = data.map((r) => ({ ...r, _tnum: taskNum(r) }));
  const stages = [];
  for (const conf of STAGE_WEIGHTS) {
    const subset = withNum.filter(
      (r) => r._tnum !== null && r._tnum >= conf.min && r._tnum <= conf.max
    );
    const tasks = subset.map((r) => ({
      id: r.subtask_id ? r.subtask_id : r.task_id,
      is_subtask: !!r.subtask_id,
      name: r.task,
      responsible: r.responsible,
      start: r.start_date,
      end: r.end_date,
      status: r.status_raw,
      status_norm: r.status_norm,
      color: statusColor(r.status_norm, r.end_date),
      remarks: r.remarks,
    }));
    stages.push({
      label: conf.label,
      weight: conf.weight,
      pct: pctForTasks(subset),
      color: stageDominantColor(subset),
      tasks,
    });
  }
  return stages;
}

function overallPct(stages) {
  let s = 0;
  for (const st of stages) s += st.pct * st.weight;
  return Math.round(s * 10) / 10;
}

function getOverallColor(stages) {
  const colors = stages.map((s) => s.color);
  for (const c of ["red", "yellow", "green", "blue"]) {
    if (colors.includes(c)) return c;
  }
  return "gray";
}

function fmtDate(d) {
  if (!d) return "—";
  const dt = asDate(d) || new Date(d);
  if (isNaN(dt.getTime())) return "—";
  const dd = String(dt.getDate()).padStart(2, "0");
  const mm = String(dt.getMonth() + 1).padStart(2, "0");
  const yyyy = dt.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// ─── HTML builders (mirrors render_html() in farol_pmo.py) ──────────────
function dot(color, size = 14) {
  const hex = COLOR_HEX[color] || "#94a3b8";
  return `<span class="dot" style="background:${hex};width:${size}px;height:${size}px;box-shadow:0 0 6px ${hex}88"></span>`;
}

function buildFarol(activeColor) {
  const colors = ["gray", "red", "yellow", "green", "blue"];
  let lights = "";
  for (const c of colors) {
    const hex = COLOR_HEX[c] || "#94a3b8";
    const style =
      c === activeColor
        ? `background:${hex}; opacity:1; box-shadow:0 0 8px ${hex}`
        : `background:${hex};`;
    lights += `<div class="farol-luz" style="${style}"></div>`;
  }
  return `<div class="farol">${lights}</div>`;
}

function buildMacroFarol(activeColor) {
  const items = [
    ["blue", "Concluído"],
    ["green", "On track"],
    ["red", "Atrasado"],
    ["gray", "Não iniciado"],
  ];
  let html = "";
  for (const [c, label] of items) {
    const isActive = c === activeColor;
    const hex = COLOR_HEX[c] || "#94a3b8";
    const shadow = isActive ? `box-shadow: 0 0 8px ${hex};` : "";
    html += `<div class="mf-item ${isActive ? "active" : ""}"><div class="mf-light" style="background:${hex}; ${shadow}"></div><span>${label}</span></div>`;
  }
  return `<div class="macro-farol">${html}</div>`;
}

function renderReport(stages, filename, meta) {
  const overall = overallPct(stages);
  const overallColor = getOverallColor(stages);
  const todayStr = fmtDate(new Date());

  let allRows = "";
  stages.forEach((s, i) => {
    const hex = COLOR_HEX[s.color];

    let taskRows = "";
    for (const t of s.tasks) {
      const tc = COLOR_HEX[t.color] || "#94a3b8";
      const indent = t.is_subtask ? "padding-left:28px" : "";
      const prefix = t.is_subtask ? "└ " : "";
      const sr = t.is_subtask ? "subtask-row" : "";
      taskRows += `<tr class="task-row ${sr}" data-status-norm="${escapeHtml(t.status_norm)}">` +
        `<td style="${indent}">${dot(t.color, 10)} ${prefix}<code>${escapeHtml(t.id)}</code></td>` +
        `<td style="color: var(--text)">${escapeHtml(t.name)}</td>` +
        `<td class="resp-col">${escapeHtml(t.responsible)}</td>` +
        `<td class="date-col">${fmtDate(t.start)}</td>` +
        `<td class="date-col">${fmtDate(t.end)}</td>` +
        `<td><span class="status-badge" style="background:${tc}22;color:${tc};border:1px solid ${tc}44">${escapeHtml(t.status)}</span></td>` +
        `<td class="remarks-col">${escapeHtml(t.remarks)}</td>` +
        `</tr>`;
    }
    taskRows += `<tr class="no-match-row" style="display:none"><td colspan="7" class="no-match-cell">Nenhuma task com o status selecionado nesta etapa.</td></tr>`;

    allRows += `<tr class="stage-row" onclick="toggleDetail('detail-${i}')">` +
      `<td class="stage-name"><span class="expand-icon">▶</span>` +
      `<strong style="color: var(--text)">${escapeHtml(s.label)}</strong></td>` +
      `<td class="weight-col">${Math.round(s.weight * 100)}%</td>` +
      `<td><div class="pbar-wrap"><div class="pbar-fill" style="width:${s.pct}%;background:${hex}"></div></div>` +
      `<span class="pct-label">${s.pct}%</span></td>` +
      `<td class="status-col">${buildFarol(s.color)}` +
      `<span style="color:${hex};font-weight:600;font-size:12px;margin-left:8px">${COLOR_LABEL[s.color]}</span></td>` +
      `</tr>` +
      `<tr id="detail-${i}" class="detail-section" style="display:none">` +
      `<td colspan="4" style="padding:0"><div class="detail-inner">` +
      `<table class="inner-table"><thead><tr>` +
      `<th>ID</th><th>Task / Subtask</th><th>Responsável</th>` +
      `<th>Início</th><th>Fim</th><th>Status</th><th>Comentários</th>` +
      `</tr></thead><tbody>${taskRows}</tbody></table>` +
      `</div></td></tr>`;
  });

  document.getElementById("report-root").innerHTML = `
<div class="header">
  <img src="fotos/logo.png" alt="DSV" class="dsv-logo" onerror="this.style.display='none'">
  <h1>Farol PMO — Implementation Report</h1>
  <div class="sub">Arquivo: ${escapeHtml(filename)} &nbsp;·&nbsp; Gerado em: ${todayStr}</div>

  <div class="header-actions">
    <button onclick="toggleTheme()" class="theme-btn">Alternar Modo Claro/Escuro</button>
    <button onclick="window.exportOpenTasksPDF()" class="theme-btn theme-btn-accent">Exportar PDF (tasks em aberto)</button>
    <button onclick="window.resetUpload()" class="theme-btn">Carregar outra planilha</button>
  </div>

  <div class="project-info">
    <div class="info-item"><strong>Responsável</strong><span class="info-value">${escapeHtml(meta.responsible)}</span></div>
    <div class="info-item"><strong>Última Atualização</strong><span class="info-value">${escapeHtml(meta.latest_update)}</span></div>
    <div class="info-item"><strong>Go-Live (Wave 1)</strong><span class="info-value">${escapeHtml(meta.go_live)}</span></div>
    <div class="info-item"><strong>Project Closure</strong><span class="info-value">${escapeHtml(meta.closure)}</span></div>
  </div>

  <div class="header-metrics">
    <div class="gauge-wrap">
      <div class="gauge-ring" style="background: conic-gradient(var(--accent) calc(${overall}% * 3.6deg), var(--track-bg) 0deg);">
        <span class="gauge-pct">${overall}%</span>
      </div>
      <div class="gauge-label">Conclusão Geral</div>
    </div>

    <div class="macro-farol-wrap">
      ${buildMacroFarol(overallColor)}
      <div class="gauge-label">Status do Projeto</div>
    </div>
  </div>
</div>

<div id="status-filter-bar" class="filter-bar"></div>

<div class="card">
  <table>
    <thead>
      <tr>
        <th>Etapa</th>
        <th>Peso</th>
        <th>Progresso</th>
        <th>Status</th>
      </tr>
    </thead>
    <tbody>
      ${allRows}
    </tbody>
  </table>
</div>

<div class="footer">
  <span>Farol PMO · rodando 100% no navegador · ${todayStr}</span>
  <span>Pesos: Initiation 5% · Contrato 10% · Cadastros 10% · Go Live 10% · Design 30% · Systems 30% · Evaluation 5%</span>
</div>`;
}

// ─── Orchestration ────────────────────────────────────────────────────────
// Kept around after render so the status filter and the PDF export can
// re-read the same data without re-parsing the workbook.
let LAST_STAGES = null;
let LAST_META = null;
let LAST_FILENAME = null;

function processWorkbook(workbook, filename) {
  const data = loadData(workbook);
  const stages = buildStages(data);
  const { latest_update, responsible } = getProjectMetadata(workbook);
  const go_live = getMilestoneDate(data, "602");
  const closure = getMilestoneDate(data, "706");
  const meta = { latest_update, responsible, go_live, closure };

  LAST_STAGES = stages;
  LAST_META = meta;
  LAST_FILENAME = filename;

  renderReport(stages, filename, meta);
  renderStatusFilterBar(stages);
}

// ─── Status filter ─────────────────────────────────────────────────────
function collectStatuses(stages) {
  // Map keyed by normalized status so "On Track" / "on track" dedupe,
  // preserving the first-seen raw label (for display) and its color.
  const map = new Map();
  for (const s of stages) {
    for (const t of s.tasks) {
      if (!map.has(t.status_norm)) {
        map.set(t.status_norm, { raw: t.status || t.status_norm, color: t.color });
      }
    }
  }
  return map;
}

function renderStatusFilterBar(stages) {
  const bar = document.getElementById("status-filter-bar");
  if (!bar) return;
  const statuses = collectStatuses(stages);
  if (statuses.size === 0) {
    bar.innerHTML = "";
    return;
  }

  let chips = `<span class="filter-label">Filtrar por status:</span>`;
  chips += `<label class="filter-chip filter-chip-all"><input type="checkbox" class="status-filter-cb" data-all="1" checked> Todos</label>`;
  for (const [norm, { raw, color }] of statuses) {
    const hex = COLOR_HEX[color] || "#94a3b8";
    chips += `<label class="filter-chip" style="--chip-color:${hex}">` +
      `<input type="checkbox" class="status-filter-cb" data-status="${escapeHtml(norm)}" checked>` +
      `${dot(color, 9)} ${escapeHtml(raw)}</label>`;
  }
  bar.innerHTML = chips;

  const allCb = bar.querySelector('[data-all="1"]');
  const statusCbs = Array.from(bar.querySelectorAll("[data-status]"));

  allCb.addEventListener("change", () => {
    statusCbs.forEach((cb) => (cb.checked = allCb.checked));
    applyStatusFilter();
  });
  statusCbs.forEach((cb) =>
    cb.addEventListener("change", () => {
      allCb.checked = statusCbs.every((c) => c.checked);
      applyStatusFilter();
    })
  );
}

function applyStatusFilter() {
  const bar = document.getElementById("status-filter-bar");
  if (!bar) return;
  const statusCbs = Array.from(bar.querySelectorAll("[data-status]"));
  const checked = new Set(statusCbs.filter((cb) => cb.checked).map((cb) => cb.dataset.status));
  const isFullSet = checked.size === statusCbs.length;

  document.querySelectorAll(".detail-inner .inner-table tbody").forEach((tbody) => {
    let anyVisible = false;
    tbody.querySelectorAll(".task-row").forEach((row) => {
      const match = checked.has(row.dataset.statusNorm);
      row.style.display = match ? "" : "none";
      if (match) anyVisible = true;
    });
    const emptyRow = tbody.querySelector(".no-match-row");
    if (emptyRow) emptyRow.style.display = anyVisible ? "none" : "table-row";
  });

  // When a real filter is active, auto-expand every stage so the matching
  // tasks are visible without having to click each row open manually.
  if (!isFullSet) {
    document.querySelectorAll(".detail-section").forEach((el) => {
      el.style.display = "table-row";
      const row = el.previousElementSibling;
      if (row) row.classList.add("expanded");
    });
  }
}
window.applyStatusFilter = applyStatusFilter;

// ─── PDF export (tasks em aberto) ───────────────────────────────────────
function exportOpenTasksPDF() {
  if (!LAST_STAGES) return;
  const view = document.getElementById("print-view");
  if (!view) return;

  const today = fmtDate(new Date());
  let sections = "";
  let totalOpen = 0;

  for (const s of LAST_STAGES) {
    const open = s.tasks.filter((t) => !t.status_norm.includes("finished"));
    if (open.length === 0) continue;
    totalOpen += open.length;

    let rows = "";
    for (const t of open) {
      const hex = COLOR_HEX[t.color] || "#94a3b8";
      rows += `<tr>` +
        `<td>${escapeHtml(t.id)}</td>` +
        `<td>${escapeHtml(t.name)}</td>` +
        `<td>${escapeHtml(t.responsible)}</td>` +
        `<td>${fmtDate(t.start)}</td>` +
        `<td>${fmtDate(t.end)}</td>` +
        `<td><span class="pdf-status" style="border-color:${hex};color:${hex}">${escapeHtml(t.status)}</span></td>` +
        `</tr>`;
    }

    sections += `<h3 class="pdf-stage-title">${escapeHtml(s.label)}</h3>` +
      `<table class="pdf-table"><thead><tr>` +
      `<th>ID</th><th>Task</th><th>Responsável</th><th>Início</th><th>Fim</th><th>Status</th>` +
      `</tr></thead><tbody>${rows}</tbody></table>`;
  }

  if (totalOpen === 0) {
    sections = `<p class="pdf-empty">Não há tasks em aberto — todas as tasks estão marcadas como "Finished".</p>`;
  }

  view.innerHTML = `
    <div class="pdf-letterhead">
      <img src="fotos/logo.png" alt="DSV" class="pdf-logo" onerror="this.style.display='none'">
      <div class="pdf-titleblock">
        <h1>Farol PMO — Tasks em Aberto</h1>
        <div class="pdf-sub">Arquivo: ${escapeHtml(LAST_FILENAME || "")} &nbsp;·&nbsp; Gerado em: ${today} &nbsp;·&nbsp; Responsável: ${escapeHtml(LAST_META ? LAST_META.responsible : "")}</div>
      </div>
    </div>
    ${sections}
    <div class="pdf-footer">Farol PMO · gerado em ${today} · ${totalOpen} task(s) em aberto</div>
  `;

 document.body.classList.add("print-mode");

  const cleanup = () => {
    document.body.classList.remove("print-mode");
    window.removeEventListener("afterprint", cleanup);
    window.removeEventListener("focus", cleanup);
  };

  // Aguarda o navegador fechar a janela de impressão nativamente
  window.addEventListener("afterprint", cleanup);
  
  // Fallback de segurança: remove o modo de impressão se a janela principal recuperar o foco
  window.addEventListener("focus", cleanup);

  // Dá 100ms para o CSS ser aplicado antes de invocar a impressão
  setTimeout(() => {
    window.print();
  }, 100);
}
window.exportOpenTasksPDF = exportOpenTasksPDF;

function toggleDetail(id) {
  const el = document.getElementById(id);
  const row = el.previousElementSibling;
  const visible = el.style.display !== "none";
  el.style.display = visible ? "none" : "table-row";
  row.classList.toggle("expanded", !visible);
}

function toggleTheme() {
  document.body.classList.toggle("light-theme");
  try {
    localStorage.setItem(
      "farol-theme",
      document.body.classList.contains("light-theme") ? "light" : "dark"
    );
  } catch (e) {
    /* localStorage unavailable — ignore */
  }
}

window.toggleDetail = toggleDetail;
window.toggleTheme = toggleTheme;
window.processWorkbookFile = function (file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array", cellDates: true });
      processWorkbook(workbook, file.name);
      document.getElementById("upload-screen").style.display = "none";
      document.getElementById("report-root").style.display = "block";
      try {
        localStorage.setItem("farol-last-file-name", file.name);
      } catch (e2) {}
    } catch (err) {
      showUploadError(
        `Não consegui ler esse arquivo (${err.message}). Confirme que é o Implementation Plan .xlsx com a aba "Implementation Plan".`
      );
    }
  };
  reader.onerror = () => showUploadError("Falha ao ler o arquivo.");
  reader.readAsArrayBuffer(file);
};

function showUploadError(msg) {
  const el = document.getElementById("upload-error");
  el.textContent = msg;
  el.style.display = "block";
}

window.resetUpload = function () {
  document.getElementById("report-root").style.display = "none";
  document.getElementById("report-root").innerHTML = "";
  document.getElementById("upload-screen").style.display = "flex";
  const printView = document.getElementById("print-view");
  if (printView) printView.innerHTML = "";
  LAST_STAGES = null;
  LAST_META = null;
  LAST_FILENAME = null;
  const input = document.getElementById("file-input");
  if (input) input.value = "";
};
