const fileInput = document.getElementById("fileInput");
const parseBtn = document.getElementById("parseBtn");
const statusEl = document.getElementById("status");
const sourceConfigPreviewEl = document.getElementById("sourceConfigPreview");
const sourceConfigListEl = document.getElementById("sourceConfigList");
const runCompareBtn = document.getElementById("runCompareBtn");
const chartsPreview = document.getElementById("chartsPreview");
const chartsRoot = document.getElementById("chartsRoot");
const resultMeta = document.getElementById("resultMeta");
const applyFirstConfigBtn = document.getElementById("applyFirstConfigBtn");
const resultViewModeEl = document.getElementById("resultViewMode");
const resultTopNEl = document.getElementById("resultTopN");
const resultWindowSizeEl = document.getElementById("resultWindowSize");
const resultWindowStartEl = document.getElementById("resultWindowStart");
const applyResultViewBtn = document.getElementById("applyResultViewBtn");

const state = {
  sources: [],
  latestResult: null,
  tinyChartObservers: [],
  parseIssues: [],
  resultView: {
    mode: "full",
    topN: 80,
    windowSize: 80,
    windowStart: 0
  },
  resultModalFocus: null
};

const CHART_COLORS = ["#007aff", "#34c759", "#ff9500", "#af52de", "#5ac8fa", "#ff2d55", "#32ade6", "#5856d6"];

bindModalEvents();
bootstrapEmptyStates();
initResultViewControls();

parseBtn.addEventListener("click", () => {
  parseFiles().catch((err) => setStatus(`解析失败：${err.message}`, true));
});

fileInput.addEventListener("change", () => {
  const files = Array.from(fileInput.files || []);
  if (!files.length) {
    setStatus("未选择文件，请先上传 CSV / Excel。", true);
    return;
  }
  const names = files.slice(0, 3).map((f) => f.name).join("，");
  const more = files.length > 3 ? ` 等 ${files.length - 3} 个` : "";
  setStatus(`已选择 ${files.length} 个文件：${names}${more}`);
});

runCompareBtn.addEventListener("click", () => {
  try {
    const chartTypes = getSelectedChartTypes();
    if (!chartTypes.length) throw new Error("请至少选择一种图表类型");

    const selectedSources = state.sources.filter((s) => s.selected);
    if (selectedSources.length < 2) throw new Error("请至少选择 2 个 sheet / 数据源参与对比");

    const normalized = normalizeSources(selectedSources);
    const { mergedRows, dimLabels, metricName } = buildMergedRows(normalized);
    const payload = { normalized, mergedRows, dimLabels, metricName, chartTypes };
    state.latestResult = payload;
    state.resultModalFocus = null;

    syncResultViewControlsToState();
    renderResults(payload);
    renderResultPreview(payload);
    setStatus(`生成完成：共 ${chartTypes.length} 类结果`);
  } catch (err) {
    setStatus(err.message, true);
  }
});

applyFirstConfigBtn.addEventListener("click", () => {
  const selected = state.sources.filter((s) => s.selected);
  if (!selected.length) {
    setStatus("没有已选数据源可复制配置", true);
    return;
  }
  const base = selected[0].config;
  selected.forEach((src, idx) => {
    if (idx === 0) return;
    src.config.dimensions = [...base.dimensions];
    src.config.metric = base.metric;
    src.config.agg = base.agg;
    src.config.filter = base.filter;
  });
  renderSourceConfigList();
  renderSourceConfigPreview();
  setStatus("已将第一个配置应用到其它已选数据源");
});

if (applyResultViewBtn) {
  applyResultViewBtn.addEventListener("click", () => {
    syncResultViewControlsToState();
    if (!state.latestResult) {
      setStatus("暂无结果可应用视图，请先生成对比", true);
      return;
    }
    renderResults(state.latestResult);
    const prevFocus = state.resultModalFocus;
    state.resultModalFocus = null;
    renderResultPreview(state.latestResult);
    state.resultModalFocus = prevFocus;
    setStatus("已应用结果视图设置");
  });
}

if (resultViewModeEl) {
  resultViewModeEl.addEventListener("change", () => {
    toggleResultViewParams(resultViewModeEl.value);
  });
}

function bootstrapEmptyStates() {
  renderSourceConfigPreview();
  chartsPreview.innerHTML = '<p class="mini-empty">暂无结果，生成对比后会显示多缩略卡片。</p>';
}

function bindModalEvents() {
  document.querySelectorAll("[data-modal-open]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const modalId = btn.getAttribute("data-modal-open");
      if (modalId === "resultModal") {
        state.resultModalFocus = null;
        if (state.latestResult) {
          renderResults(state.latestResult);
        }
      }
      openModal(modalId);
    });
  });
  document.querySelectorAll("[data-modal-close]").forEach((el) => {
    el.addEventListener("click", closeAllModals);
  });
}

function openResultModalWithFocus(focusType) {
  state.resultModalFocus = focusType || null;
  if (state.latestResult) {
    renderResults(state.latestResult);
  }
  openModal("resultModal");
}

function resolveFocusedChartTypes(chartTypes) {
  const focusType = state.resultModalFocus;
  if (!focusType) return chartTypes;
  return chartTypes.includes(focusType) ? [focusType] : chartTypes;
}

function focusTypeLabel(type) {
  const labels = {
    line: "折线图",
    bar: "柱状图",
    delta: "差值图",
    pct: "百分比图",
    table: "表格"
  };
  return labels[type] || type;
}

function openModal(modalId) {
  const modal = document.getElementById(modalId);
  if (!modal) return;
  modal.classList.remove("hidden");
  modal.setAttribute("aria-hidden", "false");

  if (modalId === "resultModal") {
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        resizeResultModalCharts();
      });
    });
  }
}

function closeAllModals() {
  document.querySelectorAll(".modal").forEach((modal) => {
    modal.classList.add("hidden");
    modal.setAttribute("aria-hidden", "true");
  });
}

function resizeResultModalCharts() {
  chartsRoot.querySelectorAll(".plot").forEach((plotDiv) => {
    try {
      Plotly.Plots.resize(plotDiv);
    } catch (_) {
      // Ignore transient resize errors during modal transition/reflow.
    }
  });
}

function syncResultViewControlsToState() {
  if (!resultViewModeEl) return;
  state.resultView.mode = resultViewModeEl.value || "full";
  state.resultView.topN = clampInt(resultTopNEl ? resultTopNEl.value : 80, 10, 5000, 80);
  state.resultView.windowSize = clampInt(resultWindowSizeEl ? resultWindowSizeEl.value : 80, 20, 5000, 80);
  state.resultView.windowStart = clampInt(resultWindowStartEl ? resultWindowStartEl.value : 0, 0, 999999, 0);
}

function initResultViewControls() {
  const mode = resultViewModeEl ? resultViewModeEl.value : "full";
  toggleResultViewParams(mode);
}

function toggleResultViewParams(mode) {
  const targets = document.querySelectorAll(".result-control-item[data-mode-visible]");
  targets.forEach((el) => {
    const visibleMode = el.getAttribute("data-mode-visible");
    const shouldShow = visibleMode === mode;
    el.classList.toggle("is-hidden", !shouldShow);

    const input = el.querySelector("input,select");
    if (input) {
      input.disabled = !shouldShow;
    }
  });
}

function setStatus(msg, isError = false) {
  statusEl.textContent = msg;
  statusEl.classList.toggle("error", isError);
}

function getSelectedChartTypes() {
  return Array.from(document.querySelectorAll(".chart-type:checked")).map((el) => el.value);
}

async function parseFiles() {
  const files = Array.from(fileInput.files || []);
  if (!files.length) {
    setStatus("请先选择至少一个文件", true);
    return;
  }
  setStatus("正在解析文件...");

  const parsedSources = [];
  const parseIssues = [];
  for (const file of files) {
    const lower = file.name.toLowerCase();
    try {
      if (lower.endsWith(".csv")) {
        const text = await file.text();
        const wb = XLSX.read(text, { type: "string" });
        const sheetName = wb.SheetNames[0] || "Sheet1";
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: null });
        const reason = getSheetInvalidReason(rows);
        if (reason) {
          parseIssues.push({ fileName: file.name, sheetName, reason });
        } else {
          parsedSources.push(buildSource(file.name, sheetName, rows));
        }
        continue;
      }
      if (lower.endsWith(".xlsx") || lower.endsWith(".xls")) {
        const arr = await file.arrayBuffer();
        const wb = XLSX.read(arr, { type: "array" });
        wb.SheetNames.forEach((sheetName) => {
          const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: null });
          const reason = getSheetInvalidReason(rows);
          if (reason) {
            parseIssues.push({ fileName: file.name, sheetName, reason });
          } else {
            parsedSources.push(buildSource(file.name, sheetName, rows));
          }
        });
        continue;
      }
      parseIssues.push({ fileName: file.name, sheetName: "-", reason: "不支持的文件类型（仅支持 csv/xlsx/xls）" });
    } catch (err) {
      parseIssues.push({ fileName: file.name, sheetName: "-", reason: `文件读取失败：${err.message || "未知错误"}` });
    }
  }

  state.sources = parsedSources;
  state.parseIssues = parseIssues;
  if (!state.sources.length) {
    renderParseFailureDetails(parseIssues);
    throw new Error(buildParseFailureMessage(parseIssues));
  }

  renderSourceConfigList();
  renderSourceConfigPreview();
  if (parseIssues.length) {
    setStatus(`解析完成：可用 ${state.sources.length} 个，跳过 ${parseIssues.length} 个异常 sheet`);
  } else {
    setStatus(`解析完成：共 ${state.sources.length} 个可用 sheet / 数据源`);
  }
}

function getSheetInvalidReason(rows) {
  if (!Array.isArray(rows) || !rows.length) return "未读取到数据行（可能只有表头或整表为空）";
  const fields = Object.keys(rows[0] || {});
  if (!fields.length) return "未识别到表头字段";
  return null;
}

function buildParseFailureMessage(issues) {
  const intro = "未解析到可用数据。";
  if (!issues || !issues.length) {
    return `${intro} 请检查文件是否有表头和数据行。`;
  }
  const top = issues.slice(0, 5).map((issue, idx) => `${idx + 1}. ${issue.fileName} / ${issue.sheetName}：${issue.reason}`).join("；");
  const more = issues.length > 5 ? `；另有 ${issues.length - 5} 项` : "";
  return `${intro} ${top}${more}`;
}

function renderParseFailureDetails(issues) {
  const safeIssues = issues && issues.length ? issues : [{ fileName: "-", sheetName: "-", reason: "未知错误" }];
  const rows = safeIssues
    .map(
      (issue) =>
        `<tr><td>${escapeHtml(issue.fileName)}</td><td>${escapeHtml(issue.sheetName)}</td><td>${escapeHtml(issue.reason)}</td></tr>`
    )
    .join("");
  const detailHtml = `
    <div class="source-item parse-issue-wrap">
      <div class="source-title">
        <div class="source-name">解析失败明细</div>
        <span class="tag">${safeIssues.length} 项</span>
      </div>
      <div class="table-wrap parse-issue-table-wrap">
        <table class="parse-issue-table">
          <thead><tr><th>文件</th><th>Sheet</th><th>原因</th></tr></thead>
          <tbody>${rows}</tbody>
        </table>
      </div>
    </div>
  `;
  sourceConfigPreviewEl.innerHTML = detailHtml;
  sourceConfigListEl.innerHTML = detailHtml;
}

function buildSource(fileName, sheetName, rows) {
  const fields = rows.length ? Object.keys(rows[0]) : [];
  const dimDefault = inferDimensionField(fields);
  const metricDefault = inferMetricField(rows, fields, dimDefault);
  return {
    id: crypto.randomUUID(),
    fileName,
    sheetName,
    name: `${fileName} / ${sheetName}`,
    rows,
    fields,
    selected: false,
    config: {
      dimensions: dimDefault ? [dimDefault] : [],
      metric: metricDefault || "",
      agg: "sum",
      filter: ""
    }
  };
}

function inferDimensionField(fields) {
  const preferred = ["date", "日期", "day", "month", "period", "时间", "时间段"];
  const lowerMap = fields.map((f) => ({ raw: f, low: String(f).toLowerCase() }));
  const hit = lowerMap.find((f) => preferred.some((p) => f.low.includes(p)));
  return hit ? hit.raw : fields[0];
}

function inferMetricField(rows, fields, dimField) {
  const candidates = fields.filter((f) => f !== dimField);
  const numeric = candidates.filter((f) => {
    let seen = 0;
    let valid = 0;
    for (const row of rows.slice(0, 80)) {
      const v = row[f];
      if (v === null || v === undefined || String(v).trim() === "") continue;
      seen += 1;
      if (!Number.isNaN(toNumber(v))) valid += 1;
    }
    return seen > 0 && valid / seen > 0.7;
  });
  return numeric[0] || candidates[0] || "";
}

function toNumber(v) {
  if (typeof v === "number") return v;
  const cleaned = String(v).replace(/,/g, "").trim();
  return Number(cleaned);
}

function renderSourceConfigPreview() {
  if (!state.sources.length) {
    sourceConfigPreviewEl.innerHTML = '<p class="mini-empty">暂无数据源，上传并解析后会显示缩略配置。</p>';
    return;
  }
  const selectedCount = state.sources.filter((s) => s.selected).length;
  const rows = state.sources
    .slice(0, 6)
    .map((s) => {
      const tag = s.selected ? "已选" : "未选";
      const dimMeta = s.config.dimensions && s.config.dimensions.length ? s.config.dimensions.join(" + ") : "-";
      const meta = `${dimMeta} / ${s.config.metric || "-"} / ${s.config.agg}`;
      return `<div class="source-item"><div class="source-title"><div class="source-name">${escapeHtml(s.name)}</div><span class="tag">${tag}</span></div><div class="field-block"><label>${escapeHtml(meta)}</label></div></div>`;
    })
    .join("");
  const more = state.sources.length > 6 ? `<p class="hint">还有 ${state.sources.length - 6} 个数据源，点击“展开”查看全部。</p>` : "";
  sourceConfigPreviewEl.innerHTML = `<p class="result-mini-title">共 ${state.sources.length} 个数据源，已选 ${selectedCount} 个</p>${rows}${more}`;
}

function renderSourceConfigList() {
  if (!state.sources.length) {
    sourceConfigListEl.innerHTML = '<p class="mini-empty">暂无数据源，请先上传并解析文件。</p>';
    return;
  }
  sourceConfigListEl.innerHTML = "";

  state.sources.forEach((src) => {
    const item = document.createElement("div");
    item.className = "source-item";
    item.innerHTML = `
      <div class="source-title">
        <div class="source-name">${escapeHtml(src.name)}</div>
        <label><input type="checkbox" data-role="selected" data-id="${src.id}" ${src.selected ? "checked" : ""}/> 参与对比</label>
      </div>
      <div class="source-fields">
        <div class="field-block">
          <label>维度字段（可多选）</label>
          <div class="dim-picker" data-id="${src.id}">
            <div class="dim-toolbar">
              <input type="text" data-role="dimension-search" data-id="${src.id}" placeholder="搜索字段..." class="dim-search" />
              <button type="button" class="ghost mini-action" data-role="dim-select-all" data-id="${src.id}">全选</button>
              <button type="button" class="ghost mini-action" data-role="dim-clear" data-id="${src.id}">清空</button>
            </div>
            <div class="dim-options" data-role="dimension-options" data-id="${src.id}">
              ${src.fields
                .map((f, i) => {
                  const checked = (src.config.dimensions || []).includes(f) ? "checked" : "";
                  const safeId = `dim_${src.id}_${i}`;
                  return `<label class="dim-option" data-field="${escapeAttr(String(f).toLowerCase())}">
                    <input type="checkbox" id="${safeId}" data-role="dimension-item" data-id="${src.id}" value="${escapeAttr(f)}" ${checked} />
                    <span>${escapeHtml(f)}</span>
                  </label>`;
                })
                .join("")}
            </div>
            <p class="hint dim-summary" data-role="dimension-summary" data-id="${src.id}">${buildDimensionSummaryText(src.config.dimensions || [])}</p>
          </div>
        </div>
        <div class="field-block">
          <label>聚合字段（Y 值）</label>
          <select data-role="metric" data-id="${src.id}">
            ${src.fields.map((f) => `<option value="${escapeAttr(f)}" ${src.config.metric === f ? "selected" : ""}>${escapeHtml(f)}</option>`).join("")}
          </select>
        </div>
        <div class="field-block">
          <label>聚合方式</label>
          <select data-role="agg" data-id="${src.id}">
            ${[
              ["sum", "sum 求和"],
              ["avg", "avg 平均值"],
              ["count", "count 计数"],
              ["min", "min 最小值"],
              ["max", "max 最大值"]
            ]
              .map(([v, label]) => `<option value="${v}" ${src.config.agg === v ? "selected" : ""}>${label}</option>`)
              .join("")}
          </select>
        </div>
        <div class="field-block">
          <label>筛选条件（可选，格式：字段=值）</label>
          <input type="text" data-role="filter" data-id="${src.id}" value="${escapeAttr(src.config.filter || "")}" placeholder="如：region=North" />
        </div>
      </div>
    `;
    sourceConfigListEl.appendChild(item);
  });

  sourceConfigListEl.querySelectorAll("input,select").forEach((el) => {
    el.addEventListener("change", syncSourceConfigFromUi);
    if (el.dataset.role === "filter" || el.dataset.role === "dimension-search") {
      el.addEventListener("input", syncSourceConfigFromUi);
    }
  });

  sourceConfigListEl.querySelectorAll("button[data-role]").forEach((btn) => {
    btn.addEventListener("click", syncSourceConfigFromUi);
  });
}

function syncSourceConfigFromUi(evt) {
  const el = evt.target;
  const src = state.sources.find((s) => s.id === el.dataset.id);
  if (!src) return;

  if (el.dataset.role === "selected") src.selected = el.checked;
  if (el.dataset.role === "metric") src.config.metric = el.value;
  if (el.dataset.role === "agg") src.config.agg = el.value;
  if (el.dataset.role === "filter") src.config.filter = el.value;

  if (el.dataset.role === "dimension-item") {
    src.config.dimensions = getCheckedDimensions(src.id);
    updateDimensionSummary(src.id, src.config.dimensions);
  }
  if (el.dataset.role === "dimension-search") {
    filterDimensionOptions(src.id, el.value);
  }
  if (el.dataset.role === "dim-select-all") {
    setAllDimensions(src.id, true);
    src.config.dimensions = getCheckedDimensions(src.id);
    updateDimensionSummary(src.id, src.config.dimensions);
  }
  if (el.dataset.role === "dim-clear") {
    setAllDimensions(src.id, false);
    src.config.dimensions = [];
    updateDimensionSummary(src.id, src.config.dimensions);
  }

  renderSourceConfigPreview();
}

function normalizeSources(sources) {
  return sources.map((src) => {
    if (!src.config.dimensions || !src.config.dimensions.length || !src.config.metric) {
      throw new Error(`请完整配置字段：${src.name}`);
    }
    const grouped = aggregateRows(
      src.rows,
      src.config.dimensions,
      src.config.metric,
      src.config.agg,
      parseFilter(src.config.filter)
    );
    return {
      id: src.id,
      name: src.name,
      dimensions: src.config.dimensions,
      metric: src.config.metric,
      agg: src.config.agg,
      grouped
    };
  });
}

function parseFilter(filterText) {
  const raw = String(filterText || "").trim();
  if (!raw) return null;
  const idx = raw.indexOf("=");
  if (idx < 1 || idx === raw.length - 1) throw new Error(`筛选条件格式错误：${raw}（正确示例：region=North）`);
  return { field: raw.slice(0, idx).trim(), value: raw.slice(idx + 1).trim() };
}

function aggregateRows(rows, dimensionFields, metricField, agg, filter) {
  const bucket = new Map();
  for (const row of rows) {
    if (filter && String(row[filter.field] ?? "") !== filter.value) continue;

    const dim = dimensionFields
      .map((field) => {
        const val = row[field];
        const safeVal = val === null || val === undefined || val === "" ? "(空值)" : String(val);
        return `${field}=${safeVal}`;
      })
      .join(" | ");

    const metricNum = toNumber(row[metricField]);
    const current = bucket.get(dim) || { count: 0, sum: 0, min: Infinity, max: -Infinity };
    current.count += 1;
    if (!Number.isNaN(metricNum)) {
      current.sum += metricNum;
      current.min = Math.min(current.min, metricNum);
      current.max = Math.max(current.max, metricNum);
    }
    bucket.set(dim, current);
  }

  const result = new Map();
  bucket.forEach((v, k) => {
    let val = 0;
    if (agg === "sum") val = v.sum;
    if (agg === "avg") val = v.count ? v.sum / v.count : 0;
    if (agg === "count") val = v.count;
    if (agg === "min") val = Number.isFinite(v.min) ? v.min : 0;
    if (agg === "max") val = Number.isFinite(v.max) ? v.max : 0;
    result.set(k, val);
  });
  return result;
}

function buildMergedRows(sources) {
  const dimSet = new Set();
  sources.forEach((s) => s.grouped.forEach((_, dim) => dimSet.add(dim)));
  const dimLabels = Array.from(dimSet).sort(sortDimensionValue);
  const mergedRows = dimLabels.map((dim) => {
    const row = { dimension: dim };
    sources.forEach((src) => {
      row[src.name] = src.grouped.get(dim) ?? 0;
    });
    return row;
  });
  return { mergedRows, dimLabels, metricName: sources[0].metric };
}

function renderResults(payload) {
  const { normalized, mergedRows, dimLabels, metricName, chartTypes } = payload;
  const viewed = prepareViewedData(mergedRows, dimLabels, normalized, state.resultView);
  const viewRows = viewed.rows;
  const viewDims = viewed.dimLabels;
  const focusedTypes = resolveFocusedChartTypes(chartTypes);

  chartsRoot.innerHTML = "";
  const baseName = normalized[0].name;
  const aggMode = normalized[0].agg;
  const focusSuffix = state.resultModalFocus ? ` | 仅看：${focusTypeLabel(state.resultModalFocus)}` : "";
  resultMeta.textContent = `${focusedTypes.length} 类结果 | 显示 ${viewDims.length}/${dimLabels.length} 维度 | ${metricName} | ${aggMode}${focusSuffix}`;

  if (focusedTypes.includes("line")) {
    const plotDiv = createChartCard(chartsRoot, "聚合趋势折线图", "line_chart");
    Plotly.newPlot(plotDiv, buildSeriesTrace(normalized, viewDims, viewRows, "scatter"), buildMainLayout(metricName, viewDims), {
      responsive: true
    });
  }

  if (focusedTypes.includes("bar")) {
    const plotDiv = createChartCard(chartsRoot, "聚合值柱状图", "bar_chart");
    Plotly.newPlot(
      plotDiv,
      buildSeriesTrace(normalized, viewDims, viewRows, "bar"),
      { ...buildMainLayout(metricName, viewDims), barmode: "group" },
      { responsive: true }
    );
  }

  if (focusedTypes.includes("delta") || focusedTypes.includes("pct")) {
    const deltaRows = buildDeltaRows(normalized, viewDims);
    if (focusedTypes.includes("delta")) {
      const plotDiv = createChartCard(chartsRoot, `与基准差值图（基准：${baseName}）`, "delta_chart");
      const traces = normalized.slice(1).map((s, idx) => ({
        x: toAxisLabels(viewDims),
        y: deltaRows.map((r) => r[`${s.name}__delta`]),
        mode: "lines+markers",
        type: "scatter",
        name: `${s.name} - ${baseName}`,
        marker: { color: CHART_COLORS[(idx + 1) % CHART_COLORS.length] },
        customdata: viewDims,
        hovertemplate: "维度=%{customdata}<br>值=%{y}<extra>%{fullData.name}</extra>"
      }));
      Plotly.newPlot(plotDiv, traces, buildMainLayout(metricName, viewDims), { responsive: true });
    }

    if (focusedTypes.includes("pct")) {
      const plotDiv = createChartCard(chartsRoot, `与基准百分比差异图（基准：${baseName}）`, "pct_chart");
      const traces = normalized.slice(1).map((s, idx) => ({
        x: toAxisLabels(viewDims),
        y: deltaRows.map((r) => r[`${s.name}__pct`]),
        mode: "lines+markers",
        type: "scatter",
        name: `${s.name} vs ${baseName}`,
        marker: { color: CHART_COLORS[(idx + 2) % CHART_COLORS.length] },
        customdata: viewDims,
        hovertemplate: "维度=%{customdata}<br>差异=%{y}%<extra>%{fullData.name}</extra>"
      }));
      Plotly.newPlot(
        plotDiv,
        traces,
        {
          ...buildMainLayout(metricName, viewDims),
          yaxis: {
            title: { text: "差异 (%)", standoff: 18 },
            ticksuffix: "%",
            automargin: true,
            zeroline: true,
            zerolinecolor: "#dbe0ea"
          }
        },
        { responsive: true }
      );
    }
  }

  if (focusedTypes.includes("table")) {
    createResultTable(chartsRoot, viewRows, normalized);
  }
}

function renderResultPreview(payload) {
  const { normalized, mergedRows, dimLabels, metricName, chartTypes } = payload;
  const viewed = prepareViewedData(mergedRows, dimLabels, normalized, state.resultView);
  const viewRows = viewed.rows;
  const viewDims = viewed.dimLabels;
  const focusedTypes = resolveFocusedChartTypes(chartTypes);

  cleanupTinyPreviewArtifacts();
  chartsPreview.innerHTML = "";

  if (!normalized.length) {
    chartsPreview.innerHTML = '<p class="mini-empty">暂无结果，生成对比后会显示多缩略卡片。</p>';
    return;
  }

  const title = document.createElement("p");
  title.className = "result-mini-title";
  const previewSuffix = state.resultModalFocus ? `，当前单图：${focusTypeLabel(state.resultModalFocus)}` : "";
  title.textContent = `缩略总览：已生成 ${chartTypes.length} 类结果（显示 ${viewDims.length}/${dimLabels.length} 维度${previewSuffix}）`;
  chartsPreview.appendChild(title);

  const grid = document.createElement("div");
  grid.className = "preview-grid";
  chartsPreview.appendChild(grid);

  if (focusedTypes.includes("line")) {
    const host = createPreviewTile(grid, "折线图", "聚合趋势", "line");
    renderTinyChart(host, buildSeriesTrace(normalized, viewDims, viewRows, "scatter"), metricName);
  }
  if (focusedTypes.includes("bar")) {
    const host = createPreviewTile(grid, "柱状图", "聚合值对比", "bar");
    renderTinyChart(host, buildSeriesTrace(normalized, viewDims, viewRows, "bar"), metricName);
  }
  if (focusedTypes.includes("delta") || focusedTypes.includes("pct")) {
    const deltaRows = buildDeltaRows(normalized, viewDims);
    if (focusedTypes.includes("delta")) {
      const host = createPreviewTile(grid, "差值图", "与基准差值", "delta");
      const traces = normalized.slice(1).map((s, idx) => ({
        x: toAxisLabels(viewDims),
        y: deltaRows.map((r) => r[`${s.name}__delta`]),
        type: "scatter",
        mode: "lines",
        name: s.name,
        marker: { color: CHART_COLORS[(idx + 1) % CHART_COLORS.length] }
      }));
      renderTinyChart(host, traces, "delta");
    }
    if (focusedTypes.includes("pct")) {
      const host = createPreviewTile(grid, "百分比图", "与基准百分比", "pct");
      const traces = normalized.slice(1).map((s, idx) => ({
        x: toAxisLabels(viewDims),
        y: deltaRows.map((r) => r[`${s.name}__pct`]),
        type: "scatter",
        mode: "lines",
        name: s.name,
        marker: { color: CHART_COLORS[(idx + 2) % CHART_COLORS.length] }
      }));
      renderTinyChart(host, traces, "%");
    }
  }
  if (focusedTypes.includes("table")) {
    renderTinyTable(grid, viewRows, normalized, "table");
  }
}

function createPreviewTile(grid, title, sub, focusType) {
  const tile = document.createElement("div");
  tile.className = "preview-tile";
  tile.setAttribute("role", "button");
  tile.tabIndex = 0;
  tile.innerHTML = `<p class="preview-title">${escapeHtml(title)}</p><p class="preview-sub">${escapeHtml(sub)}</p>`;
  const openResultModal = () => openResultModalWithFocus(focusType);
  tile.addEventListener("click", openResultModal);
  tile.addEventListener("keydown", (evt) => {
    if (evt.key === "Enter" || evt.key === " ") {
      evt.preventDefault();
      openResultModal();
    }
  });
  const host = document.createElement("div");
  host.className = "mini-chart-host";
  tile.appendChild(host);
  grid.appendChild(tile);
  return host;
}

function renderTinyChart(host, traces, yLabel) {
  Plotly.purge(host);
  Plotly.newPlot(
    host,
    traces.map((t) => ({
      ...t,
      mode: t.mode || "lines",
      line: { width: 1.8 }
    })),
    {
      margin: { l: 28, r: 10, t: 2, b: 22 },
      paper_bgcolor: "#fff",
      plot_bgcolor: "#fff",
      showlegend: false,
      xaxis: { showgrid: false, tickfont: { size: 9 } },
      yaxis: { title: yLabel, titlefont: { size: 9 }, tickfont: { size: 9 } }
    },
    { responsive: true, displayModeBar: false, staticPlot: true }
  )
    .then(() => {
      requestAnimationFrame(() => {
        try {
          Plotly.Plots.resize(host);
        } catch (_) {
          // Ignore transient resize errors during initial paint.
        }
      });
    })
    .catch(() => {
      host.innerHTML = '<p class="mini-empty">缩略图渲染失败</p>';
    });

  if (window.ResizeObserver) {
    const ro = new ResizeObserver(() => {
      try {
        Plotly.Plots.resize(host);
      } catch (_) {
        // Ignore transient resize errors while charts are being re-mounted.
      }
    });
    ro.observe(host);
    state.tinyChartObservers.push(ro);
  }
}

function renderTinyTable(grid, mergedRows, sources, focusType = "table") {
  const tile = document.createElement("div");
  tile.className = "preview-tile";
  tile.setAttribute("role", "button");
  tile.tabIndex = 0;
  tile.innerHTML = '<p class="preview-title">表格</p><p class="preview-sub">聚合结果前几行</p>';
  const openResultModal = () => openResultModalWithFocus(focusType);
  tile.addEventListener("click", openResultModal);
  tile.addEventListener("keydown", (evt) => {
    if (evt.key === "Enter" || evt.key === " ") {
      evt.preventDefault();
      openResultModal();
    }
  });

  const table = document.createElement("table");
  table.className = "mini-table";
  const cols = ["维度", ...sources.map((s) => s.name)];
  const head = `<thead><tr>${cols.slice(0, 3).map((c) => `<th>${escapeHtml(c)}</th>`).join("")}</tr></thead>`;
  const bodyRows = mergedRows
    .slice(0, 5)
    .map((row) => {
      const values = [row.dimension, ...sources.map((s) => fmtNum(row[s.name]))].slice(0, 3);
      return `<tr>${values.map((v) => `<td>${escapeHtml(v)}</td>`).join("")}</tr>`;
    })
    .join("");
  table.innerHTML = `${head}<tbody>${bodyRows}</tbody>`;
  tile.appendChild(table);
  grid.appendChild(tile);
}

function cleanupTinyPreviewArtifacts() {
  if (state.tinyChartObservers.length) {
    state.tinyChartObservers.forEach((ro) => ro.disconnect());
    state.tinyChartObservers = [];
  }
  chartsPreview.querySelectorAll(".mini-chart-host").forEach((host) => {
    try {
      Plotly.purge(host);
    } catch (_) {
      // Ignore if host has not been initialized by Plotly.
    }
  });
}

function createChartCard(container, title, exportName = "chart") {
  const card = document.createElement("div");
  card.className = "chart-card-panel";

  const head = document.createElement("div");
  head.className = "chart-card-head";
  const h3 = document.createElement("h3");
  h3.textContent = title;
  const downloadBtn = document.createElement("button");
  downloadBtn.type = "button";
  downloadBtn.className = "ghost mini-action";
  downloadBtn.textContent = "下载 PNG";

  const plot = document.createElement("div");
  plot.className = "plot";

  downloadBtn.addEventListener("click", () => {
    downloadPlotAsPngWithWhiteBackground(plot, buildExportFileName(exportName)).catch(() => {
      setStatus("下载失败，请重试", true);
    });
  });

  head.append(h3, downloadBtn);
  card.append(head, plot);
  container.appendChild(card);
  return plot;
}

function createResultTable(container, mergedRows, sources) {
  const card = document.createElement("div");
  card.className = "chart-card-panel";

  const head = document.createElement("div");
  head.className = "chart-card-head";
  const title = document.createElement("h3");
  title.textContent = "聚合结果对比表";

  const actions = document.createElement("div");
  actions.className = "table-actions";
  const csvBtn = document.createElement("button");
  csvBtn.type = "button";
  csvBtn.className = "ghost mini-action";
  csvBtn.textContent = "下载 CSV";
  const xlsxBtn = document.createElement("button");
  xlsxBtn.type = "button";
  xlsxBtn.className = "ghost mini-action";
  xlsxBtn.textContent = "下载 Excel";
  actions.append(csvBtn, xlsxBtn);
  head.append(title, actions);

  const wrap = document.createElement("div");
  wrap.className = "table-wrap";
  const table = document.createElement("table");
  const headers = ["维度", ...sources.map((s) => s.name)];
  const thead = document.createElement("thead");
  thead.innerHTML = `<tr>${headers.map((h) => `<th>${escapeHtml(h)}</th>`).join("")}</tr>`;

  const tbody = document.createElement("tbody");
  mergedRows.forEach((row) => {
    const tr = document.createElement("tr");
    const cols = [row.dimension, ...sources.map((s) => fmtNum(row[s.name]))];
    tr.innerHTML = cols.map((c) => `<td>${escapeHtml(c)}</td>`).join("");
    tbody.appendChild(tr);
  });

  table.append(thead, tbody);
  wrap.appendChild(table);
  card.append(head, wrap);
  container.appendChild(card);

  const tableData = mergedRows.map((row) => {
    const obj = { 维度: row.dimension };
    sources.forEach((s) => {
      obj[s.name] = row[s.name];
    });
    return obj;
  });

  csvBtn.addEventListener("click", () => downloadTableAsCsv(tableData, buildExportFileName("comparison_table") + ".csv"));
  xlsxBtn.addEventListener("click", () => downloadTableAsExcel(tableData, buildExportFileName("comparison_table") + ".xlsx"));
}

function downloadTableAsCsv(rows, filename) {
  const ws = XLSX.utils.json_to_sheet(rows);
  const csv = XLSX.utils.sheet_to_csv(ws);
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  triggerBrowserDownload(blob, filename);
}

function downloadTableAsExcel(rows, filename) {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, "comparison");
  const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  triggerBrowserDownload(blob, filename);
}

function triggerBrowserDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function buildExportFileName(prefix) {
  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  return `data_diff_${prefix}_${stamp}`;
}

async function downloadPlotAsPngWithWhiteBackground(plotDiv, filename) {
  const currentLayout = plotDiv && plotDiv.layout ? plotDiv.layout : {};
  const originalPaper = currentLayout.paper_bgcolor || "rgba(255,255,255,0)";
  const originalPlot = currentLayout.plot_bgcolor || "#fff";

  try {
    await Plotly.relayout(plotDiv, {
      paper_bgcolor: "#ffffff",
      plot_bgcolor: "#ffffff"
    });
    await Plotly.downloadImage(plotDiv, {
      format: "png",
      filename,
      width: 1600,
      height: 900,
      scale: 1
    });
  } finally {
    try {
      await Plotly.relayout(plotDiv, {
        paper_bgcolor: originalPaper,
        plot_bgcolor: originalPlot
      });
    } catch (_) {
      // Ignore restore errors and keep UI responsive.
    }
  }
}

function buildSeriesTrace(normalized, dimLabels, mergedRows, type) {
  const axisLabels = toAxisLabels(dimLabels);
  return normalized.map((s, idx) => {
    const common = {
      x: axisLabels,
      y: dimLabels.map((d) => findRowValue(mergedRows, d, s.name)),
      name: s.name,
      marker: { color: CHART_COLORS[idx % CHART_COLORS.length] },
      customdata: dimLabels,
      hovertemplate: "维度=%{customdata}<br>值=%{y}<extra>%{fullData.name}</extra>"
    };
    return type === "bar" ? { ...common, type: "bar" } : { ...common, type: "scatter", mode: "lines+markers" };
  });
}

function buildDeltaRows(normalized, dimLabels) {
  const baseline = normalized[0];
  const others = normalized.slice(1);
  return dimLabels.map((dim) => {
    const baseVal = baseline.grouped.get(dim) ?? 0;
    const row = { dim };
    others.forEach((s) => {
      const val = s.grouped.get(dim) ?? 0;
      row[`${s.name}__delta`] = val - baseVal;
      row[`${s.name}__pct`] = baseVal === 0 ? null : ((val - baseVal) / baseVal) * 100;
    });
    return row;
  });
}

function buildMainLayout(metricName, dimLabels) {
  const tickStep = getTickStep(dimLabels.length);
  return {
    margin: { l: 88, r: 20, t: 12, b: 44 },
    paper_bgcolor: "rgba(255,255,255,0)",
    plot_bgcolor: "#fff",
    legend: { orientation: "h", y: -0.22 },
    xaxis: {
      title: "维度",
      tickangle: -20,
      automargin: true,
      ticklabelstep: tickStep,
      ticklabeloverflow: "hide past div"
    },
    yaxis: {
      title: { text: `聚合值 (${metricName})`, standoff: 18 },
      automargin: true,
      zeroline: true,
      zerolinecolor: "#dbe0ea"
    }
  };
}

function toAxisLabels(dimLabels) {
  return dimLabels.map((label, idx) => shortDimensionLabel(label, idx));
}

function shortDimensionLabel(label, idx) {
  const raw = String(label || "");
  const compact = raw.replace(/\s+/g, " ").trim();
  if (compact.length <= 20) return compact;
  return `${compact.slice(0, 17)}...#${idx + 1}`;
}

function getTickStep(count) {
  if (count <= 30) return 1;
  if (count <= 80) return 2;
  if (count <= 150) return 3;
  if (count <= 300) return 5;
  if (count <= 600) return 8;
  return 12;
}

function prepareViewedData(mergedRows, dimLabels, normalized, viewConfig) {
  const mode = (viewConfig && viewConfig.mode) || "full";
  const topN = clampInt(viewConfig ? viewConfig.topN : 80, 10, 5000, 80);
  const windowSize = clampInt(viewConfig ? viewConfig.windowSize : 80, 20, 5000, 80);
  const windowStart = clampInt(viewConfig ? viewConfig.windowStart : 0, 0, 999999, 0);

  if (mode === "topn") {
    const baseline = normalized[0];
    const scored = dimLabels.map((dim) => ({
      dim,
      score: Math.abs((baseline.grouped.get(dim) ?? 0))
    }));
    scored.sort((a, b) => b.score - a.score);
    const chosenSet = new Set(scored.slice(0, Math.min(topN, scored.length)).map((x) => x.dim));
    const viewDims = dimLabels.filter((d) => chosenSet.has(d));
    return {
      dimLabels: viewDims,
      rows: mergedRows.filter((r) => chosenSet.has(r.dimension))
    };
  }

  if (mode === "window") {
    const start = Math.min(windowStart, Math.max(0, dimLabels.length - 1));
    const end = Math.min(dimLabels.length, start + windowSize);
    const viewDims = dimLabels.slice(start, end);
    const set = new Set(viewDims);
    return {
      dimLabels: viewDims,
      rows: mergedRows.filter((r) => set.has(r.dimension))
    };
  }

  return { dimLabels, rows: mergedRows };
}

function clampInt(value, min, max, fallback) {
  const n = Number.parseInt(value, 10);
  if (Number.isNaN(n)) return fallback;
  return Math.max(min, Math.min(max, n));
}

function findRowValue(rows, dim, sourceName) {
  const row = rows.find((r) => r.dimension === dim);
  return row ? row[sourceName] : 0;
}

function sortDimensionValue(a, b) {
  const na = Number(a);
  const nb = Number(b);
  if (!Number.isNaN(na) && !Number.isNaN(nb)) return na - nb;
  const da = Date.parse(a);
  const db = Date.parse(b);
  if (!Number.isNaN(da) && !Number.isNaN(db)) return da - db;
  return String(a).localeCompare(String(b), "zh-CN");
}

function getCheckedDimensions(sourceId) {
  return Array.from(
    sourceConfigListEl.querySelectorAll(`input[data-role="dimension-item"][data-id="${sourceId}"]:checked`)
  ).map((el) => el.value);
}

function setAllDimensions(sourceId, checked) {
  sourceConfigListEl
    .querySelectorAll(`input[data-role="dimension-item"][data-id="${sourceId}"]`)
    .forEach((el) => {
      el.checked = checked;
    });
}

function filterDimensionOptions(sourceId, keyword) {
  const kw = String(keyword || "").trim().toLowerCase();
  sourceConfigListEl
    .querySelectorAll(`.dim-option input[data-id="${sourceId}"]`)
    .forEach((checkbox) => {
      const label = checkbox.closest(".dim-option");
      if (!label) return;
      const hit = !kw || (label.dataset.field || "").includes(kw);
      label.style.display = hit ? "" : "none";
    });
}

function updateDimensionSummary(sourceId, selectedDimensions) {
  const summaryEl = sourceConfigListEl.querySelector(`.dim-summary[data-id="${sourceId}"]`);
  if (!summaryEl) return;
  summaryEl.textContent = buildDimensionSummaryText(selectedDimensions);
}

function buildDimensionSummaryText(selectedDimensions) {
  const dims = selectedDimensions || [];
  if (!dims.length) return "未选择维度字段";
  if (dims.length <= 3) return `已选 ${dims.length} 项：${dims.join("、")}`;
  return `已选 ${dims.length} 项：${dims.slice(0, 3).join("、")} 等`;
}

function fmtNum(v) {
  const n = Number(v);
  if (Number.isNaN(n)) return String(v);
  return Math.abs(n) >= 10000 ? n.toFixed(0) : n.toFixed(2).replace(/\.00$/, "");
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function escapeAttr(str) {
  return escapeHtml(str).replace(/`/g, "&#96;");
}
