/* global XLSX */

const WORKBOOK_URL = "cslb-stores.xlsx";

const DASHBOARD_CONFIG = {
  defaultMode: "stores",
  modes: {
    stores: {
      label: "Stores",
      sheetCandidates: ["Store", "Store Sheet", "Store Data"],
      districtColumnLetter: "A",
      nameColumnLetter: "D",
      nameHeader: "Store Name",
      itemLabelPlural: "stores",
      metrics: [
        {
          label: "Ranking",
          columnLetter: "F",
          valueType: "rank",
        },
        {
          label: "GP Per Labor Hour Actual",
          columnLetter: "I",
          valueType: "currency",
        },
        {
          label: "PP Act %Tgt",
          columnLetter: "O",
          valueType: "percent",
        },
        {
          label: "Rebiz Conv",
          columnLetter: "S",
          valueType: "percent",
        },
        {
          label: "Acc GP Pct Actual",
          columnLetter: "W",
          valueType: "percent",
        },
        {
          label: "CSAT Actual",
          columnLetter: "Y",
          valueType: "number",
        },
        {
          label: "Visa Priority Rate",
          columnLetter: "AC",
          valueType: "percent",
        },
        {
          label: "Indexed P360 Attach Rate",
          columnLetter: "AG",
          valueType: "percent",
        },
        {
          label: "Premium Mix Rate",
          columnLetter: "AM",
          valueType: "percent",
        },
      ],
    },
    employees: {
      label: "Employees",
      sheetCandidates: ["Employee", "Employee Sheet", "Employees", "Employee Data"],
      districtColumnLetter: "A",
      storeNameColumnLetter: "F",
      employeeNameColumnLetter: "E",
      renderAsEmployeeMode: true,
      itemLabelPlural: "employees",
      metrics: [
        {
          label: "Ranking",
          columnLetter: "H",
          valueType: "rank",
        },
        {
          label: "GP per Labor Hour Actual",
          columnLetter: "K",
          valueType: "currency",
        },
        {
          label: "Rebiz Act Conversion",
          columnLetter: "Q",
          valueType: "percent",
        },
        {
          label: "Rebiz Upgrade Conversion",
          columnLetter: "W",
          valueType: "percent",
        },
        {
          label: "Accessory Per Phone",
          columnLetter: "AA",
          valueType: "number",
        },
        {
          label: "CSAT",
          columnLetter: "AC",
          valueType: "number",
        },
        {
          label: "Visa Priority",
          columnLetter: "AG",
          valueType: "percent",
        },
        {
          label: "P360 Attach",
          columnLetter: "AK",
          valueType: "percent",
        },
        {
          label: "Premium Mix",
          columnLetter: "AQ",
          valueType: "percent",
        },
      ],
    },
  },
};

const state = {
  workbook: null,
  dataByMode: {
    stores: [],
    employees: [],
  },
};

const el = {
  statusBar: document.getElementById("statusBar"),
  metricsHost: document.getElementById("metricsHost"),
  dashboardModeSelect: document.getElementById("dashboardModeSelect"),
  viewModeSelect: document.getElementById("viewModeSelect"),
  districtControl: document.getElementById("districtControl"),
  districtSelect: document.getElementById("districtSelect"),
};

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function normalizeText(value) {
  return String(value ?? "").trim().replace(/\s+/g, " ");
}

function parseNumeric(value) {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "number") return value;
  const raw = String(value).trim();
  if (!raw) return null;
  const stripped = raw.replace(/[(),$%]/g, "").replace(/\s+/g, "");
  const num = Number(stripped);
  return Number.isNaN(num) ? null : num;
}

function formatMetricValue(value, valueType) {
  const numeric = parseNumeric(value);
  if (numeric === null) return "";
  if (valueType === "currency") return `$${numeric.toFixed(2)}`;
  if (valueType === "number") return numeric.toFixed(2);
  if (valueType === "rank") return String(Math.round(numeric));
  if (valueType === "percent") return `${(numeric * 100).toFixed(2)}%`;
  if (valueType === "percentRaw") return `${numeric.toFixed(2)}%`;
  return String(value);
}

function formatRawMetricValue(value, valueType) {
  const raw = value === null || value === undefined ? "" : String(value).trim();
  if (!raw) return "";

  const numeric = parseNumeric(raw);
  if (numeric === null) return escapeHtml(raw);

  if (valueType === "currency") return `$${numeric.toFixed(2)}`;
  if (valueType === "number") return numeric.toFixed(2);
  if (valueType === "rank") return String(Math.round(numeric));
  if (valueType === "percent") return `${(numeric * 100).toFixed(2)}%`;
  if (valueType === "percentRaw") return `${numeric.toFixed(2)}%`;
  return escapeHtml(raw);
}

function getSheetWithFallback(workbook, candidates) {
  const names = workbook?.SheetNames || [];
  for (const candidate of candidates) {
    const exact = names.find((name) => normalizeText(name).toLowerCase() === normalizeText(candidate).toLowerCase());
    if (exact) return workbook.Sheets[exact];
  }
  for (const candidate of candidates) {
    const found = names.find((name) => normalizeText(name).toLowerCase().includes(normalizeText(candidate).toLowerCase()));
    if (found) return workbook.Sheets[found];
  }
  return null;
}

function columnLetterToIndex(letter) {
  let result = 0;
  const clean = String(letter || "").toUpperCase();
  for (let i = 0; i < clean.length; i++) {
    result = result * 26 + (clean.charCodeAt(i) - 64);
  }
  return result - 1;
}

function getRowValueByColumnLetter(row, columnLetter) {
  const index = columnLetterToIndex(columnLetter);
  return row?.[index] ?? "";
}

function formatExcelDate(cell) {
  if (!cell) return "";

  if (cell.w) return normalizeText(cell.w);

  const value = cell.v ?? cell;

  if (value instanceof Date) {
    return value.toLocaleDateString("en-US", {
      month: "long",
      day: "numeric",
      year: "numeric",
    });
  }

  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return "";
    const date = new Date(parsed.y, parsed.m - 1, parsed.d);
    return date.toLocaleDateString("en-US", {
      month: "long",
      day: "numeric",
      year: "numeric",
    });
  }

  return normalizeText(value);
}

function parseRowsForMode(sheet, modeConfig) {
  if (!sheet) return [];

  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", header: 1 });
  if (!rows.length) return [];

  return rows
    .slice(1)
    .filter((row) => {
      if (modeConfig.renderAsEmployeeMode) {
        const columnC = normalizeText(row?.[2]);
        return !columnC.toUpperCase().includes("RSM");
      }
      return true;
    })
    .map((row) => ({
      __districtName: normalizeText(getRowValueByColumnLetter(row, modeConfig.districtColumnLetter)),
      __storeName: normalizeText(getRowValueByColumnLetter(row, modeConfig.storeNameColumnLetter)),
      __employeeName: normalizeText(getRowValueByColumnLetter(row, modeConfig.employeeNameColumnLetter)),
      __itemName: normalizeText(getRowValueByColumnLetter(row, modeConfig.nameColumnLetter)),
      __rawRow: row,
    }));
}

function getActiveModeKey() {
  return el.dashboardModeSelect.value || DASHBOARD_CONFIG.defaultMode;
}

function getActiveModeConfig() {
  return DASHBOARD_CONFIG.modes[getActiveModeKey()] || DASHBOARD_CONFIG.modes[DASHBOARD_CONFIG.defaultMode];
}

function getFilteredRowsForActiveMode() {
  const modeKey = getActiveModeKey();
  const rows = state.dataByMode[modeKey] || [];
  const selectedDistrict = el.districtSelect.value;
  if (!selectedDistrict || selectedDistrict === "all") return rows;
  return rows.filter((row) => normalizeText(row.__districtName) === selectedDistrict);
}

function populateDistrictOptions() {
  const previous = el.districtSelect.value;
  const districts = Array.from(
    new Set(
      (state.dataByMode.employees || [])
        .map((row) => normalizeText(row.__districtName))
        .filter(Boolean)
    )
  ).sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));

  const options = ['<option value="all">All Districts</option>']
    .concat(
      districts.map((district) => `<option value="${escapeHtml(district)}">${escapeHtml(district)}</option>`)
    )
    .join("");

  el.districtSelect.innerHTML = options;
  if (previous && districts.includes(previous)) {
    el.districtSelect.value = previous;
  } else {
    el.districtSelect.value = "all";
  }
}

function updateControlVisibility() {
  const showDistrict = true;
  el.districtControl.hidden = !showDistrict;
}

function renderMetricTable(metricGroup, modeConfig, sourceRows) {
  const viewMode = el.viewModeSelect.value;
  const isRankMetric = metricGroup.valueType === "rank";
  const sortDirection = isRankMetric
    ? viewMode === "highest"
      ? 1
      : -1
    : viewMode === "highest"
      ? -1
      : 1;

  const rows = sourceRows
    .map((row) => {
      const metricValue = row.__rawRow?.[columnLetterToIndex(metricGroup.columnLetter)] ?? "";
      return {
        districtName: normalizeText(row.__districtName) || "N/A",
        storeName: normalizeText(row.__storeName) || "N/A",
        employeeName: normalizeText(row.__employeeName) || "N/A",
        itemName: normalizeText(row.__itemName) || "N/A",
        metricValue,
        sortValue: parseNumeric(metricValue),
      };
    })
    .filter((row) => row.employeeName || row.itemName)
    .sort((a, b) => {
      const aLabel = modeConfig.renderAsEmployeeMode ? a.employeeName : a.itemName;
      const bLabel = modeConfig.renderAsEmployeeMode ? b.employeeName : b.itemName;

      if (a.sortValue === b.sortValue) {
        return aLabel.toLowerCase().localeCompare(bLabel.toLowerCase());
      }
      if (a.sortValue === null) return 1;
      if (b.sortValue === null) return -1;
      return (a.sortValue - b.sortValue) * sortDirection;
    })
    .slice(0, 20);

  const card = document.createElement("section");
  card.className = "metric-card";

  const title = document.createElement("h2");
  title.textContent = metricGroup.label;

  const caption = document.createElement("p");
  caption.className = "caption";
  caption.textContent = `${viewMode === "highest" ? "Top 20" : "Bottom 20"} ${modeConfig.itemLabelPlural}`;

  const wrap = document.createElement("div");
  wrap.className = "table-wrap";

  const table = document.createElement("table");
  table.className = "data-table";

  if (modeConfig.renderAsEmployeeMode) {
    table.innerHTML = `
      <thead>
        <tr>
          <th>Store Name</th>
          <th>Employee</th>
          <th>Metric</th>
        </tr>
      </thead>
      <tbody>
        ${rows
          .map(
            (row) => `
              <tr>
                <td>${escapeHtml(row.storeName)}</td>
                <td>${escapeHtml(row.employeeName)}</td>
                <td>${formatRawMetricValue(row.metricValue, metricGroup.valueType)}</td>
              </tr>
            `
          )
          .join("")}
      </tbody>
    `;
  } else {
    table.innerHTML = `
      <thead>
        <tr>
          <th>District</th>
          <th>${escapeHtml(modeConfig.nameHeader)}</th>
          <th>${escapeHtml(metricGroup.label)}</th>
        </tr>
      </thead>
      <tbody>
        ${rows
          .map(
            (row) => `
              <tr>
                <td>${escapeHtml(row.districtName)}</td>
                <td>${escapeHtml(row.itemName)}</td>
                <td>${escapeHtml(formatMetricValue(row.metricValue, metricGroup.valueType))}</td>
              </tr>
            `
          )
          .join("")}
      </tbody>
    `;
  }

  wrap.appendChild(table);
  card.appendChild(title);
  card.appendChild(caption);
  card.appendChild(wrap);

  return card;
}

function renderAllMetrics() {
  const modeConfig = getActiveModeConfig();
  const rows = getFilteredRowsForActiveMode();

  el.metricsHost.innerHTML = "";
  modeConfig.metrics.forEach((metricGroup) => {
    el.metricsHost.appendChild(renderMetricTable(metricGroup, modeConfig, rows));
  });
}

async function loadWorkbook() {
  try {
    const response = await fetch(WORKBOOK_URL);
    if (!response.ok) throw new Error(`Could not fetch ${WORKBOOK_URL} (${response.status})`);
    const buffer = await response.arrayBuffer();
    state.workbook = XLSX.read(buffer, { type: "array" });

    Object.entries(DASHBOARD_CONFIG.modes).forEach(([modeKey, modeConfig]) => {
      const sheet = getSheetWithFallback(state.workbook, modeConfig.sheetCandidates || []);
      state.dataByMode[modeKey] = parseRowsForMode(sheet, modeConfig);
    });

    const dateSheet = getSheetWithFallback(state.workbook, ["Date"]);
    const dataThrough = formatExcelDate(dateSheet?.A1);

    populateDistrictOptions();
    updateControlVisibility();
    renderAllMetrics();

    const refreshedAt = new Date().toLocaleString("en-US", {
      month: "long",
      day: "numeric",
      year: "numeric",
      hour: "numeric",
      minute: "2-digit",
      second: "2-digit",
    });

    el.statusBar.innerHTML = dataThrough
      ? `<strong>Data Through:</strong> ${escapeHtml(dataThrough)} <span class="subtle">(${escapeHtml(refreshedAt)})</span>`
      : `<strong>Data Through:</strong> <em>Not found</em> <span class="subtle">(${escapeHtml(refreshedAt)})</span>`;
  } catch (err) {
    console.error(err);
    const refreshedAt = new Date().toLocaleString("en-US", {
      month: "long",
      day: "numeric",
      year: "numeric",
      hour: "numeric",
      minute: "2-digit",
      second: "2-digit",
    });
    el.statusBar.innerHTML = `
      <strong>Unable to load workbook.</strong>
      Please make sure <code>${escapeHtml(WORKBOOK_URL)}</code> exists in the repository root and is being served by the site.
      <br /><br />
      ${escapeHtml(err.message || String(err))}
      <br /><br />
      <span class="subtle">${escapeHtml(refreshedAt)}</span>
    `;
  }
}

el.dashboardModeSelect.addEventListener("change", () => {
  updateControlVisibility();
  renderAllMetrics();
});
el.viewModeSelect.addEventListener("change", renderAllMetrics);
el.districtSelect.addEventListener("change", renderAllMetrics);

loadWorkbook();
