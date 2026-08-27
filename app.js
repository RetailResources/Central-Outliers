/* global XLSX */

const WORKBOOK_URL = "sales-dashboard.xlsx";

const METRIC_GROUPS = [
  {
    label: "GP Per Labor Hour",
    valueCandidates: ["GP Per Labor Hour %Tgt", "GP Per Labor Hour % Tgt", "GP/LH %Tgt", "GPH %Tgt"],
    rankCandidates: ["GP Per Labor Hour Rank", "GP/LH Rank", "GPH Rank"],
  },
  {
    label: "PP Act",
    valueCandidates: ["PP Act %Tgt", "PP Act % Tgt", "PP Activity %Tgt"],
    rankCandidates: ["PP Act Attain Rank", "PP Act Rank", "PP Activity Rank"],
  },
  {
    label: "ReBiz Conv",
    valueCandidates: ["ReBiz Conv %Tgt", "ReBiz Conv % Tgt", "ReBiz %Tgt"],
    rankCandidates: ["ReBiz Conv Rank", "ReBiz Rank", "Reservation Conversion Rank"],
  },
  {
    label: "Acc GP Pct",
    valueCandidates: ["Acc GP Pct Actual", "Acc GP Pct", "Acc GP Actual"],
    rankCandidates: ["Acc GP Pct Rank", "Acc GP Rank", "Adjusted GP Rank"],
  },
  {
    label: "CSAT",
    valueCandidates: ["CSAT Actual", "CSAT", "Customer Satisfaction Actual"],
    rankCandidates: ["CSAT Rank", "Customer Satisfaction Rank"],
  },
  {
    label: "Visa Priority",
    valueCandidates: ["Visa Priority Rate %Tg", "Visa Priority Rate %Tgt", "Visa Priority Rate", "Visa %Tgt"],
    rankCandidates: ["Visa Priority Rate Rank", "Visa Rank", "Priority Visa Rank"],
  },
  {
    label: "P360 Attach",
    valueCandidates: ["P360 Attach Rate %Tgt", "P360 Attach Rate % Tgt", "P360 Attach %Tgt"],
    rankCandidates: ["P360 Attach Rate Rank", "P360 Rank", "Premium Attach Rank"],
  },
  {
    label: "Premium Mix",
    valueCandidates: ["Premium Mix Rate %Tgt", "Premium Mix Rate % Tgt", "Premium %Tgt"],
    rankCandidates: ["Premium Rate Plan Rank", "Premium Mix Rank", "Premium Rank"],
  },
];

const state = {
  workbook: null,
  stores: [],
  selectedMetric: null,
};

const el = {
  statusBar: document.getElementById("statusBar"),
  metricsHost: document.getElementById("metricsHost"),
  metricSelect: document.getElementById("metricSelect"),
  modeSelect: document.getElementById("modeSelect"),
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

function normalizeLookup(value) {
  return normalizeText(value).toLowerCase();
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

function formatMetricValue(value, isCsat) {
  const numeric = parseNumeric(value);
  if (numeric === null) return "";
  return isCsat ? numeric.toFixed(2) : `${numeric.toFixed(2)}%`;
}

function normalizeKey(key) {
  return String(key || "").replace(/\s+/g, " ").trim();
}

function sheetToObjects(sheet) {
  if (!sheet) return [];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  return rows.map((row) => {
    const normalized = {};
    Object.entries(row).forEach(([key, value]) => {
      normalized[normalizeKey(key)] = typeof value === "string" ? normalizeText(value) : value;
    });
    return normalized;
  });
}

function getSheetWithFallback(workbook, candidates) {
  const names = workbook?.SheetNames || [];
  for (const candidate of candidates) {
    const exact = names.find((name) => normalizeLookup(name) === normalizeLookup(candidate));
    if (exact) return workbook.Sheets[exact];
  }
  for (const candidate of candidates) {
    const found = names.find((name) => normalizeLookup(name).includes(normalizeLookup(candidate)));
    if (found) return workbook.Sheets[found];
  }
  return null;
}

function findColumn(row, candidates) {
  const keys = Object.keys(row || {});
  for (const candidate of candidates) {
    const exact = keys.find((k) => normalizeLookup(k) === normalizeLookup(candidate));
    if (exact) return exact;
  }
  for (const candidate of candidates) {
    const fuzzy = keys.find((k) => normalizeLookup(k).includes(normalizeLookup(candidate)));
    if (fuzzy) return fuzzy;
  }
  return "";
}

function getRowValue(row, candidates) {
  const col = findColumn(row, candidates);
  return col ? row[col] : "";
}

function parseStores(rows) {
  return rows.map((row) => {
    const storeName =
      getRowValue(row, ["Sap: Loaction", "Sap: Location", "STORE NAME", "Store Name", "STORE CODE", "SAP"]) || "";
    return {
      ...row,
      __storeName: normalizeText(storeName),
    };
  });
}

function renderMetricTable(metricGroup) {
  const mode = el.modeSelect.value;
  const sortDirection = mode === "highest" ? -1 : 1;
  const isCsat = normalizeLookup(metricGroup.label).includes("csat");

  const rows = state.stores
    .map((row) => {
      const storeName = normalizeText(row.__storeName);
      const metricValue = getRowValue(row, metricGroup.valueCandidates);
      return {
        storeName,
        metricValue,
        sortValue: parseNumeric(metricValue),
      };
    })
    .filter((row) => row.storeName && row.sortValue !== null)
    .sort((a, b) => {
      if (a.sortValue === b.sortValue) {
        return a.storeName.toLowerCase().localeCompare(b.storeName.toLowerCase());
      }
      return (a.sortValue - b.sortValue) * sortDirection;
    })
    .slice(0, 20);

  const card = document.createElement("section");
  card.className = "metric-card";

  const title = document.createElement("h2");
  title.textContent = metricGroup.label;

  const caption = document.createElement("p");
  caption.className = "caption";
  caption.textContent = `${mode === "highest" ? "Top 20" : "Bottom 20"} stores`;

  const wrap = document.createElement("div");
  wrap.className = "table-wrap";

  const table = document.createElement("table");
  table.className = "data-table";

  table.innerHTML = `
    <thead>
      <tr>
        <th>Store Name</th>
        <th>${isCsat ? "CSAT Actual" : "Percent to Target"}</th>
      </tr>
    </thead>
    <tbody>
      ${rows
        .map(
          (row) => `
            <tr>
              <td>${escapeHtml(row.storeName)}</td>
              <td>${escapeHtml(formatMetricValue(row.metricValue, isCsat))}</td>
            </tr>
          `
        )
        .join("")}
    </tbody>
  `;

  wrap.appendChild(table);
  card.appendChild(title);
  card.appendChild(caption);
  card.appendChild(wrap);

  return card;
}

function renderSelectedMetric() {
  const selected = METRIC_GROUPS.find((group) => group.label === el.metricSelect.value) || METRIC_GROUPS[0];
  if (!selected) return;

  el.metricsHost.innerHTML = "";
  el.metricsHost.appendChild(renderMetricTable(selected));
}

function populateMetricSelect() {
  el.metricSelect.innerHTML = METRIC_GROUPS.map(
    (group) => `<option value="${escapeHtml(group.label)}">${escapeHtml(group.label)}</option>`
  ).join("");
  state.selectedMetric = METRIC_GROUPS[0]?.label || null;
  el.metricSelect.value = state.selectedMetric || "";
}

async function loadWorkbook() {
  try {
    const response = await fetch(WORKBOOK_URL);
    if (!response.ok) throw new Error(`Could not fetch ${WORKBOOK_URL} (${response.status})`);
    const buffer = await response.arrayBuffer();
    state.workbook = XLSX.read(buffer, { type: "array" });

    const storeSheet = getSheetWithFallback(state.workbook, ["Store Sheet", "Store", "Store Data"]);
    const storeRows = sheetToObjects(storeSheet);
    state.stores = parseStores(storeRows);

    populateMetricSelect();
    renderSelectedMetric();

    el.statusBar.innerHTML = `Loaded <strong>${escapeHtml(WORKBOOK_URL)}</strong> successfully.`;
  } catch (err) {
    console.error(err);
    el.statusBar.innerHTML = `
      <strong>Unable to load workbook.</strong>
      Please make sure <code>${escapeHtml(WORKBOOK_URL)}</code> exists in the repository root.
      <br /><br />
      ${escapeHtml(err.message || String(err))}
    `;
  }
}

el.metricSelect.addEventListener("change", renderSelectedMetric);
el.modeSelect.addEventListener("change", renderSelectedMetric);

loadWorkbook();
