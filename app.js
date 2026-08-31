/* global XLSX */

const WORKBOOK_URL = "cslb-stores.xlsx";

const METRIC_GROUPS = [
  {
    label: "Ranking",
    columnLetter: "F",
    valueType: "number",
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
];

const state = {
  workbook: null,
  stores: [],
};

const el = {
  statusBar: document.getElementById("statusBar"),
  metricsHost: document.getElementById("metricsHost"),
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
  if (valueType === "percent") return `${(numeric * 100).toFixed(2)}%`;
  if (valueType === "percentRaw") return `${numeric.toFixed(2)}%`;
  return String(value);
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

function parseStores(sheet) {
  if (!sheet) return [];

  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", header: 1 });
  if (!rows.length) return [];

  const storeRows = rows.slice(1);

  return storeRows.map((row) => {
    const districtName = normalizeText(row[0]); // Column A
    const storeName = normalizeText(row[3]); // Column D

    return {
      __districtName: districtName,
      __storeName: storeName,
      __rawRow: row,
    };
  });
}

function renderMetricTable(metricGroup) {
  const mode = el.modeSelect.value;
  const sortDirection = mode === "highest" ? -1 : 1;

  const rows = state.stores
    .map((row) => {
      const districtName = normalizeText(row.__districtName);
      const storeName = normalizeText(row.__storeName);
      const metricValue = getRowValueByColumnLetter(row.__rawRow, metricGroup.columnLetter);

      return {
        districtName: districtName || "N/A",
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
        <th>District</th>
        <th>Store Name</th>
        <th>${escapeHtml(metricGroup.label)}</th>
      </tr>
    </thead>
    <tbody>
      ${rows
        .map(
          (row) => `
            <tr>
              <td>${escapeHtml(row.districtName)}</td>
              <td>${escapeHtml(row.storeName)}</td>
              <td>${escapeHtml(formatMetricValue(row.metricValue, metricGroup.valueType))}</td>
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

function renderAllMetrics() {
  el.metricsHost.innerHTML = "";
  METRIC_GROUPS.forEach((metricGroup) => {
    const card = renderMetricTable(metricGroup);
    el.metricsHost.appendChild(card);
  });
}

async function loadWorkbook() {
  try {
    const response = await fetch(WORKBOOK_URL);
    if (!response.ok) throw new Error(`Could not fetch ${WORKBOOK_URL} (${response.status})`);
    const buffer = await response.arrayBuffer();
    state.workbook = XLSX.read(buffer, { type: "array" });

    const storeSheet = getSheetWithFallback(state.workbook, ["Store Sheet", "Store", "Store Data"]);
    state.stores = parseStores(storeSheet);

    const dateSheet = state.workbook.Sheets["Date"];
    const dataThrough = formatExcelDate(dateSheet?.A1);

    renderAllMetrics();

    el.statusBar.innerHTML = dataThrough
      ? `<strong>Data Through:</strong> ${escapeHtml(dataThrough)}`
      : `<strong>Data Through:</strong> <em>Not found</em>`;
  } catch (err) {
    console.error(err);
    el.statusBar.innerHTML = `
      <strong>Unable to load workbook.</strong>
      Please make sure <code>${escapeHtml(WORKBOOK_URL)}</code> exists in the repository root and is being served by the site.
      <br /><br />
      ${escapeHtml(err.message || String(err))}
    `;
  }
}

el.modeSelect.addEventListener("change", renderAllMetrics);

loadWorkbook();
