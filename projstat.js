const LOCAL_SOURCE_KEY = "projectStatusSheetLink";

const COLUMN_ALIASES = {
  system: ["system", "project name", "system project name", "system (project name)"],
  milestone: ["milestone", "next milestone"],
  developer: ["assigned developer", "developer"],
  manager: ["assigned project manager", "project manager"]
};

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function normalizeHeader(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/[()]/g, " ")
    .replace(/[^a-z0-9\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function parseGvizResponse(text) {
  const start = text.indexOf("{");
  const end = text.lastIndexOf("}");

  if (start === -1 || end === -1 || end <= start) {
    throw new Error("Invalid response from Google Sheets.");
  }

  return JSON.parse(text.slice(start, end + 1));
}

function parseGoogleSheetLink(rawValue) {
  const trimmed = String(rawValue ?? "").trim();
  if (!trimmed) {
    throw new Error("Please paste a Google Sheets link.");
  }

  const idOnlyMatch = trimmed.match(/^[a-zA-Z0-9-_]{30,}$/);
  if (idOnlyMatch) {
    const url = new URL(`https://docs.google.com/spreadsheets/d/${idOnlyMatch[0]}/edit`);
    return { raw: trimmed, url, displayUrl: url.toString() };
  }

  const withProtocol = /^https?:\/\//i.test(trimmed) ? trimmed : `https://${trimmed}`;
  let url;

  try {
    url = new URL(withProtocol);
  } catch {
    const rawId = extractSheetIdFromRaw(trimmed);
    if (!rawId) {
      throw new Error("Invalid URL. Paste a valid Google Sheets link.");
    }

    url = new URL(`https://docs.google.com/spreadsheets/d/${rawId}/edit`);
  }

  const isDocsHost = url.hostname === "docs.google.com" || url.hostname.endsWith(".docs.google.com");
  const isDriveHost = url.hostname === "drive.google.com" || url.hostname.endsWith(".drive.google.com");
  const hasRawId = Boolean(extractSheetId(url) || extractSheetIdFromRaw(trimmed));

  if ((!isDocsHost && !isDriveHost) && !hasRawId) {
    throw new Error("Only Google Sheets links are supported.");
  }

  return { raw: trimmed, url, displayUrl: withProtocol };
}

function extractSheetIdFromRaw(rawValue) {
  const text = String(rawValue ?? "");
  const patterns = [
    /\/spreadsheets\/d\/([a-zA-Z0-9-_]{20,})/,
    /[?&]id=([a-zA-Z0-9-_]{20,})/,
    /[?&]key=([a-zA-Z0-9-_]{20,})/,
    /\/file\/d\/([a-zA-Z0-9-_]{20,})/
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match?.[1]) {
      return match[1];
    }
  }

  return "";
}

function extractSheetId(url) {
  const pathMatch = url.pathname.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (pathMatch?.[1]) {
    return pathMatch[1];
  }

  const driveFileMatch = url.pathname.match(/\/file\/d\/([a-zA-Z0-9-_]+)/);
  if (driveFileMatch?.[1]) {
    return driveFileMatch[1];
  }

  const idParam = String(url.searchParams.get("id") || "").trim();
  if (idParam) {
    return idParam;
  }

  const keyParam = String(url.searchParams.get("key") || "").trim();
  if (keyParam) {
    return keyParam;
  }

  return "";
}

function getSheetSource(url, rawValue = "") {
  const sheetId = extractSheetId(url) || extractSheetIdFromRaw(rawValue);
  if (!sheetId) {
    throw new Error("Unable to read this Google Sheets link.");
  }

  const hashParams = new URLSearchParams(String(url.hash ?? "").replace(/^#/, ""));

  return {
    sheetId,
    gid: String(url.searchParams.get("gid") || hashParams.get("gid") || "").trim(),
    sheet: String(url.searchParams.get("sheet") || hashParams.get("sheet") || "").trim()
  };
}

function buildSheetApiUrl(source) {
  const params = new URLSearchParams({ tqx: "out:json", headers: "1" });

  if (source.gid) {
    params.set("gid", source.gid);
  } else if (source.sheet) {
    params.set("sheet", source.sheet);
  }

  return `https://docs.google.com/spreadsheets/d/${encodeURIComponent(source.sheetId)}/gviz/tq?${params.toString()}`;
}

function buildJsonpSheetApiUrl(source, callbackName) {
  const params = new URLSearchParams({
    tqx: `out:json;responseHandler:${callbackName}`,
    headers: "1"
  });

  if (source.gid) {
    params.set("gid", source.gid);
  } else if (source.sheet) {
    params.set("sheet", source.sheet);
  }

  return `https://docs.google.com/spreadsheets/d/${encodeURIComponent(source.sheetId)}/gviz/tq?${params.toString()}`;
}

function buildCsvApiUrl(source) {
  const params = new URLSearchParams({ format: "csv" });

  if (source.gid) {
    params.set("gid", source.gid);
  } else if (source.sheet) {
    params.set("sheet", source.sheet);
  }

  return `https://docs.google.com/spreadsheets/d/${encodeURIComponent(source.sheetId)}/export?${params.toString()}`;
}

function buildOpenSheetApiUrl(source) {
  if (!source.sheet) {
    return "";
  }

  return `https://opensheet.elk.sh/${encodeURIComponent(source.sheetId)}/${encodeURIComponent(source.sheet)}`;
}

function formatCellValue(cell, columnType) {
  if (!cell) return "";

  if (typeof cell.f === "string" && cell.f.trim()) {
    return cell.f;
  }

  if (cell.v == null) {
    return "";
  }

  if (columnType === "date" && typeof cell.v === "string") {
    const parts = cell.v.match(/^Date\((\d+),(\d+),(\d+)/);
    if (parts) {
      const year = Number(parts[1]);
      const month = Number(parts[2]);
      const day = Number(parts[3]);
      return new Date(year, month, day).toLocaleDateString();
    }
  }

  return String(cell.v);
}

function mapPayload(payload) {
  const columns = (payload.table?.cols ?? []).map((column, index) => ({
    key: normalizeHeader(column.label) || `column ${index + 1}`,
    label: String(column.label || `Column ${index + 1}`).trim(),
    type: column.type || "string"
  }));

  const rows = (payload.table?.rows ?? []).map((row) => {
    const mapped = {};

    columns.forEach((column, index) => {
      mapped[column.key] = formatCellValue(row.c?.[index], column.type);
    });

    return mapped;
  });

  return { columns, rows };
}

function parseCsvLine(line) {
  const values = [];
  let current = "";
  let inQuotes = false;

  for (let i = 0; i < line.length; i += 1) {
    const char = line[i];

    if (char === '"') {
      if (inQuotes && line[i + 1] === '"') {
        current += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (char === "," && !inQuotes) {
      values.push(current);
      current = "";
      continue;
    }

    current += char;
  }

  values.push(current);
  return values;
}

function parseCsvText(csvText) {
  const lines = String(csvText ?? "")
    .replace(/^\uFEFF/, "")
    .split(/\r?\n/)
    .filter((line) => line.trim().length > 0);

  if (!lines.length) {
    throw new Error("No rows found in CSV export.");
  }

  const headerValues = parseCsvLine(lines[0]);
  const columns = headerValues.map((header, index) => ({
    key: normalizeHeader(header) || `column ${index + 1}`,
    label: String(header || `Column ${index + 1}`).trim(),
    type: "string"
  }));

  const rows = lines.slice(1).map((line) => {
    const values = parseCsvLine(line);
    const mapped = {};

    columns.forEach((column, index) => {
      mapped[column.key] = String(values[index] ?? "").trim();
    });

    return mapped;
  });

  return { columns, rows };
}

function fetchSheetPayloadWithJsonp(source) {
  return new Promise((resolve, reject) => {
    const callbackName = `sheetCallback_${Date.now()}_${Math.floor(Math.random() * 1e6)}`;
    const script = document.createElement("script");

    const timeout = window.setTimeout(() => {
      cleanup();
      reject(new Error("Timed out while loading sheet data."));
    }, 15000);

    function cleanup() {
      window.clearTimeout(timeout);
      delete window[callbackName];
      script.remove();
    }

    window[callbackName] = (payload) => {
      cleanup();
      resolve(payload);
    };

    script.onerror = () => {
      cleanup();
      reject(new Error("Could not load sheet data via JSONP."));
    };

    script.src = buildJsonpSheetApiUrl(source, callbackName);
    document.head.appendChild(script);
  });
}

async function fetchSheetMappedData(source) {
  try {
    const gvizResponse = await fetch(buildSheetApiUrl(source), { cache: "no-store" });
    if (!gvizResponse.ok) {
      throw new Error(`Google Sheet API returned ${gvizResponse.status}`);
    }

    const payload = parseGvizResponse(await gvizResponse.text());
    if (payload?.table) {
      return mapPayload(payload);
    }
  } catch {
    // Fall through to JSONP/CSV strategies.
  }

  try {
    const jsonpPayload = await fetchSheetPayloadWithJsonp(source);
    if (jsonpPayload?.table) {
      return mapPayload(jsonpPayload);
    }
  } catch {
    // Fall through to CSV strategy.
  }

  try {
    const csvResponse = await fetch(buildCsvApiUrl(source), { cache: "no-store" });
    if (csvResponse.ok) {
      return parseCsvText(await csvResponse.text());
    }
  } catch {
    // Fall through to OpenSheet fallback.
  }

  const openSheetUrl = buildOpenSheetApiUrl(source);
  if (openSheetUrl) {
    const openSheetResponse = await fetch(openSheetUrl, { cache: "no-store" });
    if (openSheetResponse.ok) {
      const rows = await openSheetResponse.json();
      if (Array.isArray(rows) && rows.length) {
        const firstRow = rows[0];
        const columns = Object.keys(firstRow).map((label, index) => ({
          key: normalizeHeader(label) || `column ${index + 1}`,
          label: String(label || `Column ${index + 1}`).trim(),
          type: "string"
        }));
        const mappedRows = rows.map((row) => {
          const mapped = {};
          columns.forEach((column) => {
            mapped[column.key] = String(row[column.label] ?? "").trim();
          });
          return mapped;
        });
        return { columns, rows: mappedRows };
      }
    }
  }

  throw new Error("Unable to load sheet data. Make sure the sheet is shared to Anyone with the link (Viewer) or published to the web.");
}

function findColumnKey(columns, aliases) {
  const normalizedAliases = aliases.map((alias) => normalizeHeader(alias));

  for (const alias of normalizedAliases) {
    const exact = columns.find((column) => column.key === alias);
    if (exact) return exact.key;
  }

  for (const alias of normalizedAliases) {
    const partial = columns.find((column) => column.key.includes(alias) || alias.includes(column.key));
    if (partial) return partial.key;
  }

  return "";
}

function uniqueValues(rows, key) {
  if (!key) return [];

  return [...new Set(rows
    .map((row) => String(row[key] ?? "").trim())
    .filter(Boolean))].sort((a, b) => a.localeCompare(b));
}

function setFeedback(message, isError = false) {
  const feedback = document.getElementById("sheet-source-feedback");
  feedback.textContent = message;
  feedback.style.color = isError ? "#ef4444" : "";
}

const state = {
  columns: [],
  rows: [],
  filteredRows: [],
  filters: {
    system: "",
    milestone: "",
    search: ""
  },
  filterColumns: {
    system: "",
    milestone: "",
    developer: "",
    manager: ""
  }
};

function determineFilterColumns(columns) {
  const systemKey = findColumnKey(columns, COLUMN_ALIASES.system) || columns[0]?.key || "";
  const milestoneKey = findColumnKey(columns, COLUMN_ALIASES.milestone) || columns[1]?.key || columns[0]?.key || "";
  const developerKey = findColumnKey(columns, COLUMN_ALIASES.developer);
  const managerKey = findColumnKey(columns, COLUMN_ALIASES.manager);

  return {
    system: systemKey,
    milestone: milestoneKey,
    developer: developerKey,
    manager: managerKey
  };
}

function renderSummary() {
  const totalProjects = document.getElementById("kpi-total-projects");
  const totalMilestones = document.getElementById("kpi-total-milestones");

  const projectKey = state.filterColumns.system || state.columns[0]?.key || "";
  const projectValues = uniqueValues(state.filteredRows, projectKey);

  totalProjects.textContent = String(projectValues.length);
  totalMilestones.textContent = String(state.filteredRows.length);
}

function renderFilterOptions() {
  const systemSelect = document.getElementById("system-filter");
  const milestoneSelect = document.getElementById("milestone-filter");

  const systemOptions = uniqueValues(state.rows, state.filterColumns.system);
  const milestoneOptions = uniqueValues(state.rows, state.filterColumns.milestone);

  const buildOptions = (values, selected) => {
    const options = ["<option value=\"\">All</option>"];

    values.forEach((value) => {
      const safe = escapeHtml(value);
      const selectedMarker = value === selected ? " selected" : "";
      options.push(`<option value="${safe}"${selectedMarker}>${safe}</option>`);
    });

    return options.join("");
  };

  systemSelect.innerHTML = buildOptions(systemOptions, state.filters.system);
  milestoneSelect.innerHTML = buildOptions(milestoneOptions, state.filters.milestone);

  systemSelect.disabled = !systemOptions.length;
  milestoneSelect.disabled = !milestoneOptions.length;
}

function rowMatchesSearch(row, query) {
  if (!query) return true;

  const keys = [
    state.filterColumns.system,
    state.filterColumns.milestone,
    state.filterColumns.developer,
    state.filterColumns.manager
  ].filter(Boolean);

  return keys.some((key) => String(row[key] ?? "").toLowerCase().includes(query));
}

function renderTable() {
  const head = document.getElementById("status-table-head");
  const body = document.getElementById("status-table-body");

  if (!state.columns.length) {
    head.innerHTML = "";
    body.innerHTML = '<tr><td colspan="1">No results found</td></tr>';
    return;
  }

  head.innerHTML = `
    <tr>
      ${state.columns.map((column) => `<th>${escapeHtml(column.label)}</th>`).join("")}
    </tr>
  `;

  if (!state.filteredRows.length) {
    body.innerHTML = `<tr><td colspan="${state.columns.length}">No results found</td></tr>`;
    return;
  }

  body.innerHTML = state.filteredRows
    .map((row) => `
      <tr>
        ${state.columns.map((column) => `<td>${escapeHtml(row[column.key] ?? "")}</td>`).join("")}
      </tr>
    `)
    .join("");
}

function applyFilters() {
  const searchQuery = state.filters.search.trim().toLowerCase();

  state.filteredRows = state.rows.filter((row) => {
    if (state.filters.system && String(row[state.filterColumns.system] ?? "") !== state.filters.system) {
      return false;
    }

    if (state.filters.milestone && String(row[state.filterColumns.milestone] ?? "") !== state.filters.milestone) {
      return false;
    }

    return rowMatchesSearch(row, searchQuery);
  });

  renderTable();
  renderSummary();
}

function bindFilters() {
  document.getElementById("system-filter").addEventListener("change", (event) => {
    state.filters.system = event.target.value;
    applyFilters();
  });

  document.getElementById("milestone-filter").addEventListener("change", (event) => {
    state.filters.milestone = event.target.value;
    applyFilters();
  });

  document.getElementById("sheet-search").addEventListener("input", (event) => {
    state.filters.search = event.target.value;
    applyFilters();
  });
}

async function loadSheetData(parsedLink) {
  const source = getSheetSource(parsedLink.url, parsedLink.raw);
  const mapped = await fetchSheetMappedData(source);

  if (!mapped.columns.length) {
    throw new Error("No columns found in the selected sheet.");
  }

  state.columns = mapped.columns;
  state.rows = mapped.rows;
  state.filterColumns = determineFilterColumns(mapped.columns);
  state.filters.system = "";
  state.filters.milestone = "";

  renderFilterOptions();
  applyFilters();
}

async function handleSubmit(event) {
  event.preventDefault();

  const input = document.getElementById("sheet-source-input");

  try {
    const parsed = parseGoogleSheetLink(input.value);
    localStorage.setItem(LOCAL_SOURCE_KEY, parsed.displayUrl);
    input.value = parsed.displayUrl;

    await loadSheetData(parsed);
    setFeedback("Sheet loaded successfully.");
  } catch (error) {
    state.columns = [];
    state.rows = [];
    state.filteredRows = [];
    state.filterColumns = { system: "", milestone: "", developer: "", manager: "" };

    renderFilterOptions();
    renderTable();
    renderSummary();
    setFeedback(error.message || "Unable to load Google Sheets data.", true);
  }
}

function resetView() {
  localStorage.removeItem(LOCAL_SOURCE_KEY);

  document.getElementById("sheet-source-input").value = "";
  document.getElementById("sheet-search").value = "";

  state.columns = [];
  state.rows = [];
  state.filteredRows = [];
  state.filters = { system: "", milestone: "", search: "" };
  state.filterColumns = { system: "", milestone: "", developer: "", manager: "" };

  renderFilterOptions();
  renderTable();
  renderSummary();
  setFeedback("Paste a Google Sheets link and click View Data.");
}

function init() {
  document.getElementById("sheet-source-form").addEventListener("submit", handleSubmit);
  document.getElementById("sheet-source-reset").addEventListener("click", resetView);

  bindFilters();
  resetView();

  const savedLink = localStorage.getItem(LOCAL_SOURCE_KEY);
  if (savedLink) {
    document.getElementById("sheet-source-input").value = savedLink;
    document.getElementById("sheet-source-form").requestSubmit();
  }
}

init();