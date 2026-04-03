const state = {
  activeSheetIndex: 0,
  result: null,
};

const currentTool = document.body.dataset.tool || "";
const toolConfig = {
  darnel: {
    endpoint: "/api/process/darnel",
    pendingLabel: "Procesando pedido y catálogo...",
    buttonLabel: "Procesar reporte",
  },
  zaplast: {
    endpoint: "/api/process/zaplast",
    pendingLabel: "Procesando PDFs...",
    buttonLabel: "Procesar PDFs",
  },
};

const config = toolConfig[currentTool];
if (config) {
  initToolPage(config);
}

function initToolPage(configForTool) {
  const statusBanner = document.getElementById("status-banner");
  const emptyState = document.getElementById("empty-state");
  const resultsCard = document.querySelector(".results-card");
  const resultsShell = document.getElementById("results-shell");
  const statsRow = document.getElementById("stats-row");
  const zoneSummary = document.getElementById("zone-summary");
  const warningBox = document.getElementById("warning-box");
  const sheetTabs = document.getElementById("sheet-tabs");
  const sheetPreview = document.getElementById("sheet-preview");
  const downloadButton = document.getElementById("download-button");
  const form = document.getElementById("tool-form");
  const fileOutputs = [...document.querySelectorAll("[data-file-output]")];

  [...document.querySelectorAll('input[type="file"]')].forEach((input) => {
    input.addEventListener("change", () => renderSelectedFiles(input, fileOutputs));
    renderSelectedFiles(input, fileOutputs);
  });

  form.addEventListener("submit", (event) => {
    event.preventDefault();
    processForm({
      configForTool,
      form,
      statusBanner,
      emptyState,
      resultsCard,
      resultsShell,
      statsRow,
      zoneSummary,
      warningBox,
      sheetTabs,
      sheetPreview,
    });
  });

  downloadButton.addEventListener("click", () => {
    if (!state.result) {
      return;
    }

    const bytes = base64ToBytes(state.result.workbookBase64);
    const blob = new Blob([bytes], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = state.result.downloadName;
    anchor.click();
    URL.revokeObjectURL(url);
  });
}

async function processForm({
  configForTool,
  form,
  statusBanner,
  emptyState,
  resultsCard,
  resultsShell,
  statsRow,
  zoneSummary,
  warningBox,
  sheetTabs,
  sheetPreview,
}) {
  const submitButton = form.querySelector('button[type="submit"]');
  const formData = new FormData(form);

  if ([...formData.values()].every((value) => !value || value.size === 0)) {
    showStatus(statusBanner, "Seleccioná los archivos antes de procesar.", "error");
    return;
  }

  setBusyState(submitButton, true, configForTool.pendingLabel);
  showStatus(statusBanner, configForTool.pendingLabel, "success");

  try {
    const response = await fetch(configForTool.endpoint, {
      method: "POST",
      body: formData,
    });
    const data = await response.json();
    if (!response.ok || !data.ok) {
      throw new Error(data.error || "No se pudo procesar el archivo.");
    }

    state.result = data;
    state.activeSheetIndex = 0;
    renderResult({
      result: data,
      emptyState,
      resultsCard,
      resultsShell,
      statsRow,
      zoneSummary,
      warningBox,
      sheetTabs,
      sheetPreview,
    });
    showStatus(
      statusBanner,
      "Reporte generado correctamente. Ya podés revisar la vista previa y descargar el Excel.",
      "success",
    );
  } catch (error) {
    showStatus(statusBanner, error.message, "error");
  } finally {
    setBusyState(submitButton, false, configForTool.buttonLabel);
  }
}

function setBusyState(button, isBusy, label) {
  button.disabled = isBusy;
  button.textContent = label;
}

function showStatus(element, message, variant) {
  element.hidden = false;
  element.classList.remove("is-error", "is-success");
  element.classList.add(variant === "error" ? "is-error" : "is-success");
  element.textContent = message;
}

function renderSelectedFiles(input, fileOutputs) {
  const output = fileOutputs.find((item) => item.dataset.fileOutput === input.name);
  if (!output) {
    return;
  }

  const files = [...(input.files || [])];
  if (files.length === 0) {
    output.innerHTML = "";
    return;
  }

  output.innerHTML = files
    .map((file) => `<span class="file-chip">${escapeHtml(file.name)}</span>`)
    .join("");
}

function renderResult({
  result,
  emptyState,
  resultsCard,
  resultsShell,
  statsRow,
  zoneSummary,
  warningBox,
  sheetTabs,
  sheetPreview,
}) {
  resultsCard.hidden = false;
  emptyState.hidden = true;
  resultsShell.hidden = false;
  renderStats(statsRow, result.preview.summary, result.type);
  renderZoneSummary(zoneSummary, result.preview.zoneSummary);
  renderWarnings(warningBox, result.warnings);
  renderSheetTabs(sheetTabs, result.preview.sheets, sheetPreview);
  renderActiveSheet(sheetPreview);
}

function renderStats(container, summary, type) {
  const itemsLabel = type === "darnel" ? "Productos" : "Artículos";
  const itemsValue = type === "darnel" ? summary.productos : summary.articulos;
  container.innerHTML = `
    <article class="stat-card">
      <div class="stat-value">${formatInteger(itemsValue)}</div>
      <div class="stat-label">${itemsLabel}</div>
    </article>
    <article class="stat-card">
      <div class="stat-value">${formatInteger(summary.zonas)}</div>
      <div class="stat-label">Zonas</div>
    </article>
    <article class="stat-card">
      <div class="stat-value">${formatDecimal(summary.palletsTotales)}</div>
      <div class="stat-label">Pallets Totales</div>
    </article>
  `;
}

function renderZoneSummary(container, zones) {
  container.innerHTML = zones
    .map(
      (zone) => `
        <article class="zone-pill">
          <strong>ZONA ${zone.zona}</strong>
          <span class="zone-meta">${formatInteger(zone.items)} items · ${formatInteger(zone.units)} uds · ${formatDecimal(zone.pallets)} pallets</span>
        </article>
      `,
    )
    .join("");
}

function renderWarnings(container, warnings) {
  if (!warnings || warnings.length === 0) {
    container.hidden = true;
    container.innerHTML = "";
    return;
  }

  container.hidden = false;
  container.innerHTML = `
    <strong>${warnings.length} coincidencia(s) pendientes</strong>
    <span>${warnings.map((warning) => escapeHtml(warning)).join(", ")}</span>
  `;
}

function renderSheetTabs(container, sheets, sheetPreview) {
  container.innerHTML = sheets
    .map(
      (sheet, index) => `
        <button class="sheet-tab ${index === state.activeSheetIndex ? "is-active" : ""}" data-sheet-index="${index}" type="button">
          ${escapeHtml(sheet.name)}
        </button>
      `,
    )
    .join("");

  [...container.querySelectorAll(".sheet-tab")].forEach((tab) => {
    tab.addEventListener("click", () => {
      state.activeSheetIndex = Number(tab.dataset.sheetIndex);
      renderSheetTabs(container, state.result.preview.sheets, sheetPreview);
      renderActiveSheet(sheetPreview);
    });
  });
}

function renderActiveSheet(container) {
  const sheet = state.result.preview.sheets[state.activeSheetIndex];
  if (!sheet) {
    container.innerHTML = "";
    return;
  }

  container.innerHTML = `
    <article class="sheet-shell">
      <header class="sheet-head">
        <h3 class="sheet-title">${escapeHtml(sheet.headline)}</h3>
        ${sheet.subheadline ? `<p class="sheet-subtitle">${escapeHtml(sheet.subheadline)}</p>` : ""}
      </header>
      ${sheet.groups.map((group) => renderGroup(sheet.headers, group)).join("")}
      ${
        sheet.grandTotal
          ? `<footer class="grand-total">
              <strong>${escapeHtml(sheet.grandTotal.label)}</strong>
              <div class="total-values">${renderTotals(sheet.grandTotal)}</div>
            </footer>`
          : ""
      }
    </article>
  `;
}

function renderGroup(headers, group) {
  return `
    <section class="sheet-group">
      <h4 class="group-label">${escapeHtml(group.label)}</h4>
      <div class="table-shell">
        <table>
          <thead>
            <tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr>
          </thead>
          <tbody>
            ${group.rows.map((row) => renderRow(headers, row)).join("")}
          </tbody>
        </table>
      </div>
      <div class="totals-bar">
        <strong>${escapeHtml(group.totals.label)}</strong>
        <div class="total-values">${renderTotals(group.totals)}</div>
      </div>
    </section>
  `;
}

function renderRow(headers, row) {
  return `
    <tr>
      ${headers.map((header) => `<td>${formatCell(row[header])}</td>`).join("")}
    </tr>
  `;
}

function renderTotals(totals) {
  return Object.entries(totals)
    .filter(([key]) => key !== "label")
    .map(([key, value]) => `<span>${escapeHtml(key)}: ${formatCell(value)}</span>`)
    .join("");
}

function formatCell(value) {
  if (value === null || value === undefined || value === "") {
    return "—";
  }

  if (typeof value === "number") {
    if (Number.isInteger(value)) {
      return formatInteger(value);
    }
    return formatDecimal(value);
  }

  return escapeHtml(String(value));
}

function formatInteger(value) {
  return new Intl.NumberFormat("es-CO", {
    maximumFractionDigits: 0,
  }).format(value);
}

function formatDecimal(value) {
  return new Intl.NumberFormat("es-CO", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(value);
}

function base64ToBytes(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
