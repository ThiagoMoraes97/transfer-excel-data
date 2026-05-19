const state = {
  source: { token: null, sheets: [], columns: [], filename: "" },
  dest: { token: null, sheets: [], columns: [], filename: "" },
};

const els = {
  sourceFile: document.getElementById("sourceFile"),
  destFile: document.getElementById("destFile"),
  sourceFileName: document.getElementById("sourceFileName"),
  destFileName: document.getElementById("destFileName"),
  sourceSheet: document.getElementById("sourceSheet"),
  destSheet: document.getElementById("destSheet"),
  sourceColumns: document.getElementById("sourceColumns"),
  destColumns: document.getElementById("destColumns"),
  sourceStartRow: document.getElementById("sourceStartRow"),
  destStartRow: document.getElementById("destStartRow"),
  separator: document.getElementById("separator"),
  skipEmpty: document.getElementById("skipEmpty"),
  filterList: document.getElementById("filterList"),
  filterRowTemplate: document.getElementById("filterRowTemplate"),
  addFilterBtn: document.getElementById("addFilterBtn"),
  mappingList: document.getElementById("mappingList"),
  mappingRowTemplate: document.getElementById("mappingRowTemplate"),
  addMappingBtn: document.getElementById("addMappingBtn"),
  runBtn: document.getElementById("runBtn"),
  status: document.getElementById("status"),
};

function setStatus(message, type = "info") {
  els.status.textContent = message;
  els.status.dataset.type = type;
}

function columnDisplay(col) {
  const header = col.header ? ` | ${col.header}` : "";
  return `${col.label}${header}`;
}

function clearElement(el) {
  while (el.firstChild) el.removeChild(el.firstChild);
}

function renderColumnChips(container, columns) {
  clearElement(container);
  if (!columns.length) {
    const chip = document.createElement("span");
    chip.className = "chip empty";
    chip.textContent = "Sem colunas com dados";
    container.appendChild(chip);
    return;
  }

  columns.forEach((col) => {
    const chip = document.createElement("span");
    chip.className = "chip";
    chip.textContent = columnDisplay(col);
    chip.title = col.sample ? `Exemplo: ${col.sample}` : "";
    container.appendChild(chip);
  });
}

function fillSheetSelect(selectEl, sheets, selected) {
  clearElement(selectEl);
  sheets.forEach((sheet) => {
    const opt = document.createElement("option");
    opt.value = sheet;
    opt.textContent = sheet;
    if (sheet === selected) opt.selected = true;
    selectEl.appendChild(opt);
  });
  selectEl.disabled = sheets.length === 0;
}

function optionFromColumn(col, includeBlank = false) {
  const fragment = document.createDocumentFragment();
  if (includeBlank) {
    const blank = document.createElement("option");
    blank.value = "";
    blank.textContent = "(vazio)";
    fragment.appendChild(blank);
  }

  const opt = document.createElement("option");
  opt.value = col.letter;
  opt.textContent = columnDisplay(col);
  fragment.appendChild(opt);

  return fragment;
}

function fillColumnSelect(selectEl, columns, includeBlank = false, selected = "") {
  clearElement(selectEl);

  if (includeBlank) {
    const blank = document.createElement("option");
    blank.value = "";
    blank.textContent = "(vazio)";
    selectEl.appendChild(blank);
  }

  columns.forEach((col) => {
    const opt = document.createElement("option");
    opt.value = col.letter;
    opt.textContent = columnDisplay(col);
    if (col.letter === selected) opt.selected = true;
    selectEl.appendChild(opt);
  });

  if (includeBlank && !selected) {
    selectEl.value = "";
  }
}

function createFilterRow(defaults = {}) {
  const row = els.filterRowTemplate.content.firstElementChild.cloneNode(true);
  const column = row.querySelector(".filter-column");
  const operator = row.querySelector(".filter-operator");
  const value = row.querySelector(".filter-value");
  const removeBtn = row.querySelector(".remove");

  fillColumnSelect(column, state.source.columns, false, defaults.column || "");
  if (defaults.operator) operator.value = defaults.operator;
  if (defaults.value !== undefined) value.value = defaults.value;

  function updateValueField() {
    const op = operator.value;
    if (op === "empty" || op === "not_empty" || op === "equals_zero" || op === "not_zero") {
      value.disabled = true;
      value.placeholder = "(não aplicável)";
    } else if (op === "greater_than" || op === "less_than") {
      value.disabled = false;
      value.placeholder = "ex: 100";
    } else if (op === "all_in_list") {
      value.disabled = false;
      value.placeholder = "ex: Cartão de Crédito, Dinheiro";
    } else {
      value.disabled = false;
      value.placeholder = "ex: débito";
    }
  }
  operator.addEventListener("change", updateValueField);
  updateValueField();

  removeBtn.addEventListener("click", () => row.remove());

  els.filterList.appendChild(row);
}

function refreshFilterRows() {
  const rows = Array.from(els.filterList.querySelectorAll(".filter-row")).map((row) => {
    const column = row.querySelector(".filter-column").value;
    const operator = row.querySelector(".filter-operator").value;
    const value = row.querySelector(".filter-value").value;
    return { column, operator, value };
  });

  clearElement(els.filterList);
  rows.forEach((r) => createFilterRow(r));
}

function getFiltersPayload() {
  const rows = Array.from(els.filterList.querySelectorAll(".filter-row"));
  const filters = [];

  rows.forEach((row) => {
    const column = row.querySelector(".filter-column").value;
    const operator = row.querySelector(".filter-operator").value;
    const value = row.querySelector(".filter-value").value;

    if (!column || !operator) return;

    const noValueNeeded = ["empty", "not_empty", "equals_zero", "not_zero"];
    if (!noValueNeeded.includes(operator) && value === "") return;

    filters.push({ column, operator, value });
  });

  return filters;
}

function createMappingRow(defaults = {}) {
  const row = els.mappingRowTemplate.content.firstElementChild.cloneNode(true);
  const source1 = row.querySelector(".source-1");
  const source2 = row.querySelector(".source-2");
  const target = row.querySelector(".target");
  const removeBtn = row.querySelector(".remove");

  fillColumnSelect(source1, state.source.columns, false, defaults.source1 || "");
  fillColumnSelect(source2, state.source.columns, true, defaults.source2 || "");
  fillColumnSelect(target, state.dest.columns, false, defaults.target || "");

  removeBtn.addEventListener("click", () => {
    row.remove();
    if (!els.mappingList.children.length) createMappingRow();
  });

  els.mappingList.appendChild(row);
}

function refreshMappingRows() {
  const rows = Array.from(els.mappingList.querySelectorAll(".mapping-row")).map((row) => {
    const source1 = row.querySelector(".source-1").value;
    const source2 = row.querySelector(".source-2").value;
    const target = row.querySelector(".target").value;
    return { source1, source2, target };
  });

  clearElement(els.mappingList);

  if (!rows.length) {
    createMappingRow();
    return;
  }

  rows.forEach((r) => createMappingRow(r));
}

function getMappingsPayload() {
  const rows = Array.from(els.mappingList.querySelectorAll(".mapping-row"));
  const mappings = [];

  rows.forEach((row) => {
    const source1 = row.querySelector(".source-1").value;
    const source2 = row.querySelector(".source-2").value;
    const target = row.querySelector(".target").value;

    if (!source1 || !target) return;

    const sources = source2 ? [source1, source2] : [source1];
    mappings.push({ sources, target });
  });

  return mappings;
}

async function postJson(url, payload) {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });

  const data = await response.json();
  if (!response.ok || !data.ok) {
    throw new Error(data.error || "Erro de comunicacao com o servidor.");
  }
  return data;
}

async function uploadWorkbook(file, kind) {
  const formData = new FormData();
  formData.append("file", file);

  const response = await fetch("/api/upload", {
    method: "POST",
    body: formData,
  });

  const data = await response.json();
  if (!response.ok || !data.ok) {
    throw new Error(data.error || "Falha ao enviar arquivo.");
  }

  state[kind].token = data.token;
  state[kind].filename = data.workbook.filename;
  state[kind].sheets = data.workbook.sheets;
  state[kind].columns = data.workbook.columns;

  if (kind === "source") {
    els.sourceFileName.textContent = data.workbook.filename;
    fillSheetSelect(els.sourceSheet, data.workbook.sheets, data.workbook.default_sheet);
    renderColumnChips(els.sourceColumns, data.workbook.columns);
  } else {
    els.destFileName.textContent = data.workbook.filename;
    fillSheetSelect(els.destSheet, data.workbook.sheets, data.workbook.default_sheet);
    renderColumnChips(els.destColumns, data.workbook.columns);
  }

  refreshMappingRows();
}

async function loadSheetColumns(kind, sheetName) {
  const token = state[kind].token;
  if (!token) return;

  const data = await postJson("/api/sheet-columns", {
    token,
    sheet: sheetName,
  });

  state[kind].columns = data.columns;

  if (kind === "source") {
    renderColumnChips(els.sourceColumns, data.columns);
    refreshFilterRows();
  } else {
    renderColumnChips(els.destColumns, data.columns);
  }

  refreshMappingRows();
}

async function handleRun() {
  try {
    if (!state.source.token || !state.dest.token) {
      throw new Error("Envie as duas planilhas antes de gerar.");
    }

    const mappings = getMappingsPayload();
    if (!mappings.length) {
      throw new Error("Adicione pelo menos um mapeamento valido.");
    }

    const payload = {
      source_token: state.source.token,
      dest_token: state.dest.token,
      source_sheet: els.sourceSheet.value,
      dest_sheet: els.destSheet.value,
      source_start_row: Number(els.sourceStartRow.value || 2),
      dest_start_row: Number(els.destStartRow.value || 2),
      separator: els.separator.value,
      skip_empty_rows: els.skipEmpty.checked,
      filters: getFiltersPayload(),
      mappings,
    };

    setStatus("Processando transferencia...", "info");
    els.runBtn.disabled = true;

    const response = await fetch("/api/transfer", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      const data = await response.json();
      throw new Error(data.error || "Falha na transferencia.");
    }

    const blob = await response.blob();
    const disposition = response.headers.get("Content-Disposition") || "";
    const fileNameMatch = disposition.match(/filename\*=UTF-8''([^;]+)|filename=\"?([^\";]+)\"?/i);
    const fileName = decodeURIComponent(fileNameMatch?.[1] || fileNameMatch?.[2] || "planilha_atualizada.xlsx");

    const rows = response.headers.get("X-Copied-Rows") || "0";
    const cells = response.headers.get("X-Copied-Cells") || "0";

    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);

    setStatus(`Concluido: ${rows} linhas e ${cells} celulas copiadas.`, "ok");
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    els.runBtn.disabled = false;
  }
}

els.sourceFile.addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) return;
  try {
    setStatus("Carregando planilha de origem...", "info");
    await uploadWorkbook(file, "source");
    setStatus("Planilha de origem carregada.", "ok");
  } catch (error) {
    setStatus(error.message, "error");
  }
});

els.destFile.addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) return;
  try {
    setStatus("Carregando planilha de destino...", "info");
    await uploadWorkbook(file, "dest");
    setStatus("Planilha de destino carregada.", "ok");
  } catch (error) {
    setStatus(error.message, "error");
  }
});

els.sourceSheet.addEventListener("change", async () => {
  try {
    setStatus("Atualizando colunas da aba de origem...", "info");
    await loadSheetColumns("source", els.sourceSheet.value);
    setStatus("Colunas de origem atualizadas.", "ok");
  } catch (error) {
    setStatus(error.message, "error");
  }
});

els.destSheet.addEventListener("change", async () => {
  try {
    setStatus("Atualizando colunas da aba de destino...", "info");
    await loadSheetColumns("dest", els.destSheet.value);
    setStatus("Colunas de destino atualizadas.", "ok");
  } catch (error) {
    setStatus(error.message, "error");
  }
});

els.addMappingBtn.addEventListener("click", () => createMappingRow());
els.addFilterBtn.addEventListener("click", () => createFilterRow());
els.runBtn.addEventListener("click", handleRun);

createMappingRow();
