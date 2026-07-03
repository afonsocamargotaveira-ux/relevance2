const $ = (sel) => document.querySelector(sel);

const fileInput = $("#fileInput");
const runBtn = $("#runBtn");
const statusEl = $("#status");
const introEl = $("#intro");
const columnConfigEl = $("#columnConfig");
const resultsEl = $("#results");

let detectedColumns = null; // resposta de /api/columns
let headerLabels = [];

// ---------- helpers de UI ----------

function showAlert(type, html) {
  statusEl.innerHTML = `<div class="alert alert-${type}">${html}</div>`;
}

function clearAlert() {
  statusEl.innerHTML = "";
}

function setLoading(isLoading, label) {
  runBtn.disabled = isLoading;
  runBtn.innerHTML = isLoading
    ? `<span class="spinner"></span>${label || "Processando…"}`
    : "▶ Avaliar Proposições";
}

function buildSelect(selectEl, options, selectedValue, allowNone) {
  selectEl.innerHTML = "";
  if (allowNone) {
    const opt = document.createElement("option");
    opt.value = "";
    opt.textContent = "— nenhuma —";
    selectEl.appendChild(opt);
  }
  options.forEach((label, i) => {
    const opt = document.createElement("option");
    opt.value = i;
    opt.textContent = `${i} — ${label.slice(0, 30)}`;
    selectEl.appendChild(opt);
  });
  selectEl.value = (selectedValue === null || selectedValue === undefined) ? "" : String(selectedValue);
}

// ---------- termos (contagem ao vivo, espelha parse_terms no back-end de forma simplificada) ----------

function countTermsPreview(raw) {
  if (!raw || !raw.trim()) return 0;
  // contagem aproximada só para feedback visual rápido (a contagem real acontece no servidor)
  const quoted = (raw.match(/["“”‘’«»][^"“”‘’«»]+["“”‘’«»]/g) || []).length;
  const rest = raw.replace(/["“”‘’«»][^"“”‘’«»]+["“”‘’«»]/g, " ");
  const words = rest.split(/[\s,;()]+/).filter(Boolean);
  return quoted + words.length;
}

function updateTermCount() {
  const theme = $("#mainTheme").value;
  const extras = $("#extraTerms").value;
  if (!theme.trim()) {
    $("#termCount").textContent = "";
    return;
  }
  const tt = countTermsPreview(theme);
  const et = countTermsPreview(extras);
  $("#termCount").innerHTML = `🔍 <b>${tt}</b> termo(s) no tema · <b>${et}</b> termos adicionais (aprox.)`;
}
$("#mainTheme").addEventListener("input", updateTermCount);
$("#extraTerms").addEventListener("input", updateTermCount);

// ---------- fluxo: upload -> detectar colunas ----------

fileInput.addEventListener("change", async () => {
  resultsEl.classList.add("hidden");
  columnConfigEl.classList.add("hidden");
  clearAlert();
  detectedColumns = null;

  const file = fileInput.files[0];
  if (!file) return;

  setLoading(true, "Lendo planilha…");
  try {
    const fd = new FormData();
    fd.append("file", file);
    const resp = await fetch("/api/columns", { method: "POST", body: fd });
    const data = await resp.json();
    if (!resp.ok) {
      showAlert("error", data.error || "Erro ao ler o arquivo.");
      return;
    }
    detectedColumns = data;
    headerLabels = data.header;

    buildSelect($("#colEmenta"), headerLabels, data.col_ementa, false);
    buildSelect($("#colIndexa"), headerLabels, data.col_indexa, true);
    buildSelect($("#colExtraText"), headerLabels, data.col_extra_text, true);
    buildSelect($("#colTeor"), headerLabels, data.col_teor, true);

    columnConfigEl.classList.remove("hidden");
    introEl.classList.add("hidden");

    if (data.col_teor !== null && data.col_teor !== undefined) {
      if (data.n_teor_links > 0) {
        showAlert("info", `🔗 ${data.n_teor_links} links do Inteiro Teor detectados — proposições sem aderência na ementa/indexação serão consultadas → índice 2.`);
      } else {
        showAlert("warning", "⚠️ Coluna de Inteiro Teor detectada, mas nenhum hyperlink foi encontrado no arquivo.");
      }
    } else {
      clearAlert();
    }
  } catch (e) {
    showAlert("error", `Erro ao processar o arquivo: ${e}`);
  } finally {
    setLoading(false);
  }
});

// ---------- fluxo: avaliar ----------

runBtn.addEventListener("click", async () => {
  clearAlert();
  const file = fileInput.files[0];
  if (!file) {
    showAlert("warning", "⚠️ Faça upload de uma planilha .xlsx.");
    return;
  }

  const colEmenta = $("#colEmenta").value;
  if (colEmenta === "") {
    showAlert("error", "Selecione ao menos a coluna da Ementa.");
    return;
  }

  const fd = new FormData();
  fd.append("file", file);
  fd.append("main_theme", $("#mainTheme").value);
  fd.append("extra_terms", $("#extraTerms").value);
  fd.append("exclude_terms", $("#excludeTerms").value);
  fd.append("exclude_scope", document.querySelector('input[name="excludeScope"]:checked').value);
  fd.append("fetch_teor", $("#fetchTeor").checked ? "true" : "false");
  fd.append("col_ementa", colEmenta);
  fd.append("col_indexa", $("#colIndexa").value);
  fd.append("col_extra_text", $("#colExtraText").value);
  fd.append("col_teor", $("#colTeor").value);

  setLoading(true, "Avaliando proposições…");
  resultsEl.classList.add("hidden");

  try {
    const resp = await fetch("/api/process", { method: "POST", body: fd });
    const data = await resp.json();
    if (!resp.ok) {
      showAlert("error", data.error || "Erro ao avaliar as proposições.");
      return;
    }
    renderResults(data);
    clearAlert();
  } catch (e) {
    showAlert("error", `Erro ao avaliar: ${e}`);
  } finally {
    setLoading(false);
  }
});

// ---------- renderização de resultados ----------

const scoreColors = { 1: "#e74c3c", 2: "#e67e22", 3: "#f1c40f", 4: "#2ecc71", 5: "#1a6fd4" };

function renderResults(data) {
  const { propositions, scores, justifications, counts, average, excel_base64,
          n_teor_links, teor_checked, teor_available } = data;

  // resumo
  const summaryCards = $("#summaryCards");
  summaryCards.innerHTML = "";
  for (let s = 1; s <= 5; s++) {
    summaryCards.innerHTML += `
      <div class="result-card">
        <div class="value" style="color:${scoreColors[s]}">${counts[s] || 0}</div>
        <div class="label">Índice ${s}</div>
      </div>`;
  }
  summaryCards.innerHTML += `
    <div class="result-card" style="border-color:#1a6fd4;">
      <div class="value" style="color:#1a6fd4">${average.toFixed(1)}</div>
      <div class="label">Média</div>
    </div>`;

  // tabela
  const tbody = document.querySelector("#resultsTable tbody");
  tbody.innerHTML = "";
  const frag = document.createDocumentFragment();
  for (let i = 0; i < propositions.length; i++) {
    const tr = document.createElement("tr");
    const s = scores[i];
    tr.innerHTML = `
      <td>${escapeHtml(propositions[i])}</td>
      <td><span class="score-badge score-${s}">${s}</span></td>
      <td>${escapeHtml(justifications[i])}</td>`;
    frag.appendChild(tr);
  }
  tbody.appendChild(frag);

  // download
  const blob = b64ToBlob(excel_base64, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  const url = URL.createObjectURL(blob);
  const downloadBtn = $("#downloadBtn");
  downloadBtn.href = url;

  if (n_teor_links > 0 && teor_checked === 0 && teor_available === false) {
    showAlert("warning", `⚠️ Não foi possível acessar o servidor da Câmara para consultar o Inteiro Teor a partir desta função serverless. Rode a etapa de Inteiro Teor localmente, se necessário.`);
  }

  resultsEl.classList.remove("hidden");
}

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML;
}

function b64ToBlob(b64, mime) {
  const byteChars = atob(b64);
  const byteNumbers = new Array(byteChars.length);
  for (let i = 0; i < byteChars.length; i++) {
    byteNumbers[i] = byteChars.charCodeAt(i);
  }
  const byteArray = new Uint8Array(byteNumbers);
  return new Blob([byteArray], { type: mime });
}
