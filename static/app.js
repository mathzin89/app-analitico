// =======================
// Opções marcadas / NIW
// =======================
function selectedOptions() {
  return Array.from(document.querySelectorAll('.options input[type=checkbox]:checked')).map(cb => cb.value);
}
function niwMode() {
  const r = document.querySelector('.options input[name=niw]:checked');
  return r ? r.value : "none";
}

// =======================
// Filtro digitável
// =======================
function filterSelect(selectEl, query){
  const q = (query || "").toLowerCase();
  Array.from(selectEl.options).forEach(opt => { opt.hidden = !opt.text.toLowerCase().includes(q); });
  if (selectEl.selectedOptions.length && selectEl.selectedOptions[0].hidden){ selectEl.selectedIndex = -1; }
}
function attachFilter(inputId, selectId){
  const input = document.getElementById(inputId);
  const select = document.getElementById(selectId);
  input.addEventListener("input", () => filterSelect(select, input.value));
  input.addEventListener("keydown", (e) => { if(e.key === "Enter"){ select.focus(); }});
}

// =======================
// Tema claro/escuro
// =======================
function applyTheme(theme){
  document.body.classList.toggle('theme-light', theme === 'light');
  localStorage.setItem('theme', theme);
}
document.addEventListener('DOMContentLoaded', () => {
  applyTheme(localStorage.getItem('theme') || 'dark');
});
document.getElementById('themeToggle').addEventListener('click', () => {
  const next = document.body.classList.contains('theme-light') ? 'dark' : 'light';
  applyTheme(next);
});

// =======================
// Status cabeçalho
// =======================
function setStatusOnline(filename){
  const s = document.getElementById("liveStatus"); const t = document.getElementById("fileStatus");
  s.classList.remove("offline"); s.classList.add("online"); t.textContent = `Online — ${filename}`;
}
function setStatusOffline(){
  const s = document.getElementById("liveStatus"); const t = document.getElementById("fileStatus");
  s.classList.remove("online"); s.classList.add("offline"); t.textContent = "Aguardando arquivo…";
}

// =======================
// Chips
// =======================
function renderChips(selectId, chipsId){
  const vals = Array.from(document.getElementById(selectId).selectedOptions).map(o => o.value);
  const box = document.getElementById(chipsId); box.innerHTML = "";
  vals.forEach(v => {
    const chip = document.createElement("button");
    chip.className = "chip"; chip.type = "button"; chip.title = "Remover"; chip.textContent = v + " ×";
    chip.addEventListener("click", () => {
      const sel = document.getElementById(selectId);
      Array.from(sel.options).forEach(o => { if(o.value === v) o.selected = false; });
      renderChips(selectId, chipsId);
    });
    box.appendChild(chip);
  });
}
function selectVisible(selectId, chipsId){
  const sel = document.getElementById(selectId);
  Array.from(sel.options).forEach(o => { if(!o.hidden) o.selected = true; });
  renderChips(selectId, chipsId);
}
function clearSelection(selectId, chipsId){
  const sel = document.getElementById(selectId);
  Array.from(sel.options).forEach(o => o.selected = false);
  renderChips(selectId, chipsId);
}

// =======================
// Upload / Close
// =======================
document.getElementById("uploadForm").addEventListener("submit", async (e) => {
  e.preventDefault();
  const file = document.getElementById("fileInput").files[0];
  const msg = document.getElementById("uploadMsg");
  const clearBtn = document.getElementById("clearFileBtn");
  if (!file) { msg.textContent = "Selecione um arquivo .sav."; return; }
  const fd = new FormData(); fd.append("file", file);
  try {
    const res = await fetch("/upload", { method: "POST", body: fd });
    const payload = await res.json();
    if (!res.ok) throw new Error(payload.error || "Falha no upload.");
    msg.textContent = payload.message;

    const rowSel = document.getElementById("rowSelect");
    const colSel = document.getElementById("colSelect");
    const varsList = document.getElementById("varsList");
    rowSel.innerHTML = ""; colSel.innerHTML = ""; varsList.innerHTML = "";

    const names = payload.variables.map(v => v.name).sort((a,b)=> a.localeCompare(b,'pt-BR'));
    names.forEach(name => {
      rowSel.add(new Option(name, name));
      colSel.add(new Option(name, name));
      const opt = document.createElement("option"); opt.value = name; varsList.appendChild(opt);
    });
    if (payload.suggested_weight) document.getElementById("weightInput").value = payload.suggested_weight;

    attachFilter("rowSearch", "rowSelect");
    attachFilter("colSearch", "colSelect");
    clearBtn.style.display = "inline-block";
    setStatusOnline(payload.filename || "arquivo carregado");
  } catch (err) {
    msg.textContent = err.message;
  }
});

document.getElementById("clearFileBtn").addEventListener("click", async () => {
  const msg = document.getElementById("uploadMsg");
  const results = document.getElementById("results");
  try {
    const res = await fetch("/close", { method: "POST" });
    const payload = await res.json();
    if (!res.ok) throw new Error(payload.error || "Falha ao fechar.");
    msg.textContent = payload.message;
    document.getElementById("rowSelect").innerHTML = "";
    document.getElementById("colSelect").innerHTML = "";
    document.getElementById("rowChips").innerHTML = "";
    document.getElementById("colChips").innerHTML = "";
    results.innerHTML = "";
    document.getElementById("clearFileBtn").style.display = "none";
    document.getElementById("fileInput").value = "";
    setStatusOffline();
  } catch (e) {
    msg.textContent = e.message;
  }
});

// =======================
// Listeners + util
// =======================
["rowSelect","colSelect"].forEach(id=>{
  document.getElementById(id).addEventListener("change", () => {
    renderChips("rowSelect","rowChips");
    renderChips("colSelect","colChips");
  });
});
document.getElementById("rowSelectAll").addEventListener("click", ()=>selectVisible("rowSelect","rowChips"));
document.getElementById("colSelectAll").addEventListener("click", ()=>selectVisible("colSelect","colChips"));
document.getElementById("rowClear").addEventListener("click", ()=>clearSelection("rowSelect","rowChips"));
document.getElementById("colClear").addEventListener("click", ()=>clearSelection("colSelect","colChips"));

function getSelectedValues(selectId){
  const sel = document.getElementById(selectId);
  return Array.from(sel.selectedOptions).map(o => o.value);
}

// =======================
// Render: célula com linhas e destaque pelo flagMask
// =======================
function renderCellLines(lines, flagged){
  if (!Array.isArray(lines)) return lines ?? "";
  return lines.map(line => {
    if (flagged && /^Adjusted Standardized Residual:/i.test(line)) {
      const num = line.split(":").slice(1).join(":").trim();
      return `<div>Adjusted Standardized Residual: <span style="background:#fde68a;padding:0 4px;border-radius:4px;font-weight:700">${num}</span></div>`;
    }
    return `<div>${line}</div>`;
  }).join("");
}

// =======================
// Gerar tabelas + habilitar export
// =======================
document.getElementById("runBtn").addEventListener("click", async () => {
  const rows = getSelectedValues("rowSelect");
  const cols = getSelectedValues("colSelect");
  const weight = document.getElementById("weightInput").value || "peso";
  const options = selectedOptions();
  const mode = niwMode();
  const err = document.getElementById("errorMsg");
  err.textContent = "";
  if (!rows.length || !cols.length) { err.textContent = "Selecione ao menos uma variável para linha e coluna."; return; }

  try {
    const res = await fetch("/crosstab", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ rows, cols, weight, options, niw_mode: mode })
    });
    const payload = await res.json();
    if (!res.ok) throw new Error(payload.error || "Erro ao gerar tabela.");

    const resultsDiv = document.getElementById("results");
    resultsDiv.innerHTML = "";
    const tables = payload.tables || { "Tabela": payload.table };

    Object.entries(tables).forEach(([key, t]) => {
      const h3 = document.createElement("h3");
      h3.textContent = t.title || key;
      resultsDiv.appendChild(h3);

      const wrap = document.createElement("div"); wrap.className = "table-wrapper";
      const table = document.createElement("table");

      const thead = document.createElement("thead");
      thead.innerHTML = "<tr><th></th>" + t.columns.map(c => `<th>${c}</th>`).join("") + "</tr>";
      table.appendChild(thead);

      const tbody = document.createElement("tbody");
      t.data.forEach((rowVals, i) => {
        const tr = document.createElement("tr");
        tr.innerHTML = `<th>${t.index[i]}</th>` + rowVals.map((v, j) => {
          const flagged = t.flagMask ? !!t.flagMask[i][j] : false;
          return `<td>${renderCellLines(v, flagged)}</td>`;
        }).join("");
        tbody.appendChild(tr);
      });
      table.appendChild(tbody);
      wrap.appendChild(table);
      resultsDiv.appendChild(wrap);
    });

    document.getElementById("exportBtn").disabled = false;
  } catch (e) {
    err.textContent = e.message;
  }
});

// =======================
// Export Excel
// =======================
document.getElementById("exportBtn").addEventListener("click", () => {
  window.location.href = "/export_excel";
});

// =======================
// Defaults
// =======================
document.addEventListener("DOMContentLoaded", () => {
  const observed = document.querySelector('input[value="observed"]'); if (observed) observed.checked = true;
  const colPct = document.querySelector('input[value="col_pct"]'); if (colPct) colPct.checked = true;
  const adjStd = document.querySelector('input[value="adj_std_resid"]'); if (adjStd) adjStd.checked = true;
  const roundCell = document.querySelector('input[name="niw"][value="round_cell"]'); if (roundCell) roundCell.checked = true;
});
