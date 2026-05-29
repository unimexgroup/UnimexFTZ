/* Unimex FTZ Processor — browser front-end. */

const statusEl      = document.getElementById("status");
const dropzone      = document.getElementById("dropzone");
const fileInput     = document.getElementById("file-input");
const pickerBtn     = document.getElementById("picker-btn");
const resultsEl     = document.getElementById("results");
const runInfoEl     = document.getElementById("run-info");
const downloadAllBtn= document.getElementById("download-all");
const modal         = document.getElementById("preview-modal");
const modalTitle    = document.getElementById("preview-title");
const modalBody     = document.getElementById("preview-body");
const modalClose    = document.getElementById("preview-close");

let pyodide = null;
let pyReady = null;
let currentResults = []; // last successful results: [{shipment_id, filename, xlsx, ...}]

const PYODIDE_VERSION = "v0.26.4";

function setStatus(msg, kind = "info") {
  statusEl.textContent = msg;
  statusEl.dataset.kind = kind;
}

async function bootPyodide() {
  try {
    setStatus("Loading Python runtime (one-time, ~20s)...");
    pyodide = await loadPyodide({
      indexURL: `https://cdn.jsdelivr.net/pyodide/${PYODIDE_VERSION}/full/`,
    });
    setStatus("Installing pandas + openpyxl...");
    await pyodide.loadPackage(["pandas", "openpyxl"]);
    setStatus("Loading processor...");
    const [pyMain, pyWeb] = await Promise.all([
      fetch("ftz_processor.py", { cache: "no-cache" }).then(r => r.text()),
      fetch("web_processor.py", { cache: "no-cache" }).then(r => r.text()),
    ]);
    pyodide.FS.writeFile("ftz_processor.py", pyMain);
    pyodide.FS.writeFile("web_processor.py", pyWeb);
    await pyodide.runPythonAsync("import web_processor");
    setStatus("Ready. Drop your files.", "ok");
    dropzone.classList.add("ready");
  } catch (e) {
    console.error(e);
    setStatus("Failed to load Python runtime — check your connection and refresh.", "err");
  }
}

pyReady = bootPyodide();

/* ------------------------------------------------------------ */
/* Drag & drop                                                   */
/* ------------------------------------------------------------ */
function preventDefaults(e) { e.preventDefault(); e.stopPropagation(); }

["dragenter", "dragover"].forEach(evt => {
  dropzone.addEventListener(evt, e => {
    preventDefaults(e);
    dropzone.classList.add("dragover");
  });
});
["dragleave", "drop"].forEach(evt => {
  dropzone.addEventListener(evt, e => {
    preventDefaults(e);
    dropzone.classList.remove("dragover");
  });
});
dropzone.addEventListener("drop", e => {
  const files = [...(e.dataTransfer?.files || [])];
  if (files.length) handleFiles(files);
});

pickerBtn.addEventListener("click", () => fileInput.click());
fileInput.addEventListener("change", () => {
  if (fileInput.files.length) handleFiles([...fileInput.files]);
  fileInput.value = "";
});

/* ------------------------------------------------------------ */
/* Processing                                                    */
/* ------------------------------------------------------------ */
async function handleFiles(files) {
  await pyReady;
  if (!pyodide) return;

  const xlsx = files.filter(f => /\.(xlsx|xls|xlsm)$/i.test(f.name) && !f.name.startsWith("~$"));
  if (!xlsx.length) {
    setStatus("No Excel files in that drop.", "warn");
    return;
  }

  setStatus(`Processing ${xlsx.length} file(s)...`);
  resultsEl.innerHTML = `<p class="empty">Working...</p>`;
  downloadAllBtn.hidden = true;
  runInfoEl.hidden = true;

  // Convert files to a JS array Pyodide will marshal as list[dict[str, bytes]]
  const pyFiles = [];
  for (const f of xlsx) {
    const ab = await f.arrayBuffer();
    pyFiles.push({ name: f.name, data: new Uint8Array(ab) });
  }

  try {
    pyodide.globals.set("_js_files", pyodide.toPy(pyFiles));
    const proxy = await pyodide.runPythonAsync("web_processor.process_files(_js_files)");
    const result = proxy.toJs({ dict_converter: Object.fromEntries });
    proxy.destroy();
    renderResults(result);
    setStatus(
      `Done. ${result.results.length} processed, ${result.skipped.length} skipped, ${result.unrecognized.length} unrecognized.`,
      "ok"
    );
  } catch (e) {
    console.error(e);
    setStatus(`Error: ${e.message || e}`, "err");
    resultsEl.innerHTML = `<p class="empty err">Processing failed. See browser console for details.</p>`;
  }
}

/* ------------------------------------------------------------ */
/* Rendering                                                     */
/* ------------------------------------------------------------ */
function renderResults({ results, skipped, unrecognized }) {
  currentResults = results;
  resultsEl.innerHTML = "";

  if (!results.length && !skipped.length && !unrecognized.length) {
    resultsEl.innerHTML = `<p class="empty">Nothing processed.</p>`;
    return;
  }

  for (const r of results) {
    resultsEl.appendChild(buildResultCard(r));
  }
  for (const s of skipped) {
    resultsEl.appendChild(buildSkippedCard(s));
  }
  for (const u of unrecognized) {
    resultsEl.appendChild(buildUnrecognizedCard(u));
  }

  downloadAllBtn.hidden = results.length < 2;
}

function buildResultCard(r) {
  const card = document.createElement("article");
  card.className = "card card-ok";

  const head = document.createElement("div");
  head.className = "card-head";
  head.innerHTML = `<h3>${escapeHtml(r.shipment_id)}</h3>
                    <span class="badge ok">OK</span>`;
  card.appendChild(head);

  const summary = document.createElement("p");
  summary.className = "summary";
  summary.textContent = r.summary;
  card.appendChild(summary);

  if (r.warnings && r.warnings.length) {
    const w = document.createElement("ul");
    w.className = "warnings";
    for (const msg of r.warnings) {
      const li = document.createElement("li");
      li.textContent = msg;
      w.appendChild(li);
    }
    card.appendChild(w);
  }

  const actions = document.createElement("div");
  actions.className = "actions";

  const viewBtn = document.createElement("button");
  viewBtn.type = "button";
  viewBtn.className = "btn btn-secondary";
  viewBtn.textContent = "View";
  viewBtn.addEventListener("click", () => openPreview(r));
  actions.appendChild(viewBtn);

  const dlBtn = document.createElement("button");
  dlBtn.type = "button";
  dlBtn.className = "btn btn-primary";
  dlBtn.textContent = "Download .xlsx";
  dlBtn.addEventListener("click", () => downloadXlsx(r));
  actions.appendChild(dlBtn);

  card.appendChild(actions);
  return card;
}

function buildSkippedCard(s) {
  const card = document.createElement("article");
  card.className = "card card-skip";
  card.innerHTML = `<div class="card-head">
      <h3>${escapeHtml(s.shipment_id || "(unknown shipment)")}</h3>
      <span class="badge skip">Skipped</span>
    </div>
    <p class="summary">${escapeHtml(s.reason)}</p>`;
  return card;
}

function buildUnrecognizedCard(u) {
  const card = document.createElement("article");
  card.className = "card card-unrec";
  card.innerHTML = `<div class="card-head">
      <h3>${escapeHtml(u.name)}</h3>
      <span class="badge unrec">Not recognized</span>
    </div>
    <p class="summary">${escapeHtml(u.reason)}</p>`;
  return card;
}

/* ------------------------------------------------------------ */
/* Download                                                      */
/* ------------------------------------------------------------ */
function downloadXlsx(r) {
  const blob = new Blob([r.xlsx], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  triggerDownload(blob, r.filename);
}

downloadAllBtn.addEventListener("click", async () => {
  if (!currentResults.length) return;
  downloadAllBtn.disabled = true;
  const orig = downloadAllBtn.textContent;
  downloadAllBtn.textContent = "Zipping...";
  try {
    const zip = new JSZip();
    for (const r of currentResults) {
      zip.file(r.filename, r.xlsx);
    }
    const blob = await zip.generateAsync({ type: "blob" });
    const stamp = new Date().toISOString().slice(0, 10);
    triggerDownload(blob, `FTZ_${stamp}.zip`);
  } finally {
    downloadAllBtn.disabled = false;
    downloadAllBtn.textContent = orig;
  }
});

function triggerDownload(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/* ------------------------------------------------------------ */
/* Preview modal                                                 */
/* ------------------------------------------------------------ */
function openPreview(r) {
  modalTitle.textContent = `${r.shipment_id} — ${r.preview_rows.length} HS code(s)`;
  modalBody.innerHTML = "";

  const table = document.createElement("table");
  table.className = "preview-table";
  const thead = document.createElement("thead");
  const trh = document.createElement("tr");
  for (const h of r.preview_headers) {
    const th = document.createElement("th");
    th.textContent = h;
    trh.appendChild(th);
  }
  thead.appendChild(trh);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  for (const row of r.preview_rows) {
    const tr = document.createElement("tr");
    row.forEach((cell, i) => {
      const td = document.createElement("td");
      td.textContent = formatCell(r.preview_headers[i], cell);
      if (typeof cell === "number") td.classList.add("num");
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);
  modalBody.appendChild(table);

  if (typeof modal.showModal === "function") modal.showModal();
  else modal.setAttribute("open", "");
}

function formatCell(header, value) {
  if (value === null || value === undefined || value === "") return "";
  if (typeof value !== "number") return String(value);
  if (header === "Value") return "$" + value.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  if (header === "Weight") return value.toLocaleString(undefined, { minimumFractionDigits: 3, maximumFractionDigits: 3 });
  if (header === "Quantity" || header === "Charges") return value.toLocaleString();
  return String(value);
}

modalClose.addEventListener("click", () => {
  if (typeof modal.close === "function") modal.close();
  else modal.removeAttribute("open");
});
modal.addEventListener("click", e => {
  if (e.target === modal) modalClose.click();
});

/* ------------------------------------------------------------ */
/* Auto-derive Source link from GitHub Pages URL                 */
(() => {
  const link = document.getElementById("repo-link");
  if (!link) return;
  const host = location.hostname;
  const m = host.match(/^([^.]+)\.github\.io$/);
  if (m) {
    const user = m[1];
    const repo = location.pathname.split("/").filter(Boolean)[0];
    if (repo) link.href = `https://github.com/${user}/${repo}`;
    else      link.href = `https://github.com/${user}`;
  }
})();

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  })[c]);
}
