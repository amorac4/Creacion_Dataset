const state = {
  summary: null,
  queueOffset: 0,
  queueLimit: 100,
};

const $ = (id) => document.getElementById(id);

function fmt(n) {
  return Number(n || 0).toLocaleString("es-MX");
}

function toast(message) {
  const el = $("toast");
  el.textContent = message;
  el.classList.remove("hidden");
  setTimeout(() => el.classList.add("hidden"), 3500);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

async function api(path, options = {}) {
  const res = await fetch(path, {
    headers: { "Content-Type": "application/json" },
    ...options,
  });
  const data = await res.json();
  if (!res.ok || data.error) throw new Error(data.error || res.statusText);
  return data;
}

function table(columns, rows) {
  const head = columns.map(([key, label]) => `<th>${label || key}</th>`).join("");
  const body = rows.map((row) => {
    const cells = columns.map(([key, _label, render]) => {
      const value = render ? render(row) : row[key];
      return `<td>${value ?? ""}</td>`;
    }).join("");
    return `<tr>${cells}</tr>`;
  }).join("");
  return `<table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>`;
}

function listRows(rows) {
  return rows.map((r) => `
    <div class="listRow">
      <span>${r.name}</span>
      <strong>${fmt(r.count)}</strong>
    </div>
  `).join("");
}

async function loadSummary() {
  const data = await api("/api/summary");
  state.summary = data;
  $("dbPath").textContent = data.db;
  $("cards").innerHTML = [
    ["Muestras", data.totals.muestras],
    ["Requieren revisión", data.totals.requieren_revision],
    ["Pendientes", data.totals.pendientes],
    ["Revisiones hash", data.totals.revisiones_hash],
    ["Revisiones par", data.totals.revisiones_par],
  ].map(([label, value]) => `
    <div class="card"><div class="muted">${label}</div><div class="value">${fmt(value)}</div></div>
  `).join("");
  $("confidenceList").innerHTML = listRows(data.confidence);
  $("finalList").innerHTML = listRows(data.final_counts);
  $("topPairs").innerHTML = renderPairsTable(data.top_pairs);
}

function pairActions(row) {
  const local = encodeURIComponent(row.familia_local);
  const av = encodeURIComponent(row.familia_avclass);
  return `
    <div class="actions">
      <button onclick="reviewPair('${local}','${av}','accept_local')">Local</button>
      <button onclick="reviewPair('${local}','${av}','accept_avclass')">AVClass</button>
      <button onclick="reviewPairManual('${local}','${av}')">Manual</button>
    </div>
  `;
}

function renderPairsTable(rows) {
  return table([
    ["prioridad_revision", "Prio"],
    ["familia_local", "Local"],
    ["familia_avclass", "AVClass"],
    ["muestras", "Muestras", (r) => fmt(r.muestras)],
    ["ejemplo_hash", "Hash ejemplo", (r) => r.ejemplo_hash ? hashButton(r.ejemplo_hash) : ""],
    ["familia_final_sugerida", "Sugerida"],
    ["confianza_sugerida", "Confianza"],
    ["detecciones_top", "Detecciones top", (r) => {
      const text = escapeHtml(r.detecciones_top || "");
      return `<div class="detectionsCell" title="${text}">${text}</div>`;
    }],
    ["acciones", "Acciones", pairActions],
  ], rows);
}

async function loadPairs() {
  const q = encodeURIComponent($("pairSearch").value || "");
  const review = $("onlyReviewPairs").checked ? "1" : "0";
  const data = await api(`/api/pairs?q=${q}&review=${review}&limit=300`);
  $("pairsTable").innerHTML = renderPairsTable(data.rows);
}

function hashButton(hash) {
  return `<button class="hashBtn" onclick="openHash('${hash}')">${hash}</button>`;
}

function queueSingleEnabled() {
  return Boolean($("queueSingle")?.checked);
}

function queueActions(row) {
  const hash = encodeURIComponent(row.hash_md5);
  return `
    <div class="actions">
      <button onclick="reviewHash('${hash}','accept_local')">Local</button>
      <button onclick="reviewHash('${hash}','accept_avclass')">AVClass</button>
      <button onclick="reviewHash('${hash}','accept_suggested')">Sugerida</button>
      <button onclick="reviewHashManual('${hash}')">Manual</button>
      <button onclick="reviewHash('${hash}','sin_inferir')">Sin inferir</button>
    </div>
  `;
}

function renderQueueCard(row) {
  if (!row) {
    return `<div class="emptyState">No hay muestras para esos filtros.</div>`;
  }
  const hash = encodeURIComponent(row.hash_md5);
  const detections = escapeHtml(row.detecciones_top || "");
  return `
    <div class="reviewCard">
      <div class="reviewHeader">
        <div>
          <div class="muted">Prioridad ${escapeHtml(row.prioridad_revision)} - ${escapeHtml(row.lote_origen)}</div>
          <h2>${hashButton(row.hash_md5)}</h2>
        </div>
        <div class="badge">${escapeHtml(row.confianza_sugerida)}</div>
      </div>
      <div class="reviewLabels">
        <div><span>Local</span><strong>${escapeHtml(row.familia_local)}</strong></div>
        <div><span>AVClass</span><strong>${escapeHtml(row.familia_avclass)}</strong></div>
        <div><span>Sugerida</span><strong>${escapeHtml(row.familia_final_sugerida)}</strong></div>
      </div>
      <div class="reviewEvidence">
        <div>
          <h3>Clases AVClass</h3>
          <div class="pre compact">${escapeHtml(row.clases_avclass || "")}</div>
        </div>
        <div>
          <h3>Detecciones top</h3>
          <div class="detectionsCell wide" title="${detections}">${detections}</div>
        </div>
      </div>
      <div class="actions reviewActions">
        <button onclick="reviewHash('${hash}','accept_local')">Local</button>
        <button onclick="reviewHash('${hash}','accept_avclass')">AVClass</button>
        <button onclick="reviewHash('${hash}','accept_suggested')">Sugerida</button>
        <button onclick="reviewHashManual('${hash}')">Manual</button>
        <button onclick="reviewHash('${hash}','sin_inferir')">Sin inferir</button>
        <button onclick="openHash('${hash}')">Ver detalle</button>
      </div>
    </div>
  `;
}

function renderQueueTable(rows) {
  return table([
    ["prioridad_revision", "Prio"],
    ["hash_md5", "Hash", (r) => hashButton(r.hash_md5)],
    ["lote_origen", "Lote"],
    ["familia_local", "Local"],
    ["familia_avclass", "AVClass"],
    ["familia_final_sugerida", "Sugerida"],
    ["confianza_sugerida", "Confianza"],
    ["clases_avclass", "Clases"],
    ["acciones", "Acciones", queueActions],
  ], rows);
}

async function loadQueue(offset = 0) {
  state.queueOffset = offset;
  const status = $("queueStatus").value;
  const priority = $("queuePriority").value;
  const q = encodeURIComponent($("queueSearch").value || "");
  const single = queueSingleEnabled();
  const limit = single ? 1 : state.queueLimit;
  const data = await api(`/api/queue?status=${status}&priority=${priority}&q=${q}&limit=${limit}&offset=${offset}`);
  $("queueMeta").innerHTML = single
    ? `Muestra ${data.rows.length ? fmt(offset + 1) : "0"} de ${fmt(data.total)}.`
    : `Mostrando ${fmt(data.rows.length)} de ${fmt(data.total)}. Offset ${fmt(offset)}.`;
  $("queueTable").classList.toggle("singleQueue", single);
  $("queueTable").innerHTML = single ? renderQueueCard(data.rows[0]) : renderQueueTable(data.rows);
  $("queueNextBtn").disabled = !single || offset + 1 >= data.total;
}

async function reviewHash(hash, action, manual = "", note = "") {
  const data = await api("/api/review/hash", {
    method: "POST",
    body: JSON.stringify({ hash_md5: decodeURIComponent(hash), action, familia_manual: manual, note }),
  });
  toast(`Hash revisado: ${data.hash_md5} -> ${data.familia_final}`);
  await loadSummary();
  await loadQueue(queueSingleEnabled() ? 0 : state.queueOffset);
}

async function reviewHashManual(hash) {
  const value = prompt("Familia manual:");
  if (!value) return;
  const note = prompt("Nota de revisión:", "") || "";
  await reviewHash(hash, "manual", value, note);
}

async function reviewPair(local, avclass, action, manual = "", note = "") {
  const data = await api("/api/review/pair", {
    method: "POST",
    body: JSON.stringify({
      familia_local: decodeURIComponent(local),
      familia_avclass: decodeURIComponent(avclass),
      action,
      familia_manual: manual,
      note,
    }),
  });
  toast(`Par revisado: ${data.familia_local} -> ${data.familia_avclass} = ${data.familia_final}`);
  await loadSummary();
  await loadPairs();
  await loadQueue(0);
}

async function reviewPairManual(local, avclass) {
  const value = prompt("Familia final para este par:");
  if (!value) return;
  const note = prompt("Nota de revisión:", "") || "";
  await reviewPair(local, avclass, "manual", value, note);
}

async function openHash(hash) {
  document.querySelectorAll(".tab").forEach((b) => b.classList.remove("active"));
  document.querySelector('[data-tab="sample"]').classList.add("active");
  document.querySelectorAll(".panel").forEach((p) => p.classList.remove("active"));
  $("sample").classList.add("active");
  $("hashInput").value = decodeURIComponent(hash);
  await loadSample();
}

function kv(label, value) {
  return `<div class="kv"><strong>${label}</strong><span>${value ?? ""}</span></div>`;
}

async function loadSample() {
  const hash = $("hashInput").value.trim().toLowerCase();
  if (!hash) return;
  const data = await api(`/api/sample/${encodeURIComponent(hash)}`);
  const s = data.sample || {};
  const h = data.hybrid || {};
  if (!s.hash_md5) {
    $("sampleDetail").innerHTML = `<section>No se encontró el hash.</section>`;
    return;
  }
  const tagRows = table([
    ["tag_category", "Cat"],
    ["tag_name", "Tag"],
    ["votes", "Votos"],
  ], data.tags || []);
  $("sampleDetail").innerHTML = `
    <div class="detailGrid">
      <section>
        <h2>${s.hash_md5}</h2>
        ${kv("Lote", s.lote_origen)}
        ${kv("Local", h.familia_local)}
        ${kv("AVClass", h.familia_avclass)}
        ${kv("Final actual", s.familia_final_actual)}
        ${kv("Fuente", s.fuente_final_actual)}
        ${kv("Decisión", s.decision_actual)}
        ${kv("Confianza", h.confianza_final)}
        ${kv("Detección", s.detection_percent)}
        <h3>Acciones</h3>
        <div class="actions">
          <button onclick="reviewHash('${s.hash_md5}','accept_local')">Aceptar local</button>
          <button onclick="reviewHash('${s.hash_md5}','accept_avclass')">Aceptar AVClass</button>
          <button onclick="reviewHash('${s.hash_md5}','accept_suggested')">Aceptar sugerida</button>
          <button onclick="reviewHashManual('${s.hash_md5}')">Manual</button>
          <button onclick="reviewHash('${s.hash_md5}','sin_inferir')">Sin inferir</button>
        </div>
      </section>
      <section>
        <h2>Evidencia</h2>
        <h3>Detecciones top</h3>
        <div class="pre">${s.detecciones_top || ""}</div>
        <h3>Clases AVClass</h3>
        <div class="pre">${h.clases_avclass || ""}</div>
        <h3>Behaviors AVClass</h3>
        <div class="pre">${h.behaviors_avclass || ""}</div>
      </section>
    </div>
    <section>
      <h2>Tags AVClass</h2>
      <div class="tableWrap">${tagRows}</div>
    </section>
  `;
}

async function exportFinal() {
  const data = await api("/api/export/final", { method: "POST", body: JSON.stringify({}) });
  toast(`Exportado ${fmt(data.rows)} filas: ${data.output}`);
}

function setupTabs() {
  document.querySelectorAll(".tab").forEach((button) => {
    button.addEventListener("click", () => {
      document.querySelectorAll(".tab").forEach((b) => b.classList.remove("active"));
      document.querySelectorAll(".panel").forEach((p) => p.classList.remove("active"));
      button.classList.add("active");
      $(button.dataset.tab).classList.add("active");
    });
  });
}

window.reviewHash = reviewHash;
window.reviewHashManual = reviewHashManual;
window.reviewPair = reviewPair;
window.reviewPairManual = reviewPairManual;
window.openHash = openHash;

setupTabs();
$("refreshBtn").addEventListener("click", loadSummary);
$("loadPairsBtn").addEventListener("click", loadPairs);
$("loadQueueBtn").addEventListener("click", () => loadQueue(0));
$("queueNextBtn").addEventListener("click", () => loadQueue(state.queueOffset + 1));
$("queueSingle").addEventListener("change", () => loadQueue(0));
$("loadSampleBtn").addEventListener("click", loadSample);
$("exportBtn").addEventListener("click", exportFinal);

loadSummary().then(loadPairs).then(() => loadQueue(0)).catch((err) => toast(err.message));
