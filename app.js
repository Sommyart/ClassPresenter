import { init as createPptxPreviewer } from "https://esm.sh/pptx-preview@1.0.7";
import { getDocument, GlobalWorkerOptions } from "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.min.mjs";

GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.worker.min.mjs";

const $ = (selector) => document.querySelector(selector);
const api = async (path, options = {}) => { const response = await fetch(`/api${path}`, { credentials: "include", headers: { "Content-Type": "application/json", ...(options.headers || {}) }, ...options }); if (!response.ok) throw new Error((await response.json().catch(() => ({}))).error || "Request failed"); return response.json(); };
const localUser = JSON.parse(localStorage.getItem("class-presenter-user") || "null");
const auth = { user: localUser, signup: false };
const defaults = { mode: "countdown", duration: 300, position: "top-right", theme: "minimal", opacity: 92, warnings: [1800, 600, 300, 60, 10, 5], warningSound: false, presenterView: false, alwaysOnTop: false, displayTarget: "presentation", timerDisplay: "both", autoAdvance: false, slideDurations: [] };
const saved = (() => { try { return JSON.parse(localStorage.getItem("class-presenter-preferences") || "null"); } catch (error) { console.warn("Preferences could not be loaded", error); return null; } })();
const state = {
  presentation: null, config: { ...defaults, ...(saved || {}), warnings: Array.isArray(saved?.warnings) ? saved.warnings : defaults.warnings }, previewer: null, pdfDocument: null, slide: 1, session: null,
  timer: { startedAt: 0, pausedAt: 0, pausedTotal: 0, elapsedBeforePause: 0, running: false, paused: false, ended: false, warned: [] }, slideTimers: {}, hudTimeout: null, sort: "added", query: "",
};
const updateViewportProfile = () => {
  const width = Math.round(window.visualViewport?.width || window.innerWidth);
  const height = Math.round(window.visualViewport?.height || window.innerHeight);
  const breakpoint = width < 640 ? "mobile" : width < 1024 ? "tablet" : width < 1440 ? "desktop" : "large-desktop";
  const orientation = width >= height ? "landscape" : "portrait";
  const pointer = matchMedia("(pointer: coarse)").matches ? "coarse" : "fine";
  document.documentElement.dataset.breakpoint = breakpoint;
  document.documentElement.dataset.orientation = orientation;
  document.documentElement.dataset.pointer = pointer;
  document.documentElement.style.setProperty("--viewport-width", `${width}px`);
  document.documentElement.style.setProperty("--viewport-height", `${height}px`);
  document.documentElement.style.setProperty("--safe-top", "env(safe-area-inset-top, 0px)");
  document.documentElement.style.setProperty("--safe-right", "env(safe-area-inset-right, 0px)");
  document.documentElement.style.setProperty("--safe-bottom", "env(safe-area-inset-bottom, 0px)");
  document.documentElement.style.setProperty("--safe-left", "env(safe-area-inset-left, 0px)");
};
updateViewportProfile();
window.addEventListener("resize", updateViewportProfile, { passive: true });
window.visualViewport?.addEventListener("resize", updateViewportProfile, { passive: true });
screen.orientation?.addEventListener("change", updateViewportProfile);

const dbPromise = new Promise((resolve, reject) => {
  const request = indexedDB.open("class-presenter", 2);
  request.onupgradeneeded = () => {
    const database = request.result;
    if (!database.objectStoreNames.contains("presentations")) database.createObjectStore("presentations", { keyPath: "id" });
    if (!database.objectStoreNames.contains("presets")) database.createObjectStore("presets", { keyPath: "id" });
    if (!database.objectStoreNames.contains("sessions")) database.createObjectStore("sessions", { keyPath: "id" });
  };
  request.onsuccess = () => resolve(request.result);
  request.onerror = () => reject(request.error);
});
const db = async (storeName, operation, value) => {
  const database = await dbPromise;
  return new Promise((resolve, reject) => {
    const transaction = database.transaction(storeName, operation === "get" ? "readonly" : "readwrite");
    const store = transaction.objectStore(storeName);
    const request = operation === "get" ? store.getAll() : operation === "put" ? store.put(value) : store.delete(value);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
};
const formatTime = (seconds, hours = seconds >= 3600) => {
  const safe = Math.max(0, Math.floor(seconds)); const h = Math.floor(safe / 3600); const m = Math.floor((safe % 3600) / 60); const s = safe % 60;
  return hours || h ? [h, m, s].map((n) => String(n).padStart(2, "0")).join(":") : `${String(m).padStart(2, "0")}:${String(s).padStart(2, "0")}`;
};
const showToast = (message) => { const toast = $("#toast"); toast.textContent = message; toast.classList.add("show"); clearTimeout(showToast.timeout); showToast.timeout = setTimeout(() => toast.classList.remove("show"), 3000); };
const showView = (id) => document.querySelectorAll(".view").forEach((view) => view.classList.toggle("hidden", view.id !== id));
const fileDate = (timestamp) => new Intl.DateTimeFormat("en", { month: "short", day: "numeric", year: "numeric" }).format(timestamp);
const escapeHtml = (value) => String(value).replace(/[&<>"']/g, (character) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[character]);
const persistPreferences = () => localStorage.setItem("class-presenter-preferences", JSON.stringify(state.config));
const playWarningTone = () => {
  if (!state.config.warningSound) return;
  try {
    const context = new AudioContext(); const oscillator = context.createOscillator(); const gain = context.createGain();
    oscillator.frequency.value = 660; gain.gain.setValueAtTime(0.045, context.currentTime); gain.gain.exponentialRampToValueAtTime(0.001, context.currentTime + 0.18);
    oscillator.connect(gain).connect(context.destination); oscillator.start(); oscillator.stop(context.currentTime + 0.18);
  } catch (error) { console.warn("Warning sound unavailable", error); }
};
const parseDuration = () => { const parts = $("#custom-duration").value.trim().split(":").map(Number); return parts.length === 3 && parts.every((part) => Number.isFinite(part) && part >= 0) ? parts[0] * 3600 + parts[1] * 60 + parts[2] : null; };

let libraryRenderToken = 0;
const renderLibrary = async () => {
  const renderToken = ++libraryRenderToken;
  const all = await db("presentations", "get");
  if (renderToken !== libraryRenderToken) return;
  const items = all.filter((item) => item.name.toLowerCase().includes(state.query.toLowerCase())).sort((a, b) => state.sort === "alphabetical" ? a.name.localeCompare(b.name) : state.sort === "presented" ? (b.lastPresented || 0) - (a.lastPresented || 0) : b.createdAt - a.createdAt);
  $("#presentation-count").textContent = all.length; $("#empty-state").classList.toggle("hidden", all.length > 0); $("#library-grid").innerHTML = items.map((item, index) => { const safeName = escapeHtml(item.name); const title = escapeHtml(item.name.replace(/\.(pptx|pdf)$/i, "").slice(0, 25)); const type = item.type === "pdf" ? "PDF" : "PPTX"; const thumb = item.thumbnail ? `<img class="real-thumb" src="${item.thumbnail}" alt="First slide of ${safeName}" loading="lazy" />` : `<div class="thumb-fallback"><div class="thumb-text">${title}<small>CLASS PRESENTER</small></div></div>`; return `<article class="library-card"><div class="thumb">${thumb}<span class="thumb-count">${item.slideCount ? `${item.slideCount} slides` : type}</span></div><div class="card-info"><h3 title="${safeName}">${safeName}</h3><div class="card-meta"><span>${fileDate(item.createdAt)}</span><span>${item.lastPresented ? `Last ${fileDate(item.lastPresented)}` : "Not presented"}</span></div><div class="card-actions"><button class="open-card" data-open="${item.id}">Present</button><button class="settings-card" data-open="${item.id}">Settings</button><button class="delete-card" data-delete="${item.id}">Delete</button></div></div></article>`; }).join("");
  document.querySelectorAll("[data-open]").forEach((button) => button.addEventListener("click", () => openSetup(items.find((item) => item.id === button.dataset.open))));
  document.querySelectorAll("[data-delete]").forEach((button) => button.addEventListener("click", async () => { await db("presentations", "delete", button.dataset.delete); renderLibrary(); }));
};
const syncSetup = () => {
  const c = state.config; document.querySelectorAll(".mode-button").forEach((item) => item.classList.toggle("active", item.dataset.mode === c.mode));
  $("#countdown-options").classList.toggle("hidden", c.mode !== "countdown"); $("#countup-options").classList.toggle("hidden", c.mode !== "countup");
  document.querySelectorAll(".position-option").forEach((item) => item.classList.toggle("selected", item.dataset.position === c.position));
  $("#theme-select").value = c.theme; $("#timer-display-mode").value = c.timerDisplay || "both"; $("#opacity-range").value = c.opacity; $("#presenter-view-option").checked = c.presenterView; $("#warning-sound").checked = c.warningSound;
  document.querySelectorAll(".warning-chips input").forEach((item) => item.checked = c.warnings.includes(Number(item.value))); $("#auto-advance-option").checked = Boolean(c.autoAdvance); renderSlideDurations();
};
const openSetup = (presentation) => { state.presentation = presentation; $("#setup-filename").textContent = presentation.name; $("#setup-slide-count").textContent = presentation.slideCount ? `${presentation.slideCount} slides · ${presentation.type === "pdf" ? "PDF presentation" : "PowerPoint presentation"}` : "Presentation"; $("#custom-duration").value = formatTime(state.config.duration, true); syncSetup(); showView("setup-view"); };
const renderSlideDurations = () => {
  const total = state.presentation?.slideCount || 0;
  if (!total) { $("#slide-duration-list").innerHTML = '<p class="empty-presets">Slide thumbnails become available after the presentation is loaded.</p>'; return; }
  const each = state.config.duration / total;
  if (!Array.isArray(state.config.slideDurations) || state.config.slideDurations.length !== total) state.config.slideDurations = Array.from({ length: total }, () => 0);
  $("#slide-duration-list").classList.remove("hidden");
  $("#slide-duration-list").innerHTML = state.config.slideDurations.map((seconds, index) => `<label class="slide-timer-row"><span class="slide-thumb">S${index + 1}</span><span>Slide ${index + 1}</span><input type="checkbox" data-slide-enabled="${index}" ${seconds > 0 ? "checked" : ""} /><input data-slide-duration="${index}" value="${formatTime(seconds, true)}" inputmode="numeric" /></label>`).join("");
  document.querySelectorAll("[data-slide-duration]").forEach((input) => input.addEventListener("change", () => { const parts = input.value.split(":").map(Number); if (parts.length === 3 && parts.every((part) => Number.isFinite(part) && part >= 0)) { state.config.slideDurations[Number(input.dataset.slideDuration)] = parts[0] * 3600 + parts[1] * 60 + parts[2]; persistPreferences(); } }));
  document.querySelectorAll("[data-slide-enabled]").forEach((input) => input.addEventListener("change", () => { const index = Number(input.dataset.slideEnabled); if (!input.checked) state.config.slideDurations[index] = 0; else if (!state.config.slideDurations[index]) state.config.slideDurations[index] = Math.round(each); persistPreferences(); renderSlideDurations(); }));
  $("#setup-slide-summary").textContent = `${state.config.slideDurations.filter(Boolean).length} slide timers`;
  $("#setup-global-summary").textContent = `${state.config.timerDisplay || "both"} timer · ${formatTime(state.config.duration, true)}`;
};

const currentTimes = () => {
  if (!state.session) return { elapsed: 0, remaining: state.config.duration };
  const now = Date.now(); const elapsed = state.timer.paused ? state.timer.elapsedBeforePause : Math.max(0, (now - state.timer.startedAt - state.timer.pausedTotal) / 1000);
  return { elapsed, remaining: Math.max(0, state.config.duration - elapsed) };
};
const currentSlideTimes = () => {
  const duration = Number(state.config.slideDurations[state.slide - 1] || 0);
  if (!duration) return { elapsed: 0, remaining: 0, configured: false };
  const timer = state.slideTimers[state.slide - 1] || { startedAt: 0, pausedAt: 0, pausedTotal: 0, elapsedBeforePause: 0 };
  const elapsed = timer.pausedAt ? timer.elapsedBeforePause : timer.startedAt ? Math.max(0, (Date.now() - timer.startedAt - timer.pausedTotal) / 1000) : 0;
  return { elapsed, remaining: Math.max(0, duration - elapsed), configured: true };
};
const ensureSlideTimer = () => {
  const duration = Number(state.config.slideDurations[state.slide - 1] || 0);
  if (!duration || state.timer.paused || !state.timer.running || state.slideTimers[state.slide - 1]?.startedAt) return;
  state.slideTimers[state.slide - 1] = { startedAt: Date.now(), pausedAt: 0, pausedTotal: 0, elapsedBeforePause: 0, ended: false };
};
const updateTimer = () => {
  const times = currentTimes(); ensureSlideTimer(); const slideTimes = currentSlideTimes(); const globalSeconds = state.config.mode === "countdown" ? times.remaining : times.elapsed; const slideSeconds = state.config.mode === "countdown" ? slideTimes.remaining : slideTimes.elapsed;
  const globalText = formatTime(globalSeconds, state.config.mode === "countup" || globalSeconds >= 3600); const slideText = formatTime(slideSeconds, state.config.mode === "countup" || slideSeconds >= 3600);
  const display = state.config.timerDisplay || "both"; $("#timer-display").textContent = display === "global" ? globalText : display === "slide" ? (slideTimes.configured ? slideText : "—") : display === "none" ? "" : `${globalText}  ·  ${slideTimes.configured ? slideText : "—"}`; $("#timer-state").textContent = state.timer.ended ? "TIME ENDED" : state.timer.paused ? "PAUSED" : state.timer.running ? "LIVE" : "READY";
  $("#timer-overlay").classList.toggle("ended", state.timer.ended); $("#timer-overlay").style.opacity = state.config.opacity / 100; $("#timer-overlay").dataset.theme = state.config.theme; $("#pause-timer").textContent = state.timer.paused ? "▶" : "Ⅱ";
  $("#presenter-elapsed").textContent = formatTime(times.elapsed, times.elapsed >= 3600); $("#presenter-remaining").textContent = formatTime(times.remaining, times.remaining >= 3600);
  if (state.config.mode === "countdown" && state.timer.running && !state.timer.paused) {
    state.config.warnings.forEach((warning) => { if (times.remaining <= warning && !state.timer.warned.includes(warning)) { state.timer.warned.push(warning); $("#timer-overlay").classList.add("warning"); showToast(warning <= 10 ? `${warning} seconds remaining` : `${formatTime(warning)} remaining`); playWarningTone(); } });
    if (times.remaining <= 0) { state.timer.ended = true; state.timer.running = false; showToast("Allocated time has ended. You can continue presenting."); }
  }
  if (slideTimes.configured && state.timer.running && !state.timer.paused && slideTimes.remaining <= 0 && !state.slideTimers[state.slide - 1]?.ended) {
    state.slideTimers[state.slide - 1].ended = true;
    if (state.config.autoAdvance && state.slide < (state.presentation?.slideCount || 0)) renderSlide(state.slide);
  }
};
let timerTick;
const startTimerTicker = () => {
  if (timerTick || !state.timer.running) return;
  timerTick = window.setInterval(() => {
    if (!document.hidden || state.timer.running) updateTimer();
    if (!state.timer.running) { clearInterval(timerTick); timerTick = null; }
  }, 250);
};
document.addEventListener("visibilitychange", () => { if (state.timer.running) { updateTimer(); startTimerTicker(); } });
document.addEventListener("fullscreenchange", () => { updateViewportProfile(); if (state.presentation && !$("#present-view").classList.contains("hidden")) renderSlide(state.slide - 1); });

const renderSlide = async (index) => {
  const total = state.presentation?.type === "pdf" ? state.pdfDocument?.numPages : state.previewer?.slideCount;
  if (!total) return;
  const safe = Math.max(0, Math.min(index, total - 1));
  if (state.presentation.type === "pdf") {
    const page = await state.pdfDocument.getPage(safe + 1);
    const stage = $("#presentation-stage"); const baseViewport = page.getViewport({ scale: 1 }); const availableWidth = Math.max(1, stage.clientWidth); const availableHeight = Math.max(1, stage.clientHeight); const viewport = page.getViewport({ scale: Math.min(availableWidth / baseViewport.width, availableHeight / baseViewport.height) });
    const canvas = document.createElement("canvas"); canvas.width = viewport.width; canvas.height = viewport.height; canvas.className = "pdf-slide";
    $("#pptx-canvas").replaceChildren(canvas); await page.render({ canvasContext: canvas.getContext("2d"), viewport }).promise;
  } else state.previewer.renderSingleSlide(safe);
  state.slide = safe + 1; ensureSlideTimer(); $("#current-slide").textContent = state.slide; $("#total-slides").textContent = total; $("#presenter-current-number").textContent = state.slide; $("#presenter-current-preview").textContent = `Slide ${state.slide}`; $("#presenter-next-preview").textContent = state.slide < total ? `Slide ${state.slide + 1}` : "End of presentation"; updateTimer();
};
const startPresentation = async () => {
  const duration = state.config.mode === "countdown" ? parseDuration() : 0;
  if (state.config.mode === "countdown" && (!duration || duration <= 0)) return showToast("Enter a duration greater than zero.");
  state.config.duration = duration; if (!Array.isArray(state.config.slideDurations) || state.config.slideDurations.length !== (state.presentation.slideCount || 0)) state.config.slideDurations = Array.from({ length: state.presentation.slideCount || 1 }, () => 0); persistPreferences(); if (state.config.alwaysOnTop) showToast("Always on Top activates in the desktop wrapper."); state.slideTimers = {}; state.timer = { startedAt: Date.now(), pausedAt: 0, pausedTotal: 0, elapsedBeforePause: 0, running: true, paused: false, ended: false, warned: [] }; startTimerTicker();
  state.session = { id: crypto.randomUUID(), presentationId: state.presentation.id, startTime: Date.now(), mode: state.config.mode, duration, completed: false };
  await db("sessions", "put", state.session); state.presentation.lastPresented = Date.now(); await db("presentations", "put", state.presentation);
  $("#deck-title").textContent = state.presentation.name; $("#presenter-deck-name").textContent = state.presentation.name; $("#timer-overlay").className = `timer-overlay ${state.config.position}`; $("#pptx-canvas").innerHTML = '<div class="render-status"><span class="loading-spinner"></span><p>Preparing your slides…</p></div>'; $("#presenter-panel").classList.toggle("hidden", !state.config.presenterView); showView("present-view"); updateTimer();
  try {
    if (state.presentation.type === "pdf") { state.pdfDocument = await getDocument({ data: state.presentation.data }).promise; await renderSlide(0); }
    else { state.previewer = createPptxPreviewer($("#pptx-canvas"), { width: 1280, height: 720, mode: "slide" }); await state.previewer.preview(state.presentation.data); await renderSlide(0); }
  } catch (error) { console.error("Presentation rendering failed", error); $("#pptx-canvas").innerHTML = '<div class="render-status"><p>We could not render this presentation.</p><small>Check that the file or link is a valid PDF or PPTX.</small></div>'; showToast("This presentation could not be rendered."); }
};
const normalizeImportUrl = (rawUrl) => {
  const url = new URL(rawUrl.trim());
  if (url.hostname === "docs.google.com" && url.pathname.includes("/presentation/d/")) {
    const match = url.pathname.match(/\/presentation\/d\/([^/]+)/);
    if (match) return { url: `https://docs.google.com/presentation/d/${match[1]}/export/pptx`, name: "Google Slides presentation.pptx" };
  }
  if (url.hostname === "canva.com" || url.hostname.endsWith(".canva.com")) return { url: url.href, name: "Canva presentation" };
  throw new Error("Paste a Google Slides or Canva link.");
};
const importFromUrl = async () => {
  try {
    const source = normalizeImportUrl($("#presentation-url").value);
    $("#fetch-link-button").disabled = true; $("#fetch-link-button").textContent = "Fetching…";
    let response;
    try { response = await fetch(source.url); } catch (directError) { response = await fetch(`/api/import-url?url=${encodeURIComponent(source.url)}`); }
    if (!response.ok) {
      if (response.status === 501 || response.status === 404) throw new Error("This browser could not fetch the link directly. Run ClassPresenter with node server.cjs for server-assisted imports.");
      throw new Error(`The link returned HTTP ${response.status}.`);
    }
    let contentType = response.headers.get("content-type") || "";
    let data = await response.arrayBuffer();
    if (!contentType.includes("pdf") && !contentType.includes("presentation") && !contentType.includes("octet-stream")) {
      const assisted = await fetch(`/api/import-url?url=${encodeURIComponent(source.url)}`);
      if (!assisted.ok) throw new Error("This link did not return a downloadable PDF or PPTX. For Canva, use a public download/export link.");
      response = assisted; contentType = response.headers.get("content-type") || ""; data = await response.arrayBuffer();
    }
    const isPdf = contentType.includes("pdf") || source.url.toLowerCase().includes(".pdf");
    const isPptx = contentType.includes("presentation") || source.url.toLowerCase().includes(".pptx") || source.name.endsWith(".pptx");
    if (!isPdf && !isPptx) throw new Error("This link did not return a downloadable PDF or PPTX. For Canva, use a public download/export link.");
    await savePresentation({ name: isPdf ? source.name.replace(/\.pptx$/i, ".pdf") : source.name, data, type: isPdf ? "pdf" : "pptx" });
    $("#import-dialog").close(); $("#presentation-url").value = ""; showToast("Presentation imported.");
  } catch (error) { showToast(error.message || "Could not fetch that presentation link."); } finally { $("#fetch-link-button").disabled = false; $("#fetch-link-button").textContent = "Fetch presentation"; }
};
const savePresentation = async (item) => {
  const record = { id: crypto.randomUUID(), name: item.name, createdAt: Date.now(), data: item.data, type: item.type, slideCount: null, thumbnail: null };
  try {
    if (record.type === "pdf") record.slideCount = (await getDocument({ data: record.data.slice(0) }).promise).numPages;
    else { const host = document.createElement("div"); host.hidden = true; document.body.append(host); const parser = createPptxPreviewer(host, { width: 1280, height: 720, mode: "slide" }); await parser.preview(record.data.slice(0)); record.slideCount = parser.slideCount; parser.destroy(); host.remove(); }
  } catch (error) { console.warn("Slide count unavailable", error); }
  await db("presentations", "put", record); await renderLibrary();
};
const endSession = async () => { const times = currentTimes(); state.timer.running = false; if (state.session) { state.session.endTime = Date.now(); state.session.elapsed = times.elapsed; state.session.completed = state.timer.ended; await db("sessions", "put", state.session); } showView("setup-view"); showToast(`Session ended · ${formatTime(times.elapsed)} presented`); };


const generateFirstSlideThumbnail = async (record) => {
  if (record.type === "pdf") {
    const doc = await getDocument({ data: record.data.slice(0) }).promise;
    const page = await doc.getPage(1);
    const base = page.getViewport({ scale: 1 });
    const scale = Math.min(480 / base.width, 270 / base.height);
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(viewport.width));
    canvas.height = Math.max(1, Math.round(viewport.height));
    await page.render({ canvasContext: canvas.getContext("2d"), viewport }).promise;
    return canvas.toDataURL("image/jpeg", 0.78);
  }
  const host = document.createElement("div");
  host.style.cssText = "position:fixed;left:-10000px;top:-10000px;width:480px;height:270px;overflow:hidden;pointer-events:none";
  document.body.append(host);
  const parser = createPptxPreviewer(host, { width: 480, height: 270, mode: "slide" });
  await parser.preview(record.data.slice(0));
  await new Promise((resolve) => requestAnimationFrame(resolve));
  const canvas = host.querySelector("canvas");
  const result = canvas && canvas.width ? canvas.toDataURL("image/jpeg", 0.78) : null;
  parser.destroy();
  host.remove();
  return result;
};

const backfillThumbnails = async () => {
  const items = await db("presentations", "get");
  for (const item of items) {
    if (item.thumbnail) continue;
    try {
      item.thumbnail = await generateFirstSlideThumbnail(item);
      if (item.thumbnail) await db("presentations", "put", item);
    } catch (error) { console.warn("Thumbnail backfill skipped", error); }
  }
  renderLibrary();
};

const showAnalytics = async () => {
  const sessions = (await db("sessions", "get")).sort((a, b) => (b.startTime || 0) - (a.startTime || 0));
  const presentations = await db("presentations", "get");
  const names = new Map(presentations.map((p) => [p.id, p.name]));
  const totalSeconds = sessions.reduce((sum, s) => sum + Number(s.elapsed || 0), 0);
  const average = sessions.length ? totalSeconds / sessions.length : 0;
  const completed = sessions.filter((s) => s.completed).length;
  let dialog = $("#analytics-dialog");
  if (!dialog) {
    dialog = document.createElement("dialog");
    dialog.id = "analytics-dialog";
    dialog.className = "preset-dialog analytics-dialog";
    document.body.append(dialog);
  }
  const rows = sessions.slice(0, 12).map((s) => {
    const name = escapeHtml((names.get(s.presentationId) || "Presentation").replace(/\.(pptx|pdf)$/i, ""));
    return `<div class="analytics-row"><div><strong>${name}</strong><small>${fileDate(s.startTime)} · ${s.mode}</small></div><b>${formatTime(s.elapsed || 0, (s.elapsed || 0) >= 3600)}</b><span>${s.completed ? "Completed" : "Ended early"}</span></div>`;
  }).join("");
  dialog.innerHTML = `<form method="dialog"><div class="dialog-heading"><div><p class="eyebrow">PRESENTATION INSIGHTS</p><h2>Session analytics</h2></div><button class="dialog-close" value="cancel">×</button></div><div class="analytics-summary"><div><strong>${sessions.length}</strong><small>Sessions</small></div><div><strong>${formatTime(totalSeconds, totalSeconds >= 3600)}</strong><small>Total focus time</small></div><div><strong>${formatTime(average, average >= 3600)}</strong><small>Average session</small></div><div><strong>${completed}</strong><small>Completed</small></div></div><div class="analytics-list">${rows || '<div class="analytics-empty">No presentation sessions yet. Start a presentation to build your history.</div>'}</div></form>`;
  dialog.showModal();
};

const analyticsButton = document.createElement("button");
analyticsButton.id = "analytics-button";
analyticsButton.className = "text-button analytics-button";
analyticsButton.textContent = "Session analytics";
analyticsButton.addEventListener("click", showAnalytics);
$(".library-tools")?.append(analyticsButton);
backfillThumbnails().catch((error) => console.warn("Thumbnail backfill failed", error));

const upload = async (file) => {
  if (!file) return; if (!/\.(pptx|pdf)$/i.test(file.name)) return showToast("Please choose a .pptx or PDF presentation.");
  try { await savePresentation({ name: file.name, data: await file.arrayBuffer(), type: /\.pdf$/i.test(file.name) ? "pdf" : "pptx" }); showToast("Presentation added to your library."); } catch (error) { showToast(error.message); }
};

$("#upload-button").addEventListener("click", () => $("#file-input").click()); $("#empty-upload-button").addEventListener("click", () => $("#file-input").click()); $("#file-input").addEventListener("change", (event) => { upload(event.target.files[0]); event.target.value = ""; });
$("#import-link-button").addEventListener("click", () => $("#import-dialog").showModal()); $("#fetch-link-button").addEventListener("click", importFromUrl);
let searchTimeout;
$("#back-to-dashboard").addEventListener("click", () => showView("dashboard-view")); $("#library-search").addEventListener("input", (event) => { clearTimeout(searchTimeout); searchTimeout = setTimeout(() => { state.query = event.target.value; renderLibrary(); }, 120); }); $("#sort-select").addEventListener("change", (event) => { state.sort = event.target.value; renderLibrary(); });
document.querySelectorAll(".mode-button").forEach((button) => button.addEventListener("click", () => { state.config.mode = button.dataset.mode; syncSetup(); }));
document.querySelectorAll("[data-seconds]").forEach((button) => button.addEventListener("click", () => { $("#custom-duration").value = formatTime(Number(button.dataset.seconds), true); state.config.duration = Number(button.dataset.seconds); }));
document.querySelectorAll(".position-option").forEach((button) => button.addEventListener("click", () => { state.config.position = button.dataset.position; syncSetup(); }));
$("#theme-select").addEventListener("change", (event) => { state.config.theme = event.target.value; persistPreferences(); }); $("#timer-display-mode").addEventListener("change", (event) => { state.config.timerDisplay = event.target.value; persistPreferences(); updateTimer(); }); $("#opacity-range").addEventListener("input", (event) => { state.config.opacity = Number(event.target.value); persistPreferences(); }); $("#presenter-view-option").addEventListener("change", (event) => { state.config.presenterView = event.target.checked; persistPreferences(); }); $("#auto-advance-option").addEventListener("change", (event) => { state.config.autoAdvance = event.target.checked; persistPreferences(); }); $("#display-target").addEventListener("change", (event) => { state.config.displayTarget = event.target.value; persistPreferences(); });
document.querySelectorAll(".warning-chips input").forEach((input) => input.addEventListener("change", () => { state.config.warnings = [...document.querySelectorAll(".warning-chips input:checked")].map((item) => Number(item.value)); persistPreferences(); })); $("#warning-sound").addEventListener("change", (event) => { state.config.warningSound = event.target.checked; persistPreferences(); });
$("#start-presentation").addEventListener("click", startPresentation); $("#prev-slide").addEventListener("click", () => renderSlide(state.slide - 2)); $("#next-slide").addEventListener("click", () => renderSlide(state.slide));
$("#pause-timer").addEventListener("click", () => { if (state.timer.ended) return; const now = Date.now(); if (state.timer.paused) { state.timer.pausedTotal += now - state.timer.pausedAt; Object.values(state.slideTimers).forEach((timer) => { if (timer.pausedAt) { timer.pausedTotal += now - timer.pausedAt; timer.pausedAt = 0; } }); state.timer.paused = false; ensureSlideTimer(); } else { const active = state.slideTimers[state.slide - 1]; const activeElapsed = active ? currentSlideTimes().elapsed : 0; state.timer.pausedAt = now; state.timer.elapsedBeforePause = currentTimes().elapsed; if (active?.startedAt && !active.ended) { active.pausedAt = now; active.elapsedBeforePause = activeElapsed; } state.timer.paused = true; } updateTimer(); startTimerTicker(); });
$("#reset-timer").addEventListener("click", () => { state.timer.startedAt = Date.now(); state.timer.pausedTotal = 0; state.timer.elapsedBeforePause = 0; state.timer.paused = false; state.timer.ended = false; state.timer.running = true; state.timer.warned = []; state.slideTimers = {}; ensureSlideTimer(); updateTimer(); startTimerTicker(); });
$("#hide-timer").addEventListener("click", () => { $("#timer-overlay").classList.add("hidden"); $("#show-timer").classList.remove("hidden"); }); $("#show-timer").addEventListener("click", () => { $("#timer-overlay").classList.remove("hidden"); $("#show-timer").classList.add("hidden"); });
$("#exit-presentation").addEventListener("click", endSession); $("#fullscreen-button").addEventListener("click", async () => { try { if (!document.fullscreenElement) await $("#presentation-stage").requestFullscreen(); else await document.exitFullscreen(); } catch (error) { showToast("Fullscreen is not available in this browser."); } });
$("#presenter-view-button").addEventListener("click", () => $("#presenter-panel").classList.toggle("hidden")); $("#display-settings-button").addEventListener("click", () => showToast(window.getScreenDetails ? "Display selection is available in the desktop wrapper." : "Browser display selection is limited; use the desktop wrapper for projector routing."));
const revealHud = () => { $("#presentation-hud").classList.remove("faded"); clearTimeout(state.hudTimeout); state.hudTimeout = setTimeout(() => $("#presentation-hud").classList.add("faded"), 5000); };
document.addEventListener("keydown", (event) => { if ($("#present-view").classList.contains("hidden")) return; if (event.key === "ArrowRight" || event.key === "ArrowDown" || event.key === " ") renderSlide(state.slide); if (event.key === "ArrowLeft" || event.key === "ArrowUp") renderSlide(state.slide - 2); if (event.key.toLowerCase() === "p") $("#pause-timer").click(); if (event.key.toLowerCase() === "r") $("#reset-timer").click(); if (event.key.toLowerCase() === "t") ($("#timer-overlay").classList.contains("hidden") ? $("#show-timer") : $("#hide-timer")).click(); if (event.key.toLowerCase() === "f") $("#fullscreen-button").click(); if (event.key === "Escape" && !document.fullscreenElement) endSession(); if (event.key !== "Tab") revealHud(); });
$("#presentation-stage").addEventListener("mousemove", revealHud); $("#presentation-stage").addEventListener("pointerdown", revealHud);
$("#presentation-stage").addEventListener("wheel", (event) => { event.preventDefault(); renderSlide(state.slide + (event.deltaY > 0 ? 0 : -2)); }, { passive: false });
let touchStartX = 0; $("#presentation-stage").addEventListener("touchstart", (event) => { touchStartX = event.changedTouches[0].clientX; revealHud(); }, { passive: true }); $("#presentation-stage").addEventListener("touchend", (event) => { const delta = event.changedTouches[0].clientX - touchStartX; if (Math.abs(delta) > 35) renderSlide(state.slide + (delta < 0 ? 0 : -2)); }, { passive: true });

const renderPresets = async () => { const presets = await db("presets", "get"); $("#preset-list").innerHTML = presets.length ? presets.map((preset) => `<div class="preset-row"><div><strong>${preset.name}</strong><small>${preset.mode} · ${formatTime(preset.duration, true)} · ${preset.theme}</small></div><button data-use-preset="${preset.id}">Use</button><button data-delete-preset="${preset.id}">×</button></div>`).join("") : '<p class="empty-presets">No saved presets yet.</p>'; document.querySelectorAll("[data-use-preset]").forEach((button) => button.addEventListener("click", async () => { const preset = presets.find((item) => item.id === button.dataset.usePreset); state.config = { ...state.config, ...preset }; $("#preset-dialog").close(); if (state.presentation) openSetup(state.presentation); showToast(`${preset.name} applied.`); })); document.querySelectorAll("[data-delete-preset]").forEach((button) => button.addEventListener("click", async () => { await db("presets", "delete", button.dataset.deletePreset); renderPresets(); })); };
$("#preset-manager-button").addEventListener("click", () => { renderPresets(); $("#preset-dialog").showModal(); }); $("#save-preset").addEventListener("click", async () => { const name = $("#preset-name").value.trim(); if (!name) return showToast("Name your preset first."); await db("presets", "put", { id: crypto.randomUUID(), name, ...state.config }); $("#preset-name").value = ""; renderPresets(); showToast("Preset saved."); });
const presentationResizeObserver = new ResizeObserver(() => {
  updateViewportProfile();
  if (state.presentation && !$("#present-view").classList.contains("hidden") && state.presentation.type === "pdf") renderSlide(state.slide - 1);
});
presentationResizeObserver.observe($("#presentation-stage"));
renderLibrary();
