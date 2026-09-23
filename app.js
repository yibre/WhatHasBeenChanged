// app.js — 화면과 파이썬 워커를 잇는다. 파일 바이트는 워커 밖으로 나가지 않는다.

const SAME = 0, MOD = 1, ADD = 2, DEL = 3, NONE = 4;
const RENDER_CAP = 1200;          // 한 번에 그릴 최대 행 수
const CLS = { [ADD]: "cadd", [DEL]: "cdel", [MOD]: "cmod", [NONE]: "cnone" };

const $ = (s) => document.querySelector(s);
const files = { a: null, b: null };
let worker = null, ready = false, result = null, activeSheet = 0, shown = RENDER_CAP;
let pending = null, seq = 0;

/* ── 워커 감시(워치독) ──────────────────────────
   worker.js의 자체 타임아웃(withTimeout)은 동기 호출(pyodide.pyimport 등)이
   워커 스레드를 완전히 막아버리면 자신도 함께 멈춰서 발동하지 못한다.
   그러면 워커는 어떤 메시지도 못 보내고, 오버레이는 영원히 떠 있게 된다.
   그래서 메인 스레드 쪽에서 별도로 "이 시간 안에 다음 소식이 없으면 죽은 걸로 본다"를
   지키고, 응답이 없으면 워커를 강제로 terminate() 해서 화면을 풀어준다. */
const STAGE_TIMEOUT_MS = {
  runtime: 120000,
  excel_lib: 120000,
  diff_engine: 30000,
};
const WATCHDOG_DEFAULT_MS = 15000;   // 첫 status가 오기 전 기본 대기 시간
const WATCHDOG_GRACE_MS = 5000;      // 워커 자체 타임아웃이 정상 발동할 시간을 먼저 준다
let watchdogTimer = null;
let lastStageKey = null;             // 언어 전환 시 오버레이/상태 텍스트를 다시 그리는 데 쓴다

function armWatchdog(stageKey) {
  clearTimeout(watchdogTimer);
  const ms = (stageKey ? STAGE_TIMEOUT_MS[stageKey] : null) ?? WATCHDOG_DEFAULT_MS;
  watchdogTimer = setTimeout(() => onWorkerStuck(stageKey), ms + WATCHDOG_GRACE_MS);
}

function disarmWatchdog() {
  clearTimeout(watchdogTimer);
  watchdogTimer = null;
}

function onWorkerStuck(stageKey) {
  console.error("[엑셀 diff] 워커가 응답 없이 멈춰서 강제 종료합니다:", stageKey);
  worker && worker.terminate();
  worker = null;
  ready = false;
  overlay(false);
  setRuntime(t("stuck_title", stageKey), "error");
  fail("stuck_message", stageKey);
}

/* ── 워커 ──────────────────────────────────── */
function startWorker() {
  worker = new Worker("worker.js");
  worker.onmessage = ({ data }) => {
    if (data.type === "status") {
      lastStageKey = data.stage;
      armWatchdog(data.stage);
      const label = t(`stage_${data.stage}`);
      setRuntime(label, "loading");
      $("#overlayText").textContent = label;
    } else if (data.type === "ready") {
      disarmWatchdog();
      lastStageKey = null;
      ready = true;
      setRuntime(t("stage_ready"), "ready");
      overlay(false);
      refreshRunButton();
    } else if (data.type === "result") {
      disarmWatchdog();
      overlay(false);
      render(JSON.parse(data.json));
    } else if (data.type === "report") {
      disarmWatchdog();
      overlay(false);
      const bin = atob(data.b64);
      const buf = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
      save(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), t("filename_report"));
    } else if (data.type === "error") {
      disarmWatchdog();
      overlay(false);
      if (data.stage) setRuntime(t("error_title", data.stage), "error");
      fail("error_message", data.stage, data.message);
    }
  };
  worker.onerror = () => {
    disarmWatchdog();
    overlay(false);
    setRuntime(t("worker_load_fail_title"), "error");
    fail("worker_load_fail_msg");
  };
  overlay(true, t("overlay_starting"));
  worker.postMessage({ type: "init", lang });
  armWatchdog(null);
}

const setRuntime = (text, state) => {
  const el = $("#runtime");
  el.textContent = text;
  el.dataset.state = state;
};

function overlay(on, text) {
  $("#overlay").hidden = !on;
  if (text) $("#overlayText").textContent = text;
}

let failState = null;   // { key, args } — 언어 전환 시 다시 그리는 데 쓴다

function renderFail() {
  const h = $("#hint");
  if (failState) {
    h.textContent = t(failState.key, ...failState.args);
    h.classList.add("bad");
  } else {
    h.textContent = t("hint_default");
    h.classList.remove("bad");
  }
}

function fail(key, ...args) {
  failState = { key, args };
  renderFail();
}

function clearFail() {
  failState = null;
  renderFail();
}

/* ── 파일 선택 ─────────────────────────────── */
function setupWell(well) {
  const side = well.dataset.side;
  const input = well.querySelector("input");

  well.addEventListener("click", () => input.click());
  well.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); input.click(); }
  });
  input.addEventListener("change", () => input.files[0] && accept(side, input.files[0]));

  ["dragenter", "dragover"].forEach((t) =>
    well.addEventListener(t, (e) => { e.preventDefault(); well.classList.add("over"); }));
  ["dragleave", "drop"].forEach((t) =>
    well.addEventListener(t, (e) => { e.preventDefault(); well.classList.remove("over"); }));
  well.addEventListener("drop", (e) => {
    const f = e.dataTransfer.files[0];
    if (f) accept(side, f);
  });
}

function accept(side, file) {
  if (!/\.(xlsx|xlsm|xltx)$/i.test(file.name)) {
    fail("file_type_error", file.name);
    return;
  }
  clearFail();
  files[side] = file;
  paint(side);
  worker && worker.postMessage({ type: "reset" });
  refreshRunButton();
}

function paint(side) {
  const well = side === "a" ? $("#wellA") : $("#wellB");
  const f = files[side];
  well.classList.toggle("filled", !!f);
  well.querySelector(".well-name").textContent = f ? f.name : t("well_placeholder");
  well.querySelector(".well-meta").textContent = f
    ? `${(f.size / 1024).toFixed(0)} KB · ${new Date(f.lastModified).toLocaleDateString(t("date_locale"))}`
    : "";
}

const refreshRunButton = () => { $("#run").disabled = !(ready && files.a && files.b); };

$("#swap").addEventListener("click", () => {
  [files.a, files.b] = [files.b, files.a];
  paint("a"); paint("b");
  worker && worker.postMessage({ type: "reset" });
  refreshRunButton();
});

/* ── 비교 실행 ─────────────────────────────── */
$("#run").addEventListener("click", async () => {
  if (!files.a || !files.b) return;
  clearFail();
  overlay(true, t("overlay_reading"));
  const [a, b] = await Promise.all([files.a.arrayBuffer(), files.b.arrayBuffer()]);
  overlay(true, t("overlay_comparing"));
  const ua = new Uint8Array(a), ub = new Uint8Array(b);
  worker.postMessage({
    type: "compare", id: ++seq,
    a: ua, b: ub,
    nameA: files.a.name, nameB: files.b.name,
    options: {
      compare: $("#optCompare").value,
      view: $("#optView").value,
      ignore_case: $("#optCase").checked,
      ignore_space: $("#optSpace").checked,
    },
  }, [ua.buffer, ub.buffer]);
});

/* ── 결과 그리기 ───────────────────────────── */
const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

function render(res) {
  result = res;
  shown = RENDER_CAP;
  activeSheet = res.sheets.findIndex((s) => s.status !== "same");
  if (activeSheet < 0) activeSheet = 0;
  $("#result").hidden = false;
  drawSummary();
  drawTabs();
  drawGrid();
  drawList();
  if (res.row_limit_hit)
    fail("row_limit_hit", res.row_limit);
}

function drawSummary() {
  const s = result.summary;
  const items = [
    [t("sum_cell_mod"), s.cell_mod, "mod"],
    [t("sum_row_add"), s.row_add, "add"],
    [t("sum_row_del"), s.row_del, "del"],
    [t("sum_row_mod"), s.row_mod, "mod"],
    [t("sum_col_add"), s.col_add, "add"],
    [t("sum_col_del"), s.col_del, "del"],
    [t("sum_sheet_add"), s.sheet_add, "add"],
    [t("sum_sheet_del"), s.sheet_del, "del"],
  ];
  const total = items.reduce((n, [, v]) => n + v, 0);
  $("#summary").innerHTML = total === 0
    ? `<div class="stat"><b>0</b><span>${t("summary_same")}</span></div>`
    : items.map(([label, v, kind]) =>
      `<div class="stat ${v ? kind : "zero"}"><b>${v}</b><span>${label}</span></div>`).join("");
}

function drawTabs() {
  const dot = { added: "add", deleted: "del", changed: "mod" };
  $("#sheettabs").innerHTML = result.sheets.map((sh, i) => {
    const n = Object.values(sh.stat).reduce((a, b) => a + b, 0);
    const d = dot[sh.status] ? `<i class="dot ${dot[sh.status]}"></i>` : "";
    const label = sh.status === "added" ? t("tab_sheet_added") : sh.status === "deleted" ? t("tab_sheet_deleted") : n ? t("tab_count", n) : "";
    return `<button type="button" data-i="${i}" class="${i === activeSheet ? "on" : ""}">${d}${esc(sh.name)}${label ? `<span class="cnt">${label}</span>` : ""}</button>`;
  }).join("");
  $("#sheettabs").querySelectorAll("button").forEach((b) =>
    b.addEventListener("click", () => { activeSheet = +b.dataset.i; shown = RENDER_CAP; drawTabs(); drawGrid(); }));
}

function drawGrid() {
  const sh = result.sheets[activeSheet];
  const wrap = $("#gridWrap");
  if (!sh || !sh.rows.length) {
    wrap.innerHTML = `<p class="empty">${t("grid_sheet_empty", esc(sh ? sh.name : ""))}</p>`;
    return;
  }

  const head = sh.cols.map((c) => {
    const cls = c.s === ADD ? "add" : c.s === DEL ? "del" : "";
    const letter = c.s === DEL ? c.a : c.b;
    const moved = c.a && c.b && c.a !== c.b ? ` title="${t("grid_col_moved_title", c.a, c.b)}"` : "";
    return `<th class="${cls}"${moved}>${letter || "·"}${moved ? "*" : ""}</th>`;
  }).join("");

  const slice = sh.rows.slice(0, shown);
  const body = slice.map((r) => {
    if (r.gap !== undefined)
      return `<tr class="gaprow"><td colspan="${sh.cols.length + 2}">${t("grid_gap_row", r.gap)}</td></tr>`;
    const rc = r.s === ADD ? "add" : r.s === DEL ? "del" : r.s === MOD ? "mod" : "";
    const cells = r.c.map((c) => {
      const cls = CLS[c.s] || "";
      let inner;
      if (c.s === MOD) inner = `<span class="was">${esc(c.o)}</span><span class="now">${esc(c.n)}</span>`;
      else if (c.s === DEL) inner = esc(c.o);
      else if (c.s === NONE) inner = "";
      else inner = esc(c.n);
      const title = c.s === MOD ? ` title="${esc(c.o)} → ${esc(c.n)}"` : ` title="${esc(c.n ?? c.o ?? "")}"`;
      return `<td class="${cls}"${title}>${inner}</td>`;
    }).join("");
    return `<tr class="${rc}"><td class="gut a">${r.a ?? ""}</td><td class="gut b">${r.b ?? ""}</td>${cells}</tr>`;
  }).join("");

  wrap.innerHTML =
    `<table class="grid"><thead><tr>
      <th class="gut a" title="${t("grid_th_prev_title")}">${t("grid_th_prev")}</th>
      <th class="gut b" title="${t("grid_th_next_title")}">${t("grid_th_next")}</th>${head}
    </tr></thead><tbody>${body}</tbody></table>` +
    (sh.rows.length > shown
      ? `<button type="button" class="more">${t("grid_more", sh.rows.length - shown)}</button>` : "");

  const more = wrap.querySelector(".more");
  more && more.addEventListener("click", () => { shown += RENDER_CAP; drawGrid(); });
}

function drawList(q = "") {
  const rows = result.changes.filter((c) =>
    !q || c.o.toLowerCase().includes(q) || c.n.toLowerCase().includes(q) || c.sheet.toLowerCase().includes(q));
  $("#listWrap").innerHTML = rows.length === 0
    ? `<p class="empty">${result.changes.length ? t("list_no_match") : t("list_no_changes")}</p>`
    : `<table class="list"><thead><tr>
        <th>${t("list_th_sheet")}</th><th>${t("list_th_prev_pos")}</th><th>${t("list_th_next_pos")}</th><th>${t("list_th_prev_val")}</th><th>${t("list_th_next_val")}</th>
      </tr></thead><tbody>${rows.slice(0, 4000).map((c) =>
      `<tr><td>${esc(c.sheet)}</td><td class="addr">${c.a}</td><td class="addr">${c.b}</td>
        <td class="o">${esc(c.o)}</td><td class="n">${esc(c.n)}</td></tr>`).join("")}</tbody></table>`;
}

/* ── 보기 전환·저장 ─────────────────────────── */
document.querySelectorAll(".tab").forEach((tab) =>
  tab.addEventListener("click", () => {
    document.querySelectorAll(".tab").forEach((x) => x.classList.toggle("on", x === tab));
    const list = tab.dataset.view === "list";
    $("#gridWrap").hidden = list;
    $("#listWrap").hidden = !list;
    $("#sheettabs").hidden = list;
    $("#filter").hidden = !list;
    $(".legend").hidden = list;
  }));

$("#filter").addEventListener("input", (e) => drawList(e.target.value.trim().toLowerCase()));

function save(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = name; a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

$("#dlReport").addEventListener("click", () => {
  if (!result) return;
  overlay(true, t("overlay_report"));
  worker.postMessage({ type: "report", id: ++seq });
});

$("#dlJson").addEventListener("click", () => {
  if (!result) return;
  save(new Blob([JSON.stringify(result, null, 2)], { type: "application/json" }), t("filename_json"));
});

/* ── 언어 전환 ─────────────────────────────── */
function switchLang(next) {
  lang = next;
  localStorage.setItem("lang", lang);
  applyStaticI18n();
  paint("a"); paint("b");
  renderFail();
  if (ready) {
    setRuntime(t("stage_ready"), "ready");
  } else if (lastStageKey) {
    const label = t(`stage_${lastStageKey}`);
    setRuntime(label, "loading");
    if (!$("#overlay").hidden) $("#overlayText").textContent = label;
  }
  if (result) { drawSummary(); drawTabs(); drawGrid(); drawList($("#filter").value.trim().toLowerCase()); }
}

$("#langToggle").addEventListener("click", () => switchLang(lang === "ko" ? "en" : "ko"));

/* ── 시작 ─────────────────────────────────── */
applyStaticI18n();
[$("#wellA"), $("#wellB")].forEach(setupWell);
window.addEventListener("dragover", (e) => e.preventDefault());
window.addEventListener("drop", (e) => e.preventDefault());
startWorker();
