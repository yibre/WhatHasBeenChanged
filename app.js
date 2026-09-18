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
  "파이썬 런타임 내려받는 중": 120000,
  "엑셀 라이브러리 설치 중": 120000,
  "비교 엔진 올리는 중": 30000,
};
const WATCHDOG_DEFAULT_MS = 15000;   // 첫 status가 오기 전 기본 대기 시간
const WATCHDOG_GRACE_MS = 5000;      // 워커 자체 타임아웃이 정상 발동할 시간을 먼저 준다
let watchdogTimer = null;

function armWatchdog(stage) {
  clearTimeout(watchdogTimer);
  const ms = (stage ? STAGE_TIMEOUT_MS[stage] : null) ?? WATCHDOG_DEFAULT_MS;
  watchdogTimer = setTimeout(() => onWorkerStuck(stage), ms + WATCHDOG_GRACE_MS);
}

function disarmWatchdog() {
  clearTimeout(watchdogTimer);
  watchdogTimer = null;
}

function onWorkerStuck(stage) {
  console.error("[엑셀 diff] 워커가 응답 없이 멈춰서 강제 종료합니다:", stage);
  worker && worker.terminate();
  worker = null;
  ready = false;
  overlay(false);
  const label = stage || "시작";
  setRuntime(`${label} 단계에서 멈춤`, "error");
  fail(`"${label}" 단계에서 응답이 없어 강제로 멈췄습니다 (브라우저 개발자 도구 콘솔에 자세한 내용이 있습니다). 페이지를 새로고침해서 다시 시도해주세요.`);
}

/* ── 워커 ──────────────────────────────────── */
function startWorker() {
  worker = new Worker("worker.js");
  worker.onmessage = ({ data }) => {
    if (data.type === "status") {
      armWatchdog(data.stage);
      setRuntime(data.stage, "loading");
      $("#overlayText").textContent = data.stage;
    } else if (data.type === "ready") {
      disarmWatchdog();
      ready = true;
      setRuntime("파이썬 준비됨", "ready");
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
      save(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), "변경점리포트.xlsx");
    } else if (data.type === "error") {
      disarmWatchdog();
      overlay(false);
      if (data.stage) setRuntime(`${data.stage} 실패`, "error");
      fail(`${data.stage ? data.stage + " 단계에서 막혔습니다. " : ""}${data.message} (브라우저 개발자 도구 콘솔에 자세한 내용이 있습니다)`);
    }
  };
  worker.onerror = () => {
    disarmWatchdog();
    overlay(false);
    setRuntime("런타임을 불러오지 못함", "error");
    fail("파이썬 런타임을 불러오지 못했습니다. 인터넷 연결을 확인하거나, serve.py로 페이지를 연 것이 맞는지 확인하세요.");
  };
  overlay(true, "시작하는 중");
  armWatchdog(null);
  worker.postMessage({ type: "init" });
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

function fail(msg) {
  const h = $("#hint");
  h.textContent = msg;
  h.classList.add("bad");
}

function clearFail() {
  const h = $("#hint");
  h.textContent = "파일은 이 브라우저 안에서만 열립니다. 어디로도 전송되지 않습니다.";
  h.classList.remove("bad");
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
    fail(`${file.name}은(는) 읽을 수 없습니다. xlsx, xlsm, xltx 파일만 비교할 수 있습니다. 예전 xls 파일은 엑셀에서 xlsx로 저장한 뒤 다시 올려주세요.`);
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
  well.querySelector(".well-name").textContent = f ? f.name : "파일을 끌어다 놓거나 눌러서 선택";
  well.querySelector(".well-meta").textContent = f
    ? `${(f.size / 1024).toFixed(0)} KB · ${new Date(f.lastModified).toLocaleDateString("ko-KR")}`
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
  overlay(true, "파일 읽는 중");
  const [a, b] = await Promise.all([files.a.arrayBuffer(), files.b.arrayBuffer()]);
  overlay(true, "변경점 찾는 중");
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
    fail(`시트가 매우 커서 앞쪽 ${res.row_limit.toLocaleString()}행까지만 비교했습니다. 전체를 비교하려면 터미널에서 python diff_engine.py 이전.xlsx 이후.xlsx 를 실행하세요.`);
}

function drawSummary() {
  const s = result.summary;
  const items = [
    ["셀 변경", s.cell_mod, "mod"],
    ["행 추가", s.row_add, "add"],
    ["행 삭제", s.row_del, "del"],
    ["행 수정", s.row_mod, "mod"],
    ["열 추가", s.col_add, "add"],
    ["열 삭제", s.col_del, "del"],
    ["시트 추가", s.sheet_add, "add"],
    ["시트 삭제", s.sheet_del, "del"],
  ];
  const total = items.reduce((n, [, v]) => n + v, 0);
  $("#summary").innerHTML = total === 0
    ? `<div class="stat"><b>0</b><span>두 파일의 내용이 같습니다</span></div>`
    : items.map(([label, v, kind]) =>
      `<div class="stat ${v ? kind : "zero"}"><b>${v}</b><span>${label}</span></div>`).join("");
}

function drawTabs() {
  const dot = { added: "add", deleted: "del", changed: "mod" };
  $("#sheettabs").innerHTML = result.sheets.map((sh, i) => {
    const n = Object.values(sh.stat).reduce((a, b) => a + b, 0);
    const d = dot[sh.status] ? `<i class="dot ${dot[sh.status]}"></i>` : "";
    const label = sh.status === "added" ? "새 시트" : sh.status === "deleted" ? "삭제된 시트" : n ? `${n}건` : "";
    return `<button type="button" data-i="${i}" class="${i === activeSheet ? "on" : ""}">${d}${esc(sh.name)}${label ? `<span class="cnt">${label}</span>` : ""}</button>`;
  }).join("");
  $("#sheettabs").querySelectorAll("button").forEach((b) =>
    b.addEventListener("click", () => { activeSheet = +b.dataset.i; shown = RENDER_CAP; drawTabs(); drawGrid(); }));
}

function drawGrid() {
  const sh = result.sheets[activeSheet];
  const wrap = $("#gridWrap");
  if (!sh || !sh.rows.length) {
    wrap.innerHTML = `<p class="empty">${esc(sh ? sh.name : "")} 시트는 비어 있습니다.</p>`;
    return;
  }

  const head = sh.cols.map((c) => {
    const cls = c.s === ADD ? "add" : c.s === DEL ? "del" : "";
    const letter = c.s === DEL ? c.a : c.b;
    const moved = c.a && c.b && c.a !== c.b ? ` title="이전 ${c.a}열 → 이후 ${c.b}열"` : "";
    return `<th class="${cls}"${moved}>${letter || "·"}${moved ? "*" : ""}</th>`;
  }).join("");

  const slice = sh.rows.slice(0, shown);
  const body = slice.map((r) => {
    if (r.gap !== undefined)
      return `<tr class="gaprow"><td colspan="${sh.cols.length + 2}">동일한 ${r.gap}행 접힘</td></tr>`;
    const rc = r.s === ADD ? "add" : r.s === DEL ? "del" : r.s === MOD ? "mod" : "";
    const cells = r.c.map((c) => {
      const cls = CLS[c.s] || "";
      let inner;
      if (c.s === MOD) inner = `<span class="was">${esc(c.o)}</span><span class="now">${esc(c.n)}</span>`;
      else if (c.s === DEL) inner = esc(c.o);
      else if (c.s === NONE) inner = "";
      else inner = esc(c.n);
      const t = c.s === MOD ? ` title="${esc(c.o)} → ${esc(c.n)}"` : ` title="${esc(c.n ?? c.o ?? "")}"`;
      return `<td class="${cls}"${t}>${inner}</td>`;
    }).join("");
    return `<tr class="${rc}"><td class="gut a">${r.a ?? ""}</td><td class="gut b">${r.b ?? ""}</td>${cells}</tr>`;
  }).join("");

  wrap.innerHTML =
    `<table class="grid"><thead><tr>
      <th class="gut a" title="이전 파일의 행 번호">이전</th>
      <th class="gut b" title="이후 파일의 행 번호">이후</th>${head}
    </tr></thead><tbody>${body}</tbody></table>` +
    (sh.rows.length > shown
      ? `<button type="button" class="more">남은 ${sh.rows.length - shown}행 더 보기</button>` : "");

  const more = wrap.querySelector(".more");
  more && more.addEventListener("click", () => { shown += RENDER_CAP; drawGrid(); });
}

function drawList(q = "") {
  const rows = result.changes.filter((c) =>
    !q || c.o.toLowerCase().includes(q) || c.n.toLowerCase().includes(q) || c.sheet.toLowerCase().includes(q));
  $("#listWrap").innerHTML = rows.length === 0
    ? `<p class="empty">${result.changes.length ? "찾는 값이 없습니다." : "값이 바뀐 셀은 없습니다. 행이나 열 단위 변경은 표로 보기에서 확인하세요."}</p>`
    : `<table class="list"><thead><tr>
        <th>시트</th><th>이전 위치</th><th>이후 위치</th><th>이전 값</th><th>이후 값</th>
      </tr></thead><tbody>${rows.slice(0, 4000).map((c) =>
      `<tr><td>${esc(c.sheet)}</td><td class="addr">${c.a}</td><td class="addr">${c.b}</td>
        <td class="o">${esc(c.o)}</td><td class="n">${esc(c.n)}</td></tr>`).join("")}</tbody></table>`;
}

/* ── 보기 전환·저장 ─────────────────────────── */
document.querySelectorAll(".tab").forEach((t) =>
  t.addEventListener("click", () => {
    document.querySelectorAll(".tab").forEach((x) => x.classList.toggle("on", x === t));
    const list = t.dataset.view === "list";
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
  overlay(true, "리포트 만드는 중");
  worker.postMessage({ type: "report", id: ++seq });
});

$("#dlJson").addEventListener("click", () => {
  if (!result) return;
  save(new Blob([JSON.stringify(result, null, 2)], { type: "application/json" }), "변경점.json");
});

/* ── 시작 ─────────────────────────────────── */
[$("#wellA"), $("#wellB")].forEach(setupWell);
window.addEventListener("dragover", (e) => e.preventDefault());
window.addEventListener("drop", (e) => e.preventDefault());
startWorker();
