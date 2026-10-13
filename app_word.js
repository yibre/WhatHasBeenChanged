// app_word.js — 워드 탭의 화면 로직. app.js가 이미 띄워둔 worker를 그대로 같이 쓴다
// (파일 바이트는 여기서도 워커 밖으로 나가지 않는다).

// 블록 상태 코드는 app.js의 SAME/MOD/ADD/DEL과 동일한 값을 쓴다(파이썬 엔진과 공유).
const wordFiles = { a: null, b: null };
let wordResult = null, wordShown = RENDER_CAP, wordFailState = null;

/* ── 파일 선택 ─────────────────────────────── */
const WORD_EXT_RE = /\.(docx|docm|dotx)$/i;

function renderWordFail() {
  const h = $("#hintWord");
  if (wordFailState) {
    h.textContent = t(wordFailState.key, ...wordFailState.args);
    h.classList.add("bad");
  } else {
    h.textContent = t("hint_default");
    h.classList.remove("bad");
  }
}

function failWord(key, ...args) {
  wordFailState = { key, args };
  renderWordFail();
}

function clearFailWord() {
  wordFailState = null;
  renderWordFail();
}

async function acceptWord(side, file) {
  if (!WORD_EXT_RE.test(file.name)) {
    failWord("file_type_error_word", file.name);
    return;
  }
  const head = new Uint8Array(await file.slice(0, 8).arrayBuffer());
  if (OLE_SIGNATURE.every((b, i) => head[i] === b)) {
    failWord("encrypted_file_word", file.name);
    return;
  }
  clearFailWord();
  wordFiles[side] = file;
  paintWord(side);
  worker && worker.postMessage({ type: "reset", kind: "word" });
  refreshRunButtonWord();
}

function paintWord(side) {
  const well = side === "a" ? $("#wellWordA") : $("#wellWordB");
  const f = wordFiles[side];
  well.classList.toggle("filled", !!f);
  well.querySelector(".well-name").textContent = f ? f.name : t("well_placeholder");
  well.querySelector(".well-meta").textContent = f
    ? `${(f.size / 1024).toFixed(0)} KB · ${new Date(f.lastModified).toLocaleDateString(t("date_locale"))}`
    : "";
}

function refreshRunButtonWord() {
  const btn = $("#runWord");
  if (btn) btn.disabled = !(ready && wordFiles.a && wordFiles.b);
}

function setupWordWell(well) {
  const side = well.dataset.side;
  const input = well.querySelector("input");
  well.addEventListener("click", () => input.click());
  well.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); input.click(); }
  });
  input.addEventListener("change", () => input.files[0] && acceptWord(side, input.files[0]));
  ["dragenter", "dragover"].forEach((ev) =>
    well.addEventListener(ev, (e) => { e.preventDefault(); well.classList.add("over"); }));
  ["dragleave", "drop"].forEach((ev) =>
    well.addEventListener(ev, (e) => { e.preventDefault(); well.classList.remove("over"); }));
  well.addEventListener("drop", (e) => {
    const f = e.dataTransfer.files[0];
    if (f) acceptWord(side, f);
  });
}

$("#swapWord").addEventListener("click", () => {
  [wordFiles.a, wordFiles.b] = [wordFiles.b, wordFiles.a];
  paintWord("a"); paintWord("b");
  worker && worker.postMessage({ type: "reset", kind: "word" });
  refreshRunButtonWord();
});

/* ── 비교 실행 ─────────────────────────────── */
$("#runWord").addEventListener("click", async () => {
  if (!wordFiles.a || !wordFiles.b) return;
  clearFailWord();
  overlay(true, t("overlay_reading"));
  const [a, b] = await Promise.all([wordFiles.a.arrayBuffer(), wordFiles.b.arrayBuffer()]);
  overlay(true, t("overlay_comparing"));
  const ua = new Uint8Array(a), ub = new Uint8Array(b);
  worker.postMessage({
    type: "compare", kind: "word", id: ++seq,
    a: ua, b: ub,
    nameA: wordFiles.a.name, nameB: wordFiles.b.name,
    options: {
      ignore_case: $("#optCaseWord").checked,
      ignore_space: $("#optSpaceWord").checked,
    },
  }, [ua.buffer, ub.buffer]);
});

/* ── 결과 그리기 ───────────────────────────── */
function renderWord(res) {
  wordResult = res;
  wordShown = RENDER_CAP;
  $("#sampleWord").hidden = true;
  $("#resultWord").hidden = false;
  drawWordSummary();
  drawWordDoc();
  drawWordList();
  if (res.para_limit_hit) failWord("para_limit_hit", res.para_limit);
}

function drawWordSummary() {
  const s = wordResult.summary;
  const items = [
    [t("sum_para_add"), s.para_add, "add"],
    [t("sum_para_del"), s.para_del, "del"],
    [t("sum_para_mod"), s.para_mod, "mod"],
    [t("sum_word_add"), s.word_add, "add"],
    [t("sum_word_del"), s.word_del, "del"],
  ];
  const total = s.para_add + s.para_del + s.para_mod;
  $("#summaryWord").innerHTML = total === 0
    ? `<div class="stat"><b>0</b><span>${t("summary_same")}</span></div>`
    : items.map(([label, v, kind]) =>
      `<div class="stat ${v ? kind : "zero"}"><b>${v}</b><span>${label}</span></div>`).join("");
}

function spansHtml(spans) {
  return spans.map((sp) => {
    if (sp.t === "add") return `<ins>${esc(sp.s)}</ins>`;
    if (sp.t === "del") return `<del>${esc(sp.s)}</del>`;
    return esc(sp.s);
  }).join("");
}

function blockLabel(b) {
  return b.kind === "row" ? `<span class="rowtag">${t("word_table_row_label")}</span>` : "";
}

function drawWordDoc() {
  const wrap = $("#wordDocWrap");
  const blocks = wordResult.blocks;
  if (!blocks.length) {
    wrap.innerHTML = `<p class="empty">${t("word_doc_empty")}</p>`;
    return;
  }

  const slice = blocks.slice(0, wordShown);
  const html = slice.map((b) => {
    if (b.gap !== undefined) return `<p class="wgap">${t("word_gap", b.gap)}</p>`;
    const cls = b.s === ADD ? "add" : b.s === DEL ? "del" : b.s === MOD ? "mod" : "same";
    let inner;
    if (b.s === MOD) inner = spansHtml(b.d);
    else inner = esc(b.text);
    return `<p class="wp ${cls}">${blockLabel(b)}${inner || "&nbsp;"}</p>`;
  }).join("");

  wrap.innerHTML = html +
    (blocks.length > wordShown
      ? `<button type="button" class="more">${t("word_more", blocks.length - wordShown)}</button>` : "");

  const more = wrap.querySelector(".more");
  more && more.addEventListener("click", () => { wordShown += RENDER_CAP; drawWordDoc(); });
}

function drawWordList(q = "") {
  const rows = wordResult.changes.filter((c) =>
    !q || c.o.toLowerCase().includes(q) || c.n.toLowerCase().includes(q));
  $("#wordListWrap").innerHTML = rows.length === 0
    ? `<p class="empty">${wordResult.changes.length ? t("list_no_match") : t("list_no_changes")}</p>`
    : `<table class="list"><thead><tr>
        <th>${t("list_th_prev_pos")}</th><th>${t("list_th_next_pos")}</th><th>${t("list_th_prev_val")}</th><th>${t("list_th_next_val")}</th>
      </tr></thead><tbody>${rows.slice(0, 4000).map((c) =>
      `<tr><td class="addr">${c.a}</td><td class="addr">${c.b}</td>
        <td class="o">${esc(c.o)}</td><td class="n">${esc(c.n)}</td></tr>`).join("")}</tbody></table>`;
}

$("#filterWord").addEventListener("input", (e) => drawWordList(e.target.value.trim().toLowerCase()));

$("#dlJsonWord").addEventListener("click", () => {
  if (!wordResult) return;
  save(new Blob([JSON.stringify(wordResult, null, 2)], { type: "application/json" }), t("filename_json_word"));
});

/* ── 언어 전환 ─────────────────────────────── */
function onLangSwitchWord() {
  paintWord("a"); paintWord("b");
  renderWordFail();
  if (wordResult) { drawWordSummary(); drawWordDoc(); drawWordList($("#filterWord").value.trim().toLowerCase()); }
}

/* ── 시작 ─────────────────────────────────── */
[$("#wellWordA"), $("#wellWordB")].forEach(setupWordWell);
