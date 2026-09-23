// worker.js — 파이썬 런타임(Pyodide)을 백그라운드 스레드에서 구동한다.
// 파일 바이트는 이 워커 안에서만 다뤄지고, 어떤 네트워크 요청에도 실리지 않는다.

importScripts("i18n.js");   // STRINGS/t()/stageLabel() 공유 — 이 안에서 쓰는 lang은 init 메시지로 맞춘다

// 파이썬 런타임(Pyodide)과 micropip 휠을 CDN 대신 이 저장소의 pyodide/ 폴더에서 그대로 읽는다.
// 매 방문마다 CDN을 거치지 않으니 로딩이 훨씬 빠르고, CDN이 막힌 환경에서도 동작한다.
const CDN = "pyodide/";

let pyodide = null;
let engine = null;

const say = (stageKey) => {
  console.log("[엑셀 diff]", stageKey);
  postMessage({ type: "status", stage: stageKey });
};

// 어느 단계에서든 매달려 있지 않도록 제한 시간을 둔다
function withTimeout(promise, ms, stageKey) {
  let timer;
  const guard = new Promise((_, reject) => {
    timer = setTimeout(
      () => reject(new Error(t("worker_timeout", stageLabel(stageKey), ms / 1000))),
      ms
    );
  });
  return Promise.race([promise, guard]).finally(() => clearTimeout(timer));
}

async function step(stageKey, ms, fn) {
  say(stageKey);
  try {
    return await withTimeout(Promise.resolve().then(fn), ms, stageKey);
  } catch (err) {
    console.error(`[엑셀 diff] "${stageKey}" 실패`, err);
    err.stage = stageKey;
    throw err;
  }
}

/** openpyxl은 Pyodide 기본 패키지가 아니라서 wheel로 설치해야 한다.
 *  wheels/ 폴더에 넣어둔 게 있으면 그걸 쓰고, 없으면 PyPI에서 받는다. */
async function installOpenpyxl() {
  await pyodide.loadPackage(["micropip"]);

  let local = [];
  try {
    const res = await fetch("wheels/list.json", { cache: "no-store" });
    if (res.ok) {
      const names = await res.json();
      local = names.map((n) => new URL(`wheels/${n}`, location.href).href);
    }
  } catch (_) {
    /* 로컬 wheel이 없으면 PyPI로 간다 */
  }

  if (local.length) {
    console.log("[엑셀 diff] 로컬 wheel 사용:", local);
    self._wheels = local;
    await pyodide.runPythonAsync(`
import js, micropip
await micropip.install([str(u) for u in js.self._wheels])
`);
  } else {
    console.log("[엑셀 diff] PyPI에서 openpyxl 내려받는 중");
    await pyodide.runPythonAsync("import micropip\nawait micropip.install('openpyxl')");
  }

  pyodide.runPython("import openpyxl");
}

async function boot() {
  await step("runtime", 120000, async () => {
    importScripts(`${CDN}pyodide.js`);
    pyodide = await loadPyodide({ indexURL: CDN });
    console.log("[엑셀 diff] Pyodide", pyodide.version);
  });

  await step("excel_lib", 120000, installOpenpyxl);

  await step("diff_engine", 30000, async () => {
    const res = await fetch("diff_engine.py", { cache: "no-store" });
    if (!res.ok) throw new Error(t("worker_fetch_fail", res.status));
    const src = await res.text();
    if (!src.includes("def compare")) throw new Error(t("worker_bad_content"));

    pyodide.FS.mkdirTree("/lib/exceldiff");
    pyodide.FS.writeFile("/lib/exceldiff/diff_engine.py", src);
    pyodide.runPython("import sys; sys.path.insert(0, '/lib/exceldiff')");
    engine = pyodide.pyimport("diff_engine");
  });

  postMessage({ type: "ready", version: String(pyodide.version) });
}

onmessage = async (ev) => {
  const { type, id } = ev.data;
  try {
    if (type === "init") {
      if (ev.data.lang) lang = ev.data.lang;
      await boot();
    } else if (type === "reset") {
      engine && engine.reset();
    } else if (type === "compare") {
      const { a, b, nameA, nameB, options } = ev.data;
      const json = engine.compare(a, b, JSON.stringify(options), nameA, nameB);
      postMessage({ type: "result", id, json: String(json) });
    } else if (type === "report") {
      postMessage({ type: "report", id, b64: String(engine.report_b64()) });
    }
  } catch (err) {
    const detail = String((err && err.message) || err);
    console.error("[엑셀 diff] 오류", err);
    postMessage({
      type: "error",
      id,
      stage: (err && err.stage) || "",
      message: detail.split("\n").slice(-6).join("\n"),
    });
  }
};
