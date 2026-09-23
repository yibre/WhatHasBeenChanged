// i18n.js — 화면에 쓰는 모든 문구를 한/영으로 들고 있다가, 토글 시 다시 그려준다.

const STRINGS = {
  ko: {
    page_title: "엑셀 변경점 추적기",
    lang_toggle: "EN",
    tagline_loading: "파이썬 런타임 준비 중",

    well_a_role: "이전 버전",
    well_b_role: "이후 버전",
    well_placeholder: "파일을 끌어다 놓거나 눌러서 선택",
    aria_pick_files: "비교할 파일 선택",
    aria_pick_a: "이전 버전 파일 선택",
    aria_pick_b: "이후 버전 파일 선택",
    aria_swap: "두 파일 위치 바꾸기",
    btn_swap: "위치 바꾸기",

    opt_compare_label: "비교 기준",
    opt_compare_value: "계산된 값",
    opt_compare_formula: "수식",
    opt_view_label: "보기",
    opt_view_auto: "자동",
    opt_view_changes: "변경된 부분만",
    opt_view_all: "전체 행",
    chk_ignore_case: "대소문자 무시",
    chk_ignore_space: "공백 무시",
    btn_run: "변경점 비교",

    hint_default: "파일은 이 브라우저 안에서만 열립니다. 어디로도 전송되지 않습니다.",

    tab_grid: "표로 보기",
    tab_list: "변경 목록",
    filter_placeholder: "값으로 찾기",
    btn_report: "엑셀 리포트 저장",
    btn_json: "JSON 저장",
    legend_add: "추가",
    legend_del: "삭제",
    legend_mod: "수정",
    sheettabs_aria: "시트",
    result_view_aria: "결과 보기 방식",
    overlay_default: "준비 중",

    stage_runtime: "파이썬 런타임 내려받는 중",
    stage_excel_lib: "엑셀 라이브러리 설치 중",
    stage_diff_engine: "비교 엔진 올리는 중",
    stage_ready: "파이썬 준비됨",
    overlay_starting: "시작하는 중",
    overlay_reading: "파일 읽는 중",
    overlay_comparing: "변경점 찾는 중",
    overlay_report: "리포트 만드는 중",

    stuck_default_label: "시작",
    stuck_title: (k) => `${stageLabel(k)} 단계에서 멈춤`,
    stuck_message: (k) => `"${stageLabel(k)}" 단계에서 응답이 없어 강제로 멈췄습니다 (브라우저 개발자 도구 콘솔에 자세한 내용이 있습니다). 페이지를 새로고침해서 다시 시도해주세요.`,
    error_title: (k) => `${stageLabel(k)} 실패`,
    error_message: (k, m) => `${k ? stageLabel(k) + " 단계에서 막혔습니다. " : ""}${m} (브라우저 개발자 도구 콘솔에 자세한 내용이 있습니다)`,
    worker_load_fail_title: "런타임을 불러오지 못함",
    worker_load_fail_msg: "파이썬 런타임을 불러오지 못했습니다. 인터넷 연결을 확인하거나 페이지를 새로고침해 보세요.",
    worker_timeout: (l, sec) => `${l} 단계가 ${sec}초 안에 끝나지 않았습니다. 네트워크가 막혀 있을 수 있습니다.`,
    worker_fetch_fail: (status) => `diff_engine.py를 서버에서 받지 못했습니다 (HTTP ${status}). 같은 폴더에 있는지 확인하세요.`,
    worker_bad_content: "diff_engine.py 내용이 올바르지 않습니다.",

    file_type_error: (name) => `${name}은(는) 읽을 수 없습니다. xlsx, xlsm, xltx 파일만 비교할 수 있습니다. 예전 xls 파일은 엑셀에서 xlsx로 저장한 뒤 다시 올려주세요.`,
    row_limit_hit: (limit) => `시트가 매우 커서 앞쪽 ${limit.toLocaleString()}행까지만 비교했습니다. 전체를 비교하려면 터미널에서 python diff_engine.py 이전.xlsx 이후.xlsx 를 실행하세요.`,

    sum_cell_mod: "셀 변경", sum_row_add: "행 추가", sum_row_del: "행 삭제", sum_row_mod: "행 수정",
    sum_col_add: "열 추가", sum_col_del: "열 삭제", sum_sheet_add: "시트 추가", sum_sheet_del: "시트 삭제",
    summary_same: "두 파일의 내용이 같습니다",

    tab_sheet_added: "새 시트",
    tab_sheet_deleted: "삭제된 시트",
    tab_count: (n) => `${n}건`,

    grid_sheet_empty: (name) => `${name} 시트는 비어 있습니다.`,
    grid_col_moved_title: (a, b) => `이전 ${a}열 → 이후 ${b}열`,
    grid_th_prev: "이전",
    grid_th_next: "이후",
    grid_th_prev_title: "이전 파일의 행 번호",
    grid_th_next_title: "이후 파일의 행 번호",
    grid_gap_row: (n) => `동일한 ${n}행 접힘`,
    grid_more: (n) => `남은 ${n}행 더 보기`,

    list_no_match: "찾는 값이 없습니다.",
    list_no_changes: "값이 바뀐 셀은 없습니다. 행이나 열 단위 변경은 표로 보기에서 확인하세요.",
    list_th_sheet: "시트", list_th_prev_pos: "이전 위치", list_th_next_pos: "이후 위치",
    list_th_prev_val: "이전 값", list_th_next_val: "이후 값",

    filename_report: "변경점리포트.xlsx",
    filename_json: "변경점.json",
    date_locale: "ko-KR",
  },

  en: {
    page_title: "Excel Diff Tracker",
    lang_toggle: "KO",
    tagline_loading: "Preparing the Python runtime",

    well_a_role: "Before",
    well_b_role: "After",
    well_placeholder: "Drag a file here, or click to choose",
    aria_pick_files: "Choose files to compare",
    aria_pick_a: "Choose the before file",
    aria_pick_b: "Choose the after file",
    aria_swap: "Swap the two files",
    btn_swap: "Swap",

    opt_compare_label: "Compare by",
    opt_compare_value: "Calculated value",
    opt_compare_formula: "Formula",
    opt_view_label: "View",
    opt_view_auto: "Auto",
    opt_view_changes: "Changes only",
    opt_view_all: "All rows",
    chk_ignore_case: "Ignore case",
    chk_ignore_space: "Ignore whitespace",
    btn_run: "Compare",

    hint_default: "Files are opened only inside this browser. Nothing is ever sent anywhere.",

    tab_grid: "Grid view",
    tab_list: "Change list",
    filter_placeholder: "Search by value",
    btn_report: "Download Excel report",
    btn_json: "Download JSON",
    legend_add: "Added",
    legend_del: "Deleted",
    legend_mod: "Modified",
    sheettabs_aria: "Sheets",
    result_view_aria: "Result view",
    overlay_default: "Preparing…",

    stage_runtime: "Downloading the Python runtime",
    stage_excel_lib: "Installing the Excel library",
    stage_diff_engine: "Loading the diff engine",
    stage_ready: "Python ready",
    overlay_starting: "Starting…",
    overlay_reading: "Reading files…",
    overlay_comparing: "Finding changes…",
    overlay_report: "Building the report…",

    stuck_default_label: "Start",
    stuck_title: (k) => `Stuck at ${stageLabel(k)}`,
    stuck_message: (k) => `No response at the "${stageLabel(k)}" step, so it was force-stopped (see the browser console for details). Please reload the page and try again.`,
    error_title: (k) => `${stageLabel(k)} failed`,
    error_message: (k, m) => `${k ? "Got stuck at the " + stageLabel(k) + " step. " : ""}${m} (see the browser console for details)`,
    worker_load_fail_title: "Couldn't load the runtime",
    worker_load_fail_msg: "Couldn't load the Python runtime. Check your internet connection or try reloading the page.",
    worker_timeout: (l, sec) => `The "${l}" step didn't finish within ${sec}s. The network might be blocked.`,
    worker_fetch_fail: (status) => `Couldn't fetch diff_engine.py from the server (HTTP ${status}). Make sure it's in the same folder.`,
    worker_bad_content: "diff_engine.py content looks invalid.",

    file_type_error: (name) => `${name} can't be read. Only xlsx, xlsm, and xltx files are supported. If it's an old .xls file, save it as .xlsx in Excel first and try again.`,
    row_limit_hit: (limit) => `The sheet is very large, so only the first ${limit.toLocaleString()} rows were compared. To compare everything, run "python diff_engine.py before.xlsx after.xlsx" in a terminal.`,

    sum_cell_mod: "Cells changed", sum_row_add: "Rows added", sum_row_del: "Rows deleted", sum_row_mod: "Rows modified",
    sum_col_add: "Columns added", sum_col_del: "Columns deleted", sum_sheet_add: "Sheets added", sum_sheet_del: "Sheets deleted",
    summary_same: "The two files are identical",

    tab_sheet_added: "New sheet",
    tab_sheet_deleted: "Deleted sheet",
    tab_count: (n) => `${n}`,

    grid_sheet_empty: (name) => `The "${name}" sheet is empty.`,
    grid_col_moved_title: (a, b) => `Column ${a} → ${b}`,
    grid_th_prev: "Before",
    grid_th_next: "After",
    grid_th_prev_title: "Row number in the before file",
    grid_th_next_title: "Row number in the after file",
    grid_gap_row: (n) => `${n} unchanged rows collapsed`,
    grid_more: (n) => `Show ${n} more rows`,

    list_no_match: "No matching value found.",
    list_no_changes: "No cell values changed. Check the grid view for row/column-level changes.",
    list_th_sheet: "Sheet", list_th_prev_pos: "Before pos.", list_th_next_pos: "After pos.",
    list_th_prev_val: "Before value", list_th_next_val: "After value",

    filename_report: "diff-report.xlsx",
    filename_json: "changes.json",
    date_locale: "en-US",
  },
};

// Worker 안에서도 그대로 쓸 수 있게(localStorage 없음) 존재 여부부터 확인한다.
let lang = (typeof localStorage !== "undefined" && localStorage.getItem("lang"))
  || (typeof navigator !== "undefined" && navigator.language && navigator.language.toLowerCase().startsWith("ko") ? "ko" : "en");

function t(key, ...args) {
  const entry = (STRINGS[lang] && STRINGS[lang][key]) ?? STRINGS.ko[key];
  return typeof entry === "function" ? entry(...args) : entry;
}

// 부팅 단계 키("runtime" 등)를 현재 언어의 표시 문구로 바꾼다. 키가 없으면(예: 초기화 전에 멈춤) 기본 라벨.
function stageLabel(stageKey) {
  return stageKey ? t(`stage_${stageKey}`) : t("stuck_default_label");
}

function applyStaticI18n() {
  document.documentElement.lang = lang;
  document.title = t("page_title");
  document.querySelectorAll("[data-i18n]").forEach((el) => { el.textContent = t(el.dataset.i18n); });
  document.querySelectorAll("[data-i18n-placeholder]").forEach((el) => { el.placeholder = t(el.dataset.i18nPlaceholder); });
  document.querySelectorAll("[data-i18n-title]").forEach((el) => { el.title = t(el.dataset.i18nTitle); });
  document.querySelectorAll("[data-i18n-aria]").forEach((el) => { el.setAttribute("aria-label", t(el.dataset.i18nAria)); });
  const toggle = document.getElementById("langToggle");
  if (toggle) toggle.textContent = t("lang_toggle");
}
