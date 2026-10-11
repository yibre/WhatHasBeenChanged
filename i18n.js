// i18n.js — 화면에 쓰는 모든 문구를 한/영으로 들고 있다가, 토글 시 다시 그려준다.

const STRINGS = {
  ko: {
    page_title: "엑셀 변경점 추적기",
    apptab_excel: "Excel Diff Tracker",
    apptab_word: "Word Diff Tracker",
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
    worker_fetch_fail: (status, file) => `${file}를 서버에서 받지 못했습니다 (HTTP ${status}). 같은 폴더에 있는지 확인하세요.`,
    worker_bad_content: (file) => `${file} 내용이 올바르지 않습니다.`,

    file_type_error: (name) => `${name}은(는) 읽을 수 없습니다. xlsx, xlsm, xltx 파일만 비교할 수 있습니다. 예전 xls 파일은 엑셀에서 xlsx로 저장한 뒤 다시 올려주세요.`,
    encrypted_file: (name) => `${name} 파일은 암호가 걸려 있거나 회사 DRM(정보 관리·IRM)으로 보호되어 있어서 브라우저에서 열 수 없습니다. 엑셀에서 파일 → 정보 → 통합 문서 보호 메뉴로 암호나 액세스 제한을 해제한 뒤 다시 저장해서 올려주세요. 조직에서 배포한 문서라면 관리자에게 보호 해제를 요청해야 할 수도 있습니다.`,
    row_limit_hit: (limit) => `시트가 매우 커서 앞쪽 ${limit.toLocaleString()}행까지만 비교했습니다. 전체를 비교하려면 터미널에서 python diff_engine.py 이전.xlsx 이후.xlsx 를 실행하세요.`,

    file_type_error_word: (name) => `${name}은(는) 읽을 수 없습니다. docx, docm, dotx 파일만 비교할 수 있습니다. 예전 doc 파일은 워드에서 docx로 저장한 뒤 다시 올려주세요.`,
    encrypted_file_word: (name) => `${name} 파일은 암호가 걸려 있거나 회사 DRM(정보 관리·IRM)으로 보호되어 있어서 브라우저에서 열 수 없습니다. 워드에서 파일 → 정보 → 문서 보호 메뉴로 암호나 액세스 제한을 해제한 뒤 다시 저장해서 올려주세요. 조직에서 배포한 문서라면 관리자에게 보호 해제를 요청해야 할 수도 있습니다.`,
    para_limit_hit: (limit) => `문서가 매우 커서 앞쪽 ${limit.toLocaleString()}개 문단까지만 비교했습니다.`,

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

    tab_doc: "문서 보기",
    sum_para_add: "문단 추가", sum_para_del: "문단 삭제", sum_para_mod: "문단 수정",
    sum_word_add: "단어 추가", sum_word_del: "단어 삭제",
    word_table_row_label: "표",
    word_doc_empty: "문서에 내용이 없습니다.",
    word_gap: (n) => `동일한 ${n}개 단락 접힘`,
    word_more: (n) => `남은 ${n}개 더 보기`,
    filename_json_word: "문서변경점.json",
  },

  en: {
    page_title: "Excel Diff Tracker",
    apptab_excel: "Excel Diff Tracker",
    apptab_word: "Word Diff Tracker",
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
    worker_fetch_fail: (status, file) => `Couldn't fetch ${file} from the server (HTTP ${status}). Make sure it's in the same folder.`,
    worker_bad_content: (file) => `${file} content looks invalid.`,

    file_type_error: (name) => `${name} can't be read. Only xlsx, xlsm, and xltx files are supported. If it's an old .xls file, save it as .xlsx in Excel first and try again.`,
    encrypted_file: (name) => `${name} looks password-protected or restricted by corporate DRM (IRM), so it can't be opened in the browser. In Excel, go to File → Info → Protect Workbook, remove the password or access restriction, then save and upload again. If it's a DRM-labeled document from your organization, you may need to ask your admin to remove the protection.`,
    row_limit_hit: (limit) => `The sheet is very large, so only the first ${limit.toLocaleString()} rows were compared. To compare everything, run "python diff_engine.py before.xlsx after.xlsx" in a terminal.`,

    file_type_error_word: (name) => `${name} can't be read. Only docx, docm, and dotx files are supported. If it's an old .doc file, save it as .docx in Word first and try again.`,
    encrypted_file_word: (name) => `${name} looks password-protected or restricted by corporate DRM (IRM), so it can't be opened in the browser. In Word, go to File → Info → Protect Document, remove the password or access restriction, then save and upload again. If it's a DRM-labeled document from your organization, you may need to ask your admin to remove the protection.`,
    para_limit_hit: (limit) => `The document is very large, so only the first ${limit.toLocaleString()} paragraphs were compared.`,

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

    tab_doc: "Document view",
    sum_para_add: "Paragraphs added", sum_para_del: "Paragraphs deleted", sum_para_mod: "Paragraphs modified",
    sum_word_add: "Words added", sum_word_del: "Words deleted",
    word_table_row_label: "Table",
    word_doc_empty: "The document has no content.",
    word_gap: (n) => `${n} unchanged paragraphs collapsed`,
    word_more: (n) => `Show ${n} more`,
    filename_json_word: "word-diff.json",
  },

  zh: {
    page_title: "Excel 变更追踪器",
    apptab_excel: "Excel Diff Tracker",
    apptab_word: "Word Diff Tracker",
    tagline_loading: "正在准备 Python 运行环境",

    well_a_role: "之前",
    well_b_role: "之后",
    well_placeholder: "拖拽文件到此处，或点击选择",
    aria_pick_files: "选择要比较的文件",
    aria_pick_a: "选择之前版本的文件",
    aria_pick_b: "选择之后版本的文件",
    aria_swap: "交换两个文件",
    btn_swap: "交换",

    opt_compare_label: "比较依据",
    opt_compare_value: "计算值",
    opt_compare_formula: "公式",
    opt_view_label: "视图",
    opt_view_auto: "自动",
    opt_view_changes: "仅显示更改",
    opt_view_all: "全部行",
    chk_ignore_case: "忽略大小写",
    chk_ignore_space: "忽略空格",
    btn_run: "比较",

    hint_default: "文件只会在此浏览器中打开，不会上传到任何地方。",

    tab_grid: "表格视图",
    tab_list: "更改列表",
    filter_placeholder: "按值搜索",
    btn_report: "下载 Excel 报告",
    btn_json: "下载 JSON",
    legend_add: "新增",
    legend_del: "删除",
    legend_mod: "修改",
    sheettabs_aria: "工作表",
    result_view_aria: "结果视图",
    overlay_default: "准备中…",

    stage_runtime: "正在下载 Python 运行环境",
    stage_excel_lib: "正在安装 Excel 库",
    stage_diff_engine: "正在加载比较引擎",
    stage_ready: "Python 已就绪",
    overlay_starting: "正在启动…",
    overlay_reading: "正在读取文件…",
    overlay_comparing: "正在查找更改…",
    overlay_report: "正在生成报告…",

    stuck_default_label: "开始",
    stuck_title: (k) => `卡在「${stageLabel(k)}」步骤`,
    stuck_message: (k) => `"${stageLabel(k)}" 步骤长时间无响应，已被强制停止（详情请查看浏览器开发者工具控制台）。请刷新页面重试。`,
    error_title: (k) => `${stageLabel(k)} 失败`,
    error_message: (k, m) => `${k ? "在「" + stageLabel(k) + "」步骤出错。" : ""}${m}（详情请查看浏览器开发者工具控制台）`,
    worker_load_fail_title: "无法加载运行环境",
    worker_load_fail_msg: "无法加载 Python 运行环境。请检查网络连接或尝试刷新页面。",
    worker_timeout: (l, sec) => `"${l}" 步骤在 ${sec} 秒内未完成。网络可能被阻止。`,
    worker_fetch_fail: (status, file) => `无法从服务器获取 ${file}（HTTP ${status}）。请确认它与本页面在同一文件夹中。`,
    worker_bad_content: (file) => `${file} 的内容似乎不正确。`,

    file_type_error: (name) => `无法读取 ${name}。仅支持 xlsx、xlsm、xltx 文件。如果是旧版 .xls 文件，请先在 Excel 中另存为 .xlsx 后重新上传。`,
    encrypted_file: (name) => `${name} 似乎已设置密码保护，或受公司 DRM（信息权限管理）保护，无法在浏览器中打开。请在 Excel 中依次打开「文件 → 信息 → 保护工作簿」，移除密码或访问限制后重新保存并上传。如果这是组织分发的受 DRM 保护的文档，可能需要请管理员解除保护。`,
    row_limit_hit: (limit) => `工作表非常大，因此仅比较了前 ${limit.toLocaleString()} 行。如需比较全部内容，请在终端运行 "python diff_engine.py before.xlsx after.xlsx"。`,

    file_type_error_word: (name) => `无法读取 ${name}。仅支持 docx、docm、dotx 文件。如果是旧版 .doc 文件，请先在 Word 中另存为 .docx 后重新上传。`,
    encrypted_file_word: (name) => `${name} 似乎已设置密码保护，或受公司 DRM（信息权限管理）保护，无法在浏览器中打开。请在 Word 中依次打开「文件 → 信息 → 保护文档」，移除密码或访问限制后重新保存并上传。如果这是组织分发的受 DRM 保护的文档，可能需要请管理员解除保护。`,
    para_limit_hit: (limit) => `文档非常大，因此仅比较了前 ${limit.toLocaleString()} 个段落。`,

    sum_cell_mod: "单元格更改", sum_row_add: "新增行", sum_row_del: "删除行", sum_row_mod: "修改行",
    sum_col_add: "新增列", sum_col_del: "删除列", sum_sheet_add: "新增工作表", sum_sheet_del: "删除工作表",
    summary_same: "两个文件内容相同",

    tab_sheet_added: "新工作表",
    tab_sheet_deleted: "已删除的工作表",
    tab_count: (n) => `${n} 项`,

    grid_sheet_empty: (name) => `工作表「${name}」为空。`,
    grid_col_moved_title: (a, b) => `之前第 ${a} 列 → 之后第 ${b} 列`,
    grid_th_prev: "之前",
    grid_th_next: "之后",
    grid_th_prev_title: "之前文件中的行号",
    grid_th_next_title: "之后文件中的行号",
    grid_gap_row: (n) => `已折叠 ${n} 行相同内容`,
    grid_more: (n) => `显示剩余 ${n} 行`,

    list_no_match: "未找到匹配的值。",
    list_no_changes: "没有单元格的值发生变化。行或列级别的更改请查看表格视图。",
    list_th_sheet: "工作表", list_th_prev_pos: "之前位置", list_th_next_pos: "之后位置",
    list_th_prev_val: "之前的值", list_th_next_val: "之后的值",

    filename_report: "差异报告.xlsx",
    filename_json: "更改.json",
    date_locale: "zh-CN",

    tab_doc: "文档视图",
    sum_para_add: "新增段落", sum_para_del: "删除段落", sum_para_mod: "修改段落",
    sum_word_add: "新增字词", sum_word_del: "删除字词",
    word_table_row_label: "表格",
    word_doc_empty: "文档没有内容。",
    word_gap: (n) => `已折叠 ${n} 个相同段落`,
    word_more: (n) => `显示剩余 ${n} 项`,
    filename_json_word: "文档差异.json",
  },

  es: {
    page_title: "Rastreador de cambios en Excel",
    apptab_excel: "Excel Diff Tracker",
    apptab_word: "Word Diff Tracker",
    tagline_loading: "Preparando el entorno de Python",

    well_a_role: "Antes",
    well_b_role: "Después",
    well_placeholder: "Arrastra un archivo aquí o haz clic para elegirlo",
    aria_pick_files: "Elige los archivos a comparar",
    aria_pick_a: "Elige el archivo anterior",
    aria_pick_b: "Elige el archivo posterior",
    aria_swap: "Intercambiar los dos archivos",
    btn_swap: "Intercambiar",

    opt_compare_label: "Comparar por",
    opt_compare_value: "Valor calculado",
    opt_compare_formula: "Fórmula",
    opt_view_label: "Vista",
    opt_view_auto: "Automática",
    opt_view_changes: "Solo cambios",
    opt_view_all: "Todas las filas",
    chk_ignore_case: "Ignorar mayúsculas/minúsculas",
    chk_ignore_space: "Ignorar espacios",
    btn_run: "Comparar",

    hint_default: "Los archivos solo se abren dentro de este navegador. Nunca se envían a ningún lugar.",

    tab_grid: "Vista de tabla",
    tab_list: "Lista de cambios",
    filter_placeholder: "Buscar por valor",
    btn_report: "Descargar informe Excel",
    btn_json: "Descargar JSON",
    legend_add: "Añadido",
    legend_del: "Eliminado",
    legend_mod: "Modificado",
    sheettabs_aria: "Hojas",
    result_view_aria: "Vista de resultados",
    overlay_default: "Preparando…",

    stage_runtime: "Descargando el entorno de Python",
    stage_excel_lib: "Instalando la librería de Excel",
    stage_diff_engine: "Cargando el motor de comparación",
    stage_ready: "Python listo",
    overlay_starting: "Iniciando…",
    overlay_reading: "Leyendo archivos…",
    overlay_comparing: "Buscando cambios…",
    overlay_report: "Generando el informe…",

    stuck_default_label: "Inicio",
    stuck_title: (k) => `Bloqueado en "${stageLabel(k)}"`,
    stuck_message: (k) => `No hubo respuesta en el paso "${stageLabel(k)}", así que se detuvo forzosamente (hay más detalles en la consola de las herramientas de desarrollador del navegador). Recarga la página e inténtalo de nuevo.`,
    error_title: (k) => `Fallo en ${stageLabel(k)}`,
    error_message: (k, m) => `${k ? "Se bloqueó en el paso " + stageLabel(k) + ". " : ""}${m} (hay más detalles en la consola de las herramientas de desarrollador del navegador)`,
    worker_load_fail_title: "No se pudo cargar el entorno",
    worker_load_fail_msg: "No se pudo cargar el entorno de Python. Comprueba tu conexión a internet o recarga la página.",
    worker_timeout: (l, sec) => `El paso "${l}" no terminó en ${sec} s. Puede que la red esté bloqueada.`,
    worker_fetch_fail: (status, file) => `No se pudo obtener ${file} del servidor (HTTP ${status}). Comprueba que esté en la misma carpeta.`,
    worker_bad_content: (file) => `El contenido de ${file} no parece válido.`,

    file_type_error: (name) => `${name} no se puede leer. Solo se admiten archivos xlsx, xlsm y xltx. Si es un archivo .xls antiguo, guárdalo como .xlsx en Excel y vuelve a subirlo.`,
    encrypted_file: (name) => `El archivo ${name} parece estar protegido con contraseña o restringido por DRM corporativo (IRM), por lo que no se puede abrir en el navegador. En Excel, ve a Archivo → Información → Proteger libro, quita la contraseña o la restricción de acceso, y vuelve a guardarlo y subirlo. Si es un documento con DRM distribuido por tu organización, es posible que debas pedirle a tu administrador que quite la protección.`,
    row_limit_hit: (limit) => `La hoja es muy grande, así que solo se compararon las primeras ${limit.toLocaleString()} filas. Para comparar todo, ejecuta "python diff_engine.py antes.xlsx despues.xlsx" en una terminal.`,

    file_type_error_word: (name) => `${name} no se puede leer. Solo se admiten archivos docx, docm y dotx. Si es un archivo .doc antiguo, guárdalo como .docx en Word y vuelve a subirlo.`,
    encrypted_file_word: (name) => `El archivo ${name} parece estar protegido con contraseña o restringido por DRM corporativo (IRM), por lo que no se puede abrir en el navegador. En Word, ve a Archivo → Información → Proteger documento, quita la contraseña o la restricción de acceso, y vuelve a guardarlo y subirlo. Si es un documento con DRM distribuido por tu organización, es posible que debas pedirle a tu administrador que quite la protección.`,
    para_limit_hit: (limit) => `El documento es muy grande, así que solo se compararon los primeros ${limit.toLocaleString()} párrafos.`,

    sum_cell_mod: "Celdas modificadas", sum_row_add: "Filas añadidas", sum_row_del: "Filas eliminadas", sum_row_mod: "Filas modificadas",
    sum_col_add: "Columnas añadidas", sum_col_del: "Columnas eliminadas", sum_sheet_add: "Hojas añadidas", sum_sheet_del: "Hojas eliminadas",
    summary_same: "Los dos archivos son idénticos",

    tab_sheet_added: "Hoja nueva",
    tab_sheet_deleted: "Hoja eliminada",
    tab_count: (n) => `${n}`,

    grid_sheet_empty: (name) => `La hoja "${name}" está vacía.`,
    grid_col_moved_title: (a, b) => `Columna ${a} (antes) → ${b} (después)`,
    grid_th_prev: "Antes",
    grid_th_next: "Después",
    grid_th_prev_title: "Número de fila en el archivo anterior",
    grid_th_next_title: "Número de fila en el archivo posterior",
    grid_gap_row: (n) => `${n} filas sin cambios colapsadas`,
    grid_more: (n) => `Mostrar ${n} filas más`,

    list_no_match: "No se encontró ningún valor coincidente.",
    list_no_changes: "No hay celdas con valores modificados. Revisa la vista de tabla para ver cambios de filas o columnas.",
    list_th_sheet: "Hoja", list_th_prev_pos: "Pos. anterior", list_th_next_pos: "Pos. posterior",
    list_th_prev_val: "Valor anterior", list_th_next_val: "Valor posterior",

    filename_report: "informe-de-cambios.xlsx",
    filename_json: "cambios.json",
    date_locale: "es-ES",

    tab_doc: "Vista de documento",
    sum_para_add: "Párrafos añadidos", sum_para_del: "Párrafos eliminados", sum_para_mod: "Párrafos modificados",
    sum_word_add: "Palabras añadidas", sum_word_del: "Palabras eliminadas",
    word_table_row_label: "Tabla",
    word_doc_empty: "El documento no tiene contenido.",
    word_gap: (n) => `${n} párrafos sin cambios colapsados`,
    word_more: (n) => `Mostrar ${n} más`,
    filename_json_word: "diferencias-word.json",
  },
};

// Worker 안에서도 그대로 쓸 수 있게(localStorage 없음) 존재 여부부터 확인한다.
function detectLang() {
  if (typeof localStorage !== "undefined") {
    const saved = localStorage.getItem("lang");
    if (saved && STRINGS[saved]) return saved;
  }
  const nav = typeof navigator !== "undefined" && navigator.language ? navigator.language.toLowerCase() : "";
  if (nav.startsWith("ko")) return "ko";
  if (nav.startsWith("zh")) return "zh";
  if (nav.startsWith("es")) return "es";
  return "en";
}
let lang = detectLang();

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
  if (toggle) toggle.value = lang;
}
