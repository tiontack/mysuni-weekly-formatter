const referenceData = window.MYSUNI_REFERENCE_DATA || {};
const defaultGlossary = Array.isArray(referenceData.defaultGlossary) ? referenceData.defaultGlossary : [];
const defaultNames = Array.isArray(referenceData.defaultNames) ? referenceData.defaultNames : [];

const ELEMENT_NODE = 1;
const EMU_PER_PX = 9525; // 1 EMU = 1/9525 px (96 DPI 기준)

function cloneValue(value) {
  if (typeof structuredClone === "function") {
    return structuredClone(value);
  }
  return JSON.parse(JSON.stringify(value));
}

function createId() {
  const cryptoApi = typeof crypto !== "undefined" ? crypto : window.crypto;
  if (cryptoApi && typeof cryptoApi.randomUUID === "function") {
    return cryptoApi.randomUUID();
  }
  return `mysuni-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

function replaceEvery(value, search, replacement) {
  return String(value || "").split(search).join(replacement);
}

const defaultGuideItems = [
  "1. 파일명",
  "- mySUNI Weekly_위클리날짜_조직명.docx로 작성",
  "- 버전이 여러 개일 경우, 파일명 끝에 '_v2'와 같은 형식으로 명시",
  "(예, mySUNI Weekly_260102_변화추진_2.docx)",
  "2. 기본 서식",
  "- 3단 목록 스타일로 구성 ( [카테고리] / n 과제 / ­ 내용 )",
  "- 글꼴 : __카테고리__ – 나눔스퀘어_ac Extrabold (14pt) : __과제__ – 나눔스퀘어_ac Bold (13pt) : __내용__ – 나눔스퀘어_ac (12pt)",
  "- 단계 간 Tab 또는 Shift+Tab으로만 이동",
  "- 카테고리의 경우 텍스트 양 옆에 괄호([ ])를 삽입",
  "- '표'를 삽입하는 경우, 표준 양식 및 가이드 준수하되, 글꼴은 11pt로 조정",
  "3. 날짜, 시간, 기간 표기",
  "- 날짜: YYYY-MM-DD(요일) (예, 2026-01-02(금)) : 년도가 없거나 슬래시(/), 점(.)만으로 된 날짜표기 사용하지 않음",
  "- 시간: HH:MM (예, 14:30) : 24시간제로 표기",
  "- 기간: 시작일~종료일 전체 명기 (예, 2026-01-07(수)~2026-01-13(화)) : '금주', '다음달 초', '차주' 등의 표현은 사용하지 않음",
  "4. 용어, 단위, 약어 표기",
  "- 조직명, 프로그램명, 시스템명, 회의명 등의 고유명사는 동일한 용어를 사용",
  "- 새로운 약어 사용 시 문서 내 최초 1회에 전체 용어+약어를 함께 병기 (예, Performance Indicator(PI))",
  "- 숫자, 단위는 가능한 범위에서 실체 수치를 명확히 기입",
  "5. 기타",
  "- 한 '내용' 안에 여러 정보가 있을 경우 줄바꿈(Shift+Enter)을 이용해 의미 구분",
  "- '이것', '여기', '위와 같음' 등의 모호한 지시어는 사용하지 않음",
  "- 중요 사항은 가급적 문구 앞에 라벨링 (예, [변경], [중요])",
  "- 변경 사항이 있을 경우 '변경 전', '변경 후'를 모두 명시 (예, 변경 전: 2026-12-11(목) 10:30~12:00, 변경 후: 2026-12-15(월) 10:30~12:00)",
];

const STORAGE_KEYS = {
  guide: "mysuni_weekly_guide_items",
  glossary: "mysuni_weekly_glossary",
  names: "mysuni_weekly_names",
  archives: "mysuni_weekly_archives",
  documentArchives: "mysuni_weekly_document_archives",
  aiModel: "mysuni_weekly_openai_model",
  documentDraft: "mysuni_weekly_document_draft_v2",
  editorWide: "mysuni_weekly_editor_wide"
};

const INTEGRATION_ORGS = ["변화추진", "러닝플랫폼", "AI역량육성", "경영관리역량", "미래반도체", "SKMS실천", "리더십개발"];

const labelTranslations = {
  "일자": "Date",
  "부서": "Department",
  "조직": "Organization",
  "카테고리": "Category",
  "기타": "Other",
  "작성 Guide": "Writing Guide",
  "작성 가이드": "Writing Guide",
  "과제명": "Agenda",
  "내용": "Details",
  "비고": "Note",
  "일정": "Schedule",
  "과정명": "Program",
  "학습인원": "Participants",
  "장소": "Location",
  "대상": "Target",
  "목적": "Purpose",
  "계획": "Plan",
  "현황": "Status",
  "결과": "Result"
};

const phraseTranslations = {
  "일시": "Date & Time",
  "일자:": "Date:",
  "부서:": "Department:",
  "조직명": "Organization",
  "변경 전:": "Previous:",
  "변경 후:": "New:",
  "기간:": "Period:",
  "응답률": "response rate",
  "중요": "Important",
  "작성": "Drafted",
  "추가 조사": "Additional survey",
  "조사 결과 분석": "Survey result analysis",
  "대상:": "Target:",
  "장소:": "Location:",
  "목적:": "Purpose:",
  "비고:": "Note:",
  "이전:": "Previous:",
  "변경:": "New:",
  "신규:": "New:",
  "세션": "Session",
  "강연": "Lecture",
  "토픽": "Topic",
  "발표자": "Speaker",
  "발표": "Presentation",
  "예정": "Planned",
  "완료": "Completed",
  "진행 중": "In progress",
  "공유 예정": "To be shared",
  "참석": "Participants",
  "논의": "Discussion",
  "계약": "Contract",
  "업무량": "Workload",
  "계약금액": "Contract amount",
  "상반기": "first half",
  "하반기": "second half",
  "감소": "reduction",
  "증가": "increase"
};

const weekdayMap = {
  월: "Mon",
  화: "Tue",
  수: "Wed",
  목: "Thu",
  금: "Fri",
  토: "Sat",
  일: "Sun"
};

const sentenceTranslations = [
  ["조직별 간담회", "Organization-Level Roundtable Meetings"],
  ["연말 구성원 이동", "Year-End Employee Rotation"],
  ["토요 임원 인사이트 프로그램", "Saturday Executive Insight Program"],
  ["mySUNI 임원 워크숍", "mySUNI Executive Workshop (W/S) for Executives"],
  ["장소 및 상세 Agenda 추후 공유", "Note: Venue and detailed agenda will be shared later"],
  ["구성원/팀장 평가 결과 조직별 안내 및 개별 성과 피드백", "Communication of evaluation results for employees and team leaders, followed by individual performance feedback"],
  ["조직별로 시행, 각 사에 최종 평가 결과 통보", "To be conducted by each organization, and final evaluation results to be reported to each affiliate company"],
  ["임원 평가 결과", "Executive evaluation feedback"],
  ["지원범위 및 금액 협의", "Agreement on scope of support and contract amount"],
  ["설문 결과 분석", "Survey result analysis"],
  ["추가 조사", "Additional survey"],
  ["조직운영 방향", "organizational operation direction"],
  ["주요과제", "key initiatives"]
];

const DEFAULT_TABLE_TEMPLATE = {
  rows: [
    ["과정명", "일정", "학습인원"],
    ["", "", ""],
    ["", "", ""],
    ["", "", ""]
  ]
};
const DETAIL_BULLET_TOKEN = "__MYSUNI_DETAIL_BULLET__";

const DEFAULT_AI_CONFIG = {
  model: "gemini-2.0-flash"
};

const emptyDocumentState = {
  title: "",
  date: "",
  org: "",
  filename: "",
  sourceFileName: "",
  sections: []
};

const state = {
  document: cloneValue(emptyDocumentState), // 페이지 로드/새로고침 시 항상 빈 상태로 시작
  subDocuments: null,
  activeSubDocIndex: 0,
  view: "ko",
  glossary: buildDictionary(loadStoredPairs(STORAGE_KEYS.glossary, defaultGlossary)),
  names: buildDictionary(loadStoredPairs(STORAGE_KEYS.names, defaultNames)),
  guideItems: loadStoredGuideItems(),
  activeSelection: { type: "overview", sectionIndex: null },
  activeTab: "main",
  adminUnlocked: false,
  aiConfig: loadStoredAiConfig(),
  browserTranslator: null,
  aiTranslationCache: { signature: "", document: null },
  aiTranslationStatus: "idle",
  aiTranslationError: "",
  aiTranslationTimer: null,
  aiTranslationRequestId: 0,
  editorComposing: false,
  previewRenderTimer: null,
  draftSaveTimer: null,
  draftStatus: "자동 임시 저장됨",
  lastDraftSignature: "",
  editorWide: loadStoredEditorWide(),
  integrationDate: "",
  integrationSlots: INTEGRATION_ORGS.map((org) => ({ org, submitted: false, document: null })),
  archives: loadStoredArchives(),
  documentArchives: loadStoredDocumentArchives(),
  tableSelection: null,  // { sI, iI, tI, cells: Set<"rowIdx,cellIdx"> }
  outputLang: "ko-only",           // "ko-only" | "en-only" | "ko-en" | "en-ko"
  integrationOutputLang: "ko-en"   // "ko-en" | "en-ko"
};

const els = {
  mainTabButton: document.getElementById("mainTabButton"),
  integrationTabButton: document.getElementById("integrationTabButton"),
  archiveTabButton: document.getElementById("archiveTabButton"),
  translateTabButton: document.getElementById("translateTabButton"),
  guideTabButton: document.getElementById("guideTabButton"),
  adminTabButton: document.getElementById("adminTabButton"),
  mainTab: document.getElementById("mainTab"),
  integrationTab: document.getElementById("integrationTab"),
  archiveTab: document.getElementById("archiveTab"),
  translateTab: document.getElementById("translateTab"),
  translateInput: document.getElementById("translateInput"),
  translateOutput: document.getElementById("translateOutput"),
  translateCharCount: document.getElementById("translateCharCount"),
  translateClearBtn: document.getElementById("translateClearBtn"),
  translateCopyBtn: document.getElementById("translateCopyBtn"),
  translateApiStatus: document.getElementById("translateApiStatus"),
  translateLoadingDot: document.getElementById("translateLoadingDot"),
  guideTab: document.getElementById("guideTab"),
  adminTab: document.getElementById("adminTab"),
  docxInput: document.getElementById("docxInput"),
  titleInput: document.getElementById("titleInput"),
  dateInput: document.getElementById("dateInput"),
  datePicker: document.getElementById("datePicker"),
  datePickerButton: document.getElementById("datePickerButton"),
  orgInput: document.getElementById("orgInput"),
  filenamePreview: document.getElementById("filenamePreview"),
  uploadedFileName: document.getElementById("uploadedFileName"),
  clearDocxButton: document.getElementById("clearDocxButton"),
  sectionSelect: document.getElementById("sectionSelect"),
  sectionItemCount: document.getElementById("sectionItemCount"),
  editorRoot: document.getElementById("editorRoot"),
  orgTabBar: document.getElementById("orgTabBar"),
  sheetSticky: document.querySelector(".sheet-sticky"),
  previewTitle: document.getElementById("previewTitle"),
  previewDate: document.getElementById("previewDate"),
  previewOrg: document.getElementById("previewOrg"),
  dateLabel: document.getElementById("dateLabel"),
  orgLabel: document.getElementById("orgLabel"),
  previewBody: document.getElementById("previewBody"),
  integrationDateInput: document.getElementById("integrationDateInput"),
  integrationDatePicker: document.getElementById("integrationDatePicker"),
  integrationDatePickerButton: document.getElementById("integrationDatePickerButton"),
  integrationStatusList: document.getElementById("integrationStatusList"),
  integrationExportButton: document.getElementById("integrationExportButton"),
  integrationArchiveButton: document.getElementById("integrationArchiveButton"),
  archiveList: document.getElementById("archiveList"),
  documentArchiveList: document.getElementById("documentArchiveList"),
  guideList: document.getElementById("guideList"),
  memberSearchInput: document.getElementById("memberSearchInput"),
  memberSearchResult: document.getElementById("memberSearchResult"),
  glossaryGuideTable: document.getElementById("glossaryGuideTable"),
  koreanViewButton: document.getElementById("koreanViewButton"),
  englishViewButton: document.getElementById("englishViewButton"),
  outputLangGroup: document.getElementById("outputLangGroup"),
  integrationLangGroup: document.getElementById("integrationLangGroup"),
  glossaryCount: document.getElementById("glossaryCount"),
  nameCount: document.getElementById("nameCount"),
  translationStatus: document.getElementById("translationStatus"),
  archiveSelectButton: document.getElementById("archiveSelectButton"),
  archiveSelectPanel:  document.getElementById("archiveSelectPanel"),
  archiveSelectWrap:   document.getElementById("archiveSelectWrap"),
  addSectionButton: document.getElementById("addSectionButton"),
  removeSectionButton: document.getElementById("removeSectionButton"),
  moveSectionUpButton: document.getElementById("moveSectionUpButton"),
  moveSectionDownButton: document.getElementById("moveSectionDownButton"),
  toggleEditorWidthButton: document.getElementById("toggleEditorWidthButton"),
  draftStatus: document.getElementById("draftStatus"),
  saveDraftButton: document.getElementById("saveDraftButton"),
  printButton: document.getElementById("printButton"),
  exportWordButton: document.getElementById("exportWordButton"),
  adminLocked: document.getElementById("adminLocked"),
  adminUnlocked: document.getElementById("adminUnlocked"),
  adminPasswordInput: document.getElementById("adminPasswordInput"),
  adminLoginButton: document.getElementById("adminLoginButton"),
  adminLogoutButton: document.getElementById("adminLogoutButton"),
  adminGuideInput: document.getElementById("adminGuideInput"),
  adminGlossaryInput: document.getElementById("adminGlossaryInput"),
  adminNamesInput: document.getElementById("adminNamesInput"),
  adminAiModelInput: document.getElementById("adminAiModelInput"),
  adminSaveButton: document.getElementById("adminSaveButton"),
  adminResetButton: document.getElementById("adminResetButton")
};

bootstrap();

function updateTopChromeHeight() {
  const el = document.querySelector(".top-chrome");
  if (!el) return;
  const h = el.getBoundingClientRect().height;
  document.documentElement.style.setProperty("--top-chrome-height", h + "px");
}

function bootstrap() {
  try {
    state.lastDraftSignature = getDocumentSignature(state.document);
    syncInputsFromState();
    bindEvents();
    render();
    // .top-chrome 실제 높이를 측정해 CSS 변수 동기화 (sticky 오프셋 정확성 확보)
    updateTopChromeHeight();
    const topChrome = document.querySelector(".top-chrome");
    if (topChrome && window.ResizeObserver) {
      new ResizeObserver(updateTopChromeHeight).observe(topChrome);
    } else {
      window.addEventListener("resize", updateTopChromeHeight);
    }
  } catch (error) {
    console.error(error);
    showFatalError(error);
  }
}

function showFatalError(error) {
  const message = error?.message || "알 수 없는 오류";
  const shell = document.querySelector(".page-shell");
  if (!shell) {
    return;
  }
  const existing = document.getElementById("fatalErrorBanner");
  if (existing) {
    existing.textContent = `페이지 초기화 중 오류가 발생했습니다: ${message}`;
    return;
  }
  const banner = document.createElement("div");
  banner.id = "fatalErrorBanner";
  banner.className = "panel";
  banner.innerHTML = `
    <div class="panel-head">
      <h2>페이지 오류</h2>
    </div>
    <p class="hint">페이지 초기화 중 오류가 발생했습니다: ${escapeHtml(message)}</p>
    <p class="hint">브라우저를 새로고침한 뒤 다시 시도해 주세요. 문제가 계속되면 관리자에게 알려주세요.</p>
  `;
  shell.prepend(banner);
}

function bindEvents() {
  els.mainTabButton?.addEventListener("click", () => switchTab("main"));
  els.integrationTabButton?.addEventListener("click", () => switchTab("integration"));
  els.archiveTabButton?.addEventListener("click", () => switchTab("archive"));
  els.translateTabButton?.addEventListener("click", () => switchTab("translate"));
  els.guideTabButton?.addEventListener("click", () => switchTab("guide"));
  els.adminTabButton?.addEventListener("click", () => switchTab("admin"));

  els.titleInput?.addEventListener("input", () => {
    const previousSuggested = buildSuggestedFilename(state.document);
    state.document.title = els.titleInput?.value.trim() || "";
    syncFilenameWithSuggestion(previousSuggested);
    render();
  });

  els.dateInput?.addEventListener("input", () => {
    const previousSuggested = buildSuggestedFilename(state.document);
    state.document.date = els.dateInput?.value.trim() || "";
    if (els.datePicker) {
      els.datePicker.value = extractDateValue(state.document.date);
    }
    syncFilenameWithSuggestion(previousSuggested);
    render();
  });

  els.datePicker?.addEventListener("input", () => {
    const previousSuggested = buildSuggestedFilename(state.document);
    state.document.date = formatPickedDate(els.datePicker?.value || "");
    syncFilenameWithSuggestion(previousSuggested);
    render();
  });
  if (els.datePickerButton) {
    els.datePickerButton.addEventListener("click", () => {
      if (typeof els.datePicker?.showPicker === "function") {
        els.datePicker.showPicker();
      } else if (els.datePicker) {
        els.datePicker.focus();
        els.datePicker.click();
      }
    });
  }

  els.orgInput?.addEventListener("input", () => {
    const previousSuggested = buildSuggestedFilename(state.document);
    state.document.org = els.orgInput?.value.trim() || "";
    syncFilenameWithSuggestion(previousSuggested);
    render();
  });

  els.filenamePreview?.addEventListener("input", () => {
    state.document.filename = els.filenamePreview?.value.trim() || "";
    render();
  });

  els.docxInput?.addEventListener("change", onDocxSelected);
  els.clearDocxButton?.addEventListener("click", clearUploadedDocument);

  // 저장 기록 드롭다운
  els.archiveSelectButton?.addEventListener("click", (e) => {
    e.stopPropagation();
    const panel = els.archiveSelectPanel;
    if (!panel) return;
    if (panel.hidden) {
      renderArchiveSelectPanel();
      panel.hidden = false;
    } else {
      panel.hidden = true;
    }
  });
  els.archiveSelectPanel?.addEventListener("click", onArchiveSelectPanelClick);
  document.addEventListener("click", (e) => {
    if (els.archiveSelectWrap && !els.archiveSelectWrap.contains(e.target)) {
      if (els.archiveSelectPanel) els.archiveSelectPanel.hidden = true;
    }
  });
  els.orgTabBar?.addEventListener("click", (e) => {
    const tabBtn = e.target.closest("[data-action='switch-org-tab']");
    if (tabBtn) {
      switchToSubDoc(Number(tabBtn.dataset.subDocIndex));
      return;
    }
    const moveBtn = e.target.closest("[data-action='org-move-left'], [data-action='org-move-right']");
    if (!moveBtn) return;
    const idx = Number(moveBtn.dataset.subDocIndex);
    const docs = state.subDocuments;
    if (!docs) return;
    if (moveBtn.dataset.action === "org-move-left" && idx > 0) {
      [docs[idx - 1], docs[idx]] = [docs[idx], docs[idx - 1]];
      if (state.activeSubDocIndex === idx) state.activeSubDocIndex = idx - 1;
      else if (state.activeSubDocIndex === idx - 1) state.activeSubDocIndex = idx;
    } else if (moveBtn.dataset.action === "org-move-right" && idx < docs.length - 1) {
      [docs[idx], docs[idx + 1]] = [docs[idx + 1], docs[idx]];
      if (state.activeSubDocIndex === idx) state.activeSubDocIndex = idx + 1;
      else if (state.activeSubDocIndex === idx + 1) state.activeSubDocIndex = idx;
    }
    state.document = state.subDocuments[state.activeSubDocIndex];
    renderOrgTabBar();
  });

  els.koreanViewButton?.addEventListener("click", () => {
    state.view = "ko";
    state.outputLang = "ko-only";
    render();
  });

  els.englishViewButton?.addEventListener("click", () => {
    state.view = "en";
    state.outputLang = "en-only";
    render();
  });

  els.outputLangGroup?.addEventListener("click", (e) => {
    const btn = e.target.closest("[data-output-lang]");
    if (!btn) return;
    const val = btn.dataset.outputLang;
    state.outputLang = val;
    // view = first language in the order
    state.view = val.startsWith("ko") ? "ko" : "en";
    render();
  });

  els.addSectionButton?.addEventListener("click", () => {
    state.document.sections.push(createEmptySection());
    state.activeSelection = { type: "section", sectionIndex: state.document.sections.length - 1 };
    render();
  });
  els.removeSectionButton?.addEventListener("click", removeActiveSection);
  els.moveSectionUpButton?.addEventListener("click", () => {
    const idx = getEditingSectionIndex();
    if (idx > 0) {
      const secs = state.document.sections;
      [secs[idx - 1], secs[idx]] = [secs[idx], secs[idx - 1]];
      state.activeSelection = { type: "section", sectionIndex: idx - 1 };
      render();
    }
  });
  els.moveSectionDownButton?.addEventListener("click", () => {
    const idx = getEditingSectionIndex();
    const secs = state.document.sections;
    if (idx < secs.length - 1) {
      [secs[idx], secs[idx + 1]] = [secs[idx + 1], secs[idx]];
      state.activeSelection = { type: "section", sectionIndex: idx + 1 };
      render();
    }
  });
  els.sectionSelect?.addEventListener("change", () => {
    const sectionIndex = Number(els.sectionSelect?.value);
    if (Number.isInteger(sectionIndex) && sectionIndex >= 0) {
      state.activeSelection = { type: "section", sectionIndex };
      render();
    }
  });
  els.saveDraftButton?.addEventListener("click", () => {
    saveDocumentDraft(true);
    saveDocumentArchiveSnapshot();
    renderDraftStatus();
    alert("현재 작성 내용이 저장되었습니다.");
  });
  if (els.toggleEditorWidthButton) {
    els.toggleEditorWidthButton.addEventListener("click", () => {
      state.editorWide = !state.editorWide;
      localStorage.setItem(STORAGE_KEYS.editorWide, state.editorWide ? "1" : "0");
      render();
    });
  }

  els.printButton?.addEventListener("click", printDocument);
  els.exportWordButton?.addEventListener("click", exportWordDocument);
  els.integrationDateInput?.addEventListener("input", () => {
    state.integrationDate = els.integrationDateInput?.value.trim() || "";
    syncIntegrationDatePicker();
  });
  els.integrationDatePicker?.addEventListener("input", () => {
    state.integrationDate = formatPickedDate(els.integrationDatePicker?.value || "");
    renderIntegrationTab();
  });
  if (els.integrationDatePickerButton) {
    els.integrationDatePickerButton.addEventListener("click", () => {
      if (typeof els.integrationDatePicker?.showPicker === "function") {
        els.integrationDatePicker.showPicker();
      } else if (els.integrationDatePicker) {
        els.integrationDatePicker.focus();
        els.integrationDatePicker.click();
      }
    });
  }
  els.integrationStatusList?.addEventListener("change", onIntegrationUpload);
  els.integrationExportButton?.addEventListener("click", exportIntegratedWordDocument);
  els.integrationArchiveButton?.addEventListener("click", saveIntegrationArchive);
  els.integrationLangGroup?.addEventListener("click", (e) => {
    const btn = e.target.closest("[data-integ-lang]");
    if (!btn) return;
    state.integrationOutputLang = btn.dataset.integLang;
    if (els.integrationLangGroup) {
      els.integrationLangGroup.querySelectorAll("[data-integ-lang]").forEach((b) => {
        b.classList.toggle("active", b.dataset.integLang === state.integrationOutputLang);
      });
    }
  });
  els.archiveList?.addEventListener("click", onArchiveClick);
  els.documentArchiveList?.addEventListener("click", onDocumentArchiveClick);
  els.editorRoot?.addEventListener("click", onEditorClick);
  els.editorRoot?.addEventListener("keydown", onEditorKeydown);
  els.editorRoot?.addEventListener("input", onEditorInput);
  els.editorRoot?.addEventListener("change", onEditorInput);
  els.editorRoot?.addEventListener("focusout", () => {
    schedulePreviewRender();
  });
  els.editorRoot?.addEventListener("compositionstart", () => {
    state.editorComposing = true;
  });
  els.editorRoot?.addEventListener("compositionend", (event) => {
    state.editorComposing = false;
    onEditorInput(event);
  });
  els.previewBody?.addEventListener("click", onPreviewClick);
  els.memberSearchInput?.addEventListener("input", renderGuidePage);
  els.sheetSticky?.addEventListener("click", () => {
    state.activeSelection = { type: "overview", sectionIndex: null };
    renderSectionList();
    renderEditor();
    renderPreview();
  });
  els.adminLoginButton?.addEventListener("click", unlockAdmin);
  els.adminLogoutButton?.addEventListener("click", lockAdmin);
  els.adminSaveButton?.addEventListener("click", saveAdminData);
  els.adminResetButton?.addEventListener("click", resetAdminData);
}

async function onDocxSelected(event) {
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }

  try {
    const arrayBuffer = await file.arrayBuffer();
    const parsed = await parseDocx(arrayBuffer);

    if (parsed?._multiDoc && parsed.documents.length > 1) {
      // 멀티 조직 모드
      const docs = parsed.documents;
      docs.forEach((doc) => {
        doc.sourceFileName = file.name;
        doc.filename = buildSuggestedFilename(doc);
      });
      state.subDocuments = docs;
      state.activeSubDocIndex = 0;
      state.document = docs[0];
    } else {
      // 단일 조직 모드
      state.subDocuments = null;
      state.activeSubDocIndex = 0;
      parsed.filename = buildSuggestedFilename(parsed);
      parsed.sourceFileName = file.name;
      state.document = parsed;
    }

    state.activeSelection = state.document.sections.length
      ? { type: "section", sectionIndex: 0 }
      : { type: "overview", sectionIndex: null };
    syncInputsFromState();
    render();

    if (!state.document.sections.length) {
      alert("파일은 업로드되었지만 본문 구조를 완전히 인식하지 못했습니다. 개요만 반영되었거나 본문은 수동 편집이 필요할 수 있습니다.");
    }
  } catch (error) {
    console.error(error);
    alert("docx 파싱 중 오류가 발생했습니다. 텍스트 붙여넣기 방식으로 다시 시도해 주세요.");
  }
}

// ── 표 셀 선택 관리 ──────────────────────────────────────────────────────────

function selKey(rowIndex, cellIndex) {
  return `${rowIndex},${cellIndex}`;
}

function isCellSelected(sI, iI, tI, rI, cI) {
  const s = state.tableSelection;
  return s && s.sI === sI && s.iI === iI && s.tI === tI && s.cells.has(selKey(rI, cI));
}

function selectTableCell(sI, iI, tI, rI, cI, addMode) {
  const key = selKey(rI, cI);
  const s = state.tableSelection;
  if (addMode && s && s.sI === sI && s.iI === iI && s.tI === tI) {
    if (s.cells.has(key)) {
      s.cells.delete(key);
      if (!s.cells.size) state.tableSelection = null;
    } else {
      s.cells.add(key);
    }
  } else {
    state.tableSelection = { sI, iI, tI, cells: new Set([key]) };
  }
  renderEditor();
}

function selectTableRow(sI, iI, tI, rI, addMode) {
  const table = state.document.sections[sI].items[iI].tables[tI];
  if (!table?.rows?.[rI]) return;
  const keys = table.rows[rI].map((_, ci) => selKey(rI, ci));
  const s = state.tableSelection;
  if (addMode && s && s.sI === sI && s.iI === iI && s.tI === tI) {
    keys.forEach((k) => s.cells.add(k));
  } else {
    state.tableSelection = { sI, iI, tI, cells: new Set(keys) };
  }
  renderEditor();
}

function clearTableSelection() {
  state.tableSelection = null;
}

function getSelectedCellCoords() {
  if (!state.tableSelection) return [];
  return [...state.tableSelection.cells].map((k) => {
    const [r, c] = k.split(",").map(Number);
    return { r, c };
  });
}

function getSelectedRowIndices() {
  return [...new Set(getSelectedCellCoords().map((coord) => coord.r))].sort((a, b) => a - b);
}

function getSelectedColIndices() {
  return [...new Set(getSelectedCellCoords().map((coord) => coord.c))].sort((a, b) => a - b);
}

function selectTableCol(sI, iI, tI, cI, addMode) {
  const table = state.document.sections[sI].items[iI].tables[tI];
  if (!table?.rows) return;
  // rowSpan=0인 숨김 셀을 제외하고 해당 열의 셀 키 수집
  const keys = table.rows
    .map((row, ri) => {
      if (cI >= row.length) return null;
      return normalizeTableCellData(row[cI]).rowSpan !== 0 ? selKey(ri, cI) : null;
    })
    .filter(Boolean);
  const s = state.tableSelection;
  if (addMode && s && s.sI === sI && s.iI === iI && s.tI === tI) {
    keys.forEach((k) => s.cells.add(k));
  } else {
    state.tableSelection = { sI, iI, tI, cells: new Set(keys) };
  }
  renderEditor();
}

// 선택된 셀 합치기 (같은 행, 인접 셀)
function mergeSelectedCells() {
  const s = state.tableSelection;
  if (!s || s.cells.size < 2) return;
  const coords = getSelectedCellCoords();
  const rowSet = new Set(coords.map((c) => c.r));
  if (rowSet.size !== 1) {
    alert("같은 행의 셀만 합칠 수 있습니다.");
    return;
  }
  const rowIdx = [...rowSet][0];
  const cellIndices = coords.map((c) => c.c).sort((a, b) => a - b);
  for (let i = 1; i < cellIndices.length; i++) {
    if (cellIndices[i] !== cellIndices[i - 1] + 1) {
      alert("인접한 셀만 합칠 수 있습니다.");
      return;
    }
  }
  const table = state.document.sections[s.sI].items[s.iI].tables[s.tI];
  const row = table.rows[rowIdx];
  // Bug B fix: 합칠 셀들의 rowSpan이 모두 같아야 함
  const rowSpans = cellIndices.map((ci) => normalizeTableCellData(row[ci]).rowSpan);
  if (new Set(rowSpans).size > 1) {
    alert("가로로 합칠 셀의 세로 높이(rowSpan)가 모두 같아야 합니다.\n먼저 세로 병합을 해제해 주세요.");
    return;
  }
  let totalColSpan = 0;
  const texts = [];
  for (const ci of cellIndices) {
    const cell = normalizeTableCellData(row[ci]);
    totalColSpan += cell.colSpan;
    if (cell.text) texts.push(cell.text);
  }
  const firstCell = normalizeTableCellData(row[cellIndices[0]]);
  row[cellIndices[0]] = { text: texts.join("\n"), colSpan: totalColSpan, rowSpan: firstCell.rowSpan };
  for (let i = cellIndices.length - 1; i >= 1; i--) {
    row.splice(cellIndices[i], 1);
  }
  clearTableSelection();
  render();
}

// 선택된 셀 분리 (colspan > 1)
function splitSelectedCell() {
  const s = state.tableSelection;
  if (!s || s.cells.size !== 1) return;
  const [{ r, c }] = getSelectedCellCoords();
  const table = state.document.sections[s.sI].items[s.iI].tables[s.tI];
  const row = table.rows[r];
  const cell = normalizeTableCellData(row[c]);
  if (cell.colSpan <= 1) return;
  const newCells = [{ text: cell.text, colSpan: 1, rowSpan: cell.rowSpan, align: cell.align || null }];
  for (let i = 1; i < cell.colSpan; i++) {
    newCells.push({ text: "", colSpan: 1, rowSpan: 1, align: null });
  }
  row.splice(c, 1, ...newCells);
  clearTableSelection();
  render();
}

// 선택된 셀 세로 합치기 (rowspan 증가)
function mergeSelectedCellsVertical() {
  const s = state.tableSelection;
  if (!s || s.cells.size < 2) return;
  const coords = getSelectedCellCoords();
  const colSet = new Set(coords.map((co) => co.c));
  if (colSet.size !== 1) {
    alert("같은 열의 셀만 세로로 합칠 수 있습니다.");
    return;
  }
  const colIdx = [...colSet][0];
  const rowIndices = coords.map((co) => co.r).sort((a, b) => a - b);
  for (let i = 1; i < rowIndices.length; i++) {
    if (rowIndices[i] !== rowIndices[i - 1] + 1) {
      alert("인접한 행의 셀만 합칠 수 있습니다.");
      return;
    }
  }
  const table = state.document.sections[s.sI].items[s.iI].tables[s.tI];
  // Bug A fix: 합칠 셀들의 colSpan이 모두 같아야 함
  const colSpans = rowIndices.map((ri) => normalizeTableCellData(table.rows[ri][colIdx]).colSpan);
  if (new Set(colSpans).size > 1) {
    alert("세로로 합칠 셀의 가로 너비(colSpan)가 모두 같아야 합니다.\n먼저 가로 병합을 해제해 주세요.");
    return;
  }
  const texts = [];
  for (const ri of rowIndices) {
    const cell = normalizeTableCellData(table.rows[ri][colIdx]);
    if (cell.text) texts.push(cell.text);
  }
  const firstCell = normalizeTableCellData(table.rows[rowIndices[0]][colIdx]);
  table.rows[rowIndices[0]][colIdx] = {
    text: texts.join("\n"),
    colSpan: firstCell.colSpan,
    rowSpan: rowIndices.length,
    align: firstCell.align || null
  };
  for (let i = 1; i < rowIndices.length; i++) {
    table.rows[rowIndices[i]][colIdx] = { text: "", colSpan: firstCell.colSpan, rowSpan: 0, align: null };
  }
  clearTableSelection();
  render();
}

// 선택된 셀 세로 분리 (rowspan → 1)
function splitSelectedCellVertical() {
  const s = state.tableSelection;
  if (!s || s.cells.size !== 1) return;
  const [{ r, c }] = getSelectedCellCoords();
  const table = state.document.sections[s.sI].items[s.iI].tables[s.tI];
  const cell = normalizeTableCellData(table.rows[r][c]);
  if (cell.rowSpan <= 1) return;
  table.rows[r][c] = { text: cell.text, colSpan: cell.colSpan, rowSpan: 1, align: cell.align || null };
  for (let ri = r + 1; ri < r + cell.rowSpan; ri++) {
    if (table.rows[ri]) {
      table.rows[ri][c] = { text: "", colSpan: cell.colSpan, rowSpan: 1, align: null };
    }
  }
  clearTableSelection();
  render();
}

// 선택된 셀 정렬 적용
function alignSelectedCells(alignValue) {
  const s = state.tableSelection;
  if (!s) return;
  const coords = getSelectedCellCoords();
  const table = state.document.sections[s.sI].items[s.iI].tables[s.tI];
  for (const { r, c } of coords) {
    if (table.rows[r]?.[c] !== undefined) {
      const norm = normalizeTableCellData(table.rows[r][c]);
      table.rows[r][c] = { ...norm, align: alignValue };
    }
  }
  render();
}

// 행 삽입 (Bug C fix: rowSpan이 삽입 위치를 가로지르는 셀 처리)
function insertTableRowAt(sI, iI, tI, insertIdx) {
  const table = state.document.sections[sI].items[iI].tables[tI];
  const colCount = getTableColumnCount(table);
  const newRow = [];
  for (let ci = 0; ci < colCount; ci++) {
    // 위쪽 행 중 rowSpan이 insertIdx를 가로지르는 셀이 있는지 탐색
    let spannerRi = null;
    let spannerCell = null;
    for (let ri = insertIdx - 1; ri >= 0; ri--) {
      if (ci >= table.rows[ri].length) continue;
      const cell = normalizeTableCellData(table.rows[ri][ci]);
      if (cell.rowSpan > 0) {
        if (ri + cell.rowSpan > insertIdx) {
          spannerRi = ri;
          spannerCell = cell;
        }
        break;
      }
    }
    if (spannerCell !== null) {
      // 기존 head 셀의 rowSpan을 +1
      const headCell = normalizeTableCellData(table.rows[spannerRi][ci]);
      table.rows[spannerRi][ci] = { ...headCell, rowSpan: headCell.rowSpan + 1 };
      // 새 행에는 hidden 셀(rowSpan: 0)
      newRow.push({ text: "", colSpan: spannerCell.colSpan, rowSpan: 0, align: null });
    } else {
      newRow.push({ text: "", colSpan: 1, rowSpan: 1, align: null });
    }
  }
  table.rows.splice(insertIdx, 0, newRow);
  clearTableSelection();
  render();
}

// 선택된 행 삭제 (Bug D fix: rowSpan 그룹 부분 삭제 방지)
function deleteSelectedTableRows() {
  const s = state.tableSelection;
  if (!s) return;
  const table = state.document.sections[s.sI].items[s.iI].tables[s.tI];
  const rowIndices = getSelectedRowIndices().sort((a, b) => a - b);
  if (table.rows.length - rowIndices.length < 1) {
    alert("표에는 최소 1개의 행이 있어야 합니다.");
    return;
  }
  const rowSet = new Set(rowIndices);
  // rowSpan 그룹이 선택 범위를 벗어나는지 확인
  for (const ri of rowIndices) {
    for (let ci = 0; ci < table.rows[ri].length; ci++) {
      const cell = normalizeTableCellData(table.rows[ri][ci]);
      if (cell.rowSpan > 1) {
        // head 셀: span 범위 내 모든 행이 선택돼 있어야 함
        for (let offset = 1; offset < cell.rowSpan; offset++) {
          if (!rowSet.has(ri + offset)) {
            alert("합쳐진 셀(rowSpan)이 포함된 행은 병합된 전체 행을 함께 선택해야 삭제할 수 있습니다.");
            return;
          }
        }
      }
      if (cell.rowSpan === 0) {
        // hidden 셀: head 행도 함께 선택돼 있어야 함
        for (let headRi = ri - 1; headRi >= 0; headRi--) {
          if (ci >= table.rows[headRi].length) continue;
          const headCell = normalizeTableCellData(table.rows[headRi][ci]);
          if (headCell.rowSpan > 0) {
            if (!rowSet.has(headRi)) {
              alert("합쳐진 셀(rowSpan)이 포함된 행은 병합된 전체 행을 함께 선택해야 삭제할 수 있습니다.");
              return;
            }
            break;
          }
        }
      }
    }
  }
  rowIndices.reverse().forEach((ri) => table.rows.splice(ri, 1));
  clearTableSelection();
  render();
}

// 열 삽입 (insertBefore=true: 선택 열 앞에, false: 선택 열 뒤에)
function insertTableColAt(sI, iI, tI, colIdx, insertBefore = true) {
  const table = state.document.sections[sI].items[iI].tables[tI];
  if (!table?.rows) return;
  const insertIdx = insertBefore ? colIdx : colIdx + 1;
  table.rows.forEach((row) => {
    row.splice(insertIdx, 0, { text: "", colSpan: 1, rowSpan: 1 });
  });
  clearTableSelection();
  render();
}

// 선택된 열 삭제 (Bug E fix: colSpan 셀이 선택 범위를 벗어나는지 확인)
function deleteSelectedTableCols() {
  const s = state.tableSelection;
  if (!s) return;
  const table = state.document.sections[s.sI].items[s.iI].tables[s.tI];
  const colIndices = getSelectedColIndices().sort((a, b) => a - b);
  const totalCols = getTableColumnCount(table);
  if (totalCols - colIndices.length < 1) {
    alert("표에는 최소 1개의 열이 있어야 합니다.");
    return;
  }
  const colSet = new Set(colIndices);
  // colSpan 그룹이 선택 범위를 벗어나는지 확인
  for (const row of table.rows) {
    for (let ci = 0; ci < row.length; ci++) {
      const cell = normalizeTableCellData(row[ci]);
      if (cell.colSpan > 1 && colSet.has(ci)) {
        for (let offset = 1; offset < cell.colSpan; offset++) {
          if (!colSet.has(ci + offset)) {
            alert("합쳐진 셀(colSpan)이 포함된 열은 병합된 전체 열을 함께 선택해야 삭제할 수 있습니다.");
            return;
          }
        }
      }
    }
  }
  colIndices.reverse().forEach((ci) => {
    table.rows.forEach((row) => {
      if (ci < row.length) row.splice(ci, 1);
    });
  });
  clearTableSelection();
  render();
}

function switchToSubDoc(index) {
  if (!state.subDocuments || index < 0 || index >= state.subDocuments.length) {
    return;
  }
  state.activeSubDocIndex = index;
  state.document = state.subDocuments[index];
  state.activeSelection = state.document.sections.length
    ? { type: "section", sectionIndex: 0 }
    : { type: "overview", sectionIndex: null };
  syncInputsFromState();
  render();
}

// ── 표 에디터 렌더링 ──────────────────────────────────────────────────────────
function renderEditorTable(table, sI, iI, tI) {
  const sel = state.tableSelection;
  const isThisTable = sel && sel.sI === sI && sel.iI === iI && sel.tI === tI;
  const selCount = isThisTable ? sel.cells.size : 0;

  // 툴바 버튼 활성화 판단
  let canMerge = false;       // 가로 합치기 (같은 행, 인접 셀)
  let canSplit = false;       // 가로 분리 (colSpan > 1)
  let canRowMerge = false;    // 세로 합치기 (같은 열, 인접 행)
  let canRowSplit = false;    // 세로 분리 (rowSpan > 1)
  let currentAlign = null;    // 현재 선택 셀 정렬
  let toolbarRowIdx = 0;
  let canColOp = false;       // 열 삽입/삭제 가능 (같은 열 선택)
  let toolbarColIdx = 0;

  if (isThisTable && selCount >= 1) {
    const coords = getSelectedCellCoords();
    const rowIndices = getSelectedRowIndices();
    toolbarRowIdx = rowIndices[0] ?? 0;

    if (selCount === 1) {
      const [{ r, c }] = coords;
      const cell = normalizeTableCellData(table.rows[r]?.[c]);
      if (cell.colSpan > 1) canSplit = true;
      if (cell.rowSpan > 1) canRowSplit = true;
      currentAlign = cell.align || "left";
    }

    if (selCount >= 2) {
      // 가로 합치기: 같은 행, 인접 열
      const rowSet = new Set(coords.map((co) => co.r));
      if (rowSet.size === 1) {
        const sortedCols = coords.map((co) => co.c).sort((a, b) => a - b);
        let adjacent = true;
        for (let i = 1; i < sortedCols.length; i++) {
          if (sortedCols[i] !== sortedCols[i - 1] + 1) { adjacent = false; break; }
        }
        if (adjacent) canMerge = true;
      }
      // 세로 합치기: 같은 열, 인접 행
      const colSet = new Set(coords.map((co) => co.c));
      if (colSet.size === 1) {
        const sortedRows = coords.map((co) => co.r).sort((a, b) => a - b);
        let consecutive = true;
        for (let i = 1; i < sortedRows.length; i++) {
          if (sortedRows[i] !== sortedRows[i - 1] + 1) { consecutive = false; break; }
        }
        if (consecutive) canRowMerge = true;
      }
    }
    // 열 작업: 모든 선택된 셀이 같은 열 인덱스인 경우
    const selColSet = new Set(coords.map((c) => c.c));
    if (selColSet.size === 1) {
      canColOp = true;
      toolbarColIdx = [...selColSet][0];
    }
  }

  // 정렬 버튼 active 판단
  const alignBtns = ["left", "center", "right"].map((al) => {
    const label = al === "left" ? "◀ 좌" : al === "center" ? "● 중앙" : "▶ 우";
    const isActive = currentAlign === al;
    return `<button class="sel-btn sel-align-btn${isActive ? " active" : ""}" data-action="align-selected-cells" data-align="${al}">${label}</button>`;
  }).join("");

  // 컨텍스트 툴바
  const toolbarHtml = (isThisTable && selCount > 0) ? `
    <div class="table-sel-toolbar">
      <span class="table-sel-label">${selCount}개 셀 선택됨</span>
      <span class="sel-sep">│</span>
      <span class="sel-group-label">가로</span>
      <button class="sel-btn sel-merge-btn" data-action="merge-selected-cells"${canMerge ? "" : " disabled"}>⬛ 합치기</button>
      <button class="sel-btn" data-action="split-selected-cell"${canSplit ? "" : " disabled"}>⬜ 분리</button>
      <span class="sel-sep">│</span>
      <span class="sel-group-label">세로</span>
      <button class="sel-btn sel-merge-btn" data-action="merge-selected-cells-vertical"${canRowMerge ? "" : " disabled"}>⬛ 합치기</button>
      <button class="sel-btn" data-action="split-selected-cell-vertical"${canRowSplit ? "" : " disabled"}>⬜ 분리</button>
      <span class="sel-sep">│</span>
      <span class="sel-group-label">정렬</span>
      ${alignBtns}
      <span class="sel-sep">│</span>
      <button class="sel-btn" data-action="insert-row-above"
        data-section-index="${sI}" data-item-index="${iI}"
        data-table-index="${tI}" data-row-index="${toolbarRowIdx}">↑ 행 추가</button>
      <button class="sel-btn" data-action="insert-row-below"
        data-section-index="${sI}" data-item-index="${iI}"
        data-table-index="${tI}" data-row-index="${toolbarRowIdx}">↓ 행 추가</button>
      <button class="sel-btn sel-danger-btn" data-action="delete-selected-rows">행 삭제</button>
      ${canColOp ? `<span class="sel-sep">│</span>
      <span class="sel-group-label">열</span>
      <button class="sel-btn" data-action="insert-col-before"
        data-section-index="${sI}" data-item-index="${iI}"
        data-table-index="${tI}" data-col-index="${toolbarColIdx}">← 삽입</button>
      <button class="sel-btn" data-action="insert-col-after"
        data-section-index="${sI}" data-item-index="${iI}"
        data-table-index="${tI}" data-col-index="${toolbarColIdx}">삽입 →</button>
      <button class="sel-btn sel-danger-btn" data-action="delete-selected-cols">열 삭제</button>` : ""}
      <button class="sel-btn sel-clear-btn" data-action="clear-table-selection" title="선택 해제">✕</button>
    </div>
  ` : `
    <div class="table-sel-toolbar table-sel-hint">
      <span class="table-sel-label">셀 클릭으로 편집 · 셀 테두리 클릭으로 선택 · 열 머리(A,B,C) 클릭으로 열 전체 선택</span>
    </div>
  `;

  // 행 렌더링
  const rowsHtml = table.rows.map((row, rI) => {
    const isRowSelected = isThisTable && row.some((_, cI) => isCellSelected(sI, iI, tI, rI, cI));

    const rowSelectorHtml = `<td
      class="row-selector-col${isRowSelected ? " row-selected" : ""}"
      data-action="select-table-row"
      data-section-index="${sI}" data-item-index="${iI}"
      data-table-index="${tI}" data-row-index="${rI}"
      title="행 전체 선택 (Ctrl+클릭으로 다중 행 선택)"
    >${rI + 1}</td>`;

    const cellsHtml = row.map((cell, cI) => {
      const normalized = normalizeTableCellData(cell);
      const rs = normalized.rowSpan;
      const cs = normalized.colSpan;

      if (rs === 0) {
        // 세로 병합된 숨김 셀: 부모 셀이 이미 rowspan 속성을 가지므로
        // 추가 <td>를 렌더링하면 브라우저가 열을 초과해 표가 찌그러짐 → 빈 문자열 반환
        return "";
      }

      const selected = isCellSelected(sI, iI, tI, rI, cI);
      const csAttr = cs > 1 ? ` colspan="${cs}"` : "";
      const rsAttr = rs > 1 ? ` rowspan="${rs}"` : "";
      const selClass = selected ? " cell-selected" : "";
      const alignClass = normalized.align ? ` align-${normalized.align}` : "";
      // 병합 뱃지
      const mergeBadge = (cs > 1 || rs > 1)
        ? `<span class="cell-merge-badge">${cs > 1 ? `←${cs}→` : ""}${cs > 1 && rs > 1 ? " " : ""}${rs > 1 ? `↕${rs}` : ""}</span>`
        : "";
      // 정렬 아이콘
      const alignIcon = normalized.align
        ? `<span class="cell-align-badge">${normalized.align === "left" ? "◀" : normalized.align === "center" ? "●" : normalized.align === "right" ? "▶" : "≡"}</span>`
        : "";

      const isHeaderRow = rI === 0;
      const headerClass = isHeaderRow ? " editor-header-cell" : "";
      // 헤더 행 기본 정렬: 명시된 align 없으면 center
      const inputStyle = normalized.align
        ? `text-align:${normalized.align}`
        : isHeaderRow ? "text-align:center" : "";

      // 셀 이미지 썸네일 + 삭제 버튼
      const cellImgList = (normalized.images || []).map((img, imgIdx) => {
        const wPx = img.cx > 0 ? Math.min(Math.round(img.cx / EMU_PER_PX), 160) : 100;
        return `<div class="cell-img-row">
          <img class="cell-img-thumb" src="${img.dataUrl}" style="max-width:${wPx}px" alt="" />
          <button class="cell-img-remove-btn"
            data-action="table-cell-remove-image"
            data-section-index="${sI}" data-item-index="${iI}"
            data-table-index="${tI}" data-row-index="${rI}" data-cell-index="${cI}"
            data-image-index="${imgIdx}" title="이미지 삭제">✕</button>
        </div>`;
      }).join("");

      return `<td${csAttr}${rsAttr}
        class="editor-cell-wrap-td${selClass}${alignClass}${headerClass}"
        data-action="select-table-cell"
        data-section-index="${sI}" data-item-index="${iI}"
        data-table-index="${tI}" data-row-index="${rI}" data-cell-index="${cI}"
      >${mergeBadge}${alignIcon}<textarea
        data-action="table-cell"
        data-section-index="${sI}" data-item-index="${iI}"
        data-table-index="${tI}" data-row-index="${rI}" data-cell-index="${cI}"
        placeholder="셀 내용"
        style="${inputStyle}"
      >${escapeHtml(normalized.text)}</textarea>${cellImgList}<button class="cell-img-add-btn"
        data-action="table-cell-add-image"
        data-section-index="${sI}" data-item-index="${iI}"
        data-table-index="${tI}" data-row-index="${rI}" data-cell-index="${cI}"
        title="셀에 이미지 추가">＋이미지</button></td>`;
    }).join("");

    return `<tr>${rowSelectorHtml}${cellsHtml}</tr>`;
  }).join("");

  // 열 헤더 행 (A, B, C ... 클릭으로 열 전체 선택)
  const totalCols = getTableColumnCount(table);
  const colHeaderCells = Array.from({ length: totalCols }, (_, cI) => {
    const letter = cI < 26
      ? String.fromCharCode(65 + cI)
      : String.fromCharCode(65 + Math.floor(cI / 26) - 1) + String.fromCharCode(65 + (cI % 26));
    const isColSelected = isThisTable && [...(sel?.cells || [])].some((k) => {
      const [, c] = k.split(",").map(Number);
      return c === cI;
    });
    return `<th class="col-selector-col${isColSelected ? " col-selected" : ""}"
      data-action="select-table-col"
      data-section-index="${sI}" data-item-index="${iI}"
      data-table-index="${tI}" data-col-index="${cI}"
      title="${letter}열 전체 선택 (Ctrl+클릭 다중 선택)"
    >${letter}</th>`;
  }).join("");
  const colHeaderRow = `<tr><th class="col-selector-corner"></th>${colHeaderCells}</tr>`;

  return `<div class="editor-table-card">
    <div class="editor-table-head">
      <span class="table-label">표 ${tI + 1}</span>
      <div class="table-head-btns">
        <button class="icon-button" data-action="add-table-row"
          data-section-index="${sI}" data-item-index="${iI}" data-table-index="${tI}">+ 행</button>
        <button class="icon-button" data-action="add-table-col"
          data-section-index="${sI}" data-item-index="${iI}" data-table-index="${tI}">+ 열 (끝)</button>
        <button class="icon-button" data-action="remove-table"
          data-section-index="${sI}" data-item-index="${iI}" data-table-index="${tI}">표 삭제</button>
      </div>
    </div>
    ${toolbarHtml}<div class="editor-table-wrap">
      <table class="editor-table">
        <thead>${colHeaderRow}</thead>
        <tbody>${rowsHtml}</tbody>
      </table>
    </div>
  </div>`;
}

function renderOrgTabBar() {
  if (!els.orgTabBar) {
    return;
  }
  if (!state.subDocuments || state.subDocuments.length <= 1) {
    els.orgTabBar.innerHTML = "";
    return;
  }
  const tabsHtml = state.subDocuments
    .map(
      (doc, i) => `
        <div class="org-tab-group">
          <button
            class="org-tab ${i === state.activeSubDocIndex ? "active" : ""}"
            data-action="switch-org-tab"
            data-sub-doc-index="${i}"
            title="${escapeHtml(doc.org || `조직 ${i + 1}`)}"
          >${escapeHtml(doc.org || `조직 ${i + 1}`)}</button>
          <div class="org-tab-move-btns">
            <button class="org-tab-move-btn" data-action="org-move-left" data-sub-doc-index="${i}"
              ${i === 0 ? "disabled" : ""} title="조직 왼쪽으로">◀</button>
            <button class="org-tab-move-btn" data-action="org-move-right" data-sub-doc-index="${i}"
              ${i === state.subDocuments.length - 1 ? "disabled" : ""} title="조직 오른쪽으로">▶</button>
          </div>
        </div>
      `
    )
    .join("");
  els.orgTabBar.innerHTML = `
    <div class="org-tab-bar-header">
      <span class="org-tab-bar-label">조직 선택</span>
      <span class="org-tab-bar-count">${state.subDocuments.length}개 조직 · 선택 조직의 개요·본문이 아래에 표시됩니다</span>
    </div>
    <div class="org-tab-list">${tabsHtml}</div>
  `;
}

function clearUploadedDocument() {
  state.document = cloneValue(emptyDocumentState);
  state.subDocuments = null;
  state.activeSubDocIndex = 0;
  state.activeSelection = { type: "overview", sectionIndex: null };
  if (els.docxInput) {
    els.docxInput.value = "";
  }
  render();
}

function removeActiveSection() {
  const sectionIndex = getEditingSectionIndex();
  if (sectionIndex < 0) {
    return;
  }
  state.document.sections.splice(sectionIndex, 1);
  if (state.document.sections.length) {
    state.activeSelection = { type: "section", sectionIndex: Math.max(0, sectionIndex - 1) };
  } else {
    state.activeSelection = { type: "overview", sectionIndex: null };
  }
  render();
}

async function parseDocx(arrayBuffer) {
  const xmlParsed = await parseDocxXml(arrayBuffer).catch(() => null);
  if (xmlParsed?._multiDoc) {
    return xmlParsed;
  }
  if (xmlParsed?.sections?.length) {
    return xmlParsed;
  }

  const styleMap = [
    "p[style-name='타이틀'] => h1.docx-title:fresh",
    "p[style-name='기본 작성 정보'] => p.docx-meta:fresh",
    "p[style-name='카테고리'] => p.docx-category:fresh",
    "p[style-name='과제'] => p.docx-item-title:fresh",
    "p[style-name='내용_'] => p.docx-detail:fresh",
    "p[style-name='내용'] => p.docx-detail:fresh"
  ];

  const result = await mammoth.convertToHtml(
    { arrayBuffer },
    { styleMap, includeDefaultStyleMap: true }
  );

  const wrapper = document.createElement("div");
  wrapper.innerHTML = result.value;
  const parsed = parseStyledHtml(wrapper);
  if (parsed.sections.length) {
    return parsed;
  }

  const raw = await mammoth.extractRawText({ arrayBuffer });
  const plainParsed = parsePlainText(raw.value);
  if (plainParsed.sections.length || plainParsed.title || plainParsed.date || plainParsed.org) {
    return plainParsed;
  }

  return buildFallbackDocumentFromRawText(raw.value);
}

async function parseDocxXml(arrayBuffer) {
  if (!window.JSZip) {
    return null;
  }

  const zip = await window.JSZip.loadAsync(arrayBuffer);
  const documentXml = await zip.file("word/document.xml")?.async("string");
  if (!documentXml) {
    return null;
  }

  // 이미지 관계 파일 파싱
  const relsXml = await zip.file("word/_rels/document.xml.rels")?.async("string");
  const relsMap = parseWordRels(relsXml);

  // 번호 매기기 정보 파싱 (숫자 목록 복원용)
  const numberingXml = await zip.file("word/numbering.xml")?.async("string");
  const numberingData = parseWordNumbering(numberingXml);
  // 이미지 파일을 base64로 미리 로드
  const imageCache = {};
  for (const [rId, target] of Object.entries(relsMap)) {
    const normalizedTarget = target.replace(/^\//, "");
    const mediaPath = normalizedTarget.startsWith("word/") ? normalizedTarget : `word/${normalizedTarget}`;
    const imgFile = zip.file(mediaPath);
    if (imgFile) {
      const mime = getImageMimeType(mediaPath);
      if (mime) {
        const base64 = await imgFile.async("base64");
        imageCache[rId] = `data:${mime};base64,${base64}`;
      }
    }
  }

  const parser = new DOMParser();
  const xml = parser.parseFromString(documentXml, "application/xml");
  const body = xml.getElementsByTagNameNS(WORD_NS, "body")[0];
  if (!body) {
    return null;
  }

  const allParsed = [{ title: "mySUNI Weekly", date: "", org: "", sections: [] }];
  let parsed = allParsed[0];
  let currentSection = null;
  let currentItem = null;

  for (const node of [...body.childNodes]) {
    if (node.nodeType !== ELEMENT_NODE) {
      continue;
    }

    if (node.localName === "p") {
      const styleId = getWordParagraphStyle(node);
      const paragraphInfo = getWordParagraphInfo(node);
      // pStyle이 없는 단락 → numbering 체인으로 스타일 추론 (예: SKMS실천 과제 단락)
      const effectiveStyleId = styleId ||
        (paragraphInfo.hasNumbering && paragraphInfo.numId > 0
          ? (getStyleFromNumbering(numberingData, paragraphInfo.numId, paragraphInfo.ilvl) || "")
          : "");
      const text = cleanMultilineText(extractWordNodeText(node));

      // 단락 내 이미지(drawing) 추출 — cx/cy/align/distT/distB(EMU) 포함
      const drawingInfos = extractDrawingInfos(node);
      const drawingImages = drawingInfos
        .filter((info) => imageCache[info.rId])
        .map((info) => {
          // inline 이미지는 anchor positionH가 없으므로, 단락의 <w:jc>를 수평 정렬로 사용
          const pPrDraw = [...node.childNodes].find(
            (n) => n.nodeType === ELEMENT_NODE && n.localName === "pPr"
          );
          const jcDraw = pPrDraw
            ? [...pPrDraw.childNodes].find((n) => n.nodeType === ELEMENT_NODE && n.localName === "jc")
            : null;
          const paraAlign = jcDraw
            ? (jcDraw.getAttributeNS(WORD_NS, "val") || jcDraw.getAttribute("w:val") || null)
            : null;
          return {
            dataUrl: imageCache[info.rId],
            cx: info.cx,
            cy: info.cy,
            align: info.align || paraAlign || null,
            distT: info.distT || 0,
            distB: info.distB || 0,
          };
        });
      if (drawingImages.length) {
        if (!currentItem) {
          if (!currentSection) {
            currentSection = createEmptySection("[Other]");
            parsed.sections.push(currentSection);
          }
          currentItem = { id: createId(), title: "", details: [], tables: [], images: [] };
          currentSection.items.push(currentItem);
        }
        if (!currentItem.images) currentItem.images = [];
        currentItem.images.push(...drawingImages);
      }

      if (!text) {
        continue;
      }

      if (effectiveStyleId === "ad" || /^mysuni weekly$/i.test(text)) {
        // 이미 내용이 있으면 → 새 조직 문서 시작
        if (parsed.sections.length > 0 || parsed.date || parsed.org) {
          parsed = { title: text, date: "", org: "", sections: [] };
          allParsed.push(parsed);
          currentSection = null;
          currentItem = null;
        } else {
          parsed.title = text;
        }
        continue;
      }

      if (effectiveStyleId === "af" || /^(일자|date|부서|department|조직|organization)\s*:/.test(text)) {
        const lines = text.split("\n");
        let newDate = null;
        let newOrg = null;
        lines.forEach((line) => {
          const info = parseMetaLine(line);
          if (info.date) newDate = info.date;
          if (info.org) newOrg = info.org;
        });
        // 이미 일자/부서가 있고 본문 내용도 있으면 → 새 조직 문서 시작
        if ((newDate || newOrg) && (parsed.date || parsed.org) && parsed.sections.length > 0) {
          parsed = { title: "mySUNI Weekly", date: newDate || "", org: newOrg || "", sections: [] };
          allParsed.push(parsed);
          currentSection = null;
          currentItem = null;
        } else {
          if (newDate) parsed.date = newDate;
          if (newOrg) parsed.org = newOrg;
        }
        continue;
      }

      // 카테고리: 반드시 [대괄호] 형식이어야 함 — () 등 다른 괄호는 카테고리로 처리하지 않음
      if (/^\[[^\]]+\]$/.test(text)) {
        if (shouldAppendTitleLikeParagraphAsDetail(currentItem, paragraphInfo, text)) {
          appendDetailLine(currentItem, text);
          continue;
        }
        currentSection = {
          id: createId(),
          category: text,
          items: []
        };
        parsed.sections.push(currentSection);
        currentItem = null;
        continue;
      }

      if (effectiveStyleId === "a") {
        if (shouldTreatAAsDetail(currentItem, paragraphInfo, text)) {
          appendParsedDetailLine(currentItem, paragraphInfo, text);
          continue;
        }
        if (!currentSection) {
          currentSection = createEmptySection("[Other]");
          parsed.sections.push(currentSection);
        }
        currentItem = {
          id: createId(),
          title: text,
          details: [],
          tables: []
        };
        currentSection.items.push(currentItem);
        continue;
      }

      if (effectiveStyleId === "a9" && shouldTreatA9AsItemTitle(node, text)) {
        if (!currentSection) {
          currentSection = createEmptySection("[Other]");
          parsed.sections.push(currentSection);
        }
        currentItem = {
          id: createId(),
          title: text,
          details: [],
          tables: []
        };
        currentSection.items.push(currentItem);
        continue;
      }

      if (effectiveStyleId === "a1" || effectiveStyleId === "a9") {
        if (!currentItem) {
          if (!currentSection) {
            currentSection = createEmptySection("[Other]");
            parsed.sections.push(currentSection);
          }
          currentItem = createEmptyItem("");
          currentSection.items.push(currentItem);
        }
        // 원문 텍스트를 그대로 보존 — 번호/불릿 생성 없이 파싱
        appendParsedDetailLine(currentItem, paragraphInfo, text);
      }
    }

    if (node.localName === "tbl") {
      if (!currentItem) {
        if (!currentSection) {
          currentSection = createEmptySection("[Other]");
          parsed.sections.push(currentSection);
        }
        currentItem = createEmptyItem("");
        currentSection.items.push(currentItem);
      }

      // 표 파싱 — imageCache 전달하여 셀별 이미지를 cell.images로 저장
      currentItem.tables.push(parseWordXmlTable(node, imageCache));
    }
  }

  const finalizedDocs = allParsed.map(finalizeParsedDocument);
  if (finalizedDocs.length > 1) {
    return { _multiDoc: true, documents: finalizedDocs };
  }
  return finalizedDocs[0];
}

function parseStyledHtml(wrapper) {
  const parsed = { title: "mySUNI Weekly", date: "", org: "", sections: [] };
  let currentSection = null;
  let currentItem = null;

  const children = [...wrapper.children];
  for (const node of children) {
    const text = cleanText(node.textContent);
    const multilineText = extractNodeTextWithBreaks(node);
    if (!text) {
      continue;
    }

    if (node.matches(".docx-title")) {
      parsed.title = text;
      continue;
    }

    if (node.matches(".docx-meta")) {
      const info = parseMetaLine(text);
      if (info.date) parsed.date = info.date;
      if (info.org) parsed.org = info.org;
      continue;
    }

    if (node.matches(".docx-category")) {
      // 카테고리는 반드시 [대괄호] 형식이어야 함 — () 등은 과제 제목으로 처리
      if (/^\[[^\]]+\]$/.test(text)) {
        currentSection = {
          id: createId(),
          category: text,
          items: []
        };
        parsed.sections.push(currentSection);
        currentItem = null;
        continue;
      }
      // [대괄호] 아닌 경우 → 과제 제목으로 fall-through
    }

    if (node.matches(".docx-item-title") || (node.matches(".docx-category") && !/^\[[^\]]+\]$/.test(text))) {
      if (!currentSection) {
        currentSection = createEmptySection("[Other]");
        parsed.sections.push(currentSection);
      }
      currentItem = {
        id: createId(),
        title: text,
        details: [],
        tables: []
      };
      currentSection.items.push(currentItem);
      continue;
    }

    if (node.matches(".docx-detail")) {
      if (!currentItem) {
        if (!currentSection) {
          currentSection = createEmptySection("[Other]");
          parsed.sections.push(currentSection);
        }
        currentItem = createEmptyItem("");
        currentSection.items.push(currentItem);
      }
      currentItem.details.push(markDetailBullet(multilineText || text));
      continue;
    }

    if (node.tagName === "TABLE") {
      if (!currentItem) {
        if (!currentSection) {
          currentSection = createEmptySection("[Other]");
          parsed.sections.push(currentSection);
        }
        currentItem = createEmptyItem("");
        currentSection.items.push(currentItem);
      }
      currentItem.tables.push(parseHtmlTable(node));
      continue;
    }
  }

  return finalizeParsedDocument(parsed);
}

function parsePlainText(text) {
  const parsed = {
    title: "mySUNI Weekly",
    date: "",
    org: "",
    sections: []
  };

  let currentSection = null;
  let currentItem = null;

  const lines = text
    .split(/\r?\n/)
    .map((line) => cleanText(line))
    .filter((line) => line);

  for (const line of lines) {
    if (/^mysuni weekly$/i.test(line)) {
      parsed.title = line;
      continue;
    }

    const meta = parseMetaLine(line);
    if (meta.date || meta.org) {
      if (meta.date) parsed.date = meta.date;
      if (meta.org) parsed.org = meta.org;
      continue;
    }

    if (/^\[[^\]]+\]$/.test(line)) {
      currentSection = createEmptySection(line);
      parsed.sections.push(currentSection);
      currentItem = null;
      continue;
    }

    if (!currentSection) {
      currentSection = createEmptySection("[Other]");
      parsed.sections.push(currentSection);
    }

    if (isLikelyTitle(line, currentItem)) {
      currentItem = createEmptyItem(line);
      currentSection.items.push(currentItem);
      continue;
    }

    if (!currentItem) {
      currentItem = createEmptyItem("");
      currentSection.items.push(currentItem);
    }

    currentItem.details.push(markDetailBullet(line.replace(/^[•\-]\s*/, "")));
  }

  return finalizeParsedDocument(parsed);
}

function parseMetaLine(line) {
  const normalized = line.replace(/\s+/g, " ").trim();
  const info = {};

  const dateMatch = normalized.match(/^(일자|date)\s*:\s*(.+)$/i);
  if (dateMatch) {
    info.date = dateMatch[2].trim();
  }

  const orgMatch = normalized.match(/^(부서|department|조직|organization)\s*:\s*(.+)$/i);
  if (orgMatch) {
    info.org = orgMatch[2].trim();
  }

  return info;
}

function parseHtmlTable(table) {
  const rows = [...table.querySelectorAll("tr")].map((row) =>
    [...row.children].map((cell) => ({
      text: cleanText(cell.textContent),
      colSpan: Number.parseInt(cell.getAttribute("colspan") || "1", 10) || 1
    }))
  );
  return { rows };
}

function parseWordXmlTable(tableNode, imageCache = {}) {
  const trNodes = [...tableNode.childNodes].filter(
    (n) => n.nodeType === ELEMENT_NODE && n.localName === "tr"
  );
  if (!trNodes.length) {
    return { rows: [] };
  }

  // 1차: 각 행의 셀 원시 데이터 추출 (colspan, vMerge 포함)
  const rawRows = trNodes.map((tr) =>
    [...tr.childNodes]
      .filter((n) => n.nodeType === ELEMENT_NODE && n.localName === "tc")
      .map((cell) => {
        const cellText = [...cell.childNodes]
          .filter((n) => n.nodeType === ELEMENT_NODE && n.localName === "p")
          .map((p) => cleanMultilineText(extractWordNodeText(p)))
          .filter(Boolean)
          .join("\n");

        const tcPr = [...cell.childNodes].find(
          (n) => n.nodeType === ELEMENT_NODE && n.localName === "tcPr"
        );

        // colspan (w:gridSpan)
        const spanNode = tcPr
          ? [...tcPr.childNodes].find((n) => n.nodeType === ELEMENT_NODE && n.localName === "gridSpan")
          : null;
        const colSpan =
          Number.parseInt(
            spanNode?.getAttributeNS(WORD_NS, "val") || spanNode?.getAttribute("w:val") || "1",
            10
          ) || 1;

        // rowspan (w:vMerge)
        const vMergeNode = tcPr
          ? [...tcPr.childNodes].find((n) => n.nodeType === ELEMENT_NODE && n.localName === "vMerge")
          : null;
        let vMerge = null;
        if (vMergeNode) {
          const val =
            vMergeNode.getAttributeNS(WORD_NS, "val") || vMergeNode.getAttribute("w:val") || "";
          vMerge = val === "restart" ? "restart" : "continue";
        }

        // 수평 정렬 (w:jc) — 셀 내 첫 단락에서 추출
        let align = null;
        const cellParas = [...cell.childNodes].filter(
          (n) => n.nodeType === ELEMENT_NODE && n.localName === "p"
        );
        for (const para of cellParas) {
          const pPrNode = [...para.childNodes].find(
            (n) => n.nodeType === ELEMENT_NODE && n.localName === "pPr"
          );
          const jcNode = pPrNode
            ? [...pPrNode.childNodes].find((n) => n.nodeType === ELEMENT_NODE && n.localName === "jc")
            : null;
          if (jcNode) {
            const val = jcNode.getAttributeNS(WORD_NS, "val") || jcNode.getAttribute("w:val") || "";
            if (val) {
              align = val === "both" ? "justify" : val;
              break;
            }
          }
        }

        // 셀 내 이미지 추출 (cx/cy EMU 포함)
        const cellDrawingInfos = extractDrawingInfos(cell);
        const cellImages = cellDrawingInfos
          .filter((info) => imageCache[info.rId])
          .map((info) => ({ dataUrl: imageCache[info.rId], cx: info.cx, cy: info.cy }));

        return { text: cellText, colSpan, vMerge, align, images: cellImages };
      })
  );

  // 각 행에서 셀이 차지하는 그리드 열 위치 계산
  const rowGridPositions = rawRows.map((row) => {
    const positions = [];
    let col = 0;
    for (const cell of row) {
      positions.push(col);
      col += cell.colSpan;
    }
    return positions;
  });

  // 특정 그리드 열에 있는 셀 조회
  const getCellAtGridCol = (rowIdx, targetCol) => {
    const row = rawRows[rowIdx];
    const positions = rowGridPositions[rowIdx];
    for (let i = 0; i < row.length; i++) {
      if (positions[i] === targetCol) return row[i];
      if (positions[i] > targetCol) return null;
    }
    return null;
  };

  // 2차: rowspan 계산 및 최종 데이터 구조 생성
  // rowSpan=0 → 위 셀에 의해 병합된 숨김 셀
  const computedRows = rawRows.map((row, rowIdx) => {
    const positions = rowGridPositions[rowIdx];
    return row.map((cell, cellIdx) => {
      const gridCol = positions[cellIdx];
      if (cell.vMerge === "continue") {
        return { text: "", colSpan: cell.colSpan, rowSpan: 0, align: cell.align || null, images: [] };
      }
      if (cell.vMerge === "restart") {
        let rowSpan = 1;
        for (let r = rowIdx + 1; r < rawRows.length; r++) {
          const next = getCellAtGridCol(r, gridCol);
          if (next && next.vMerge === "continue") {
            rowSpan++;
          } else {
            break;
          }
        }
        return { text: cell.text, colSpan: cell.colSpan, rowSpan, align: cell.align || null, images: cell.images || [] };
      }
      return { text: cell.text, colSpan: cell.colSpan, rowSpan: 1, align: cell.align || null, images: cell.images || [] };
    });
  }).filter((row) => row.length);

  return { rows: computedRows };
}

function getWordParagraphStyle(paragraphNode) {
  const pPr = [...paragraphNode.childNodes].find((node) => node.nodeType === ELEMENT_NODE && node.localName === "pPr");
  if (!pPr) {
    return "";
  }
  const pStyle = [...pPr.childNodes].find((node) => node.nodeType === ELEMENT_NODE && node.localName === "pStyle");
  return pStyle?.getAttributeNS(WORD_NS, "val") || pStyle?.getAttribute("w:val") || "";
}

function getWordParagraphInfo(paragraphNode) {
  const pPr = [...paragraphNode.childNodes].find((node) => node.nodeType === ELEMENT_NODE && node.localName === "pPr");
  if (!pPr) {
    return {
      styleId: "",
      hasNumbering: false,
      leftIndent: 0,
      hangingIndent: 0,
      firstLineIndent: 0
    };
  }

  const ind = [...pPr.childNodes].find((node) => node.nodeType === ELEMENT_NODE && node.localName === "ind");
  const numPr = [...pPr.childNodes].find((node) => node.nodeType === ELEMENT_NODE && node.localName === "numPr");

  let numId = 0;
  let ilvl = 0;
  if (numPr) {
    const numIdEl = [...numPr.childNodes].find((n) => n.nodeType === ELEMENT_NODE && n.localName === "numId");
    const ilvlEl = [...numPr.childNodes].find((n) => n.nodeType === ELEMENT_NODE && n.localName === "ilvl");
    numId = Number(numIdEl?.getAttributeNS(WORD_NS, "val") || numIdEl?.getAttribute("w:val") || "0");
    ilvl = Number(ilvlEl?.getAttributeNS(WORD_NS, "val") || ilvlEl?.getAttribute("w:val") || "0");
  }

  return {
    styleId: getWordParagraphStyle(paragraphNode),
    hasNumbering: Boolean(numPr),
    numId,
    ilvl,
    leftIndent: parseWordIndentValue(ind, "left"),
    hangingIndent: parseWordIndentValue(ind, "hanging"),
    firstLineIndent: parseWordIndentValue(ind, "firstLine")
  };
}

function parseWordIndentValue(indNode, key) {
  if (!indNode) {
    return 0;
  }
  const value = indNode.getAttributeNS(WORD_NS, key) || indNode.getAttribute(`w:${key}`) || "0";
  return Number.parseInt(value, 10) || 0;
}

function extractWordNodeText(node, _state) {
  // _state는 재귀 호출 시 공유되는 가변 상태 객체
  // inFieldResult: true일 때 <w:t> 내용을 스킵 (자동 번호/기호 캐시 결과 무시)
  const state = _state || { inFieldResult: false };
  let output = "";
  for (const child of [...node.childNodes]) {
    if (child.nodeType === Node.TEXT_NODE) {
      output += child.textContent || "";
      continue;
    }

    if (child.nodeType !== ELEMENT_NODE) {
      continue;
    }

    // fldChar: 필드 코드 상태 추적
    // separate ~ end 사이의 <w:t>는 자동 생성 번호/기호 캐시이므로 스킵
    if (child.localName === "fldChar") {
      const fldType = child.getAttributeNS(WORD_NS, "fldCharType") || child.getAttribute("w:fldCharType") || "";
      if (fldType === "separate") state.inFieldResult = true;
      else if (fldType === "end") state.inFieldResult = false;
      continue;
    }

    // instrText: 필드 명령어 텍스트 (예: " AUTONUM ", " SEQ ") → 화면에 표시 안되므로 스킵
    if (child.localName === "instrText" || child.localName === "delText") {
      continue;
    }

    if (child.localName === "t") {
      if (!state.inFieldResult) output += child.textContent || "";
      continue;
    }

    if (child.localName === "br") {
      output += "\n";
      continue;
    }

    if (child.localName === "sym") {
      output += mapWordSymbol(child.getAttribute("w:font") || "", child.getAttribute("w:char") || "");
      continue;
    }

    // 흰색(FFFFFF) 텍스트 런 스킵 — Word에서 배경과 동일한 색으로 숨겨진 텍스트(예: 보이지 않는 '1' 자리수)
    if (child.localName === "r") {
      const rPr = [...child.childNodes].find((n) => n.nodeType === ELEMENT_NODE && n.localName === "rPr");
      if (rPr) {
        const colorEl = [...rPr.childNodes].find((n) => n.nodeType === ELEMENT_NODE && n.localName === "color");
        if (colorEl) {
          const colorVal = (colorEl.getAttributeNS(WORD_NS, "val") || colorEl.getAttribute("w:val") || "").toUpperCase();
          if (colorVal === "FFFFFF") continue;
        }
      }
    }

    output += extractWordNodeText(child, state);
  }
  return output;
}

function mapWordSymbol(font, charCode) {
  const key = `${String(font || "").toLowerCase()}:${String(charCode || "").toUpperCase()}`;
  const symbolMap = {
    "wingdings:F0E8": "→",
    "wingdings:F0DF": "↑",
    "wingdings:F0E0": "←",
    "wingdings:F0E1": "↓"
  };
  return symbolMap[key] || "";
}

function shouldAppendTitleLikeParagraphAsDetail(currentItem, paragraphInfo, text) {
  if (!currentItem) {
    return false;
  }

  if (!currentItem.details.length && !currentItem.tables.length) {
    return false;
  }

  if (!paragraphInfo || paragraphInfo.styleId !== "a0") {
    return false;
  }

  if (text.startsWith("[")) {
    return false;
  }

  if (text.startsWith("*") || text.startsWith(":") || text.startsWith("→")) {
    return true;
  }

  if (paragraphInfo.hasNumbering || paragraphInfo.leftIndent >= 900 || paragraphInfo.firstLineIndent > 0) {
    return true;
  }

  const lastDetail = currentItem.details[currentItem.details.length - 1] || "";
  if (/\*$/.test(lastDetail) || /[→:]\s*$/.test(lastDetail)) {
    return true;
  }

  return false;
}

function shouldTreatAAsDetail(currentItem, paragraphInfo, text) {
  if (!currentItem) {
    return false;
  }

  if (!paragraphInfo || paragraphInfo.styleId !== "a") {
    return false;
  }

  if (paragraphInfo.hangingIndent > 0) {
    return false;
  }

  if (paragraphInfo.leftIndent < 900) {
    return false;
  }

  return Boolean(text);
}

function shouldTreatA9AsItemTitle(paragraphNode, text) {
  if (!text) {
    return false;
  }

  let sibling = paragraphNode.nextSibling;
  while (sibling) {
    if (sibling.nodeType !== ELEMENT_NODE) {
      sibling = sibling.nextSibling;
      continue;
    }

    if (sibling.localName === "tbl") {
      return true;
    }

    if (sibling.localName !== "p") {
      sibling = sibling.nextSibling;
      continue;
    }

    const siblingText = cleanMultilineText(extractWordNodeText(sibling));
    if (!siblingText) {
      sibling = sibling.nextSibling;
      continue;
    }

    const siblingStyle = getWordParagraphStyle(sibling);
    if (siblingStyle === "a1") {
      sibling = sibling.nextSibling;
      continue;
    }

    return false;
  }

  return false;
}

function appendParsedDetailLine(item, paragraphInfo, text) {
  const normalizedText = normalizeParsedDetailText(paragraphInfo, text);
  if (!normalizedText) {
    return;
  }

  if (shouldAppendParsedDetailToPrevious(item, paragraphInfo, text)) {
    appendDetailLine(item, normalizedText);
    return;
  }

  item.details.push(markDetailBullet(normalizedText));
}

function shouldAppendParsedDetailToPrevious(item, paragraphInfo, text) {
  if (!item?.details?.length) {
    return false;
  }

  if (!paragraphInfo) {
    return false;
  }

  if (paragraphInfo.hasNumbering) {
    return true;
  }

  if (paragraphInfo.styleId === "a" && paragraphInfo.leftIndent >= 900 && !/^[*•\-→:]/.test(cleanMultilineText(text))) {
    return true;
  }

  return false;
}

function normalizeParsedDetailText(paragraphInfo, text) {
  let normalized = cleanMultilineText(text);
  if (!normalized) {
    return "";
  }

  // NOTE: hasNumbering 기반 숫자/기호 제거 로직을 삭제함.
  // OOXML에서 자동 생성된 번호는 텍스트 내용에 포함되지 않으므로
  // 제거 불필요. 사용자가 직접 입력한 "1. 항목" 등의 번호/기호를 보존.

  if (paragraphInfo?.styleId === "a" && paragraphInfo.leftIndent >= 900) {
    normalized = normalized.replace(/^\*\s*/, "");
  }

  return cleanMultilineText(normalized);
}

function finalizeParsedDocument(parsed) {
  const cleanSections = parsed.sections
    .map((section) => ({
      ...section,
      items: section.items
        .map((item) => ({
          ...item,
          details: item.details.filter(Boolean),
          tables: item.tables || [],
          images: item.images || []
        }))
        .filter((item) => item.title || item.details.length || item.tables.length || item.images?.length)
    }))
    .filter((section) => section.category || section.items.length);

  return {
    title: parsed.title || "",
    date: parsed.date || "",
    org: parsed.org || "",
    sections: cleanSections
  };
}

function buildFallbackDocumentFromRawText(text) {
  const lines = String(text || "")
    .split(/\r?\n/)
    .map((line) => cleanMultilineText(line))
    .filter(Boolean);

  if (!lines.length) {
    return cloneValue(emptyDocumentState);
  }

  const [firstLine, ...rest] = lines;
  const detailText = rest.join("\n");

  return {
    title: /^mysuni weekly$/i.test(firstLine) ? firstLine : "mySUNI Weekly",
    date: "",
    org: "",
    filename: "",
    sourceFileName: "",
    sections: [
      {
        id: createId(),
        category: "[기타]",
        items: [
          {
            id: createId(),
            title: firstLine,
            details: detailText ? [detailText] : [],
            tables: []
          }
        ]
      }
    ]
  };
}

function render() {
  scheduleDocumentDraftSave();
  renderCounts();
  renderTabs();
  renderEditorWidthMode();
  renderUploadStatus();
  renderOrgTabBar();
  renderSectionList();
  renderEditor();
  renderPreview();
  renderViewButtons();
  renderFilename();
  renderDraftStatus();
  renderIntegrationTab();
  renderDocumentArchiveTab();
  renderGuidePage();
  renderAdminPage();
}

function renderCounts() {
  if (els.glossaryCount) {
    els.glossaryCount.textContent = `용어 ${state.glossary.length}건`;
  }
  if (els.nameCount) {
    els.nameCount.textContent = `이름 ${state.names.length}건`;
  }
}

function renderTabs() {
  els.mainTabButton?.classList.toggle("active", state.activeTab === "main");
  els.integrationTabButton?.classList.toggle("active", state.activeTab === "integration");
  els.archiveTabButton?.classList.toggle("active", state.activeTab === "archive");
  els.translateTabButton?.classList.toggle("active", state.activeTab === "translate");
  els.guideTabButton?.classList.toggle("active", state.activeTab === "guide");
  els.adminTabButton?.classList.toggle("active", state.activeTab === "admin");
  els.mainTab?.classList.toggle("active", state.activeTab === "main");
  els.integrationTab?.classList.toggle("active", state.activeTab === "integration");
  els.archiveTab?.classList.toggle("active", state.activeTab === "archive");
  els.translateTab?.classList.toggle("active", state.activeTab === "translate");
  els.guideTab?.classList.toggle("active", state.activeTab === "guide");
  els.adminTab?.classList.toggle("active", state.activeTab === "admin");
}

function renderIntegrationTab() {
  if (!els.integrationDateInput || !els.integrationDatePicker || !els.integrationStatusList || !els.archiveList) {
    return;
  }

  els.integrationDateInput.value = state.integrationDate || "";
  els.integrationDatePicker.value = extractDateValue(state.integrationDate);
  els.integrationStatusList.innerHTML = state.integrationSlots
    .map(
      (slot, index) => `
        <div class="integration-card ${slot.submitted ? "submitted" : ""}">
          <div class="integration-card-header">
            <p class="integration-card-title">${escapeHtml(slot.org)}</p>
            <span class="integration-badge ${slot.submitted ? "done" : ""}">${slot.submitted ? "제출" : "미제출"}</span>
          </div>
          <label class="field">
            <span>문서 업로드</span>
            <input type="file" accept=".docx" data-action="integration-upload" data-slot-index="${index}" />
          </label>
          <p class="hint">${slot.submitted ? `${escapeHtml(slot.document?.title || "문서")} / ${escapeHtml(slot.document?.org || slot.org)}` : "아직 업로드되지 않았습니다."}</p>
        </div>
      `
    )
    .join("");

  els.archiveList.innerHTML = state.archives.length
    ? state.archives
        .map(
          (archive, index) => `
            <div class="archive-card">
              <div>
                <strong>${escapeHtml(archive.date)}</strong>
                <div class="hint">${archive.submittedCount}/${INTEGRATION_ORGS.length} 제출</div>
              </div>
              <button class="ghost-button" data-action="restore-archive" data-archive-index="${index}">불러오기</button>
            </div>
          `
        )
        .join("")
    : '<div class="editor-empty">저장된 아카이브가 없습니다.</div>';
}

function renderDocumentArchiveTab() {
  if (!els.documentArchiveList) {
    return;
  }

  els.documentArchiveList.innerHTML = state.documentArchives.length
    ? state.documentArchives
        .map(
          (archive, index) => `
            <div class="archive-card">
              <div>
                <strong>${escapeHtml(archive.title || "제목")}</strong>
                <div class="hint">${escapeHtml(archive.date || "-")} / ${escapeHtml(archive.org || "조직명")} / ${archive.sectionCount || 0}개 카테고리</div>
                <div class="hint">${escapeHtml(archive.savedAtLabel || "")}</div>
              </div>
              <div class="button-row">
                <button class="ghost-button" data-action="restore-document-archive" data-archive-index="${index}">불러오기</button>
                <button class="ghost-button" data-action="delete-document-archive" data-archive-index="${index}">삭제</button>
              </div>
            </div>
          `
        )
        .join("")
    : '<div class="editor-empty">저장된 작성 아카이브가 없습니다.</div>';
}

/** 저장 기록 드롭다운 패널 렌더링 */
function renderArchiveSelectPanel() {
  const panel = els.archiveSelectPanel;
  if (!panel) return;

  const archives = state.documentArchives;
  if (!archives.length) {
    panel.innerHTML = `<div class="archive-select-empty">저장된 문서가 없습니다.<br><span style="font-size:10px;">작성 탭에서 '저장' 버튼을 눌러 기록을 만들어 보세요.</span></div>`;
    return;
  }

  panel.innerHTML =
    `<div class="archive-select-header">저장된 문서 (${archives.length}개)</div>` +
    archives.map((a, i) => `
      <div class="archive-select-item" data-action="load-from-archive" data-archive-index="${i}">
        <div class="archive-select-item-title">${escapeHtml(a.title || "(제목 없음)")}</div>
        <div class="archive-select-item-meta">${escapeHtml(a.date || "날짜 없음")} · ${escapeHtml(a.org || "조직명 없음")} · ${a.sectionCount || 0}개 카테고리 · ${escapeHtml(a.savedAtLabel || "")}</div>
      </div>
    `).join("");
}

/** 저장 기록 드롭다운 항목 클릭 처리 */
function onArchiveSelectPanelClick(event) {
  const item = event.target.closest("[data-action='load-from-archive']");
  if (!item) return;

  const archive = state.documentArchives[Number(item.dataset.archiveIndex)];
  if (!archive) return;

  state.document = cloneValue(archive.document);
  state.activeSelection = state.document.sections.length
    ? { type: "section", sectionIndex: 0 }
    : { type: "overview", sectionIndex: null };
  clearAiTranslationCache();
  syncInputsFromState();
  render();

  // 패널 닫기
  if (els.archiveSelectPanel) els.archiveSelectPanel.hidden = true;
}

function syncIntegrationDatePicker() {
  if (!els.integrationDatePicker) {
    return;
  }
  els.integrationDatePicker.value = extractDateValue(state.integrationDate);
}

function renderSectionList() {
  if (!els.sectionSelect || !els.removeSectionButton) {
    return;
  }

  const selectedIndex = getEditingSectionIndex();
  els.sectionSelect.innerHTML = state.document.sections.length
    ? state.document.sections
        .map(
          (section, sectionIndex) =>
            `<option value="${sectionIndex}" ${sectionIndex === selectedIndex ? "selected" : ""}>${escapeHtml(section.category)}</option>`
        )
        .join("")
    : '<option value="">카테고리 없음</option>';
  els.sectionSelect.disabled = !state.document.sections.length;
  els.removeSectionButton.disabled = !state.document.sections.length;
  if (els.moveSectionUpButton) {
    els.moveSectionUpButton.disabled = selectedIndex <= 0;
  }
  if (els.moveSectionDownButton) {
    els.moveSectionDownButton.disabled = selectedIndex < 0 || selectedIndex >= state.document.sections.length - 1;
  }
  if (els.sectionItemCount) {
    const currentSection = state.document.sections[selectedIndex];
    els.sectionItemCount.textContent = currentSection ? `${currentSection.items.length}개` : "0개";
  }
}

function autoResizeSingleTextarea(ta) {
  ta.style.height = "auto";
  ta.style.height = ta.scrollHeight + "px";
}

function autoResizeTableTextareas() {
  if (!els.editorRoot) return;
  els.editorRoot.querySelectorAll('textarea[data-action="table-cell"]').forEach(autoResizeSingleTextarea);
}

function renderEditor() {
  if (!els.editorRoot) {
    return;
  }
  const sectionIndex = getEditingSectionIndex();
  const section = state.document.sections[sectionIndex];

  if (!section) {
    els.editorRoot.innerHTML = `<div class="editor-empty">카테고리를 추가한 뒤 본문을 편집해 주세요.</div>`;
    return;
  }

  const itemsHtml = section.items
    .map((item, itemIndex) => {
      const detailHtml = item.details
        .map(
          (detail, detailIndex) => `
            <div class="detail-row">
              <textarea
                rows="3"
                data-action="detail-input"
                data-detail-bullet="${hasDetailBullet(detail) ? "true" : "false"}"
                data-section-index="${sectionIndex}"
                data-item-index="${itemIndex}"
                data-detail-index="${detailIndex}"
                placeholder="내용 입력 (Shift+Enter 줄바꿈 유지)"
              >${escapeHtml(stripDetailBullet(detail))}</textarea>
              <div class="detail-row-actions">
                <button class="icon-button move-button" data-action="detail-move-up"
                  data-section-index="${sectionIndex}" data-item-index="${itemIndex}" data-detail-index="${detailIndex}"
                  ${detailIndex === 0 ? "disabled" : ""} title="내용 위로">▲</button>
                <button class="icon-button move-button" data-action="detail-move-down"
                  data-section-index="${sectionIndex}" data-item-index="${itemIndex}" data-detail-index="${detailIndex}"
                  ${detailIndex === item.details.length - 1 ? "disabled" : ""} title="내용 아래로">▼</button>
                <button class="icon-button" data-action="remove-detail"
                  data-section-index="${sectionIndex}" data-item-index="${itemIndex}" data-detail-index="${detailIndex}">삭제</button>
              </div>
            </div>
          `
        )
        .join("");

      const tablesHtml = item.tables
        .map((table, tableIndex) => renderEditorTable(table, sectionIndex, itemIndex, tableIndex))
        .join("");

      const imagesHtml = (item.images || []).map((image, imageIndex) => `
        <div class="editor-image-row">
          <img class="editor-image-thumb" src="${image.dataUrl}" alt="삽입 이미지" />
          <button class="icon-button icon-button-danger" data-action="remove-image"
            data-section-index="${sectionIndex}" data-item-index="${itemIndex}"
            data-image-index="${imageIndex}">✕ 삭제</button>
        </div>
      `).join("");

      return `
        <div class="editor-item">
          <div class="item-header">
            <input
              class="item-title-input"
              value="${escapeHtml(item.title)}"
              data-action="item-title"
              data-section-index="${sectionIndex}"
              data-item-index="${itemIndex}"
            />
            <button class="icon-button move-button" data-action="item-move-up"
              data-section-index="${sectionIndex}" data-item-index="${itemIndex}"
              ${itemIndex === 0 ? "disabled" : ""} title="과제 위로">▲</button>
            <button class="icon-button move-button" data-action="item-move-down"
              data-section-index="${sectionIndex}" data-item-index="${itemIndex}"
              ${itemIndex === section.items.length - 1 ? "disabled" : ""} title="과제 아래로">▼</button>
            <button class="icon-button" data-action="add-detail" data-section-index="${sectionIndex}" data-item-index="${itemIndex}">내용 추가</button>
            <button class="icon-button" data-action="add-table" data-section-index="${sectionIndex}" data-item-index="${itemIndex}">표 추가</button>
            <button class="icon-button" data-action="add-image" data-section-index="${sectionIndex}" data-item-index="${itemIndex}">🖼 이미지</button>
            <button class="icon-button" data-action="remove-item" data-section-index="${sectionIndex}" data-item-index="${itemIndex}">과제 삭제</button>
          </div>
          <div class="detail-list">${detailHtml || '<p class="hint">아직 내용이 없습니다.</p>'}</div>
          <div class="table-list">${tablesHtml}</div>
          ${imagesHtml ? `<div class="image-list">${imagesHtml}</div>` : ""}
        </div>
      `;
    })
    .join("");

  els.editorRoot.innerHTML = `
    <div class="editor-section">
      <div class="section-header">
        <input
          class="section-title-input"
          value="${escapeHtml(section.category)}"
          data-action="section-title"
          data-section-index="${sectionIndex}"
        />
        <button class="icon-button" data-action="add-item" data-section-index="${sectionIndex}">과제 추가</button>
      </div>
      ${itemsHtml}
    </div>
  `;
  // 표 셀 textarea 높이 자동 맞춤 (innerHTML 교체 후)
  autoResizeTableTextareas();
}

function renderUploadStatus() {
  if (!els.uploadedFileName) {
    return;
  }
  els.uploadedFileName.textContent = state.document.sourceFileName || "현재 업로드된 파일 없음";
  els.clearDocxButton.disabled = !state.document.sourceFileName;
}

function getEditingSectionIndex() {
  if (state.activeSelection.type === "section" && Number.isInteger(state.activeSelection.sectionIndex)) {
    return Math.max(0, Math.min(state.activeSelection.sectionIndex, state.document.sections.length - 1));
  }
  return state.document.sections.length ? 0 : -1;
}

function renderDocumentSections(doc, opts = {}) {
  return doc.sections
    .map(
      (section, sectionIndex) => `
        <section class="doc-section ${!opts.secondary && state.activeSelection.type === "section" && state.activeSelection.sectionIndex === sectionIndex ? "is-selected" : ""}" data-section-index="${sectionIndex}">
          ${!opts.secondary ? `<span class="doc-section-actions">
            <button class="doc-section-act-btn" data-action="copy-section" data-section-index="${sectionIndex}" title="표준 양식 스타일로 클립보드 복사 → 동일 템플릿 기반 Word 문서에 바로 붙여넣기 가능">📋 복사</button>
            <button class="doc-section-act-btn" data-action="download-section" data-section-index="${sectionIndex}" title="표준 양식 .docx 파일로 저장">⬇ Word 저장</button>
          </span>` : ""}
          <p class="doc-category">${escapeHtml(section.category)}</p>
          ${section.items
            .map(
              (item) => `
                <article class="doc-item">
                  ${item.title ? `<p class="doc-item-title">${escapeHtml(item.title)}</p>` : ""}
                  ${item.details.map(renderDetail).join("")}
                  ${item.images?.map(renderDocImage).join("") || ""}
                  ${item.tables.map(renderTable).join("")}
                </article>
              `
            )
            .join("")}
        </section>
      `
    )
    .join("");
}

/* ─────────────────────────────────────────────────────────────
   섹션 액션: 복사 (Word HTML) + Word 저장 (.docx 다운로드)
───────────────────────────────────────────────────────────── */

/**
 * 표준 양식 스타일(카테고리/과제/내용_)을 Word HTML 클립보드로 복사.
 *
 * 핵심: @list CSS + mso-list: l0 levelX lfo1 으로 진짜 다단계 목록 구성
 *   → Word에서 Tab/Shift+Tab 으로 레벨 간 이동 가능
 *   → ■ / – 은 목록 정의에서 자동 삽입 (인라인 텍스트 불필요)
 *
 * 실측 template.docx/styles.xml 값:
 *   level1 = 카테고리 (a,  numId=12 ilvl=0): 불릿 없음
 *   level2 = 과제명   (a0, numId=12 ilvl=1): ■ Wingdings \F0A0, left=49.6pt hang=28.35pt
 *   level3 = 내용_    (a1, numId=11 ilvl=2): – 나눔스퀘어_ac,   left=49.65pt hang=14.2pt
 */
async function copySectionToClipboard(sectionIndex, btn) {
  const doc = getPreviewDocument();
  const section = doc.sections[sectionIndex];
  if (!section) return;

  const lang = state.outputLang.startsWith("ko") ? "KO" : "EN";
  const originalLabel = btn ? btn.innerHTML : "";
  const markDone  = () => { if (btn) { btn.innerHTML = "✓ 복사됨"; btn.classList.add("copied"); setTimeout(() => { btn.innerHTML = originalLabel; btn.classList.remove("copied"); }, 2000); } };
  const markError = () => { if (btn) { btn.innerHTML = originalLabel; } };

  // ── 표 셀 공통 inline style ────────────────────────────────
  const S_TD = "font-family:'나눔스퀘어_ac','맑은 고딕',sans-serif;" +
    "mso-fareast-font-family:'나눔스퀘어_ac';mso-ascii-font-family:'나눔스퀘어_ac';mso-hansi-font-family:'나눔스퀘어_ac';" +
    "font-size:10.0pt;color:#1A1A1A;padding:4pt 6pt;vertical-align:middle;border:1pt solid #999999;";

  // ── HTML 본문 조립 ─────────────────────────────────────────
  let rows = "";

  // 카테고리 (level1: 불릿 없음, Word mso-list로 처리)
  rows += `<p class=a><span lang=${lang}>${escapeHtml(section.category)}</span></p>\n`;

  for (const item of section.items) {
    // 과제명 (level2: ■ 은 @list l0:level2 정의에서 자동 삽입)
    if (item.title) {
      rows += `<!--[if !supportLists]--><p class=a0><span style='font-family:Wingdings;mso-fareast-font-family:"나눔스퀘어_ac Bold"'>&#xF0A0;<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span><!--[endif]--><span lang=${lang}>${escapeHtml(item.title)}</span></p>\n`;
    }

    // 내용_ (level3: – 은 @list l0:level3 정의에서 자동 삽입)
    for (const detail of item.details) {
      const hasBullet = hasDetailBullet(detail);
      const text = stripDetailBullet(detail);
      text.split("\n").forEach((line, i) => {
        if (!line.trim() && i > 0) return;
        if (i === 0 && hasBullet) {
          // 불릿 있는 내용: level3 목록 항목
          rows += `<!--[if !supportLists]--><p class=a1><span lang=${lang} style='font-family:"나눔스퀘어_ac"'>&#8211;<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span><!--[endif]--><span lang=${lang}>${escapeHtml(line)}</span></p>\n`;
        } else {
          // 불릿 없는 내용 / 연속행: 수동 들여쓰기
          rows += `<p class=a1nb><span lang=${lang}>${escapeHtml(line)}</span></p>\n`;
        }
      });
    }

    // 표
    for (const table of item.tables) {
      if (!table.rows?.length) continue;
      rows += `<table style="border-collapse:collapse;width:100%;mso-table-lspace:.05pt;mso-table-rspace:.05pt;margin:4pt 0 4pt 49.65pt;">\n`;
      table.rows.forEach((row, rI) => {
        const isHeader = rI === 0;
        rows += "<tr>\n";
        row.forEach(cell => {
          const norm = normalizeTableCellData(cell);
          if (norm.rowSpan === 0) return;
          const tag  = isHeader ? "th" : "td";
          const cs   = norm.colSpan > 1 ? ` colspan="${norm.colSpan}"` : "";
          const rs   = norm.rowSpan > 1 ? ` rowspan="${norm.rowSpan}"` : "";
          const align = norm.align || (isHeader ? "center" : "left");
          const hExtra = isHeader ? "background:#E8E8E8;font-weight:bold;text-align:center;" : `text-align:${align};`;
          const cellHtml = getTableCellDisplayValue(cell, isHeader).split("\n").map(l => escapeHtml(l)).join("<br>");
          rows += `<${tag}${cs}${rs} style="${S_TD}${hExtra}"><span lang=${lang}>${cellHtml}</span></${tag}>\n`;
        });
        rows += "</tr>\n";
      });
      rows += "</table>\n";
    }

    // 이미지
    for (const img of (item.images || [])) {
      if (!img?.dataUrl) continue;
      const wPx = img.cx > 0 ? Math.round(img.cx / 914400 * 96) : null;
      rows += `<p style="margin:4pt 0 4pt 49.65pt;"><img src="${img.dataUrl}"${wPx ? ` width="${wPx}"` : ""} style="max-width:100%;"></p>\n`;
    }
  }

  // ── @list + 단락 스타일 정의 ──────────────────────────────────
  // @list l0: Word 다단계 목록 정의 (template.docx abstractNumId=7 구조 반영)
  //   level2 = 과제명: Wingdings \F0A0(■), left=49.6pt, hang=28.35pt → ■ at 21.25pt
  //   level3 = 내용_:  나눔스퀘어_ac – ,  left=49.65pt, hang=14.2pt → – at 35.45pt
  const styleSection = `<style>
@list l0 {mso-list-id:20260316;mso-list-type:hybrid;mso-list-template-ids:20260316 -1 -1 -1 -1 -1 -1 -1 -1;}
@list l0:level1 {mso-level-number-format:none;mso-level-text:"";mso-level-tab-stop:none;mso-level-number-position:left;margin-left:0pt;text-indent:0pt;}
@list l0:level2 {
  mso-level-number-format:bullet;mso-level-text:"\F0A0";
  mso-level-tab-stop:49.6pt;mso-level-number-position:left;
  margin-left:49.6pt;text-indent:-28.35pt;
  font-family:Wingdings;mso-bidi-font-family:Wingdings;}
@list l0:level3 {
  mso-level-number-format:bullet;mso-level-text:"\2013";
  mso-level-tab-stop:49.65pt;mso-level-number-position:left;
  margin-left:49.65pt;text-indent:-14.2pt;
  font-family:"나눔스퀘어_ac";mso-bidi-font-family:"나눔스퀘어_ac";}
p.a {
  mso-style-name:"카테고리";
  mso-list:l0 level1 lfo1;
  margin:0pt;line-height:150%;
  font-family:'나눔스퀘어_ac ExtraBold','나눔스퀘어_ac','맑은 고딕',sans-serif;
  mso-fareast-font-family:'나눔스퀘어_ac ExtraBold';mso-ascii-font-family:'나눔스퀘어_ac ExtraBold';mso-hansi-font-family:'나눔스퀘어_ac ExtraBold';
  font-size:14.0pt;font-weight:bold;color:#0070C0;}
p.a0 {
  mso-style-name:"과제";
  mso-list:l0 level2 lfo1;
  margin:0pt;margin-left:49.6pt;text-indent:-28.35pt;line-height:115%;
  font-family:'나눔스퀘어_ac Bold','나눔스퀘어_ac','맑은 고딕',sans-serif;
  mso-fareast-font-family:'나눔스퀘어_ac Bold';mso-ascii-font-family:'나눔스퀘어_ac Bold';mso-hansi-font-family:'나눔스퀘어_ac Bold';
  font-size:13.0pt;font-weight:bold;color:#000000;}
p.a1 {
  mso-style-name:"내용_";
  mso-list:l0 level3 lfo1;
  margin:0pt;margin-left:49.65pt;text-indent:-14.2pt;line-height:115%;
  font-family:'나눔스퀘어_ac','맑은 고딕',sans-serif;
  mso-fareast-font-family:'나눔스퀘어_ac';mso-ascii-font-family:'나눔스퀘어_ac';mso-hansi-font-family:'나눔스퀘어_ac';
  font-size:12.0pt;color:#1A1A1A;}
p.a1nb {
  mso-style-name:"내용_";
  margin:0pt;margin-left:49.65pt;text-indent:0pt;line-height:115%;
  font-family:'나눔스퀘어_ac','맑은 고딕',sans-serif;
  mso-fareast-font-family:'나눔스퀘어_ac';mso-ascii-font-family:'나눔스퀘어_ac';mso-hansi-font-family:'나눔스퀘어_ac';
  font-size:12.0pt;color:#1A1A1A;}
</style>`;

  const wordHtml = [
    `<html xmlns:o='urn:schemas-microsoft-com:office:office'`,
    `      xmlns:w='urn:schemas-microsoft-com:office:word'`,
    `      xmlns='http://www.w3.org/TR/REC-html40'>`,
    `<head><meta charset="UTF-8">`,
    styleSection,
    `<!--[if gte mso 9]><xml>`,
    `<w:WordDocument><w:View>Normal</w:View><w:Zoom>0</w:Zoom>`,
    `<w:PunctuationKerning/></w:WordDocument>`,
    `</xml><![endif]-->`,
    `</head>`,
    `<body lang=${lang} style='tab-interval:40.0pt;word-wrap:break-word'>`,
    `<div class=WordSection1>`,
    rows,
    `</div></body></html>`,
  ].join("\n");

  // ── 클립보드에 쓰기 ─────────────────────────────────────────
  try {
    await navigator.clipboard.write([
      new ClipboardItem({ "text/html": new Blob([wordHtml], { type: "text/html" }) })
    ]);
    markDone();
  } catch {
    // Clipboard API 미지원 시 execCommand 폴백
    const el = document.createElement("div");
    el.innerHTML = wordHtml;
    Object.assign(el.style, { position: "fixed", top: "-9999px", left: "-9999px" });
    document.body.appendChild(el);
    const range = document.createRange();
    range.selectNodeContents(el);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    const ok = document.execCommand("copy");
    sel.removeAllRanges();
    document.body.removeChild(el);
    ok ? markDone() : markError();
  }
}

/** 해당 섹션을 표준 template.docx 기반 .docx 파일로 다운로드 */
async function downloadSectionDocx(sectionIndex, btn) {
  if (!window.JSZip) {
    alert("JSZip 라이브러리를 불러오지 못했습니다. 잠시 후 다시 시도해 주세요.");
    return;
  }

  const doc = getPreviewDocument();
  const section = doc.sections[sectionIndex];
  if (!section) return;

  const originalLabel = btn ? btn.innerHTML : "";
  const markLoading = () => { if (btn) { btn.innerHTML = "⏳ 생성 중…"; btn.disabled = true; } };
  const markDone    = () => { if (btn) { btn.innerHTML = "✓ 다운로드됨"; btn.classList.add("copied"); setTimeout(() => { btn.innerHTML = originalLabel; btn.classList.remove("copied"); btn.disabled = false; }, 2500); } };
  const markError   = () => { if (btn) { btn.innerHTML = originalLabel; btn.disabled = false; } };

  markLoading();

  try {
    const templateBuffer = await fetch("./assets/template.docx").then(r => r.arrayBuffer());
    const zip = await window.JSZip.loadAsync(templateBuffer);

    // 단일 섹션 문서 구성 (섹션만 포함, 헤더 없음)
    const singleSectionDoc = { ...doc, sections: [section] };
    const imageDataMap = await injectImagesToWordZip(zip, [singleSectionDoc]);

    const xmlString = await zip.file("word/document.xml").async("string");
    const parser = new DOMParser();
    const xml = parser.parseFromString(xmlString, "application/xml");
    const body = xml.getElementsByTagNameNS(WORD_NS, "body")[0];
    const sectPr = body.getElementsByTagNameNS(WORD_NS, "sectPr")[0].cloneNode(true);

    while (body.firstChild) body.removeChild(body.firstChild);

    // 헤더 없이 섹션 본문만 빌드
    buildTemplateBody(xml, body, singleSectionDoc, {
      lang: state.outputLang.startsWith("ko") ? "ko" : "en",
      includeHeader: false,
      imageDataMap,
    });
    body.appendChild(sectPr);

    zip.file("word/document.xml", new XMLSerializer().serializeToString(xml));
    const blob = await zip.generateAsync({ type: "blob" });

    // 파일명: 카테고리명 기반
    const safeName = section.category.replace(/[\\/:*?"<>|]/g, "_").trim() || `section_${sectionIndex + 1}`;
    downloadBlob(blob, `${safeName}.docx`);
    markDone();
  } catch (err) {
    console.error(err);
    alert("섹션 파일 생성에 실패했습니다.");
    markError();
  }
}

function renderPreview() {
  if (!els.previewTitle || !els.previewDate || !els.previewOrg || !els.dateLabel || !els.orgLabel || !els.translationStatus || !els.previewBody) {
    return;
  }

  const isDual = state.outputLang === "ko-en" || state.outputLang === "en-ko";
  const primaryLang = state.outputLang.startsWith("ko") ? "ko" : "en";
  const doc = getPreviewDocument();
  const labels = primaryLang === "ko" ? { date: "일자", org: "부서" } : { date: "Date", org: "Department" };

  els.previewTitle.textContent = doc.title || "제목";
  els.previewDate.textContent = doc.date || "";
  els.previewOrg.textContent = doc.org || "";
  els.dateLabel.textContent = labels.date;
  els.orgLabel.textContent = labels.org;
  els.translationStatus.textContent = getTranslationStatusMessage();
  els.translationStatus.className =
    state.aiTranslationStatus === "loading"
      ? "hint translating"
      : state.aiTranslationStatus === "error"
        ? "hint translate-error"
        : "hint";

  if (!doc.sections.length) {
    els.previewBody.innerHTML = '<div class="empty-state">본문이 없습니다.</div>';
    if (primaryLang === "en" || isDual) {
      scheduleAiTranslation();
    }
    return;
  }

  if (!isDual) {
    // 단일 언어 모드 (기존 동작)
    els.previewBody.innerHTML = renderDocumentSections(doc);
    renderPreviewSelection();
    if (primaryLang === "en") {
      scheduleAiTranslation();
    }
    return;
  }

  // 이중 언어 모드 (ko-en / en-ko)
  const koDoc = state.document;
  const enDoc = (() => {
    const sig = getDocumentSignature(state.document);
    if (state.aiTranslationCache.signature === sig && state.aiTranslationCache.document) {
      return state.aiTranslationCache.document;
    }
    return translateDocument(state.document);
  })();

  const firstDoc  = state.outputLang === "ko-en" ? koDoc : enDoc;
  const secondDoc = state.outputLang === "ko-en" ? enDoc  : koDoc;
  const firstLabel  = state.outputLang === "ko-en" ? "국문" : "English";
  const secondLabel = state.outputLang === "ko-en" ? "English" : "국문";
  const secondDateLabel = state.outputLang === "ko-en" ? "Date" : "일자";
  const secondOrgLabel  = state.outputLang === "ko-en" ? "Department" : "부서";

  const secondBodyHtml = secondDoc.sections.length
    ? renderDocumentSections(secondDoc, { secondary: true })
    : `<div class="empty-state">${state.aiTranslationStatus === "loading" ? "번역 중..." : "본문이 없습니다."}</div>`;

  els.previewBody.innerHTML = `
    <div class="lang-section-wrap">
      <div class="lang-section-label">${firstLabel}</div>
      ${renderDocumentSections(firstDoc)}
    </div>
    <div class="lang-divider">
      <span class="lang-divider-label">${secondLabel}</span>
    </div>
    <div class="lang-section-wrap secondary-lang-section">
      <div class="lang-meta-block">
        <p><span>${secondDateLabel}</span>: <strong>${escapeHtml(secondDoc.date || "")}</strong></p>
        <p><span>${secondOrgLabel}</span>: <strong>${escapeHtml(secondDoc.org || "")}</strong></p>
      </div>
      ${secondBodyHtml}
    </div>
  `;
  renderPreviewSelection();
  scheduleAiTranslation();
}

function renderDetail(detail) {
  const hasBullet = hasDetailBullet(detail);
  const text = stripDetailBullet(detail);
  const lines = text.split("\n");

  if (lines.length <= 1) {
    return `<p class="doc-detail ${hasBullet ? "with-bullet" : "without-bullet"}">${escapeHtml(text)}</p>`;
  }

  return lines
    .map((line, idx) => {
      if (idx === 0) {
        return `<p class="doc-detail ${hasBullet ? "with-bullet" : "without-bullet"}">${escapeHtml(line)}</p>`;
      }
      // Shift+Enter 연속 행: 대쉬 없이 같은 들여쓰기
      if (!line.trim()) {
        return "";
      }
      return `<p class="doc-detail continuation">${escapeHtml(line)}</p>`;
    })
    .join("");
}

function renderDocImage(image) {
  if (!image?.dataUrl) {
    return "";
  }
  // EMU → px 변환 (EMU_PER_PX: 96 DPI 기준)
  const imgStyles = [];
  if (image.cx > 0) imgStyles.push(`width:${Math.round(image.cx / EMU_PER_PX)}px`);
  const imgStyleAttr = imgStyles.length ? ` style="${imgStyles.join(";")}"` : "";

  // 수평 정렬 (center/right) + 위아래 간격 (distT/distB EMU → px)
  const wrapStyles = [];
  const align = image.align;
  if (align === "center") {
    // 중앙 정렬: flex로 수평 중앙
    wrapStyles.push("display:flex", "justify-content:center", "margin-left:0");
  } else if (align === "right") {
    // 오른쪽 정렬
    wrapStyles.push("display:flex", "justify-content:flex-end", "margin-left:0");
  }
  const tPx = Math.round((image.distT || 0) / EMU_PER_PX);
  const bPx = Math.round((image.distB || 0) / EMU_PER_PX);
  if (tPx > 0) wrapStyles.push(`margin-top:${tPx}px`);
  if (bPx > 0) wrapStyles.push(`margin-bottom:${bPx}px`);
  const wrapStyleAttr = wrapStyles.length ? ` style="${wrapStyles.join(";")}"` : "";

  return `<div class="doc-image-wrap"${wrapStyleAttr}><img class="doc-image"${imgStyleAttr} src="${image.dataUrl}" alt="삽입 이미지" /></div>`;
}

function renderTable(table) {
  if (!table.rows?.length) {
    return "";
  }

  const renderCell = (cell, isHeader) => {
    const normalized = normalizeTableCellData(cell);
    const rs = normalized.rowSpan;
    const cs = normalized.colSpan;
    if (rs === 0) return ""; // 병합된 숨김 셀
    const tag = isHeader ? "th" : "td";
    const attrs = [
      cs > 1 ? `colspan="${cs}"` : "",
      rs > 1 ? `rowspan="${rs}"` : ""
    ].filter(Boolean).join(" ");
    const styleAttr = normalized.align ? ` style="text-align:${normalized.align}"` : "";
    const cellText = getTableCellDisplayValue(cell, isHeader);
    const cellHtml = cellText.split("\n").map((line) => escapeHtml(line)).join("<br>");
    const cellImgsHtml = (normalized.images || []).map((img) => renderDocImage(img)).join("");
    return `<${tag}${attrs ? " " + attrs : ""}${styleAttr}>${cellHtml}${cellImgsHtml}</${tag}>`;
  };

  const [headerRow, ...bodyRows] = table.rows;
  return `
    <div class="doc-table-wrap">
      <table class="doc-table">
        <thead>
          <tr>${headerRow.map((c) => renderCell(c, true)).join("")}</tr>
        </thead>
        <tbody>
          ${bodyRows.map((row) => `<tr>${row.map((c) => renderCell(c, false)).join("")}</tr>`).join("")}
        </tbody>
      </table>
    </div>
  `;
}

function renderViewButtons() {
  els.koreanViewButton?.classList.toggle("active", state.view === "ko");
  els.englishViewButton?.classList.toggle("active", state.view === "en");
  if (els.outputLangGroup) {
    els.outputLangGroup.querySelectorAll("[data-output-lang]").forEach((btn) => {
      btn.classList.toggle("active", btn.dataset.outputLang === state.outputLang);
    });
  }
}

function renderPreviewSelection() {
  els.sheetSticky?.classList.toggle("is-selected", state.activeSelection.type === "overview");
}

function schedulePreviewRender() {
  if (state.previewRenderTimer) {
    cancelAnimationFrame(state.previewRenderTimer);
  }
  state.previewRenderTimer = requestAnimationFrame(() => {
    state.previewRenderTimer = null;
    renderPreview();
  });
}

/** 가이드 항목 텍스트를 리치 HTML로 변환 */
function formatGuideText(raw) {
  let html = escapeHtml(raw);
  // __텍스트__ → 강조 bold
  html = html.replace(/__([^_]+)__/g, '<strong class="guide-em">$1</strong>');
  // (예, ...) 인라인 → pill 스타일
  html = html.replace(/(\(예,\s*[^)]+\))/g, '<span class="guide-inline-eg">$1</span>');
  return html;
}

/** 가이드 항목 배열 → 섹션 구조 HTML 생성 */
function renderGuideContent(items) {
  const SECTION_COLORS = [
    { bg: "#e8f4fd", border: "#0070c0" }, // 파란색
    { bg: "#e6f7f2", border: "#0891b2" }, // 청록색
    { bg: "#fef9e7", border: "#b45309" }, // 황금색
    { bg: "#f3eeff", border: "#7c3aed" }, // 보라색
    { bg: "#f0f4f8", border: "#475569" }, // 회청색
  ];

  let html = '<div class="guide-content">';
  let sectionIdx = -1;
  let inSection = false;

  items.forEach((item) => {
    // 섹션 헤딩: "1. 파일명", "2. 기본 서식" 등
    const sectionMatch = item.match(/^(\d+)\.\s+(.+)$/);
    if (sectionMatch) {
      if (inSection) html += "</div></div>"; // 이전 섹션 body + block 닫기
      sectionIdx++;
      const c = SECTION_COLORS[sectionIdx % SECTION_COLORS.length];
      html += `<div class="guide-section-block" style="--sc:${c.border};--sb:${c.bg}">
        <div class="guide-section-head">
          <span class="guide-section-badge">${escapeHtml(sectionMatch[1])}</span>
          <span class="guide-section-title">${escapeHtml(sectionMatch[2])}</span>
        </div>
        <div class="guide-section-body">`;
      inSection = true;
      return;
    }

    // 불릿 항목: "- 텍스트" — " : " 구분자로 서브 라인 처리
    if (/^-\s/.test(item)) {
      const rawText = item.replace(/^-\s*/, "");
      const parts = rawText.split(" : ");
      const mainHtml = formatGuideText(parts[0]);
      const subHtml = parts
        .slice(1)
        .map((p) => `<div class="guide-sub-row">${formatGuideText(p)}</div>`)
        .join("");
      html += `<div class="guide-bullet-row">
        <span class="guide-bullet-mark">–</span>
        <div class="guide-bullet-body">
          <div class="guide-bullet-main">${mainHtml}</div>
          ${subHtml}
        </div>
      </div>`;
      return;
    }

    // 예시 라인: "(예, ...)"
    if (/^\(예,/.test(item)) {
      const inner = item.replace(/^\(예,\s*/, "").replace(/\)$/, "");
      html += `<div class="guide-eg-row">
        <span class="guide-eg-badge">예시</span>
        <span class="guide-eg-text">${escapeHtml(inner)}</span>
      </div>`;
      return;
    }

    // 일반 텍스트 폴백
    html += `<div class="guide-plain">${escapeHtml(item)}</div>`;
  });

  if (inSection) html += "</div></div>"; // 마지막 섹션 닫기
  html += "</div>";
  return html;
}

function renderGuidePage() {
  if (!els.guideList || !els.memberSearchInput || !els.memberSearchResult || !els.glossaryGuideTable) {
    return;
  }

  els.guideList.innerHTML = renderGuideContent(state.guideItems);

  const keyword = cleanText(els.memberSearchInput.value).toLowerCase();
  const matches = state.names.filter(([kr, en]) => {
    if (!keyword) {
      return true;
    }
    return kr.toLowerCase().includes(keyword) || en.toLowerCase().includes(keyword);
  });

  els.memberSearchResult.innerHTML = matches.length
    ? matches
        .slice(0, 50)
        .map(
          ([kr, en]) => `
            <div class="search-card">
              <strong>${escapeHtml(kr)}</strong>
              <span>${escapeHtml(en)}</span>
            </div>
          `
        )
        .join("")
    : '<div class="editor-empty">검색 결과가 없습니다.</div>';

  els.glossaryGuideTable.innerHTML = state.glossary
    .slice(0, 120)
    .map(
      ([ko, en]) => `
        <div class="guide-dictionary-row">
          <span>${escapeHtml(ko)}</span>
          <span>${escapeHtml(en)}</span>
        </div>
      `
    )
    .join("");
}

function renderAdminPage() {
  if (!els.adminLocked || !els.adminUnlocked) {
    return;
  }

  // 잠금 상태 변화 감지: 잠금 해제 직후에만 textarea를 state 값으로 초기화
  const wasLocked = els.adminUnlocked.classList.contains("hidden");

  els.adminLocked.classList.toggle("hidden", state.adminUnlocked);
  els.adminUnlocked.classList.toggle("hidden", !state.adminUnlocked);

  if (!state.adminUnlocked) {
    return;
  }

  // 매 render()마다 덮어쓰지 않고, 잠금 해제 시 1회만 초기값 세팅
  // (이후 render()가 발생해도 사용자의 미저장 편집 내용을 보호)
  if (wasLocked) {
    els.adminGuideInput.value = state.guideItems.join("\n");
    els.adminGlossaryInput.value = state.glossary.map(([ko, en]) => `${ko}\t${en}`).join("\n");
    els.adminNamesInput.value = state.names.map(([kr, en]) => `${kr}\t${en}`).join("\n");
    els.adminAiModelInput.value = state.aiConfig.model;
  }
}

function getPreviewDocument() {
  if (state.view === "ko") {
    return state.document;
  }

  const signature = getDocumentSignature(state.document);
  if (state.aiTranslationCache.signature === signature && state.aiTranslationCache.document) {
    return state.aiTranslationCache.document;
  }

  return translateDocument(state.document);
}

function getTranslationStatusMessage() {
  if (state.view === "ko") {
    return "국문 원본 구조를 그대로 보여줍니다.";
  }

  const currentSignature = getDocumentSignature(state.document);
  const hasFreshAiResult =
    state.aiTranslationCache.signature === currentSignature && state.aiTranslationCache.document;

  if (state.aiTranslationStatus === "loading") {
    return "브라우저 내장 번역을 우선 사용해 영문 번역 중입니다.";
  }

  if (state.aiTranslationStatus === "error") {
    return `브라우저 내장 번역을 사용할 수 없어 임시 규칙 기반 번역으로 표시합니다. ${state.aiTranslationError}`;
  }

  if (hasFreshAiResult) {
    return "영문 보기는 브라우저 내장 번역 결과를 우선 표시합니다. 용어집과 영문 이름도 함께 반영했습니다.";
  }

  return "영문 보기는 브라우저 내장 번역을 준비 중입니다.";
}

function getDocumentSignature(doc) {
  return JSON.stringify(doc);
}

function clearAiTranslationCache() {
  if (state.aiTranslationTimer) {
    clearTimeout(state.aiTranslationTimer);
  }
  state.aiTranslationCache = { signature: "", document: null };
  state.aiTranslationStatus = "idle";
  state.aiTranslationError = "";
  state.aiTranslationTimer = null;
}

function scheduleAiTranslation() {
  if (state.view !== "en") {
    return;
  }

  const signature = getDocumentSignature(state.document);
  if (state.aiTranslationCache.signature === signature && state.aiTranslationCache.document) {
    return;
  }

  if (state.aiTranslationStatus === "loading") {
    return;
  }

  if (state.aiTranslationTimer) {
    clearTimeout(state.aiTranslationTimer);
  }

  state.aiTranslationStatus = "loading";
  state.aiTranslationError = "";
  state.aiTranslationTimer = window.setTimeout(() => {
    ensureAiTranslation(state.document).catch((error) => {
      console.error(error);
    });
  }, 500);
}

async function ensureAiTranslation(doc) {
  const signature = getDocumentSignature(doc);
  if (state.aiTranslationCache.signature === signature && state.aiTranslationCache.document) {
    return state.aiTranslationCache.document;
  }

  const requestId = ++state.aiTranslationRequestId;
  if (state.aiTranslationTimer) {
    clearTimeout(state.aiTranslationTimer);
    state.aiTranslationTimer = null;
  }
  state.aiTranslationStatus = "loading";
  state.aiTranslationError = "";
  renderPreview();

  try {
    const translated = await requestAiTranslation(doc);
    if (requestId !== state.aiTranslationRequestId) {
      return translated;
    }

    state.aiTranslationCache = { signature, document: translated };
    state.aiTranslationStatus = "ready";
    state.aiTranslationError = "";
    renderPreview();
    return translated;
  } catch (error) {
    if (requestId !== state.aiTranslationRequestId) {
      return translateDocument(doc);
    }

    state.aiTranslationStatus = "error";
    state.aiTranslationError = error?.message || "AI API 호출 오류";
    renderPreview();
    return translateDocument(doc);
  }
}

async function getEnglishDocumentForOutput(doc) {
  const signature = getDocumentSignature(doc);
  if (signature === getDocumentSignature(state.document)) {
    return ensureAiTranslation(doc);
  }

  try {
    return await requestAiTranslation(doc);
  } catch (error) {
    console.error(error);
    return translateDocument(doc);
  }
}

async function requestAiTranslation(doc) {
  try {
    return await requestBuiltInBrowserTranslation(doc);
  } catch (browserError) {
    console.warn(browserError);
  }

  return translateDocument(doc);
}

async function requestBuiltInBrowserTranslation(doc) {
  const translator = await getBuiltInTranslator();

  return {
    title: await translateTextWithBuiltInApi(translator, doc.title),
    date: await translateTextWithBuiltInApi(translator, doc.date),
    org: await translateTextWithBuiltInApi(translator, doc.org),
    sections: await Promise.all(
      doc.sections.map(async (section) => ({
        ...section,
        category: await translateCategoryWithBuiltInApi(translator, section.category),
        items: await Promise.all(
          section.items.map(async (item) => ({
            ...item,
            title: await translateTextWithBuiltInApi(translator, item.title),
            details: await Promise.all(item.details.map((detail) => translateTextWithBuiltInApi(translator, detail))),
            tables: await Promise.all(
              item.tables.map(async (table) => ({
                rows: await Promise.all(
                  table.rows.map((row) =>
                    Promise.all(
                      row.map(async (cell) => ({
                        ...normalizeTableCellData(cell),
                        text: await translateTextWithBuiltInApi(translator, getTableCellText(cell))
                      }))
                    )
                  )
                )
              }))
            )
          }))
        )
      }))
    )
  };
}

async function getBuiltInTranslator() {
  if (state.browserTranslator) {
    return state.browserTranslator;
  }

  const translatorApi = window.Translator;
  if (!translatorApi || typeof translatorApi.create !== "function") {
    throw new Error("이 브라우저는 내장 번역 API를 지원하지 않습니다. 최신 Chrome에서 확인해 주세요.");
  }

  if (typeof translatorApi.availability === "function") {
    const availability = await translatorApi.availability({
      sourceLanguage: "ko",
      targetLanguage: "en"
    });

    if (availability === "unavailable") {
      throw new Error("브라우저 내장 한영 번역을 사용할 수 없는 환경입니다.");
    }
  }

  state.browserTranslator = await translatorApi.create({
    sourceLanguage: "ko",
    targetLanguage: "en"
  });

  return state.browserTranslator;
}

async function translateCategoryWithBuiltInApi(translator, category) {
  const inside = category.replace(/^\[|\]$/g, "");
  const translated = await translateTextWithBuiltInApi(translator, inside);
  return translated ? `[${translated}]` : "";
}

async function translateTextWithBuiltInApi(translator, text) {
  const detailBullet = hasDetailBullet(text);
  let source = cleanMultilineText(stripDetailBullet(text));
  if (!source) {
    return "";
  }

  const protectedSource = protectDictionaryTerms(source, [...state.names, ...state.glossary]);

  if (!containsHangul(protectedSource.text)) {
    const output = finalizeEnglishText(restoreProtectedTerms(protectedSource.text, protectedSource.tokens));
    return detailBullet ? markDetailBullet(output) : output;
  }

  const translated = await translator.translate(protectedSource.text);
  const output = finalizeEnglishText(restoreProtectedTerms(translated, protectedSource.tokens));
  return detailBullet ? markDetailBullet(output) : output;
}

function getTranslateApiUrl() {
  const { protocol, hostname, port, origin } = window.location;
  if (protocol === "file:" || (hostname === "127.0.0.1" && port === "8001")) {
    return "http://127.0.0.1:8002/api/translate";
  }
  if (hostname === "localhost" && port === "8001") {
    return "http://localhost:8002/api/translate";
  }
  return `${origin}/api/translate`;
}

function formatTranslateError(status, rawText) {
  const text = cleanText(rawText);
  if (!text) {
    return `HTTP ${status}`;
  }

  if (text.includes("insufficient_quota")) {
    return "Gemini API 사용량 또는 결제/프로젝트 한도를 확인해 주세요. 서버의 GEMINI_API_KEY 상태를 점검해야 합니다.";
  }

  if (status === 404) {
    return "번역 API 경로를 찾지 못했습니다. 로컬에서는 http://127.0.0.1:8002 로 접속해야 합니다.";
  }

  if (status === 401 || status === 403) {
    return "서버의 Gemini API 인증에 실패했습니다. 환경변수 GEMINI_API_KEY를 확인해 주세요.";
  }

  return text;
}

function sanitizeTranslatedDocument(translatedDoc, sourceDoc) {
  return {
    title: cleanText(translatedDoc?.title) || translateText(sourceDoc.title),
    date: cleanText(translatedDoc?.date) || translateDateLine(sourceDoc.date),
    org: cleanText(translatedDoc?.org) || translateText(sourceDoc.org),
    sections: sourceDoc.sections.map((section, sectionIndex) => {
      const translatedSection = translatedDoc?.sections?.[sectionIndex] || {};
      return {
        ...section,
        category: cleanText(translatedSection.category) || translateCategory(section.category),
        items: section.items.map((item, itemIndex) => {
          const translatedItem = translatedSection.items?.[itemIndex] || {};
          return {
            ...item,
            title: cleanText(translatedItem.title) || translateText(item.title),
            details: item.details.map(
              (detail, detailIndex) => cleanText(translatedItem.details?.[detailIndex]) || translateText(detail)
            ),
            tables: item.tables.map((table, tableIndex) => {
              const translatedTable = translatedItem.tables?.[tableIndex];
              return {
                rows: table.rows.map((row, rowIndex) =>
                  row.map(
                    (cell, cellIndex) => ({
                      ...normalizeTableCellData(cell),
                      text:
                        cleanText(getTableCellText(translatedTable?.rows?.[rowIndex]?.[cellIndex])) ||
                        translateText(getTableCellText(cell))
                    })
                  )
                )
              };
            })
          };
        })
      };
    })
  };
}

function renderFilename() {
  if (!els.filenamePreview) {
    return;
  }
  const suggestedFilename = buildSuggestedFilename(state.document);
  els.filenamePreview.value = state.document.filename || suggestedFilename;
}

function renderDraftStatus() {
  if (els.draftStatus) {
    els.draftStatus.textContent = state.draftStatus;
  }
}

function renderEditorWidthMode() {
  els.mainTab?.classList.toggle("editor-wide", state.editorWide);
  if (els.toggleEditorWidthButton) {
    els.toggleEditorWidthButton.textContent = state.editorWide ? "기본 폭으로 보기" : "편집 넓게 보기";
  }
}

/** 파일 선택 다이얼로그를 열고, 선택된 이미지의 dataUrl을 콜백으로 전달 */
function pickImageFile(onDataUrl) {
  const fileInput = document.createElement("input");
  fileInput.type = "file";
  fileInput.accept = "image/*";
  fileInput.onchange = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => onDataUrl(ev.target.result);
    reader.readAsDataURL(file);
  };
  fileInput.click();
}

function onEditorClick(event) {
  // 셀 선택 클릭 처리 (td 요소)
  const cellEl = event.target.closest("td[data-action='select-table-cell']");
  if (cellEl) {
    // ★ textarea 직접 클릭 시 selectTableCell/re-render를 건너뜀 → 브라우저 기본 포커스 허용
    if (event.target.tagName === "INPUT" || event.target.tagName === "TEXTAREA") {
      return;
    }
    const sI = Number(cellEl.dataset.sectionIndex);
    const iI = Number(cellEl.dataset.itemIndex);
    const tI = Number(cellEl.dataset.tableIndex);
    const rI = Number(cellEl.dataset.rowIndex);
    const cI = Number(cellEl.dataset.cellIndex);
    selectTableCell(sI, iI, tI, rI, cI, event.ctrlKey || event.metaKey);
    // re-render 후 해당 셀 textarea에 자동 포커스 (단일 클릭으로 바로 편집 가능)
    if (!event.ctrlKey && !event.metaKey) {
      requestAnimationFrame(() => {
        const sel = `textarea[data-action="table-cell"][data-section-index="${sI}"][data-item-index="${iI}"][data-table-index="${tI}"][data-row-index="${rI}"][data-cell-index="${cI}"]`;
        const ta = els.editorRoot?.querySelector(sel);
        if (ta) ta.focus();
      });
    }
    return;
  }
  // 행 선택 클릭 처리
  const rowSelEl = event.target.closest("td[data-action='select-table-row']");
  if (rowSelEl) {
    const sI = Number(rowSelEl.dataset.sectionIndex);
    const iI = Number(rowSelEl.dataset.itemIndex);
    const tI = Number(rowSelEl.dataset.tableIndex);
    const rI = Number(rowSelEl.dataset.rowIndex);
    selectTableRow(sI, iI, tI, rI, event.ctrlKey || event.metaKey);
    return;
  }
  // 열 선택 클릭 처리 (열 헤더 th 요소)
  const colSelEl = event.target.closest("th[data-action='select-table-col']");
  if (colSelEl) {
    const sI = Number(colSelEl.dataset.sectionIndex);
    const iI = Number(colSelEl.dataset.itemIndex);
    const tI = Number(colSelEl.dataset.tableIndex);
    const cI = Number(colSelEl.dataset.colIndex);
    selectTableCol(sI, iI, tI, cI, event.ctrlKey || event.metaKey);
    return;
  }

  const button = event.target.closest("button[data-action]");
  if (!button) {
    return;
  }

  const action = button.dataset.action;
  const sectionIndex = Number(button.dataset.sectionIndex);
  const itemIndex = Number(button.dataset.itemIndex);
  const detailIndex = Number(button.dataset.detailIndex);

  if (action === "focus-overview-title") {
    els.titleInput.focus();
    return;
  }

  if (action === "focus-overview-date") {
    els.dateInput.focus();
    return;
  }

  if (action === "focus-overview-org") {
    els.orgInput.focus();
    return;
  }

  if (action === "add-item") {
    state.document.sections[sectionIndex].items.push(createEmptyItem());
  }

  if (action === "remove-section") {
    state.document.sections.splice(sectionIndex, 1);
    if (!state.document.sections.length || state.activeSelection.sectionIndex === sectionIndex) {
      state.activeSelection = { type: "overview", sectionIndex: null };
    } else if (state.activeSelection.sectionIndex > sectionIndex) {
      state.activeSelection.sectionIndex -= 1;
    }
  }

  if (action === "add-detail") {
    state.document.sections[sectionIndex].items[itemIndex].details.push(markDetailBullet(""));
  }

  if (action === "add-table") {
    state.document.sections[sectionIndex].items[itemIndex].tables.push(createDefaultTable());
  }

  if (action === "add-image") {
    pickImageFile((dataUrl) => {
      const item = state.document.sections[sectionIndex].items[itemIndex];
      if (!item.images) item.images = [];
      item.images.push({ dataUrl });
      scheduleDocumentDraftSave();
      render();
    });
    return;
  }

  if (action === "remove-image") {
    const imageIndex = Number(button.dataset.imageIndex);
    state.document.sections[sectionIndex].items[itemIndex].images.splice(imageIndex, 1);
    scheduleDocumentDraftSave();
    render();
    return;
  }

  if (action === "table-cell-add-image") {
    const tableIndex = Number(button.dataset.tableIndex);
    const rowIndex = Number(button.dataset.rowIndex);
    const cellIndex = Number(button.dataset.cellIndex);
    pickImageFile((dataUrl) => {
      const cell = state.document.sections[sectionIndex].items[itemIndex].tables[tableIndex].rows[rowIndex][cellIndex];
      const norm = normalizeTableCellData(cell);
      norm.images.push({ dataUrl, cx: 0, cy: 0 });
      state.document.sections[sectionIndex].items[itemIndex].tables[tableIndex].rows[rowIndex][cellIndex] = norm;
      scheduleDocumentDraftSave();
      render();
    });
    return;
  }

  if (action === "table-cell-remove-image") {
    const tableIndex = Number(button.dataset.tableIndex);
    const rowIndex = Number(button.dataset.rowIndex);
    const cellIndex = Number(button.dataset.cellIndex);
    const imageIndex = Number(button.dataset.imageIndex);
    const cell = state.document.sections[sectionIndex].items[itemIndex].tables[tableIndex].rows[rowIndex][cellIndex];
    const norm = normalizeTableCellData(cell);
    norm.images.splice(imageIndex, 1);
    state.document.sections[sectionIndex].items[itemIndex].tables[tableIndex].rows[rowIndex][cellIndex] = norm;
    scheduleDocumentDraftSave();
    render();
    return;
  }

  if (action === "remove-table") {
    const tableIndex = Number(button.dataset.tableIndex);
    state.document.sections[sectionIndex].items[itemIndex].tables.splice(tableIndex, 1);
  }

  if (action === "add-table-row") {
    const tableIndex = Number(button.dataset.tableIndex);
    const table = state.document.sections[sectionIndex].items[itemIndex].tables[tableIndex];
    const columnCount = getTableColumnCount(table);
    if (table?.rows?.[0] && columnCount) {
      table.rows.push(Array.from({ length: columnCount }, () => ({ text: "", colSpan: 1, rowSpan: 1 })));
    }
  }

  if (action === "remove-table-row") {
    const tableIndex = Number(button.dataset.tableIndex);
    const rowIndex = Number(button.dataset.rowIndex);
    const table = state.document.sections[sectionIndex].items[itemIndex].tables[tableIndex];
    if (table && rowIndex > 0) {
      table.rows.splice(rowIndex, 1);
    }
  }

  // 열 추가
  if (action === "add-table-col") {
    const tableIndex = Number(button.dataset.tableIndex);
    const table = state.document.sections[sectionIndex].items[itemIndex].tables[tableIndex];
    if (table?.rows) {
      table.rows.forEach((row) => row.push({ text: "", colSpan: 1, rowSpan: 1 }));
    }
  }

  // 오른쪽 셀과 합치기 (colspan 증가)
  if (action === "merge-cell-right") {
    const tableIndex = Number(button.dataset.tableIndex);
    const rowIndex = Number(button.dataset.rowIndex);
    const cellIndex = Number(button.dataset.cellIndex);
    const row = state.document.sections[sectionIndex].items[itemIndex].tables[tableIndex].rows[rowIndex];
    if (row && cellIndex + 1 < row.length) {
      const cur = normalizeTableCellData(row[cellIndex]);
      const nxt = normalizeTableCellData(row[cellIndex + 1]);
      row[cellIndex] = {
        text: cur.text + (nxt.text ? "\n" + nxt.text : ""),
        colSpan: cur.colSpan + nxt.colSpan,
        rowSpan: cur.rowSpan
      };
      row.splice(cellIndex + 1, 1);
    }
  }

  // 셀 분리 (colspan → 1씩 나누기)
  if (action === "split-cell") {
    const tableIndex = Number(button.dataset.tableIndex);
    const rowIndex = Number(button.dataset.rowIndex);
    const cellIndex = Number(button.dataset.cellIndex);
    const row = state.document.sections[sectionIndex].items[itemIndex].tables[tableIndex].rows[rowIndex];
    if (row) {
      const cell = normalizeTableCellData(row[cellIndex]);
      if (cell.colSpan > 1) {
        const newCells = [{ text: cell.text, colSpan: 1, rowSpan: cell.rowSpan }];
        for (let i = 1; i < cell.colSpan; i++) {
          newCells.push({ text: "", colSpan: 1, rowSpan: 1 });
        }
        row.splice(cellIndex, 1, ...newCells);
      }
    }
  }

  if (action === "item-move-up") {
    const items = state.document.sections[sectionIndex].items;
    if (itemIndex > 0) {
      [items[itemIndex - 1], items[itemIndex]] = [items[itemIndex], items[itemIndex - 1]];
    }
  }

  if (action === "item-move-down") {
    const items = state.document.sections[sectionIndex].items;
    if (itemIndex < items.length - 1) {
      [items[itemIndex], items[itemIndex + 1]] = [items[itemIndex + 1], items[itemIndex]];
    }
  }

  if (action === "remove-item") {
    state.document.sections[sectionIndex].items.splice(itemIndex, 1);
    if (!state.document.sections[sectionIndex].items.length) {
      state.document.sections[sectionIndex].items.push(createEmptyItem());
    }
  }

  if (action === "detail-move-up") {
    const details = state.document.sections[sectionIndex].items[itemIndex].details;
    if (detailIndex > 0) {
      [details[detailIndex - 1], details[detailIndex]] = [details[detailIndex], details[detailIndex - 1]];
    }
  }

  if (action === "detail-move-down") {
    const details = state.document.sections[sectionIndex].items[itemIndex].details;
    if (detailIndex < details.length - 1) {
      [details[detailIndex], details[detailIndex + 1]] = [details[detailIndex + 1], details[detailIndex]];
    }
  }

  if (action === "remove-detail") {
    state.document.sections[sectionIndex].items[itemIndex].details.splice(detailIndex, 1);
  }

  // ── 표 선택 기반 액션 ─────────────────────────────────────────────────────
  if (action === "merge-selected-cells") {
    mergeSelectedCells();
    return;
  }
  if (action === "split-selected-cell") {
    splitSelectedCell();
    return;
  }
  if (action === "merge-selected-cells-vertical") {
    mergeSelectedCellsVertical();
    return;
  }
  if (action === "split-selected-cell-vertical") {
    splitSelectedCellVertical();
    return;
  }
  if (action === "align-selected-cells") {
    alignSelectedCells(button.dataset.align);
    return;
  }
  if (action === "insert-row-above") {
    const tI = Number(button.dataset.tableIndex);
    const rI = Number(button.dataset.rowIndex);
    insertTableRowAt(sectionIndex, itemIndex, tI, rI);
    return;
  }
  if (action === "insert-row-below") {
    const tI = Number(button.dataset.tableIndex);
    const rI = Number(button.dataset.rowIndex);
    insertTableRowAt(sectionIndex, itemIndex, tI, rI + 1);
    return;
  }
  if (action === "delete-selected-rows") {
    deleteSelectedTableRows();
    return;
  }
  if (action === "insert-col-before") {
    const tI = Number(button.dataset.tableIndex);
    const cI = Number(button.dataset.colIndex);
    insertTableColAt(sectionIndex, itemIndex, tI, cI, true);
    return;
  }
  if (action === "insert-col-after") {
    const tI = Number(button.dataset.tableIndex);
    const cI = Number(button.dataset.colIndex);
    insertTableColAt(sectionIndex, itemIndex, tI, cI, false);
    return;
  }
  if (action === "delete-selected-cols") {
    deleteSelectedTableCols();
    return;
  }
  if (action === "clear-table-selection") {
    clearTableSelection();
    renderEditor();
    return;
  }

  render();
}

function onSectionListClick(event) {
  const button = event.target.closest("[data-action]");
  if (!button) {
    return;
  }

  const action = button.dataset.action;
  if (action === "select-overview") {
    state.activeSelection = { type: "overview", sectionIndex: null };
    renderSectionList();
    renderEditor();
    renderPreview();
    return;
  }

  if (action === "select-section") {
    state.activeSelection = { type: "section", sectionIndex: Number(button.dataset.sectionIndex) };
    renderEditor();
    renderSectionList();
    renderPreview();
    return;
  }

  if (action === "remove-section") {
    const sectionIndex = Number(button.dataset.sectionIndex);
    state.document.sections.splice(sectionIndex, 1);
    state.activeSelection = { type: "overview", sectionIndex: null };
    render();
  }
}

function onEditorKeydown(event) {
  const target = event.target;
  if (target.dataset?.action !== "table-cell") return;

  if (event.key === "Tab") {
    // Tab / Shift+Tab: 표 셀 간 이동 (워드 수준 키보드 네비게이션)
    event.preventDefault();
    const sI = Number(target.dataset.sectionIndex);
    const iI = Number(target.dataset.itemIndex);
    const tI = Number(target.dataset.tableIndex);
    const selector = `textarea[data-action="table-cell"][data-section-index="${sI}"][data-item-index="${iI}"][data-table-index="${tI}"]`;
    const allCells = Array.from(els.editorRoot.querySelectorAll(selector));
    allCells.sort((a, b) => {
      const rDiff = Number(a.dataset.rowIndex) - Number(b.dataset.rowIndex);
      if (rDiff !== 0) return rDiff;
      return Number(a.dataset.cellIndex) - Number(b.dataset.cellIndex);
    });
    const currentIdx = allCells.indexOf(target);
    if (currentIdx === -1) return;
    const nextIdx = event.shiftKey
      ? (currentIdx - 1 + allCells.length) % allCells.length
      : (currentIdx + 1) % allCells.length;
    const nextCell = allCells[nextIdx];
    nextCell.focus();
    nextCell.setSelectionRange(0, nextCell.value.length);
  }
}

function onEditorInput(event) {
  if (state.editorComposing) {
    return;
  }

  const target = event.target;
  const action = target.dataset.action;
  if (!action) {
    return;
  }

  const sectionIndex = Number(target.dataset.sectionIndex);
  const itemIndex = Number(target.dataset.itemIndex);
  const detailIndex = Number(target.dataset.detailIndex);
  const tableIndex = Number(target.dataset.tableIndex);
  const rowIndex = Number(target.dataset.rowIndex);
  const cellIndex = Number(target.dataset.cellIndex);

  if (action === "section-title") {
    state.document.sections[sectionIndex].category = target.value;
  }

  if (action === "item-title") {
    state.document.sections[sectionIndex].items[itemIndex].title = target.value;
  }

  if (action === "detail-input") {
    state.document.sections[sectionIndex].items[itemIndex].details[detailIndex] =
      target.dataset.detailBullet === "true" ? markDetailBullet(target.value) : cleanMultilineText(target.value);
  }

  if (action === "table-cell") {
    const currentCell = state.document.sections[sectionIndex].items[itemIndex].tables[tableIndex].rows[rowIndex][cellIndex];
    state.document.sections[sectionIndex].items[itemIndex].tables[tableIndex].rows[rowIndex][cellIndex] = {
      ...normalizeTableCellData(currentCell),
      text: target.value
    };
    autoResizeSingleTextarea(target);
  }

  if (action === "section-title") {
    renderSectionList();
  }

  scheduleDocumentDraftSave();
  renderDraftStatus();
  schedulePreviewRender();
}

function onPreviewClick(event) {
  // 섹션 액션 버튼 (복사 / Word 저장) — 섹션 선택보다 먼저 처리
  const actBtn = event.target.closest("[data-action='copy-section'],[data-action='download-section']");
  if (actBtn) {
    event.stopPropagation();
    const sI = Number(actBtn.dataset.sectionIndex);
    if (actBtn.dataset.action === "copy-section") {
      copySectionToClipboard(sI, actBtn);
    } else {
      downloadSectionDocx(sI, actBtn);
    }
    return;
  }

  const section = event.target.closest(".doc-section");
  if (!section) {
    return;
  }

  state.activeSelection = { type: "section", sectionIndex: Number(section.dataset.sectionIndex) };
  renderSectionList();
  renderEditor();
  renderPreview();
}

function switchTab(tab) {
  state.activeTab = tab;
  renderTabs();
  if (tab === "integration") {
    renderIntegrationTab();
  }
  if (tab === "translate") {
    initTranslateTab();
  }
  if (tab === "guide") {
    renderGuidePage();
  }
  if (tab === "admin") {
    renderAdminPage();
  }
}

// ── 번역 탭 ──────────────────────────────────────────────────────────────────

/** Chrome AI Translator 인스턴스 (한→영) */
let _translateTabTranslator = null;
/** 번역 탭 디바운스 타이머 */
let _translateDebounceTimer = null;
/** 번역 탭 초기화 완료 여부 */
let _translateTabReady = false;

/**
 * 번역기 인스턴스 생성 시도 순서:
 * 1. Chrome 131+ Translation API (window.translation)
 * 2. Chrome AI API (window.ai.translator)
 * 3. Google Translate 비공개 API (API 키 불필요, 실시간 번역)
 * 4. null → 규칙 기반 translateText() 폴백
 */
async function createBrowserTranslator() {
  // 1. Chrome 131+ Translation API (window.translation)
  if (typeof window.translation !== "undefined") {
    try {
      const pair = await window.translation.canTranslate({ sourceLanguage: "ko", targetLanguage: "en" });
      if (pair !== "no") {
        const t = await window.translation.createTranslator({ sourceLanguage: "ko", targetLanguage: "en" });
        return { type: "browser", translator: { translate: (text) => t.translate(text) } };
      }
    } catch (_) { /* fall through */ }
  }
  // 2. Chrome AI API (window.ai.translator)
  if (typeof window.ai !== "undefined" && typeof window.ai.translator !== "undefined") {
    try {
      const caps = await window.ai.translator.capabilities();
      const avail = caps.languagePairAvailable ? caps.languagePairAvailable("ko", "en") : caps.available;
      if (avail !== "no" && avail !== "unavailable") {
        const t = await window.ai.translator.create({ sourceLanguage: "ko", targetLanguage: "en" });
        return { type: "browser", translator: { translate: (text) => t.translate(text) } };
      }
    } catch (_) { /* fall through */ }
  }
  // 3. Google Translate 비공개 API (API 키 불필요)
  try {
    const testUrl = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=ko&tl=en&dt=t&q=${encodeURIComponent("테스트")}`;
    const testRes = await fetch(testUrl, { signal: AbortSignal.timeout(5000) });
    if (testRes.ok) {
      return {
        type: "google",
        translator: {
          translate: async (text) => {
            // 긴 텍스트는 문단 단위로 분할해 번역 (URL 길이 제한 대응)
            const MAX_CHUNK = 1000;
            const lines = text.split("\n");
            const chunks = [];
            let current = "";
            for (const line of lines) {
              if ((current + "\n" + line).length > MAX_CHUNK && current) {
                chunks.push(current);
                current = line;
              } else {
                current = current ? current + "\n" + line : line;
              }
            }
            if (current) chunks.push(current);

            const translated = await Promise.all(chunks.map(async (chunk) => {
              const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=ko&tl=en&dt=t&q=${encodeURIComponent(chunk)}`;
              const res = await fetch(url);
              if (!res.ok) throw new Error(`HTTP ${res.status}`);
              const data = await res.json();
              return data[0].map((seg) => seg[0]).join("");
            }));
            return translated.join("\n");
          }
        }
      };
    }
  } catch (_) { /* fall through */ }
  return null;
}

/** 번역 탭 초기화 (탭 전환 시 1회만 실행) */
async function initTranslateTab() {
  if (_translateTabReady) return;
  _translateTabReady = true;

  // 이벤트 바인딩
  els.translateInput?.addEventListener("input", onTranslateInput);
  els.translateClearBtn?.addEventListener("click", () => {
    if (els.translateInput) els.translateInput.value = "";
    setTranslateOutput("", false);
    if (els.translateCharCount) els.translateCharCount.textContent = "0자";
  });
  els.translateCopyBtn?.addEventListener("click", async () => {
    const text = els.translateOutput?.innerText?.trim() || "";
    if (!text || text === "번역 결과가 여기에 표시됩니다.") return;
    try {
      await navigator.clipboard.writeText(text);
      const btn = els.translateCopyBtn;
      const prev = btn.textContent;
      btn.textContent = "✓ 복사됨";
      setTimeout(() => { btn.textContent = prev; }, 1500);
    } catch (_) { alert("클립보드 복사에 실패했습니다."); }
  });

  // 브라우저 번역기 초기화
  if (els.translateApiStatus) {
    els.translateApiStatus.textContent = "번역 엔진 초기화 중…";
    els.translateApiStatus.className = "translate-api-badge translate-api-loading";
  }
  const result = await createBrowserTranslator();
  if (result) {
    _translateTabTranslator = result.translator;
    if (els.translateApiStatus) {
      const label = result.type === "google" ? "✓ Google 번역" : "✓ 브라우저 번역 엔진";
      els.translateApiStatus.textContent = label;
      els.translateApiStatus.className = "translate-api-badge translate-api-ok";
    }
  } else {
    if (els.translateApiStatus) {
      els.translateApiStatus.textContent = "오프라인 모드 (규칙 기반 번역)";
      els.translateApiStatus.className = "translate-api-badge translate-api-fallback";
    }
  }
}

/** 입력 이벤트 핸들러 (300ms 디바운스) */
function onTranslateInput() {
  const text = els.translateInput?.value || "";
  if (els.translateCharCount) {
    els.translateCharCount.textContent = `${text.length}자`;
  }
  clearTimeout(_translateDebounceTimer);
  if (!text.trim()) {
    setTranslateOutput("", false);
    return;
  }
  setTranslateLoading(true);
  _translateDebounceTimer = setTimeout(() => runTranslation(text), 300);
}

/** 번역 실행 */
async function runTranslation(text) {
  try {
    let result = "";
    if (_translateTabTranslator) {
      // 브라우저 내장 번역 API — 용어집·이름 사전 보호 후 번역, 복원
      const dictProtected = protectDictionaryTerms(text, [...state.names, ...state.glossary]);
      const translated = await _translateTabTranslator.translate(dictProtected.text);
      result = restoreProtectedTerms(translated, dictProtected.tokens);
    } else {
      // 규칙 기반 번역 (translateText 내부에서 이미 용어집 적용됨)
      // 여러 줄일 경우 줄별로 처리
      result = text
        .split("\n")
        .map((line) => (line.trim() ? translateText(line) : ""))
        .join("\n");
    }
    setTranslateOutput(result, false);
  } catch (e) {
    setTranslateOutput("", true, "번역 중 오류가 발생했습니다: " + e.message);
  }
}

/** 번역 결과/로딩 상태 UI 업데이트 */
function setTranslateLoading(loading) {
  if (els.translateLoadingDot) {
    els.translateLoadingDot.classList.toggle("hidden", !loading);
  }
}

function setTranslateOutput(text, isError, errorMsg) {
  setTranslateLoading(false);
  if (!els.translateOutput) return;
  if (!text && !isError) {
    els.translateOutput.innerHTML = `<span class="translate-placeholder">번역 결과가 여기에 표시됩니다.</span>`;
    return;
  }
  if (isError) {
    els.translateOutput.innerHTML = `<span class="translate-error-msg">${escapeHtml(errorMsg || "오류")}</span>`;
    return;
  }
  // 줄바꿈을 <br>로 변환하여 출력
  els.translateOutput.innerHTML = text
    .split("\n")
    .map((line) => (line ? `<p class="translate-result-line">${escapeHtml(line)}</p>` : `<p class="translate-result-line translate-result-empty"></p>`))
    .join("");
}

function unlockAdmin() {
  if (els.adminPasswordInput.value !== "1234") {
    alert("비밀번호가 올바르지 않습니다.");
    return;
  }
  state.adminUnlocked = true;
  els.adminPasswordInput.value = "";
  renderAdminPage();
}

function lockAdmin() {
  state.adminUnlocked = false;
  renderAdminPage();
}

function saveAdminData() {
  const guideItems = els.adminGuideInput.value
    .split(/\r?\n/)
    .map((line) => cleanText(line))
    .filter(Boolean);
  const glossary = parseTabSeparatedPairs(els.adminGlossaryInput.value);
  const names = parseTabSeparatedPairs(els.adminNamesInput.value);
  const model = cleanText(els.adminAiModelInput.value) || "gemini-2.0-flash";

  state.guideItems = guideItems.length ? guideItems : [...defaultGuideItems];
  state.glossary = glossary.length ? glossary : buildDictionary(defaultGlossary);
  state.names = names.length ? names : buildDictionary(defaultNames);
  state.aiConfig = { model };
  clearAiTranslationCache();

  localStorage.setItem(STORAGE_KEYS.guide, JSON.stringify(state.guideItems));
  localStorage.setItem(STORAGE_KEYS.glossary, JSON.stringify(state.glossary));
  localStorage.setItem(STORAGE_KEYS.names, JSON.stringify(state.names));
  localStorage.setItem(STORAGE_KEYS.aiModel, state.aiConfig.model);

  render();
  alert("Admin 데이터가 저장되었습니다.");
}

function resetAdminData() {
  state.guideItems = [...defaultGuideItems];
  state.glossary = buildDictionary(defaultGlossary);
  state.names = buildDictionary(defaultNames);
  state.aiConfig = { model: "gemini-2.0-flash" };
  clearAiTranslationCache();

  localStorage.removeItem(STORAGE_KEYS.guide);
  localStorage.removeItem(STORAGE_KEYS.glossary);
  localStorage.removeItem(STORAGE_KEYS.names);
  localStorage.removeItem(STORAGE_KEYS.aiModel);

  render();
}

async function onIntegrationUpload(event) {
  const input = event.target.closest('[data-action="integration-upload"]');
  if (!input) {
    return;
  }

  const file = input.files?.[0];
  if (!file) {
    return;
  }

  try {
    const arrayBuffer = await file.arrayBuffer();
    const parsed = await parseDocx(arrayBuffer);
    const slotIndex = Number(input.dataset.slotIndex);
    state.integrationSlots[slotIndex] = {
      org: state.integrationSlots[slotIndex].org,
      submitted: true,
      document: {
        ...parsed,
        org: parsed.org || state.integrationSlots[slotIndex].org,
        date: parsed.date || state.integrationDate
      }
    };
    renderIntegrationTab();
  } catch (error) {
    console.error(error);
    alert("통합용 문서 업로드 중 오류가 발생했습니다.");
  }
}

async function exportIntegratedWordDocument() {
  const submittedDocs = state.integrationSlots.filter((slot) => slot.submitted && slot.document);
  if (!submittedDocs.length) {
    alert("업로드된 조직 문서가 없습니다.");
    return;
  }

  try {
    const templateBuffer = await fetch("./assets/template.docx").then((response) => response.arrayBuffer());
    const zip = await window.JSZip.loadAsync(templateBuffer);
    const xmlString = await zip.file("word/document.xml").async("string");
    const parser = new DOMParser();
    const xml = parser.parseFromString(xmlString, "application/xml");
    const body = xml.getElementsByTagNameNS(WORD_NS, "body")[0];
    const sectPr = body.getElementsByTagNameNS(WORD_NS, "sectPr")[0].cloneNode(true);

    while (body.firstChild) {
      body.removeChild(body.firstChild);
    }

    const koDocs = submittedDocs.map((slot) => ({ ...slot.document, date: state.integrationDate || slot.document.date }));
    const enDocs = await Promise.all(koDocs.map((doc) => getEnglishDocumentForOutput(doc)));

    // 이미지 ZIP 주입
    const imageDataMap = await injectImagesToWordZip(zip, [...koDocs, ...enDocs].filter(Boolean));

    body.appendChild(createStyledParagraph(xml, "ad", ["mySUNI Weekly Integrated"]));
    body.appendChild(createMetaParagraph(xml, "일자: ", state.integrationDate || "-"));
    body.appendChild(createEmptyParagraph(xml, "나눔스퀘어_ac Bold"));

    const integOrder = state.integrationOutputLang || "ko-en";
    if (integOrder === "ko-en") {
      body.appendChild(createStyledParagraph(xml, "a", ["[국문 통합본]"]));
      koDocs.forEach((doc, index) => appendIntegratedDocument(xml, body, doc, index > 0, "ko", imageDataMap));
      body.appendChild(createStyledParagraph(xml, "a", ["[English Integrated Version]"]));
      enDocs.forEach((doc, index) => appendIntegratedDocument(xml, body, doc, index > 0, "en", imageDataMap));
    } else {
      body.appendChild(createStyledParagraph(xml, "a", ["[English Integrated Version]"]));
      enDocs.forEach((doc, index) => appendIntegratedDocument(xml, body, doc, index > 0, "en", imageDataMap));
      body.appendChild(createStyledParagraph(xml, "a", ["[국문 통합본]"]));
      koDocs.forEach((doc, index) => appendIntegratedDocument(xml, body, doc, index > 0, "ko", imageDataMap));
    }
    body.appendChild(sectPr);

    zip.file("word/document.xml", new XMLSerializer().serializeToString(xml));
    const blob = await zip.generateAsync({ type: "blob" });
    const filename = `mySUNI Weekly_${(state.integrationDate || "통합본").replace(/[^\dA-Za-z가-힣_-]/g, "")}_Integrated.docx`;
    downloadBlob(blob, filename);
  } catch (error) {
    console.error(error);
    alert("통합 Word 생성에 실패했습니다.");
  }
}

function appendIntegratedDocument(xml, body, doc, withSpacer, lang = "ko", imageDataMap = null) {
  if (withSpacer) {
    body.appendChild(createPageBreakParagraph(xml));
  }
  const orgLabel = lang === "ko" ? "부서: " : "Department: ";
  body.appendChild(createMetaParagraph(xml, orgLabel, doc.org || "-"));
  buildTemplateBody(xml, body, doc, { includeHeader: false, lang, imageDataMap });
}

function saveIntegrationArchive() {
  const archive = {
    date: state.integrationDate || new Date().toISOString().slice(0, 10),
    submittedCount: state.integrationSlots.filter((slot) => slot.submitted).length,
    slots: state.integrationSlots.map((slot) => ({
      org: slot.org,
      submitted: slot.submitted,
      document: slot.document
    }))
  };

  state.archives = [archive, ...state.archives.filter((item) => item.date !== archive.date)].slice(0, 30);
  localStorage.setItem(STORAGE_KEYS.archives, JSON.stringify(state.archives));
  renderIntegrationTab();
}

function onArchiveClick(event) {
  const button = event.target.closest('[data-action="restore-archive"]');
  if (!button) {
    return;
  }
  const archive = state.archives[Number(button.dataset.archiveIndex)];
  if (!archive) {
    return;
  }
  state.integrationDate = archive.date;
  state.integrationSlots = normalizeIntegrationSlots(archive.slots);
  switchTab("integration");
  renderIntegrationTab();
}

function onDocumentArchiveClick(event) {
  const button = event.target.closest("[data-action]");
  if (!button) {
    return;
  }

  const archiveIndex = Number(button.dataset.archiveIndex);
  const archive = state.documentArchives[archiveIndex];
  if (!archive) {
    return;
  }

  if (button.dataset.action === "restore-document-archive") {
    state.document = cloneValue(archive.document);
    state.activeSelection = state.document.sections.length ? { type: "section", sectionIndex: 0 } : { type: "overview", sectionIndex: null };
    clearAiTranslationCache();
    syncInputsFromState();
    switchTab("main");
    render();
    return;
  }

  if (button.dataset.action === "delete-document-archive") {
    state.documentArchives.splice(archiveIndex, 1);
    localStorage.setItem(STORAGE_KEYS.documentArchives, JSON.stringify(state.documentArchives));
    renderDocumentArchiveTab();
  }
}

function normalizeIntegrationSlots(slots) {
  const slotMap = new Map(
    (Array.isArray(slots) ? slots : []).map((slot) => [slot?.org, { ...slot }])
  );

  return INTEGRATION_ORGS.map((org) => {
    const existing = slotMap.get(org);
    return {
      org,
      submitted: Boolean(existing?.submitted),
      document: existing?.document || null
    };
  });
}

async function exportWordDocument() {
  if (!window.JSZip) {
    alert("Word 템플릿 라이브러리를 불러오지 못했습니다. 잠시 후 다시 시도해 주세요.");
    return;
  }

  // 여러 조직 파일인 경우 전체 조직 내보내기
  if (state.subDocuments && state.subDocuments.length > 1) {
    return exportAllOrgsWordDocument();
  }

  try {
    const order = state.outputLang; // "ko-only" | "en-only" | "ko-en" | "en-ko"
    const koDoc = state.document;
    const needEn = order !== "ko-only";
    const enDoc = needEn ? await getEnglishDocumentForOutput(koDoc) : null;

    const templateBuffer = await fetch("./assets/template.docx").then((response) => response.arrayBuffer());
    const zip = await window.JSZip.loadAsync(templateBuffer);

    // 이미지 ZIP 주입 (이미지가 있는 경우만 처리)
    const imageDataMap = await injectImagesToWordZip(zip, [koDoc, enDoc].filter(Boolean));

    const xmlString = await zip.file("word/document.xml").async("string");
    const parser = new DOMParser();
    const xml = parser.parseFromString(xmlString, "application/xml");
    const body = xml.getElementsByTagNameNS(WORD_NS, "body")[0];
    const sectPr = body.getElementsByTagNameNS(WORD_NS, "sectPr")[0].cloneNode(true);

    while (body.firstChild) {
      body.removeChild(body.firstChild);
    }

    if (order === "ko-only") {
      buildTemplateBody(xml, body, koDoc, { lang: "ko", imageDataMap });
    } else if (order === "en-only") {
      buildTemplateBody(xml, body, enDoc, { lang: "en", imageDataMap });
    } else if (order === "ko-en") {
      buildTemplateBody(xml, body, koDoc, { lang: "ko", imageDataMap });
      body.appendChild(createPageBreakParagraph(xml));
      buildTemplateBody(xml, body, enDoc, { lang: "en", includeHeader: true, imageDataMap });
    } else if (order === "en-ko") {
      buildTemplateBody(xml, body, enDoc, { lang: "en", imageDataMap });
      body.appendChild(createPageBreakParagraph(xml));
      buildTemplateBody(xml, body, koDoc, { lang: "ko", includeHeader: true, imageDataMap });
    }

    body.appendChild(sectPr);

    const serializer = new XMLSerializer();
    zip.file("word/document.xml", serializer.serializeToString(xml));
    const blob = await zip.generateAsync({ type: "blob" });
    const filename = (els.filenamePreview.value || "mySUNI Weekly_export").replace(/\.docx$/i, "");
    downloadBlob(blob, `${filename}.docx`);
  } catch (error) {
    console.error(error);
    alert("Word 파일 생성에 실패했습니다.");
  }
}

// 여러 조직 파일 업로드 시 전체 조직 Word 내보내기
async function exportAllOrgsWordDocument() {
  try {
    const order = state.outputLang;
    const koDocs = state.subDocuments;
    const needEn = order !== "ko-only";
    const enDocs = needEn
      ? await Promise.all(koDocs.map((doc) => getEnglishDocumentForOutput(doc)))
      : null;

    const templateBuffer = await fetch("./assets/template.docx").then((r) => r.arrayBuffer());
    const zip = await window.JSZip.loadAsync(templateBuffer);

    // 모든 조직 문서의 이미지를 ZIP에 주입
    const allDocs = [...koDocs, ...(enDocs || [])].filter(Boolean);
    const imageDataMap = await injectImagesToWordZip(zip, allDocs);

    const xmlString = await zip.file("word/document.xml").async("string");
    const parser = new DOMParser();
    const xml = parser.parseFromString(xmlString, "application/xml");
    const body = xml.getElementsByTagNameNS(WORD_NS, "body")[0];
    const sectPr = body.getElementsByTagNameNS(WORD_NS, "sectPr")[0].cloneNode(true);

    while (body.firstChild) body.removeChild(body.firstChild);

    if (order === "ko-only") {
      koDocs.forEach((doc, i) => {
        if (i > 0) body.appendChild(createPageBreakParagraph(xml));
        buildTemplateBody(xml, body, doc, { lang: "ko", imageDataMap });
      });
    } else if (order === "en-only") {
      enDocs.forEach((doc, i) => {
        if (i > 0) body.appendChild(createPageBreakParagraph(xml));
        buildTemplateBody(xml, body, doc, { lang: "en", imageDataMap });
      });
    } else if (order === "ko-en") {
      // 국문 전체 조직 → 영문 전체 조직
      koDocs.forEach((doc, i) => {
        if (i > 0) body.appendChild(createPageBreakParagraph(xml));
        buildTemplateBody(xml, body, doc, { lang: "ko", imageDataMap });
      });
      enDocs.forEach((doc, i) => {
        body.appendChild(createPageBreakParagraph(xml));
        buildTemplateBody(xml, body, doc, { lang: "en", imageDataMap });
      });
    } else if (order === "en-ko") {
      // 영문 전체 조직 → 국문 전체 조직
      enDocs.forEach((doc, i) => {
        if (i > 0) body.appendChild(createPageBreakParagraph(xml));
        buildTemplateBody(xml, body, doc, { lang: "en", imageDataMap });
      });
      koDocs.forEach((doc, i) => {
        body.appendChild(createPageBreakParagraph(xml));
        buildTemplateBody(xml, body, doc, { lang: "ko", imageDataMap });
      });
    }

    body.appendChild(sectPr);
    zip.file("word/document.xml", new XMLSerializer().serializeToString(xml));
    const blob = await zip.generateAsync({ type: "blob" });
    const filename = (els.filenamePreview?.value || "mySUNI Weekly_export").replace(/\.docx$/i, "");
    downloadBlob(blob, `${filename}.docx`);
  } catch (error) {
    console.error(error);
    alert("Word 파일 생성에 실패했습니다.");
  }
}

// 인쇄: 여러 조직이면 전체 조직 포함
function printDocument() {
  if (state.subDocuments && state.subDocuments.length > 1) {
    printAllOrgs();
  } else {
    window.print();
  }
}

// 여러 조직 전체 인쇄 — preview 영역을 임시로 전체 내용으로 교체 후 인쇄
function printAllOrgs() {
  const docs = state.subDocuments;
  if (!docs?.length || !els.previewBody) { window.print(); return; }

  // 현재 미리보기 상태 저장
  const savedTitle     = els.previewTitle?.textContent;
  const savedDate      = els.previewDate?.textContent;
  const savedOrg       = els.previewOrg?.textContent;
  const savedDateLabel = els.dateLabel?.textContent;
  const savedOrgLabel  = els.orgLabel?.textContent;
  const savedBody      = els.previewBody.innerHTML;

  // 첫 번째 조직으로 헤더 설정
  const firstDoc = docs[0];
  if (els.previewTitle)  els.previewTitle.textContent  = firstDoc.title || "제목";
  if (els.previewDate)   els.previewDate.textContent   = firstDoc.date  || "";
  if (els.previewOrg)    els.previewOrg.textContent    = firstDoc.org   || "";
  if (els.dateLabel)     els.dateLabel.textContent     = "일자";
  if (els.orgLabel)      els.orgLabel.textContent      = "부서";

  // 전체 조직 HTML 조합 (2번째 조직부터 자체 헤더 포함)
  const allHtml = docs.map((doc, i) => {
    const sectHtml = doc.sections.length
      ? renderDocumentSections(doc, { secondary: true })
      : '<div class="empty-state">본문이 없습니다.</div>';
    if (i === 0) return sectHtml;
    return `
      <div class="print-org-break">
        <div class="print-org-header">
          <div class="print-org-title">${escapeHtml(doc.title || "제목")}</div>
          <div class="print-org-meta">
            <span>일자: ${escapeHtml(doc.date || "")}</span>&emsp;<span>부서: ${escapeHtml(doc.org || "")}</span>
          </div>
        </div>
        ${sectHtml}
      </div>`;
  }).join("");

  els.previewBody.innerHTML = allHtml;
  window.print();

  // 원래 상태 복원
  if (els.previewTitle)  els.previewTitle.textContent  = savedTitle;
  if (els.previewDate)   els.previewDate.textContent   = savedDate;
  if (els.previewOrg)    els.previewOrg.textContent    = savedOrg;
  if (els.dateLabel)     els.dateLabel.textContent     = savedDateLabel;
  if (els.orgLabel)      els.orgLabel.textContent      = savedOrgLabel;
  if (els.previewBody)   els.previewBody.innerHTML     = savedBody;
}

function translateDocument(doc) {
  return {
    title: translateText(doc.title),
    date: translateDateLine(doc.date),
    org: translateText(doc.org),
    sections: doc.sections.map((section) => ({
      ...section,
      category: translateCategory(section.category),
      items: section.items.map((item) => ({
        ...item,
        title: translateText(item.title),
        details: item.details.map((detail) => translateText(detail)),
        tables: item.tables.map((table) => ({
          rows: table.rows.map((row) =>
            row.map((cell) => ({
              ...normalizeTableCellData(cell),
              text: translateText(getTableCellText(cell))
            }))
          )
        }))
      }))
    }))
  };
}

function translateCategory(category) {
  const inside = category.replace(/^\[|\]$/g, "");
  return `[${translateText(inside)}]`;
}

function translateText(text) {
  const detailBullet = hasDetailBullet(text);
  let translated = cleanMultilineText(stripDetailBullet(text));
  if (!translated) {
    return "";
  }

  translated = applySentenceLevelTranslation(translated);
  translated = applyGuideBasedTranslation(translated);
  const protectedTranslation = protectDictionaryTerms(translated, [...state.names, ...state.glossary]);
  translated = restoreProtectedTerms(protectedTranslation.text, protectedTranslation.tokens);

  for (const [source, target] of Object.entries(labelTranslations)) {
    translated = replaceEvery(translated, source, target);
  }

  for (const [source, target] of Object.entries(phraseTranslations)) {
    translated = replaceEvery(translated, source, target);
  }

  translated = finalizeEnglishText(translated);
  return detailBullet ? markDetailBullet(translated) : translated;
}

function finalizeEnglishText(text) {
  return cleanMultilineText(text)
    .replace(/([0-9]{4}-[0-9]{2}-[0-9]{2})\((.)\)/g, (_, date, weekday) => `${date}(${weekdayMap[weekday] || weekday})`)
    .replace(/([0-9]{4})년/g, "$1 ")
    .replace(/([0-9]{1,2})월/g, (_, value) => `${value} month`)
    .replace(/([0-9]{1,2})일/g, (_, value) => `${value} day`)
    .replace(/([0-9]{1,2})명/g, (_, value) => `${value} people`)
    .replace(/약\s*/g, "Approx. ")
    .replace(/명/g, " people")
    .replace(/조직명/g, "Organization")
    .replace(/부서명/g, "Department")
    .replace(/카테고리/g, "Category")
    .replace(/과제/g, "Agenda")
    .replace(/과제명/g, "Agenda")
    .replace(/내용/g, "Details")
    .replace(/후속 면담/g, "follow-up meeting")
    .replace(/추후 공유/g, "to be shared later")
    .replace(/논의 예정/g, "scheduled for discussion")
    .replace(/사전 기획/g, "preliminary planning")
    .replace(/정기 성과 리뷰/g, "regular performance review")
    .replace(/커뮤니케이션/g, "communication")
    .replace(/[^\S\n]+/g, " ")
    .trim();
}

function translateDateLine(text) {
  return translateText(text)
    .replace(/(\d{4}) (\d{1,2}) month (\d{1,2}) day/g, (_, y, m, d) => `${y}-${pad(m)}-${pad(d)}`)
    .replace(/month/g, "")
    .replace(/day/g, "")
    .replace(/\s{2,}/g, " ")
    .trim();
}

function replaceFromDictionary(text, entries) {
  const sorted = [...entries].sort((a, b) => b[0].length - a[0].length);
  let output = text;
  for (const [source, target] of sorted) {
    output = replaceEvery(output, source, target);
  }
  return output;
}

function protectDictionaryTerms(text, entries) {
  const sorted = [...entries].sort((a, b) => b[0].length - a[0].length);
  let output = String(text || "");
  const tokens = [];

  for (const [source, target] of sorted) {
    if (!source || !target || !output.includes(source)) {
      continue;
    }
    const token = `MYSUNIXXTERM${tokens.length}XX`;
    output = replaceEvery(output, source, token);
    tokens.push([token, target]);
  }

  return { text: output, tokens };
}

function restoreProtectedTerms(text, tokens) {
  let output = String(text || "");
  for (const [token, target, index] of tokens.map((entry, index) => [...entry, index])) {
    const escapedToken = token.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    output = output.replace(new RegExp(escapedToken, "g"), target);

    // Handle degraded token variants from browser translation/casing changes.
    const legacyPatterns = [
      new RegExp(`__mysuni[_\\s-]*term[_\\s-]*${index}__`, "gi"),
      new RegExp(`mysuni[_\\s-]*term[_\\s-]*${index}`, "gi"),
      new RegExp(`mysunixxterm${index}xx`, "gi")
    ];

    legacyPatterns.forEach((pattern) => {
      output = output.replace(pattern, target);
    });
  }
  return output;
}

function containsHangul(text) {
  return /[가-힣]/.test(text);
}

function applySentenceLevelTranslation(text) {
  let output = text;

  for (const [source, target] of sentenceTranslations) {
    output = replaceEvery(output, source, target);
  }

  output = output
    .replace(/^\[변경\]\s*/g, "[Change] ")
    .replace(/^\[중요\]\s*/g, "[Important] ")
    .replace(/^일시:\s*/i, "Date & Time: ")
    .replace(/^일자:\s*/i, "Date: ")
    .replace(/^부서:\s*/i, "Department: ")
    .replace(/^조직명:\s*/i, "Organization: ")
    .replace(/^목적:\s*/i, "Purpose: ")
    .replace(/^대상:\s*/i, "Target: ")
    .replace(/^장소:\s*/i, "Location: ")
    .replace(/^비고:\s*/i, "Note: ")
    .replace(/^기간:\s*/i, "Period: ")
    .replace(/^변경 전:\s*/i, "Previous: ")
    .replace(/^변경 후:\s*/i, "New: ")
    .replace(/^추가 조사:\s*/i, "Additional survey: ")
    .replace(/^응답률\s*/i, "response rate ")
    .replace(/종일\+만찬/g, "full day + dinner")
    .replace(/추후 공유/g, "will be shared later")
    .replace(/논의 예정/g, "scheduled for discussion")
    .replace(/구성원 대상 Survey 완료/g, "Employee survey completed")
    .replace(/국내 주요그룹 대상 '?AI 활용 현황'?\s*설문 진행 중/g, '“Current Status of AI Utilization” survey at major domestic groups, in progress')
    .replace(/([0-9]+)명 응답/g, (_, count) => `Responses: ${count}`)
    .replace(/([0-9.]+)%/g, (_, rate) => `${rate}%`);

  return output;
}

function createEmptySection(category = "[카테고리]") {
  return {
    id: createId(),
    category,
    items: [createEmptyItem()]
  };
}

function createEmptyItem(title = "과제명") {
  return {
    id: createId(),
    title,
    details: [markDetailBullet("내용")],
    tables: [],
    images: []
  };
}

function createDefaultTable() {
  return {
    rows: DEFAULT_TABLE_TEMPLATE.rows.map((row) => row.map((cell) => ({ text: cell, colSpan: 1 })))
  };
}

function isLikelyTitle(line, currentItem) {
  if (!currentItem) {
    return true;
  }

  if (/^(date|department|speaker|topic|purpose|target|location|note)\s*:/i.test(line)) {
    return false;
  }

  if (/^[•\-]/.test(line)) {
    return false;
  }

  return line.length <= 80;
}

function normalizeCell(value) {
  return cleanText(String(value || ""));
}

function buildDictionary(entries) {
  return entries
    .map(([source, target]) => [cleanText(source), cleanText(target)])
    .filter(([source, target]) => source && target);
}

function formatDetailForWord(detail) {
  // "–" 불릿은 template numbering(numId=11, ilvl=2)이 자동으로 추가
  // 이 함수는 순수 텍스트만 반환 (불릿 접두사 없음)
  return stripDetailBullet(detail);
}

function appendDetailLine(item, line) {
  const text = cleanMultilineText(line);
  if (!text) {
    return;
  }

  if (!item.details.length) {
    item.details.push(text);
    return;
  }

  item.details[item.details.length - 1] = `${item.details[item.details.length - 1]}\n${text}`;
}

function applyGuideBasedTranslation(text) {
  let output = text;

  const ruleSet = [
    ["작성 가이드", "Writing Guide"],
    ["영문 참고 사전", "English Reference Glossary"],
    ["구성원 영문이름", "Employee English Name"],
    ["조직별", "by organization"],
    ["운영 방향", "operation direction"],
    ["핵심 추진 과제", "key initiatives"],
    ["평가 결과 안내", "communication of evaluation results"],
    ["평가 피드백", "evaluation feedback"],
    ["별도 일정", "separate schedule"],
    ["일정 조정", "schedule adjustment"],
    ["변경 전", "Previous"],
    ["변경 후", "New"],
    ["응답률", "response rate"],
    ["중요", "Important"],
    ["기간", "Period"],
    ["사전", "preliminary"],
    ["예정", "planned"]
  ];

  for (const [source, target] of ruleSet) {
    output = replaceEvery(output, source, target);
  }

  // Admin에서 수정되는 가이드 문구를 기반으로 자주 쓰는 한글 키워드를 영문화 힌트로 활용
  state.guideItems.forEach((item) => {
    if (item.includes("카테고리")) {
      output = replaceEvery(output, "카테고리", "Category");
    }
    if (item.includes("과제")) {
      output = replaceEvery(output, "과제", "Agenda");
    }
    if (item.includes("내용")) {
      output = replaceEvery(output, "내용", "Details");
    }
    if (item.includes("조직명")) {
      output = replaceEvery(output, "조직명", "Organization");
    }
  });

  return output;
}

function parseTabSeparatedPairs(text) {
  return buildDictionary(
    text
      .split(/\r?\n/)
      .map((line) => line.split("\t"))
      .filter((parts) => parts.length >= 2)
      .map((parts) => [parts[0], parts.slice(1).join("\t")])
  );
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

// ── 번호 매기기(numbering) 파싱 헬퍼 ──────────────────────────────────────

function parseWordNumbering(numberingXml) {
  if (!numberingXml) return null;
  const parser = new DOMParser();
  const doc = parser.parseFromString(numberingXml, "application/xml");

  // abstractNum 파싱: abstractNumId → { [ilvl]: { numFmt, start, lvlText } }
  const abstractNums = {};
  const abstractNumEls = [...doc.getElementsByTagNameNS(WORD_NS, "abstractNum")];
  for (const abstractNumEl of abstractNumEls) {
    const id = Number(
      abstractNumEl.getAttributeNS(WORD_NS, "abstractNumId") ||
      abstractNumEl.getAttribute("w:abstractNumId") || "0"
    );
    const levels = {};
    const lvlEls = [...abstractNumEl.getElementsByTagNameNS(WORD_NS, "lvl")];
    for (const lvlEl of lvlEls) {
      const ilvl = Number(
        lvlEl.getAttributeNS(WORD_NS, "ilvl") ||
        lvlEl.getAttribute("w:ilvl") || "0"
      );
      const startEl = lvlEl.getElementsByTagNameNS(WORD_NS, "start")[0];
      const numFmtEl = lvlEl.getElementsByTagNameNS(WORD_NS, "numFmt")[0];
      const lvlTextEl = lvlEl.getElementsByTagNameNS(WORD_NS, "lvlText")[0];
      // pStyle: numbering 레벨에 연결된 단락 스타일 (없는 경우 null)
      const lvlPStyleEl = lvlEl.getElementsByTagNameNS(WORD_NS, "pStyle")[0];
      levels[ilvl] = {
        numFmt: numFmtEl?.getAttributeNS(WORD_NS, "val") || numFmtEl?.getAttribute("w:val") || "decimal",
        start: Number(startEl?.getAttributeNS(WORD_NS, "val") || startEl?.getAttribute("w:val") || "1"),
        lvlText: lvlTextEl?.getAttributeNS(WORD_NS, "val") ?? lvlTextEl?.getAttribute("w:val") ?? "%1.",
        pStyle: lvlPStyleEl?.getAttributeNS(WORD_NS, "val") || lvlPStyleEl?.getAttribute("w:val") || null
      };
    }
    abstractNums[id] = levels;
  }

  // num 파싱: numId → { abstractNumId, levelOverrides }
  const numMap = {};
  const numEls = [...doc.getElementsByTagNameNS(WORD_NS, "num")];
  for (const numEl of numEls) {
    const numId = Number(
      numEl.getAttributeNS(WORD_NS, "numId") ||
      numEl.getAttribute("w:numId") || "0"
    );
    const abstractNumIdEl = numEl.getElementsByTagNameNS(WORD_NS, "abstractNumId")[0];
    const abstractNumId = Number(
      abstractNumIdEl?.getAttributeNS(WORD_NS, "val") ||
      abstractNumIdEl?.getAttribute("w:val") || "-1"
    );
    const levelOverrides = {};
    const overrideEls = [...numEl.getElementsByTagNameNS(WORD_NS, "lvlOverride")];
    for (const overrideEl of overrideEls) {
      const ilvl = Number(
        overrideEl.getAttributeNS(WORD_NS, "ilvl") ||
        overrideEl.getAttribute("w:ilvl") || "0"
      );
      const startOverrideEl = overrideEl.getElementsByTagNameNS(WORD_NS, "startOverride")[0];
      if (startOverrideEl) {
        levelOverrides[ilvl] = Number(
          startOverrideEl.getAttributeNS(WORD_NS, "val") ||
          startOverrideEl.getAttribute("w:val") || "1"
        );
      }
    }
    numMap[numId] = { abstractNumId, levelOverrides };
  }

  return { abstractNums, numMap };
}

function getNumberingInfo(numberingData, numId, ilvl) {
  if (!numberingData || numId <= 0) return null;
  const numInfo = numberingData.numMap[numId];
  if (!numInfo || numInfo.abstractNumId < 0) return null;
  const levels = numberingData.abstractNums[numInfo.abstractNumId];
  if (!levels) return null;
  const levelInfo = levels[ilvl];
  if (!levelInfo) return null;
  const start = numInfo.levelOverrides[ilvl] ?? levelInfo.start;
  return {
    abstractNumId: numInfo.abstractNumId,
    numFmt: levelInfo.numFmt,
    start,
    lvlText: levelInfo.lvlText
  };
}

// pStyle 없는 단락의 스타일을 numbering 체인으로 추론
// 1순위: abstractNum 레벨에 연결된 pStyle
// 2순위: ilvl 기반 mySUNI 관례 (ilvl=1 → 과제 "a")
function getStyleFromNumbering(numberingData, numId, ilvl) {
  if (!numberingData || !numId || numId <= 0) return null;
  const numInfo = numberingData.numMap?.[numId];
  if (!numInfo || numInfo.abstractNumId < 0) return null;
  const levels = numberingData.abstractNums?.[numInfo.abstractNumId];
  const levelInfo = levels?.[ilvl];
  // numbering 레벨에 pStyle이 연결돼 있으면 그대로 사용
  if (levelInfo?.pStyle) return levelInfo.pStyle;
  // pStyle 없는 ilvl=1 → 과제("a") 로 추론
  // (SKMS실천 등 numPr만으로 서식이 적용된 과제 단락 처리)
  if (ilvl === 1) return "a";
  return null;
}

function toRomanNumeral(n) {
  const vals = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1];
  const syms = ["m", "cm", "d", "cd", "c", "xc", "l", "xl", "x", "ix", "v", "iv", "i"];
  let result = "";
  let num = n;
  for (let i = 0; i < vals.length; i++) {
    while (num >= vals[i]) { result += syms[i]; num -= vals[i]; }
  }
  return result;
}

function formatListNumber(numFmt, n, lvlText) {
  let numStr;
  switch (numFmt) {
    case "decimal": numStr = String(n); break;
    case "decimalZero": numStr = n < 10 ? "0" + n : String(n); break;
    case "upperLetter": numStr = String.fromCharCode(64 + ((n - 1) % 26 + 1)); break;
    case "lowerLetter": numStr = String.fromCharCode(96 + ((n - 1) % 26 + 1)); break;
    case "upperRoman": numStr = toRomanNumeral(n).toUpperCase(); break;
    case "lowerRoman": numStr = toRomanNumeral(n); break;
    case "koreanCounting":
    case "ideographKorean": numStr = String(n); break; // 단순화
    default: return null; // bullet 등 비수치 → 추가 안 함
  }
  if (lvlText) {
    // %1, %2, ... → numStr 치환
    return lvlText.replace(/%\d+/g, numStr).trim();
  }
  return numStr;
}

// ─────────────────────────────────────────────────────────────────────────────

function parseWordRels(relsXml) {
  if (!relsXml) {
    return {};
  }
  const parser = new DOMParser();
  const doc = parser.parseFromString(relsXml, "application/xml");
  const rels = {};
  for (const rel of [...doc.querySelectorAll("Relationship")]) {
    const id = rel.getAttribute("Id");
    const type = rel.getAttribute("Type") || "";
    const target = rel.getAttribute("Target") || "";
    if (id && target && type.includes("image")) {
      rels[id] = target;
    }
  }
  return rels;
}

function extractDrawingRIds(paragraphNode) {
  return extractDrawingInfos(paragraphNode).map((info) => info.rId);
}

// DrawingML / VML drawing에서 rId + 원본 크기(cx, cy: EMU 단위)를 함께 추출
function extractDrawingInfos(node) {
  const infos = [];

  // <wp:inline> / <wp:anchor> 기준으로 extent + blip + 정렬 + 간격 연결
  function processInlineOrAnchor(inlineNode) {
    let cx = 0, cy = 0, align = null;
    const isAnchor = inlineNode.localName === "anchor";
    // distT/distB: 이미지와 텍스트 사이 간격 (EMU)
    const distT = Number(inlineNode.getAttribute("distT") || "0");
    const distB = Number(inlineNode.getAttribute("distB") || "0");

    for (const child of [...inlineNode.childNodes]) {
      if (child.nodeType !== ELEMENT_NODE) continue;
      // <wp:extent cx="..." cy="..."> 에서 크기 추출
      if (child.localName === "extent") {
        cx = Number(child.getAttribute("cx") || "0");
        cy = Number(child.getAttribute("cy") || "0");
      }
      // anchor 전용: <wp:positionH><wp:align>center</wp:align></wp:positionH>
      if (isAnchor && child.localName === "positionH") {
        const alignEl = [...child.childNodes].find(
          (n) => n.nodeType === ELEMENT_NODE && n.localName === "align"
        );
        if (alignEl) align = alignEl.textContent.trim() || null;
      }
    }
    // blip rId 탐색 (그래픽 구조 안에 중첩)
    function findBlip(n) {
      for (const c of [...n.childNodes]) {
        if (c.nodeType !== ELEMENT_NODE) continue;
        if (c.localName === "blip") {
          const rId = c.getAttributeNS(REL_NS, "embed") || c.getAttribute("r:embed");
          if (rId) infos.push({ rId, cx, cy, align, distT, distB });
          return;
        }
        findBlip(c);
      }
    }
    findBlip(inlineNode);
  }

  function traverse(n) {
    for (const child of [...n.childNodes]) {
      if (child.nodeType !== ELEMENT_NODE) continue;
      if (child.localName === "drawing") {
        // <wp:inline> 또는 <wp:anchor> 자식 처리
        for (const gc of [...child.childNodes]) {
          if (gc.nodeType !== ELEMENT_NODE) continue;
          if (gc.localName === "inline" || gc.localName === "anchor") {
            processInlineOrAnchor(gc);
          }
        }
        continue;
      }
      // VML: <v:imagedata r:id="rIdN"> — 크기 정보 없으므로 0 저장
      if (child.localName === "imagedata") {
        const rId = child.getAttributeNS(REL_NS, "id") || child.getAttribute("r:id");
        if (rId) infos.push({ rId, cx: 0, cy: 0 });
        continue;
      }
      traverse(child);
    }
  }

  traverse(node);
  return infos;
}

function getImageMimeType(path) {
  const ext = (path || "").split(".").pop()?.toLowerCase() || "";
  const mimes = {
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    gif: "image/gif",
    webp: "image/webp",
    svg: "image/svg+xml",
    bmp: "image/bmp",
    tiff: "image/tiff",
    tif: "image/tiff"
  };
  // wmf/emf는 브라우저에서 렌더링 불가 → null 반환하여 이미지캐시에서 제외
  return mimes[ext] || null;
}

function buildTemplateBody(xml, body, docSource, options = {}) {
  const lang = options.lang || state.view || "ko";
  if (options.includeHeader !== false) {
    body.appendChild(createStyledParagraph(xml, "ad", [docSource.title || "mySUNI Weekly"]));
    body.appendChild(createEmptyParagraph(xml));
    body.appendChild(createMetaParagraph(xml, lang === "ko" ? "일자: " : "Date: ", docSource.date || "-"));
    body.appendChild(createMetaParagraph(xml, lang === "ko" ? "부서: " : "Department: ", docSource.org || "-"));
    body.appendChild(createEmptyParagraph(xml, "나눔스퀘어_ac Bold"));
  }

  docSource.sections.forEach((section, sectionIndex) => {
    body.appendChild(createStyledParagraph(xml, "a", [section.category]));

    section.items.forEach((item, itemIndex) => {
      body.appendChild(createStyledParagraph(xml, "a0", [item.title]));
      item.details.forEach((detail) => {
        // bullet 있는 내용: template numId=11/ilvl=2 의 "–" 불릿 사용
        // bullet 없는 내용: numbering 억제하여 들여쓰기만 유지
        body.appendChild(createStyledParagraph(xml, "a1", [formatDetailForWord(detail)], { suppressNumbering: !hasDetailBullet(detail) }));
      });

      item.tables.forEach((table) => {
        if (table.rows?.length) {
          body.appendChild(createTable(xml, table, options.imageDataMap || null));
        }
      });

      // 이미지 단락 (imageDataMap에 등록된 경우만 출력, 원문 정렬 적용)
      (item.images || []).forEach((image) => {
        const imgData = options.imageDataMap?.get(image.dataUrl);
        if (imgData) {
          body.appendChild(createImageParagraph(xml, imgData.rId, imgData.cx, imgData.cy, image.align || null));
        }
      });

      const isLastItemInSection = itemIndex === section.items.length - 1;
      const isLastSection = sectionIndex === docSource.sections.length - 1;
      if (!isLastItemInSection || !isLastSection) {
        body.appendChild(createEmptyParagraph(xml));
      }
    });
  });
}

function createStyledParagraph(xml, styleId, texts, options = {}) {
  const paragraph = createElement(xml, "p");
  const pPr = createElement(xml, "pPr");
  const pStyle = createElement(xml, "pStyle");
  pStyle.setAttributeNS(WORD_NS, "w:val", styleId);
  pPr.appendChild(pStyle);
  // 목록 서식 억제: 단락 스타일에 numPr가 정의돼 있어도 강제로 비활성화
  if (options.suppressNumbering) {
    const numPr = createElement(xml, "numPr");
    const ilvl = createElement(xml, "ilvl");
    ilvl.setAttributeNS(WORD_NS, "w:val", "0");
    const numId = createElement(xml, "numId");
    numId.setAttributeNS(WORD_NS, "w:val", "0");
    numPr.appendChild(ilvl);
    numPr.appendChild(numId);
    pPr.appendChild(numPr);
  }
  // pPr/rPr — 표준 양식과 동일한 단락 서식 지시자
  {
    const pRpr = createElement(xml, "rPr");
    if (styleId === "ad") {
      // 타이틀: sz/szCs만 override (나눔스퀘어_ac ExtraBold, 36/40)
      const szEl = createElement(xml, "sz"); szEl.setAttributeNS(WORD_NS, "w:val", "36");
      const szCsEl = createElement(xml, "szCs"); szCsEl.setAttributeNS(WORD_NS, "w:val", "40");
      pRpr.append(szEl, szCsEl);
    } else {
      // 나머지: rStyle="ae" + b/bCs/i/iCs/spacing 초기화 (표준 양식 구조와 동일)
      const rStyleEl = createElement(xml, "rStyle"); rStyleEl.setAttributeNS(WORD_NS, "w:val", "ae");
      const bEl = createElement(xml, "b"); bEl.setAttributeNS(WORD_NS, "w:val", "0");
      const bCsEl = createElement(xml, "bCs"); bCsEl.setAttributeNS(WORD_NS, "w:val", "0");
      const iEl = createElement(xml, "i"); iEl.setAttributeNS(WORD_NS, "w:val", "0");
      const iCsEl = createElement(xml, "iCs"); iCsEl.setAttributeNS(WORD_NS, "w:val", "0");
      const spacingEl = createElement(xml, "spacing"); spacingEl.setAttributeNS(WORD_NS, "w:val", "0");
      pRpr.append(rStyleEl, bEl, bCsEl, iEl, iCsEl, spacingEl);
    }
    pPr.appendChild(pRpr);
  }
  paragraph.appendChild(pPr);

  texts.forEach((text) => {
    const parts = String(text || "").split("\n");
    parts.forEach((part, index) => {
      const run = createElement(xml, "r");
      // run rPr — 표준 양식과 동일한 런 서식
      {
        const rPr = createElement(xml, "rPr");
        if (styleId === "ad") {
          const rFonts = createElement(xml, "rFonts"); rFonts.setAttributeNS(WORD_NS, "w:hint", "eastAsia");
          const szEl = createElement(xml, "sz"); szEl.setAttributeNS(WORD_NS, "w:val", "36");
          const szCsEl = createElement(xml, "szCs"); szCsEl.setAttributeNS(WORD_NS, "w:val", "40");
          rPr.append(rFonts, szEl, szCsEl);
        } else {
          const rStyleEl = createElement(xml, "rStyle"); rStyleEl.setAttributeNS(WORD_NS, "w:val", "ae");
          const rFonts = createElement(xml, "rFonts"); rFonts.setAttributeNS(WORD_NS, "w:hint", "eastAsia");
          const bEl = createElement(xml, "b"); bEl.setAttributeNS(WORD_NS, "w:val", "0");
          const bCsEl = createElement(xml, "bCs"); bCsEl.setAttributeNS(WORD_NS, "w:val", "0");
          const iEl = createElement(xml, "i"); iEl.setAttributeNS(WORD_NS, "w:val", "0");
          const iCsEl = createElement(xml, "iCs"); iCsEl.setAttributeNS(WORD_NS, "w:val", "0");
          const spacingEl = createElement(xml, "spacing"); spacingEl.setAttributeNS(WORD_NS, "w:val", "0");
          rPr.append(rStyleEl, rFonts, bEl, bCsEl, iEl, iCsEl, spacingEl);
        }
        run.appendChild(rPr);
      }
      const textNode = createElement(xml, "t");
      if (/^\s|\s$/.test(part)) {
        textNode.setAttributeNS(XML_NS, "xml:space", "preserve");
      }
      textNode.textContent = part;
      run.appendChild(textNode);
      paragraph.appendChild(run);
      if (index < parts.length - 1) {
        const breakRun = createElement(xml, "r");
        breakRun.appendChild(createElement(xml, "br"));
        paragraph.appendChild(breakRun);
      }
    });
  });

  return paragraph;
}

function createMetaParagraph(xml, label, value) {
  // createStyledParagraph이 pPr/rPr 및 run rPr를 자동으로 추가하므로 직접 처리 불필요
  return createStyledParagraph(xml, "af", [label, value]);
}

function createEmptyParagraph(xml, fontName = "") {
  const paragraph = createElement(xml, "p");
  const pPr = createElement(xml, "pPr");
  const rPr = createElement(xml, "rPr");
  const sz = createElement(xml, "sz");
  sz.setAttributeNS(WORD_NS, "w:val", "14");
  const szCs = createElement(xml, "szCs");
  szCs.setAttributeNS(WORD_NS, "w:val", "14");
  if (fontName) {
    const fonts = createElement(xml, "rFonts");
    fonts.setAttributeNS(WORD_NS, "w:ascii", fontName);
    fonts.setAttributeNS(WORD_NS, "w:eastAsia", fontName);
    fonts.setAttributeNS(WORD_NS, "w:hAnsi", fontName);
    rPr.appendChild(fonts);
  }
  rPr.appendChild(sz);
  rPr.appendChild(szCs);
  pPr.appendChild(rPr);
  paragraph.appendChild(pPr);
  return paragraph;
}

function createPageBreakParagraph(xml) {
  const paragraph = createElement(xml, "p");
  const run = createElement(xml, "r");
  const breakEl = createElement(xml, "br");
  breakEl.setAttributeNS(WORD_NS, "w:type", "page");
  run.appendChild(breakEl);
  paragraph.appendChild(run);
  return paragraph;
}

function createTable(xml, table, imageDataMap = null) {
  const totalColumns = getTableColumnCount(table);
  const widths = resolveTableWidths(totalColumns);
  const tbl = createElement(xml, "tbl");
  const tblPr = createElement(xml, "tblPr");
  const tblStyle = createElement(xml, "tblStyle");
  tblStyle.setAttributeNS(WORD_NS, "w:val", "af4");
  const tblW = createElement(xml, "tblW");
  tblW.setAttributeNS(WORD_NS, "w:w", "8500");
  tblW.setAttributeNS(WORD_NS, "w:type", "dxa");
  const tblInd = createElement(xml, "tblInd");
  tblInd.setAttributeNS(WORD_NS, "w:w", "993");
  tblInd.setAttributeNS(WORD_NS, "w:type", "dxa");
  const tblLook = createElement(xml, "tblLook");
  tblLook.setAttributeNS(WORD_NS, "w:val", "04A0");
  tblLook.setAttributeNS(WORD_NS, "w:firstRow", "1");
  tblLook.setAttributeNS(WORD_NS, "w:lastRow", "0");
  tblLook.setAttributeNS(WORD_NS, "w:firstColumn", "1");
  tblLook.setAttributeNS(WORD_NS, "w:lastColumn", "0");
  tblLook.setAttributeNS(WORD_NS, "w:noHBand", "0");
  tblLook.setAttributeNS(WORD_NS, "w:noVBand", "1");
  tblPr.append(tblStyle, tblW, tblInd, tblLook);
  tbl.appendChild(tblPr);

  const tblGrid = createElement(xml, "tblGrid");
  widths.forEach((width) => {
    const gridCol = createElement(xml, "gridCol");
    gridCol.setAttributeNS(WORD_NS, "w:w", String(width));
    tblGrid.appendChild(gridCol);
  });
  tbl.appendChild(tblGrid);

  table.rows.forEach((row, rowIndex) => {
    const tr = createElement(xml, "tr");
    let columnOffset = 0;
    row.forEach((cell, cellIndex) => {
      const colSpan = getTableCellColSpan(cell);
      const tc = createElement(xml, "tc");
      const tcPr = createElement(xml, "tcPr");
      const tcW = createElement(xml, "tcW");
      const width = widths.slice(columnOffset, columnOffset + colSpan).reduce((sum, value) => sum + value, 0) || widths[0];
      tcW.setAttributeNS(WORD_NS, "w:w", String(width));
      tcW.setAttributeNS(WORD_NS, "w:type", "dxa");
      tcPr.appendChild(tcW);

      if (colSpan > 1) {
        const gridSpan = createElement(xml, "gridSpan");
        gridSpan.setAttributeNS(WORD_NS, "w:val", String(colSpan));
        tcPr.appendChild(gridSpan);
      }

      if (rowIndex === 0) {
        const shd = createElement(xml, "shd");
        shd.setAttributeNS(WORD_NS, "w:val", "clear");
        shd.setAttributeNS(WORD_NS, "w:color", "auto");
        shd.setAttributeNS(WORD_NS, "w:fill", "F2F2F2");
        tcPr.appendChild(shd);
      }

      const vAlign = createElement(xml, "vAlign");
      vAlign.setAttributeNS(WORD_NS, "w:val", "center");
      tcPr.appendChild(vAlign);
      tc.appendChild(tcPr);

      const paragraph = createElement(xml, "p");
      const pPr = createElement(xml, "pPr");
      const pStyle = createElement(xml, "pStyle");
      pStyle.setAttributeNS(WORD_NS, "w:val", "a1");
      pPr.appendChild(pStyle);
      // 표 셀에서 단락 스타일의 목록 서식을 명시적으로 비활성화 (대시 중복 방지)
      const numPr = createElement(xml, "numPr");
      const ilvl = createElement(xml, "ilvl");
      ilvl.setAttributeNS(WORD_NS, "w:val", "0");
      const numId = createElement(xml, "numId");
      numId.setAttributeNS(WORD_NS, "w:val", "0"); // 0 = 목록 없음
      numPr.appendChild(ilvl);
      numPr.appendChild(numId);
      pPr.appendChild(numPr);
      // 정렬: 원문에서 파싱한 align 값 우선, 없으면 헤더행=center, 본문=left
      const cellNorm = normalizeTableCellData(cell);
      const cellAlignRaw = cellNorm.align || (rowIndex === 0 ? "center" : "left");
      const cellAlignWord = cellAlignRaw === "justify" ? "both" : cellAlignRaw;
      const jc = createElement(xml, "jc");
      jc.setAttributeNS(WORD_NS, "w:val", cellAlignWord);
      pPr.appendChild(jc);
      // 표 셀 단락 pPr/rPr — 표준 양식과 동일: rStyle="ae"+b/bCs/i/iCs/spacing=0+sz=20/szCs=20 (10pt)
      {
        const pCellRpr = createElement(xml, "rPr");
        const rStyleEl = createElement(xml, "rStyle"); rStyleEl.setAttributeNS(WORD_NS, "w:val", "ae");
        const bEl = createElement(xml, "b"); bEl.setAttributeNS(WORD_NS, "w:val", "0");
        const bCsEl = createElement(xml, "bCs"); bCsEl.setAttributeNS(WORD_NS, "w:val", "0");
        const iEl = createElement(xml, "i"); iEl.setAttributeNS(WORD_NS, "w:val", "0");
        const iCsEl = createElement(xml, "iCs"); iCsEl.setAttributeNS(WORD_NS, "w:val", "0");
        const spacingEl = createElement(xml, "spacing"); spacingEl.setAttributeNS(WORD_NS, "w:val", "0");
        const szEl = createElement(xml, "sz"); szEl.setAttributeNS(WORD_NS, "w:val", "20");
        const szCsEl = createElement(xml, "szCs"); szCsEl.setAttributeNS(WORD_NS, "w:val", "20");
        pCellRpr.append(rStyleEl, bEl, bCsEl, iEl, iCsEl, spacingEl, szEl, szCsEl);
        pPr.appendChild(pCellRpr);
      }
      paragraph.appendChild(pPr);

      // 셀 텍스트 — 줄바꿈을 <w:br/>로, 각 라인을 별도 run으로 처리
      const cellFont = rowIndex === 0 ? "나눔스퀘어_ac Bold" : "나눔스퀘어_ac";
      const cellDisplayText = getTableCellDisplayValue(cell, rowIndex === 0);
      const cellLines = cellDisplayText.split("\n");
      cellLines.forEach((line, lineIdx) => {
        const run = createElement(xml, "r");
        run.appendChild(createRunProperties(xml, { font: cellFont, size: "20" }));
        const textNode = createElement(xml, "t");
        if (/^\s|\s$/.test(line)) {
          textNode.setAttributeNS(XML_NS, "xml:space", "preserve");
        }
        textNode.textContent = line;
        run.appendChild(textNode);
        paragraph.appendChild(run);
        if (lineIdx < cellLines.length - 1) {
          const breakRun = createElement(xml, "r");
          breakRun.appendChild(createElement(xml, "br"));
          paragraph.appendChild(breakRun);
        }
      });
      tc.appendChild(paragraph);

      // 셀 내 이미지: 각 이미지마다 별도 단락으로 추가
      if (imageDataMap) {
        const MAX_CELL_EMU = 2500000; // 약 2.6인치 (셀 너비 상한)
        for (const img of (cellNorm.images || [])) {
          const imgData = imageDataMap.get(img.dataUrl);
          if (!imgData) continue;
          // 원본 cx/cy 우선 사용, 없으면 injectImagesToWordZip이 계산한 값 사용
          let imgCx = img.cx > 0 ? img.cx : imgData.cx;
          let imgCy = img.cy > 0 ? img.cy : imgData.cy;
          // 셀 너비 상한으로 비율 축소
          if (imgCx > MAX_CELL_EMU) {
            imgCy = imgCx > 0 ? Math.round((MAX_CELL_EMU * imgCy) / imgCx) : imgCy;
            imgCx = MAX_CELL_EMU;
          }
          tc.appendChild(createImageParagraph(xml, imgData.rId, imgCx, imgCy, img.align || null));
        }
      }

      tr.appendChild(tc);
      columnOffset += colSpan;
    });
    tbl.appendChild(tr);
  });

  return tbl;
}

// ── 이미지 Word 내보내기 ─────────────────────────────────────────────────────

/** 이미지의 자연 크기를 비동기로 읽어 EMU 단위로 반환 */
async function getImageNaturalSize(dataUrl) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve({ w: img.naturalWidth, h: img.naturalHeight });
    img.onerror = () => resolve({ w: 0, h: 0 });
    img.src = dataUrl;
  });
}

/** OOXML inline image 단락 생성 (align: "left"|"center"|"right"|null) */
function createImageParagraph(xml, rId, cx, cy, align = null) {
  const imgId = Math.floor(Math.random() * 90000) + 10000;
  const jcXml = align ? `<w:jc w:val="${align === "justify" ? "both" : align}"/>` : "";
  const drawingXmlStr = `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:pPr>
      <w:pStyle w:val="a1"/>
      <w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>
      ${jcXml}
      <w:rPr><w:rStyle w:val="ae"/><w:b w:val="0"/><w:bCs w:val="0"/><w:i w:val="0"/><w:iCs w:val="0"/><w:spacing w:val="0"/></w:rPr>
    </w:pPr>
    <w:r>
      <w:drawing>
        <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0">
          <wp:extent cx="${cx}" cy="${cy}"/>
          <wp:effectExtent l="0" t="0" r="0" b="0"/>
          <wp:docPr id="${imgId}" name="Image_${imgId}"/>
          <wp:cNvGraphicFramePr>
            <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
          </wp:cNvGraphicFramePr>
          <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:nvPicPr>
                  <pic:cNvPr id="${imgId}" name="Image_${imgId}"/>
                  <pic:cNvPicPr/>
                </pic:nvPicPr>
                <pic:blipFill>
                  <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${rId}"/>
                  <a:stretch><a:fillRect/></a:stretch>
                </pic:blipFill>
                <pic:spPr>
                  <a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
                  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                </pic:spPr>
              </pic:pic>
            </a:graphicData>
          </a:graphic>
        </wp:inline>
      </w:drawing>
    </w:r>
  </w:p>`;
  const drawingDoc = new DOMParser().parseFromString(drawingXmlStr, "application/xml");
  return xml.importNode(drawingDoc.documentElement, true);
}

/**
 * 문서 목록에서 이미지를 수집하여 Word ZIP에 주입하고,
 * dataUrl → { rId, cx, cy } 맵을 반환
 */
async function injectImagesToWordZip(zip, docs) {
  const MAX_W_EMU = 5000000; // ≈ 5.47 인치

  // 1단계: 고유 dataUrl을 삽입 순서 유지하며 수집 (동기)
  const seen = new Set();
  const orderedUrls = [];
  function collectUrl(dataUrl) {
    if (dataUrl && !seen.has(dataUrl)) {
      seen.add(dataUrl);
      orderedUrls.push(dataUrl);
    }
  }
  for (const doc of docs) {
    if (!doc) continue;
    for (const section of doc.sections || []) {
      for (const item of section.items || []) {
        for (const img of item.images || []) collectUrl(img?.dataUrl);
        for (const table of item.tables || []) {
          for (const row of table.rows || []) {
            for (const cell of row) {
              for (const img of (cell?.images || [])) collectUrl(img?.dataUrl);
            }
          }
        }
      }
    }
  }

  if (orderedUrls.length === 0) return new Map();

  // 2단계: 이미지 자연 크기를 병렬 조회
  const sizes = await Promise.all(orderedUrls.map((url) => getImageNaturalSize(url)));

  // 3단계: dataUrl → { rId, ext, cx, cy } 맵 구성
  const uniqueImages = new Map();
  let rIdCounter = 100;
  orderedUrls.forEach((dataUrl, i) => {
    const { w, h } = sizes[i];
    const mimeMatch = dataUrl.match(/^data:image\/([a-z]+);base64,/);
    const ext = mimeMatch ? (mimeMatch[1] === "jpeg" ? "jpg" : mimeMatch[1]) : "png";
    const cx = MAX_W_EMU;
    const cy = w > 0 ? Math.round((MAX_W_EMU * h) / w) : Math.round(MAX_W_EMU * 0.75);
    uniqueImages.set(dataUrl, { rId: `rId${rIdCounter++}`, ext, cx, cy });
  });

  if (uniqueImages.size === 0) return uniqueImages;

  // 이미지 파일을 word/media/ 에 추가
  for (const [dataUrl, { rId, ext }] of uniqueImages) {
    const base64Data = dataUrl.split(",")[1];
    zip.file(`word/media/img_${rId}.${ext}`, base64Data, { base64: true });
  }

  // word/_rels/document.xml.rels 에 관계 추가
  const relsXml = await zip.file("word/_rels/document.xml.rels").async("string");
  const relsDoc = new DOMParser().parseFromString(relsXml, "application/xml");
  const relsRoot = relsDoc.documentElement;
  const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
  for (const [, { rId, ext }] of uniqueImages) {
    const relEl = relsDoc.createElementNS(PKG_REL_NS, "Relationship");
    relEl.setAttribute("Id", rId);
    relEl.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
    relEl.setAttribute("Target", `media/img_${rId}.${ext}`);
    relsRoot.appendChild(relEl);
  }
  zip.file("word/_rels/document.xml.rels", new XMLSerializer().serializeToString(relsDoc));

  // [Content_Types].xml 에 이미지 확장자 추가 (없는 것만)
  const ctXml = await zip.file("[Content_Types].xml").async("string");
  const ctDoc = new DOMParser().parseFromString(ctXml, "application/xml");
  const ctRoot = ctDoc.documentElement;
  const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
  const existingExts = new Set([...ctRoot.getElementsByTagNameNS(CT_NS, "Default")].map((el) => el.getAttribute("Extension")));
  const mimeMap = { png: "image/png", jpg: "image/jpeg", gif: "image/gif", webp: "image/webp", bmp: "image/bmp", svg: "image/svg+xml" };
  for (const [, { ext }] of uniqueImages) {
    if (!existingExts.has(ext)) {
      const defEl = ctDoc.createElementNS(CT_NS, "Default");
      defEl.setAttribute("Extension", ext);
      defEl.setAttribute("ContentType", mimeMap[ext] || "image/png");
      ctRoot.appendChild(defEl);
      existingExts.add(ext);
    }
  }
  zip.file("[Content_Types].xml", new XMLSerializer().serializeToString(ctDoc));

  return uniqueImages; // Map<dataUrl, { rId, cx, cy }>
}

function createRunProperties(xml, options = {}) {
  const rPr = createElement(xml, "rPr");
  const rStyle = createElement(xml, "rStyle");
  rStyle.setAttributeNS(WORD_NS, "w:val", "ae");
  rPr.appendChild(rStyle);

  const fontName = options.font || "";
  if (fontName) {
    const fonts = createElement(xml, "rFonts");
    fonts.setAttributeNS(WORD_NS, "w:ascii", fontName);
    fonts.setAttributeNS(WORD_NS, "w:eastAsia", fontName);
    fonts.setAttributeNS(WORD_NS, "w:hAnsi", fontName);
    fonts.setAttributeNS(WORD_NS, "w:hint", "eastAsia");
    rPr.appendChild(fonts);
  } else {
    const fonts = createElement(xml, "rFonts");
    fonts.setAttributeNS(WORD_NS, "w:hint", "eastAsia");
    rPr.appendChild(fonts);
  }

  const b = createElement(xml, "b");
  b.setAttributeNS(WORD_NS, "w:val", "0");
  const bCs = createElement(xml, "bCs");
  bCs.setAttributeNS(WORD_NS, "w:val", "0");
  const i = createElement(xml, "i");
  i.setAttributeNS(WORD_NS, "w:val", "0");
  const iCs = createElement(xml, "iCs");
  iCs.setAttributeNS(WORD_NS, "w:val", "0");
  const spacing = createElement(xml, "spacing");
  spacing.setAttributeNS(WORD_NS, "w:val", "0");
  rPr.append(b, bCs, i, iCs, spacing);

  if (options.size) {
    const sz = createElement(xml, "sz");
    sz.setAttributeNS(WORD_NS, "w:val", options.size);
    const szCs = createElement(xml, "szCs");
    szCs.setAttributeNS(WORD_NS, "w:val", options.size);
    rPr.append(sz, szCs);
  }

  return rPr;
}

function resolveTableWidths(columnCount) {
  if (columnCount === 3) {
    return [2688, 3685, 2127];
  }
  const total = 8500;
  const base = Math.floor(total / Math.max(columnCount, 1));
  const widths = Array.from({ length: columnCount }, () => base);
  const remainder = total - base * columnCount;
  if (widths.length) {
    widths[widths.length - 1] += remainder;
  }
  return widths;
}

function getTableColumnCount(table) {
  if (!table?.rows?.length) {
    return 0;
  }
  return Math.max(
    ...table.rows.map((row) => row.reduce((sum, cell) => sum + getTableCellColSpan(cell), 0)),
    0
  );
}

function createElement(xml, tagName) {
  return xml.createElementNS(WORD_NS, `w:${tagName}`);
}

const XML_NS = "http://www.w3.org/XML/1998/namespace";

function findHeaderIndexes(rows, koHeader, enHeader, fallbackKo, fallbackEn) {
  for (const row of rows) {
    const normalized = row.map((cell) => normalizeCell(cell));
    const koIndex = normalized.indexOf(koHeader);
    const enIndex = normalized.indexOf(enHeader);
    if (koIndex >= 0 && enIndex >= 0) {
      return { koIndex, enIndex };
    }
  }

  return { koIndex: fallbackKo, enIndex: fallbackEn };
}

function syncInputsFromState() {
  if (els.titleInput) {
    els.titleInput.value = state.document.title;
  }
  if (els.dateInput) {
    els.dateInput.value = state.document.date;
  }
  if (els.datePicker) {
    els.datePicker.value = extractDateValue(state.document.date);
  }
  if (els.orgInput) {
    els.orgInput.value = state.document.org;
  }
  if (els.filenamePreview) {
    els.filenamePreview.value = state.document.filename || buildSuggestedFilename(state.document);
  }
}

function loadStoredPairs(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) {
      return fallback;
    }
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : fallback;
  } catch {
    return fallback;
  }
}

function loadStoredGuideItems() {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.guide);
    if (!raw) {
      return [...defaultGuideItems];
    }
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed) || !parsed.length) {
      return [...defaultGuideItems];
    }

    const isLegacyDefaultGuide =
      parsed.length === 6 &&
      parsed[0] === "상단 제목, 일자, 조직명은 고정 레이아웃으로 유지" &&
      parsed[5] === "권장 파일명: mySUNI Weekly_위클리날짜_조직명_v2.docx";

    // 구버전 기본 가이드(비구조화 flat 형식) 감지 → 신규 가이드로 자동 업그레이드
    const isOldDefaultGuide = parsed[0] === "파일명";

    return (isLegacyDefaultGuide || isOldDefaultGuide) ? [...defaultGuideItems] : parsed;
  } catch {
    return [...defaultGuideItems];
  }
}

function loadStoredAiConfig() {
  try {
    return {
      model: localStorage.getItem(STORAGE_KEYS.aiModel) || DEFAULT_AI_CONFIG.model
    };
  } catch {
    return { ...DEFAULT_AI_CONFIG };
  }
}

function loadStoredEditorWide() {
  try {
    return localStorage.getItem(STORAGE_KEYS.editorWide) === "1";
  } catch {
    return false;
  }
}

function loadStoredDocumentDraft() {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.documentDraft);
    if (!raw) {
      return cloneValue(emptyDocumentState);
    }
    const parsed = JSON.parse(raw);
    return sanitizeLoadedDocumentDraft(parsed);
  } catch {
    return cloneValue(emptyDocumentState);
  }
}

function sanitizeLoadedDocumentDraft(doc) {
  const safeDoc = doc && typeof doc === "object" ? doc : {};
  return {
    title: cleanText(safeDoc.title),
    date: cleanText(safeDoc.date),
    org: cleanText(safeDoc.org),
    filename: cleanText(safeDoc.filename),
    sourceFileName: cleanText(safeDoc.sourceFileName),
    sections: Array.isArray(safeDoc.sections)
      ? safeDoc.sections.map((section) => ({
          id: section.id || createId(),
          category: cleanText(section.category),
          items: Array.isArray(section.items)
            ? section.items.map((item) => ({
                id: item.id || createId(),
                title: cleanText(item.title),
                details: Array.isArray(item.details) ? item.details.map((detail) => String(detail || "")) : [],
                tables: Array.isArray(item.tables) ? item.tables.map((tbl) => ({
                  rows: Array.isArray(tbl.rows) ? tbl.rows.map((row) =>
                    Array.isArray(row) ? row.map(normalizeTableCellData) : []
                  ) : []
                })) : [],
                images: Array.isArray(item.images) ? item.images.filter((img) => img && typeof img.dataUrl === "string") : []
              }))
            : []
        }))
      : []
  };
}

function loadStoredArchives() {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.archives);
    if (!raw) {
      return [];
    }
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function loadStoredDocumentArchives() {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.documentArchives);
    if (!raw) {
      return [];
    }
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function scheduleDocumentDraftSave() {
  const signature = getDocumentSignature(state.document);
  if (signature === state.lastDraftSignature) {
    return;
  }

  state.lastDraftSignature = signature;
  state.draftStatus = "임시 저장 중...";

  if (state.draftSaveTimer) {
    clearTimeout(state.draftSaveTimer);
  }

  state.draftSaveTimer = window.setTimeout(() => {
    saveDocumentDraft(false);
    renderDraftStatus();
  }, 500);
}

function saveDocumentDraft(isManual) {
  if (state.draftSaveTimer) {
    clearTimeout(state.draftSaveTimer);
    state.draftSaveTimer = null;
  }

  localStorage.setItem(STORAGE_KEYS.documentDraft, JSON.stringify(state.document));
  state.lastDraftSignature = getDocumentSignature(state.document);
  state.draftStatus = isManual ? "수동 저장됨" : "자동 임시 저장됨";
}

function saveDocumentArchiveSnapshot() {
  const timestamp = new Date();
  const archive = {
    id: createId(),
    title: state.document.title || "제목",
    date: state.document.date || "",
    org: state.document.org || "",
    sectionCount: state.document.sections.length,
    savedAt: timestamp.toISOString(),
    savedAtLabel: timestamp.toLocaleString("ko-KR"),
    document: cloneValue(state.document)
  };

  state.documentArchives = [archive, ...state.documentArchives].slice(0, 50);
  localStorage.setItem(STORAGE_KEYS.documentArchives, JSON.stringify(state.documentArchives));
  renderDocumentArchiveTab();
}

function buildSuggestedFilename(doc) {
  const compactDate = String(doc?.date || "")
    .replace(/[^\d]/g, "")
    .slice(2, 8);
  const org = String(doc?.org || "조직명").replace(/\s+/g, "");
  const version = extractFilenameVersion(doc?.filename) || "2";
  return `mySUNI Weekly_${compactDate || "위클리날짜"}_${org || "조직명"}_v${version}.docx`;
}

function syncFilenameWithSuggestion(previousSuggested) {
  const currentFilename = cleanText(state.document.filename);
  if (!currentFilename || currentFilename === previousSuggested) {
    state.document.filename = buildSuggestedFilename(state.document);
  }
}

function extractFilenameVersion(filename) {
  const match = String(filename || "").match(/_v(\d+)(?=\.docx$|$)/i);
  return match ? match[1] : "";
}

function extractDateValue(value) {
  const match = String(value || "").match(/(\d{4}-\d{2}-\d{2})/);
  return match ? match[1] : "";
}

function formatPickedDate(value) {
  const dateValue = extractDateValue(value);
  if (!dateValue) {
    return "";
  }

  const date = new Date(`${dateValue}T00:00:00`);
  const weekdays = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  return `${dateValue}(${weekdays[date.getDay()]})`;
}

function cleanText(value) {
  return String(value || "")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function cleanMultilineText(value) {
  return normalizeWordSymbols(
    String(value || "")
    .replace(/\u00a0/g, " ")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .split("\n")
    .map((line) => line.replace(/[ \t]+/g, " ").trim())
    .join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim()
  );
}

function markDetailBullet(value) {
  return `${DETAIL_BULLET_TOKEN}${cleanMultilineText(value)}`;
}

function hasDetailBullet(value) {
  return String(value || "").startsWith(DETAIL_BULLET_TOKEN);
}

function stripDetailBullet(value) {
  return cleanMultilineText(String(value || "").replace(DETAIL_BULLET_TOKEN, ""));
}

function normalizeTableCellData(cell) {
  if (cell && typeof cell === "object" && !Array.isArray(cell)) {
    return {
      text: cleanMultilineText(cell.text),
      colSpan: Number.parseInt(cell.colSpan, 10) || 1,
      rowSpan: cell.rowSpan !== undefined ? (Number.parseInt(cell.rowSpan, 10) ?? 1) : 1,
      align: cell.align || null,
      images: Array.isArray(cell.images) ? cell.images.filter((img) => img?.dataUrl) : []
    };
  }

  return {
    text: cleanMultilineText(cell),
    colSpan: 1,
    rowSpan: 1,
    align: null,
    images: []
  };
}

function getTableCellText(cell) {
  return normalizeTableCellData(cell).text;
}

function getTableCellColSpan(cell) {
  return normalizeTableCellData(cell).colSpan;
}

function getTableCellRowSpan(cell) {
  return normalizeTableCellData(cell).rowSpan;
}

function getTableCellDisplayValue(value, isHeader = false) {
  const normalized = getTableCellText(value);
  if (normalized) {
    return normalized;
  }
  return "";
}

function extractNodeTextWithBreaks(node) {
  if (!node) {
    return "";
  }
  const clone = node.cloneNode(true);
  clone.querySelectorAll("br").forEach((br) => {
    br.replaceWith("\n");
  });
  return cleanMultilineText(clone.textContent || "");
}

function normalizeWordSymbols(value) {
  return String(value || "")
    .replace(/[\uf0e8]/g, "->")
    .replace(/[\uf0e0]/g, "←")
    .replace(/[\uf0df]/g, "↑")
    .replace(/[\uf0e1]/g, "↓")
    .replace(/[\uf0b7]/g, "•")
    .replace(/[\uf0a7]/g, "▪")
    .replace(/[\uf0a8]/g, "■")
    .replace(/[\uf0fc]/g, "✓")
    .replace(/[\uf0fd]/g, "☑")
    .replace(/[\uf0fe]/g, "☒");
}

function escapeHtml(value) {
  let output = String(value || "");
  output = replaceEvery(output, "&", "&amp;");
  output = replaceEvery(output, "<", "&lt;");
  output = replaceEvery(output, ">", "&gt;");
  output = replaceEvery(output, '"', "&quot;");
  return output;
}

function pad(value) {
  return String(value).padStart(2, "0");
}
