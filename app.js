(function () {
  "use strict";

  const root = typeof window !== "undefined" ? window : globalThis;
  const STORAGE_KEY = "excel-auto-tool-fields-v2";
  const DRAFT_DB_NAME = "excel-auto-tool-drafts";
  const DRAFT_DB_VERSION = 1;
  const DRAFT_STORE_NAME = "drafts";
  const DRAFT_VERSION = 1;
  const MAX_DRAFTS = 2;
  const DRAFT_SAVE_DELAY = 600;
  const DEFAULT_FIELDS = [
    { key: "install_location", label: "安装位置", group: "", type: "text", required: false, options: [] },
    { key: "model", label: "型号", group: "", type: "text", required: false, options: [] },
    { key: "image", label: "图片", group: "", type: "image", required: false, options: [] },
    { key: "finished_width", label: "宽", group: "成品规格(M)", type: "number", required: false, options: [] },
    { key: "finished_height", label: "高", group: "成品规格(M)", type: "number", required: false, options: [] },
    { key: "finished_ratio", label: "比例", group: "成品规格(M)", type: "text", required: false, options: [] },
    { key: "finished_style", label: "款式", group: "成品规格(M)", type: "text", required: false, options: [] },
    { key: "material_meters", label: "米数", group: "用料", type: "number", required: false, options: [] },
    { key: "material_unit_price", label: "元/米", group: "单价", type: "number", required: false, options: [] },
    { key: "amount", label: "金额", group: "", type: "number", required: false, options: [] },
    { key: "track_name", label: "品名", group: "轨道", type: "text", required: false, options: [] },
    { key: "track_quantity", label: "数量", group: "轨道", type: "number", required: false, options: [] },
    { key: "track_unit_price", label: "单价", group: "轨道", type: "number", required: false, options: [] },
    { key: "track_amount", label: "金额", group: "轨道", type: "number", required: false, options: [] },
    { key: "remark", label: "备注", group: "", type: "text", required: false, options: [] },
  ];
  const CUSTOM_ROW_TYPE_KEY = "custom_type";
  const CUSTOM_ROW_QUANTITY_KEY = "custom_quantity";
  const CUSTOM_ROW_UNIT_PRICE_KEY = "custom_unit_price";
  const CUSTOM_ROW_AMOUNT_KEY = "custom_amount";
  const CUSTOM_ROW_FIELDS = [
    { key: CUSTOM_ROW_TYPE_KEY, label: "类型", group: "", type: "text", required: false, options: [] },
    { key: CUSTOM_ROW_QUANTITY_KEY, label: "数量", group: "", type: "number", required: false, options: [] },
    { key: CUSTOM_ROW_UNIT_PRICE_KEY, label: "单价", group: "", type: "number", required: false, options: [] },
    { key: CUSTOM_ROW_AMOUNT_KEY, label: "金额", group: "", type: "number", required: false, options: [] },
  ];
  const FIELD_TYPES = ["text", "number", "date", "select", "image"];
  const TYPE_TEXT = {
    text: "文本",
    number: "数字",
    date: "日期",
    select: "下拉",
    image: "图片",
  };
  const MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  const XML_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const XML_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const ZIP_EPOCH = new Date("1980-01-01T00:00:00Z");
  const SUPPORTED_IMAGE_MIMES = ["image/png", "image/jpeg", "image/gif"];
  const NATURAL_FILL_ENDPOINT = "/api/agent-skills/doubao-excel-natural-fill/extract";
  const IMAGE_CELL_WIDTH_PX = 132;
  const IMAGE_CELL_HEIGHT_PX = 96;
  const IMAGE_CELL_PADDING_PX = 6;
  const EMU_PER_PIXEL = 9525;
  const CALCULATED_STYLE_ID = 5;
  const CUSTOM_TEXT_STYLE_ID = 8;
  const CUSTOM_DATE_STYLE_ID = 9;
  const CUSTOM_NUMBER_STYLE_ID = 10;
  const CUSTOM_CALCULATED_STYLE_ID = 11;
  const EXPORT_NOTICE_TEXT = "\u6e29\u99a8\u63d0\u793a\uff1a1\uff1a\u7a97\u5e18\u4e3a\u5b9a\u5236\u4ea7\u54c1\uff0c\u82e5\u65e0\u8d28\u91cf\u95ee\u9898\u4e0d\u4e88\u9000\u6362\u30022\uff1a\u9762\u6599\u5899\u7eb8\u8272\u5dee\u5141\u8bb8\u5728\u56fd\u5bb6\u89c4\u5b9a\u8272\u5dee\u5141\u8bb85%\u4e4b\u5185\u3002\uff08\u5ba3\u745e\u8f6f\u88c5 \u9648\u59d0\uff1a13688428383\uff09";

  let state = {
    documentName: "",
    depositAmount: "",
    sharedRemark: "",
    fields: [],
    rows: [],
    invalidCells: new Set(),
    invalidSharedFields: new Set(),
    invalidFields: new Set(),
    configOpen: false,
    drafts: [],
    activeDraftId: "",
    draftReady: false,
    draftStorageAvailable: false,
    naturalFillRunning: false,
  };

  let els = {};
  let statusTimer = 0;
  let draftDbPromise = null;
  let draftSaveTimer = 0;

  function cloneField(field) {
    const key = String(field.key || makeFieldKey("field"));
    const label = String(field.label || "");
    const inferredType = label.trim() === "图片" || key.toLowerCase() === "image" ? "image" : "text";
    const explicitType = FIELD_TYPES.includes(field.type) ? field.type : "";
    return {
      key,
      label,
      group: String(field.group || ""),
      type: explicitType === "text" && inferredType === "image" ? "image" : explicitType || inferredType,
      required: Boolean(field.required),
      options: Array.isArray(field.options)
        ? field.options.map(String).filter(Boolean)
        : String(field.options || "")
            .split(",")
            .map((item) => item.trim())
            .filter(Boolean),
    };
  }

  function normalizeFields(fields) {
    if (!Array.isArray(fields)) {
      return DEFAULT_FIELDS.map(cloneField);
    }

    const seen = new Set();
    const normalized = fields
      .map(cloneField)
      .map((field, index) => {
        let key = field.key || makeFieldKey(`field_${index + 1}`);
        while (seen.has(key)) {
          key = makeFieldKey(key);
        }
        seen.add(key);
        return { ...field, key };
      });

    return normalized.length ? normalized : DEFAULT_FIELDS.map(cloneField);
  }

  function loadFields() {
    try {
      const saved = root.localStorage ? root.localStorage.getItem(STORAGE_KEY) : null;
      return normalizeFields(saved ? JSON.parse(saved) : DEFAULT_FIELDS);
    } catch (error) {
      return DEFAULT_FIELDS.map(cloneField);
    }
  }

  function saveFields() {
    if (!root.localStorage) {
      return;
    }
    try {
      root.localStorage.setItem(STORAGE_KEY, JSON.stringify(state.fields));
    } catch (error) {
      // Field config persistence is only a convenience fallback.
    }
  }

  function canUseDraftStorage() {
    return Boolean(root.indexedDB);
  }

  function openDraftDb() {
    if (!canUseDraftStorage()) {
      return Promise.reject(new Error("IndexedDB is not available"));
    }

    if (draftDbPromise) {
      return draftDbPromise;
    }

    draftDbPromise = new Promise((resolve, reject) => {
      const request = root.indexedDB.open(DRAFT_DB_NAME, DRAFT_DB_VERSION);

      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(DRAFT_STORE_NAME)) {
          db.createObjectStore(DRAFT_STORE_NAME, { keyPath: "id" });
        }
      };
      request.onsuccess = () => {
        const db = request.result;
        db.onversionchange = () => db.close();
        resolve(db);
      };
      request.onerror = () => {
        draftDbPromise = null;
        reject(request.error || new Error("Unable to open draft storage"));
      };
    });

    return draftDbPromise;
  }

  function runDraftTransaction(mode, callback) {
    return openDraftDb().then((db) => {
      return new Promise((resolve, reject) => {
        const transaction = db.transaction(DRAFT_STORE_NAME, mode);
        const store = transaction.objectStore(DRAFT_STORE_NAME);
        let result;

        transaction.oncomplete = () => resolve(result);
        transaction.onerror = () => reject(transaction.error || new Error("Draft storage failed"));
        transaction.onabort = () => reject(transaction.error || new Error("Draft storage aborted"));

        try {
          result = callback(store);
        } catch (error) {
          reject(error);
        }
      });
    });
  }

  function readRequest(request) {
    return new Promise((resolve, reject) => {
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error || new Error("Draft request failed"));
    });
  }

  async function loadDraftsFromStorage() {
    const drafts = await runDraftTransaction("readonly", (store) => readRequest(store.getAll()));
    return sortDrafts(drafts);
  }

  function putDraftToStorage(draft) {
    return runDraftTransaction("readwrite", (store) => {
      store.put(draft);
    });
  }

  function deleteDraftFromStorage(id) {
    return runDraftTransaction("readwrite", (store) => {
      store.delete(id);
    });
  }

  function makeDraftId() {
    return `draft_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}`;
  }

  function normalizeDraftList(drafts) {
    return sortDrafts(drafts).slice(0, MAX_DRAFTS);
  }

  function sortDrafts(drafts) {
    return (Array.isArray(drafts) ? drafts : [])
      .map(normalizeDraft)
      .filter(Boolean)
      .sort((left, right) => right.updatedAt - left.updatedAt);
  }

  function normalizeDraft(draft) {
    if (!draft || typeof draft !== "object") {
      return null;
    }

    const fields = normalizeFields(draft.fields);
    const rows = Array.isArray(draft.rows)
      ? draft.rows.map((row) => normalizeAppRow(row, fields))
      : [];
    const documentName = String(draft.documentName || "");
    const depositAmount = normalizeNonNegativeNumberValue(draft.depositAmount || "");
    const updatedAt = Number.isFinite(Number(draft.updatedAt)) ? Number(draft.updatedAt) : Date.now();

    return {
      id: String(draft.id || makeDraftId()),
      name: draftDisplayName(draft.name || documentName),
      updatedAt,
      documentName,
      depositAmount,
      sharedRemark: String(draft.sharedRemark || ""),
      fields,
      rows,
      version: Number(draft.version) || DRAFT_VERSION,
    };
  }

  function draftDisplayName(value) {
    const text = String(value || "").trim();
    return text || "未命名表格";
  }

  function createDraftPayload() {
    ensureSharedRemarkState();
    const fields = normalizeFields(state.fields);
    const rows = Array.isArray(state.rows) ? state.rows.map((row) => normalizeAppRow(row, fields)) : [];
    const documentName = String(state.documentName || "");
    const depositAmount = normalizeNonNegativeNumberValue(state.depositAmount || "");
    const updatedAt = Date.now();

    return {
      id: state.activeDraftId || makeDraftId(),
      name: draftDisplayName(documentName),
      updatedAt,
      documentName,
      depositAmount,
      sharedRemark: String(state.sharedRemark || ""),
      fields,
      rows,
      version: DRAFT_VERSION,
    };
  }

  async function initializeDraftState() {
    state.fields = loadFields();
    state.rows = [createEmptyRow()];
    state.activeDraftId = makeDraftId();

    if (!canUseDraftStorage()) {
      state.draftStorageAvailable = false;
      state.draftReady = false;
      return;
    }

    try {
      state.draftStorageAvailable = true;
      state.drafts = normalizeDraftList(await loadDraftsFromStorage());
      if (state.drafts.length) {
        applyDraft(state.drafts[0]);
      }
      state.draftReady = true;
    } catch (error) {
      state.draftStorageAvailable = false;
      state.draftReady = false;
      state.drafts = [];
    }
  }

  function applyDraft(draft) {
    const normalizedDraft = normalizeDraft(draft);
    if (!normalizedDraft) {
      return false;
    }

    state.activeDraftId = normalizedDraft.id;
    state.documentName = normalizedDraft.documentName;
    state.depositAmount = normalizedDraft.depositAmount;
    state.sharedRemark = normalizedDraft.sharedRemark;
    state.fields = normalizedDraft.fields;
    state.rows = normalizedDraft.rows.length ? normalizedDraft.rows : [createEmptyRow()];
    clearValidationState();
    ensureSharedRemarkState();
    saveFields();
    return true;
  }

  function clearValidationState() {
    state.invalidCells.clear();
    state.invalidSharedFields.clear();
    state.invalidFields.clear();
  }

  function syncDocumentNameInput() {
    if (els.documentName && els.documentName.value !== state.documentName) {
      els.documentName.value = state.documentName;
    }
  }

  function syncDepositAmountInput() {
    if (els.depositAmount && els.depositAmount.value !== state.depositAmount) {
      els.depositAmount.value = state.depositAmount;
    }
  }

  function scheduleDraftSave() {
    if (!state.draftReady || !state.draftStorageAvailable) {
      return;
    }

    clearTimeout(draftSaveTimer);
    draftSaveTimer = setTimeout(() => {
      flushDraftSave({ showSaved: true });
    }, DRAFT_SAVE_DELAY);
  }

  async function flushDraftSave(options) {
    const shouldShowSaved = !options || options.showSaved !== false;
    clearTimeout(draftSaveTimer);

    if (!state.draftReady || !state.draftStorageAvailable) {
      return true;
    }

    try {
      const draft = createDraftPayload();
      state.activeDraftId = draft.id;
      await putDraftToStorage(draft);
      await pruneDrafts();
      if (shouldShowSaved) {
        setStatus("已自动保存", "success");
      }
      renderDraftPanel();
      return true;
    } catch (error) {
      setStatus("图片过大或浏览器空间不足，请减少图片后重试", "warning");
      return false;
    }
  }

  async function pruneDrafts() {
    const drafts = await loadDraftsFromStorage();
    const extras = drafts.slice(MAX_DRAFTS);
    await Promise.all(extras.map((draft) => deleteDraftFromStorage(draft.id)));
    state.drafts = drafts.slice(0, MAX_DRAFTS);
  }

  function startNewDraft() {
    state.activeDraftId = makeDraftId();
    state.documentName = "";
    state.depositAmount = "";
    state.sharedRemark = "";
    state.rows = [createEmptyRow()];
    clearValidationState();
    ensureSharedRemarkState();
    saveFields();
    render();
    flushDraftSave({ showSaved: false }).then(() => {
      setStatus("已创建新表格", "success");
    });
  }

  function makeFieldKey(label) {
    const base = String(label || "field")
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "")
      .slice(0, 32);
    return `${base || "field"}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 6)}`;
  }

  function createEmptyValues(fields) {
    return (fields || state.fields).reduce((row, field) => {
      row[field.key] = "";
      return row;
    }, {});
  }

  function createNormalRow() {
    return {
      kind: "normal",
      values: createEmptyValues(),
      bold: false,
      mergeRemark: false,
    };
  }

  function createCustomRow() {
    return {
      kind: "custom",
      values: createEmptyValues(CUSTOM_ROW_FIELDS),
      bold: true,
      mergeRemark: false,
    };
  }

  function createEmptyRow() {
    return createNormalRow();
  }

  function getRowValues(row) {
    return row && row.values && typeof row.values === "object" ? row.values : row || {};
  }

  function isCustomRow(row) {
    return Boolean(row && row.kind === "custom");
  }

  function normalizeAppRow(row, fields) {
    const targetFields = fields || state.fields;
    const hasWrappedShape = row && typeof row === "object" && row.values && typeof row.values === "object";
    const kind = hasWrappedShape && row.kind === "custom" ? "custom" : "normal";
    const rawValues = hasWrappedShape ? row.values : row;
    const values = kind === "custom"
      ? normalizeCustomRowValues(rawValues, targetFields)
      : normalizeRowShape(rawValues, targetFields);

    return {
      kind,
      values,
      bold: kind === "custom" ? row.bold !== false : Boolean(row && row.bold),
      mergeRemark: kind === "custom" ? false : Boolean(row && row.mergeRemark),
    };
  }

  function cloneAppRow(row) {
    const normalized = normalizeAppRow(row);
    return {
      ...normalized,
      values: { ...normalized.values },
    };
  }

  function ensureRowShape(row) {
    return normalizeAppRow(row, state.fields);
  }

  function normalizeRowShape(row, fields) {
    const next = {};
    fields.forEach((field) => {
      const value = row && row[field.key] != null ? String(row[field.key]) : "";
      next[field.key] = field.type === "number" ? normalizeNonNegativeNumberValue(value) : value;
    });
    return next;
  }

  function normalizeCustomRowValues(row, fields) {
    const next = normalizeRowShape(row, CUSTOM_ROW_FIELDS);
    const exportMap = getCustomRowExportMap(fields || state.fields);

    if (!next[CUSTOM_ROW_TYPE_KEY]) {
      next[CUSTOM_ROW_TYPE_KEY] = getCustomRowFallbackValue(row, exportMap.typeField, fields, isCustomTypeFallbackField);
    }
    if (!next[CUSTOM_ROW_QUANTITY_KEY]) {
      next[CUSTOM_ROW_QUANTITY_KEY] = normalizeNonNegativeNumberValue(getCustomRowFallbackValue(row, exportMap.quantityField, fields, isQuantityField));
    }
    if (!next[CUSTOM_ROW_UNIT_PRICE_KEY]) {
      next[CUSTOM_ROW_UNIT_PRICE_KEY] = normalizeNonNegativeNumberValue(getCustomRowFallbackValue(row, exportMap.unitPriceField, fields, isUnitPriceField));
    }

    syncCustomRowAmountValue(next);
    return next;
  }

  function getCustomRowFallbackValue(row, preferredField, fields, predicate) {
    if (!row || typeof row !== "object") {
      return "";
    }

    if (preferredField && row[preferredField.key] != null && String(row[preferredField.key]).trim() !== "") {
      return String(row[preferredField.key]);
    }

    const fallbackField = (fields || []).find((field) => predicate(field) && row[field.key] != null && String(row[field.key]).trim() !== "");
    return fallbackField ? String(row[fallbackField.key]) : "";
  }

  function syncCustomRowAmountValue(values) {
    const calculatedValue = calculateCustomRowAmount(values);
    values[CUSTOM_ROW_AMOUNT_KEY] = calculatedValue === "" ? "" : formatCalculatedValue(calculatedValue);
  }

  function normalizeNonNegativeNumberValue(value) {
    const text = String(value == null ? "" : value).trim();
    if (text === "") {
      return "";
    }

    const numericValue = Number(text.replace(/,/g, ""));
    if (!Number.isFinite(numericValue)) {
      return text.includes("-") ? "" : text;
    }

    return numericValue < 0 ? "0" : text;
  }

  async function init() {
    els = {
      fieldCount: document.getElementById("fieldCount"),
      rowCount: document.getElementById("rowCount"),
      generateBtn: document.getElementById("generateBtn"),
      documentName: document.getElementById("documentName"),
      depositAmount: document.getElementById("depositAmount"),
      draftPanel: document.getElementById("draftPanel"),
      configPanel: document.getElementById("configPanel"),
      toggleConfigBtn: document.getElementById("toggleConfigBtn"),
      addFieldBtn: document.getElementById("addFieldBtn"),
      fieldList: document.getElementById("fieldList"),
      addRowTopBtn: document.getElementById("addRowTopBtn"),
      addCustomRowTopBtn: document.getElementById("addCustomRowTopBtn"),
      addRowBottomBtn: document.getElementById("addRowBottomBtn"),
      addCustomRowBottomBtn: document.getElementById("addCustomRowBottomBtn"),
      clearRowsBtn: document.getElementById("clearRowsBtn"),
      naturalFillBtn: document.getElementById("naturalFillBtn"),
      naturalFillPanel: document.getElementById("naturalFillPanel"),
      naturalFillText: document.getElementById("naturalFillText"),
      naturalFillSubmitBtn: document.getElementById("naturalFillSubmitBtn"),
      naturalFillCancelBtn: document.getElementById("naturalFillCancelBtn"),
      rowList: document.getElementById("rowList"),
      status: document.getElementById("status"),
      sharedRemarkPanel: document.getElementById("sharedRemarkPanel"),
    };

    if (!hasRequiredElements()) {
      showStartupError("页面结构加载不完整，请刷新后重试。");
      return;
    }

    await initializeDraftState();
    if (els.documentName) {
      els.documentName.addEventListener("input", handleDocumentNameInput);
    }
    if (els.depositAmount) {
      els.depositAmount.addEventListener("beforeinput", handleDepositAmountBeforeInput);
      els.depositAmount.addEventListener("input", handleDepositAmountInput);
      els.depositAmount.addEventListener("change", handleDepositAmountInput);
    }

    els.toggleConfigBtn.addEventListener("click", toggleConfigPanel);
    els.addFieldBtn.addEventListener("click", addField);
    els.addRowTopBtn.addEventListener("click", () => addRow());
    els.addCustomRowTopBtn.addEventListener("click", () => addCustomRow());
    els.addRowBottomBtn.addEventListener("click", () => addRow());
    els.addCustomRowBottomBtn.addEventListener("click", () => addCustomRow());
    els.clearRowsBtn.addEventListener("click", clearRows);
    els.generateBtn.addEventListener("click", generateExcel);
    els.naturalFillBtn.addEventListener("click", openNaturalFillPanel);
    els.naturalFillSubmitBtn.addEventListener("click", handleNaturalFillSubmit);
    els.naturalFillCancelBtn.addEventListener("click", closeNaturalFillPanel);
    els.naturalFillText.addEventListener("keydown", handleNaturalFillTextKeydown);
    document.addEventListener("keydown", handleDocumentKeydown);
    els.fieldList.addEventListener("input", handleFieldInput);
    els.fieldList.addEventListener("change", handleFieldChange);
    els.fieldList.addEventListener("click", handleFieldClick);
    els.sharedRemarkPanel.addEventListener("beforeinput", handleSharedRemarkBeforeInput);
    els.sharedRemarkPanel.addEventListener("input", handleSharedRemarkInput);
    els.sharedRemarkPanel.addEventListener("change", handleSharedRemarkChange);
    els.rowList.addEventListener("beforeinput", handleRowBeforeInput);
    els.rowList.addEventListener("input", handleRowInput);
    els.rowList.addEventListener("change", handleRowChange);
    els.rowList.addEventListener("click", handleRowClick);
    els.draftPanel.addEventListener("click", handleDraftPanelClick);

    render();
    if (state.draftStorageAvailable && state.drafts.length) {
      setStatus("已恢复草稿", "success");
    } else if (!state.draftStorageAvailable) {
      setStatus("浏览器不支持本地草稿保存，仍可正常生成 Excel。", "warning");
    }
  }

  function hasRequiredElements() {
    return Object.keys(els).every((key) => Boolean(els[key]));
  }

  function render() {
    syncDocumentNameInput();
    syncDepositAmountInput();
    renderDraftPanel();
    renderFields();
    renderSharedRemark();
    renderRows();
    updateConfigPanel();
    updateNaturalFillControls();
    updateSummary();
  }

  function updateSummary() {
    els.fieldCount.textContent = `${state.fields.length} 个表头`;
    els.rowCount.textContent = `${state.rows.length} 行数据`;
  }

  function updateConfigPanel() {
    els.configPanel.classList.toggle("is-open", state.configOpen);
    els.toggleConfigBtn.setAttribute("aria-expanded", String(state.configOpen));

    const toggleText = els.toggleConfigBtn.querySelector(".config-toggle-text");
    if (toggleText) {
      toggleText.textContent = state.configOpen ? "收起" : "展开";
    }
  }

  function toggleConfigPanel() {
    state.configOpen = !state.configOpen;
    updateConfigPanel();
  }

  function updateNaturalFillControls() {
    if (!els.naturalFillBtn) {
      return;
    }

    const isOpen = !els.naturalFillPanel.hidden;
    els.naturalFillBtn.disabled = state.naturalFillRunning;
    els.naturalFillBtn.classList.toggle("is-loading", state.naturalFillRunning);
    els.naturalFillBtn.setAttribute("aria-expanded", String(isOpen));
    els.naturalFillBtn.setAttribute("aria-label", state.naturalFillRunning ? "智能填写中" : "智能填写");
    els.naturalFillBtn.title = state.naturalFillRunning ? "智能填写中" : "智能填写";
    if (els.naturalFillText) {
      els.naturalFillText.disabled = state.naturalFillRunning;
    }
    if (els.naturalFillSubmitBtn) {
      els.naturalFillSubmitBtn.disabled = state.naturalFillRunning;
      els.naturalFillSubmitBtn.textContent = state.naturalFillRunning ? "正在填入..." : "填入表格";
    }
    if (els.naturalFillCancelBtn) {
      els.naturalFillCancelBtn.disabled = state.naturalFillRunning;
    }
  }

  function renderDraftPanel() {
    if (!els.draftPanel) {
      return;
    }

    if (!state.draftStorageAvailable) {
      els.draftPanel.innerHTML = '<div class="draft-unavailable">本机浏览器暂不支持草稿保存。</div>';
      return;
    }

    const sortedDrafts = normalizeDraftList(state.drafts);
    const activeDraft = sortedDrafts.find((draft) => draft.id === state.activeDraftId);
    const activeName = draftDisplayName(state.documentName || (activeDraft && activeDraft.name));
    const activeUpdatedAt = activeDraft && activeDraft.updatedAt ? formatDraftTime(activeDraft.updatedAt) : "未保存";
    const activeRows = state.rows.length;
    const list = sortedDrafts.length
      ? sortedDrafts
          .map((draft) => {
            const isActive = draft.id === state.activeDraftId;
            return `
              <button class="draft-chip ${isActive ? "is-active" : ""}" type="button" data-draft-action="load" data-draft-id="${escapeAttr(draft.id)}">
                <span>${escapeHtml(draftDisplayName(draft.documentName || draft.name))}</span>
                <small>${escapeHtml(formatDraftTime(draft.updatedAt))} · ${draft.rows.length} 行</small>
              </button>
            `;
          })
          .join("")
      : '<span class="draft-empty">暂无已保存草稿</span>';

    els.draftPanel.innerHTML = `
      <div class="draft-current">
        <div>
          <span class="draft-label">当前草稿</span>
          <strong>${escapeHtml(activeName)}</strong>
          <small>${escapeHtml(activeUpdatedAt)} · ${activeRows} 行</small>
        </div>
        <div class="draft-actions">
          <button class="secondary-action" type="button" data-draft-action="continue">继续编辑</button>
          <button class="secondary-action" type="button" data-draft-action="new">新建表格</button>
        </div>
      </div>
      <div class="draft-list" aria-label="草稿列表">${list}</div>
    `;
  }

  function formatDraftTime(value) {
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) {
      return "时间未知";
    }

    const now = Date.now();
    const diffMs = Math.max(0, now - date.getTime());
    if (diffMs < 60000) {
      return "刚刚";
    }
    if (diffMs < 3600000) {
      return `${Math.floor(diffMs / 60000)} 分钟前`;
    }

    const pad = (number) => String(number).padStart(2, "0");
    const today = new Date(now);
    const sameDay = date.getFullYear() === today.getFullYear() && date.getMonth() === today.getMonth() && date.getDate() === today.getDate();
    const time = `${pad(date.getHours())}:${pad(date.getMinutes())}`;
    if (sameDay) {
      return `今天 ${time}`;
    }

    return `${date.getMonth() + 1}/${date.getDate()} ${time}`;
  }

  function handleDraftPanelClick(event) {
    const button = event.target.closest("[data-draft-action]");
    if (!button) {
      return;
    }

    const action = button.dataset.draftAction;
    if (action === "new") {
      startNewDraft();
      return;
    }

    if (action === "continue") {
      setStatus("已恢复草稿", "success");
      return;
    }

    if (action === "load") {
      const draftId = button.dataset.draftId;
      if (draftId === state.activeDraftId) {
        flushDraftSave({ showSaved: false }).finally(() => {
          setStatus("已恢复草稿", "success");
          renderDraftPanel();
        });
        return;
      }

      const draft = state.drafts.find((item) => item.id === draftId);
      if (!draft) {
        return;
      }

      flushDraftSave({ showSaved: false }).finally(() => {
        if (applyDraft(draft)) {
          render();
          setStatus("已恢复草稿", "success");
        }
      });
    }
  }

  function renderFields() {
    if (!state.fields.length) {
      els.fieldList.innerHTML = '<div class="empty-state">暂无表头</div>';
      return;
    }

    els.fieldList.innerHTML = state.fields.map((field, index) => fieldTemplate(field, index)).join("");
  }

  function fieldTemplate(field, index) {
    const optionsValue = field.options.join(", ");
    const selectOptions = FIELD_TYPES.map((type) => {
      const selected = field.type === type ? " selected" : "";
      return `<option value="${type}"${selected}>${TYPE_TEXT[type]}</option>`;
    }).join("");
    const optionsInput = field.type === "select"
      ? `<label class="full">下拉选项<input data-field-prop="options" data-field-index="${index}" value="${escapeAttr(optionsValue)}" placeholder="选项一, 选项二"></label>`
      : "";
    const invalidText = state.invalidFields.has(field.key) ? "表头名称不能为空" : "";

    return `
      <article class="field-card" data-field-index="${index}">
        <div class="card-heading">
          <div class="card-title">字段 ${index + 1}</div>
          <div class="action-strip">
            <button class="small-icon" type="button" data-field-action="up" data-field-index="${index}" aria-label="上移" title="上移">↑</button>
            <button class="small-icon" type="button" data-field-action="down" data-field-index="${index}" aria-label="下移" title="下移">↓</button>
            <button class="small-icon danger" type="button" data-field-action="delete" data-field-index="${index}" aria-label="删除" title="删除">×</button>
          </div>
        </div>
        <div class="field-grid">
          <label>表头名称
            <input data-field-prop="label" data-field-index="${index}" value="${escapeAttr(field.label)}" class="${state.invalidFields.has(field.key) ? "is-invalid" : ""}">
            <span class="field-error">${invalidText}</span>
          </label>
          <label>输入类型
            <select data-field-prop="type" data-field-index="${index}">${selectOptions}</select>
          </label>
          <label class="full">上级表头
            <input data-field-prop="group" data-field-index="${index}" value="${escapeAttr(field.group)}" placeholder="如：成品规格(M)、轨道；留空表示独立表头">
          </label>
          <label class="checkbox-label full">
            <input type="checkbox" data-field-prop="required" data-field-index="${index}" ${field.required ? "checked" : ""}>
            必填
          </label>
          ${optionsInput}
        </div>
      </article>
    `;
  }

  function renderSharedRemark() {
    const field = getSharedRemarkField(state.fields);
    if (!field) {
      state.invalidSharedFields.clear();
      els.sharedRemarkPanel.hidden = true;
      els.sharedRemarkPanel.innerHTML = "";
      return;
    }

    ensureSharedRemarkState();
    els.sharedRemarkPanel.hidden = false;
    els.sharedRemarkPanel.innerHTML = sharedRemarkTemplate(field);
  }

  function sharedRemarkTemplate(field) {
    const invalidClass = state.invalidSharedFields.has(field.key) ? "is-invalid" : "";
    const requiredMark = field.required ? '<span class="required-dot">*</span>' : "";
    const title = displayFieldLabel(field);
    const label = `
      <span class="label-row">
        <span>${escapeHtml(title || "备注")}</span>
        ${requiredMark}
      </span>
    `;
    const commonAttrs = `data-shared-field-key="${escapeAttr(field.key)}" data-shared-field-type="${escapeAttr(field.type)}" class="${invalidClass}"`;

    if (field.type === "select") {
      const options = ['<option value="">请选择</option>']
        .concat(field.options.map((option) => {
          const selected = state.sharedRemark === option ? " selected" : "";
          return `<option value="${escapeAttr(option)}"${selected}>${escapeHtml(option)}</option>`;
        }))
        .join("");
      return `<label>${label}<select ${commonAttrs}>${options}</select></label>`;
    }

    const type = field.type === "number" ? "number" : field.type === "date" ? "date" : "text";
    const numberAttrs = field.type === "number" ? ' step="any" min="0" inputmode="decimal"' : "";
    return `<label>${label}<input type="${type}"${numberAttrs} ${commonAttrs} value="${escapeAttr(state.sharedRemark)}" placeholder="所有行共用此备注"></label>`;
  }

  function renderRows() {
    if (!state.fields.length) {
      els.rowList.innerHTML = '<div class="empty-state">请先添加表头</div>';
      updateSummary();
      return;
    }

    if (!state.rows.length) {
      els.rowList.innerHTML = '<div class="empty-state">暂无数据行</div>';
      updateSummary();
      return;
    }

    ensureSharedRemarkState();
    const calculatedFields = getCalculatedFields(state.fields);
    state.rows = state.rows.map(ensureRowShape);
    state.rows.forEach((row) => syncCalculatedRow(row, calculatedFields));
    els.rowList.innerHTML = state.rows
      .map((row, index) => rowTemplate(row, index, calculatedFields))
      .join("");
    updateSummary();
  }

  function rowTemplate(row, rowIndex, calculatedFields) {
    const custom = isCustomRow(row);
    const values = getRowValues(row);
    const entryFields = getRowEntryFields(state.fields, row);
    const controls = entryFields
      .map((field) => inputTemplate(field, values[field.key] || "", rowIndex, row, getRowCalculatedField(row, field.key, calculatedFields)))
      .join("");
    return `
      <article class="row-card ${custom ? "custom-row-card" : ""}" data-row-index="${rowIndex}">
        <div class="card-heading">
          <div class="card-title">${custom ? "自由行" : `第 ${rowIndex + 1} 行`}</div>
          <div class="action-strip">
            <button class="row-action" type="button" data-row-action="copy" data-row-index="${rowIndex}">复制</button>
            <button class="row-action danger" type="button" data-row-action="delete" data-row-index="${rowIndex}">删除</button>
          </div>
        </div>
        <div class="row-input-grid">${controls}</div>
      </article>
    `;
  }

  function inputTemplate(field, value, rowIndex, row, calculatedField) {
    const cellKey = `${rowIndex}:${field.key}`;
    const invalidClass = state.invalidCells.has(cellKey) ? "is-invalid" : "";
    const requiredMark = field.required ? '<span class="required-dot">*</span>' : "";
    const title = displayFieldLabel(field);
    const label = `
      <span class="label-row">
        <span>${escapeHtml(title || "未命名表头")}</span>
        ${requiredMark}
      </span>
    `;
    const commonAttrs = `data-row-index="${rowIndex}" data-field-key="${escapeAttr(field.key)}" data-field-type="${escapeAttr(field.type)}" class="${invalidClass}"`;
    const customFieldClass = isCustomRow(row)
      ? isCustomPrimaryField(field) ? "custom-primary-field" : "custom-secondary-field"
      : "";
    const labelClass = customFieldClass ? ` class="${customFieldClass}"` : "";

    if (calculatedField && !isCustomRow(row)) {
      const calculatedValue = calculateFieldValue(row, calculatedField);
      const displayValue = calculatedValue === "" ? "" : formatCalculatedValue(calculatedValue);
      return `<label class="calculated-label">${label}<input type="number" step="any" min="0" readonly data-row-index="${rowIndex}" data-calculated-field-key="${escapeAttr(field.key)}" class="${invalidClass}" value="${escapeAttr(displayValue)}" placeholder="自动计算" title="自动计算"></label>`;
    }

    if (field.type === "select") {
      const options = ['<option value="">请选择</option>']
        .concat(field.options.map((option) => {
          const selected = value === option ? " selected" : "";
          return `<option value="${escapeAttr(option)}"${selected}>${escapeHtml(option)}</option>`;
        }))
        .join("");
      return `<label${labelClass}>${label}<select ${commonAttrs}>${options}</select></label>`;
    }

    if (field.type === "image") {
      const image = parseImageValue(value);
      const imageClass = ["image-input", invalidClass, customFieldClass].filter(Boolean).join(" ");
      const preview = image
        ? `<div class="image-preview">
            <img src="${escapeAttr(image.dataUrl)}" alt="${escapeAttr(image.name || "已选择图片")}">
            <div class="image-meta">
              <span>${escapeHtml(image.name || "已选择图片")}</span>
              <button type="button" data-image-action="clear" data-row-index="${rowIndex}" data-field-key="${escapeAttr(field.key)}">清除</button>
            </div>
          </div>`
        : '<div class="image-empty">未选择图片</div>';
      return `<label class="${escapeAttr(imageClass)}">
        ${label}
        <span class="image-select-button">选择本地图片</span>
        <input type="file" accept="image/png,image/jpeg,image/gif" data-image-input="true" ${commonAttrs}>
        ${preview}
      </label>`;
    }

    const type = field.type === "number" ? "number" : field.type === "date" ? "date" : "text";
    const numberAttrs = field.type === "number" ? ' step="any" min="0" inputmode="decimal"' : "";
    const customCalculatedValue = calculatedField && isCustomRow(row)
      ? calculateFieldValue(row, calculatedField)
      : "";
    const displayValue = customCalculatedValue === "" ? value : formatCalculatedValue(customCalculatedValue);
    const calculatedAttrs = calculatedField && isCustomRow(row) ? ' readonly placeholder="自动计算" title="自动计算"' : "";
    return `<label${labelClass}>${label}<input type="${type}"${numberAttrs}${calculatedAttrs} ${commonAttrs} value="${escapeAttr(displayValue)}"></label>`;
  }

  function displayFieldLabel(field) {
    const label = String(field.label || "");
    const group = String(field.group || "").trim();
    return group ? `${group} - ${label}` : label;
  }

  function isRemarkField(field) {
    const normalizedKey = String(field && field.key ? field.key : "").trim().toLowerCase();
    const normalizedLabel = normalizeLabel(displayFieldLabel(field || {}));
    return normalizedKey === "remark" || normalizedLabel === "备注" || normalizedLabel.endsWith("-备注");
  }

  function getSharedRemarkField(fields) {
    return (fields || []).find(isRemarkField) || null;
  }

  function isCustomPrimaryField(field) {
    return isRemarkField(field) ||
      isTypeField(field) ||
      isAmountField(field) ||
      isQuantityField(field) ||
      isUnitPriceField(field) ||
      isNameField(field) ||
      isModelField(field);
  }

  function getRowEntryFields(fields, row) {
    if (isCustomRow(row)) {
      return CUSTOM_ROW_FIELDS;
    }

    const sharedRemarkField = getSharedRemarkField(fields);
    return sharedRemarkField ? fields.filter((field) => field.key !== sharedRemarkField.key) : fields;
  }

  function getRowCalculatedField(row, fieldKey, calculatedFields) {
    if (isCustomRow(row)) {
      return fieldKey === CUSTOM_ROW_AMOUNT_KEY ? makeCustomRowAmountField() : null;
    }

    return calculatedFields.get(fieldKey) || null;
  }

  function ensureSharedRemarkState() {
    const field = getSharedRemarkField(state.fields);
    if (!field) {
      state.sharedRemark = "";
      return;
    }

    state.rows = state.rows.map(ensureRowShape);
    const firstRowWithRemark = state.rows.find((row) => !isCustomRow(row) && getRowValues(row)[field.key]);
    if (!state.sharedRemark && firstRowWithRemark) {
      state.sharedRemark = String(getRowValues(firstRowWithRemark)[field.key] || "");
    }

    state.rows.forEach((row) => {
      if (row && !isCustomRow(row)) {
        getRowValues(row)[field.key] = state.sharedRemark;
      }
    });
  }

  function handleDocumentNameInput(event) {
    state.documentName = event.target.value;
    renderDraftPanel();
    scheduleDraftSave();
  }

  function handleDepositAmountBeforeInput(event) {
    if (event.data && event.data.includes("-")) {
      event.preventDefault();
    }
  }

  function handleDepositAmountInput(event) {
    const value = normalizeNonNegativeNumberValue(event.target.value);
    if (event.target.value !== value) {
      event.target.value = value;
    }

    state.depositAmount = value;
    renderDraftPanel();
    scheduleDraftSave();
  }

  function handleFieldInput(event) {
    const target = event.target;
    if (!target.dataset.fieldProp) {
      return;
    }

    const index = Number(target.dataset.fieldIndex);
    const prop = target.dataset.fieldProp;

    if (!Number.isInteger(index) || !state.fields[index] || !prop) {
      return;
    }

    if (prop === "label") {
      state.fields[index].label = target.value;
      state.invalidFields.delete(state.fields[index].key);
    } else if (prop === "group") {
      state.fields[index].group = target.value;
    } else if (prop === "options") {
      state.fields[index].options = splitOptions(target.value);
    }

    saveFields();
    renderSharedRemark();
    renderRows();
    updateSummary();
    renderDraftPanel();
    scheduleDraftSave();
  }

  function handleFieldChange(event) {
    const target = event.target;
    if (!target.dataset.fieldProp) {
      return;
    }

    const index = Number(target.dataset.fieldIndex);
    const prop = target.dataset.fieldProp;

    if (!Number.isInteger(index) || !state.fields[index] || !prop) {
      return;
    }

    if (prop === "type") {
      state.fields[index].type = FIELD_TYPES.includes(target.value) ? target.value : "text";
      if (state.fields[index].type === "select" && !state.fields[index].options.length) {
        state.fields[index].options = ["选项一", "选项二"];
      }
      saveFields();
      render();
      scheduleDraftSave();
      return;
    }

    if (prop === "required") {
      state.fields[index].required = target.checked;
      saveFields();
      renderSharedRemark();
      renderRows();
      updateSummary();
      renderDraftPanel();
      scheduleDraftSave();
    }
  }

  function handleFieldClick(event) {
    const button = event.target.closest("[data-field-action]");
    if (!button) {
      return;
    }

    const index = Number(button.dataset.fieldIndex);
    const action = button.dataset.fieldAction;

    if (!Number.isInteger(index) || !state.fields[index]) {
      return;
    }

    if (action === "delete") {
      const [removed] = state.fields.splice(index, 1);
      state.rows.forEach((row) => {
        delete getRowValues(row)[removed.key];
      });
      syncSharedRemarkFromRemovedField(removed);
    } else if (action === "up" && index > 0) {
      [state.fields[index - 1], state.fields[index]] = [state.fields[index], state.fields[index - 1]];
    } else if (action === "down" && index < state.fields.length - 1) {
      [state.fields[index + 1], state.fields[index]] = [state.fields[index], state.fields[index + 1]];
    }

    saveFields();
    state.invalidCells.clear();
    state.invalidSharedFields.clear();
    state.invalidFields.clear();
    render();
    scheduleDraftSave();
  }

  function handleRowBeforeInput(event) {
    const target = event.target;
    if (!isNumberValueTarget(target)) {
      return;
    }

    if (event.data && event.data.includes("-")) {
      event.preventDefault();
    }
  }

  function handleSharedRemarkBeforeInput(event) {
    const target = event.target;
    if (!isNumberValueTarget(target)) {
      return;
    }

    if (event.data && event.data.includes("-")) {
      event.preventDefault();
    }
  }

  function handleSharedRemarkInput(event) {
    const target = event.target;
    const fieldKey = target.dataset.sharedFieldKey;
    if (!fieldKey) {
      return;
    }

    const field = state.fields.find((item) => item.key === fieldKey);
    const value = field && field.type === "number" ? normalizeNonNegativeNumberValue(target.value) : target.value;
    if (target.value !== value) {
      target.value = value;
    }

    state.sharedRemark = value;
    state.invalidSharedFields.delete(fieldKey);
    target.classList.remove("is-invalid");
    ensureSharedRemarkState();
    renderDraftPanel();
    scheduleDraftSave();
  }

  function handleSharedRemarkChange(event) {
    handleSharedRemarkInput(event);
  }

  function isNumberValueTarget(target) {
    return Boolean(
      target &&
        target.dataset &&
        target.dataset.imageInput !== "true" &&
        (target.dataset.fieldKey || target.dataset.sharedFieldKey) &&
        (target.dataset.fieldType === "number" || target.type === "number")
    );
  }

  function handleRowInput(event) {
    const target = event.target;
    if (target.dataset.imageInput === "true") {
      return;
    }

    const rowIndex = Number(target.dataset.rowIndex);
    const fieldKey = target.dataset.fieldKey;

    if (!Number.isInteger(rowIndex) || !state.rows[rowIndex] || !fieldKey) {
      return;
    }

    const row = state.rows[rowIndex];
    const values = getRowValues(row);
    const field = getRowEntryFields(state.fields, row).find((item) => item.key === fieldKey);
    const value = field && field.type === "number" ? normalizeNonNegativeNumberValue(target.value) : target.value;
    if (target.value !== value) {
      target.value = value;
    }

    values[fieldKey] = value;
    if (isCustomRow(row)) {
      syncCustomRowAmountValue(values);
    } else {
      syncCalculatedRow(row);
    }
    state.invalidCells.delete(`${rowIndex}:${fieldKey}`);
    target.classList.remove("is-invalid");
    updateCalculatedOutputs(rowIndex);
    renderDraftPanel();
    scheduleDraftSave();
  }

  function handleRowChange(event) {
    const target = event.target;
    if (target.dataset.imageInput === "true") {
      handleImageInput(target);
      return;
    }

    if (target.dataset.rowOption) {
      handleRowOptionChange(target);
      return;
    }

    handleRowInput(event);
  }

  function handleRowOptionChange(input) {
    const rowIndex = Number(input.dataset.rowIndex);
    const option = input.dataset.rowOption;

    if (!Number.isInteger(rowIndex) || !state.rows[rowIndex] || !isCustomRow(state.rows[rowIndex])) {
      return;
    }

    if (option === "bold") {
      state.rows[rowIndex].bold = input.checked;
    } else if (option === "mergeRemark") {
      state.rows[rowIndex].mergeRemark = input.checked;
    }
    renderDraftPanel();
    scheduleDraftSave();
  }

  function handleImageInput(input) {
    const rowIndex = Number(input.dataset.rowIndex);
    const fieldKey = input.dataset.fieldKey;
    const file = input.files && input.files[0];

    if (!Number.isInteger(rowIndex) || !state.rows[rowIndex] || !fieldKey) {
      return;
    }

    if (!file) {
      return;
    }

    if (!SUPPORTED_IMAGE_MIMES.includes(file.type)) {
      input.value = "";
      setStatus("请选择 PNG、JPG 或 GIF 图片", "warning");
      return;
    }

    const reader = new FileReader();
    reader.onload = () => {
      getRowValues(state.rows[rowIndex])[fieldKey] = JSON.stringify({
        kind: "image",
        name: file.name,
        type: file.type,
        dataUrl: String(reader.result || ""),
      });
      state.invalidCells.delete(`${rowIndex}:${fieldKey}`);
      renderRows();
      setStatus("图片已读取", "success");
      scheduleDraftSave();
    };
    reader.onerror = () => {
      input.value = "";
      setStatus("图片读取失败，请重新选择", "error");
    };
    reader.readAsDataURL(file);
  }

  function handleRowClick(event) {
    const button = event.target.closest("[data-row-action]");
    const imageButton = event.target.closest("[data-image-action]");
    if (imageButton) {
      event.preventDefault();
      if (handleImageClear(imageButton)) {
        return;
      }
    }

    if (!button) {
      return;
    }

    const rowIndex = Number(button.dataset.rowIndex);
    const action = button.dataset.rowAction;

    if (!Number.isInteger(rowIndex) || !state.rows[rowIndex]) {
      return;
    }

    if (action === "delete") {
      state.rows.splice(rowIndex, 1);
    } else if (action === "copy") {
      state.rows.splice(rowIndex + 1, 0, cloneAppRow(state.rows[rowIndex]));
    }

    state.invalidCells.clear();
    renderRows();
    renderDraftPanel();
    scheduleDraftSave();
  }

  function handleImageClear(button) {
    const rowIndex = Number(button.dataset.rowIndex);
    const fieldKey = button.dataset.fieldKey;

    if (!Number.isInteger(rowIndex) || !state.rows[rowIndex] || !fieldKey) {
      return false;
    }

    getRowValues(state.rows[rowIndex])[fieldKey] = "";
    state.invalidCells.delete(`${rowIndex}:${fieldKey}`);
    renderRows();
    setStatus("图片已清除", "success");
    scheduleDraftSave();
    return true;
  }

  function syncSharedRemarkFromRemovedField(removed) {
    if (!removed || !isRemarkField(removed)) {
      return;
    }

    state.sharedRemark = "";
    state.invalidSharedFields.clear();
  }

  function addField() {
    const key = makeFieldKey("field");
    state.fields.push({
      key,
      label: "新表头",
      group: "",
      type: "text",
      required: false,
      options: [],
    });
    state.rows = state.rows.map((row) => {
      const normalized = normalizeAppRow(row);
      return {
        ...normalized,
        values: { ...normalized.values, [key]: "" },
      };
    });
    saveFields();
    render();
    scheduleDraftSave();
    setStatus("已添加表头", "success");
  }

  function addRow(copy) {
    if (!state.fields.length) {
      setStatus("请先添加表头", "warning");
      return;
    }

    ensureSharedRemarkState();
    state.rows.push(copy ? cloneAppRow(copy) : createNormalRow());
    ensureSharedRemarkState();
    renderRows();
    renderDraftPanel();
    scheduleDraftSave();
    setStatus("已添加行", "success");
  }

  function addCustomRow() {
    if (!state.fields.length) {
      setStatus("请先添加表头", "warning");
      return;
    }

    ensureSharedRemarkState();
    state.rows.push(createCustomRow());
    renderRows();
    renderDraftPanel();
    scheduleDraftSave();
    setStatus("已添加自由行", "success");
  }

  function clearRows() {
    state.rows = [];
    state.invalidCells.clear();
    renderRows();
    renderDraftPanel();
    scheduleDraftSave();
    setStatus("已清空数据行", "success");
  }

  function openNaturalFillPanel() {
    if (state.naturalFillRunning) {
      return;
    }

    if (!state.fields.length) {
      setStatus("请先添加表头", "warning");
      return;
    }

    const extractableFields = getNaturalFillFields();
    if (!extractableFields.length) {
      setStatus("没有可智能填充的文本或数字表头", "warning");
      return;
    }

    els.naturalFillPanel.hidden = false;
    updateNaturalFillControls();
    els.naturalFillText.focus();
  }

  function closeNaturalFillPanel() {
    if (state.naturalFillRunning) {
      return;
    }

    els.naturalFillPanel.hidden = true;
    updateNaturalFillControls();
  }

  function handleNaturalFillTextKeydown(event) {
    if ((event.ctrlKey || event.metaKey) && event.key === "Enter") {
      event.preventDefault();
      handleNaturalFillSubmit();
    }
  }

  function handleDocumentKeydown(event) {
    if (event.key === "Escape" && els.naturalFillPanel && !els.naturalFillPanel.hidden) {
      closeNaturalFillPanel();
    }
  }

  async function handleNaturalFillSubmit() {
    if (state.naturalFillRunning) {
      return;
    }

    const text = String(els.naturalFillText.value || "").trim();
    if (!text) {
      setStatus("请输入要填入表格的自然语言内容", "warning");
      els.naturalFillText.focus();
      return;
    }

    if (!state.fields.length) {
      setStatus("请先添加表头", "warning");
      return;
    }

    const fields = getNaturalFillFields();
    if (!fields.length) {
      setStatus("没有可智能填充的文本或数字表头", "warning");
      return;
    }

    state.naturalFillRunning = true;
    updateNaturalFillControls();
    setStatus("正在根据描述填写表格...", "success");

    try {
      const result = await requestNaturalFillExtraction(text, fields);
      const addedCount = appendNaturalFillRows(result.rows || []);
      if (!addedCount) {
        setStatus("未解析到可填写的数据", "warning");
        return;
      }

      renderRows();
      renderDraftPanel();
      scheduleDraftSave();
      els.naturalFillText.value = "";
      closeNaturalFillPanel();
      setStatus(`已智能填入 ${addedCount} 行`, result.warnings && result.warnings.length ? "warning" : "success");
    } catch (error) {
      setStatus(`智能填行失败：${error.message || "请稍后重试"}`, "error");
    } finally {
      state.naturalFillRunning = false;
      updateNaturalFillControls();
    }
  }

  function getNaturalFillFields() {
    return state.fields
      .filter((field) => field.type !== "image" && !isRemarkField(field))
      .map((field) => ({
        key: field.key,
        label: field.label,
        group: field.group,
        type: field.type,
        options: field.options,
        required: field.required,
      }));
  }

  async function requestNaturalFillExtraction(text, fields) {
    const body = JSON.stringify({ text, fields });
    let response;
    try {
      response = await fetch(NATURAL_FILL_ENDPOINT, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body,
      });
    } catch (error) {
      response = await fetchLocalNaturalFillEndpoint(body, error);
    }

    if (response.status === 404 && shouldTryLocalNaturalFillEndpoint()) {
      response = await fetchLocalNaturalFillEndpoint(body);
    }

    let payload = null;
    try {
      payload = await response.json();
    } catch (error) {
      payload = null;
    }

    if (!response.ok) {
      throw new Error((payload && payload.error) || getNaturalFillHttpError(response.status));
    }

    if (!payload || !Array.isArray(payload.rows)) {
      throw new Error("智能填行服务返回格式不正确");
    }

    return payload;
  }

  async function fetchLocalNaturalFillEndpoint(body, originalError) {
    if (!shouldTryLocalNaturalFillEndpoint()) {
      throw new Error(getNaturalFillConnectionError(originalError));
    }

    try {
      return await fetch(`${getNaturalFillFallbackOrigin()}${NATURAL_FILL_ENDPOINT}`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body,
      });
    } catch (error) {
      throw new Error(getNaturalFillConnectionError(originalError || error));
    }
  }

  function shouldTryLocalNaturalFillEndpoint() {
    if (!root.location) {
      return false;
    }

    const host = root.location.host;
    return root.location.protocol === "file:" || (host && host !== "127.0.0.1:4173" && host !== "localhost:4173");
  }

  function getNaturalFillFallbackOrigin() {
    if (!root.location || root.location.protocol === "file:") {
      return "http://127.0.0.1:4173";
    }

    const hostname = root.location.hostname || "127.0.0.1";
    return `http://${hostname}:4173`;
  }

  function getNaturalFillConnectionError(error) {
    if (root.location && root.location.protocol === "file:") {
      return "请先启动本地服务 node server.js，并通过服务启动后显示的地址打开页面";
    }

    return `无法连接智能填行服务，请确认电脑端已运行 node server.js，并用服务启动后显示的地址打开页面（${error.message || "网络错误"}）`;
  }

  function getNaturalFillHttpError(status) {
    if (status === 404) {
      return "当前页面不是由智能填行服务打开，请使用服务启动后显示的地址访问后再智能填行";
    }

    return `本地智能填行服务返回 ${status}`;
  }

  function appendNaturalFillRows(rows) {
    const fieldsByKey = new Map(state.fields.map((field) => [field.key, field]));
    const sharedRemarkField = getSharedRemarkField(state.fields);
    const calculatedFields = getCalculatedFields(state.fields);
    const rowsToAdd = [];

    rows.forEach((item) => {
      const sourceValues = item && item.values && typeof item.values === "object" ? item.values : {};
      const row = createNormalRow();
      const values = getRowValues(row);
      let hasValue = false;

      Object.keys(sourceValues).forEach((key) => {
        const field = fieldsByKey.get(key);
        if (!field || field.type === "image" || (sharedRemarkField && field.key === sharedRemarkField.key)) {
          return;
        }

        const value = normalizeNaturalFillValue(sourceValues[key], field);
        if (value === "") {
          return;
        }

        values[field.key] = value;
        hasValue = true;
      });

      if (!hasValue) {
        return;
      }

      if (sharedRemarkField) {
        values[sharedRemarkField.key] = state.sharedRemark;
      }
      syncCalculatedRow(row, calculatedFields);
      rowsToAdd.push(row);
    });

    if (rowsToAdd.length) {
      if (state.rows.length === 1 && isEmptyNormalRow(state.rows[0])) {
        state.rows = [];
      }
      state.rows.push(...rowsToAdd);
      ensureSharedRemarkState();
    }

    return rowsToAdd.length;
  }

  function isEmptyNormalRow(row) {
    if (!row || isCustomRow(row)) {
      return false;
    }

    const values = getRowValues(row);
    return Object.keys(values).every((key) => String(values[key] == null ? "" : values[key]).trim() === "");
  }

  function normalizeNaturalFillValue(value, field) {
    const text = String(value == null ? "" : value).trim();
    if (!text) {
      return "";
    }

    if (field.type === "number") {
      const cleaned = text.replace(/,/g, "").replace(/[^\d.]/g, "");
      return cleaned ? normalizeNonNegativeNumberValue(cleaned) : "";
    }

    if (field.type === "select" && field.options.length) {
      const exact = field.options.find((option) => option === text);
      if (exact) {
        return exact;
      }
      const normalizedText = normalizeLabel(text);
      const fuzzy = field.options.find((option) => normalizeLabel(option) === normalizedText);
      return fuzzy || text;
    }

    return text;
  }

  async function generateExcel() {
    const validation = validateBeforeExport();
    if (!validation.ok) {
      setStatus(validation.message, validation.level || "error");
      renderFields();
      renderSharedRemark();
      renderRows();
      return;
    }

    try {
      ensureSharedRemarkState();
      await flushDraftSave({ showSaved: false });
      const blob = await createXlsxBlob(state.fields, state.rows, state.documentName, state.depositAmount);
      const filename = `自动生成表格_${formatTimestamp(new Date())}.xlsx`;
      downloadBlob(blob, filename);
      setStatus("Excel 已生成", "success");
    } catch (error) {
      setStatus(`生成失败：${error.message || "未知错误"}`, "error");
    }
  }

  function validateBeforeExport() {
    state.invalidCells.clear();
    state.invalidSharedFields.clear();
    state.invalidFields.clear();

    if (!state.fields.length) {
      return { ok: false, message: "请先添加表头", level: "warning" };
    }

    state.fields.forEach((field) => {
      if (!field.label.trim()) {
        state.invalidFields.add(field.key);
      }
    });

    if (state.invalidFields.size) {
      return { ok: false, message: "表头名称不能为空" };
    }

    if (!state.rows.length) {
      return { ok: false, message: "请至少添加一行数据", level: "warning" };
    }

    ensureSharedRemarkState();
    const sharedRemarkField = getSharedRemarkField(state.fields);
    const hasNormalRows = state.rows.some((row) => !isCustomRow(row));
    if (sharedRemarkField && sharedRemarkField.required && hasNormalRows && String(state.sharedRemark || "").trim() === "") {
      state.invalidSharedFields.add(sharedRemarkField.key);
      return { ok: false, message: "请填写备注" };
    }

    state.rows.forEach((row, rowIndex) => {
      const values = getRowValues(row);
      state.fields.forEach((field) => {
        if (isCustomRow(row) || (sharedRemarkField && field.key === sharedRemarkField.key)) {
          return;
        }
        const value = values[field.key];
        if (field.required && String(value == null ? "" : value).trim() === "") {
          state.invalidCells.add(`${rowIndex}:${field.key}`);
        }
      });
    });

    if (state.invalidCells.size) {
      return { ok: false, message: "请填写所有必填项" };
    }

    return { ok: true };
  }

  function splitOptions(value) {
    return String(value || "")
      .split(",")
      .map((item) => item.trim())
      .filter(Boolean);
  }

  function parseImageValue(value) {
    if (!value) {
      return null;
    }

    try {
      const image = JSON.parse(String(value));
      if (image && image.kind === "image" && image.dataUrl && image.type) {
        return image;
      }
    } catch (error) {
      return null;
    }
    return null;
  }

  function collectWorksheetImages(fields, rows) {
    const images = [];
    const dataStartRow = 5;

    rows.forEach((row, rowIndex) => {
      const values = getRowValues(row);
      fields.forEach((field, fieldIndex) => {
        if (field.type !== "image") {
          return;
        }

        const image = parseImageValue(values[field.key]);
        if (!image || !SUPPORTED_IMAGE_MIMES.includes(image.type)) {
          return;
        }

        const bytes = dataUrlToBytes(image.dataUrl);
        if (!bytes.length) {
          return;
        }

        images.push({
          index: images.length + 1,
          name: image.name || `图片 ${images.length + 1}`,
          type: image.type,
          extension: imageExtension(image.type),
          bytes,
          colIndex: fieldIndex,
          rowIndex: rowIndex + dataStartRow - 1,
          rowNumber: rowIndex + dataStartRow,
        });
      });
    });

    return images;
  }

  function rowHasImage(images, rowNumber) {
    return images.some((image) => image.rowNumber === rowNumber);
  }

  function imageExtension(type) {
    if (type === "image/jpeg") {
      return "jpg";
    }
    if (type === "image/gif") {
      return "gif";
    }
    return "png";
  }

  function imageContentType(extension) {
    if (extension === "jpg" || extension === "jpeg") {
      return "image/jpeg";
    }
    if (extension === "gif") {
      return "image/gif";
    }
    return "image/png";
  }

  function dataUrlToBytes(dataUrl) {
    const match = String(dataUrl || "").match(/^data:[^;]+;base64,(.+)$/);
    if (!match) {
      return new Uint8Array();
    }

    return base64ToBytes(match[1]);
  }

  function base64ToBytes(base64) {
    if (typeof root.atob === "function") {
      const binary = root.atob(base64);
      const bytes = new Uint8Array(binary.length);
      for (let index = 0; index < binary.length; index += 1) {
        bytes[index] = binary.charCodeAt(index);
      }
      return bytes;
    }

    if (typeof Buffer !== "undefined") {
      return new Uint8Array(Buffer.from(base64, "base64"));
    }

    return new Uint8Array();
  }

  function normalizeLabel(label) {
    return String(label || "")
      .replace(/\s+/g, "")
      .replace(/[()（）]/g, "")
      .toLowerCase();
  }

  function normalizeKey(field) {
    return String(field && field.key ? field.key : "").trim().toLowerCase();
  }

  function isMeterField(field) {
    const normalized = normalizeLabel(displayFieldLabel(field));
    const key = normalizeKey(field);
    return normalized.includes("米数") || normalized.includes("长度") || key.includes("meter") || key.includes("length");
  }

  function isWidthField(field) {
    const normalized = normalizeLabel(displayFieldLabel(field));
    return normalized.includes("宽") || normalizeKey(field).includes("width");
  }

  function isHeightField(field) {
    const normalized = normalizeLabel(displayFieldLabel(field));
    return normalized.includes("高") || normalizeKey(field).includes("height");
  }

  function isUnitPriceField(field) {
    const normalized = normalizeLabel(displayFieldLabel(field));
    const key = normalizeKey(field);
    return normalized.includes("单价") ||
      normalized.includes("元/米") ||
      normalized.includes("元每米") ||
      normalized.includes("元米") ||
      key.includes("unit_price") ||
      key.includes("unitprice") ||
      key.includes("price");
  }

  function isQuantityField(field) {
    const normalized = normalizeLabel(displayFieldLabel(field));
    const key = normalizeKey(field);
    return normalized.includes("数量") || key.includes("quantity") || key.includes("qty");
  }

  function isTypeField(field) {
    const normalized = normalizeLabel(displayFieldLabel(field));
    return normalized.includes("类型") || normalizeKey(field).includes("type");
  }

  function isNameField(field) {
    const normalized = normalizeLabel(displayFieldLabel(field));
    return normalized.includes("品名") || normalized.includes("项目") || normalized.includes("名称");
  }

  function isModelField(field) {
    const normalized = normalizeLabel(displayFieldLabel(field));
    return normalized.includes("型号") || normalized.includes("规格");
  }

  function isAmountField(field) {
    const normalized = normalizeLabel(displayFieldLabel(field));
    return normalized.includes("金额") || normalizeKey(field).includes("amount");
  }

  function isCustomTypeFallbackField(field) {
    return isTypeField(field) || isNameField(field) || isModelField(field);
  }

  function sameGroup(left, right) {
    return normalizeLabel(left.group) === normalizeLabel(right.group);
  }

  function getCalculatedFields(fields) {
    const calculatedFields = new Map();
    const widthField = fields.find(isWidthField);
    const heightField = fields.find(isHeightField);
    const meterField = fields.find(isMeterField);

    if (widthField && heightField && meterField) {
      calculatedFields.set(meterField.key, makeCalculatedField("sum", fields, widthField, heightField));
    }

    fields.forEach((field, index) => {
      if (!isAmountField(field)) {
        return;
      }

      const previousFields = fields.slice(0, index);
      const groupPreviousFields = previousFields.filter((candidate) => sameGroup(candidate, field));
      const sameGroupSources = normalizeLabel(field.group)
        ? groupPreviousFields.length
          ? groupPreviousFields
          : previousFields
        : previousFields;
      const meterField = sameGroupSources.findLast ? sameGroupSources.findLast(isMeterField) : findLast(sameGroupSources, isMeterField);
      const unitPriceField = sameGroupSources.findLast ? sameGroupSources.findLast(isUnitPriceField) : findLast(sameGroupSources, isUnitPriceField);
      if (meterField && unitPriceField && !field.group) {
        calculatedFields.set(field.key, makeCalculatedField("meterAmount", fields, meterField, unitPriceField));
        return;
      }

      const quantityField = sameGroupSources.findLast ? sameGroupSources.findLast(isQuantityField) : findLast(sameGroupSources, isQuantityField);
      const priceField = sameGroupSources.findLast
        ? sameGroupSources.findLast((candidate) => isUnitPriceField(candidate) && candidate.key !== field.key)
        : findLast(sameGroupSources, (candidate) => isUnitPriceField(candidate) && candidate.key !== field.key);
      if (quantityField && priceField) {
        calculatedFields.set(field.key, makeCalculatedField("quantityAmount", fields, quantityField, priceField));
      }
    });

    return calculatedFields;
  }

  function findLast(items, predicate) {
    for (let index = items.length - 1; index >= 0; index -= 1) {
      if (predicate(items[index], index, items)) {
        return items[index];
      }
    }
    return null;
  }

  function makeCalculatedField(type, fields, firstField, secondField) {
    return {
      type,
      sourceKeys: [firstField.key, secondField.key],
      sourceIndexes: [fields.indexOf(firstField), fields.indexOf(secondField)],
    };
  }

  function makeCustomRowAmountField() {
    return {
      type: "quantityAmount",
      sourceKeys: [CUSTOM_ROW_QUANTITY_KEY, CUSTOM_ROW_UNIT_PRICE_KEY],
      sourceIndexes: [1, 2],
    };
  }

  function parseNumber(value) {
    if (value == null || String(value).trim() === "") {
      return null;
    }
    const normalizedValue = normalizeNonNegativeNumberValue(value);
    if (normalizedValue === "") {
      return null;
    }
    const numericValue = Number(String(normalizedValue).replace(/,/g, ""));
    return Number.isFinite(numericValue) ? numericValue : null;
  }

  function calculateFieldValue(row, calculatedField) {
    const rowValues = getRowValues(row);
    const values = calculatedField.sourceKeys.map((key) => parseNumber(rowValues[key]));
    if (values.some((value) => value == null)) {
      return "";
    }
    if (calculatedField.type === "sum") {
      return values[0] + values[1];
    }
    return values[0] * values[1];
  }

  function calculateCustomRowAmount(values) {
    const quantity = parseNumber(values && values[CUSTOM_ROW_QUANTITY_KEY]);
    const unitPrice = parseNumber(values && values[CUSTOM_ROW_UNIT_PRICE_KEY]);
    return quantity == null || unitPrice == null ? "" : quantity * unitPrice;
  }

  function formatCalculatedValue(value) {
    const rounded = Math.round((Number(value) + Number.EPSILON) * 100) / 100;
    return Number.isInteger(rounded) ? String(rounded) : rounded.toFixed(2);
  }

  function syncCalculatedRow(row, calculatedFields) {
    if (isCustomRow(row)) {
      return;
    }

    const values = getRowValues(row);
    const fieldsToCalculate = calculatedFields || getCalculatedFields(state.fields);
    fieldsToCalculate.forEach((calculatedField, fieldKey) => {
      const value = calculateFieldValue(row, calculatedField);
      values[fieldKey] = value === "" ? "" : formatCalculatedValue(value);
    });
  }

  function updateCalculatedOutputs(rowIndex) {
    const row = state.rows[rowIndex];
    if (!row || !els.rowList) {
      return;
    }

    const values = getRowValues(row);
    if (isCustomRow(row)) {
      const input = els.rowList.querySelector(`[data-row-index="${rowIndex}"][data-field-key="${CUSTOM_ROW_AMOUNT_KEY}"]`);
      if (!input) {
        return;
      }
      const value = calculateCustomRowAmount(values);
      input.value = value === "" ? "" : formatCalculatedValue(value);
      return;
    }

    const calculatedFields = getCalculatedFields(state.fields);
    calculatedFields.forEach((calculatedField, fieldKey) => {
      const selector = `[data-row-index="${rowIndex}"][data-calculated-field-key="${fieldKey}"]`;
      const input = els.rowList.querySelector(selector);
      if (!input) {
        return;
      }
      const value = calculateFieldValue(row, calculatedField);
      input.value = value === "" ? "" : formatCalculatedValue(value);
    });
  }

  function setStatus(message, type) {
    if (!els.status) {
      return;
    }

    clearTimeout(statusTimer);
    els.status.textContent = message;
    els.status.className = `status is-visible is-${type || "success"}`;
    statusTimer = setTimeout(() => {
      els.status.className = "status";
      els.status.textContent = "";
    }, 2600);
  }

  function downloadBlob(blob, filename) {
    if (root.navigator && /MicroMessenger/i.test(root.navigator.userAgent || "")) {
      setStatus("微信内置浏览器可能拦截下载，请用系统浏览器打开后生成 Excel。", "warning");
    }

    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    if (anchor.parentNode) {
      anchor.parentNode.removeChild(anchor);
    }
    setTimeout(() => URL.revokeObjectURL(url), 1200);
  }

  async function createXlsxBlob(fields, rows, documentName, depositAmount) {
    const normalizedFields = normalizeFields(fields).filter((field) => field.label.trim());
    const normalizedRows = Array.isArray(rows) ? rows.map((row) => normalizeAppRow(row, normalizedFields)) : [];
    const images = collectWorksheetImages(normalizedFields, normalizedRows);
    const xmlFiles = buildWorkbookFiles(normalizedFields, normalizedRows, images, documentName, depositAmount);
    const zipBytes = createZip(xmlFiles);
    return new Blob([zipBytes], { type: MIME_XLSX });
  }

  function buildWorkbookFiles(fields, rows, images, documentName, depositAmount) {
    const worksheetImages = Array.isArray(images) ? images : collectWorksheetImages(fields, rows);
    const sheetXml = buildSheetXml(fields, rows, worksheetImages, documentName, depositAmount);
    const contentTypes = buildContentTypesXml(worksheetImages);
    const worksheetRelationships = worksheetImages.length ? buildWorksheetRelsXml() : null;
    const files = [
      {
        name: "[Content_Types].xml",
        content: contentTypes,
      },
      {
        name: "_rels/.rels",
        content: xmlDecl(`\
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>\
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>\
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>\
</Relationships>`),
      },
      {
        name: "docProps/core.xml",
        content: xmlDecl(`\
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\
<dc:title>自动生成表格</dc:title>\
<dc:creator>Excel 自动生成工具</dc:creator>\
<cp:lastModifiedBy>Excel 自动生成工具</cp:lastModifiedBy>\
<dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:created>\
<dcterms:modified xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:modified>\
</cp:coreProperties>`),
      },
      {
        name: "docProps/app.xml",
        content: xmlDecl(`\
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">\
<Application>Excel 自动生成工具</Application>\
<DocSecurity>0</DocSecurity>\
<ScaleCrop>false</ScaleCrop>\
<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs>\
<TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>数据</vt:lpstr></vt:vector></TitlesOfParts>\
</Properties>`),
      },
      {
        name: "xl/workbook.xml",
        content: xmlDecl(`\
<workbook xmlns="${XML_NS_MAIN}" xmlns:r="${XML_NS_REL}">\
<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="23426"/>\
<workbookPr defaultThemeVersion="166925"/>\
<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="16000" windowHeight="9000"/></bookViews>\
<sheets><sheet name="数据" sheetId="1" r:id="rId1"/></sheets>\
<calcPr calcId="191029" calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>\
</workbook>`),
      },
      {
        name: "xl/_rels/workbook.xml.rels",
        content: xmlDecl(`\
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>\
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>\
</Relationships>`),
      },
      { name: "xl/styles.xml", content: buildStylesXml() },
      { name: "xl/worksheets/sheet1.xml", content: sheetXml },
    ];
    if (worksheetRelationships) {
      files.push({
        name: "xl/worksheets/_rels/sheet1.xml.rels",
        content: worksheetRelationships,
      });
      files.push({
        name: "xl/drawings/drawing1.xml",
        content: buildDrawingXml(worksheetImages),
      });
      files.push({
        name: "xl/drawings/_rels/drawing1.xml.rels",
        content: buildDrawingRelsXml(worksheetImages),
      });
      worksheetImages.forEach((image) => {
        files.push({
          name: `xl/media/image${image.index}.${image.extension}`,
          content: image.bytes,
        });
      });
    }
    return files;
  }

  function buildContentTypesXml(images) {
    const defaults = new Set((images || []).map((image) => image.extension));
    const imageDefaults = Array.from(defaults)
      .map((extension) => `<Default Extension="${extension}" ContentType="${imageContentType(extension)}"/>`)
      .join("");
    const drawingOverride = images.length
      ? '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
      : "";

    return xmlDecl(`\
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\
<Default Extension="xml" ContentType="application/xml"/>\
${imageDefaults}\
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>\
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>\
${drawingOverride}\
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>\
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>\
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>\
</Types>`);
  }

  function buildWorksheetRelsXml() {
    return xmlDecl(`\
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>\
</Relationships>`);
  }

  function buildDrawingRelsXml(images) {
    const relationships = images
      .map((image) => `<Relationship Id="rId${image.index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image${image.index}.${image.extension}"/>`)
      .join("");

    return xmlDecl(`\
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\
${relationships}\
</Relationships>`);
  }

  function buildDrawingXml(images) {
    const anchors = images.map((image) => buildImageAnchorXml(image)).join("");
    return xmlDecl(`\
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${XML_NS_REL}">\
${anchors}\
</xdr:wsDr>`);
  }

  function buildImageAnchorXml(image) {
    const widthEmu = Math.round((IMAGE_CELL_WIDTH_PX - IMAGE_CELL_PADDING_PX * 2) * EMU_PER_PIXEL);
    const heightEmu = Math.round((IMAGE_CELL_HEIGHT_PX - IMAGE_CELL_PADDING_PX * 2) * EMU_PER_PIXEL);
    const offsetEmu = Math.round(IMAGE_CELL_PADDING_PX * EMU_PER_PIXEL);

    return `\
<xdr:oneCellAnchor editAs="oneCell">\
<xdr:from><xdr:col>${image.colIndex}</xdr:col><xdr:colOff>${offsetEmu}</xdr:colOff><xdr:row>${image.rowIndex}</xdr:row><xdr:rowOff>${offsetEmu}</xdr:rowOff></xdr:from>\
<xdr:ext cx="${widthEmu}" cy="${heightEmu}"/>\
<xdr:pic>\
<xdr:nvPicPr><xdr:cNvPr id="${image.index}" name="${escapeXml(image.name || `图片 ${image.index}`)}"/><xdr:cNvPicPr/></xdr:nvPicPr>\
<xdr:blipFill><a:blip r:embed="rId${image.index}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>\
<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${widthEmu}" cy="${heightEmu}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>\
</xdr:pic>\
<xdr:clientData/>\
</xdr:oneCellAnchor>`;
  }

  function buildSheetXml(fields, rows, images, documentName, depositAmount) {
    const calculatedFields = getCalculatedFields(fields);
    const worksheetImages = Array.isArray(images) ? images : [];
    const sharedRemarkField = getSharedRemarkField(fields);
    const sharedRemarkIndex = sharedRemarkField ? fields.indexOf(sharedRemarkField) : -1;
    const customRowExportMap = getCustomRowExportMap(fields);
    const columnCount = Math.max(fields.length, 1);
    const titleRows = 2;
    const headerRows = 2;
    const dataStartRow = titleRows + headerRows + 1;
    const summaryRowNumber = rows.length + dataStartRow;
    const noticeRowNumber = summaryRowNumber + 1;
    const rowCount = Math.max(noticeRowNumber, dataStartRow - 1);
    const lastCell = `${columnName(columnCount)}${rowCount}`;
    const titleInfo = buildTitleRows(documentName, columnCount);
    const headerInfo = buildHeaderRows(fields, titleRows + 1);
    const summaryInfo = buildSummaryRow(fields, rows, dataStartRow, summaryRowNumber, depositAmount);
    const noticeInfo = buildNoticeRow(noticeRowNumber, columnCount);
    const customMergeRefs = [];
    const dataRows = rows
      .map((row, rowIndex) => {
        const rowNumber = rowIndex + dataStartRow;
        const values = getRowValues(row);
        const custom = isCustomRow(row);
        const customRemarkMerge = custom && row.mergeRemark !== false && sharedRemarkIndex >= 0 && sharedRemarkIndex < columnCount - 1;
        const cells = fields
          .map((field, fieldIndex) => {
            const colNumber = fieldIndex + 1;
            if (custom) {
              const customCell = makeCustomRowExportCell(rowNumber, colNumber, field, fieldIndex, values, customRowExportMap);
              if (customCell) {
                return customCell;
              }
              return makeTextCell(rowNumber, colNumber, "", CUSTOM_TEXT_STYLE_ID);
            }

            const inMergedRemarkTail = customRemarkMerge && fieldIndex > sharedRemarkIndex;
            if (inMergedRemarkTail) {
              return makeTextCell(rowNumber, colNumber, "", 2);
            }

            const calculatedField = calculatedFields.get(field.key);
            if (calculatedField) {
              return makeFormulaCell(rowNumber, colNumber, calculatedField, CALCULATED_STYLE_ID);
            }
            const value = sharedRemarkField && field.key === sharedRemarkField.key && rowIndex > 0 && !custom
              ? ""
              : values[field.key];
            return makeDataCell(rowNumber, colNumber, field, value, custom);
          })
          .join("");
        if (customRemarkMerge) {
          customMergeRefs.push(`${cellRef(rowNumber, sharedRemarkIndex + 1)}:${cellRef(rowNumber, columnCount)}`);
        }
        const rowHeight = rowHasImage(worksheetImages, rowNumber) ? 78 : custom ? 42 : 0;
        const customHeight = rowHeight ? ` ht="${rowHeight}" customHeight="1"` : "";
        return `<row r="${rowNumber}"${customHeight}>${cells}</row>`;
      })
      .join("");
    const cols = fields
      .map((field, index) => {
        const width = field.type === "image"
          ? 19
          : Math.min(Math.max(String(field.label || "").length * 2 + 8, 14), 32);
        return `<col min="${index + 1}" max="${index + 1}" width="${width}" customWidth="1"/>`;
      })
      .join("");
    const drawing = worksheetImages.length ? '<drawing r:id="rId1"/>' : "";
    const sharedRemarkMergeRefs = buildSharedRemarkMergeRefs(rows, dataStartRow, sharedRemarkIndex);

    return xmlDecl(`\
<worksheet xmlns="${XML_NS_MAIN}" xmlns:r="${XML_NS_REL}">\
<dimension ref="A1:${lastCell}"/>\
<sheetViews><sheetView workbookViewId="0"><pane ySplit="4" topLeftCell="A5" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft"/></sheetView></sheetViews>\
<sheetFormatPr defaultRowHeight="18"/>\
<cols>${cols}</cols>\
<sheetData>${titleInfo.rows}${headerInfo.rows}${dataRows}${summaryInfo.rows}${noticeInfo.rows}</sheetData>\
${mergeBlocks(titleInfo.mergeRefs.concat(headerInfo.mergeRefs, sharedRemarkMergeRefs, customMergeRefs, summaryInfo.mergeRefs, noticeInfo.mergeRefs))}\
<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>\
${drawing}\
</worksheet>`);
  }

  function buildSharedRemarkMergeRefs(rows, dataStartRow, sharedRemarkIndex) {
    if (sharedRemarkIndex < 0 || rows.length < 2) {
      return [];
    }

    const refs = [];
    let startRow = null;
    let endRow = null;
    rows.forEach((row, index) => {
      const rowNumber = dataStartRow + index;
      if (isCustomRow(row)) {
        if (startRow != null && endRow > startRow) {
          refs.push(`${cellRef(startRow, sharedRemarkIndex + 1)}:${cellRef(endRow, sharedRemarkIndex + 1)}`);
        }
        startRow = null;
        endRow = null;
        return;
      }

      if (startRow == null) {
        startRow = rowNumber;
      }
      endRow = rowNumber;
    });

    if (startRow != null && endRow > startRow) {
      refs.push(`${cellRef(startRow, sharedRemarkIndex + 1)}:${cellRef(endRow, sharedRemarkIndex + 1)}`);
    }

    return refs;
  }

  function buildNoticeRow(rowNumber, columnCount) {
    const cells = [makeTextCell(rowNumber, 1, EXPORT_NOTICE_TEXT, 2)];
    for (let colNumber = 2; colNumber <= columnCount; colNumber += 1) {
      cells.push(makeTextCell(rowNumber, colNumber, "", 2));
    }

    return {
      rows: `<row r="${rowNumber}" ht="42" customHeight="1">${cells.join("")}</row>`,
      mergeRefs: columnCount > 1 ? [`A${rowNumber}:${columnName(columnCount)}${rowNumber}`] : [],
    };
  }

  function buildSummaryRow(fields, rows, dataStartRow, rowNumber, depositAmount) {
    const columnCount = Math.max(fields.length, 1);
    const amountIndexes = fields
      .map((field, index) => isAmountField(field) ? index : -1)
      .filter((index) => index >= 0);
    const depositValue = normalizeNonNegativeNumberValue(depositAmount || "");
    const amountIndex = amountIndexes[0] != null ? amountIndexes[0] : -1;
    const depositLabelIndex = findDepositLabelIndex(columnCount, amountIndex);
    const depositValueIndex = depositLabelIndex >= 0 && depositLabelIndex + 1 < columnCount ? depositLabelIndex + 1 : -1;
    const totalFormula = makeTotalAmountFormula(amountIndexes, rows.length, dataStartRow);
    const cells = [];

    for (let fieldIndex = 0; fieldIndex < columnCount; fieldIndex += 1) {
      const colNumber = fieldIndex + 1;
      if (fieldIndex === 0) {
        cells.push(makeTextCell(rowNumber, colNumber, "总金额", CUSTOM_TEXT_STYLE_ID));
      } else if (fieldIndex === amountIndex) {
        cells.push(totalFormula
          ? `<c r="${cellRef(rowNumber, colNumber)}" s="${CUSTOM_CALCULATED_STYLE_ID}"><f>${escapeXml(totalFormula)}</f></c>`
          : makeDataCell(rowNumber, colNumber, { type: "number" }, "", true));
      } else if (fieldIndex === depositLabelIndex) {
        cells.push(makeTextCell(rowNumber, colNumber, "定金", CUSTOM_TEXT_STYLE_ID));
      } else if (fieldIndex === depositValueIndex) {
        cells.push(makeDataCell(rowNumber, colNumber, { type: "number" }, depositValue, true));
      } else {
        cells.push(makeTextCell(rowNumber, colNumber, "", CUSTOM_TEXT_STYLE_ID));
      }
    }

    return {
      rows: `<row r="${rowNumber}" ht="28" customHeight="1">${cells.join("")}</row>`,
      mergeRefs: [],
    };
  }

  function findDepositLabelIndex(columnCount, amountIndex) {
    if (columnCount < 2) {
      return -1;
    }

    const candidates = amountIndex + 2 < columnCount
      ? [amountIndex + 1, columnCount - 2, 1]
      : [columnCount - 2, 1];
    return candidates.find((index) => index >= 0 && index + 1 < columnCount && index !== amountIndex && index + 1 !== amountIndex) ?? -1;
  }

  function makeTotalAmountFormula(amountIndexes, dataRowCount, dataStartRow) {
    if (!amountIndexes.length || !dataRowCount) {
      return "";
    }

    const dataEndRow = dataStartRow + dataRowCount - 1;
    return amountIndexes
      .map((fieldIndex) => {
        const colName = columnName(fieldIndex + 1);
        return `SUM(${colName}${dataStartRow}:${colName}${dataEndRow})`;
      })
      .join("+");
  }

  function buildTitleRows(documentName, columnCount) {
    const lastColumn = columnName(columnCount);
    const labelCells = [makeTextCell(1, 1, "名称", 6)];
    const valueCells = [makeTextCell(2, 1, String(documentName || ""), 7)];
    for (let colNumber = 2; colNumber <= columnCount; colNumber += 1) {
      labelCells.push(makeTextCell(1, colNumber, "", 6));
      valueCells.push(makeTextCell(2, colNumber, "", 7));
    }
    return {
      rows: `<row r="1" ht="24" customHeight="1">${labelCells.join("")}</row><row r="2" ht="26" customHeight="1">${valueCells.join("")}</row>`,
      mergeRefs: [`A1:${lastColumn}1`, `A2:${lastColumn}2`],
    };
  }

  function buildHeaderRows(fields, startRow) {
    const topCells = [];
    const secondCells = [];
    const mergeRefs = [];
    const topRow = startRow || 1;
    const secondRow = topRow + 1;
    let index = 0;

    while (index < fields.length) {
      const field = fields[index];
      const group = String(field.group || "").trim();
      const colNumber = index + 1;

      if (!group) {
        topCells.push(makeTextCell(topRow, colNumber, field.label, 1));
        secondCells.push(makeTextCell(secondRow, colNumber, "", 1));
        mergeRefs.push(`${cellRef(topRow, colNumber)}:${cellRef(secondRow, colNumber)}`);
        index += 1;
        continue;
      }

      let endIndex = index;
      while (endIndex + 1 < fields.length && String(fields[endIndex + 1].group || "").trim() === group) {
        endIndex += 1;
      }

      topCells.push(makeTextCell(topRow, colNumber, group, 1));
      for (let blankIndex = index + 1; blankIndex <= endIndex; blankIndex += 1) {
        topCells.push(makeTextCell(topRow, blankIndex + 1, "", 1));
      }
      for (let childIndex = index; childIndex <= endIndex; childIndex += 1) {
        secondCells.push(makeTextCell(secondRow, childIndex + 1, fields[childIndex].label, 1));
      }
      if (endIndex > index) {
        mergeRefs.push(`${cellRef(topRow, colNumber)}:${cellRef(topRow, endIndex + 1)}`);
      }
      index = endIndex + 1;
    }

    const headerRows = `<row r="${topRow}" ht="24" customHeight="1">${topCells.join("")}</row><row r="${secondRow}" ht="24" customHeight="1">${secondCells.join("")}</row>`;
    return { rows: headerRows, mergeRefs };
  }

  function mergeBlocks(mergeRefs) {
    return mergeRefs.length
      ? `<mergeCells count="${mergeRefs.length}">${mergeRefs.map((ref) => `<mergeCell ref="${ref}"/>`).join("")}</mergeCells>`
      : "";
  }

  function getCustomRowExportMap(fields) {
    const sourceFields = fields || [];
    const amountField = findCustomRowAmountField(sourceFields);
    const amountIndex = amountField ? sourceFields.indexOf(amountField) : -1;
    const beforeAmountFields = amountIndex > -1 ? sourceFields.slice(0, amountIndex) : sourceFields;
    const amountGroup = amountField && normalizeLabel(amountField.group);
    const groupFields = amountGroup ? beforeAmountFields.filter((field) => sameGroup(field, amountField)) : beforeAmountFields;
    const scopedFields = groupFields.length ? groupFields : beforeAmountFields;

    const quantityField = findLast(scopedFields, isQuantityField) || findLast(beforeAmountFields, isQuantityField);
    const unitPriceField = findLast(scopedFields, (field) => isUnitPriceField(field) && field !== amountField) ||
      findLast(beforeAmountFields, (field) => isUnitPriceField(field) && field !== amountField);
    const typeField = findLast(scopedFields, isCustomTypeFallbackField) || findLast(beforeAmountFields, isCustomTypeFallbackField) || null;

    return {
      fields: sourceFields,
      typeField,
      typeIndex: typeField ? sourceFields.indexOf(typeField) : -1,
      quantityField: quantityField || null,
      quantityIndex: quantityField ? sourceFields.indexOf(quantityField) : -1,
      unitPriceField: unitPriceField || null,
      unitPriceIndex: unitPriceField ? sourceFields.indexOf(unitPriceField) : -1,
      amountField,
      amountIndex,
    };
  }

  function findCustomRowAmountField(fields) {
    const sourceFields = fields || [];
    const amountFields = sourceFields.filter(isAmountField);
    return amountFields.find((field) => {
      const index = sourceFields.indexOf(field);
      const previousFields = sourceFields.slice(0, index);
      const sameGroupFields = normalizeLabel(field.group)
        ? previousFields.filter((candidate) => sameGroup(candidate, field))
        : previousFields;
      const scopedFields = sameGroupFields.length ? sameGroupFields : previousFields;
      return findLast(scopedFields, isQuantityField) &&
        findLast(scopedFields, (candidate) => isUnitPriceField(candidate) && candidate !== field);
    }) || amountFields[0] || null;
  }

  function makeCustomRowExportCell(rowNumber, colNumber, field, fieldIndex, values, exportMap) {
    if (exportMap.typeField && field.key === exportMap.typeField.key) {
      return makeDataCell(rowNumber, colNumber, CUSTOM_ROW_FIELDS[0], values[CUSTOM_ROW_TYPE_KEY], true);
    }

    if (exportMap.quantityField && field.key === exportMap.quantityField.key) {
      return makeDataCell(rowNumber, colNumber, CUSTOM_ROW_FIELDS[1], values[CUSTOM_ROW_QUANTITY_KEY], true);
    }

    if (exportMap.unitPriceField && field.key === exportMap.unitPriceField.key) {
      return makeDataCell(rowNumber, colNumber, CUSTOM_ROW_FIELDS[2], values[CUSTOM_ROW_UNIT_PRICE_KEY], true);
    }

    if (exportMap.amountField && field.key === exportMap.amountField.key) {
      if (exportMap.quantityIndex > -1 && exportMap.unitPriceIndex > -1) {
        return makeFormulaCell(
          rowNumber,
          colNumber,
          {
            type: "quantityAmount",
            sourceIndexes: [exportMap.quantityIndex, exportMap.unitPriceIndex],
          },
          CUSTOM_CALCULATED_STYLE_ID
        );
      }
      return makeDataCell(rowNumber, colNumber, CUSTOM_ROW_FIELDS[3], values[CUSTOM_ROW_AMOUNT_KEY], true);
    }

    return "";
  }

  function buildStylesXml() {
    return xmlDecl(`\
<styleSheet xmlns="${XML_NS_MAIN}">\
<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>\
<fonts count="4">\
<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>\
<font><b/><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/><family val="2"/></font>\
<font><b/><sz val="14"/><color rgb="FF000000"/><name val="Calibri"/><family val="2"/></font>\
<font><b/><sz val="12"/><color rgb="FF000000"/><name val="Calibri"/><family val="2"/></font>\
</fonts>\
<fills count="4">\
<fill><patternFill patternType="none"/></fill>\
<fill><patternFill patternType="gray125"/></fill>\
<fill><patternFill patternType="solid"><fgColor rgb="FFC0C0C0"/><bgColor indexed="64"/></patternFill></fill>\
<fill><patternFill patternType="solid"><fgColor rgb="FFE7EAEE"/><bgColor indexed="64"/></patternFill></fill>\
</fills>\
<borders count="2">\
<border><left/><right/><top/><bottom/><diagonal/></border>\
<border><left style="thin"><color rgb="FFD8E0DE"/></left><right style="thin"><color rgb="FFD8E0DE"/></right><top style="thin"><color rgb="FFD8E0DE"/></top><bottom style="thin"><color rgb="FFD8E0DE"/></bottom><diagonal/></border>\
</borders>\
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>\
<cellXfs count="12">\
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>\
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>\
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>\
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>\
<xf numFmtId="2" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>\
<xf numFmtId="2" fontId="0" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>\
<xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>\
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>\
<xf numFmtId="0" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>\
<xf numFmtId="164" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>\
<xf numFmtId="2" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>\
<xf numFmtId="2" fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>\
</cellXfs>\
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>\
<dxfs count="0"/>\
<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>\
</styleSheet>`);
  }

  function makeDataCell(rowNumber, colNumber, field, rawValue, bold) {
    const rawText = rawValue == null ? "" : String(rawValue);
    const value = field.type === "number" ? normalizeNonNegativeNumberValue(rawText) : rawText;
    const textStyleId = bold ? CUSTOM_TEXT_STYLE_ID : 2;
    const dateStyleId = bold ? CUSTOM_DATE_STYLE_ID : 3;
    const numberStyleId = bold ? CUSTOM_NUMBER_STYLE_ID : 4;
    if (field.type === "image") {
      return makeTextCell(rowNumber, colNumber, "", textStyleId);
    }

    if (value === "") {
      return makeTextCell(rowNumber, colNumber, "", textStyleId);
    }

    if (field.type === "number") {
      const numericValue = Number(value);
      if (Number.isFinite(numericValue)) {
        return `<c r="${cellRef(rowNumber, colNumber)}" s="${numberStyleId}"><v>${numericValue}</v></c>`;
      }
    }

    if (field.type === "date" && /^\d{4}-\d{2}-\d{2}$/.test(value)) {
      return `<c r="${cellRef(rowNumber, colNumber)}" s="${dateStyleId}"><v>${dateToExcelSerial(value)}</v></c>`;
    }

    return makeTextCell(rowNumber, colNumber, value, textStyleId);
  }

  function makeFormulaCell(rowNumber, colNumber, calculatedField, styleId) {
    const [firstSourceIndex, secondSourceIndex] = calculatedField.sourceIndexes;
    const firstRef = cellRef(rowNumber, firstSourceIndex + 1);
    const secondRef = cellRef(rowNumber, secondSourceIndex + 1);
    const operator = calculatedField.type === "sum" ? "+" : "*";
    const formula = `IF(OR(${firstRef}="",${secondRef}=""),"",${firstRef}${operator}${secondRef})`;
    return `<c r="${cellRef(rowNumber, colNumber)}" s="${styleId || CALCULATED_STYLE_ID}"><f>${escapeXml(formula)}</f></c>`;
  }

  function makeTextCell(rowNumber, colNumber, value, styleId) {
    return `<c r="${cellRef(rowNumber, colNumber)}" t="inlineStr" s="${styleId}"><is><t${needsPreserveSpace(value) ? ' xml:space="preserve"' : ""}>${escapeXml(value)}</t></is></c>`;
  }

  function dateToExcelSerial(dateValue) {
    const [year, month, day] = dateValue.split("-").map(Number);
    const utcDate = Date.UTC(year, month - 1, day);
    const epoch = Date.UTC(1899, 11, 30);
    return Math.round((utcDate - epoch) / 86400000);
  }

  function formatTimestamp(date) {
    const pad = (value) => String(value).padStart(2, "0");
    return `${date.getFullYear()}${pad(date.getMonth() + 1)}${pad(date.getDate())}_${pad(date.getHours())}${pad(date.getMinutes())}`;
  }

  function xmlDecl(xml) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${xml}`;
  }

  function escapeXml(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  function escapeHtml(value) {
    return escapeXml(value);
  }

  function escapeAttr(value) {
    return escapeXml(value);
  }

  function needsPreserveSpace(value) {
    return /^\s|\s$/.test(String(value));
  }

  function cellRef(rowNumber, colNumber) {
    return `${columnName(colNumber)}${rowNumber}`;
  }

  function columnName(colNumber) {
    let name = "";
    let number = colNumber;
    while (number > 0) {
      const remainder = (number - 1) % 26;
      name = String.fromCharCode(65 + remainder) + name;
      number = Math.floor((number - 1) / 26);
    }
    return name;
  }

  function createZip(files) {
    const encoder = typeof TextEncoder === "function" ? new TextEncoder() : null;
    const localParts = [];
    const centralParts = [];
    let offset = 0;

    files.forEach((file) => {
      const nameBytes = encodeUtf8(file.name, encoder);
      const dataBytes = typeof file.content === "string" ? encodeUtf8(file.content, encoder) : file.content;
      const crc = crc32(dataBytes);
      const timestamp = dosTimestamp(new Date());
      const localHeader = new Uint8Array(30 + nameBytes.length);
      const localView = new DataView(localHeader.buffer);
      localView.setUint32(0, 0x04034b50, true);
      localView.setUint16(4, 20, true);
      localView.setUint16(6, 0x0800, true);
      localView.setUint16(8, 0, true);
      localView.setUint16(10, timestamp.time, true);
      localView.setUint16(12, timestamp.date, true);
      localView.setUint32(14, crc, true);
      localView.setUint32(18, dataBytes.length, true);
      localView.setUint32(22, dataBytes.length, true);
      localView.setUint16(26, nameBytes.length, true);
      localView.setUint16(28, 0, true);
      localHeader.set(nameBytes, 30);
      localParts.push(localHeader, dataBytes);

      const centralHeader = new Uint8Array(46 + nameBytes.length);
      const centralView = new DataView(centralHeader.buffer);
      centralView.setUint32(0, 0x02014b50, true);
      centralView.setUint16(4, 0x0314, true);
      centralView.setUint16(6, 20, true);
      centralView.setUint16(8, 0x0800, true);
      centralView.setUint16(10, 0, true);
      centralView.setUint16(12, timestamp.time, true);
      centralView.setUint16(14, timestamp.date, true);
      centralView.setUint32(16, crc, true);
      centralView.setUint32(20, dataBytes.length, true);
      centralView.setUint32(24, dataBytes.length, true);
      centralView.setUint16(28, nameBytes.length, true);
      centralView.setUint16(30, 0, true);
      centralView.setUint16(32, 0, true);
      centralView.setUint16(34, 0, true);
      centralView.setUint16(36, 0, true);
      centralView.setUint32(38, 0, true);
      centralView.setUint32(42, offset, true);
      centralHeader.set(nameBytes, 46);
      centralParts.push(centralHeader);

      offset += localHeader.length + dataBytes.length;
    });

    const centralOffset = offset;
    const centralSize = centralParts.reduce((sum, part) => sum + part.length, 0);
    const endRecord = new Uint8Array(22);
    const endView = new DataView(endRecord.buffer);
    endView.setUint32(0, 0x06054b50, true);
    endView.setUint16(4, 0, true);
    endView.setUint16(6, 0, true);
    endView.setUint16(8, files.length, true);
    endView.setUint16(10, files.length, true);
    endView.setUint32(12, centralSize, true);
    endView.setUint32(16, centralOffset, true);
    endView.setUint16(20, 0, true);

    const allParts = localParts.concat(centralParts, [endRecord]);
    const totalLength = allParts.reduce((sum, part) => sum + part.length, 0);
    const output = new Uint8Array(totalLength);
    let cursor = 0;
    allParts.forEach((part) => {
      output.set(part, cursor);
      cursor += part.length;
    });
    return output;
  }

  function encodeUtf8(value, encoder) {
    if (encoder) {
      return encoder.encode(value);
    }

    const text = unescape(encodeURIComponent(String(value)));
    const bytes = new Uint8Array(text.length);
    for (let index = 0; index < text.length; index += 1) {
      bytes[index] = text.charCodeAt(index);
    }
    return bytes;
  }

  function dosTimestamp(date) {
    const safeDate = date < ZIP_EPOCH ? ZIP_EPOCH : date;
    return {
      time: (safeDate.getHours() << 11) | (safeDate.getMinutes() << 5) | Math.floor(safeDate.getSeconds() / 2),
      date: ((safeDate.getFullYear() - 1980) << 9) | ((safeDate.getMonth() + 1) << 5) | safeDate.getDate(),
    };
  }

  const CRC_TABLE = (() => {
    const table = new Uint32Array(256);
    for (let i = 0; i < 256; i += 1) {
      let c = i;
      for (let k = 0; k < 8; k += 1) {
        c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      }
      table[i] = c >>> 0;
    }
    return table;
  })();

  function crc32(bytes) {
    let crc = 0xffffffff;
    for (let i = 0; i < bytes.length; i += 1) {
      crc = CRC_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
    }
    return (crc ^ 0xffffffff) >>> 0;
  }

  root.ExcelAutoTool = {
    createXlsxBlob,
    buildWorkbookFiles,
    collectWorksheetImages,
    calculateFieldValue,
    createNormalRow,
    createCustomRow,
    normalizeAppRow,
    getRowValues,
    isCustomRow,
    getCalculatedFields,
    normalizeFields,
    dateToExcelSerial,
    columnName,
  };

  function showStartupError(message) {
    const status = typeof document !== "undefined" ? document.getElementById("status") : null;
    if (!status) {
      return;
    }
    status.textContent = message;
    status.className = "status status-static is-error";
  }

  if (typeof document !== "undefined") {
    document.addEventListener("DOMContentLoaded", () => {
      try {
        init();
      } catch (error) {
        showStartupError(`页面加载失败：${error.message || "请刷新后重试"}`);
      }
    });
  }
})();
