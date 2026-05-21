import { downloadBlobLocally, shouldUseServerExportDownload, createServerExportDownload } from "./export-download.js";
import { requestNaturalFillExtraction } from "./natural-fill-api.js";
import { createXlsxBuilder } from "./xlsx-builder.js";
import {
  CALCULATION_OPERATORS,
  CALCULATION_TYPE_OPERATORS,
  CUSTOM_ROW_AMOUNT_KEY,
  CUSTOM_ROW_FIELDS,
  CUSTOM_ROW_QUANTITY_KEY,
  CUSTOM_ROW_TYPE_KEY,
  CUSTOM_ROW_UNIT_PRICE_KEY,
  DEFAULT_FIELDS,
  DRAFT_DB_NAME,
  DRAFT_DB_VERSION,
  DRAFT_SAVE_DELAY,
  DRAFT_STORE_NAME,
  DRAFT_VERSION,
  FIELD_TYPES,
  MAX_DRAFTS,
  STORAGE_KEY,
  SUPPORTED_IMAGE_MIMES,
  TYPE_TEXT,
} from "./app-config.js";

export function startExcelAutoTool() {
(function () {
  "use strict";

  const root = typeof window !== "undefined" ? window : globalThis;
  let state = {
    documentName: "",
    depositAmount: "",
    sharedRemark: "",
    fields: [],
    rows: [],
    invalidCells: new Set(),
    invalidSharedFields: new Set(),
    invalidFields: new Set(),
    calculationRules: [],
    disabledAutoCalculationTargets: [],
    configOpen: false,
    drafts: [],
    activeDraftId: "",
    draftReady: false,
    draftStorageAvailable: false,
    naturalFillRunning: false,
  };

  let els = {};
  let statusTimer = 0;

  const xlsxBuilder = createXlsxBuilder({
    calculateFieldValue,
    displayFieldLabel,
    findLast,
    formatCalculatedValue,
    getCalculatedFields,
    getRowValues,
    getSharedRemarkField,
    isAmountField,
    isCustomRow,
    isCustomTypeFallbackField,
    isMeterField,
    isQuantityField,
    isUnitPriceField,
    normalizeAppRow,
    normalizeCalculationRules,
    normalizeDisabledAutoCalculationTargets,
    normalizeFields,
    normalizeLabel,
    normalizeNonNegativeNumberValue,
    sameGroup,
  });
  const {
    buildWorkbookFiles,
    collectWorksheetImages,
    createXlsxBlob,
    dateToExcelSerial,
    columnName,
  } = xlsxBuilder;
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

  function normalizeCalculationRules(rules, fields) {
    const targetFields = fields || state.fields;
    const fieldKeys = new Set((targetFields || []).map((field) => field.key));
    return (Array.isArray(rules) ? rules : [])
      .map((rule) => {
        if (!rule || typeof rule !== "object") {
          return null;
        }

        const targetKey = String(rule.targetKey || "").trim();
        const sourceKeys = Array.isArray(rule.sourceKeys)
          ? rule.sourceKeys.map((key) => String(key || "").trim()).filter(Boolean)
          : [];
        const operator = CALCULATION_OPERATORS.includes(rule.operator) ? rule.operator : "";
        if (!fieldKeys.has(targetKey) || sourceKeys.length !== 2 || !operator) {
          return null;
        }
        if (!sourceKeys.every((key) => fieldKeys.has(key)) || sourceKeys.includes(targetKey)) {
          return null;
        }

        return { targetKey, sourceKeys, operator };
      })
      .filter(Boolean);
  }

  function normalizeDisabledAutoCalculationTargets(targets, fields) {
    const fieldKeys = new Set((fields || state.fields || []).map((field) => field.key));
    return Array.from(new Set((Array.isArray(targets) ? targets : [])
      .map((key) => String(key || "").trim())
      .filter((key) => fieldKeys.has(key))));
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
    const calculationRules = normalizeCalculationRules(draft.calculationRules, fields);
    const disabledAutoCalculationTargets = normalizeDisabledAutoCalculationTargets(draft.disabledAutoCalculationTargets, fields);

    return {
      id: String(draft.id || makeDraftId()),
      name: draftDisplayName(draft.name || documentName),
      updatedAt,
      documentName,
      depositAmount,
      sharedRemark: String(draft.sharedRemark || ""),
      fields,
      rows,
      calculationRules,
      disabledAutoCalculationTargets,
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
    const calculationRules = normalizeCalculationRules(state.calculationRules, fields);
    const disabledAutoCalculationTargets = normalizeDisabledAutoCalculationTargets(state.disabledAutoCalculationTargets, fields);

    return {
      id: state.activeDraftId || makeDraftId(),
      name: draftDisplayName(documentName),
      updatedAt,
      documentName,
      depositAmount,
      sharedRemark: String(state.sharedRemark || ""),
      fields,
      rows,
      calculationRules,
      disabledAutoCalculationTargets,
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
    state.calculationRules = normalizedDraft.calculationRules;
    state.disabledAutoCalculationTargets = normalizedDraft.disabledAutoCalculationTargets;
    clearValidationState();
    ensureSharedRemarkState();
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
    state.fields = loadFields();
    state.calculationRules = [];
    state.disabledAutoCalculationTargets = [];
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
    const calculatedFields = getCalculatedFields(state.fields, state.calculationRules);
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
      pruneCalculationRules();
      saveFields();
      render();
      scheduleDraftSave();
      return;
    }

    if (prop === "required") {
      state.fields[index].required = target.checked;
      saveFields();
      renderSharedRemark();
      render();
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
      removeFieldByKey(state.fields[index].key);
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
      render();
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

  function pruneCalculationRules() {
    state.calculationRules = normalizeCalculationRules(state.calculationRules, state.fields);
    state.disabledAutoCalculationTargets = normalizeDisabledAutoCalculationTargets(state.disabledAutoCalculationTargets, state.fields);
  }

  function removeFieldByKey(fieldKey) {
    const index = state.fields.findIndex((field) => field.key === fieldKey);
    if (index < 0) {
      return null;
    }

    const [removed] = state.fields.splice(index, 1);
    state.rows.forEach((row) => {
      delete getRowValues(row)[removed.key];
    });
    syncSharedRemarkFromRemovedField(removed);
    pruneCalculationRules();
    state.invalidCells.clear();
    state.invalidSharedFields.clear();
    state.invalidFields.delete(removed.key);
    return removed;
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

  function getClientCalculationRules() {
    return normalizeCalculationRules(state.calculationRules, state.fields).map((rule) => ({ ...rule }));
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
      setStatus("没有可智能处理的表头", "warning");
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
      setStatus("没有可智能处理的表头", "warning");
      return;
    }

    state.naturalFillRunning = true;
    updateNaturalFillControls();
    setStatus("正在根据描述填写表格...", "success");

    try {
      const result = await requestNaturalFillExtraction(text, fields, getClientCalculationRules());
      const fieldChangeCount = applyNaturalFillFieldChanges(result.fieldChanges || []);
      const ruleChangeCount = applyNaturalFillCalculationRuleChanges(
        result.calculationRuleChanges || result.calculationRules || []
      );
      const addedCount = appendNaturalFillRows(result.rows || []);
      if (!addedCount && !fieldChangeCount && !ruleChangeCount) {
        const hasStructuralIntent = hasNaturalFillStructuralIntent(text);
        setStatus(
          hasStructuralIntent
            ? "未解析到表头或规则变更，请确认线上后端已部署最新版后重试"
            : "未解析到可填写的数据",
          "warning"
        );
        return;
      }

      render();
      scheduleDraftSave();
      els.naturalFillText.value = "";
      closeNaturalFillPanel();
      setStatus(buildNaturalFillSuccessMessage(addedCount, fieldChangeCount, ruleChangeCount), result.warnings && result.warnings.length ? "warning" : "success");
    } catch (error) {
      setStatus(`智能填行失败：${error.message || "请稍后重试"}`, "error");
    } finally {
      state.naturalFillRunning = false;
      updateNaturalFillControls();
    }
  }

  function getNaturalFillFields() {
    return state.fields
      .map((field) => ({
        key: field.key,
        label: field.label,
        group: field.group,
        type: field.type,
        options: field.options,
        required: field.required,
      }));
  }

  function hasNaturalFillStructuralIntent(text) {
    return /(删除|删掉|移除|去掉|取消|不要|新增|添加|增加|改名|改成|修改|更改|规则|计算|公式|自动计算|字段|表头)/.test(String(text || ""));
  }

  function applyNaturalFillFieldChanges(changes) {
    let changedCount = 0;
    const normalizedChanges = Array.isArray(changes) ? changes : [];

    normalizedChanges.forEach((change) => {
      if (!change || change.action !== "delete") {
        return;
      }

      if (removeFieldByKey(String(change.key || "").trim())) {
        changedCount += 1;
      }
    });

    normalizedChanges.forEach((change) => {
      if (!change || typeof change !== "object") {
        return;
      }

      if (change.action === "delete") {
        return;
      }

      if (change.action === "update") {
        const field = state.fields.find((item) => item.key === change.key);
        if (!field) {
          return;
        }
        if (typeof change.label === "string" && change.label.trim()) {
          field.label = change.label.trim();
        }
        if (Object.prototype.hasOwnProperty.call(change, "group")) {
          field.group = String(change.group || "");
        }
        if (FIELD_TYPES.includes(change.type) && change.type !== "image") {
          field.type = change.type;
        }
        if (Array.isArray(change.options)) {
          field.options = change.options.map(String).map((item) => item.trim()).filter(Boolean);
        }
        if (Object.prototype.hasOwnProperty.call(change, "required")) {
          field.required = Boolean(change.required);
        }
        changedCount += 1;
        return;
      }

      if (change.action !== "add") {
        return;
      }

      const label = String(change.label || "").trim();
      if (!label) {
        return;
      }

      state.fields.push({
        key: makeUniqueFieldKey(change.key || label),
        label,
        group: String(change.group || ""),
        type: FIELD_TYPES.includes(change.type) && change.type !== "image" ? change.type : "text",
        required: Boolean(change.required),
        options: Array.isArray(change.options)
          ? change.options.map(String).map((item) => item.trim()).filter(Boolean)
          : [],
      });
      changedCount += 1;
    });

    if (changedCount) {
      state.fields = normalizeFields(state.fields);
      state.rows = state.rows.map((row) => normalizeAppRow(row, state.fields));
      state.calculationRules = normalizeCalculationRules(state.calculationRules, state.fields);
      state.disabledAutoCalculationTargets = normalizeDisabledAutoCalculationTargets(state.disabledAutoCalculationTargets, state.fields);
      ensureSharedRemarkState();
      syncRowsWithCurrentCalculations();
    }

    return changedCount;
  }

  function makeUniqueFieldKey(value) {
    const base = String(value || "field")
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9_]+/g, "_")
      .replace(/^_+|_+$/g, "")
      .slice(0, 32);
    let key = base && !state.fields.some((field) => field.key === base) ? base : makeFieldKey(base || value || "field");
    while (state.fields.some((field) => field.key === key)) {
      key = makeFieldKey(base || value || "field");
    }
    return key;
  }

  function applyNaturalFillCalculationRuleChanges(changes) {
    const rawChanges = Array.isArray(changes) ? changes : [];
    const fieldKeys = new Set(state.fields.map((field) => field.key));
    const rulesByTarget = new Map(normalizeCalculationRules(state.calculationRules, state.fields).map((rule) => [rule.targetKey, rule]));
    const disabledTargets = new Set(normalizeDisabledAutoCalculationTargets(state.disabledAutoCalculationTargets, state.fields));
    let changedCount = 0;

    rawChanges.forEach((change) => {
      if (!change || typeof change !== "object") {
        return;
      }

      const action = change.action === "delete" ? "delete" : change.action === "set" || !change.action ? "set" : "";
      const targetKey = String(change.targetKey || "").trim();
      if (!action || !fieldKeys.has(targetKey)) {
        return;
      }

      if (action === "delete") {
        rulesByTarget.delete(targetKey);
        disabledTargets.add(targetKey);
        changedCount += 1;
        return;
      }

      const normalizedRule = normalizeCalculationRules([change], state.fields)[0];
      if (!normalizedRule) {
        return;
      }

      rulesByTarget.set(normalizedRule.targetKey, normalizedRule);
      disabledTargets.delete(normalizedRule.targetKey);
      changedCount += 1;
    });

    if (!changedCount) {
      return 0;
    }

    state.calculationRules = normalizeCalculationRules(Array.from(rulesByTarget.values()), state.fields);
    state.disabledAutoCalculationTargets = normalizeDisabledAutoCalculationTargets(Array.from(disabledTargets), state.fields);
    syncRowsWithCurrentCalculations();
    return changedCount;
  }

  function syncRowsWithCurrentCalculations() {
    const calculatedFields = getCalculatedFields(state.fields, state.calculationRules);
    state.rows.forEach((row) => syncCalculatedRow(row, calculatedFields));
  }

  function buildNaturalFillSuccessMessage(addedCount, fieldChangeCount, ruleChangeCount) {
    const parts = [];
    if (addedCount) {
      parts.push(`填入 ${addedCount} 行`);
    }
    if (fieldChangeCount) {
      parts.push(`更新 ${fieldChangeCount} 个表头`);
    }
    if (ruleChangeCount) {
      parts.push(`应用 ${ruleChangeCount} 条规则`);
    }
    return `已智能处理：${parts.join("，")}`;
  }

  function appendNaturalFillRows(rows) {
    const fieldsByKey = new Map(state.fields.map((field) => [field.key, field]));
    const sharedRemarkField = getSharedRemarkField(state.fields);
    const calculatedFields = getCalculatedFields(state.fields, state.calculationRules);
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
      const blob = await createXlsxBlob(
        state.fields,
        state.rows,
        state.documentName,
        state.depositAmount,
        state.calculationRules,
        state.disabledAutoCalculationTargets
      );
      const filename = `自动生成表格_${formatTimestamp(new Date())}.xlsx`;
      await downloadBlob(blob, filename);
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

  function formatTimestamp(date) {
    const pad = (value) => String(value).padStart(2, "0");
    return `${date.getFullYear()}${pad(date.getMonth() + 1)}${pad(date.getDate())}_${pad(date.getHours())}${pad(date.getMinutes())}`;
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

  function getCalculatedFields(fields, calculationRules, disabledAutoTargets) {
    const calculatedFields = new Map();
    const disabledTargets = new Set(normalizeDisabledAutoCalculationTargets(disabledAutoTargets || state.disabledAutoCalculationTargets, fields));
    const widthField = fields.find(isWidthField);
    const heightField = fields.find(isHeightField);
    const meterField = fields.find(isMeterField);

    if (widthField && heightField && meterField && !disabledTargets.has(meterField.key)) {
      calculatedFields.set(meterField.key, makeCalculatedField("sum", fields, widthField, heightField));
    }

    fields.forEach((field, index) => {
      if (!isAmountField(field)) {
        return;
      }
      if (disabledTargets.has(field.key)) {
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

    normalizeCalculationRules(calculationRules || state.calculationRules, fields).forEach((rule) => {
      const sourceFields = rule.sourceKeys.map((key) => fields.find((field) => field.key === key));
      if (sourceFields.every(Boolean)) {
        calculatedFields.set(rule.targetKey, makeCalculatedField(rule.operator, fields, sourceFields[0], sourceFields[1]));
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
      operator: CALCULATION_TYPE_OPERATORS[type] || type,
      sourceKeys: [firstField.key, secondField.key],
      sourceIndexes: [fields.indexOf(firstField), fields.indexOf(secondField)],
    };
  }

  function makeCustomRowAmountField() {
    return {
      type: "quantityAmount",
      operator: "multiply",
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
    const operator = calculatedField.operator || CALCULATION_TYPE_OPERATORS[calculatedField.type] || "multiply";
    if (operator === "add") {
      return values[0] + values[1];
    }
    if (operator === "subtract") {
      return values[0] - values[1];
    }
    if (operator === "divide") {
      return values[1] === 0 ? "" : values[0] / values[1];
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
    const fieldsToCalculate = calculatedFields || getCalculatedFields(state.fields, state.calculationRules);
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

    const calculatedFields = getCalculatedFields(state.fields, state.calculationRules);
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

  async function downloadBlob(blob, filename) {
    if (shouldUseServerExportDownload()) {
      try {
        const downloadUrl = await createServerExportDownload(blob, filename);
        root.location.href = downloadUrl;
        return;
      } catch (error) {
        setStatus("正在使用浏览器下载方式，若无法打开请换系统浏览器重试。", "warning");
      }
    } else if (root.navigator && /MicroMessenger/i.test(root.navigator.userAgent || "")) {
      setStatus("微信内置浏览器可能拦截下载，请用系统浏览器打开后生成 Excel。", "warning");
    }

    downloadBlobLocally(blob, filename);
  }

  function shouldUseServerExportDownload() {
    return Boolean(root.location && /^https?:$/.test(root.location.protocol) && root.fetch);
  }

  async function createServerExportDownload(blob, filename) {
    const response = await fetch(`${EXPORT_ENDPOINT}?filename=${encodeURIComponent(filename)}`, {
      method: "POST",
      headers: {
        "Content-Type": MIME_XLSX,
      },
      body: blob,
    });
    let payload = null;
    try {
      payload = await response.json();
    } catch (error) {
      payload = null;
    }
    if (!response.ok || !payload || !payload.url) {
      throw new Error((payload && payload.error) || "无法创建下载链接");
    }
    return payload.url;
  }

  function downloadBlobLocally(blob, filename) {
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
    normalizeCalculationRules,
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

}
