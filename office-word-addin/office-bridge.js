const DOCX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

const officeElements = {
  sourceModeBadge: document.getElementById("sourceModeBadge"),
  sourceHint: document.getElementById("sourceHint"),
  officeAvailability: document.getElementById("officeAvailability"),
  officeEnvironmentValue: document.getElementById("officeEnvironmentValue"),
  officeEnvironmentHint: document.getElementById("officeEnvironmentHint"),
  officeActionNote: document.getElementById("officeActionNote"),
  loadCurrentDocumentButton: document.getElementById("loadCurrentDocumentButton"),
  applyToCurrentDocumentButton: document.getElementById("applyToCurrentDocumentButton"),
  langZhButton: document.getElementById("langZhButton"),
  langEnButton: document.getElementById("langEnButton"),
};

const OFFICE_COPY = {
  zh: {
    sourceModeEmpty: "未载入文档",
    sourceModeUpload: "上传的 DOCX",
    sourceModeWord: "当前 Word 文档",
    sourceHintIdle: "你可以先读取当前 Word 文档，也可以上传一个 DOCX 作为回退方案。",
    sourceHintUpload: "当前使用上传文件模式，适合 Word 网页版、浏览器预览或离线处理。",
    sourceHintWord: "当前使用 Word 当前文档模式，扫描完成后可以直接把修正结果写回文档。",
    availabilityPending: "正在检测 Word 宿主环境…",
    availabilityWordDesktop: "已连接 Word 桌面宿主，可直接读取当前文档并在预览确认后写回。",
    availabilityWordWeb: "已连接 Word 网页宿主。由于网页端无法直接读取整个 DOCX，请改用上传模式完成扫描和导出。",
    availabilityBrowser: "当前不在 Word 宿主环境中，仍可上传 DOCX 进行扫描、预览和导出。",
    environmentDesktop: "Word 桌面版",
    environmentWeb: "Word 网页版",
    environmentBrowser: "浏览器 / 静态页",
    environmentOther: "Office 宿主",
    environmentDesktopHint: "推荐直接读取当前文档，预览确认后可一键写回。",
    environmentWebHint: "网页版 Word 继续支持上传回退；若要直接读写当前文档，请改在桌面版打开。",
    environmentBrowserHint: "当前页面也能完成上传、扫描和导出，只是不会直接操作 Word 当前文档。",
    actionNoteIdle: "当前未连接 Word 当前文档，完成预览后可以先下载修正后的 DOCX。",
    actionNoteUpload: "上传模式下会保留下载导出；若要直接写回当前文档，请先从 Word 中读取文档。",
    actionNoteReady: "当前预览来自 Word 当前文档，确认无误后可以直接写回。",
    actionNoteNeedsPreview: "已经载入 Word 当前文档；请先扫描生成预览，再执行写回。",
    loadCurrentButton: "读取当前 Word 文档",
    loadingCurrentButton: "读取中…",
    applyCurrentButton: "写回当前文档",
    applyingCurrentButton: "写回中…",
    loadingCurrentStatus: "正在从当前 Word 文档读取完整 DOCX…",
    loadedCurrentStatus: "已读取当前 Word 文档，可以开始扫描。",
    loadUnsupportedStatus: "当前宿主不支持直接读取整个 DOCX。请改用上传模式，或在桌面版 Word 中打开此加载项。",
    loadFailedStatus: "读取当前 Word 文档失败。你可以稍后重试，或先使用上传模式处理。",
    previewRequiredStatus: "请先扫描并生成预览，再执行写回。",
    loadCurrentFirstStatus: "请先从当前 Word 文档载入内容，再执行写回。",
    applyingCurrentStatus: "正在把修正结果写回当前 Word 文档…",
    applySuccessStatus: "修正结果已写回当前 Word 文档。",
    applyUnsupportedStatus: "当前宿主不支持把修正后的文档直接写回。你仍然可以先下载修正后的 DOCX。",
    applyFailedStatus: "写回当前 Word 文档失败。你仍然可以先下载修正后的 DOCX。",
    currentDocumentName: "当前文档.docx"
  },
  en: {
    sourceModeEmpty: "No document",
    sourceModeUpload: "Uploaded DOCX",
    sourceModeWord: "Current Word document",
    sourceHintIdle: "Load the current Word document first, or upload a DOCX as a fallback.",
    sourceHintUpload: "Upload mode works well for Word on the web, browser preview, or offline review.",
    sourceHintWord: "The current preview is backed by the active Word document, so you can write the result back after checking it.",
    availabilityPending: "Detecting the Word host environment…",
    availabilityWordDesktop: "Connected to desktop Word. You can load the current document directly and write the corrected result back after previewing.",
    availabilityWordWeb: "Connected to Word on the web. Because the full DOCX cannot be read directly there, please use upload mode for scanning and export.",
    availabilityBrowser: "This page is not running inside Word, but you can still upload a DOCX, preview changes, and export the corrected file.",
    environmentDesktop: "Word desktop",
    environmentWeb: "Word on the web",
    environmentBrowser: "Browser / static page",
    environmentOther: "Office host",
    environmentDesktopHint: "Direct document read/write is available here after preview.",
    environmentWebHint: "Upload fallback stays available on Word on the web. Use desktop Word if you want direct current-document access.",
    environmentBrowserHint: "Upload, scan, and export still work here; only direct current-document access is unavailable.",
    actionNoteIdle: "A corrected DOCX can always be downloaded after preview, even when the current Word document is not connected.",
    actionNoteUpload: "Upload mode keeps export available. To write directly back into Word, load the current document first.",
    actionNoteReady: "This preview comes from the current Word document. After you confirm it, you can write it back directly.",
    actionNoteNeedsPreview: "The current Word document is loaded. Run a scan first to generate a preview before writing back.",
    loadCurrentButton: "Load Current Word Doc",
    loadingCurrentButton: "Loading…",
    applyCurrentButton: "Write Back to Word",
    applyingCurrentButton: "Writing Back…",
    loadingCurrentStatus: "Reading the full DOCX from the current Word document…",
    loadedCurrentStatus: "The current Word document is ready. You can start scanning now.",
    loadUnsupportedStatus: "This host cannot read the full DOCX directly. Please use upload mode, or open the add-in in desktop Word.",
    loadFailedStatus: "Failed to read the current Word document. You can retry, or continue in upload mode.",
    previewRequiredStatus: "Run a scan first so there is a preview to write back.",
    loadCurrentFirstStatus: "Load the current Word document first before writing back.",
    applyingCurrentStatus: "Writing the corrected result back into the current Word document…",
    applySuccessStatus: "The corrected result has been written back into the current Word document.",
    applyUnsupportedStatus: "This host cannot write the corrected document directly back. You can still download the corrected DOCX.",
    applyFailedStatus: "Writing back to the current Word document failed. You can still download the corrected DOCX.",
    currentDocumentName: "current-document.docx"
  }
};

const officeState = {
  ready: false,
  host: null,
  platform: null,
  directDocumentAccess: false,
  loadingCurrentDocument: false,
  applyingCurrentDocument: false,
};

initializeOfficeBridge();

function initializeOfficeBridge() {
  bindOfficeBridgeEvents();
  renderOfficeBridge();
  void detectOfficeHost();
}

function bindOfficeBridgeEvents() {
  officeElements.loadCurrentDocumentButton?.addEventListener("click", () => void loadCurrentWordDocument());
  officeElements.applyToCurrentDocumentButton?.addEventListener("click", () => void applyPreviewToCurrentWordDocument());

  officeElements.langZhButton?.addEventListener("click", () => requestAnimationFrame(renderOfficeBridge));
  officeElements.langEnButton?.addEventListener("click", () => requestAnimationFrame(renderOfficeBridge));

  window.addEventListener("wnh:file-selection", (event) => {
    const file = event.detail?.file ?? null;
    if (!file) {
      state.sourceKind = null;
      state.officeOriginalBase64 = null;
    } else if (file.__wnhSourceKind === "office") {
      state.sourceKind = "office";
    } else {
      state.sourceKind = "upload";
      state.officeOriginalBase64 = null;
    }

    renderOfficeBridge();
  });

  window.addEventListener("wnh:preview-rendered", renderOfficeBridge);
  window.addEventListener("wnh:preview-cleared", renderOfficeBridge);
}

async function detectOfficeHost() {
  if (!window.Office?.onReady) {
    officeState.ready = true;
    officeState.host = null;
    officeState.platform = null;
    officeState.directDocumentAccess = false;
    renderOfficeBridge();
    return;
  }

  try {
    const info = await Office.onReady();
    officeState.ready = true;
    officeState.host = info?.host ?? null;
    officeState.platform = info?.platform ?? null;
    officeState.directDocumentAccess = isWordHost(officeState.host) && !isWordWebPlatform(officeState.platform);
  } catch (error) {
    console.error(error);
    officeState.ready = true;
    officeState.host = null;
    officeState.platform = null;
    officeState.directDocumentAccess = false;
  }

  renderOfficeBridge();
}

function renderOfficeBridge() {
  const copy = getOfficeCopy();
  const sourceKind = state.sourceKind ?? null;
  const hasPreview = Boolean(state.preview?.blob);
  const isOfficeSource = sourceKind === "office";

  setTextContent(
    officeElements.sourceModeBadge,
    isOfficeSource ? copy.sourceModeWord : sourceKind === "upload" ? copy.sourceModeUpload : copy.sourceModeEmpty,
  );
  officeElements.sourceModeBadge?.classList.toggle("source-mode-badge--word", isOfficeSource);
  officeElements.sourceModeBadge?.classList.toggle("source-mode-badge--upload", sourceKind === "upload");

  setTextContent(
    officeElements.sourceHint,
    isOfficeSource ? copy.sourceHintWord : sourceKind === "upload" ? copy.sourceHintUpload : copy.sourceHintIdle,
  );

  setTextContent(officeElements.officeAvailability, getAvailabilityMessage(copy));
  setTextContent(officeElements.officeEnvironmentValue, getEnvironmentLabel(copy));
  setTextContent(officeElements.officeEnvironmentHint, getEnvironmentHint(copy));

  setTextContent(
    officeElements.officeActionNote,
    isOfficeSource
      ? hasPreview
        ? copy.actionNoteReady
        : copy.actionNoteNeedsPreview
      : sourceKind === "upload"
        ? copy.actionNoteUpload
        : copy.actionNoteIdle,
  );

  setTextContent(
    officeElements.loadCurrentDocumentButton,
    officeState.loadingCurrentDocument ? copy.loadingCurrentButton : copy.loadCurrentButton,
  );
  setTextContent(
    officeElements.applyToCurrentDocumentButton,
    officeState.applyingCurrentDocument ? copy.applyingCurrentButton : copy.applyCurrentButton,
  );

  if (officeElements.loadCurrentDocumentButton) {
    officeElements.loadCurrentDocumentButton.disabled =
      !officeState.directDocumentAccess || officeState.loadingCurrentDocument;
  }

  if (officeElements.applyToCurrentDocumentButton) {
    officeElements.applyToCurrentDocumentButton.disabled =
      !officeState.directDocumentAccess ||
      !isOfficeSource ||
      !hasPreview ||
      officeState.applyingCurrentDocument;
  }
}

async function loadCurrentWordDocument() {
  const copy = getOfficeCopy();

  if (!officeState.directDocumentAccess) {
    setRawStatus(copy.loadUnsupportedStatus, "warn");
    return;
  }

  officeState.loadingCurrentDocument = true;
  renderOfficeBridge();
  setRawStatus(copy.loadingCurrentStatus, "info");

  try {
    const base64 = await readCurrentDocumentAsBase64();
    const file = createDocxFileFromBase64(base64, copy.currentDocumentName, "office");
    state.officeOriginalBase64 = base64;
    handleFileSelection(file);
    setRawStatus(copy.loadedCurrentStatus, "info");
  } catch (error) {
    console.error(error);
    setRawStatus(resolveOfficeErrorMessage(error, copy.loadUnsupportedStatus, copy.loadFailedStatus), "warn");
  } finally {
    officeState.loadingCurrentDocument = false;
    renderOfficeBridge();
  }
}

async function applyPreviewToCurrentWordDocument() {
  const copy = getOfficeCopy();

  if (!state.preview?.blob) {
    setRawStatus(copy.previewRequiredStatus, "warn");
    return;
  }

  if (state.sourceKind !== "office") {
    setRawStatus(copy.loadCurrentFirstStatus, "warn");
    return;
  }

  if (!officeState.directDocumentAccess) {
    setRawStatus(copy.applyUnsupportedStatus, "warn");
    return;
  }

  officeState.applyingCurrentDocument = true;
  renderOfficeBridge();
  setRawStatus(copy.applyingCurrentStatus, "info");

  try {
    const base64 = await blobToBase64(state.preview.blob);
    await writeDocumentBackToWord(base64);
    state.officeOriginalBase64 = base64;
    state.file = createDocxFileFromArrayBuffer(
      await state.preview.blob.arrayBuffer(),
      state.file?.name ?? copy.currentDocumentName,
      "office",
    );
    renderFileMeta();
    setRawStatus(copy.applySuccessStatus, "info");
  } catch (error) {
    console.error(error);
    setRawStatus(resolveOfficeErrorMessage(error, copy.applyUnsupportedStatus, copy.applyFailedStatus), "warn");
  } finally {
    officeState.applyingCurrentDocument = false;
    renderOfficeBridge();
  }
}

async function readCurrentDocumentAsBase64() {
  const file = await getOfficeFileAsync(Office.FileType.Compressed, { sliceSize: 1024 * 1024 });

  try {
    const chunks = [];
    for (let index = 0; index < file.sliceCount; index += 1) {
      const slice = await getOfficeFileSliceAsync(file, index);
      chunks.push(toUint8Array(slice.data));
    }

    const merged = mergeUint8Arrays(chunks);
    return uint8ArrayToBase64(merged);
  } finally {
    await closeOfficeFileAsync(file);
  }
}

async function writeDocumentBackToWord(base64) {
  if (!window.Word?.run) {
    throw new Error("Word runtime is unavailable.");
  }

  await Word.run(async (context) => {
    if (typeof context.document.insertFileFromBase64 === "function") {
      context.document.insertFileFromBase64(base64, Word.InsertLocation.replace);
    } else {
      context.document.body.insertFileFromBase64(base64, Word.InsertLocation.replace);
    }

    await context.sync();
  });
}

function getOfficeFileAsync(fileType, options) {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(fileType, options ?? {}, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
        return;
      }

      reject(new Error(result.error?.message || "Office getFileAsync failed."));
    });
  });
}

function getOfficeFileSliceAsync(file, sliceIndex) {
  return new Promise((resolve, reject) => {
    file.getSliceAsync(sliceIndex, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
        return;
      }

      reject(new Error(result.error?.message || "Office getSliceAsync failed."));
    });
  });
}

function closeOfficeFileAsync(file) {
  return new Promise((resolve) => {
    file.closeAsync(() => resolve());
  });
}

function createDocxFileFromBase64(base64, name, sourceKind) {
  return createDocxFileFromArrayBuffer(base64ToArrayBuffer(base64), name, sourceKind);
}

function createDocxFileFromArrayBuffer(arrayBuffer, name, sourceKind) {
  const buffer =
    arrayBuffer instanceof ArrayBuffer
      ? arrayBuffer.slice(0)
      : arrayBuffer.buffer.slice(arrayBuffer.byteOffset, arrayBuffer.byteOffset + arrayBuffer.byteLength);

  if (typeof File === "function") {
    const file = new File([buffer], name, { type: DOCX_MIME_TYPE });
    file.__wnhSourceKind = sourceKind;
    return file;
  }

  return {
    __wnhSourceKind: sourceKind,
    name,
    size: buffer.byteLength,
    type: DOCX_MIME_TYPE,
    arrayBuffer: async () => buffer.slice(0),
  };
}

function toUint8Array(data) {
  if (data instanceof Uint8Array) {
    return data;
  }

  if (data instanceof ArrayBuffer) {
    return new Uint8Array(data);
  }

  if (Array.isArray(data)) {
    return Uint8Array.from(data);
  }

  return new Uint8Array();
}

function mergeUint8Arrays(chunks) {
  const totalLength = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const merged = new Uint8Array(totalLength);
  let offset = 0;

  for (const chunk of chunks) {
    merged.set(chunk, offset);
    offset += chunk.length;
  }

  return merged;
}

function uint8ArrayToBase64(bytes) {
  let binary = "";
  const chunkSize = 0x8000;

  for (let index = 0; index < bytes.length; index += chunkSize) {
    const chunk = bytes.subarray(index, index + chunkSize);
    binary += String.fromCharCode(...chunk);
  }

  return btoa(binary);
}

function base64ToArrayBuffer(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);

  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }

  return bytes.buffer;
}

function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = String(reader.result ?? "");
      resolve(result.includes(",") ? result.split(",")[1] : result);
    };
    reader.onerror = () => reject(reader.error ?? new Error("Failed to read blob."));
    reader.readAsDataURL(blob);
  });
}

function resolveOfficeErrorMessage(error, unsupportedMessage, fallbackMessage) {
  const raw = String(error?.message ?? "");
  if (!raw) {
    return fallbackMessage;
  }

  if (/not supported|unsupported|notimplemented|invalid api call|not available/i.test(raw)) {
    return unsupportedMessage;
  }

  return `${fallbackMessage} ${raw}`;
}

function getAvailabilityMessage(copy) {
  if (!officeState.ready) {
    return copy.availabilityPending;
  }

  if (isWordHost(officeState.host)) {
    return officeState.directDocumentAccess ? copy.availabilityWordDesktop : copy.availabilityWordWeb;
  }

  return copy.availabilityBrowser;
}

function getEnvironmentLabel(copy) {
  if (!officeState.ready || !officeState.host) {
    return copy.environmentBrowser;
  }

  if (isWordHost(officeState.host)) {
    return officeState.directDocumentAccess ? copy.environmentDesktop : copy.environmentWeb;
  }

  return `${copy.environmentOther}: ${officeState.host}`;
}

function getEnvironmentHint(copy) {
  if (!officeState.ready || !officeState.host) {
    return copy.environmentBrowserHint;
  }

  if (isWordHost(officeState.host)) {
    return officeState.directDocumentAccess ? copy.environmentDesktopHint : copy.environmentWebHint;
  }

  return copy.environmentBrowserHint;
}

function getOfficeCopy() {
  return OFFICE_COPY[state.language === "en" ? "en" : "zh"];
}

function isWordWebPlatform(platform) {
  return /online|web/i.test(String(platform ?? ""));
}

function isWordHost(host) {
  return /^word$/i.test(String(host ?? ""));
}

function setRawStatus(message, level = "info") {
  setStatus(message, level, { isRaw: true });
}

function setTextContent(element, value) {
  if (element) {
    element.textContent = value;
  }
}
