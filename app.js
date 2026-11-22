const PYODIDE_VERSION = "0.27.7";
const OLETOOLS_VERSION = "0.60.3";
const PCODE_CHUNK_SIZE = 16_384;
const PCODE_PREVIEW_LIMIT = 524_288;
const PCODE_DOWNLOAD_CHUNK_SIZE = 262_144;
const PCODE_LARGE_DOWNLOAD_WARNING_BYTES = 10 * 1024 * 1024;
const VBA_SOURCE_CHUNK_SIZE = 16_384;
const VBA_SOURCE_PREVIEW_LIMIT = 524_288;
const VBA_SOURCE_DOWNLOAD_CHUNK_SIZE = 262_144;
const VBA_SOURCE_LARGE_DOWNLOAD_WARNING_BYTES = 10 * 1024 * 1024;
const XLM_CHUNK_SIZE = 16_384;
const XLM_PREVIEW_LIMIT = 524_288;
const XLM_DOWNLOAD_CHUNK_SIZE = 262_144;
const XLM_LARGE_DOWNLOAD_WARNING_BYTES = 10 * 1024 * 1024;
const DEBUG_LOG_KEY = "vba-scan-debug-events";
const debugMode = new URLSearchParams(window.location.search).get("debug") === "1";

const fileInput = document.querySelector("#file-input");
const dropZone = document.querySelector("#drop-zone");
const browseButton = document.querySelector("#browse-button");
const selectedFileElement = document.querySelector("#selected-file");
const fileNameElement = document.querySelector("#file-name");
const fileSizeElement = document.querySelector("#file-size");
const removeFileButton = document.querySelector("#remove-file");
const runtimeNote = document.querySelector("#runtime-note");
const runtimeNoteContainer = runtimeNote.closest(".runtime-note");
const runtimeIndicator = document.querySelector(".runtime-dot");
const statusMessage = document.querySelector("#status-message");
const statusText = document.querySelector("#status-text");
const analysisSpinner = document.querySelector("#analysis-spinner");
const analysisProgress = document.querySelector("#analysis-progress");
const debugLogButton = document.querySelector("#download-debug-log");
const resultsElement = document.querySelector("#results");
const resultSummary = document.querySelector("#result-summary");
const extractedSourceSection = document.querySelector("#extracted-source-section");
const openVbaSourceButton = document.querySelector("#open-vba-source");
const openXlmSourceButton = document.querySelector("#open-xlm-source");
const openPcodeSourceButton = document.querySelector("#open-pcode-source");
const sourceCodeDialog = document.querySelector("#source-code-dialog");
const sourceDialogTitle = document.querySelector("#source-dialog-title");
const sourceDialogCloseButton = document.querySelector("#source-dialog-close");
const vbaSourceSection = document.querySelector("#vba-source-section");
const vbaSourcePreview = document.querySelector("#vba-source-preview");
const vbaSourceMatchLocation = document.querySelector("#vba-source-match-location");
const vbaSourceOutput = document.querySelector("#vba-source-output");
const vbaSourceControls = document.querySelector("#vba-source-controls");
const vbaSourceProgress = document.querySelector("#vba-source-progress");
const vbaSourceMoreButton = document.querySelector("#vba-source-more");
const vbaSourceLoadAboveButton = document.querySelector("#vba-source-load-above");
const vbaSourceLoadBelowButton = document.querySelector("#vba-source-load-below");
const vbaSourceDownloadButton = document.querySelector("#vba-source-download");
const vbaSourceTextDownloadButton = document.querySelector("#vba-source-text-download");
const vbaSourceLargeDownloadWarning = document.querySelector("#vba-source-large-download-warning");
const xlmFindingsGroup = document.querySelector("#xlm-findings-group");
const xlmFindingsList = document.querySelector("#xlm-findings-list");
const xlmFindingCount = document.querySelector("#xlm-finding-count");
const xlmSection = document.querySelector("#xlm-section");
const xlmOutput = document.querySelector("#xlm-output");
const xlmControls = document.querySelector("#xlm-controls");
const xlmProgress = document.querySelector("#xlm-progress");
const xlmMoreButton = document.querySelector("#xlm-more");
const xlmDownloadButton = document.querySelector("#xlm-download");
const xlmLargeDownloadWarning = document.querySelector("#xlm-large-download-warning");
const pcodeSection = document.querySelector("#pcode-section");
const pcodeOutput = document.querySelector("#pcode-output");
const pcodeControls = document.querySelector("#pcode-controls");
const pcodeProgress = document.querySelector("#pcode-progress");
const pcodeMoreButton = document.querySelector("#pcode-more");
const pcodeDownloadButton = document.querySelector("#pcode-download");
const pcodeLargeDownloadWarning = document.querySelector("#pcode-large-download-warning");
const vbaSourceFindingsList = document.querySelector("#vba-source-findings-list");
const vbaSourceFindingCount = document.querySelector("#vba-source-finding-count");
const vbaSourceFindingsGroup = document.querySelector('[aria-labelledby="vba-source-findings-heading"]');
const vbaSourceWarningIcon = document.querySelector("#vba-source-warning-icon");
const xlmWarningIcon = document.querySelector("#xlm-warning-icon");
const pcodeFindingsGroup = document.querySelector("#pcode-findings-group");
const pcodeWarningIcon = document.querySelector("#pcode-warning-icon");
const pcodeMatchNote = document.querySelector("#pcode-match-note");
const pcodeFindingsList = document.querySelector("#pcode-findings-list");
const pcodeFindingCount = document.querySelector("#pcode-finding-count");
const findingCount = document.querySelector("#finding-count");
const analysisTime = document.querySelector("#analysis-time");
const documentHash = document.querySelector("#document-hash");
const documentFilename = document.querySelector("#document-filename");
const documentCodepage = document.querySelector("#document-codepage");
const copyHashButton = document.querySelector("#copy-hash-button");
const copyStatus = document.querySelector("#copy-status");

let selectedFile = null;
const analysisWorker = new Worker(new URL("./analysis-worker.js", import.meta.url), { type: "module" });
const workerRequests = new Map();
let workerRequestSequence = 0;
let scanning = false;
let engineState = "loading";
let copyFeedbackTimeout = null;
let vbaSourceOffset = 0;
let vbaSourcePath = "";
let vbaSourceSize = 0;
let vbaSourceModuleCount = 0;
let vbaSourceArchivePath = "";
let vbaSourceArchiveSize = 0;
let vbaSourceLoading = false;
let vbaSourceDownloadInProgress = false;
let vbaSourceDecoder = new TextDecoder("utf-8");
let vbaSourceMatchContext = null;
let vbaSourceWindowReadOffset = 0;
let vbaSourceWindowPendingBytes = new Uint8Array();
let vbaSourceWindowLoading = false;
let xlmPath = "";
let xlmSize = 0;
let xlmMacroCount = 0;
let xlmOffset = 0;
let xlmLoading = false;
let xlmDownloadInProgress = false;
let xlmDecoder = new TextDecoder("utf-8");
let pcodeOffset = 0;
let pcodePath = "";
let pcodeSize = 0;
let pcodeLoading = false;
let pcodeDownloadInProgress = false;
let pcodeDecoder = new TextDecoder("utf-8");
const pcodeSourceContextCache = new Map();
const vbaSourceContextCache = new Map();
const debugEvents = [];
if (debugMode) {
  try {
    const savedEvents = JSON.parse(localStorage.getItem(DEBUG_LOG_KEY) || "[]");
    if (Array.isArray(savedEvents)) debugEvents.push(...savedEvents.slice(-500));
  } catch {}
}

debugLogButton.hidden = !debugMode;

function setEngineState(state) {
  engineState = state;
  runtimeIndicator.dataset.state = state;
}

const macroCapableOfficeExtensions = new Set([
  "xls", "xla", "xlsm", "xlam", "xlsb",
  "doc", "dot", "docm", "dotm",
  "ppt", "pot", "pps", "pptm", "potm", "ppsm",
]);
const macroFreeOfficeExtensions = new Set([
  "xlsx", "xltx", "docx", "dotx", "pptx", "potx", "ppsx",
]);

function logDebug(event, details = {}) {
  if (!debugMode) return;
  debugEvents.push({ timestamp: new Date().toISOString(), elapsed_ms: Math.round(performance.now()), event, ...details });
  if (debugEvents.length > 500) debugEvents.shift();
  try { localStorage.setItem(DEBUG_LOG_KEY, JSON.stringify(debugEvents)); } catch {}
}

function getErrorDetails(error) {
  return {
    name: error?.name || "Error",
    message: String(error?.message || error).replace(selectedFile?.name || "\0", "[selected file]"),
  };
}

window.addEventListener("error", (event) => {
  logDebug("window_error", { message: event.message, filename: event.filename, line: event.lineno });
});
window.addEventListener("unhandledrejection", (event) => {
  logDebug("unhandled_rejection", getErrorDetails(event.reason));
});
if (debugMode && "PerformanceObserver" in window) {
  try {
    new PerformanceObserver((list) => {
      for (const entry of list.getEntries()) {
        logDebug("long_task", { duration_ms: Math.round(entry.duration), start_ms: Math.round(entry.startTime) });
      }
    }).observe({ type: "longtask", buffered: true });
  } catch {}
}
logDebug("debug_session_started", { user_agent: navigator.userAgent });

function setStatus(message = "", isError = false) {
  statusText.textContent = message;
  const showProgress = scanning && !isError;
  analysisSpinner.hidden = !showProgress;
  analysisProgress.hidden = !showProgress;
  statusMessage.hidden = !message && !showProgress;
  statusMessage.classList.toggle("busy", showProgress);
  statusMessage.classList.toggle("error", isError);
}

function requestWorker(type, payload = {}, transfer = []) {
  const requestId = ++workerRequestSequence;
  return new Promise((resolve, reject) => {
    workerRequests.set(requestId, { resolve, reject });
    try {
      analysisWorker.postMessage({ type, requestId, ...payload }, transfer);
    } catch (error) {
      workerRequests.delete(requestId);
      reject(error);
    }
  });
}

analysisWorker.addEventListener("message", ({ data }) => {
  if (data.type === "engine-status") {
    setEngineState(data.state);
    if (scanning) setStatus(data.message, data.state === "error");
    else {
      runtimeNote.textContent = data.message;
      if (data.state === "error") setStatus(`Analysis engine could not be loaded: ${data.message}`, true);
    }
    return;
  }
  if (data.type === "scan-phase") {
    if (scanning) setStatus(data.message);
    return;
  }
  const pending = workerRequests.get(data.requestId);
  if (!pending) return;
  workerRequests.delete(data.requestId);
  if (data.type === "file-bytes") pending.resolve(new Uint8Array(data.bytes));
  else if (data.type === "scan-complete") pending.resolve(data);
  else if (data.type === "request-error") pending.reject(new Error(data.message));
});

analysisWorker.addEventListener("error", (event) => {
  const error = new Error(event.message || "The analysis worker stopped unexpectedly.");
  for (const { reject } of workerRequests.values()) reject(error);
  workerRequests.clear();
  setEngineState("error");
  if (scanning) {
    setBusy(false);
    setStatus(`Analysis could not be completed: ${error.message}`, true);
  } else {
    setStatus(`Analysis engine could not be loaded: ${error.message}`, true);
  }
});

function logDuration(label, startedAt) {
  const duration = Number((performance.now() - startedAt).toFixed(1));
  console.info(`[VBA Scan] ${label}: ${duration} ms`);
  logDebug("timing", { label, duration_ms: duration });
}

async function readPcodeBytes(offset, size) {
  return readTemporaryFileBytes(pcodePath, offset, size);
}

async function findPcodeSourceContext(keyword) {
  const cacheKey = keyword.trim().toLowerCase();
  if (pcodeSourceContextCache.has(cacheKey)) return pcodeSourceContextCache.get(cacheKey);

  const search = (async () => {
    const pattern = new RegExp(keyword.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "i");
    const decoder = new TextDecoder("utf-8");
    let offset = 0;
    let pendingLine = "";
    let inModuleStreams = false;
    let moduleStream = null;
    let block = null;
    const processLine = (line) => {
      const trimmedLine = line.trim();
      if (trimmedLine === "Module streams:") {
        inModuleStreams = true;
        return null;
      }
      const moduleStreamMatch = inModuleStreams && trimmedLine.match(/^(.+ - \d+ bytes)$/);
      if (moduleStreamMatch) {
        if (block?.matched) return block.context;
        moduleStream = moduleStreamMatch[1];
        block = null;
        return null;
      }
      const lineMatch = trimmedLine.match(/^Line #(\d+):$/);
      if (lineMatch) {
        if (block?.matched) return block.context;
        block = {
          matched: false,
          context: { number: Number(lineMatch[1]), text: line, module_stream: moduleStream },
        };
        return null;
      }
      if (block) {
        block.context.text += `\n${line}`;
        if (pattern.test(line)) block.matched = true;
      }
      return null;
    };

    while (offset < pcodeSize) {
      const bytes = await readPcodeBytes(offset, Math.min(PCODE_CHUNK_SIZE, pcodeSize - offset));
      if (bytes.length === 0) break;
      offset += bytes.length;
      pendingLine += decoder.decode(bytes, { stream: offset < pcodeSize });
      const lines = pendingLine.split("\n");
      pendingLine = lines.pop();
      for (const line of lines) {
        const context = processLine(line);
        if (context) return context;
      }
    }
    if (pendingLine) {
      const context = processLine(pendingLine);
      if (context) return context;
    }
    return block?.matched ? block.context : null;
  })();

  pcodeSourceContextCache.set(cacheKey, search);
  try {
    return await search;
  } catch (error) {
    pcodeSourceContextCache.delete(cacheKey);
    throw error;
  }
}

async function findVbaSourceContext(keyword) {
  const searchTerm = keyword.trim().toLowerCase();
  if (!searchTerm || !vbaSourceSize) return null;
  if (vbaSourceContextCache.has(searchTerm)) return vbaSourceContextCache.get(searchTerm);

  const search = (async () => {
    const decoder = new TextDecoder("utf-8");
    const encoder = new TextEncoder();
    let offset = 0;
    let scannedOffset = 0;
    let pendingLine = "";
    let moduleName = "";
    let moduleLineNumber = 0;
    let precedingLines = [];
    let matchedLine = null;
    let contextLines = [];
    let followingLinesRemaining = 2;
    const getContext = () => ({
      moduleName,
      lineNumber: matchedLine.number,
      lines: contextLines,
      startByteOffset: contextLines[0].byteOffset,
      endByteOffset: contextLines.at(-1).endByteOffset,
    });
    const processLine = (line, newlineBytes) => {
      const byteOffset = scannedOffset;
      scannedOffset += encoder.encode(line).length + newlineBytes;
      const sourceText = line.replace(/\r$/, "");
      const moduleHeader = sourceText.match(/^' ===== (.+) =====$/);
      if (moduleHeader) {
        if (matchedLine) return getContext();
        moduleName = moduleHeader[1];
        moduleLineNumber = 0;
        precedingLines = [];
        return null;
      }
      if (!moduleName) return null;

      const sourceLine = {
        number: ++moduleLineNumber,
        text: sourceText,
        matched: false,
        byteOffset,
        endByteOffset: scannedOffset,
      };
      if (matchedLine) {
        contextLines.push(sourceLine);
        followingLinesRemaining -= 1;
        return followingLinesRemaining === 0 ? getContext() : null;
      }
      if (line.toLowerCase().includes(searchTerm)) {
        sourceLine.matched = true;
        matchedLine = sourceLine;
        contextLines = [...precedingLines, sourceLine];
        return null;
      }
      precedingLines.push(sourceLine);
      if (precedingLines.length > 2) precedingLines.shift();
      return null;
    };

    while (offset < vbaSourceSize) {
      const bytes = await readTemporaryFileBytes(
        vbaSourcePath,
        offset,
        Math.min(VBA_SOURCE_CHUNK_SIZE, vbaSourceSize - offset),
      );
      if (bytes.length === 0) break;
      offset += bytes.length;
      pendingLine += decoder.decode(bytes, { stream: offset < vbaSourceSize });
      const lines = pendingLine.split("\n");
      pendingLine = lines.pop();
      for (const line of lines) {
        const context = processLine(line, 1);
        if (context) return context;
      }
    }
    if (pendingLine) {
      const context = processLine(pendingLine, 0);
      if (context) return context;
    }
    return matchedLine ? getContext() : null;
  })();

  vbaSourceContextCache.set(searchTerm, search);
  try {
    return await search;
  } catch (error) {
    vbaSourceContextCache.delete(searchTerm);
    throw error;
  }
}

function renderVbaSourceContext(container, context) {
  const module = document.createElement("div");
  module.className = "vba-source-context-module";
  module.textContent = `${context.moduleName} · line ${context.lineNumber}`;
  const code = document.createElement("pre");
  for (const line of context.lines) {
    const sourceLine = document.createElement("span");
    sourceLine.className = line.matched ? "vba-source-context-line is-match" : "vba-source-context-line";
    sourceLine.textContent = line.number == null
      ? line.text
      : `${String(line.number).padStart(4, " ")}  ${line.text}`;
    code.append(sourceLine, document.createTextNode("\n"));
  }
  container.replaceChildren(module, code);
}

function renderVbaSourceWindow(context) {
  const nodes = [];
  for (const line of context.lines) {
    const sourceLine = document.createElement("span");
    sourceLine.className = line.matched ? "vba-source-context-line is-match" : "vba-source-context-line";
    sourceLine.textContent = line.text;
    nodes.push(sourceLine, document.createTextNode("\n"));
  }
  vbaSourceOutput.replaceChildren(...nodes);
}

async function loadMoreVbaSourceAbove() {
  const context = vbaSourceMatchContext;
  if (!context || vbaSourceWindowLoading || context.startByteOffset <= 0) return;
  vbaSourceWindowLoading = true;
  vbaSourceLoadAboveButton.disabled = true;
  vbaSourceLoadBelowButton.disabled = true;
  const pre = vbaSourceOutput.parentElement;
  const previousHeight = pre.scrollHeight;
  try {
    const endOffset = context.startByteOffset;
    const startOffset = Math.max(0, endOffset - VBA_SOURCE_CHUNK_SIZE - 4);
    let bytes = await readTemporaryFileBytes(vbaSourcePath, startOffset, endOffset - startOffset);
    let alignedOffset = startOffset;
    if (startOffset > 0) {
      const newlineIndex = bytes.indexOf(10);
      if (newlineIndex >= 0) {
        alignedOffset += newlineIndex + 1;
        bytes = bytes.subarray(newlineIndex + 1);
      }
    }
    const text = new TextDecoder("utf-8").decode(bytes);
    const lines = text.split("\n");
    if (lines.at(-1) === "") lines.pop();
    const encoder = new TextEncoder();
    let lineOffset = alignedOffset;
    const addedLines = lines.map((line, index) => {
      const hasNewline = index < lines.length - 1 || text.endsWith("\n");
      const endByteOffset = lineOffset + encoder.encode(line).length + Number(hasNewline);
      const sourceText = line.replace(/\r$/, "");
      const sourceLine = {
        number: null,
        text: sourceText,
        matched: false,
        byteOffset: lineOffset,
        endByteOffset,
      };
      lineOffset = endByteOffset;
      return sourceLine;
    });
    if (addedLines.length) {
      context.lines.unshift(...addedLines);
      context.startByteOffset = addedLines[0].byteOffset;
      renderVbaSourceWindow(context);
      pre.scrollTop += pre.scrollHeight - previousHeight;
    }
    vbaSourceLoadAboveButton.hidden = context.startByteOffset === 0;
  } catch (error) {
    logDebug("vba_source_context_above_error", getErrorDetails(error));
    vbaSourceProgress.textContent = "Could not load earlier VBA source lines.";
  } finally {
    vbaSourceLoadAboveButton.disabled = false;
    vbaSourceLoadBelowButton.disabled = false;
    vbaSourceWindowLoading = false;
  }
}

async function loadMoreVbaSourceBelow() {
  const context = vbaSourceMatchContext;
  if (!context || vbaSourceWindowLoading || vbaSourceWindowReadOffset >= vbaSourceSize) return;
  vbaSourceWindowLoading = true;
  vbaSourceLoadBelowButton.disabled = true;
  vbaSourceLoadAboveButton.disabled = true;
  try {
    const bytes = await readTemporaryFileBytes(
      vbaSourcePath,
      vbaSourceWindowReadOffset,
      Math.min(VBA_SOURCE_CHUNK_SIZE, vbaSourceSize - vbaSourceWindowReadOffset),
    );
    vbaSourceWindowReadOffset += bytes.length;
    const combined = new Uint8Array(vbaSourceWindowPendingBytes.length + bytes.length);
    combined.set(vbaSourceWindowPendingBytes);
    combined.set(bytes, vbaSourceWindowPendingBytes.length);
    const atEnd = vbaSourceWindowReadOffset >= vbaSourceSize;
    const lastNewline = combined.lastIndexOf(10);
    const completeLength = atEnd ? combined.length : lastNewline + 1;
    const completeBytes = combined.subarray(0, Math.max(0, completeLength));
    vbaSourceWindowPendingBytes = combined.slice(Math.max(0, completeLength));
    if (completeBytes.length) {
      const text = new TextDecoder("utf-8").decode(completeBytes);
      const lines = text.split("\n");
      const endsWithNewline = completeBytes.at(-1) === 10;
      if (endsWithNewline) lines.pop();
      const encoder = new TextEncoder();
      let lineOffset = context.endByteOffset;
      const addedLines = lines.map((line, index) => {
        const hasNewline = index < lines.length - 1 || endsWithNewline;
        const endByteOffset = lineOffset + encoder.encode(line).length + Number(hasNewline);
        const sourceText = line.replace(/\r$/, "");
        const moduleHeader = sourceText.match(/^' ===== (.+) =====$/);
        const sourceLine = {
          number: moduleHeader ? null : undefined,
          text: sourceText,
          matched: false,
          byteOffset: lineOffset,
          endByteOffset,
        };
        lineOffset = endByteOffset;
        return sourceLine;
      });
      context.lines.push(...addedLines);
      context.endByteOffset = lineOffset;
      renderVbaSourceWindow(context);
    }
    vbaSourceLoadBelowButton.hidden = atEnd && vbaSourceWindowPendingBytes.length === 0;
  } catch (error) {
    logDebug("vba_source_context_below_error", getErrorDetails(error));
    vbaSourceProgress.textContent = "Could not load later VBA source lines.";
  } finally {
    vbaSourceLoadBelowButton.disabled = false;
    vbaSourceLoadAboveButton.disabled = false;
    vbaSourceWindowLoading = false;
  }
}

function openVbaSourceMatch(context) {
  for (const section of [vbaSourceSection, xlmSection, pcodeSection]) {
    section.hidden = section !== vbaSourceSection;
  }
  vbaSourceMatchContext = context;
  vbaSourceWindowReadOffset = context.endByteOffset;
  vbaSourceWindowPendingBytes = new Uint8Array();
  vbaSourceOffset = 0;
  vbaSourceOutput.replaceChildren();
  vbaSourceMatchLocation.textContent = `${context.moduleName} · line ${context.lineNumber}`;
  vbaSourceMatchLocation.hidden = false;
  vbaSourceLoadAboveButton.hidden = context.startByteOffset === 0;
  vbaSourceLoadBelowButton.hidden = vbaSourceWindowReadOffset >= vbaSourceSize;
  vbaSourceMoreButton.hidden = true;
  vbaSourceControls.hidden = false;
  vbaSourceOutput.classList.add("vba-match-window");
  vbaSourceProgress.textContent = "Showing source around the matched line";
  renderVbaSourceWindow(context);
  vbaSourcePreview.hidden = false;
  sourceDialogTitle.textContent = "VBA source code";
  sourceCodeDialog.showModal();
}

function createVbaSourceDetails(keyword) {
  const details = document.createElement("details");
  details.className = "vba-source-lines";
  const summary = document.createElement("summary");
  summary.textContent = "Show matching VBA source line";
  const content = document.createElement("div");
  content.className = "vba-source-context";
  details.append(summary, content);
  let loading = false;
  let loaded = false;
  details.addEventListener("toggle", () => {
    if (!details.open || loading || loaded) return;
    loading = true;
    content.textContent = "Searching extracted VBA source…";
    void findVbaSourceContext(keyword).then((context) => {
      content.replaceChildren();
      if (!context) {
        summary.textContent = "VBA source line (no direct match)";
        content.textContent = "No direct match was found in the extracted VBA source.";
        loaded = true;
        return;
      }
      summary.textContent = `Matched ${context.moduleName}, line ${context.lineNumber}`;
      renderVbaSourceContext(content, context);
      const seeMore = document.createElement("button");
      seeMore.type = "button";
      seeMore.className = "source-see-more";
      seeMore.textContent = "See more";
      seeMore.addEventListener("click", () => openVbaSourceMatch(context));
      content.append(seeMore);
      loaded = true;
    }).catch((error) => {
      logDebug("vba_source_context_error", getErrorDetails(error));
      content.textContent = "Could not search VBA source. Close and reopen to retry.";
    }).finally(() => {
      loading = false;
    });
  });
  return details;
}

async function readTemporaryFileBytes(path, offset, size) {
  return requestWorker("read-file", { path, offset, size });
}

async function appendVbaSourceChunk() {
  if (vbaSourceLoading || vbaSourceOffset >= Math.min(vbaSourceSize, VBA_SOURCE_PREVIEW_LIMIT)) return;
  vbaSourceLoading = true;
  vbaSourceMoreButton.disabled = true;
  const startedAt = performance.now();
  try {
    const bytes = await readTemporaryFileBytes(
      vbaSourcePath,
      vbaSourceOffset,
      VBA_SOURCE_CHUNK_SIZE,
    );
    vbaSourceOffset += bytes.length;
    vbaSourceOutput.append(
      document.createTextNode(vbaSourceDecoder.decode(bytes, { stream: vbaSourceOffset < vbaSourceSize })),
    );
    vbaSourceProgress.textContent =
      `Preview: ${vbaSourceOffset.toLocaleString()} of ${vbaSourceSize.toLocaleString()} bytes · ${vbaSourceModuleCount.toLocaleString()} modules`;
    vbaSourceMoreButton.hidden = vbaSourceOffset >= Math.min(vbaSourceSize, VBA_SOURCE_PREVIEW_LIMIT);
    logDebug("vba_source_preview_chunk", { bytes: bytes.length, total_bytes: vbaSourceSize });
  } catch (error) {
    logDebug("vba_source_preview_error", getErrorDetails(error));
    setStatus(`Could not load VBA source preview: ${error.message}`, true);
  } finally {
    vbaSourceMoreButton.disabled = false;
    logDuration("VBA source preview chunk", startedAt);
    vbaSourceLoading = false;
  }
}

async function loadVbaSourcePreview() {
  while (vbaSourceOffset < Math.min(vbaSourceSize, VBA_SOURCE_PREVIEW_LIMIT)) {
    const previousOffset = vbaSourceOffset;
    await appendVbaSourceChunk();
    if (vbaSourceOffset === previousOffset) break;
  }
}

async function appendPcodeChunk() {
  if (pcodeLoading || pcodeOffset >= Math.min(pcodeSize, PCODE_PREVIEW_LIMIT)) return;
  pcodeLoading = true;
  pcodeMoreButton.disabled = true;
  const startedAt = performance.now();
  try {
    const bytes = await readPcodeBytes(pcodeOffset, PCODE_CHUNK_SIZE);
    pcodeOffset += bytes.length;
    pcodeOutput.append(document.createTextNode(pcodeDecoder.decode(bytes, { stream: pcodeOffset < pcodeSize })));
    pcodeProgress.textContent = `Preview: ${pcodeOffset.toLocaleString()} of ${pcodeSize.toLocaleString()} bytes`;
    pcodeMoreButton.hidden = pcodeOffset >= Math.min(pcodeSize, PCODE_PREVIEW_LIMIT);
    logDebug("pcode_preview_chunk", { bytes: bytes.length, total_bytes: pcodeSize });
  } catch (error) {
    logDebug("pcode_preview_error", getErrorDetails(error));
    setStatus(`Could not load p-code preview: ${error.message}`, true);
  } finally {
    pcodeMoreButton.disabled = false;
    logDuration("P-code preview chunk", startedAt);
    pcodeLoading = false;
  }
}

async function appendXlmChunk() {
  if (xlmLoading || xlmOffset >= Math.min(xlmSize, XLM_PREVIEW_LIMIT)) return;
  xlmLoading = true;
  xlmMoreButton.disabled = true;
  try {
    const bytes = await readTemporaryFileBytes(xlmPath, xlmOffset, XLM_CHUNK_SIZE);
    xlmOffset += bytes.length;
    xlmOutput.append(document.createTextNode(xlmDecoder.decode(bytes, { stream: xlmOffset < xlmSize })));
    xlmProgress.textContent = `Preview: ${xlmOffset.toLocaleString()} of ${xlmSize.toLocaleString()} bytes`;
    xlmMoreButton.hidden = xlmOffset >= Math.min(xlmSize, XLM_PREVIEW_LIMIT);
  } catch (error) {
    logDebug("xlm_preview_error", getErrorDetails(error));
    setStatus(`Could not load XLM macro preview: ${error.message}`, true);
  } finally {
    xlmMoreButton.disabled = false;
    xlmLoading = false;
  }
}

function openSourcePreview(sourceType) {
  const sources = {
    vba: { section: vbaSourceSection, title: "VBA source code", size: vbaSourceSize, offset: vbaSourceOffset },
    xlm: { section: xlmSection, title: "XLM formula source", size: xlmSize, offset: xlmOffset },
    pcode: { section: pcodeSection, title: "P-code source", size: pcodeSize, offset: pcodeOffset },
  };
  const source = sources[sourceType];
  if (!source) return;
  for (const section of Object.values(sources).map(({ section }) => section)) {
    section.hidden = section !== source.section;
  }
  if (sourceType === "vba") {
    vbaSourceMatchContext = null;
    vbaSourceOffset = 0;
    vbaSourceDecoder = new TextDecoder("utf-8");
    vbaSourceOutput.replaceChildren();
    vbaSourceOutput.classList.remove("vba-match-window");
    vbaSourceMatchLocation.hidden = true;
    vbaSourceLoadAboveButton.hidden = true;
    vbaSourceLoadBelowButton.hidden = true;
    vbaSourceMoreButton.hidden = !source.size;
    vbaSourceMoreButton.textContent = "Show more";
    vbaSourceControls.hidden = !source.size;
    vbaSourcePreview.hidden = false;
  }
  sourceDialogTitle.textContent = source.title;
  sourceCodeDialog.showModal();
  if (!source.size) return;
  if (sourceType === "vba") {
    if (vbaSourceSize <= VBA_SOURCE_PREVIEW_LIMIT) void loadVbaSourcePreview();
    else void appendVbaSourceChunk();
    return;
  }
  if (source.offset > 0) return;
  if (sourceType === "xlm") {
    void appendXlmChunk();
  } else {
    void appendPcodeChunk();
  }
}

function downloadDebugLog() {
  const log = {
    app: "VBA Scan",
    pyodide_version: PYODIDE_VERSION,
    oletools_version: OLETOOLS_VERSION,
    user_agent: navigator.userAgent,
    events: debugEvents,
  };
  const url = URL.createObjectURL(new Blob([JSON.stringify(log, null, 2)], { type: "application/json" }));
  const link = document.createElement("a");
  link.href = url;
  link.download = "vba-scan-debug-log.json";
  link.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function setSelectedFile(file) {
  const extension = file.name.split(".").pop().toLowerCase();
  if (macroFreeOfficeExtensions.has(extension)) {
    setStatus(`.${extension} files cannot contain embedded VBA macros. Choose a macro-enabled format such as .xlsm, .docm, or .pptm.`, true);
    return;
  }
  if (!macroCapableOfficeExtensions.has(extension)) {
    setStatus("Choose a supported Office document (for example .xlsm, .docm, or .pptm).", true);
    return;
  }
  if (file.size === 0) {
    setStatus("This file is empty. Choose a document with content to analyze.", true);
    return;
  }

  selectedFile = file;
  pcodeSourceContextCache.clear();
  vbaSourceContextCache.clear();
  logDebug("file_selected", { size_bytes: file.size, extension });
  fileNameElement.textContent = file.name;
  fileSizeElement.textContent = formatFileSize(file.size);
  selectedFileElement.hidden = false;
  dropZone.hidden = true;
  resultsElement.hidden = true;
  setStatus();
  void analyzeSelectedFile();
}

function clearSelectedFile() {
  selectedFile = null;
  pcodeSourceContextCache.clear();
  vbaSourceContextCache.clear();
  fileInput.value = "";
  const temporaryPaths = [pcodePath, vbaSourcePath, vbaSourceArchivePath, xlmPath].filter(Boolean);
  if (temporaryPaths.length) analysisWorker.postMessage({ type: "clear-files", paths: temporaryPaths });
  pcodePath = "";
  vbaSourcePath = "";
  vbaSourceArchivePath = "";
  xlmPath = "";
  pcodeSize = 0;
  vbaSourceSize = 0;
  vbaSourceArchiveSize = 0;
  vbaSourceModuleCount = 0;
  xlmSize = 0;
  xlmMacroCount = 0;
  selectedFileElement.hidden = true;
  dropZone.hidden = false;
  resultsElement.hidden = true;
  setStatus();
}

function formatFileSize(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function getPcodeFilename(fileName) {
  const extensionSeparator = fileName.lastIndexOf(".");
  if (extensionSeparator <= 0 || extensionSeparator === fileName.length - 1) {
    return `${fileName}-pcode.txt`;
  }
  return `${fileName.slice(0, extensionSeparator)}-${fileName.slice(extensionSeparator + 1)}-pcode.txt`;
}

function getVbaSourceArchiveFilename(fileName) {
  const extensionSeparator = fileName.lastIndexOf(".");
  const stem = extensionSeparator > 0 ? fileName.slice(0, extensionSeparator) : fileName;
  return `${stem}-vba-source.zip`;
}

function getVbaSourceTextFilename(fileName) {
  const extensionSeparator = fileName.lastIndexOf(".");
  const stem = extensionSeparator > 0 ? fileName.slice(0, extensionSeparator) : fileName;
  return `${stem}-vba-source.txt`;
}

function getXlmFilename(fileName) {
  const extensionSeparator = fileName.lastIndexOf(".");
  const stem = extensionSeparator > 0 ? fileName.slice(0, extensionSeparator) : fileName;
  return `${stem}-xlm-macros.txt`;
}

async function downloadVbaSource(path, size, filename, mimeType, label, eventPrefix) {
  if (vbaSourceDownloadInProgress || !size || !selectedFile) return;
  vbaSourceDownloadInProgress = true;
  vbaSourceDownloadButton.disabled = true;
  vbaSourceTextDownloadButton.disabled = true;
  removeFileButton.disabled = true;
  const startedAt = performance.now();
  let written = 0;
  logDebug(`${eventPrefix}_started`, { bytes: size, download_method: "blob" });
  try {
    vbaSourceProgress.textContent = `Preparing ${label}…`;
    const chunks = [];
    for (let offset = 0; offset < size; offset += VBA_SOURCE_DOWNLOAD_CHUNK_SIZE) {
      const requestedBytes = Math.min(VBA_SOURCE_DOWNLOAD_CHUNK_SIZE, size - offset);
      const bytes = await readTemporaryFileBytes(path, offset, requestedBytes);
      if (bytes.length === 0) throw new Error(`The ${label} ended before the expected size.`);
      chunks.push(bytes);
      written += bytes.length;
      vbaSourceProgress.textContent =
        `Downloading ${written.toLocaleString()} of ${size.toLocaleString()} bytes`;
      if (written % (VBA_SOURCE_DOWNLOAD_CHUNK_SIZE * 16) === 0) {
        logDebug(`${eventPrefix}_progress`, { bytes: written, total_bytes: size });
      }
    }
    const blob = new Blob(chunks, { type: mimeType });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    link.click();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
    vbaSourceProgress.textContent = `Download complete (${written.toLocaleString()} bytes)`;
    logDebug(`${eventPrefix}_complete`, { bytes: written });
  } catch (error) {
    logDebug(`${eventPrefix}_error`, getErrorDetails(error));
    vbaSourceProgress.textContent = "Download did not complete";
    setStatus(`Could not download VBA source: ${error.message}`, true);
  } finally {
    vbaSourceDownloadButton.disabled = false;
    vbaSourceTextDownloadButton.disabled = false;
    removeFileButton.disabled = false;
    vbaSourceDownloadInProgress = false;
    logDuration(`${label} download`, startedAt);
  }
}

function downloadVbaSourceArchive() {
  if (!selectedFile) return;
  return downloadVbaSource(
    vbaSourceArchivePath,
    vbaSourceArchiveSize,
    getVbaSourceArchiveFilename(selectedFile.name),
    "application/zip",
    "VBA source archive",
    "vba_source_download",
  );
}

function downloadVbaSourceText() {
  if (!selectedFile) return;
  return downloadVbaSource(
    vbaSourcePath,
    vbaSourceSize,
    getVbaSourceTextFilename(selectedFile.name),
    "text/plain;charset=utf-8",
    "VBA source text",
    "vba_source_text_download",
  );
}

async function downloadXlmMacros() {
  if (xlmDownloadInProgress || !xlmSize || !selectedFile) return;
  xlmDownloadInProgress = true;
  xlmDownloadButton.disabled = true;
  removeFileButton.disabled = true;
  let written = 0;
  try {
    const chunks = [];
    for (let offset = 0; offset < xlmSize; offset += XLM_DOWNLOAD_CHUNK_SIZE) {
      const requestedBytes = Math.min(XLM_DOWNLOAD_CHUNK_SIZE, xlmSize - offset);
      const bytes = await readTemporaryFileBytes(xlmPath, offset, requestedBytes);
      if (bytes.length === 0) throw new Error("The XLM macro file ended before the expected size.");
      chunks.push(bytes);
      written += bytes.length;
      xlmProgress.textContent = `Downloading ${written.toLocaleString()} of ${xlmSize.toLocaleString()} bytes`;
    }
    const blob = new Blob(chunks, { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = getXlmFilename(selectedFile.name);
    link.click();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
    xlmProgress.textContent = `Download complete (${written.toLocaleString()} bytes)`;
  } catch (error) {
    logDebug("xlm_download_error", getErrorDetails(error));
    xlmProgress.textContent = "Download did not complete";
    setStatus(`Could not download XLM macros: ${error.message}`, true);
  } finally {
    xlmDownloadButton.disabled = false;
    removeFileButton.disabled = false;
    xlmDownloadInProgress = false;
  }
}

function setBusy(isBusy, message = "") {
  scanning = isBusy;
  if (!isBusy) {
    analysisSpinner.hidden = true;
    analysisProgress.hidden = true;
    statusMessage.classList.remove("busy");
  }
  removeFileButton.disabled = isBusy;
  runtimeNote.textContent = message || {
    loading: "Preparing analysis engine in background…",
    ready: "Analysis engine ready",
    error: "Engine unavailable; remove the file and upload again to retry",
  }[engineState];
}

function renderResults(result, elapsedMilliseconds, sha256, fileName) {
  documentHash.textContent = sha256;
  documentFilename.textContent = fileName;
  const projectCodepages = result.project_codepages || [];
  const codepages = [...new Map(projectCodepages.map(({ value, name }) => [value, { value, name }])).values()];
  documentCodepage.replaceChildren();
  if (codepages.length === 0) {
    documentCodepage.append(document.createTextNode("Not available"));
  } else {
    codepages.forEach(({ value, name }, index) => {
      if (index > 0) documentCodepage.append(document.createTextNode(", "));
      documentCodepage.append(document.createTextNode(String(value)));
      const info = document.createElement("span");
      info.className = "codepage-info";
      const icon = document.createElementNS("http://www.w3.org/2000/svg", "svg");
      icon.setAttribute("viewBox", "0 0 16 16");
      icon.setAttribute("aria-hidden", "true");
      const outline = document.createElementNS("http://www.w3.org/2000/svg", "circle");
      outline.setAttribute("cx", "8");
      outline.setAttribute("cy", "8");
      outline.setAttribute("r", "6.5");
      outline.setAttribute("fill", "none");
      outline.setAttribute("stroke", "currentColor");
      outline.setAttribute("stroke-width", "1.4");
      const stem = document.createElementNS("http://www.w3.org/2000/svg", "path");
      stem.setAttribute("d", "M8 7.2v4.1");
      stem.setAttribute("stroke", "currentColor");
      stem.setAttribute("stroke-width", "1.4");
      stem.setAttribute("stroke-linecap", "round");
      const dot = document.createElementNS("http://www.w3.org/2000/svg", "circle");
      dot.setAttribute("cx", "8");
      dot.setAttribute("cy", "4.8");
      dot.setAttribute("r", "0.8");
      dot.setAttribute("fill", "currentColor");
      icon.append(outline, stem, dot);
      info.append(icon);
      info.title = name || "Code page name unavailable";
      info.setAttribute("role", "img");
      info.setAttribute("tabindex", "0");
      info.setAttribute("aria-label", `Code page name: ${info.title}`);
      documentCodepage.append(document.createTextNode(" "), info);
    });
  }
  const summaries = [];
  if (result.has_xlm) {
    summaries.push({ label: "⚠️ Excel 4.0 macros detected", value: "Review this document", tone: "danger" });
  }
  if (result.stomping) {
    summaries.push({
      label: "⚠️ VBA stomping detected",
      value: "Pre-Compiled VBA differs from source",
      tone: "danger",
      note: {
        text: "ℹ️ Note that oletools can report false positives for VBA stomping. See upstream reports:",
        issues: [521, 533, 620],
      },
      mismatches: result.stomping_items || [],
    });
  }
  resultSummary.replaceChildren(...summaries.map((item) => {
    const card = document.createElement("div");
    card.className = "summary-card";
    const label = document.createElement("div");
    label.className = "summary-label";
    label.textContent = item.label;
    const value = document.createElement("div");
    value.className = `summary-value ${item.tone}`;
    const indicator = document.createElement("span");
    indicator.className = "indicator";
    value.append(indicator, document.createTextNode(item.value));
    card.append(label, value);
    if (item.note) {
      const note = document.createElement("p");
      note.className = "summary-note";
      note.append(document.createTextNode(`${item.note.text} `));
      item.note.issues.forEach((number, index) => {
        if (index > 0) note.append(document.createTextNode(", "));
        const link = document.createElement("a");
        link.href = `https://github.com/decalage2/oletools/issues/${number}`;
        link.target = "_blank";
        link.rel = "noopener noreferrer";
        link.textContent = `#${number}`;
        note.append(link);
      });
      card.append(note);
    }
    if (item.mismatches) {
      if (item.mismatches.length === 0) {
        const details = document.createElement("p");
        details.className = "summary-note stomping-mismatches";
        details.textContent = "No exact P-code/source mismatch item was available.";
        card.append(details);
      } else {
        const details = document.createElement("details");
        details.className = "summary-note stomping-mismatches";
        const summary = document.createElement("summary");
        summary.textContent = `P-code items not found in source (${item.mismatches.length})`;
        details.append(summary);
        const shownItems = item.mismatches.slice(0, 8);
        const list = document.createElement("ul");
        shownItems.forEach((mismatch) => {
          const entry = document.createElement("li");
          const code = document.createElement("code");
          code.textContent = `${mismatch.kind}: ${mismatch.value}`;
          entry.append(code);
          list.append(entry);
        });
        if (item.mismatches.length > shownItems.length) {
          const more = document.createElement("li");
          more.textContent = `And ${item.mismatches.length - shownItems.length} more`;
          list.append(more);
        }
        details.append(list);
        card.append(details);
      }
    }
    return card;
  }));
  resultSummary.hidden = summaries.length === 0;

  vbaSourceOffset = 0;
  vbaSourceSize = result.vba_source_size || 0;
  vbaSourceModuleCount = result.vba_source_module_count || 0;
  vbaSourceArchiveSize = result.vba_source_archive_size || 0;
  vbaSourceLargeDownloadWarning.hidden =
    Math.max(vbaSourceArchiveSize, vbaSourceSize) <= VBA_SOURCE_LARGE_DOWNLOAD_WARNING_BYTES;
  vbaSourceDecoder = new TextDecoder("utf-8");
  vbaSourceOutput.replaceChildren();
  vbaSourceOutput.classList.toggle("code-unavailable", !vbaSourceSize);
  if (vbaSourceSize) {
    vbaSourceControls.hidden = false;
    vbaSourceDownloadButton.hidden = false;
    vbaSourceTextDownloadButton.hidden = false;
    vbaSourceProgress.textContent =
      `Source ready (${vbaSourceModuleCount.toLocaleString()} modules, ${vbaSourceSize.toLocaleString()} bytes)`;
    vbaSourceMoreButton.hidden = false;
  } else {
    vbaSourceControls.hidden = true;
    vbaSourceDownloadButton.hidden = true;
    vbaSourceTextDownloadButton.hidden = true;
    vbaSourceOutput.textContent = result.vba_source_error || "VBA source code was not available.";
  }

  xlmFindingsGroup.hidden = !result.has_xlm;
  xlmFindingsGroup.open = true;
  xlmOffset = 0;
  xlmSize = result.xlm_size || 0;
  xlmMacroCount = result.xlm_macro_count || 0;
  xlmLargeDownloadWarning.hidden = xlmSize <= XLM_LARGE_DOWNLOAD_WARNING_BYTES;
  xlmDecoder = new TextDecoder("utf-8");
  xlmOutput.replaceChildren();
  xlmOutput.classList.toggle("code-unavailable", !xlmSize);
  if (xlmSize) {
    xlmControls.hidden = false;
    xlmDownloadButton.hidden = false;
    xlmProgress.textContent =
      `Macros ready (${xlmMacroCount.toLocaleString()} locations, ${xlmSize.toLocaleString()} bytes)`;
    xlmMoreButton.hidden = false;
  } else {
    xlmControls.hidden = true;
    xlmDownloadButton.hidden = true;
    xlmOutput.textContent = result.xlm_error || "XLM macro formulas were not available.";
  }

  pcodeOffset = 0;
  pcodeSize = result.pcode_size || 0;
  pcodeLargeDownloadWarning.hidden = pcodeSize <= PCODE_LARGE_DOWNLOAD_WARNING_BYTES;
  pcodeDecoder = new TextDecoder("utf-8");
  pcodeOutput.replaceChildren();
  pcodeOutput.classList.toggle("code-unavailable", !pcodeSize);
  if (pcodeSize) {
    pcodeControls.hidden = false;
    pcodeDownloadButton.hidden = false;
    pcodeProgress.textContent = `Disassembly ready (${pcodeSize.toLocaleString()} bytes)`;
    pcodeMoreButton.hidden = false;
  } else {
    pcodeControls.hidden = true;
    pcodeDownloadButton.hidden = true;
    pcodeOutput.textContent = result.pcode_error || "P-code disassembly was not available.";
  }
  openVbaSourceButton.hidden = !result.has_vba;
  openXlmSourceButton.hidden = !result.has_xlm;
  openPcodeSourceButton.hidden = !result.has_vba && !result.stomping;
  extractedSourceSection.hidden = openVbaSourceButton.hidden
    && openXlmSourceButton.hidden
    && openPcodeSourceButton.hidden;
  logDebug("analysis_complete", {
    duration_ms: Math.round(elapsedMilliseconds),
    has_vba: result.has_vba,
    has_xlm: result.has_xlm,
    stomping: result.stomping,
    stomping_items: result.stomping_items,
    vba_source_bytes: vbaSourceSize,
    vba_source_archive_bytes: vbaSourceArchiveSize,
    vba_source_module_count: vbaSourceModuleCount,
    pcode_bytes: pcodeSize,
    xlm_bytes: xlmSize,
    xlm_macro_count: xlmMacroCount,
    finding_count: (result.vba_source_findings || []).length + (result.pcode_findings || []).length + (result.xlm_findings || []).length,
  });

  const vbaSourceFindings = result.vba_source_findings || [];
  const pcodeFindings = result.pcode_findings || [];
  const xlmFindings = result.xlm_findings || [];
  vbaSourceWarningIcon.hidden = vbaSourceFindings.length === 0;
  xlmWarningIcon.hidden = xlmFindings.length === 0;
  const findingSignature = ({ type, keyword }) => `${type}\u0000${keyword.trim().toLowerCase()}`;
  const vbaSourceFindingSignatures = new Set(vbaSourceFindings.map(findingSignature));
  const vbaSourceFindingsBySignature = new Map(
    vbaSourceFindings.map((finding) => [findingSignature(finding), finding]),
  );
  const pcodeFindingSignatures = new Set(pcodeFindings.map(findingSignature));
  const unmatchedPcodeFindingSignatures = new Set(
    [...pcodeFindingSignatures].filter((signature) => !vbaSourceFindingSignatures.has(signature)),
  );
  const pcodeFindingsAreInVbaSource = [...pcodeFindingSignatures]
    .every((signature) => vbaSourceFindingSignatures.has(signature));
  const pcodeGroupShouldStartCollapsed = result.stomping === false && pcodeFindingsAreInVbaSource;
  vbaSourceFindingsGroup.open = true;
  pcodeFindingsGroup.open = !pcodeGroupShouldStartCollapsed;
  pcodeWarningIcon.hidden = !result.stomping && pcodeFindingsAreInVbaSource;
  pcodeMatchNote.hidden = !(result.stomping === false && pcodeFindings.length > 0 && pcodeFindingsAreInVbaSource);
  const totalFindings = vbaSourceFindings.length + pcodeFindings.length + xlmFindings.length;
  findingCount.textContent = `${totalFindings} ${totalFindings === 1 ? "indicator" : "indicators"}`;
  const renderFindingGroup = (
    listElement,
    countElement,
    findings,
    emptyMessage,
    options = {},
  ) => {
    const {
      matchingFindings = null,
      unmatchedFindingSignatures = null,
      sourceContextType = null,
    } = options;
    countElement.textContent = `${findings.length}`;
    if (findings.length === 0) {
      const empty = document.createElement("div");
      empty.className = "no-findings";
      empty.textContent = emptyMessage;
      listElement.replaceChildren(empty);
    } else {
      listElement.replaceChildren(...findings.map((finding) => {
      const row = document.createElement("div");
      row.className = "finding-row";
      const type = document.createElement("div");
      type.className = "finding-type";
      type.textContent = finding.type;
      const content = document.createElement("div");
      content.className = "finding-content";
      const keyword = document.createElement("div");
      keyword.className = "finding-keyword";
      keyword.textContent = finding.keyword;
      const description = document.createElement("div");
      description.className = "finding-description";
      description.textContent = finding.description.replace(" (option --decode to see all)", "");
      content.append(keyword, description);
      if (unmatchedFindingSignatures?.has(findingSignature(finding))) {
        const unmatched = document.createElement("div");
        unmatched.className = "pcode-unmatched";
        unmatched.textContent = "⚠ No matching VBA indicator";
        content.append(unmatched);
      }
      if (sourceContextType === "pcode") {
        const sourceDetails = document.createElement("details");
        sourceDetails.className = "pcode-source-lines";
        const sourceSummary = document.createElement("summary");
        sourceSummary.textContent = "Show matching P-code source block";
        const sourceContent = document.createElement("div");
        sourceContent.className = "pcode-source-context";
        sourceDetails.append(sourceSummary, sourceContent);
        let loading = false;
        sourceDetails.addEventListener("toggle", () => {
          if (!sourceDetails.open || loading) return;
          loading = true;
          sourceContent.textContent = "Loading P-code source block…";
          void findPcodeSourceContext(finding.keyword).then((sourceContext) => {
            sourceContent.replaceChildren();
            if (sourceContext) {
              sourceSummary.textContent = `Matching P-code Line #${sourceContext.number}`;
              if (sourceContext.module_stream) {
                const moduleStream = document.createElement("div");
                moduleStream.className = "pcode-source-module";
                moduleStream.textContent = `Module streams: ${sourceContext.module_stream}`;
                sourceContent.append(moduleStream);
              }
              const sourceCode = document.createElement("pre");
              sourceCode.textContent = sourceContext.text;
              sourceContent.append(sourceCode);
            } else {
              sourceSummary.textContent = "P-code source block (no direct match)";
              sourceContent.textContent = "This indicator may have been detected in decoded P-code content.";
            }
          }).catch((error) => {
            logDebug("pcode_source_context_error", getErrorDetails(error));
            sourceContent.textContent = "Could not load P-code source block. Close and reopen to retry.";
            loading = false;
          });
        });
        content.append(sourceDetails);
      }
      if (sourceContextType === "vba") {
        content.append(createVbaSourceDetails(finding.keyword));
      }
      const matchingFinding = matchingFindings?.get(findingSignature(finding));
      if (matchingFinding) {
        const matchDetails = document.createElement("details");
        matchDetails.className = "finding-match";
        const matchSummary = document.createElement("summary");
        matchSummary.textContent = "Likely corresponding VBA indicator";
        const matchContent = document.createElement("div");
        matchContent.className = "finding-match-content";
        const matchType = document.createElement("div");
        matchType.className = "finding-type";
        matchType.textContent = matchingFinding.type;
        const matchKeyword = document.createElement("div");
        matchKeyword.className = "finding-keyword";
        matchKeyword.textContent = matchingFinding.keyword;
        const matchDescription = document.createElement("div");
        matchDescription.className = "finding-description";
        matchDescription.textContent = matchingFinding.description.replace(" (option --decode to see all)", "");
        matchContent.append(matchType, matchKeyword, matchDescription);
        matchDetails.append(matchSummary, matchContent);
        content.append(matchDetails);
      }
      row.append(type, content);
      return row;
      }));
    }
  };
  renderFindingGroup(
    vbaSourceFindingsList,
    vbaSourceFindingCount,
    vbaSourceFindings,
    result.has_vba ? "No suspicious indicators found in VBA source code." : "No VBA source code available.",
    { sourceContextType: "vba" },
  );
  renderFindingGroup(
    pcodeFindingsList,
    pcodeFindingCount,
    pcodeFindings,
    result.pcode_size ? "No suspicious indicators found in P-code." : "P-code was not available.",
    {
      matchingFindings: result.stomping === false ? vbaSourceFindingsBySignature : null,
      unmatchedFindingSignatures: unmatchedPcodeFindingSignatures,
      sourceContextType: "pcode",
    },
  );
  renderFindingGroup(
    xlmFindingsList,
    xlmFindingCount,
    xlmFindings,
    result.has_xlm ? "No suspicious indicators found in XLM macros." : "No XLM macros available.",
  );

  analysisTime.textContent = `COMPLETED IN ${(elapsedMilliseconds / 1000).toFixed(1)}S`;
  resultsElement.hidden = false;
}

async function analyzeSelectedFile() {
  if (!selectedFile || scanning) return;

  setBusy(true, "Analyzing locally in your browser…");
  setStatus("Reading document and calculating SHA-256…");
  try {
    pcodePath = "";
    vbaSourcePath = "";
    vbaSourceArchivePath = "";
    xlmPath = "";
    pcodeSize = 0;
    vbaSourceSize = 0;
    vbaSourceArchiveSize = 0;
    vbaSourceModuleCount = 0;
    xlmSize = 0;
    xlmMacroCount = 0;

    const file = selectedFile;
    const startedAt = performance.now();
    if (!globalThis.crypto?.subtle) {
      throw new Error("SHA-256 hashing is unavailable in this browser context.");
    }
    const fileBytes = await file.arrayBuffer();
    const digest = await globalThis.crypto.subtle.digest("SHA-256", fileBytes);
    const sha256 = Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, "0")).join("");
    logDebug("analysis_started", { input_bytes: fileBytes.byteLength });
    setStatus("Preparing Python and oletools for local analysis…");
    const scanStartedAt = performance.now();
    const scan = await requestWorker("scan", { fileBytes, fileName: file.name }, [fileBytes]);
    logDuration("Document analysis", scanStartedAt);
    pcodePath = scan.paths.pcode;
    vbaSourcePath = scan.paths.vbaSource;
    vbaSourceArchivePath = scan.paths.vbaArchive;
    xlmPath = scan.paths.xlm;
    renderResults(JSON.parse(scan.result), performance.now() - startedAt, sha256, file.name);
    runtimeNoteContainer.hidden = true;
    setStatus("✅ Analysis complete.");
  } catch (error) {
    logDebug("analysis_error", getErrorDetails(error));
    setStatus(`Analysis could not be completed: ${error.message}`, true);
  } finally {
    setBusy(false);
  }
}

browseButton.addEventListener("click", () => fileInput.click());
dropZone.addEventListener("click", (event) => {
  if (event.target.closest("button")) return;
  fileInput.click();
});
fileInput.addEventListener("change", () => {
  if (fileInput.files?.[0]) setSelectedFile(fileInput.files[0]);
});
removeFileButton.addEventListener("click", clearSelectedFile);
openVbaSourceButton.addEventListener("click", () => openSourcePreview("vba"));
openXlmSourceButton.addEventListener("click", () => openSourcePreview("xlm"));
openPcodeSourceButton.addEventListener("click", () => openSourcePreview("pcode"));
sourceDialogCloseButton.addEventListener("click", () => sourceCodeDialog.close());
sourceCodeDialog.addEventListener("click", (event) => {
  if (event.target === sourceCodeDialog) sourceCodeDialog.close();
});
vbaSourceLoadAboveButton.addEventListener("click", loadMoreVbaSourceAbove);
vbaSourceLoadBelowButton.addEventListener("click", loadMoreVbaSourceBelow);
vbaSourceMoreButton.addEventListener("click", appendVbaSourceChunk);
vbaSourceDownloadButton.addEventListener("click", downloadVbaSourceArchive);
vbaSourceTextDownloadButton.addEventListener("click", downloadVbaSourceText);
xlmMoreButton.addEventListener("click", appendXlmChunk);
xlmDownloadButton.addEventListener("click", downloadXlmMacros);
pcodeMoreButton.addEventListener("click", appendPcodeChunk);
pcodeDownloadButton.addEventListener("click", async () => {
  if (pcodeDownloadInProgress || !pcodeSize || !selectedFile) return;
  const filename = getPcodeFilename(selectedFile.name);
  const downloadMethod = "blob";
  pcodeDownloadInProgress = true;
  pcodeDownloadButton.disabled = true;
  removeFileButton.disabled = true;
  const startedAt = performance.now();
  let offset = 0;
  let written = 0;
  logDebug("pcode_download_started", {
    bytes: pcodeSize,
    download_method: downloadMethod,
  });
  try {
    pcodeProgress.textContent = "Preparing p-code download…";
    const chunks = [];
    for (; offset < pcodeSize; offset += PCODE_DOWNLOAD_CHUNK_SIZE) {
      const requestedBytes = Math.min(PCODE_DOWNLOAD_CHUNK_SIZE, pcodeSize - offset);
      const bytes = await readPcodeBytes(offset, requestedBytes);
      if (bytes.length === 0) throw new Error("The p-code file ended before the expected size.");
      chunks.push(bytes);
      written += bytes.length;
      pcodeProgress.textContent = `Downloading ${written.toLocaleString()} of ${pcodeSize.toLocaleString()} bytes`;
      if (written % (PCODE_DOWNLOAD_CHUNK_SIZE * 16) === 0) {
        logDebug("pcode_download_progress", { bytes: written, total_bytes: pcodeSize });
      }
    }
    const blob = new Blob(chunks, { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    link.click();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
    pcodeProgress.textContent = `Download complete (${written.toLocaleString()} bytes)`;
    logDebug("pcode_download_complete", { bytes: written });
  } catch (error) {
    logDebug("pcode_download_error", getErrorDetails(error));
    pcodeProgress.textContent = "Download did not complete";
    setStatus(`Could not download p-code: ${error.message}`, true);
  } finally {
    pcodeDownloadButton.disabled = false;
    removeFileButton.disabled = false;
    pcodeDownloadInProgress = false;
    logDuration("P-code download", startedAt);
  }
});
debugLogButton.addEventListener("click", downloadDebugLog);
copyHashButton.addEventListener("click", async () => {
  try {
    if (!navigator.clipboard?.writeText) {
      throw new Error("Clipboard access is unavailable in this browser context.");
    }
    await navigator.clipboard.writeText(documentHash.textContent);
    copyHashButton.classList.add("copied");
    copyHashButton.setAttribute("aria-label", "SHA-256 copied");
    copyHashButton.title = "SHA-256 copied";
    copyStatus.textContent = "SHA-256 copied to clipboard.";
    window.clearTimeout(copyFeedbackTimeout);
    copyFeedbackTimeout = window.setTimeout(() => {
      copyHashButton.classList.remove("copied");
      copyHashButton.setAttribute("aria-label", "Copy SHA-256");
      copyHashButton.title = "Copy SHA-256";
      copyStatus.textContent = "";
    }, 1500);
  } catch (error) {
    setStatus(`Could not copy SHA-256: ${error.message}`, true);
  }
});
for (const eventName of ["dragenter", "dragover"]) {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.add("dragging");
  });
}
for (const eventName of ["dragleave", "drop"]) {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.remove("dragging");
  });
}
dropZone.addEventListener("drop", (event) => {
  const file = event.dataTransfer?.files?.[0];
  if (file) setSelectedFile(file);
});

analysisWorker.postMessage({ type: "initialize" });
