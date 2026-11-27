import { renderHighlightedCode, StreamingSyntaxHighlighter, ViewportSyntaxHighlighter } from "./syntax-highlighter.ts";

const PYODIDE_VERSION = "0.27.7";
const OLETOOLS_VERSION = "0.60.3";
const PCODE_CHUNK_SIZE = 16_384;
const PCODE_PREVIEW_LIMIT = 524_288;
const SYNTAX_HIGHLIGHT_CHUNK_SIZE = 8_192;
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

type EngineState = "loading" | "ready" | "error";

interface Finding {
  type: string;
  keyword: string;
  description: string;
}

interface VbaSourceLocation extends Finding {
  module_name: string;
  line_number: number;
}

interface SourceLine {
  number: number | null | undefined;
  text: string;
  matched: boolean;
  byteOffset: number;
  endByteOffset: number;
}

interface VbaSourceContext {
  moduleName: string;
  lineNumber: number;
  lines: SourceLine[];
  startByteOffset: number;
  endByteOffset: number;
}

interface PcodeSourceContext {
  number: number;
  text: string;
  module_stream: string | null;
}

interface PcodeBlock {
  matched: boolean;
  context: PcodeSourceContext;
}

interface FindingGroupOptions {
  matchingFindings?: Map<string, Finding> | null;
  unmatchedFindingSignatures?: Set<string> | null;
  sourceContextType?: "pcode" | "vba" | null;
}

interface SummaryCard {
  label: string;
  value: string;
  tone: string;
  note?: { text: string; issues: number[] };
  mismatches?: Array<{ kind: string; value: string }>;
}

interface ScanResult {
  has_vba: boolean;
  has_xlm: boolean;
  stomping: boolean | null;
  stomping_items: Array<{ kind: string; value: string }>;
  pcode_size: number;
  pcode_error: string | null;
  vba_source_size: number;
  vba_source_archive_size: number;
  vba_source_module_count: number;
  vba_source_error: string | null;
  xlm_size: number;
  xlm_error: string | null;
  xlm_macro_count: number;
  project_codepages: Array<{ value: number; name: string | null }>;
  vba_source_findings: Finding[];
  vba_source_locations: VbaSourceLocation[];
  pcode_findings: Finding[];
  xlm_findings: Finding[];
}

interface ScanWorkerResult {
  result: string;
  paths: { pcode: string; vbaSource: string; vbaArchive: string; xlm: string };
}

function requireElement<T extends HTMLElement = HTMLElement>(selector: string): T {
  const element = document.querySelector<T>(selector);
  if (!element) throw new Error(`Required page element is missing: ${selector}`);
  return element;
}

const fileInput = requireElement<HTMLInputElement>("#file-input");
const dropZone = requireElement("#drop-zone");
const browseButton = requireElement<HTMLButtonElement>("#browse-button");
const selectedFileElement = requireElement("#selected-file");
const fileNameElement = requireElement("#file-name");
const fileSizeElement = requireElement("#file-size");
const removeFileButton = requireElement<HTMLButtonElement>("#remove-file");
const runtimeNote = requireElement("#runtime-note");
const runtimeNoteContainer = requireElement(".runtime-note");
const runtimeIndicator = requireElement(".runtime-dot");
const statusMessage = requireElement("#status-message");
const statusText = requireElement("#status-text");
const analysisSpinner = requireElement("#analysis-spinner");
const analysisProgress = requireElement("#analysis-progress");
const debugLogButton = requireElement<HTMLButtonElement>("#download-debug-log");
const resultsElement = requireElement("#results");
const resultSummary = requireElement("#result-summary");
const findingsSection = requireElement("#findings-section");
const noSuspiciousIndicators = requireElement("#no-suspicious-indicators");
const extractedSourceSection = requireElement("#extracted-source-section");
const openVbaSourceButton = requireElement<HTMLButtonElement>("#open-vba-source");
const openXlmSourceButton = requireElement<HTMLButtonElement>("#open-xlm-source");
const openPcodeSourceButton = requireElement<HTMLButtonElement>("#open-pcode-source");
const sourceCodeDialog = requireElement<HTMLDialogElement>("#source-code-dialog");
const sourceDialogTitle = requireElement("#source-dialog-title");
const sourceDialogCloseButton = requireElement<HTMLButtonElement>("#source-dialog-close");
const vbaSourceSection = requireElement("#vba-source-section");
const vbaSourcePreview = requireElement("#vba-source-preview");
const vbaSourceMatchLocation = requireElement("#vba-source-match-location");
const vbaSourceOutput = requireElement("#vba-source-output");
const vbaSourceControls = requireElement("#vba-source-controls");
const vbaSourceProgress = requireElement("#vba-source-progress");
const vbaSourceHighlightStatus = requireElement("#vba-source-highlight-status");
const vbaSourceMoreButton = requireElement<HTMLButtonElement>("#vba-source-more");
const vbaSourceLoadAboveButton = requireElement<HTMLButtonElement>("#vba-source-load-above");
const vbaSourceLoadBelowButton = requireElement<HTMLButtonElement>("#vba-source-load-below");
const vbaSourceDownloadButton = requireElement<HTMLButtonElement>("#vba-source-download");
const vbaSourceTextDownloadButton = requireElement<HTMLButtonElement>("#vba-source-text-download");
const vbaSourceLargeDownloadWarning = requireElement("#vba-source-large-download-warning");
const xlmFindingsGroup = requireElement<HTMLDetailsElement>("#xlm-findings-group");
const xlmFindingsList = requireElement("#xlm-findings-list");
const xlmFindingCount = requireElement("#xlm-finding-count");
const xlmSection = requireElement("#xlm-section");
const xlmOutput = requireElement("#xlm-output");
const xlmControls = requireElement("#xlm-controls");
const xlmProgress = requireElement("#xlm-progress");
const xlmMoreButton = requireElement<HTMLButtonElement>("#xlm-more");
const xlmDownloadButton = requireElement<HTMLButtonElement>("#xlm-download");
const xlmLargeDownloadWarning = requireElement("#xlm-large-download-warning");
const pcodeSection = requireElement("#pcode-section");
const pcodeOutput = requireElement("#pcode-output");
const pcodeControls = requireElement("#pcode-controls");
const pcodeProgress = requireElement("#pcode-progress");
const pcodeHighlightButton = requireElement<HTMLButtonElement>("#pcode-highlight");
const pcodeHighlightStatus = requireElement("#pcode-highlight-status");
const pcodeMoreButton = requireElement<HTMLButtonElement>("#pcode-more");
const pcodeDownloadButton = requireElement<HTMLButtonElement>("#pcode-download");
const pcodeLargeDownloadWarning = requireElement("#pcode-large-download-warning");
const vbaSourceFindingsList = requireElement("#vba-source-findings-list");
const vbaSourceFindingCount = requireElement("#vba-source-finding-count");
const vbaSourceFindingsGroup = requireElement<HTMLDetailsElement>('[aria-labelledby="vba-source-findings-heading"]');
const vbaSourceWarningIcon = requireElement("#vba-source-warning-icon");
const xlmWarningIcon = requireElement("#xlm-warning-icon");
const pcodeFindingsGroup = requireElement<HTMLDetailsElement>("#pcode-findings-group");
const pcodeWarningIcon = requireElement("#pcode-warning-icon");
const pcodeMatchNote = requireElement("#pcode-match-note");
const pcodeFindingsList = requireElement("#pcode-findings-list");
const pcodeFindingCount = requireElement("#pcode-finding-count");
const analysisTime = requireElement("#analysis-time");
const documentHash = requireElement("#document-hash");
const documentFilename = requireElement("#document-filename");
const documentCodepage = requireElement("#document-codepage");
const copyHashButton = requireElement<HTMLButtonElement>("#copy-hash-button");
const copyStatus = requireElement("#copy-status");
const vbaSourceHighlighter = new ViewportSyntaxHighlighter(
  vbaSourceOutput,
  vbaSourceOutput.closest("pre")!,
  "vba",
);
const pcodeHighlighter = new StreamingSyntaxHighlighter(pcodeOutput, "pcode");

let selectedFile: File | null = null;
const analysisWorker = new Worker(new URL("./analysis-worker.js", import.meta.url), { type: "module" });
const workerRequests = new Map<number, { resolve: (value: unknown) => void; reject: (error: Error) => void }>();
let workerRequestSequence = 0;
let scanning = false;
let engineState: EngineState = "loading";
let copyFeedbackTimeout: ReturnType<typeof setTimeout> | null = null;
let vbaSourceOffset = 0;
let vbaSourcePath = "";
let vbaSourceSize = 0;
let vbaSourceModuleCount = 0;
let vbaSourceArchivePath = "";
let vbaSourceArchiveSize = 0;
let vbaSourceLoading = false;
let vbaSourceHighlighting = false;
let vbaSourceDownloadInProgress = false;
let vbaSourceDecoder = new TextDecoder("utf-8");
let vbaSourceMatchContext: VbaSourceContext | null = null;
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
let pcodeHighlighting = false;
let pcodeHighlightingInProgress = false;
let pcodeSyntaxHighlightGeneration = 0;
let pcodeDownloadInProgress = false;
let pcodeDecoder = new TextDecoder("utf-8");
const pcodeSourceContextCache = new Map<string, Promise<PcodeSourceContext | null>>();
const vbaSourceContextCache = new Map<string, Promise<VbaSourceContext | null>>();
const debugEvents: Array<Record<string, unknown>> = [];
if (debugMode) {
  try {
    const savedEvents = JSON.parse(localStorage.getItem(DEBUG_LOG_KEY) || "[]");
    if (Array.isArray(savedEvents)) debugEvents.push(...savedEvents.slice(-500));
  } catch {}
}

debugLogButton.hidden = !debugMode;

function setEngineState(state: EngineState) {
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

function logDebug(event: string, details: Record<string, unknown> = {}) {
  if (!debugMode) return;
  debugEvents.push({ timestamp: new Date().toISOString(), elapsed_ms: Math.round(performance.now()), event, ...details });
  if (debugEvents.length > 500) debugEvents.shift();
  try { localStorage.setItem(DEBUG_LOG_KEY, JSON.stringify(debugEvents)); } catch {}
}

function getErrorDetails(error: unknown) {
  const message = error instanceof Error ? error.message : String(error);
  return {
    name: error instanceof Error ? error.name : "Error",
    message: message.replace(selectedFile?.name || "\0", "[selected file]"),
  };
}

function getErrorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
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

function requestWorker<T = unknown>(type: string, payload: Record<string, unknown> = {}, transfer: Transferable[] = []): Promise<T> {
  const requestId = ++workerRequestSequence;
  return new Promise<T>((resolve, reject) => {
    workerRequests.set(requestId, { resolve: (value) => resolve(value as T), reject });
    try {
      analysisWorker.postMessage({ type, requestId, ...payload }, transfer);
    } catch (error) {
      workerRequests.delete(requestId);
      reject(error instanceof Error ? error : new Error(String(error)));
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
    setStatus(`Analysis could not be completed: ${getErrorMessage(error)}`, true);
  } else {
    setStatus(`Analysis engine could not be loaded: ${getErrorMessage(error)}`, true);
  }
});

function logDuration(label: string, startedAt: number) {
  const duration = Number((performance.now() - startedAt).toFixed(1));
  console.info(`[VBA-Scan] ${label}: ${duration} ms`);
  logDebug("timing", { label, duration_ms: duration });
}

async function readPcodeBytes(offset: number, size: number): Promise<Uint8Array> {
  return readTemporaryFileBytes(pcodePath, offset, size);
}

async function findPcodeSourceContext(keyword: string): Promise<PcodeSourceContext | null> {
  const cacheKey = keyword.trim().toLowerCase();
  const cached = pcodeSourceContextCache.get(cacheKey);
  if (cached) return cached;

  const search = (async () => {
    const pattern = new RegExp(keyword.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "i");
    const decoder = new TextDecoder("utf-8");
    let offset = 0;
    let pendingLine = "";
    let inModuleStreams = false;
    let moduleStream: string | null = null;
    let block: PcodeBlock | null = null;
    const processLine = (line: string): PcodeSourceContext | null => {
      const trimmedLine = line.trim();
      if (trimmedLine === "Module streams:") {
        inModuleStreams = true;
        return null;
      }
      const moduleStreamMatch = inModuleStreams && trimmedLine.match(/^(.+ - \d+ bytes)$/);
      if (moduleStreamMatch) {
        if (block?.matched) return block.context;
        moduleStream = moduleStreamMatch[1] ?? "";
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
      pendingLine = lines.pop() ?? "";
      for (const line of lines) {
        const context = processLine(line);
        if (context) return context;
      }
    }
    if (pendingLine) {
      const context = processLine(pendingLine);
      if (context) return context;
    }
    const finalBlock = block as PcodeBlock | null;
    return finalBlock?.matched ? finalBlock.context : null;
  })();

  pcodeSourceContextCache.set(cacheKey, search);
  try {
    return await search;
  } catch (error) {
    pcodeSourceContextCache.delete(cacheKey);
    throw error;
  }
}

async function findVbaSourceContext(location: VbaSourceLocation): Promise<VbaSourceContext | null> {
  if (!vbaSourceSize) return null;
  const cacheKey = `${location.module_name}\u0000${location.line_number}`;
  const cached = vbaSourceContextCache.get(cacheKey);
  if (cached) return cached;

  const search = (async (): Promise<VbaSourceContext | null> => {
    const decoder = new TextDecoder("utf-8");
    const encoder = new TextEncoder();
    let offset = 0;
    let scannedOffset = 0;
    let pendingLine = "";
    let moduleName = "";
    let moduleLineNumber = 0;
    let isTargetModule = false;
    let precedingLines: SourceLine[] = [];
    let matchedLine: SourceLine | null = null;
    let contextLines: SourceLine[] = [];
    let followingLinesRemaining = 2;
    const getContext = (matched: SourceLine): VbaSourceContext => ({
      moduleName,
      lineNumber: matched.number!,
      lines: contextLines,
      startByteOffset: contextLines[0].byteOffset,
      endByteOffset: contextLines.at(-1)!.endByteOffset,
    });
    const processLine = (line: string, newlineBytes: number): VbaSourceContext | null => {
      const byteOffset = scannedOffset;
      scannedOffset += encoder.encode(line).length + newlineBytes;
      const sourceText = line.replace(/\r$/, "");
      const moduleHeader = sourceText.match(/^' ===== (.+) =====$/);
      if (moduleHeader) {
        if (matchedLine) return getContext(matchedLine);
        moduleName = moduleHeader[1] ?? "";
        moduleLineNumber = 0;
        isTargetModule = moduleName === location.module_name;
        precedingLines = [];
        return null;
      }
      if (!isTargetModule) return null;

      const sourceLine: SourceLine = {
        number: ++moduleLineNumber,
        text: sourceText,
        matched: false,
        byteOffset,
        endByteOffset: scannedOffset,
      };
      if (matchedLine) {
        contextLines.push(sourceLine);
        followingLinesRemaining -= 1;
        return followingLinesRemaining === 0 ? getContext(matchedLine) : null;
      }
      if (sourceLine.number === location.line_number) {
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
      pendingLine = lines.pop() ?? "";
      for (const line of lines) {
        const context = processLine(line, 1);
        if (context) return context;
      }
    }
    if (pendingLine) {
      const context = processLine(pendingLine, 0);
      if (context) return context;
    }
    return matchedLine ? getContext(matchedLine) : null;
  })();

  vbaSourceContextCache.set(cacheKey, search);
  try {
    return await search;
  } catch (error) {
    vbaSourceContextCache.delete(cacheKey);
    throw error;
  }
}

function renderVbaSourceContext(container: HTMLElement, context: VbaSourceContext) {
  const module = document.createElement("div");
  module.className = "vba-source-context-module";
  module.textContent = `${context.moduleName} · line ${context.lineNumber}`;
  const code = document.createElement("pre");
  for (const line of context.lines) {
    const sourceLine = document.createElement("span");
    sourceLine.className = line.matched ? "vba-source-context-line is-match" : "vba-source-context-line";
    const prefix = line.number == null ? "" : `${String(line.number).padStart(4, " ")}  `;
    renderHighlightedCode(sourceLine, line.text, "vba");
    sourceLine.prepend(document.createTextNode(prefix));
    code.append(sourceLine, document.createTextNode("\n"));
  }
  container.replaceChildren(module, code);
}

function renderVbaSourceWindow(context: VbaSourceContext) {
  const nodes = [];
  for (const line of context.lines) {
    const sourceLine = document.createElement("span");
    sourceLine.className = line.matched ? "vba-source-context-line is-match" : "vba-source-context-line";
    renderHighlightedCode(sourceLine, line.text, "vba");
    nodes.push(sourceLine, document.createTextNode("\n"));
  }
  vbaSourceOutput.replaceChildren(...nodes);
}

function updateVbaSourceProgress(context: VbaSourceContext) {
  const lineNumbers = context.lines
    .map((line) => line.number)
    .filter((number): number is number => typeof number === "number");
  if (lineNumbers.length === 0) {
    vbaSourceProgress.textContent = `Showing source from ${context.moduleName}`;
    return;
  }
  const firstLine = Math.min(...lineNumbers);
  const lastLine = Math.max(...lineNumbers);
  const range = firstLine === lastLine ? `line ${firstLine}` : `lines ${firstLine}-${lastLine}`;
  vbaSourceProgress.textContent = `Showing ${range} of ${context.moduleName}`;
}

async function loadMoreVbaSourceAbove() {
  const context = vbaSourceMatchContext;
  if (!context || vbaSourceWindowLoading || context.startByteOffset <= 0) return;
  const pre = vbaSourceOutput.parentElement;
  if (!pre) return;
  vbaSourceWindowLoading = true;
  vbaSourceLoadAboveButton.disabled = true;
  vbaSourceLoadBelowButton.disabled = true;
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
    let precedingLineNumber = context.lines.find((line) => typeof line.number === "number")?.number;
    const addedLines = lines.map((line, index) => {
      const hasNewline = index < lines.length - 1 || text.endsWith("\n");
      const endByteOffset = lineOffset + encoder.encode(line).length + Number(hasNewline);
      const sourceText = line.replace(/\r$/, "");
      const sourceLine: SourceLine = {
        number: null,
        text: sourceText,
        matched: false,
        byteOffset: lineOffset,
        endByteOffset,
      };
      lineOffset = endByteOffset;
      return sourceLine;
    });
    for (let index = addedLines.length - 1; index >= 0; index -= 1) {
      const sourceLine = addedLines[index];
      if (/^' ===== .+ =====$/.test(sourceLine.text)) precedingLineNumber = undefined;
      else if (precedingLineNumber != null) {
        precedingLineNumber -= 1;
        sourceLine.number = precedingLineNumber;
      }
    }
    if (addedLines.length) {
      context.lines.unshift(...addedLines);
      context.startByteOffset = addedLines[0].byteOffset;
      renderVbaSourceWindow(context);
      updateVbaSourceProgress(context);
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
      let nextLineNumber = context.lines.at(-1)?.number;
      const addedLines = lines.map((line, index) => {
        const hasNewline = index < lines.length - 1 || endsWithNewline;
        const endByteOffset = lineOffset + encoder.encode(line).length + Number(hasNewline);
        const sourceText = line.replace(/\r$/, "");
        const moduleHeader = sourceText.match(/^' ===== (.+) =====$/);
        if (moduleHeader) nextLineNumber = undefined;
        else if (nextLineNumber !== undefined && nextLineNumber !== null) nextLineNumber += 1;
        const sourceLine = {
          number: moduleHeader || nextLineNumber == null ? null : nextLineNumber,
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
      updateVbaSourceProgress(context);
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

function openVbaSourceMatch(context: VbaSourceContext) {
  for (const section of [vbaSourceSection, xlmSection, pcodeSection]) {
    section.hidden = section !== vbaSourceSection;
  }
  vbaSourceMatchContext = context;
  vbaSourceWindowReadOffset = context.endByteOffset;
  vbaSourceWindowPendingBytes = new Uint8Array();
  vbaSourceOffset = 0;
  resetVbaSyntaxHighlighting();
  vbaSourceOutput.replaceChildren();
  vbaSourceMatchLocation.textContent = `${context.moduleName} · line ${context.lineNumber}`;
  vbaSourceMatchLocation.hidden = false;
  vbaSourceLoadAboveButton.hidden = context.startByteOffset === 0;
  vbaSourceLoadBelowButton.hidden = vbaSourceWindowReadOffset >= vbaSourceSize;
  vbaSourceMoreButton.hidden = true;
  vbaSourceControls.hidden = false;
  vbaSourceOutput.classList.add("vba-match-window");
  renderVbaSourceWindow(context);
  updateVbaSourceProgress(context);
  vbaSourcePreview.hidden = false;
  sourceDialogTitle.textContent = "VBA source code";
  sourceCodeDialog.showModal();
}

function createVbaSourceDetails(location: VbaSourceLocation) {
  const details = document.createElement("details");
  details.className = "vba-source-lines";
  const summary = document.createElement("summary");
  summary.textContent = `Matched ${location.module_name}, line ${location.line_number}`;
  const content = document.createElement("div");
  content.className = "vba-source-context";
  details.append(summary, content);
  let loading = false;
  let loaded = false;
  details.addEventListener("toggle", () => {
    if (!details.open || loading || loaded) return;
    loading = true;
    content.textContent = "Searching extracted VBA source…";
    void findVbaSourceContext(location).then((context) => {
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

async function readTemporaryFileBytes(path: string, offset: number, size: number): Promise<Uint8Array> {
  return requestWorker<Uint8Array>("read-file", { path, offset, size });
}

function waitForAnimationFrame(): Promise<void> {
  return new Promise((resolve) => requestAnimationFrame(() => resolve()));
}

async function applySyntaxHighlightingInFrames(
  source: string,
  output: HTMLElement,
  highlighter: StreamingSyntaxHighlighter,
  isCurrent: () => boolean,
): Promise<boolean> {
  output.replaceChildren();
  highlighter.reset();
  await waitForAnimationFrame();
  for (let offset = 0; offset < source.length; offset += SYNTAX_HIGHLIGHT_CHUNK_SIZE) {
    if (!isCurrent()) return false;
    const end = Math.min(offset + SYNTAX_HIGHLIGHT_CHUNK_SIZE, source.length);
    highlighter.append(source.slice(offset, end), end === source.length);
    await waitForAnimationFrame();
  }
  return isCurrent();
}

function showSyntaxHighlightStatus(status: HTMLElement, message: string, loading = false) {
  status.replaceChildren();
  if (loading) {
    const spinner = document.createElement("span");
    spinner.className = "syntax-highlight-spinner";
    spinner.setAttribute("aria-hidden", "true");
    status.append(spinner);
  }
  const text = document.createElement("span");
  text.textContent = message;
  status.append(text);
}

function resetVbaSyntaxHighlighting() {
  vbaSourceHighlighting = false;
  vbaSourceHighlightStatus.hidden = true;
  vbaSourceMoreButton.disabled = false;
  vbaSourceHighlighter.reset();
}

function resetPcodeSyntaxHighlighting() {
  pcodeSyntaxHighlightGeneration++;
  pcodeHighlighting = false;
  pcodeHighlightingInProgress = false;
  pcodeHighlightButton.disabled = false;
  pcodeHighlightButton.textContent = "Enable syntax highlighting";
  pcodeHighlightStatus.hidden = true;
  pcodeMoreButton.disabled = false;
  pcodeHighlighter.reset();
}

function updateVbaSyntaxHighlightStatus(metrics: ReturnType<ViewportSyntaxHighlighter["enable"]>) {
  vbaSourceHighlightStatus.hidden = false;
  showSyntaxHighlightStatus(
    vbaSourceHighlightStatus,
    metrics.sourceChars
      ? `Viewport highlighting active · ${metrics.sourceChars.toLocaleString()} chars · ${metrics.totalLines.toLocaleString()} lines`
      : "Viewport highlighting active · waiting for source",
  );
}

function startVbaSyntaxHighlighting() {
  if (vbaSourceHighlighting) return;
  try {
    const metrics = vbaSourceHighlighter.enable(vbaSourceOutput.textContent ?? "");
    vbaSourceHighlighting = true;
    updateVbaSyntaxHighlightStatus(metrics);
  } catch (error) {
    logDebug("vba_syntax_highlight_error", getErrorDetails(error));
    vbaSourceHighlighting = false;
    vbaSourceHighlighter.reset();
  }
}

async function enablePcodeSyntaxHighlighting() {
  if (pcodeHighlighting || pcodeHighlightingInProgress) return;
  pcodeHighlightingInProgress = true;
  const generation = ++pcodeSyntaxHighlightGeneration;
  let source = "";
  pcodeHighlightButton.disabled = true;
  pcodeHighlightStatus.hidden = false;
  showSyntaxHighlightStatus(pcodeHighlightStatus, "Applying syntax highlighting…", true);
  pcodeMoreButton.disabled = true;
  try {
    while (pcodeLoading && generation === pcodeSyntaxHighlightGeneration) {
      await waitForAnimationFrame();
    }
    if (generation !== pcodeSyntaxHighlightGeneration) return;
    source = pcodeOutput.textContent ?? "";
    const applied = await applySyntaxHighlightingInFrames(
      source,
      pcodeOutput,
      pcodeHighlighter,
      () => generation === pcodeSyntaxHighlightGeneration,
    );
    if (applied) {
      pcodeHighlighting = true;
      pcodeHighlightButton.textContent = "Syntax highlighting enabled";
      showSyntaxHighlightStatus(pcodeHighlightStatus, "Syntax highlighting enabled");
    }
  } catch (error) {
    logDebug("pcode_syntax_highlight_error", getErrorDetails(error));
    if (generation === pcodeSyntaxHighlightGeneration) {
      pcodeOutput.textContent = source;
      pcodeHighlighter.reset();
    }
  } finally {
    if (generation === pcodeSyntaxHighlightGeneration) {
      pcodeHighlightingInProgress = false;
      pcodeHighlightStatus.hidden = !pcodeHighlighting;
      pcodeHighlightButton.disabled = pcodeHighlighting;
      pcodeMoreButton.disabled = false;
    }
  }
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
    const previewComplete = vbaSourceOffset >= Math.min(vbaSourceSize, VBA_SOURCE_PREVIEW_LIMIT);
    const chunk = vbaSourceDecoder.decode(bytes, { stream: !previewComplete });
    if (vbaSourceHighlighting) {
      const metrics = vbaSourceHighlighter.appendSource(chunk);
      if (metrics) updateVbaSyntaxHighlightStatus(metrics);
      vbaSourceMoreButton.disabled = false;
    } else {
      vbaSourceOutput.append(document.createTextNode(chunk));
    }
    vbaSourceProgress.textContent =
      `Preview: ${vbaSourceOffset.toLocaleString()} of ${vbaSourceSize.toLocaleString()} bytes · ${vbaSourceModuleCount.toLocaleString()} modules`;
    vbaSourceMoreButton.hidden = vbaSourceOffset >= Math.min(vbaSourceSize, VBA_SOURCE_PREVIEW_LIMIT);
    logDebug("vba_source_preview_chunk", { bytes: bytes.length, total_bytes: vbaSourceSize });
  } catch (error) {
    logDebug("vba_source_preview_error", getErrorDetails(error));
    setStatus(`Could not load VBA source preview: ${getErrorMessage(error)}`, true);
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
  if (pcodeLoading || pcodeHighlightingInProgress || pcodeOffset >= Math.min(pcodeSize, PCODE_PREVIEW_LIMIT)) return;
  pcodeLoading = true;
  pcodeMoreButton.disabled = true;
  const startedAt = performance.now();
  try {
    const bytes = await readPcodeBytes(pcodeOffset, PCODE_CHUNK_SIZE);
    pcodeOffset += bytes.length;
    const previewComplete = pcodeOffset >= Math.min(pcodeSize, PCODE_PREVIEW_LIMIT);
    const chunk = pcodeDecoder.decode(bytes, { stream: !previewComplete });
    if (pcodeHighlighting) {
      pcodeHighlighter.append(chunk, previewComplete);
    } else {
      pcodeOutput.append(document.createTextNode(chunk));
    }
    pcodeProgress.textContent = `Preview: ${pcodeOffset.toLocaleString()} of ${pcodeSize.toLocaleString()} bytes`;
    pcodeMoreButton.hidden = pcodeOffset >= Math.min(pcodeSize, PCODE_PREVIEW_LIMIT);
    logDebug("pcode_preview_chunk", { bytes: bytes.length, total_bytes: pcodeSize });
  } catch (error) {
    logDebug("pcode_preview_error", getErrorDetails(error));
    setStatus(`Could not load p-code preview: ${getErrorMessage(error)}`, true);
  } finally {
    pcodeMoreButton.disabled = pcodeHighlightingInProgress;
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
    setStatus(`Could not load XLM macro preview: ${getErrorMessage(error)}`, true);
  } finally {
    xlmMoreButton.disabled = false;
    xlmLoading = false;
  }
}

function openSourcePreview(sourceType: "vba" | "xlm" | "pcode") {
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
    resetVbaSyntaxHighlighting();
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
    startVbaSyntaxHighlighting();
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
    app: "VBA-Scan",
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

function setSelectedFile(file: File) {
  const extension = file.name.split(".").pop()?.toLowerCase() ?? "";
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

function formatFileSize(bytes: number) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function getPcodeFilename(fileName: string) {
  const extensionSeparator = fileName.lastIndexOf(".");
  if (extensionSeparator <= 0 || extensionSeparator === fileName.length - 1) {
    return `${fileName}-pcode.txt`;
  }
  return `${fileName.slice(0, extensionSeparator)}-${fileName.slice(extensionSeparator + 1)}-pcode.txt`;
}

function getVbaSourceArchiveFilename(fileName: string) {
  const extensionSeparator = fileName.lastIndexOf(".");
  const stem = extensionSeparator > 0 ? fileName.slice(0, extensionSeparator) : fileName;
  return `${stem}-vba-source.zip`;
}

function getVbaSourceTextFilename(fileName: string) {
  const extensionSeparator = fileName.lastIndexOf(".");
  const stem = extensionSeparator > 0 ? fileName.slice(0, extensionSeparator) : fileName;
  return `${stem}-vba-source.txt`;
}

function getXlmFilename(fileName: string) {
  const extensionSeparator = fileName.lastIndexOf(".");
  const stem = extensionSeparator > 0 ? fileName.slice(0, extensionSeparator) : fileName;
  return `${stem}-xlm-macros.txt`;
}

async function downloadVbaSource(
  path: string,
  size: number,
  filename: string,
  mimeType: string,
  label: string,
  eventPrefix: string,
): Promise<void> {
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
    const blob = new Blob(chunks.map((chunk) => chunk.slice().buffer as ArrayBuffer), { type: mimeType });
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
    setStatus(`Could not download VBA source: ${getErrorMessage(error)}`, true);
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
    const blob = new Blob(chunks.map((chunk) => chunk.slice().buffer as ArrayBuffer), { type: "text/plain;charset=utf-8" });
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
    setStatus(`Could not download XLM macros: ${getErrorMessage(error)}`, true);
  } finally {
    xlmDownloadButton.disabled = false;
    removeFileButton.disabled = false;
    xlmDownloadInProgress = false;
  }
}

function setBusy(isBusy: boolean, message = "") {
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

function renderResults(result: ScanResult, elapsedMilliseconds: number, sha256: string, fileName: string) {
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
  resetVbaSyntaxHighlighting();
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
  resetPcodeSyntaxHighlighting();
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
  const hasSuspiciousIndicators = result.has_xlm || result.stomping === true
    || vbaSourceFindings.length > 0
    || pcodeFindings.length > 0
    || xlmFindings.length > 0;
  findingsSection.hidden = !hasSuspiciousIndicators;
  noSuspiciousIndicators.hidden = hasSuspiciousIndicators;
  vbaSourceWarningIcon.hidden = vbaSourceFindings.length === 0;
  xlmWarningIcon.hidden = xlmFindings.length === 0;
  const findingSignature = ({ type, keyword }: Finding) => `${type}\u0000${keyword.trim().toLowerCase()}`;
  const sourceLocationSignature = (finding: Finding) => finding.type === "Hex String"
    ? `${findingSignature(finding)}\u0000${finding.description.trim().toLowerCase()}`
    : findingSignature(finding);
  const vbaSourceLocationsBySignature = new Map<string, VbaSourceLocation[]>();
  result.vba_source_locations.forEach((location) => {
    const signature = sourceLocationSignature(location);
    const locations = vbaSourceLocationsBySignature.get(signature);
    if (locations) locations.push(location);
    else vbaSourceLocationsBySignature.set(signature, [location]);
  });
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
  const renderFindingGroup = (
    listElement: HTMLElement,
    countElement: HTMLElement,
    findings: Finding[],
    emptyMessage: string,
    options: FindingGroupOptions = {},
  ) => {
    const {
      matchingFindings = null,
      unmatchedFindingSignatures = null,
      sourceContextType = null,
    } = options;
    const hexStringFindings = findings.filter((finding) => finding.type === "Hex String");
    const hexSummaryFinding = findings.find((finding) =>
      finding.type === "Suspicious" && finding.keyword.trim().toLowerCase() === "hex strings",
    );
    const regularFindings = findings.filter((finding) =>
      finding.type !== "Hex String"
      && !(finding.type === "Suspicious" && finding.keyword.trim().toLowerCase() === "hex strings"),
    );
    const hasHexGroup = hexStringFindings.length > 0 || hexSummaryFinding !== undefined;
    countElement.textContent = `${regularFindings.length + Number(hasHexGroup)}`;
    if (regularFindings.length === 0 && !hasHexGroup) {
      const empty = document.createElement("div");
      empty.className = "no-findings";
      empty.textContent = emptyMessage;
      listElement.replaceChildren(empty);
    } else {
      const rows = regularFindings.map((finding) => {
      const row = document.createElement("div");
      row.className = "finding-row";
      const type = document.createElement("div");
      type.className = "finding-type";
      type.textContent = finding.type;
      const content = document.createElement("div");
      content.className = "finding-content";
      const keyword = document.createElement("div");
      keyword.className = "finding-keyword";
      const isBase64Finding = finding.type === "Base64 String";
      keyword.textContent = isBase64Finding
        ? `"${finding.description}" → "${finding.keyword}"`
        : finding.keyword;
      const description = document.createElement("div");
      description.className = "finding-description";
      const sourceLocations = sourceContextType === "vba"
        ? vbaSourceLocationsBySignature.get(findingSignature(finding)) || []
        : [];
      const locationCount = finding.keyword.trim().toLowerCase() === "lib" ? sourceLocations.length : 0;
      description.textContent = isBase64Finding
        ? "Potential Base64 string that decodes to printable characters"
        : `${finding.description.replace(" (option --decode to see all)", "")}${
          locationCount > 1 ? ` (${locationCount} source locations)` : ""
        }`;
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
              renderHighlightedCode(sourceCode, sourceContext.text, "pcode");
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
        sourceLocations.forEach((location) => content.append(createVbaSourceDetails(location)));
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
      });
      if (hasHexGroup) {
        const row = document.createElement("div");
        row.className = "finding-row hex-finding-row";
        const type = document.createElement("div");
        type.className = "finding-type";
        type.textContent = "Hex strings";
        const content = document.createElement("div");
        content.className = "finding-content";
        const keyword = document.createElement("div");
        keyword.className = "finding-keyword";
        keyword.textContent = hexStringFindings.length
          ? `${hexStringFindings.length} encoded ${hexStringFindings.length === 1 ? "string" : "strings"} detected`
          : "Hex-encoded strings detected";
        const description = document.createElement("div");
        description.className = "finding-description";
        description.textContent = hexSummaryFinding?.description.replace(" (option --decode to see all)", "")
          || "Review the encoded values and their decoded text.";
        content.append(keyword, description);
        if (hexStringFindings.length > 0) {
          const details = document.createElement("details");
          details.className = "hex-string-details";
          const summary = document.createElement("summary");
          summary.textContent = "Show encoded and decoded strings";
          const list = document.createElement("div");
          list.className = "hex-string-list";
          hexStringFindings.forEach((finding) => {
            const item = document.createElement("div");
            item.className = "hex-string-item";
            for (const [label, value] of [["Encoded", finding.description], ["Decoded", finding.keyword]]) {
              const valueRow = document.createElement("div");
              valueRow.className = "hex-string-value";
              const valueLabel = document.createElement("span");
              valueLabel.textContent = label;
              const code = document.createElement("code");
              code.textContent = label === "Decoded"
                ? value.replace(/[\x00-\x1f\x7f-\x9f]/g, (character) =>
                  `\\x${character.charCodeAt(0).toString(16).padStart(2, "0")}`)
                : value;
              valueRow.append(valueLabel, code);
              item.append(valueRow);
            }
            if (sourceContextType === "vba") {
              const locations = vbaSourceLocationsBySignature.get(sourceLocationSignature(finding)) || [];
              if (locations.length) {
                locations.forEach((location) => item.append(createVbaSourceDetails(location)));
              }
              else {
                const note = document.createElement("div");
                note.className = "finding-description";
                note.textContent = "No matching VBA source line; this value may come from embedded form data.";
                item.append(note);
              }
            }
            list.append(item);
          });
          details.append(summary, list);
          content.append(details);
        } else {
          const note = document.createElement("div");
          note.className = "finding-description";
          note.textContent = "No printable decoded strings were available to list.";
          content.append(note);
        }
        row.append(type, content);
        rows.push(row);
      }
      listElement.replaceChildren(...rows);
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
    const scan = await requestWorker<ScanWorkerResult>("scan", { fileBytes, fileName: file.name }, [fileBytes]);
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
    setStatus(`Analysis could not be completed: ${getErrorMessage(error)}`, true);
  } finally {
    setBusy(false);
  }
}

browseButton.addEventListener("click", () => fileInput.click());
dropZone.addEventListener("click", (event) => {
  if (event.target instanceof Element && event.target.closest("button")) return;
  fileInput.click();
});
fileInput.addEventListener("change", () => {
  if (fileInput.files?.[0]) setSelectedFile(fileInput.files[0]);
});
removeFileButton.addEventListener("click", clearSelectedFile);
openVbaSourceButton.addEventListener("click", () => openSourcePreview("vba"));
openXlmSourceButton.addEventListener("click", () => openSourcePreview("xlm"));
openPcodeSourceButton.addEventListener("click", () => openSourcePreview("pcode"));
pcodeHighlightButton.addEventListener("click", enablePcodeSyntaxHighlighting);
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
    const blob = new Blob(chunks.map((chunk) => chunk.slice().buffer as ArrayBuffer), { type: "text/plain;charset=utf-8" });
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
    setStatus(`Could not download p-code: ${getErrorMessage(error)}`, true);
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
    window.clearTimeout(copyFeedbackTimeout ?? undefined);
    copyFeedbackTimeout = window.setTimeout(() => {
      copyHashButton.classList.remove("copied");
      copyHashButton.setAttribute("aria-label", "Copy SHA-256");
      copyHashButton.title = "Copy SHA-256";
      copyStatus.textContent = "";
    }, 1500);
  } catch (error) {
    setStatus(`Could not copy SHA-256: ${getErrorMessage(error)}`, true);
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
