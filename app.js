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
const runtimeIndicator = document.querySelector(".runtime-dot");
const statusMessage = document.querySelector("#status-message");
const statusText = document.querySelector("#status-text");
const analysisSpinner = document.querySelector("#analysis-spinner");
const debugLogButton = document.querySelector("#download-debug-log");
const resultsElement = document.querySelector("#results");
const resultSummary = document.querySelector("#result-summary");
const vbaSourceSection = document.querySelector("#vba-source-section");
const vbaSourceOutput = document.querySelector("#vba-source-output");
const vbaSourceControls = document.querySelector("#vba-source-controls");
const vbaSourceProgress = document.querySelector("#vba-source-progress");
const vbaSourceMoreButton = document.querySelector("#vba-source-more");
const vbaSourceDownloadButton = document.querySelector("#vba-source-download");
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
let pyodidePromise = null;
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

const analyzePython = `
import json
import os
import re
import zipfile
from oletools.olevba import VBA_Parser, VBA_Project

def get_project_codepages(parser):
    codepages = []
    parsers = [parser]
    while parsers:
        current = parsers.pop()
        parsers.extend(getattr(current, "ole_subfiles", []) or [])
        if current.ole_file is None:
            continue
        try:
            projects = current.find_vba_projects() or []
        except Exception:
            continue
        for vba_root, project_path, dir_path in projects:
            try:
                project = VBA_Project(
                    current.ole_file, vba_root, project_path, dir_path, current.relaxed
                )
                codepages.append({
                    "project": str(project.projectname),
                    "value": project.codepage,
                    "name": project.codepage_name,
                })
            except Exception:
                continue
    return codepages

def scan_document(path, pcode_path, vba_source_path, vba_archive_path, xlm_path):
    parser = VBA_Parser(path)
    try:
        has_vba = parser.detect_vba_macros()
        has_xlm = parser.detect_xlm_macros()
        stomping = None
        stomping_error = None
        stomping_items = []
        try:
            stomping = parser.detect_vba_stomping()
        except Exception as error:
            stomping_error = str(error)
        if stomping:
          stomping_items = getattr(parser, "vba_stomping_items", [])

        pcode_size = 0
        pcode_error = None
        vba_source_size = 0
        vba_source_archive_size = 0
        vba_source_module_count = 0
        vba_source_error = None
        xlm_size = 0
        xlm_error = None
        xlm_sources = []
        xlm_findings = []
        if has_vba or stomping:
            try:
                pcode = parser.extract_pcode() or ""
                if pcode:
                    with open(pcode_path, "w", encoding="utf-8") as pcode_file:
                        pcode_file.write(pcode)
                    pcode_size = os.path.getsize(pcode_path)
            except Exception as error:
                pcode_error = str(error)

        if has_vba:
            try:
                used_names = set()
                source_parts = []
                with zipfile.ZipFile(vba_archive_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
                    for index, (_, stream_path, vba_filename, vba_code) in enumerate(parser.extract_macros(), start=1):
                        if stream_path == "VBA P-code" and vba_filename == "VBA_P-code.txt":
                            continue
                        if stream_path == "xlm_macro" and vba_filename == "xlm_macro.txt":
                            continue
                        filename = str(vba_filename).replace("\\\\", "/").rsplit("/", 1)[-1]
                        if filename in ("", ".", ".."):
                            filename = f"Module-{index}.bas"
                        filename = re.sub(r'[<>:"/|?*\\x00-\\x1f]', "_", filename)
                        stem, extension = os.path.splitext(filename)
                        candidate = filename
                        suffix = 2
                        while candidate.casefold() in used_names:
                            candidate = f"{stem}-{suffix}{extension}"
                            suffix += 1
                        used_names.add(candidate.casefold())
                        archive.writestr(candidate, vba_code)
                        source_parts.append(f"' ===== {candidate} =====\\n{vba_code}\\n")
                        vba_source_module_count += 1
                if vba_source_module_count:
                    with open(vba_source_path, "w", encoding="utf-8", newline="") as source_file:
                        source_file.write("\\n".join(source_parts))
                    vba_source_size = os.path.getsize(vba_source_path)
                    vba_source_archive_size = os.path.getsize(vba_archive_path)
                else:
                    os.unlink(vba_archive_path)
            except Exception as error:
                vba_source_error = str(error)
                vba_source_module_count = 0
                if os.path.exists(vba_archive_path):
                    os.unlink(vba_archive_path)
                if os.path.exists(vba_source_path):
                    os.unlink(vba_source_path)

        if has_xlm:
          try:
            xlm_sources = parser.extract_xlm_macros() or []
            xlm_findings = [
              {"type": kind, "keyword": keyword, "description": description}
              for kind, keyword, description in parser.analyze_xlm_macros()
            ]
            if xlm_sources:
              with open(xlm_path, "w", encoding="utf-8", newline="") as xlm_file:
                for index, (location, code) in enumerate(xlm_sources):
                  if index > 0:
                    xlm_file.write("\\n")
                  xlm_file.write(f"===== {location} =====\\n")
                  xlm_file.write(code.rstrip() + "\\n")
              xlm_size = os.path.getsize(xlm_path)
          except Exception as error:
            xlm_error = str(error)

        findings_by_source = parser.analyze_macros_by_source() or {}
        vba_source_findings = [
          {"type": kind, "keyword": keyword, "description": description}
          for kind, keyword, description in findings_by_source.get("vba_source", [])
        ]
        pcode_findings = [
          {"type": kind, "keyword": keyword, "description": description}
          for kind, keyword, description in findings_by_source.get("pcode", [])
        ]
        return json.dumps({
            "has_vba": has_vba,
            "has_xlm": has_xlm,
            "stomping": stomping,
            "stomping_error": stomping_error,
            "stomping_items": stomping_items,
            "pcode_size": pcode_size,
            "pcode_error": pcode_error,
            "vba_source_size": vba_source_size,
            "vba_source_archive_size": vba_source_archive_size,
            "vba_source_module_count": vba_source_module_count,
            "vba_source_error": vba_source_error,
            "xlm_size": xlm_size,
            "xlm_error": xlm_error,
            "xlm_macro_count": len(xlm_sources),
            "project_codepages": get_project_codepages(parser),
            "vba_source_findings": vba_source_findings,
            "pcode_findings": pcode_findings,
            "xlm_findings": xlm_findings,
        })
    finally:
        parser.close()

`;

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
  analysisSpinner.hidden = !scanning || isError;
  statusMessage.hidden = !message && analysisSpinner.hidden;
  statusMessage.classList.toggle("error", isError);
}

function logDuration(label, startedAt) {
  const duration = Number((performance.now() - startedAt).toFixed(1));
  console.info(`[VBA Scan] ${label}: ${duration} ms`);
  logDebug("timing", { label, duration_ms: duration });
}

async function readPcodeBytes(offset, size) {
  return readTemporaryFileBytes(pcodePath, offset, size);
}

async function readTemporaryFileBytes(path, offset, size) {
  const pyodide = await pyodidePromise;
  const stream = pyodide.FS.open(path, "r");
  const buffer = new Uint8Array(size);
  try {
    const bytesRead = pyodide.FS.read(stream, buffer, 0, size, offset);
    return buffer.subarray(0, bytesRead);
  } finally {
    pyodide.FS.close(stream);
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
  fileInput.value = "";
  const temporaryPaths = [pcodePath, vbaSourcePath, vbaSourceArchivePath, xlmPath].filter(Boolean);
  if (temporaryPaths.length && pyodidePromise) {
    pyodidePromise.then((pyodide) => {
      for (const path of temporaryPaths) {
        try { pyodide.FS.unlink(path); } catch {}
      }
    });
  }
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

function getXlmFilename(fileName) {
  const extensionSeparator = fileName.lastIndexOf(".");
  const stem = extensionSeparator > 0 ? fileName.slice(0, extensionSeparator) : fileName;
  return `${stem}-xlm-macros.txt`;
}

async function downloadVbaSourceArchive() {
  if (vbaSourceDownloadInProgress || !vbaSourceArchiveSize || !selectedFile) return;
  const filename = getVbaSourceArchiveFilename(selectedFile.name);
  vbaSourceDownloadInProgress = true;
  vbaSourceDownloadButton.disabled = true;
  removeFileButton.disabled = true;
  const startedAt = performance.now();
  let written = 0;
  logDebug("vba_source_download_started", { bytes: vbaSourceArchiveSize, download_method: "blob" });
  try {
    vbaSourceProgress.textContent = "Preparing VBA source archive…";
    const chunks = [];
    for (let offset = 0; offset < vbaSourceArchiveSize; offset += VBA_SOURCE_DOWNLOAD_CHUNK_SIZE) {
      const requestedBytes = Math.min(VBA_SOURCE_DOWNLOAD_CHUNK_SIZE, vbaSourceArchiveSize - offset);
      const bytes = await readTemporaryFileBytes(vbaSourceArchivePath, offset, requestedBytes);
      if (bytes.length === 0) throw new Error("The VBA source archive ended before the expected size.");
      chunks.push(bytes);
      written += bytes.length;
      vbaSourceProgress.textContent =
        `Downloading ${written.toLocaleString()} of ${vbaSourceArchiveSize.toLocaleString()} bytes`;
      if (written % (VBA_SOURCE_DOWNLOAD_CHUNK_SIZE * 16) === 0) {
        logDebug("vba_source_download_progress", { bytes: written, total_bytes: vbaSourceArchiveSize });
      }
    }
    const blob = new Blob(chunks, { type: "application/zip" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    link.click();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
    vbaSourceProgress.textContent = `Download complete (${written.toLocaleString()} bytes)`;
    logDebug("vba_source_download_complete", { bytes: written });
  } catch (error) {
    logDebug("vba_source_download_error", getErrorDetails(error));
    vbaSourceProgress.textContent = "Download did not complete";
    setStatus(`Could not download VBA source: ${error.message}`, true);
  } finally {
    vbaSourceDownloadButton.disabled = false;
    removeFileButton.disabled = false;
    vbaSourceDownloadInProgress = false;
    logDuration("VBA source archive download", startedAt);
  }
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
  if (!isBusy) analysisSpinner.hidden = true;
  removeFileButton.disabled = isBusy;
  runtimeNote.textContent = message || {
    loading: "Preparing analysis engine in background…",
    ready: "Analysis engine ready",
    error: "Engine unavailable; remove the file and upload again to retry",
  }[engineState];
}

async function loadPyodideEngine() {
  if (!pyodidePromise) {
    setEngineState("loading");
    pyodidePromise = (async () => {
      const engineStartedAt = performance.now();
      const script = document.createElement("script");
      script.src = `https://cdn.jsdelivr.net/pyodide/v${PYODIDE_VERSION}/full/pyodide.js`;
      script.async = true;
      const scriptStartedAt = performance.now();
      document.head.append(script);
      await new Promise((resolve, reject) => {
        script.addEventListener("load", () => {
          logDuration("Pyodide loader download and evaluation", scriptStartedAt);
          resolve();
        }, { once: true });
        script.addEventListener("error", () => {
          logDuration("Pyodide loader download failed", scriptStartedAt);
          reject(new Error("Could not load the WebAssembly runtime. Check your internet connection and try again."));
        }, { once: true });
      });

      const runtimeStartedAt = performance.now();
      const pyodide = await globalThis.loadPyodide({
        indexURL: `https://cdn.jsdelivr.net/pyodide/v${PYODIDE_VERSION}/full/`,
      });
      logDuration("Pyodide runtime startup (includes WASM asset downloads)", runtimeStartedAt);

      const micropipStartedAt = performance.now();
      await pyodide.loadPackage("micropip");
      logDuration("micropip package load", micropipStartedAt);

      const oletoolsStartedAt = performance.now();
      const oletoolsResponse = await fetch(new URL("dist/oletools.zip", import.meta.url), { cache: "no-store" });
      if (!oletoolsResponse.ok) {
        throw new Error(`Could not load the bundled analysis packages (${oletoolsResponse.status}).`);
      }
      pyodide.FS.writeFile("/tmp/oletools.zip", new Uint8Array(await oletoolsResponse.arrayBuffer()));
      await pyodide.runPythonAsync(
        `import micropip
await micropip.install([
    "pyparsing>=2.1.0,<4",
    "olefile>=0.46",
    "easygui",
    "colorclass",
    "msoffcrypto-tool",
])
import sys
sys.path.insert(0, "/tmp/oletools.zip")
import oletools.olevba`,
      );
      logDuration("Vendored oletools and dependencies setup", oletoolsStartedAt);

      const analyzerStartedAt = performance.now();
      await pyodide.runPythonAsync(analyzePython);
      logDuration("Python analyzer setup", analyzerStartedAt);
      logDuration("Total analysis engine setup", engineStartedAt);
      return pyodide;
    })();

    pyodidePromise.then(() => {
      setEngineState("ready");
    }, (error) => {
      setEngineState("error");
      pyodidePromise = null;
      console.error("[VBA Scan] Analysis engine setup failed.", error);
    });
  }
  return pyodidePromise;
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

  vbaSourceSection.hidden = !result.has_vba;
  vbaSourceOffset = 0;
  vbaSourceSize = result.vba_source_size || 0;
  vbaSourceModuleCount = result.vba_source_module_count || 0;
  vbaSourceArchiveSize = result.vba_source_archive_size || 0;
  vbaSourceLargeDownloadWarning.hidden =
    vbaSourceArchiveSize <= VBA_SOURCE_LARGE_DOWNLOAD_WARNING_BYTES;
  vbaSourceDecoder = new TextDecoder("utf-8");
  vbaSourceOutput.replaceChildren();
  vbaSourceOutput.classList.toggle("code-unavailable", !vbaSourceSize);
  if (vbaSourceSize) {
    vbaSourceControls.hidden = false;
    vbaSourceDownloadButton.hidden = false;
    vbaSourceProgress.textContent =
      `Source ready (${vbaSourceModuleCount.toLocaleString()} modules, ${vbaSourceSize.toLocaleString()} bytes)`;
    vbaSourceMoreButton.hidden = false;
  } else {
    vbaSourceControls.hidden = true;
    vbaSourceDownloadButton.hidden = true;
    vbaSourceOutput.textContent = result.vba_source_error || "VBA source code was not available.";
  }

  xlmFindingsGroup.hidden = !result.has_xlm;
  xlmFindingsGroup.open = true;
  xlmSection.hidden = !result.has_xlm;
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

  pcodeSection.hidden = !result.has_vba && !result.stomping;
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
  const pcodeFindingsAreInVbaSource = [...pcodeFindingSignatures]
    .every((signature) => vbaSourceFindingSignatures.has(signature));
  const pcodeGroupShouldStartCollapsed = result.stomping === false && pcodeFindingsAreInVbaSource;
  vbaSourceFindingsGroup.open = true;
  pcodeFindingsGroup.open = !pcodeGroupShouldStartCollapsed;
  pcodeWarningIcon.hidden = !result.stomping && pcodeFindingsAreInVbaSource;
  pcodeMatchNote.hidden = !(result.stomping === false && pcodeFindings.length > 0 && pcodeFindingsAreInVbaSource);
  const totalFindings = vbaSourceFindings.length + pcodeFindings.length + xlmFindings.length;
  findingCount.textContent = `${totalFindings} ${totalFindings === 1 ? "indicator" : "indicators"}`;
  const renderFindingGroup = (listElement, countElement, findings, emptyMessage, matchingFindings = null) => {
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
  );
  renderFindingGroup(
    pcodeFindingsList,
    pcodeFindingCount,
    pcodeFindings,
    result.pcode_size ? "No suspicious indicators found in P-code." : "P-code was not available.",
    result.stomping === false ? vbaSourceFindingsBySignature : null,
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

  setBusy(true, engineState === "ready" ? "Preparing document for analysis…" : "Loading WebAssembly analysis engine…");
  setStatus("Initializing Python and oletools. The first run may take a little longer.");
  try {
    const pyodide = await loadPyodideEngine();
    if (!selectedFile) return;

    for (const temporaryPath of [pcodePath, vbaSourcePath, vbaSourceArchivePath, xlmPath]) {
      if (temporaryPath) {
        try { pyodide.FS.unlink(temporaryPath); } catch {}
      }
    }
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

    setBusy(true, "Analyzing locally in your browser…");
    setStatus();
    const file = selectedFile;
    const startedAt = performance.now();
    if (!globalThis.crypto?.subtle) {
      throw new Error("SHA-256 hashing is unavailable in this browser context.");
    }
    const fileBytes = await file.arrayBuffer();
    const digest = await globalThis.crypto.subtle.digest("SHA-256", fileBytes);
    const sha256 = Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, "0")).join("");
    const path = `/tmp/vba-scan-${Date.now()}-${file.name.replace(/[^a-zA-Z0-9._-]/g, "_")}`;
    pcodePath = `/tmp/vba-scan-pcode-${Date.now()}.txt`;
    vbaSourcePath = `/tmp/vba-scan-source-${Date.now()}.txt`;
    vbaSourceArchivePath = `/tmp/vba-scan-source-${Date.now()}.zip`;
    xlmPath = `/tmp/vba-scan-xlm-${Date.now()}.txt`;
    pyodide.FS.writeFile(path, new Uint8Array(fileBytes));
    logDebug("analysis_started", { input_bytes: fileBytes.byteLength });
    try {
      const scanStartedAt = performance.now();
      const rawResult = await pyodide.runPythonAsync(
        `scan_document(${JSON.stringify(path)}, ${JSON.stringify(pcodePath)}, ${JSON.stringify(vbaSourcePath)}, ${JSON.stringify(vbaSourceArchivePath)}, ${JSON.stringify(xlmPath)})`,
      );
      logDuration("Document analysis", scanStartedAt);
      renderResults(JSON.parse(rawResult), performance.now() - startedAt, sha256, file.name);
    } finally {
      pyodide.FS.unlink(path);
    }
    setStatus("Analysis complete.");
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
vbaSourceMoreButton.addEventListener("click", appendVbaSourceChunk);
vbaSourceSection.addEventListener("toggle", () => {
  if (vbaSourceSection.open && vbaSourceOffset === 0 && vbaSourceSize) {
    if (vbaSourceSize <= VBA_SOURCE_PREVIEW_LIMIT) {
      void loadVbaSourcePreview();
    } else {
      void appendVbaSourceChunk();
    }
  }
});
vbaSourceDownloadButton.addEventListener("click", downloadVbaSourceArchive);
xlmMoreButton.addEventListener("click", appendXlmChunk);
xlmSection.addEventListener("toggle", () => {
  if (xlmSection.open && xlmOffset === 0 && xlmSize) void appendXlmChunk();
});
xlmDownloadButton.addEventListener("click", downloadXlmMacros);
pcodeMoreButton.addEventListener("click", appendPcodeChunk);
pcodeSection.addEventListener("toggle", () => {
  if (pcodeSection.open && pcodeOffset === 0 && pcodeSize) void appendPcodeChunk();
});
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

loadPyodideEngine().then(() => {
  if (!scanning) setBusy(false);
}).catch((error) => {
  if (!scanning) {
    setBusy(false);
    setStatus(`Analysis engine could not be loaded: ${error.message}`, true);
  }
});
