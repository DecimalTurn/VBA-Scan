const PYODIDE_VERSION = "0.27.7";
const OLETOOLS_VERSION = "0.60.2";
const PCODE_CHUNK_SIZE = 16_384;
const PCODE_PREVIEW_LIMIT = 524_288;
const PCODE_DOWNLOAD_CHUNK_SIZE = 262_144;
const PCODE_BLOB_FALLBACK_LIMIT = 8 * 1024 * 1024;
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
const pcodeSection = document.querySelector("#pcode-section");
const pcodeOutput = document.querySelector("#pcode-output");
const pcodeControls = document.querySelector("#pcode-controls");
const pcodeProgress = document.querySelector("#pcode-progress");
const pcodeMoreButton = document.querySelector("#pcode-more");
const pcodeDownloadButton = document.querySelector("#pcode-download");
const findingsList = document.querySelector("#findings-list");
const findingCount = document.querySelector("#finding-count");
const analysisTime = document.querySelector("#analysis-time");
const documentHash = document.querySelector("#document-hash");
const documentFilename = document.querySelector("#document-filename");
const copyHashButton = document.querySelector("#copy-hash-button");
const copyStatus = document.querySelector("#copy-status");

let selectedFile = null;
let pyodidePromise = null;
let scanning = false;
let engineState = "loading";
let copyFeedbackTimeout = null;
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
from oletools.olevba import VBA_Parser

def scan_document(path, pcode_path):
    parser = VBA_Parser(path)
    try:
        has_vba = parser.detect_vba_macros()
        has_xlm = parser.detect_xlm_macros()
        stomping = None
        stomping_error = None
        try:
            stomping = parser.detect_vba_stomping()
        except Exception as error:
            stomping_error = str(error)

        pcode_size = 0
        pcode_error = None
        if stomping:
            try:
                pcode = parser.extract_pcode() or ""
                if pcode:
                    with open(pcode_path, "w", encoding="utf-8") as pcode_file:
                        pcode_file.write(pcode)
                    pcode_size = os.path.getsize(pcode_path)
            except Exception as error:
                pcode_error = str(error)

        findings = [
            {"type": kind, "keyword": keyword, "description": description}
            for kind, keyword, description in (parser.analyze_macros() or [])
        ]
        return json.dumps({
            "has_vba": has_vba,
            "has_xlm": has_xlm,
            "stomping": stomping,
            "stomping_error": stomping_error,
            "pcode_size": pcode_size,
            "pcode_error": pcode_error,
            "findings": findings,
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
  const pyodide = await pyodidePromise;
  const stream = pyodide.FS.open(pcodePath, "r");
  const buffer = new Uint8Array(size);
  try {
    const bytesRead = pyodide.FS.read(stream, buffer, 0, size, offset);
    return buffer.subarray(0, bytesRead);
  } finally {
    pyodide.FS.close(stream);
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
  if (pcodePath && pyodidePromise) {
    const path = pcodePath;
    pcodePath = "";
    pyodidePromise.then((pyodide) => {
      try { pyodide.FS.unlink(path); } catch {}
    });
  }
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
      await pyodide.runPythonAsync(
        `import micropip\nawait micropip.install("oletools==${OLETOOLS_VERSION}")`,
      );
      logDuration("oletools download and install", oletoolsStartedAt);

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
  const summaries = [];
  if (result.has_xlm) {
    summaries.push({ label: "⚠️ Excel 4.0 macros detected", value: "Review this document", tone: "danger" });
  }
  if (result.stomping) {
    summaries.push({
      label: "⚠️ VBA stomping detected",
      value: "Compiled VBA differs from source",
      tone: "danger",
      note: {
        text: "ℹ️ Note that oletools can report false positives for VBA stomping. See upstream reports:",
        issues: [489, 521, 533, 620],
      },
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
    return card;
  }));
  resultSummary.hidden = summaries.length === 0;

  pcodeSection.hidden = !result.stomping;
  pcodeOffset = 0;
  pcodeSize = result.pcode_size || 0;
  pcodeDecoder = new TextDecoder("utf-8");
  pcodeOutput.replaceChildren();
  pcodeOutput.classList.toggle("pcode-unavailable", !pcodeSize);
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
    pcode_bytes: pcodeSize,
    finding_count: result.findings.length,
  });

  const isStompingFinding = (finding) =>
    [finding.type, finding.keyword, finding.description].some(
      (value) => typeof value === "string" && value.toLowerCase().includes("stomp"),
    );
  const findings = result.findings.filter((finding) => !isStompingFinding(finding));

  findingCount.textContent = `${findings.length} ${findings.length === 1 ? "indicator" : "indicators"}`;
  if (findings.length === 0) {
    const empty = document.createElement("div");
    empty.className = "no-findings";
    if (!result.has_vba) {
      const indicator = document.createElement("span");
      indicator.className = "no-vba-indicator";
      empty.append(indicator, document.createTextNode("No VBA code found in this file."));
    } else {
      empty.textContent = "No suspicious keywords were reported by oletools.";
    }
    findingsList.replaceChildren(empty);
  } else {
    findingsList.replaceChildren(...findings.map((finding) => {
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
      description.textContent = finding.description;
      content.append(keyword, description);
      row.append(type, content);
      return row;
    }));
  }

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

    if (pcodePath) {
      try { pyodide.FS.unlink(pcodePath); } catch {}
      pcodePath = "";
    }
    pcodeSize = 0;

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
    pyodide.FS.writeFile(path, new Uint8Array(fileBytes));
    logDebug("analysis_started", { input_bytes: fileBytes.byteLength });
    try {
      const scanStartedAt = performance.now();
      const rawResult = await pyodide.runPythonAsync(`scan_document(${JSON.stringify(path)}, ${JSON.stringify(pcodePath)})`);
      logDuration("Document analysis", scanStartedAt);
      renderResults(JSON.parse(rawResult), performance.now() - startedAt, sha256, file.name);
    } finally {
      pyodide.FS.unlink(path);
    }
    setStatus("Analysis complete. Review any indicators in context.");
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
pcodeMoreButton.addEventListener("click", appendPcodeChunk);
pcodeSection.addEventListener("toggle", () => {
  if (pcodeSection.open && pcodeOffset === 0 && pcodeSize) void appendPcodeChunk();
});
pcodeDownloadButton.addEventListener("click", async () => {
  if (pcodeDownloadInProgress || !pcodeSize) return;
  pcodeDownloadInProgress = true;
  pcodeDownloadButton.disabled = true;
  removeFileButton.disabled = true;
  const startedAt = performance.now();
  let writable;
  logDebug("pcode_download_started", {
    bytes: pcodeSize,
    file_save_api_available: Boolean(window.showSaveFilePicker),
  });
  try {
    pcodeProgress.textContent = "Preparing p-code download…";
    if (window.showSaveFilePicker) {
      const handle = await window.showSaveFilePicker({
        suggestedName: "vba-pcode.txt",
        types: [{ description: "Text file", accept: { "text/plain": [".txt"] } }],
      });
      writable = await handle.createWritable();
    } else if (pcodeSize > PCODE_BLOB_FALLBACK_LIMIT) {
      throw new Error("This browser cannot safely download a large p-code file. Open the page in Chrome or Edge and try again.");
    }

    const chunks = writable ? null : [];
    let written = 0;
    for (let offset = 0; offset < pcodeSize; offset += PCODE_DOWNLOAD_CHUNK_SIZE) {
      const bytes = await readPcodeBytes(offset, Math.min(PCODE_DOWNLOAD_CHUNK_SIZE, pcodeSize - offset));
      if (bytes.length === 0) throw new Error("The p-code file ended before the expected size.");
      if (writable) await writable.write(bytes);
      else chunks.push(bytes);
      written += bytes.length;
      pcodeProgress.textContent = `Downloading ${written.toLocaleString()} of ${pcodeSize.toLocaleString()} bytes`;
      if (written % (PCODE_DOWNLOAD_CHUNK_SIZE * 16) === 0) {
        logDebug("pcode_download_progress", { bytes: written, total_bytes: pcodeSize });
      }
    }
    if (writable) {
      await writable.close();
    } else {
      const url = URL.createObjectURL(new Blob(chunks, { type: "text/plain;charset=utf-8" }));
      const link = document.createElement("a");
      link.href = url;
      link.download = "vba-pcode.txt";
      link.click();
      window.setTimeout(() => URL.revokeObjectURL(url), 1000);
    }
    pcodeProgress.textContent = `Download complete (${written.toLocaleString()} bytes)`;
    logDebug("pcode_download_complete", { bytes: written });
  } catch (error) {
    if (writable) {
      try { await writable.abort(); } catch {}
    }
    if (error.name === "AbortError") {
      pcodeProgress.textContent = "Download cancelled";
    } else {
      logDebug("pcode_download_error", getErrorDetails(error));
      const message = !writable && pcodeSize > PCODE_BLOB_FALLBACK_LIMIT && window.showSaveFilePicker
        ? `The browser could not grant file access. Open this page in Chrome or Edge. ${error.message}`
        : error.message;
      pcodeProgress.textContent = "Download did not complete";
      setStatus(`Could not download p-code: ${message}`, true);
    }
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
