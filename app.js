const PYODIDE_VERSION = "0.27.7";
const OLETOOLS_VERSION = "0.60.2";

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
const resultsElement = document.querySelector("#results");
const resultSummary = document.querySelector("#result-summary");
const findingsList = document.querySelector("#findings-list");
const findingCount = document.querySelector("#finding-count");
const analysisTime = document.querySelector("#analysis-time");

let selectedFile = null;
let pyodidePromise = null;
let scanning = false;
let engineState = "loading";

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
from oletools.olevba import VBA_Parser

def scan_document(path):
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

        findings = [
            {"type": kind, "keyword": keyword, "description": description}
            for kind, keyword, description in (parser.analyze_macros() or [])
        ]
        return json.dumps({
            "has_vba": has_vba,
            "has_xlm": has_xlm,
            "stomping": stomping,
            "stomping_error": stomping_error,
            "findings": findings,
        })
    finally:
        parser.close()
`;

function setStatus(message = "", isError = false) {
  statusText.textContent = message;
  analysisSpinner.hidden = !scanning || isError;
  statusMessage.hidden = !message && analysisSpinner.hidden;
  statusMessage.classList.toggle("error", isError);
}

function logDuration(label, startedAt) {
  console.info(`[VBA Scan] ${label}: ${(performance.now() - startedAt).toFixed(1)} ms`);
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

function renderResults(result, elapsedMilliseconds) {
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

    setBusy(true, "Analyzing locally in your browser…");
    setStatus();
    const file = selectedFile;
    const path = `/tmp/vba-scan-${Date.now()}-${file.name.replace(/[^a-zA-Z0-9._-]/g, "_")}`;
    const startedAt = performance.now();
    pyodide.FS.writeFile(path, new Uint8Array(await file.arrayBuffer()));
    try {
      const rawResult = await pyodide.runPythonAsync(`scan_document(${JSON.stringify(path)})`);
      renderResults(JSON.parse(rawResult), performance.now() - startedAt);
    } finally {
      pyodide.FS.unlink(path);
    }
    setStatus("Analysis complete. Review any indicators in context.");
  } catch (error) {
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
