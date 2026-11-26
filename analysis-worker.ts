import { analyzePython } from "./analysis-python.ts";

const PYODIDE_VERSION = "0.27.7";
interface PyodideRuntime {
  FS: {
    close(stream: number): void;
    open(path: string, flags: string): number;
    read(stream: number, buffer: Uint8Array, offset: number, length: number, position: number): number;
    unlink(path: string): void;
    writeFile(path: string, data: Uint8Array): void;
  };
  loadPackage(name: string): Promise<void>;
  runPythonAsync(source: string): Promise<string>;
}

interface PyodideModule {
  loadPyodide(options: { indexURL: string }): Promise<PyodideRuntime>;
}

type WorkerRequest =
  | { type: "initialize" }
  | { type: "clear-files"; paths: string[] }
  | { type: "read-file"; requestId: number; path: string; offset: number; size: number }
  | { type: "scan"; requestId: number; fileBytes: ArrayBuffer; fileName: string };

const workerScope = self as DedicatedWorkerGlobalScope;
let pyodidePromise: Promise<PyodideRuntime> | null = null;
let pyodide: PyodideRuntime | null = null;
const temporaryPaths = new Set<string>();
let scanSequence = 0;

function sendEngineStatus(state: "loading" | "ready" | "error", message: string): void {
  workerScope.postMessage({ type: "engine-status", state, message });
}

function getErrorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

async function loadAnalysisEngine(): Promise<PyodideRuntime> {
  if (!pyodidePromise) {
    pyodidePromise = (async () => {
      sendEngineStatus("loading", "Loading Python runtime in background…");
      const { loadPyodide } = await import(
        `https://cdn.jsdelivr.net/pyodide/v${PYODIDE_VERSION}/full/pyodide.mjs`
      ) as unknown as PyodideModule;
      pyodide = await loadPyodide({
        indexURL: `https://cdn.jsdelivr.net/pyodide/v${PYODIDE_VERSION}/full/`,
      });

      sendEngineStatus("loading", "Preparing Python packages…");
      await pyodide.loadPackage("micropip");
      const oletoolsResponse = await fetch(new URL("./dist/oletools.zip", import.meta.url), {
        cache: "no-store",
      });
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
      await pyodide.runPythonAsync(analyzePython);
      sendEngineStatus("ready", "Analysis engine ready");
      return pyodide;
    })().catch((error) => {
      pyodidePromise = null;
      sendEngineStatus("error", getErrorMessage(error));
      throw error;
    });
  }
  return pyodidePromise;
}

function removeTemporaryPaths(paths: Iterable<string> = temporaryPaths): void {
  if (!pyodide) return;
  for (const path of paths) {
    if (!temporaryPaths.has(path)) continue;
    try { pyodide.FS.unlink(path); } catch {}
    temporaryPaths.delete(path);
  }
}

async function scanDocument(message: Extract<WorkerRequest, { type: "scan" }>): Promise<void> {
  const runtime = await loadAnalysisEngine();
  removeTemporaryPaths();

  const sequence = ++scanSequence;
  const safeName = message.fileName.replace(/[^a-zA-Z0-9._-]/g, "_");
  const inputPath = `/tmp/vba-scan-${sequence}-${safeName}`;
  const paths = {
    pcode: `/tmp/vba-scan-pcode-${sequence}.txt`,
    vbaSource: `/tmp/vba-scan-source-${sequence}.txt`,
    vbaArchive: `/tmp/vba-scan-source-${sequence}.zip`,
    xlm: `/tmp/vba-scan-xlm-${sequence}.txt`,
  };
  runtime.FS.writeFile(inputPath, new Uint8Array(message.fileBytes));
  temporaryPaths.add(inputPath);
  Object.values(paths).forEach((path) => temporaryPaths.add(path));
  workerScope.postMessage({ type: "scan-phase", message: "Analyzing document locally… Large files can take a while." });

  try {
    const result = await runtime.runPythonAsync(
      `scan_document(${JSON.stringify(inputPath)}, ${JSON.stringify(paths.pcode)}, ${JSON.stringify(paths.vbaSource)}, ${JSON.stringify(paths.vbaArchive)}, ${JSON.stringify(paths.xlm)})`,
    );
    workerScope.postMessage({
      type: "scan-complete",
      requestId: message.requestId,
      result,
      paths,
    });
  } finally {
    try { runtime.FS.unlink(inputPath); } catch {}
    temporaryPaths.delete(inputPath);
  }
}

function readTemporaryFile(message: Extract<WorkerRequest, { type: "read-file" }>): void {
  if (!temporaryPaths.has(message.path)) {
    throw new Error("The requested analysis output is no longer available.");
  }
  if (!pyodide) throw new Error("The analysis engine is not available.");
  const stream = pyodide.FS.open(message.path, "r");
  const buffer = new Uint8Array(message.size);
  try {
    const bytesRead = pyodide.FS.read(stream, buffer, 0, message.size, message.offset);
    const bytes = buffer.slice(0, bytesRead);
    workerScope.postMessage({
      type: "file-bytes",
      requestId: message.requestId,
      bytes: bytes.buffer as ArrayBuffer,
    }, [bytes.buffer as ArrayBuffer]);
  } finally {
    pyodide.FS.close(stream);
  }
}

async function handleMessage(data: WorkerRequest): Promise<void> {
  if (data.type === "initialize") {
    try {
      await loadAnalysisEngine();
    } catch (error) {
      console.error("[VBA-Scan] Analysis engine setup failed.", getErrorMessage(error));
    }
    return;
  }
  if (data.type === "clear-files") {
    removeTemporaryPaths(data.paths);
    return;
  }
  if (data.type === "read-file") {
    try {
      readTemporaryFile(data);
    } catch (error) {
      workerScope.postMessage({ type: "request-error", requestId: data.requestId, message: getErrorMessage(error) });
    }
    return;
  }
  if (data.type === "scan") {
    try {
      await scanDocument(data);
    } catch (error) {
      workerScope.postMessage({
        type: "request-error",
        requestId: data.requestId,
        message: getErrorMessage(error),
      });
    }
  }
}

workerScope.addEventListener("message", (event: MessageEvent<WorkerRequest>) => {
  void handleMessage(event.data);
});
