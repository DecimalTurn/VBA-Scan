# VBA Scan

A static, browser-based Office macro triage page powered by Python `oletools` running in WebAssembly. Documents are processed locally in the browser; this project does not include a server or upload endpoint.

## Run locally

Serve this directory over HTTP (ES modules and WebAssembly need a web server):

```sh
python3 -m http.server 8000
```

Open <http://localhost:8000>. On the first scan, the page downloads Pyodide and the pinned `oletools` package from their public package hosts. Scanned documents are not uploaded. An internet connection is required to initialize the analysis engine.

The scanner reports VBA and Excel 4.0 macro presence, VBA stomping checks, and suspicious keyword indicators. These are triage signals, not a verdict that a document is safe or malicious.
