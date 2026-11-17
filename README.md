# VBA-Scan

A static, browser-based Office macro triage page powered by Python `oletools` running in WebAssembly. Documents are processed locally in the browser; this project does not include a server or upload endpoint.

The scanner reports VBA and Excel 4.0 macro presence, VBA stomping checks and suspicious keyword indicators.

On the first scan, the page downloads Pyodide and the pinned `oletools` package from their public package hosts. Scanned documents are not uploaded. An internet connection is required to initialize the analysis engine.

## Run locally

Serve this directory over HTTP (ES modules and WebAssembly need a web server):

```sh
python3 -m http.server 8000
```

Open <http://localhost:8000>. 

## Acknowledgements

- [oletools](https://github.com/decalage2/oletools) for providing the core analysis functionality.
- [Pyodide](https://github.com/pyodide/pyodide) for enabling Python to run in the browser via WebAssembly.
- [Senipah](https://gist.github.com/Senipah/cbfa43cf472fd03fbb322e00ae4fe7c7) for the project logo base design.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

The oletools package remains under its own license. See the [oletools repository](https://github.com/decalage2/oletools) for more information.

