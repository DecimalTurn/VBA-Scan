# VBA-Scan

A static, browser-based Office macro analysis tool powered by Python `oletools` running in WebAssembly. Documents are processed locally in the browser; this project does not include a server or upload endpoint.

The scanner reports VBA and Excel 4.0 macro presence, VBA stomping checks and suspicious keyword indicators.

## How to use

Simply visit vba-scan.decimalturn.dev and submit a file for analysis.

## How it works

On the first scan, the page downloads Pyodide and the dependencies required by `oletools`. The `oletools` and `pcodedmp` sources are bundled from their submodules as `dist/oletools.zip`, so the page runs the checked-out versions, including local customizations. Scanned documents are not uploaded. An internet connection is required to initialize the analysis engine and install its dependencies.

After changing either package submodule, rebuild the static package asset:

```sh
python3 build-oletools-asset.py
```

Commit the regenerated `dist/oletools.zip` with the submodule update.

P-code downloads are named `<filename>-<file-ext>-pcode.txt` (for example,
`invoice-xlsm-pcode.txt`). The standalone extractor writes the same name into
its output directory. Browser downloads use the standard download mechanism;
files over 10 MiB display a warning because buffering them may use substantial
memory.

Extracted VBA source is shown inline, with large sources limited to a preview,
and can be downloaded as `<filename>-vba-source.zip` containing the individual
module files.

## Publish on GitHub Pages

The GitHub Actions workflow publishes the site on pushes to `main` and can also
be run manually from the Actions tab. To enable deployment, set the repository's
Pages build and deployment source to **GitHub Actions** in Settings → Pages.

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

