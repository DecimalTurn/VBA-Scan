# Python Setup

The devcontainer uses Python 3.12. Keep project dependencies in the repository-root `.venv`; `.gitignore` excludes it, and `.vscode/settings.json` points VS Code at its interpreter.

After changing the devcontainer image, rebuild the Codespace before creating the environment. An existing Codespace keeps its original system Python until it is rebuilt.

## First-Time Setup

Run these commands from the repository root:

```sh
# --recursive might be needed if there are nested submodules, otherwise keep it simple:
git submodule update --init
python3 --version
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r submodules/oletools/requirements.txt
```

The version check should report Python 3.12. If VS Code does not select the environment automatically, run **Python: Select Interpreter** and choose `.venv/bin/python`.

## Running Commands

Activate the environment first:

```sh
source .venv/bin/activate
```

Then Python scripts use the virtualenv's `python`:

```sh
python build-oletools-asset.py
python -m http.server 8000
```

The VBA extraction script also accepts an explicit interpreter, without activating the environment:

```sh
PYTHON="$PWD/.venv/bin/python" ./extract-vba.sh --force samples/BuiltInStringFunctions.xlsm
```

Run the oletools tests from its submodule directory:

```sh
cd submodules/oletools
../../.venv/bin/python -m unittest discover -v
```