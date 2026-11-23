# Development Cycle

## 1. Set up the environment

Use Node.js 26 for the frontend build:

```sh
nvm use 26
```

Install the root project dependencies using the package manager and lockfile already in the repository.

## 2. Main code locations

Start from the failing behavior and identify which layer owns it:

- `app.ts` and `styles.css` contain the browser UI.
- `analysis-python.ts` connects the browser analysis flow to Python.
- `submodules/oletools` is a separate Git repository containing the VBA analysis code.

Keep a change within its owning layer where possible. If oletools behavior changes, update or add a focused test alongside the implementation.

### VBA parsing and detection safety

The scanner analyzes potentially malicious VBA, so syntax-related blind spots can hide real threats. Do not assume commands use conventional spacing or formatting: VBA permits compact syntax such as a colon statement separator with no surrounding spaces (`Then:Open`). When changing parsing or detection, check the relevant VBA syntax and add regression tests for valid syntax variants as well as nearby non-command cases, so fixes do not introduce false negatives. When in doubt, it's better to report more threats than to miss potential ones.

## 3. Validate Python analysis changes

From the oletools checkout, run the full test suite:

```sh
cd submodules/oletools
PYTHONPATH=. python -m unittest discover -s tests
```

For a narrow change, run the relevant test first, then the full suite when practical. If oletools source is packaged for browser use, regenerate the packaged archive as part of the change so the browser does not keep using stale code.

## 4. Build and check the frontend

From the repository root:

```sh
npm run build
```

This runs the TypeScript check, bundles the app and analysis worker, and copies the static files into the build output. Review the app in a browser after changes to user-facing behavior, including the relevant sample file or workflow.

Before committing, check whitespace and inspect the changed files:

```sh
git diff --check
git status --short
```

## 5. Commit changes

Use conventional commit messages that describe the behavior change. When changing the nested oletools repository:

1. Commit the implementation and tests inside `submodules/oletools`.
2. Commit the updated submodule reference in the root repository, together with any root-level changes.

Stage only files related to the task. Avoid committing unrelated or pre-existing work.
