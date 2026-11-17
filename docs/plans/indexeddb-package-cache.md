# Plan: Cache the Python analysis packages in IndexedDB

## Goal

Avoid downloading and reinstalling `oletools` and its Python dependencies on
every page load, while keeping document analysis entirely in the browser.
Persist only the analysis runtime's packages; never persist the selected
document or its analysis output.

## Current setup

[`app.js`](../../app.js) pins Pyodide `0.27.7` and `oletools` `0.60.2`. On page
load it initializes Pyodide, loads `micropip`, installs `oletools` and its
dependencies, and defines the Python analyzer. The existing console timings
separate these steps, although the runtime-startup timing combines WebAssembly
downloads with initialization.

The current implementation uses Pyodide's in-memory filesystem. A reload
creates a new Python runtime, so installed package files are not retained by
that filesystem. The browser may independently cache HTTP responses; that
cache is not a guarantee that package installation can be skipped.

## Proposed implementation

1. **Establish a repeatable baseline.** Record cold and warm page-load timings,
   transferred network bytes, and package storage size using the method below.
   Keep Pyodide and `oletools` versions fixed during the comparison.
2. **Prototype persistence with the pinned Pyodide version.** Pyodide exposes
   Emscripten's `IDBFS` filesystem. Verify that the exact packages and
   dependency artifacts installed by this project can be written to and
   restored from an IndexedDB-backed filesystem across a full page reload.
   Mount and restore the cache before importing packages that depend on it.
   Confirm that persistence does not hide or replace Pyodide's built-in
   packages, native extensions, or other runtime files.
3. **Choose the smallest safe cache layout.** Prefer persisting only the
   installed Python-tool package files and required dependency artifacts.
   Include a manifest with the cache schema, Pyodide version, Python version,
   `oletools` version, and exact resolved dependency versions. Key or validate
   the cache against all of these values so incompatible packages are never
   reused.
4. **Restore or install.** On startup, mount the cache, restore it with
   `FS.syncfs(true)`, validate the manifest, and skip package installation
   only when the cache is complete and compatible. On a miss or invalid cache,
   run the current installation path, then flush the package files and
   manifest to IndexedDB. Keep defining the analyzer in each new runtime;
   Python definitions and runtime state are not preserved by this package
   cache.
5. **Handle storage failure without breaking scanning.** If IndexedDB is
   unavailable, quota-limited, evicted, or corrupt, log a clear cache
   miss/bypass reason and use the existing installation path. Do not claim a
   cache hit until restore and package imports succeed. Keep the existing
   explicit error reporting for failures that prevent installation or
   analysis.
6. **Keep the cache scoped to engine assets.** The document is currently
   written under `/tmp` for analysis. Do not place that path, document bytes,
   or results in the persistent mount.
7. **Measure and document the result.** Add separate timings for cache restore,
   cache validation, package install on a miss, cache write, and total engine
   setup. Report whether a run was a cache hit or miss. Compare cold and warm
   runs before deciding whether the added persistence complexity is worthwhile.

If installed package files cannot safely be restored from `IDBFS` with the
pinned Pyodide version, test caching the wheel artifacts in IndexedDB and
installing from those local artifacts. That fallback may save network transfer,
but it should not be counted as saving package installation time. Do not
silently assume that Pyodide's ordinary HTTP cache is an IndexedDB package
installation cache.

## Expected savings (estimates, not guarantees)

Previous console measurements for this page were:

| Setup step | Observed time |
| --- | ---: |
| Pyodide runtime startup, including WASM downloads | 1.24–2.82 s |
| `micropip` package load | 0.30–0.62 s |
| `oletools` download and install, including dependencies | 1.41–2.02 s |
| Python analyzer setup (measured separately in one run) | about 1.06 s |
| Total engine setup | 4.73–5.93 s |

These are only two runs under different cache/network conditions, not a
controlled benchmark. If a valid persistent package cache avoids the
`oletools` installation step, the directly observed opportunity is roughly
**1.4–2.0 seconds per warm load**. Avoiding the `micropip` load as well could
add up to about **0.3–0.6 seconds**, but only if the implementation no longer
needs it on cache hits. Pyodide startup and analyzer initialization still
remain; the optimization will not make a reload instantaneous.

The earlier storage estimate was **about 10 MB** for installed `oletools` plus
its dependencies. This was an order-of-magnitude estimate, not a measured
IndexedDB footprint. It excludes the Pyodide runtime. It is also **not** a
promise of 10 MB less network traffic: installed files are expanded, while
network bytes are compressed downloads, and the browser's HTTP cache may
already avoid transferring some or all of those downloads. On a warm cache
hit, actual network savings can range from zero (if HTTP caching already
avoids the transfer) up to the measured transfer size of the package and
dependency downloads.

## How to refresh the numbers when Python tools change

Whenever the Pyodide, `oletools`, or Python dependency set changes:

1. Update the version pins and cache schema/key as needed. Resolve and record
   the full dependency versions for the new combination; do not key a cache
   only by `oletools` if its transitive dependencies can change.
2. With browser site data cleared and the browser's HTTP cache disabled, run
   the scanner and record per-resource **transferred bytes** and timings in
   DevTools Network. Separate Pyodide runtime assets from `micropip`,
   `oletools`, and dependency assets. Resource sizes from a warm HTTP cache
   are not cold-download sizes.
3. Measure persistent-storage growth before and after the cache is populated.
   Use `navigator.storage.estimate()` for an origin-level estimate and, where
   possible, sum the files in the cache mount. Record these separately from
   network bytes; origin-wide storage estimates include unrelated site data.
4. Measure engine setup on a cache miss and on a cache hit using the same
   browser, device, and network conditions. Run several trials and report the
   median, including cache status and version pins. Preserve the existing
   timing breakdown so the saved step can be identified rather than inferred
   only from total load time.
5. Re-run the measurements after changing versions, dependency resolution,
   cache layout, compression, or Pyodide. Update this plan's table and
   estimates with the new measured values, date, and test environment. Keep
   estimates labeled as estimates until repeated measurements confirm them.

## Validation checklist

- First visit installs packages, analyzes a supported document, and stores
  only runtime/package assets.
- Reload restores the exact compatible package set and passes the same
  analysis tests without reinstalling.
- Changing any cache-key version or dependency-lock value causes a cache miss.
- Missing, corrupt, or incompatible cache data is discarded/rebuilt safely.
- Blocked IndexedDB, quota errors, or cache eviction fall back to the existing
  install path and do not prevent analysis when network installation works.
- Cache writes are not reported as successful until the filesystem sync
  completes; interrupted writes do not appear as valid cache entries.
- No uploaded document bytes or analysis results appear in IndexedDB.

## References

- [Pyodide file-system guide](https://pyodide.org/en/stable/usage/file-system.html)
  documents `IDBFS` and syncing filesystems between page loads.
- [micropip API](https://micropip.pyodide.org/en/stable/project/api.html)
  documents browser package downloads, wheel support, and package metadata.
- Re-check both references against the project's pinned Pyodide version
  (`0.27.7`) before implementation; the stable documentation may describe a
  newer release.
