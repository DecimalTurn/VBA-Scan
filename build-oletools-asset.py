from pathlib import Path
from tempfile import NamedTemporaryFile
from zipfile import ZIP_DEFLATED, ZipFile
import os


ROOT = Path(__file__).resolve().parent
SUBMODULE_ROOT = ROOT / "submodules" / "oletools"
PACKAGE_ROOT = SUBMODULE_ROOT / "oletools"
ASSET_ROOT = ROOT / "dist"
ASSET_PATH = ASSET_ROOT / "oletools.zip"

if not PACKAGE_ROOT.is_dir():
    raise SystemExit("The submodules/oletools submodule is not initialized.")

ASSET_ROOT.mkdir(exist_ok=True)
with NamedTemporaryFile(dir=ASSET_ROOT, suffix=".zip", delete=False) as temporary_file:
    temporary_path = Path(temporary_file.name)

try:
    with ZipFile(temporary_path, "w", ZIP_DEFLATED) as archive:
        for source_path in sorted(PACKAGE_ROOT.rglob("*")):
            if source_path.is_file() and "__pycache__" not in source_path.parts:
                archive.write(source_path, source_path.relative_to(SUBMODULE_ROOT).as_posix())
    os.chmod(temporary_path, 0o644)
    os.replace(temporary_path, ASSET_PATH)
finally:
    temporary_path.unlink(missing_ok=True)

print(f"Built {ASSET_PATH.relative_to(ROOT)} ({ASSET_PATH.stat().st_size:,} bytes)")