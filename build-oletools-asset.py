from pathlib import Path
from tempfile import NamedTemporaryFile
from zipfile import ZIP_DEFLATED, ZipFile
import os


ROOT = Path(__file__).resolve().parent
OLETOOLS_ROOT = ROOT / "submodules" / "oletools"
PCODEDMP_ROOT = ROOT / "submodules" / "pcodedmp"
PACKAGE_ROOTS = [
    OLETOOLS_ROOT / "oletools",
    PCODEDMP_ROOT / "pcodedmp",
]
ASSET_ROOT = ROOT / "dist"
ASSET_PATH = ASSET_ROOT / "oletools.zip"

for package_root in PACKAGE_ROOTS:
    if not package_root.is_dir():
        raise SystemExit(f"Required package submodule is not initialized: {package_root.parent}")

ASSET_ROOT.mkdir(exist_ok=True)
with NamedTemporaryFile(dir=ASSET_ROOT, suffix=".zip", delete=False) as temporary_file:
    temporary_path = Path(temporary_file.name)

try:
    with ZipFile(temporary_path, "w", ZIP_DEFLATED) as archive:
        for package_root in PACKAGE_ROOTS:
            for source_path in sorted(package_root.rglob("*")):
                if source_path.is_file() and "__pycache__" not in source_path.parts:
                    archive.write(source_path, source_path.relative_to(package_root.parent).as_posix())
        archive.write(PCODEDMP_ROOT / "LICENSE", "pcodedmp/LICENSE")
    os.chmod(temporary_path, 0o644)
    os.replace(temporary_path, ASSET_PATH)
finally:
    temporary_path.unlink(missing_ok=True)

print(f"Built {ASSET_PATH.relative_to(ROOT)} ({ASSET_PATH.stat().st_size:,} bytes)")