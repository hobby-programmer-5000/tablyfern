"""Build and smoke-check a portable Windows x64 release in a new directory."""

import argparse
import hashlib
import json
import os
import platform
import runpy
import shutil
import subprocess
import sys
from importlib.metadata import distribution
from pathlib import Path
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parent.parent
VERSION = runpy.run_path(str(ROOT / "version.py"))["VERSION"]
APP_MODULES = ("app.py", "converter.py", "exporters.py", "version.py")
DEPENDENCIES = (
    "PySide6-Essentials", "shiboken6", "XlsxWriter", "PyInstaller",
    "pyinstaller-hooks-contrib", "altgraph", "packaging", "pefile", "pywin32-ctypes", "setuptools",
)


def hash_file(path: Path) -> str:
    with path.open("rb") as file:
        return hashlib.file_digest(file, "sha256").hexdigest()


def version_resource() -> str:
    number = tuple(map(int, VERSION.split("."))) + (0,)
    return f"""VSVersionInfo(
  ffi=FixedFileInfo(filevers={number!r}, prodvers={number!r}, mask=0x3f,
    flags=0, OS=0x40004, fileType=1, subtype=0, date=(0, 0)),
  kids=[
    StringFileInfo([StringTable('040904B0', [
      StringStruct('FileDescription', 'Tablyfern'),
      StringStruct('FileVersion', '{VERSION}'),
      StringStruct('OriginalFilename', 'Tablyfern.exe'),
      StringStruct('ProductName', 'Tablyfern'),
      StringStruct('ProductVersion', '{VERSION}')
    ])]),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])
  ])"""


def collect_notices(output: Path) -> list[dict]:
    installed = output / "licenses" / "installed"
    installed.mkdir()
    records = []
    for name in DEPENDENCIES:
        package = distribution(name)
        records.append({
            "name": package.metadata["Name"],
            "version": package.version,
            "license": package.metadata.get("License-Expression") or package.metadata.get("License"),
        })
        for entry in package.files or []:
            if ".." in entry.parts:
                continue
            if "license" not in str(entry).lower() and Path(entry).name.lower() not in {"copying", "notice"}:
                continue
            source = Path(package.locate_file(entry))
            if source.is_file():
                target = installed / name / Path(entry)
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(source, target)
    shutil.copy2(Path(sys.base_prefix) / "LICENSE.txt", installed / "CPython-LICENSE.txt")
    return records


def use_system_msvc_runtime(internal: Path) -> None:
    prefixes = ("msvcp140", "vcruntime140", "concrt140", "vcomp140", "api-ms-win-crt-", "ucrtbase")
    for path in internal.rglob("*.dll"):
        if path.name.lower().startswith(prefixes):
            path.unlink()


def build(dist: Path) -> None:
    if sys.platform != "win32" or platform.machine().lower() not in {"amd64", "x86_64"}:
        raise SystemExit("Build the Windows x64 release on a Windows x64 interpreter.")
    dist.mkdir(parents=True)
    work = dist / ".build"
    work.mkdir()
    version_file = work / "version.txt"
    version_file.write_text(version_resource(), encoding="utf-8")

    import PySide6

    environment = os.environ.copy()
    windows = Path(os.environ["WINDIR"])
    environment["PATH"] = os.pathsep.join(map(str, (
        Path(sys.base_prefix), Path(PySide6.__file__).parent, windows / "System32", windows,
    )))
    command = [
        sys.executable, "-m", "PyInstaller", "--clean", "--onedir", "--windowed", "--noupx",
        "--name", "Tablyfern", "--version-file", str(version_file), "--specpath", str(work),
        "--distpath", str(dist), "--workpath", str(work / "work"), str(ROOT / "app.py"),
    ]
    subprocess.run(command, cwd=ROOT, env=environment, check=True)
    output = dist / "Tablyfern"
    use_system_msvc_runtime(output / "_internal")
    for name in (
        "README.md", "README.ko.md", "LICENSE", "CONTRIBUTORS.md", "CONTRIBUTORS.ko.md",
        "THIRD-PARTY-NOTICES.md", "THIRD-PARTY-NOTICES.ko.md",
    ):
        shutil.copy2(ROOT / name, output / name)
    for name in ("docs", "samples", "licenses"):
        shutil.copytree(ROOT / name, output / name)
    build_info = {
        "app": "Tablyfern", "version": VERSION, "platform": "Windows x64",
        "python": platform.python_version(), "layout": "onedir with replaceable Qt libraries",
        "windows_runtime": "system Microsoft Visual C++ v14 x64 runtime",
        "dependencies": collect_notices(output),
        "source_sha256": {name: hash_file(ROOT / name) for name in APP_MODULES},
    }
    (output / "BUILD-INFO.json").write_text(json.dumps(build_info, indent=2) + "\n", encoding="utf-8")
    binaries = {
        path.relative_to(output).as_posix(): hash_file(path)
        for path in sorted(output.rglob("*")) if path.suffix.lower() in {".exe", ".dll", ".pyd"}
    }
    (output / "BINARY-MANIFEST.json").write_text(json.dumps(binaries, indent=2) + "\n", encoding="utf-8")
    smoke_env = {name: value for name, value in environment.items() if not name.startswith(("PYTHON", "QT_"))}
    smoke_env["QT_QPA_PLATFORM"] = "windows"
    smoke = work / "smoke"
    subprocess.run([str(output / "Tablyfern.exe"), "--smoke-check", str(smoke)],
                   cwd=output, env=smoke_env, check=True, timeout=60)
    result = json.loads((smoke / "smoke.json").read_text(encoding="utf-8"))
    if result != {"version": VERSION, "frozen": True, "tabs": ["orders", "events", "empty"]}:
        raise RuntimeError(f"Unexpected frozen smoke result: {result}")
    stem = f"Tablyfern-{VERSION}-windows-x64"
    archive = Path(shutil.make_archive(str(dist / stem), "zip", dist, "Tablyfern"))
    with ZipFile(archive) as zipped:
        bad = zipped.testzip()
        if bad is not None:
            raise RuntimeError(f"ZIP CRC check failed: {bad}")
    (dist / "SHA256SUMS.txt").write_text(f"{hash_file(archive)}  {archive.name}\n", encoding="ascii")
    print(f"Verified frozen build and ZIP: {archive}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--dist-dir", type=Path, default=ROOT / "dist" / VERSION)
    build(parser.parse_args().dist_dir.resolve())
