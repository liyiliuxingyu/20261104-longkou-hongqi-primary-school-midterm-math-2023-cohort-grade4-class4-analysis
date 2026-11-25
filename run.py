#!/usr/bin/env python3
"""Portable entry point for the Grade 4 November 2026 continuation project.

Verification and analysis use only the Python standard library. Spreadsheet
rebuilding uses the host's installed @oai/artifact-tool in a temporary workspace.
No command downloads packages, uploads student data, or overwrites the baseline.
"""
from __future__ import annotations
import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import zipfile

ROOT = Path(__file__).resolve().parent
CONFIG = json.loads((ROOT / "project.json").read_text(encoding="utf-8"))
BASELINE = ROOT / CONFIG["baseline"]
EXCLUDED = {"outputs", "build", "__pycache__", "node_modules", ".git", ".runtime"}

def sha256(path):
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()

def launch(argv, env=None):
    result = subprocess.run([str(x) for x in argv], cwd=ROOT, env=env, check=False)
    if result.returncode:
        raise RuntimeError("Command failed with exit status " + str(result.returncode))

def verify():
    manifest_path = ROOT / "MANIFEST.json"
    if not manifest_path.exists():
        raise RuntimeError("MANIFEST.json is missing. Extract the complete ZIP before verification.")
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    errors = []
    for item in manifest["files"]:
        relative = Path(item["path"])
        if relative.is_absolute() or ".." in relative.parts:
            errors.append("Unsafe manifest path: " + item["path"])
            continue
        file = ROOT / relative
        if not file.is_file() or file.is_symlink():
            errors.append("Missing file or unexpected symlink: " + item["path"])
        elif file.stat().st_size != item["size_bytes"] or sha256(file) != item["sha256"]:
            errors.append("Changed or damaged file: " + item["path"])
    with zipfile.ZipFile(BASELINE) as workbook:
        if workbook.testzip() is not None:
            errors.append("The baseline XLSX has a failed ZIP CRC check.")
    result = {"status": "passed" if not errors else "failed", "verified_files": len(manifest["files"]), "errors": errors}
    print(json.dumps(result, ensure_ascii=False, indent=2))
    if errors:
        raise RuntimeError("Integrity verification failed; see the listed file paths.")

def get_runtime():
    node = os.environ.get("CODEX_PRIMARY_RUNTIME_NODE") or os.environ.get("GRADE4_NODE") or shutil.which("node")
    modules = os.environ.get("CODEX_PRIMARY_RUNTIME_NODE_MODULES") or os.environ.get("GRADE4_NODE_MODULES")
    if not modules and (ROOT / "node_modules" / "@oai" / "artifact-tool" / "package.json").exists():
        modules = str(ROOT / "node_modules")
    return node, modules

def doctor():
    node, modules = get_runtime()
    package = Path(modules) / "@oai" / "artifact-tool" / "package.json" if modules else None
    available = bool(node and package and package.is_file())
    version = json.loads(package.read_text(encoding="utf-8")).get("version") if package and package.is_file() else None
    node_version = subprocess.check_output([node, "--version"], text=True).strip() if node else None
    print(json.dumps({"python": sys.version.split()[0], "analysis_requires_third_party_packages": False,
                      "node": node_version, "artifact_tool": version, "can_rebuild": available,
                      "baseline_exists": BASELINE.is_file()}, ensure_ascii=False, indent=2))
    return available

def analyze(input_path, output_dir):
    output_dir.mkdir(parents=True, exist_ok=True)
    launch([sys.executable, ROOT / "src" / "audit.py", "--input", input_path, "--output-dir", output_dir])

def rebuild(input_path, output_dir, render=False):
    node, modules = get_runtime()
    if not node or not modules or not (Path(modules) / "@oai" / "artifact-tool" / "package.json").is_file():
        raise RuntimeError("XLSX rebuilding requires Node.js and an installed @oai/artifact-tool. "
                           "Use a ChatGPT spreadsheet workspace with CODEX_PRIMARY_RUNTIME_NODE and "
                           "CODEX_PRIMARY_RUNTIME_NODE_MODULES, or explicitly set GRADE4_NODE and "
                           "GRADE4_NODE_MODULES. Verification and analysis do not require this dependency.")
    if input_path.resolve() == (output_dir / CONFIG["rebuilt_filename"]).resolve():
        raise RuntimeError("Input and output must be different files.")
    analyze(input_path, output_dir / "analysis")
    with tempfile.TemporaryDirectory(prefix="grade4_rebuild_") as temporary:
        stage = Path(temporary)
        shutil.copyfile(ROOT / "src" / "rebuild.mjs", stage / "rebuild.mjs")
        try:
            (stage / "node_modules").symlink_to(Path(modules).resolve(), target_is_directory=True)
        except OSError as error:
            raise RuntimeError("Cannot link the host Node dependencies into the temporary build directory. "
                               "The tested full-build environment is Linux; use that environment for rebuilding.") from error
        environment = os.environ.copy()
        environment.update({"GRADE4_PROJECT_ROOT": str(ROOT), "GRADE4_INPUT_XLSX": str(input_path.resolve()),
                            "GRADE4_OUTPUT_DIR": str(output_dir.resolve()), "GRADE4_PYTHON": sys.executable,
                            "GRADE4_RENDER": "1" if render else "0"})
        launch([node, stage / "rebuild.mjs"], env=environment)
    candidate = output_dir / CONFIG["rebuilt_filename"]
    launch([sys.executable, ROOT / "src" / "validate_workbook.py", "--input", candidate,
            "--reference", input_path, "--report", output_dir / "rebuild_validation.json"])
    print("Rebuilt workbook: " + str(candidate))

def payload_files():
    files = []
    for file in sorted(ROOT.rglob("*")):
        relative = file.relative_to(ROOT)
        if any(part in EXCLUDED or part.startswith(".") for part in relative.parts):
            continue
        if not file.is_file() or file.is_symlink() or file.name == "MANIFEST.json" or file.suffix == ".zip":
            continue
        files.append(file)
    return files

def pack(destination):
    files = payload_files()
    required = [BASELINE, ROOT / "run.py", ROOT / "src/audit.py", ROOT / "src/rebuild.mjs",
                ROOT / "src/validate_workbook.py", ROOT / "src/preserve_native_metadata.py",
                ROOT / "README.md", ROOT / "HANDOFF.md", ROOT / "STATUS.json"]
    missing = [str(p.relative_to(ROOT)) for p in required if not p.is_file()]
    if missing:
        raise RuntimeError("Required files missing: " + ", ".join(missing))
    manifest = {"schema_version": 1, "project": CONFIG["project_id"],
                "manifest_scope": "All payload files except MANIFEST.json itself; transient outputs are excluded.",
                "files": [{"path": p.relative_to(ROOT).as_posix(), "size_bytes": p.stat().st_size,
                           "sha256": sha256(p)} for p in files]}
    (ROOT / "MANIFEST.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    destination.parent.mkdir(parents=True, exist_ok=True)
    folder = CONFIG["archive_root"]
    with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for p in files + [ROOT / "MANIFEST.json"]:
            archive.write(p, folder + "/" + p.relative_to(ROOT).as_posix())
    with zipfile.ZipFile(destination) as archive:
        if archive.testzip() is not None:
            raise RuntimeError("Created ZIP failed its CRC check.")
    print(json.dumps({"archive": str(destination), "payload_file_count": len(files) + 1,
                      "bytes": destination.stat().st_size, "sha256": sha256(destination)}, ensure_ascii=False, indent=2))

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    commands.add_parser("verify", help="Verify packaged SHA-256 hashes; Python standard library only.")
    commands.add_parser("doctor", help="Report local runtime availability without installing anything.")
    for verb in ["analyze", "rebuild"]:
        sub = commands.add_parser(verb)
        sub.add_argument("--input", type=Path, default=BASELINE)
        sub.add_argument("--output-dir", type=Path, default=ROOT / "outputs" / verb)
        if verb == "rebuild":
            sub.add_argument("--render", action="store_true", help="Also render diagnostic PNG previews.")
    packing = commands.add_parser("pack", help="Create a new continuation ZIP with a fresh manifest.")
    packing.add_argument("--destination", required=True, type=Path)
    args = parser.parse_args()
    if args.command == "verify": verify()
    elif args.command == "doctor": doctor()
    elif args.command == "analyze": analyze(args.input.resolve(), args.output_dir.resolve())
    elif args.command == "rebuild": rebuild(args.input.resolve(), args.output_dir.resolve(), args.render)
    elif args.command == "pack": pack(args.destination.resolve())

if __name__ == "__main__":
    try:
        main()
    except (OSError, RuntimeError, ValueError, zipfile.BadZipFile) as error:
        print("ERROR: " + str(error), file=sys.stderr)
        sys.exit(1)
