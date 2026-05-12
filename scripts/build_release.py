#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""2.3.1 release build script.

Builds the stable onefile package by default. The onedir target is kept only
as a troubleshooting option.
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path.cwd()
DIST_ROOT = PROJECT_ROOT / "release"
BUILD_ROOT = PROJECT_ROOT / "build"
VENV_PYTHON = PROJECT_ROOT / ".venv" / "Scripts" / "python.exe"
SPEC_FILES = {
    "onedir": PROJECT_ROOT / "售后登记表_v2.3.1_onedir.spec",
    "onefile": PROJECT_ROOT / "售后登记表_v2.3.1_onefile.spec",
}
ALIAS_ENV_KEY = "SHOUHOU_BUILD_ALIAS_ACTIVE"
ALIAS_DRIVES = ("X:", "Y:", "Z:")


def parse_args():
    parser = argparse.ArgumentParser(description="Build 售后登记表 2.3.1 release package")
    parser.add_argument(
        "--mode",
        choices=("onedir", "onefile", "all"),
        default="onefile",
        help="Build mode, default: onefile",
    )
    parser.add_argument("--clean", action="store_true", help="Clean build/release before building")
    return parser.parse_args()


def needs_ascii_alias(path: Path):
    return any(ord(char) > 127 for char in str(path))


def find_available_drive():
    for drive in ALIAS_DRIVES:
        if not Path(f"{drive}\\").exists():
            return drive
    raise RuntimeError("No available temporary drive letter. Please free X:, Y:, or Z:.")


def maybe_reexec_via_ascii_alias():
    if os.environ.get(ALIAS_ENV_KEY) == "1":
        return
    if not needs_ascii_alias(PROJECT_ROOT):
        return

    drive = find_available_drive()
    env = os.environ.copy()
    env[ALIAS_ENV_KEY] = "1"

    subprocess.run(["subst", drive, str(PROJECT_ROOT)], check=True)
    try:
        subprocess.run(
            [
                f"{drive}\\.venv\\Scripts\\python.exe",
                f"{drive}\\scripts\\build_release.py",
                *sys.argv[1:],
            ],
            check=True,
            env=env,
            cwd=f"{drive}\\",
        )
    finally:
        subprocess.run(["subst", drive, "/d"], check=False)

    raise SystemExit(0)


def require_local_venv():
    if not VENV_PYTHON.exists():
        raise SystemExit(
            "Missing project virtual environment: .venv\\Scripts\\python.exe\n"
            "Create .venv in the project root and install requirements.txt plus pyinstaller."
        )

    active_python = Path(sys.executable).resolve()
    expected_python = VENV_PYTHON.resolve()
    if active_python != expected_python:
        print(f"[build] Current Python: {active_python}")
        print(f"[build] Project Python: {expected_python}")
        print("[build] PyInstaller will run through the project .venv.")


def clean_outputs():
    for target in (BUILD_ROOT, DIST_ROOT):
        if target.exists():
            try:
                shutil.rmtree(target)
                print(f"[clean] Removed {target}")
            except PermissionError as exc:
                print(f"[clean] Skipped {target}: {exc}")


def run_pyinstaller(spec_path: Path, dist_subdir: str):
    if not spec_path.exists():
        raise SystemExit(f"Missing spec file: {spec_path}")

    dist_path = DIST_ROOT / dist_subdir
    work_path = BUILD_ROOT / f"2.3.1_{dist_subdir}"
    dist_path.mkdir(parents=True, exist_ok=True)
    work_path.mkdir(parents=True, exist_ok=True)

    cmd = [
        str(VENV_PYTHON),
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        f"--distpath={dist_path}",
        f"--workpath={work_path}",
        str(spec_path),
    ]
    print("[build] Running:", " ".join(cmd))
    subprocess.run(cmd, cwd=PROJECT_ROOT, check=True)


def main():
    args = parse_args()
    maybe_reexec_via_ascii_alias()
    require_local_venv()

    if args.clean:
        clean_outputs()

    modes = ("onedir", "onefile") if args.mode == "all" else (args.mode,)
    for mode in modes:
        run_pyinstaller(SPEC_FILES[mode], mode)

    print("[build] Done")


if __name__ == "__main__":
    main()
