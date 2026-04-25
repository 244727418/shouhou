#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""1.7 发布构建脚本。

默认优先构建 onedir 版本，onefile 作为可选对比产物。
强制使用项目本地 .venv，避免误用全局 Conda 环境。
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
    "onedir": PROJECT_ROOT / "售后登记表_v1.7_onedir.spec",
    "onefile": PROJECT_ROOT / "售后登记表_v1.7_onefile.spec",
}
ALIAS_ENV_KEY = "SHOUHOU_BUILD_ALIAS_ACTIVE"
ALIAS_DRIVES = ("X:", "Y:", "Z:")


def parse_args():
    parser = argparse.ArgumentParser(description="构建售后登记表 1.7 发布包")
    parser.add_argument(
        "--mode",
        choices=("onedir", "onefile", "all"),
        default="onedir",
        help="构建模式，默认 onedir",
    )
    parser.add_argument("--clean", action="store_true", help="构建前清理 build/dist 目录")
    return parser.parse_args()


def needs_ascii_alias(path: Path):
    return any(ord(char) > 127 for char in str(path))


def find_available_drive():
    for drive in ALIAS_DRIVES:
        if not Path(f"{drive}\\").exists():
            return drive
    raise RuntimeError("未找到可用的临时盘符，请先释放 X:/Y:/Z: 之一。")


def maybe_reexec_via_ascii_alias():
    if os.environ.get(ALIAS_ENV_KEY) == "1":
        return
    if not needs_ascii_alias(PROJECT_ROOT):
        return

    drive = find_available_drive()
    project_root_text = str(PROJECT_ROOT)
    alias_python = f"{drive}\\.venv\\Scripts\\python.exe"
    alias_script = f"{drive}\\scripts\\build_release.py"
    env = os.environ.copy()
    env[ALIAS_ENV_KEY] = "1"

    subprocess.run(["subst", drive, project_root_text], check=True)
    try:
        subprocess.run([alias_python, alias_script, *sys.argv[1:]], check=True, env=env, cwd=f"{drive}\\")
    finally:
        subprocess.run(["subst", drive, "/d"], check=False)

    raise SystemExit(0)


def require_local_venv():
    if not VENV_PYTHON.exists():
        raise SystemExit(
            "未找到项目本地虚拟环境解释器: .venv\\Scripts\\python.exe\n"
            "请先在项目目录创建 .venv，并安装 requirements.txt 与 pyinstaller。"
        )

    active_python = Path(sys.executable).resolve()
    expected_python = VENV_PYTHON.resolve()
    if active_python != expected_python:
        print(f"[build] 当前解释器: {active_python}")
        print(f"[build] 目标解释器: {expected_python}")
        print("[build] 将改用项目本地 .venv 执行 PyInstaller。")


def clean_outputs():
    for target in (BUILD_ROOT, DIST_ROOT):
        if target.exists():
            try:
                shutil.rmtree(target)
                print(f"[clean] 已清理 {target}")
            except PermissionError as exc:
                print(f"[clean] 跳过 {target}: {exc}")


def run_pyinstaller(spec_path: Path, dist_subdir: str):
    dist_path = DIST_ROOT / dist_subdir
    work_path = BUILD_ROOT / f"1.7_{dist_subdir}"
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
    print("[build] 执行:", " ".join(cmd))
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

    print("[build] 构建完成")


if __name__ == "__main__":
    main()
