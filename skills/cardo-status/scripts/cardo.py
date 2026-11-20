# -*- coding: utf-8 -*-
"""Cardo 插件启动器（cardo-status）。

用法（<技能目录> 取工具结果 <skill_resources> 的 resourceBase）：
    python <技能目录>/scripts/cardo.py --doctor         检查交付包与运行时
    python <技能目录>/scripts/cardo.py --where          打印交付包目录
    python <技能目录>/scripts/cardo.py [开工.py 参数...] 以当前目录为工作区运行开工.py

定位顺序：--package > CARDO_PACKAGE > CARDO_HOME/CardoStatus > 当前目录及上三级 > %USERPROFILE%/Cardo交付。
运行时顺序：CARDO_RUNTIME（runtime 目录）> 交付包自带 runtime > CARDO_RUNTIME_HOME/CardoStatus/runtime
> CARDO_HOME/CardoStatus/runtime。不联网、不装依赖。
"""
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

PACKAGE = "CardoStatus"
ENTRY = "开工.py"


def _roots():
    yield Path(__file__).resolve().parents[1]   # 派生仓库：技能目录自带 开工.py 与 Skills
    env_pkg = os.environ.get("CARDO_PACKAGE")
    if env_pkg:
        yield Path(env_pkg)
    env_home = os.environ.get("CARDO_HOME")
    if env_home:
        yield Path(env_home) / PACKAGE
    here = Path.cwd()
    for base in (here, *list(here.parents)[:3]):
        yield base / PACKAGE
        yield base / "06_Cardo交付" / PACKAGE
    yield Path.home() / "Cardo交付" / PACKAGE


def wire_runtime(pkg: Path, rt: Path) -> None:
    """外部 runtime（CARDO_RUNTIME / CARDO_RUNTIME_HOME）就位时，给包内自检补一条目录联接。

    交付包的 `开工.py` 会检查 `<包>/runtime/python`；插件不复制 2 GB 运行时，
    而是把用户已有的 runtime 接到包目录下。仅在包内 runtime 缺失、且显式配置了
    外部 runtime 时执行；建不成联接不阻断（继续用外部解释器跑）。
    """
    target = pkg / "runtime"
    if not rt.is_dir() or not (os.environ.get("CARDO_RUNTIME") or os.environ.get("CARDO_RUNTIME_HOME")):
        return
    if target.exists():
        # 已有一条联接但指向别的 runtime（换了配置项）：换指；真实目录一律不动
        try:
            stale = target.resolve() != rt.resolve()
            reparse = _is_reparse(target)
            if stale and reparse:
                os.rmdir(str(target))
            else:
                return
        except OSError:
            return
    try:
        if os.name == "nt":
            subprocess.run(["cmd", "/c", "mklink", "/J", str(target), str(rt)],
                           check=True, capture_output=True)
        else:
            os.symlink(rt, target, target_is_directory=True)
    except Exception:
        pass


def locate():
    seen = set()
    for root in _roots():
        try:
            path = root.expanduser().resolve()
        except OSError:
            continue
        if path in seen:
            continue
        seen.add(path)
        if (path / ENTRY).is_file():
            return path
    return None


def _is_reparse(path: Path) -> bool:
    if os.name != "nt":
        return os.path.islink(path)
    import ctypes
    attrs = ctypes.windll.kernel32.GetFileAttributesW(str(path))
    return attrs != -1 and bool(attrs & 0x400)


def python_for(pkg):
    """返回 (python 可执行文件, runtime 目录)：优先用用户指定/已装的运行时。"""
    exe = "python.exe" if os.name == "nt" else "python"
    dirs = []
    env_rt = os.environ.get("CARDO_RUNTIME")
    if env_rt:
        dirs.append(Path(env_rt))
    local = pkg / "runtime"
    env_home = os.environ.get("CARDO_RUNTIME_HOME") or os.environ.get("CARDO_HOME")
    home_rt = Path(env_home) / PACKAGE / "runtime" if env_home else None
    # 包内 runtime 若是上一次配置留下的联接，且这次指定了外部根，先按新配置解析
    if home_rt is not None and local.is_dir() and _is_reparse(local):
        dirs.extend([home_rt, local])
    else:
        dirs.append(local)
        if home_rt is not None:
            dirs.append(home_rt)
    for rt in dirs:
        exe_path = rt / "python" / exe
        if exe_path.is_file():
            return exe_path, rt
    # 派生仓库不带 runtime：回落到系统 Python（开工.py 会按 requirements.txt 建 .venv）
    if Path(sys.executable).is_file():
        return Path(sys.executable), dirs[0]
    return dirs[0] / "python" / exe, dirs[0]


def doctor():
    pkg = locate()
    if pkg is None:
        print("未找到 Cardo 交付包 CardoStatus。", file=sys.stderr)
        print("请设置环境变量 CARDO_HOME，指向包含 CardoStatus\\ 的目录，例如：", file=sys.stderr)
        print("  setx CARDO_HOME \"C:\\path\\to\\cardo\"", file=sys.stderr)
        print("也可用 CARDO_PACKAGE 直接指向 CardoStatus 目录。", file=sys.stderr)
        return 2
    exe, rt = python_for(pkg)
    wire_runtime(pkg, rt)
    print("交付包：%s" % pkg)
    print("入口：%s（%s）" % (pkg / ENTRY, "存在" if (pkg / ENTRY).is_file() else "缺失"))
    print("运行时：%s（%s）" % (exe, "存在" if exe.is_file() else "缺失"))
    print("运行时目录：%s" % rt)
    print("技能目录：%s" % (pkg / "Skills"))
    return 0 if exe.is_file() else 2


def workspace_for(args, pkg):
    """确定工作目录：只认用户给的路径（--workspace 或当前目录），不退回安装目录。"""
    if "--workspace" in args:
        i = args.index("--workspace")
        if i + 1 >= len(args):
            print("--workspace 需要一个目录参数，例如 --workspace <文件夹>（当前包 " + PACKAGE + "）", file=sys.stderr)
            return None
        ws = Path(args[i + 1]).expanduser()
    else:
        ws = Path.cwd()
    try:
        ws = ws.resolve()
    except OSError:
        print("工作目录无法解析：%s" % ws, file=sys.stderr)
        return None
    # 插件包根（node_modules/cardo-*）：不许把成果写进安装目录
    install_root = Path(__file__).resolve().parents[3]
    homes = [pkg.resolve(), install_root]
    if os.environ.get("DSH_HOME"):
        homes.append(Path(os.environ["DSH_HOME"]).expanduser().resolve())
    for root in homes:
        if ws == root or root in ws.parents:
            print("工作目录不能设在插件安装目录或 DSH_HOME 内：%s" % ws, file=sys.stderr)
            print("请新建一个空文件夹，用 --workspace <文件夹> 指定，或把它设为当前目录。", file=sys.stderr)
            return None
    if not ws.is_dir():
        print("工作目录不存在：%s" % ws, file=sys.stderr)
        print("请先新建一个空文件夹，再把它作为工作目录。", file=sys.stderr)
        return None
    probe = ws / (".cardo-write-test-%d" % os.getpid())
    try:
        probe.write_bytes(b"")
        probe.unlink()
    except OSError:
        print("工作目录不可写：%s" % ws, file=sys.stderr)
        print("请换一个有写权限的文件夹，或调整该目录权限后重试。", file=sys.stderr)
        return None
    return ws


def main(argv):
    if "--doctor" in argv:
        return doctor()
    if "--where" in argv:
        pkg = locate()
        if pkg is None:
            print("未找到交付包 CardoStatus", file=sys.stderr)
            return 2
        print(pkg)
        return 0
    if "--package" in argv:
        i = argv.index("--package")
        if i + 1 >= len(argv):
            print("--package 需要一个目录参数", file=sys.stderr)
            return 2
        os.environ["CARDO_PACKAGE"] = argv[i + 1]
        argv = argv[:i] + argv[i + 2:]

    pkg = locate()
    if pkg is None:
        return doctor()
    exe, rt = python_for(pkg)
    wire_runtime(pkg, rt)
    if not exe.is_file():
        print("交付包缺少运行时：%s" % exe, file=sys.stderr)
        print("可设置 CARDO_RUNTIME 指向已有 runtime 目录（含 python\\python.exe），", file=sys.stderr)
        print("或按交付包 runtime/README.txt 重新解压完整包后再试。", file=sys.stderr)
        return 2

    args = list(argv)
    ws = workspace_for(args, pkg)
    if ws is None:
        return 2
    if "--workspace" not in args:
        args = ["--workspace", str(ws)] + args

    env = dict(os.environ)
    env.update({"PYTHONIOENCODING": "utf-8", "PYTHONUTF8": "1", "PYTHONDONTWRITEBYTECODE": "1"})
    # 临时文件也留在工作区：交付包里有库会往 %TEMP% 建目录，不让它落到公共临时区
    scope_tmp = ws / ".cardo-tmp"
    try:
        scope_tmp.mkdir(exist_ok=True)
        env["TEMP"] = env["TMP"] = str(scope_tmp)
    except OSError:
        pass
    print("Cardo 包：%s" % pkg, file=sys.stderr)
    print("工作区：%s" % ws, file=sys.stderr)
    try:
        return subprocess.call([str(exe), str(pkg / ENTRY), *args], cwd=str(ws), env=env)
    except KeyboardInterrupt:
        return 130
    finally:
        try:
            scope_tmp.rmdir()   # 只在空目录时删掉，留文件就留在工作区
        except OSError:
            pass


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
