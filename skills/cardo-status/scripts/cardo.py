# -*- coding: utf-8 -*-
"""Cardo 插件启动器（cardo-status）。

用法（<技能目录> 取工具结果 <skill_resources> 的 resourceBase）：
    python <技能目录>/scripts/cardo.py --doctor         检查交付包、运行时与默认工作区
    python <技能目录>/scripts/cardo.py --where          打印交付包目录
    python <技能目录>/scripts/cardo.py [开工.py 参数...] 以当前目录为工作区运行开工.py

工作区：默认取当前目录；第一次调用先建目录架构（幂等，已存在不覆盖）并逐条说明放什么。
工作区还没有任何输入/过程件、且没有给流水线参数时，只建目录并停下等料，不自行往下跑。
显式 --workspace 越界（插件安装目录 / DSH_HOME）时拒绝并给出可用默认路径；
未显式指定且当前目录越界时，改用默认工作区 %USERPROFILE%/Cardo工作区/cardo-status-<时间戳>。

定位顺序：技能目录自带包 > --package > CARDO_PACKAGE > CARDO_HOME > 当前目录及上三级。
运行时顺序：CARDO_RUNTIME > 包内 runtime > CARDO_RUNTIME_HOME；都没有时用系统 Python，
首次运行由 开工.py 按 requirements.txt 在用户缓存建虚拟环境。不联网、不装依赖。
"""
from __future__ import annotations

import os
import subprocess
import sys
import time
from pathlib import Path

PACKAGE = "CardoStatus"
ENTRY = "开工.py"
NAME = "cardo-status"
VERSION = "0.1.1"
# 工作区目录架构与用途：与包内 开工.py 的 DIRS 对齐；只建目录，不写内容
DIR_GUIDE = [('输入文件', '放原始材料（现状底稿、项目资料，任意层级）'), ('输入文件-模板', '待填模板放这里'), ('提资数据/模板', '现状底稿模板'), ('过程文件', '流水线中间件与运行报告，跑完之前不要动'), ('成果文件', '最终交付件，跑完在这里取'), ('记忆与规则', '本项目的规则与记忆，跨会话复用')]
DIRS = [rel for rel, _ in DIR_GUIDE]


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


def _default_workspace() -> Path:
    """越界或缺失时的可用默认工作区；根目录可用 CARDO_WORKSPACE_HOME 覆盖。"""
    root = os.environ.get("CARDO_WORKSPACE_HOME") or str(Path.home() / "Cardo工作区")
    return Path(root) / ("%s-%s" % (NAME, time.strftime("%Y%m%d-%H%M%S")))


def scaffold(ws: Path):
    """按约定建工作区目录架构：幂等，已存在不动、不覆盖。"""
    made = []
    for rel in DIRS:
        target = ws / rel
        if target.is_dir():
            continue
        try:
            target.mkdir(parents=True, exist_ok=True)
            made.append(rel)
        except OSError as exc:
            print("目录建不出来：%s（%s）" % (target, exc), file=sys.stderr)
    return made


def _cell_width(text: str) -> int:
    """中文按两格算，保证目录说明上下对齐。"""
    return sum(2 if ord(ch) > 127 else 1 for ch in text)


def print_guide(ws: Path) -> None:
    """逐条说明每个目录放什么（首次调用 / 等料时打印）。"""
    print("工作区：%s" % ws, flush=True)
    print("目录架构（已存在的不覆盖）：", flush=True)
    width = max(_cell_width(rel) for rel in DIRS) + 2
    for rel, desc in DIR_GUIDE:
        label = rel + "\\"
        print("  %s%s  %s" % (label, " " * (width - _cell_width(label)), desc), flush=True)


def _has_files(ws: Path, *names) -> bool:
    for name in names:
        d = ws / name
        if d.is_dir() and any(p.is_file() for p in d.rglob("*")):
            return True
    return False


def _forbidden_roots(pkg):
    here = Path(__file__).resolve()
    roots = [here.parents[1], here.parents[2], here.parents[3]]
    if pkg is not None:
        roots.append(pkg.resolve())
    if os.environ.get("DSH_HOME"):
        try:
            roots.append(Path(os.environ["DSH_HOME"]).expanduser().resolve())
        except OSError:
            pass
    return roots


def doctor():
    pkg = locate()
    print("默认工作区：%s（--workspace 越界或当前目录越界时启用）" % _default_workspace())
    print("工作区目录架构：%s" % "、".join(DIRS))
    if pkg is None:
        print("未找到 Cardo 交付包 %s。" % PACKAGE, file=sys.stderr)
        print("请设置环境变量 CARDO_HOME，指向包含 %s\\ 的目录，例如：" % PACKAGE, file=sys.stderr)
        print("  setx CARDO_HOME \"C:\\path\\to\\cardo\"", file=sys.stderr)
        print("也可用 CARDO_PACKAGE 直接指向 %s 目录。" % PACKAGE, file=sys.stderr)
        return 2
    exe, rt = python_for(pkg)
    wire_runtime(pkg, rt)
    print("交付包：%s" % pkg)
    print("入口：%s（%s）" % (pkg / ENTRY, "存在" if (pkg / ENTRY).is_file() else "缺失"))
    print("运行时：%s（%s）" % (exe, "存在" if exe.is_file() else "缺失"))
    print("运行时目录：%s" % rt)
    print("技能目录：%s" % (pkg / "Skills"))
    return 0 if exe.is_file() else 2


def workspace_for(args, pkg, allow_fallback=True):
    """确定工作目录：只认用户给的路径（--workspace 或当前目录），不退回安装目录。

    显式 --workspace 越界或不存在：拒绝，并给出可用默认路径。
    未显式指定且当前目录越界：改用默认工作区（allow_fallback），并在输出里说明。
    """
    explicit = "--workspace" in args
    if explicit:
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

    def _inside(root: Path) -> bool:
        try:
            return ws == root or root in ws.parents
        except OSError:
            return False

    if any(_inside(root) for root in _forbidden_roots(pkg)):
        hint = _default_workspace()
        if explicit or not allow_fallback:
            print("工作目录不能设在插件安装目录或 DSH_HOME 内：%s" % ws, file=sys.stderr)
            print("请换一个独立空文件夹；也可直接用默认工作区：", file=sys.stderr)
            print("  python \"%s\" --workspace \"%s\"" % (Path(__file__).resolve(), hint), file=sys.stderr)
            return None
        try:
            hint.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            print("默认工作区建不出来：%s（%s）" % (hint, exc), file=sys.stderr)
            print("请用 --workspace 指定一个独立空文件夹后重试。", file=sys.stderr)
            return None
        print("当前目录在插件安装目录或 DSH_HOME 内，已改用默认工作区：%s" % hint, file=sys.stderr)
        print("下次可直接用 --workspace 指定自己的项目文件夹。", file=sys.stderr)
        ws = hint.resolve()

    if not ws.is_dir():
        if explicit:
            hint = _default_workspace()
            print("工作目录不存在：%s" % ws, file=sys.stderr)
            print("请先新建该文件夹，或改用默认工作区：", file=sys.stderr)
            print("  python \"%s\" --workspace \"%s\"" % (Path(__file__).resolve(), hint), file=sys.stderr)
            return None
        try:
            ws.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            print("工作目录不存在且建不出来：%s（%s）" % (ws, exc), file=sys.stderr)
            return None
        print("工作目录不存在，已新建：%s" % ws, file=sys.stderr)

    probe = ws / (".cardo-write-test-%d" % os.getpid())
    try:
        probe.write_bytes(b"")
        probe.unlink()
    except OSError:
        print("工作目录不可写：%s" % ws, file=sys.stderr)
        print("请换一个有写权限的文件夹，或调整该目录权限后重试。", file=sys.stderr)
        return None
    return ws


def _action_args(args):
    """去掉启动器自己的参数，只留下要转交 开工.py 的参数。"""
    out, skip = [], False
    for arg in args:
        if skip:
            skip = False
            continue
        if arg == "--workspace":
            skip = True
            continue
        out.append(arg)
    return out


def main(argv):
    if "--version" in argv:
        print("%s %s" % (NAME, VERSION))
        return 0
    if "--doctor" in argv:
        return doctor()
    if "--where" in argv:
        pkg = locate()
        if pkg is None:
            print("未找到交付包 %s" % PACKAGE, file=sys.stderr)
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
    args = list(argv)
    ws = workspace_for(args, pkg)
    if ws is None:
        return 2
    made = scaffold(ws)

    # 开场：建完目录先说明放什么；还没有输入/过程件且没给流水线参数时，停下等料
    intake = ("输入文件", "输入文件-模板", "提资数据", "过程文件", "成果文件")
    if made or not _has_files(ws, *intake):
        print_guide(ws)
    if not _action_args(args) and not _has_files(ws, *intake):
        print("", flush=True)
        print("还没有输入文件，先不开工。请把材料放进上面的「输入文件」目录（缺什么、放哪里见各目录说明）。", flush=True)
        print("放好后回一句「继续」，我再跑流水线。", flush=True)
        return 0

    if pkg is None:
        return doctor()
    exe, rt = python_for(pkg)
    wire_runtime(pkg, rt)
    if not exe.is_file():
        print("交付包缺少运行时：%s" % exe, file=sys.stderr)
        print("可设置 CARDO_RUNTIME 指向已有 runtime 目录（含 python\\python.exe），", file=sys.stderr)
        print("或按交付包 runtime/README.txt 重新解压完整包后再试。", file=sys.stderr)
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
    print("Cardo 包：%s" % pkg, file=sys.stderr, flush=True)
    print("工作区：%s" % ws, file=sys.stderr, flush=True)
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
