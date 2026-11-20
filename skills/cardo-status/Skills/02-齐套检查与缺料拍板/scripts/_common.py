"""cardo-project-material 三个脚本的公共工具。

约定：
- 脚本自身不含盘符、项目名等硬编码；根目录/项目名/规则一律由参数或配置传入。
- 所有文本读写使用 UTF-8（读时容忍 BOM），写文件统一 UTF-8 无 BOM + LF。
"""
from __future__ import annotations

import datetime
import hashlib
import json
import os
import sys
from pathlib import Path

# 运行时是 python313._pth 模式，脚本目录不进 sys.path：兄弟模块必须显式加进来
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from _yaml_lite import YamlError, parse  # noqa: E402

SKILL_DIR = Path(__file__).resolve().parent.parent
RULES_FILE = SKILL_DIR / "规则" / "分类规则.yaml"
LIB_CONFIG_FILE = SKILL_DIR / "规则" / "项目库配置.yaml"

GATE_EXIT = 3
MISSING_EXIT = 2

# ---- T20：用户目录（任意项目目录）口径 ----
# 项目目录的判定标记：有项目索引（真实项目）或有 .cardo/（前端建档锚）。
INDEX_NAME = "00_项目索引.md"
CARD_DIR = ".cardo"
CARD_FILE = "project.json"
USER_SKELETON_DIRS = ("01_原始输入", "02_分类材料", "03_拆解成果", "04_成果报告", "05_过程", "06_归档")


class UserProjectError(Exception):
    """T20 用户目录模式的参数/输入错误：调用方统一返回 2（旧 --root 模式行为不变）。"""

# 项目库按「方向」分层：<项目库根>\<方向>\<项目名>\
# 方向名可由 规则/项目库配置.yaml 的 directions 覆盖；这里只作为兜底默认值，
# 不出现任何具体项目名或盘符。
DEFAULT_DIRECTIONS = ("老旧街区项目", "老旧厂区项目", "老旧小区项目")
DEFAULT_DIRECTION = "老旧街区项目"


def load_lib_config() -> dict:
    if LIB_CONFIG_FILE.exists():
        try:
            return load_yaml_file(LIB_CONFIG_FILE) or {}
        except SystemExit:
            return {}
    return {}


def direction_names() -> tuple[str, ...]:
    cfg = load_lib_config()
    raw = cfg.get("directions")
    if isinstance(raw, list):
        names = tuple(str(item).strip() for item in raw if str(item).strip())
        if names:
            return names
    return DEFAULT_DIRECTIONS


def default_direction() -> str:
    cfg = load_lib_config()
    configured = cfg.get("default_direction")
    if isinstance(configured, str) and configured.strip():
        return configured.strip()
    return DEFAULT_DIRECTION


def resolve_project_dir(root: Path, project: str, direction: str | None = None) -> Path | None:
    """定位项目目录。

    优先级：显式 --direction > <root>\\<方向>\\<项目名> 唯一命中 > 旧扁平布局 <root>\\<项目名>。
    命中多个方向 → 报错让调用方走闸门（同名项目不允许跨方向重复）。
    都没命中返回 None，由调用方决定是「报错」还是「新建」。
    """
    if direction:
        direction = direction.strip()
        if direction not in direction_names():
            raise SystemExit(
                f"错误：未知方向 {direction!r}。可选方向：{'、'.join(direction_names())}"
                "（可在 规则/项目库配置.yaml 的 directions 里增改）"
            )
        candidate = root / direction / project
        return candidate if candidate.is_dir() else None

    hits = [root / d / project for d in direction_names() if (root / d / project).is_dir()]
    if len(hits) > 1:
        raise SystemExit(
            "错误：同名项目目录在多个方向下都存在，必须用 --direction 指定：\n  "
            + "\n  ".join(str(p) for p in hits)
        )
    if hits:
        return hits[0]
    legacy = root / project
    if legacy.is_dir():
        return legacy
    return None


def resolve_new_project_dir(root: Path, project: str, direction: str | None = None) -> Path:
    """新建项目时的目标目录：<root>\\<方向>\\<项目名>。"""
    direction = (direction or default_direction()).strip()
    if direction not in direction_names():
        raise SystemExit(
            f"错误：未知方向 {direction!r}。可选方向：{'、'.join(direction_names())}"
        )
    return root / direction / project


def setup_stdout() -> None:
    """Windows 控制台默认 GBK，统一改成 UTF-8，避免中文乱码。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass


def read_text(path: os.PathLike | str) -> str:
    return Path(path).read_text(encoding="utf-8-sig")


def write_text(path: os.PathLike | str, text: str) -> None:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    with target.open("w", encoding="utf-8", newline="\n") as fh:
        fh.write(text)


def load_yaml_file(path: os.PathLike | str) -> dict:
    try:
        data = parse(read_text(path))
    except YamlError as exc:
        raise SystemExit(f"错误：规则文件解析失败 {path}：{exc}")
    if not isinstance(data, dict):
        raise SystemExit(f"错误：规则文件顶层必须是映射：{path}")
    return data


def resolve_root(arg_root: str | None) -> Path:
    """根目录优先级：--root > 环境变量 CARDO_LIB_ROOT > 项目库配置.yaml。"""
    if arg_root:
        return Path(arg_root).expanduser().resolve()
    env_root = os.environ.get("CARDO_LIB_ROOT", "").strip()
    if env_root:
        return Path(env_root).expanduser().resolve()
    if LIB_CONFIG_FILE.exists():
        cfg = load_yaml_file(LIB_CONFIG_FILE)
        configured = cfg.get("project_root")
        if isinstance(configured, str) and configured.strip():
            return Path(configured.strip()).expanduser().resolve()
    raise SystemExit(
        "错误：未指定项目库根目录。请传 --root \"<标准项目库根>\"，"
        "或设置环境变量 CARDO_LIB_ROOT，或在 规则/项目库配置.yaml 里填 project_root。"
    )


def resolve_project_name(arg_project: str | None) -> str:
    if arg_project and arg_project.strip():
        name = arg_project.strip()
        if any(sep in name for sep in ("/", "\\")):
            raise SystemExit(f"错误：--project 只能是项目文件夹名，不能含路径分隔符：{name!r}")
        return name
    if LIB_CONFIG_FILE.exists():
        cfg = load_yaml_file(LIB_CONFIG_FILE)
        configured = cfg.get("default_project")
        if isinstance(configured, str) and configured.strip():
            return configured.strip()
    raise SystemExit("错误：未指定项目名。请传 --project \"<项目名>\"。")


# ---------------------------------------------------------------- T20 用户目录模式
# 约定（所有脚本一致）：
#   --project-dir <路径>  最高优先级，直接指向用户自选项目目录（任意路径，可含空格/中文）。
#   --root/--project      只用于校准集/回归（标准项目库布局），行为与 T20 前完全一致。
#   两者都不给：当前目录像项目（00_项目索引.md 或 .cardo/）就用它；否则 exit 2，不猜标准项目库。


def is_project_dir(path: os.PathLike | str) -> bool:
    """目录是否“像项目”：有 00_项目索引.md 或 .cardo/（不扫描、不读正文）。"""
    base = Path(path)
    return (base / INDEX_NAME).is_file() or (base / CARD_DIR).is_dir()


def read_card(project_dir: os.PathLike | str) -> dict:
    """只读读取 .cardo/project.json（T20 不写、不生成该文件；坏文件/非映射一律当没有）。"""
    path = Path(project_dir) / CARD_DIR / CARD_FILE
    if not path.is_file():
        return {}
    try:
        data = json.loads(read_text(path))
    except (OSError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def card_direction(project_dir: os.PathLike | str) -> str:
    """.cardo/project.json 里的方向（只读；缺失/非法返回空串，由调用方决定是否问用户）。"""
    value = read_card(project_dir).get("direction")
    return value.strip() if isinstance(value, str) else ""


def canon_direction(raw: object) -> str | None:
    """方向名归一：接受 `老旧街区项目` 或短名 `老旧街区`；不在配置名单里返回 None。"""
    text = str(raw or "").strip()
    if not text or text in ("待确认", "—", "-"):
        return None
    names = direction_names()
    if text in names:
        return text
    for candidate in (text + "项目", text[:-2] if text.endswith("项目") else ""):
        if candidate and candidate in names:
            return candidate
    return None


def resolve_user_project_dir(arg_project_dir: object, *, require_exists: bool = True) -> Path:
    """T20 统一入口：返回用户项目目录（绝对路径）。

    - 传了 --project-dir：直接用它；不存在且 require_exists=True → 报错（exit 2）。
    - 没传：当前目录像项目（有索引或 .cardo/）就用当前目录。
    - 都没有：报错并提示传 --project-dir；**不猜标准项目库、不静默指旧路径**。
    """
    raw = str(arg_project_dir or "").strip()
    if raw:
        path = Path(raw).expanduser()
        try:
            path = path.resolve()
        except OSError:  # pragma: no cover - 极少数无效路径
            path = Path(os.path.abspath(raw))
        if path.exists() and not path.is_dir():
            raise UserProjectError(f"--project-dir 不是目录：{disp_path(path)}")
        if require_exists and not path.is_dir():
            raise UserProjectError(
                f"--project-dir 不存在或不是目录：{disp_path(path)}（请指向用户自选的项目目录）")
        return path
    cwd = Path.cwd()
    if is_project_dir(cwd):
        return cwd
    raise UserProjectError(
        "未指定项目目录：请用 --project-dir <项目目录> 指向用户自选的项目文件夹；"
        "（--root/--project 只用于标准项目库的校准/回归，不再自动猜库）")


def disp_path(path: os.PathLike | str, base: os.PathLike | str | None = None) -> str:
    """控制台展示用路径：尽量相对当前目录（不打印盘符/用户目录）。"""
    try:
        rel = os.path.relpath(str(path), str(base or Path.cwd()))
    except (OSError, ValueError):  # pragma: no cover - 跨盘等
        return Path(path).name or "."
    return Path(rel).as_posix()


def portable_display(path: os.PathLike | str) -> str:
    """产物里引用 skill 外部资源时的占位展示：skill 内 → <技能目录>/…；其它 → 文件名。"""
    target = Path(path)
    try:
        return "<技能目录>/" + target.resolve().relative_to(SKILL_DIR).as_posix()
    except (OSError, ValueError):
        return target.name or "."


def gate(lines: list[str], exit_code: int = GATE_EXIT) -> int:
    """打印闸门提示并返回退出码（调用方直接 sys.exit(返回值)）。"""
    print("")
    print("[闸门] 以下情况不确定，必须先停下来问用户，不能自动继续：")
    for line in lines:
        print(f"  - {line}")
    if exit_code:
        print("")
        print(f"（退出码 {exit_code}；确认后再执行下一步，删除/覆盖/移动原始文件需用户明确授权）")
    return exit_code


def now_str() -> str:
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def today_str() -> str:
    return datetime.datetime.now().strftime("%Y-%m-%d")


def format_size(size: int) -> str:
    value = float(size)
    for unit in ("B", "KB", "MB", "GB"):
        if value < 1024 or unit == "GB":
            if unit == "B":
                return f"{int(value)} B"
            return f"{value:.1f} {unit}"
        value /= 1024.0
    return f"{size} B"


def sha256_file(path: os.PathLike | str, chunk: int = 1 << 20) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as fh:
        while True:
            block = fh.read(chunk)
            if not block:
                break
            digest.update(block)
    return digest.hexdigest()


def rel_posix(path: Path, base: Path) -> str:
    return path.relative_to(base).as_posix()


def real_key(path: os.PathLike | str) -> str:
    """目录的"真实身份"：解析联接/符号链接后的规范路径，用于识别别名。

    Windows 的目录联接（junction）、符号链接、以及大小写差异，都会让同一个物理目录
    出现多个名字。台账与投递区判定必须按真实身份去重，否则会把"别名"误判成
    "两个不同的投递区"而触发闸门。
    """
    try:
        resolved = os.path.realpath(str(path))
    except OSError:
        resolved = os.path.abspath(str(path))
    return os.path.normcase(resolved)


def group_name_aliases(project_dir: Path, names: list[str]) -> list[list[str]]:
    """把项目目录下的一组候选目录名按真实身份分组，保持传入顺序。

    返回如 [[原始提供输入输出, 01_原始输入]] 表示这两个名字指向同一个目录；
    返回多个子列表则表示确实是不同的目录（调用方应按闸门处理）。
    """
    groups: dict[str, list[str]] = {}
    order: list[str] = []
    for name in names:
        key = real_key(project_dir / name)
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(name)
    return [groups[k] for k in order]