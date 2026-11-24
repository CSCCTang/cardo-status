# 更新记录（cardo-status）

## v0.1.2 · 2026-11-24

- 修复《06-分章成文》编制依据排序：`basis_library.py` 原按固定层级字典 `{"国家","某省","某市","项目"}` 取权重，而依据库实际登记的是具体地名（如“四川省”“成都市”），导致全部落到兜底值、层级排序失效。改为查不到固定层级时按行政区后缀兜底（省/自治区→1、市→2、其他→9），库内原值不改写。
- 版本号同步至 0.1.2（`package.json`、`scripts/cardo.py` 的 `VERSION`、README 安装与版本行）。

## 2026-11-24 · 文档

- 新增本更新记录，按 tag 补齐 v0.1.0、v0.1.1 的变更说明。
- `docs/03_常见失败.md` 补充环境自检：`--doctor` 查交付包与运行时，`--where` 打印交付包目录。

## v0.1.1 · 2026-11-21

- 启动器 `scripts/cardo.py`：定位顺序改为「技能目录自带包 > `--package` > `CARDO_PACKAGE` > `CARDO_HOME` > 当前目录及上三级」，不再依赖固定的 `CardoStatus\` 相对路径；新增 `--where`（打印交付包目录）与 `--version`。
- 首次调用先按约定建工作区目录架构（幂等，已存在不覆盖）并逐条说明每个目录放什么；工作区没有输入且未给流水线参数时，只建目录、停下等料。
- 工作区越界保护：`--workspace` 指向插件安装目录或 `DSH_HOME` 时拒绝并给出可用路径；未指定且当前目录越界时，改用默认工作区 `%USERPROFILE%\Cardo工作区\cardo-status-<时间戳>`（根目录可用 `CARDO_WORKSPACE_HOME` 覆盖）。`--doctor` 一并打印默认工作区与目录架构。
- `SKILL.md`、`pipeline.py`、`check_chapter.py`、`开工.py` 里的命令统一写成 `python "<技能目录>/..."`，避免按插件安装位置猜路径。
- `index.js` 注册技能时补 `source: 'custom'`。

## v0.1.0 · 2026-11-20

- 首次发布：Host 入口 `index.js` 与 `cordis.patch.yml`，内嵌 `skills/cardo-status/` 技能目录。
- 流水线按 00–09 分段（00 总控、01 建档与材料归集、02 齐套检查与缺料拍板、03 模板与源报告转储、04 章节映射、05 事实抽取与事实库、06 分章成文、07 表 3.1 口径计算、08 组装与模板渲染、09 三重自检与边界），06 再分 A–F 章；各段带 `SKILL.md`、`scripts/` 与 `规则/`。
- 附依据库、报告模板、地名词表等资产，`docs/` 使用流程、输入准备、常见失败三篇，以及 LICENSE（AGPL-3.0-or-later）、NOTICE、第三方许可清单与第三方组件说明。
- 交付形态：`package.json` + `requirements.txt`；GitHub Actions 在打 `v*` tag 时执行 `pnpm pack`、生成 SHA256SUMS 并建草稿 Release。
