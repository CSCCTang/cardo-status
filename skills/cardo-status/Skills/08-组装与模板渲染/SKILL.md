# 08 组装与模板渲染

## 用途
把 A–F 六章组装成连续正文（缺口版另加「附件四 数据缺口索引」一页），再按标准模板渲染 docx，
并回读校验标题、层级、表格、样式与东亚字体。

## 单独运行
```powershell
python Skills/08-组装与模板渲染/scripts/assemble_chapters.py ^
  --extract-dir 过程文件/06-分章成文/分章初稿 ^
  --out 过程文件/08-组装与模板渲染/正文.md ^
  --mapping-out 过程文件/08-组装与模板渲染/来源映射.md ^
  --gap-heading "附件四 数据缺口索引" --gap-index-only --append-to-chapter F

python Skills/08-组装与模板渲染/scripts/export_evidence_pack.py ^
  --project-dir . --out-dir 成果文件/证据包

python Skills/08-组装与模板渲染/scripts/build_report_docx.py ^
  --template ..\提资数据\模板\<模板>.docx ^
  --content 过程文件/08-组装与模板渲染/正文.md ^
  --out 成果文件/<项目名>现状分析报告.docx ^
  --title "<项目名>现状分析报告" --force
```

## 输出
- `正文.md`（H1=6 / H2=30）、`来源映射.md`。
- `成果文件/*现状分析报告.docx`。
- `成果文件/证据包/`：干净版 docx / 标注版 docx / 依据对照表 xlsx（同源生成，干净版从标注版剥离）。
- 终端 `[verify] ok` 系列：段落样式、H1/H2 数、表格数/形状、表样式、Title、页面设置、eastAsia、样式保留。

## 退出码
0 渲染且回读校验通过；1 组装失败/回读校验失败；2 参数错误。

## 门禁
- 标题、H1/H2 必须与模板逐字一致；表格必须落在对应节内。
- **默认干净版（clean）**：正文不含任何缺口提示、不追加附件四；缺口另存
  `过程文件/08-组装与模板渲染/数据缺口汇总.md`，并导出《数据缺口清单.xlsx》。
- 只有显式 `--edition gap` 才追加「附件四 数据缺口索引」（指向 xlsx，不造完整表）与正文开头的缺口说明，
  且须至少登记一个缺项。**成果文件名统一为《XX现状分析报告.docx》**，不带「（缺口版）」——
  版本体现在运行元数据（`运行状态.json` 的 `edition` 字段）。
- 附件一~四须齐全（来自 F 章分章稿）；缺失时只记入 `附件齐备情况.json`，干净版正文不提示。
- `[verify] ok` 缺失即阻断（不能只看进程退出码）。

## 已知限制
- 表格由 markdown 转 docx 表格，样式为 Table Grid；不继承模板中的复杂表格样式。
- 正文长度受 06 的章级裁剪约束。
