# 09 三重自检与边界

## 用途
独立验证交付物：章节契约、事实回溯、边界（原件只读/模板未替换/无绝对路径/交付物齐套），
并输出运行指标、验证报告与验证基线。

## 单独运行
```powershell
python Skills/09-三重自检与边界/scripts/build_trace.py ^
  --project-dir . --artifact 过程文件/08-组装与模板渲染/正文.md ^
  --facts 过程文件/05-事实抽取与事实库/facts/事实库.json ^
  --out 过程文件/09-三重自检与边界/回溯/回溯表.json --baseline-version 1

python Skills/09-三重自检与边界/scripts/check_trace.py ^
  --project-dir . --artifact 过程文件/09-三重自检与边界/回溯/回溯表.json ^
  --facts 过程文件/05-事实抽取与事实库/facts/事实库.json ^
  --out 过程文件/09-三重自检与边界/回溯检查

python Skills/09-三重自检与边界/scripts/check_contracts.py ^
  --project-dir . --direction 老旧街区 ^
  --template ..\提资数据\模板\<模板>.docx ^
  --report 过程文件/08-组装与模板渲染/正文.md ^
  --docx 成果文件/<报告>.docx ^
  --contracts Skills/09-三重自检与边界/规则 ^
  --source-dump 过程文件/03-模板与源报告转储/源报告转储 ^
  --out 过程文件/09-三重自检与边界/契约检查

python Skills/09-三重自检与边界/scripts/boundary_check.py ^
  --project-dir . --fingerprint 过程文件/01-建档与材料归集/原件指纹.json ^
  --template ..\提资数据\模板\<模板>.docx ^
  --contract Skills/09-三重自检与边界/规则/章节契约.yaml ^
  --inputs 输入文件 --outputs 成果文件 ^
  --deliverables 成果文件/<报告>.docx 成果文件/输入材料补齐确认单.docx 成果文件/数据缺口清单.xlsx ^
  --report 过程文件/08-组装与模板渲染/正文.md [--gap-mode] ^
  --out 过程文件/09-三重自检与边界/边界检查.json

python Skills/09-三重自检与边界/scripts/baseline_check.py ^
  --baseline 记忆与规则/验证基线-CardoStatus.json ^
  --metrics 过程文件/09-三重自检与边界/运行指标.json

# 依据三件套校验（只校验不写产物；导出在 08 的 export_evidence_pack.py）
python Skills/09-三重自检与边界/scripts/check_evidence.py --project-dir .
```

## 输出
- `回溯/回溯表.json`（schema `cardo/trace/v1`，只存 factId/定位引用，不复制数值）、
  `回溯检查/`（覆盖率、失败项）。
- `契约检查/契约检查-<日期>.md` + `契约检查/契约检查.json`：fail/hint/skip/manual 计数与逐项差异。
- `边界检查.json`：原件 sha、模板 sha、技能绝对路径、输入/成果扩展名白名单、交付物、缺口版标注。
- `依据校验_<节>.md`：标注源逐段依据块校验（正文洁净 / 段落覆盖 / 数值有源 / 依据有效 / 剥离一致 / 正文非表格化 / 对照表回指 / 链条闭合）。成稿口径下正文 0 处「（数据来源：…）」行，出处由〔锚点〕+《依据对照表》承载；三件套导不出标注源时按失败计（exit 1），不再静默跳过。
- `运行指标.json`、`验证报告.md`、`记忆与规则/验证基线-CardoStatus.json`。

## 退出码
0 通过；1 存在契约硬失败/回溯失败/边界失败；2 输入错误。

## 判定口径
- 完整版验收：契约 fail=0、回溯覆盖率 100%、边界 PASS、三件交付物齐套。
- 缺口版：契约 fail 允许 >0，但必须逐项列入验证报告与数据缺口清单；边界仍须 PASS。
- 章内自检的提示级 FAIL/WARN 不改变 09 的契约判定，但必须在验证报告中说明。

## 规则文件
`规则/章节契约.yaml`、`表格契约.yaml`、`写作规范.yaml`、`计算口径.yaml`、`质量评分表.yaml`、
`回溯表.schema.json`、`chapter_skills.yaml`（均为上游 cardo-status-spec 的只读副本）。
