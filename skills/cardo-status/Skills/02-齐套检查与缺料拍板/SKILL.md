# 02 齐套检查与缺料拍板

## 用途
按 `规则/必备材料清单.yaml` 检查材料齐套度，输出缺料表；生成《输入材料补齐确认单》，
并把“先出稿/等材料”的拍板结果留痕。

## 单独运行
```powershell
python Skills/02-齐套检查与缺料拍板/scripts/check_materials.py ^
  --project-dir . --direction 老旧街区 --out 过程文件/02-齐套检查与缺料拍板/齐套检查

python Skills/02-齐套检查与缺料拍板/scripts/make_material_request.py ^
  --gap-json 过程文件/02-齐套检查与缺料拍板/齐套检查/材料齐套表.json ^
  --out 成果文件/输入材料补齐确认单.docx ^
  --md 过程文件/02-齐套检查与缺料拍板/补齐确认单.md ^
  --classification 过程文件/01-建档与材料归集/分类清单/分类清单.json ^
  --report-name "<项目名>现状分析报告"

python Skills/02-齐套检查与缺料拍板/scripts/export_gap_xlsx.py ^
  --materials 过程文件/02-齐套检查与缺料拍板/齐套检查/材料齐套表.json ^
  --chapters 过程文件/06-分章成文/分章初稿 ^
  --facts 过程文件/05-事实抽取与事实库/facts/事实库.json ^
  --out 成果文件/数据缺口清单.xlsx
```

## 输出
- `齐套检查/材料齐套表.json`（+ 可读表）。
- `成果文件/输入材料补齐确认单.docx`、`过程文件/.../补齐确认单.md`。
- `过程文件/02-齐套检查与缺料拍板/缺料拍板.json`（hasGap、choice、reason、时间）。
- 09 阶段调用 `export_gap_xlsx.py` 生成三张表：材料缺口 / 章节缺项 / 事实缺失。

## 退出码
- `check_materials.py`：0 齐套；1 存在缺料；其他非 0/1 视为执行失败。
- `make_material_request.py` / `export_gap_xlsx.py`：0 成功，非 0 失败。
- 缺料不阻断流程，但必须显式拍板：pipeline 写 `缺料拍板.json`，选择由
  `run --material-choice 先出稿`（缺省，缺口版继续）或 `先补料`（exit 3 停下等补齐）决定。

## 规则要点
- 缺料清单必须逐项可追溯（材料类别、用途、影响章节）。
- 不把“缺料”写成“无问题”，不静默继续。
