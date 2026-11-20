# 07 表 3.1 口径计算

## 用途
按 `规则/计算口径.yaml` 计算现状用地分类面积与占比（表 3.1），输出可审计的计算轨迹。
缺数一律走缺项，不补数、不猜数。

## 单独运行
```powershell
python Skills/07-表3.1口径计算/scripts/build_calc_input.py ^
  --facts 过程文件/05-事实抽取与事实库/facts/事实库.json ^
  --out 过程文件/07-表3.1口径计算/计算输入.json

python Skills/07-表3.1口径计算/scripts/calc_metrics.py ^
  --project-dir . --input 过程文件/07-表3.1口径计算/计算输入.json ^
  --out 过程文件/07-表3.1口径计算 --format text

python Skills/07-表3.1口径计算/scripts/build_table31.py ^
  --trace 过程文件/07-表3.1口径计算/指标计算-<日期>.json ^
  --out 过程文件/07-表3.1口径计算/表3.1.md ^
  --title "表 3.1 现状用地分类面积与占比"
```

## 输出
- `计算输入.json`（cells 来自事实库“用地+面积”事实；无事实给 `value:null` 占位并标 fallback）。
- `指标计算-<日期>.json / .md`：每指标 status（ok/gap）、缺数说明、校验结果。
- `表3.1.md`：可并入正文的表。

## 退出码
0 全部算通；1 存在缺数（gap，不是错误）；3 硬门禁（除零、单位不一致、占比超差）；2 输入错误。

## 口径要点
- 占比合计≈100%、合计=逐行之和、单位统一，任一不满足即记录而不是修数。
- 表头与 `表格契约.yaml` TBL-01 一致（费用构成表口径不同，不混用）。
