# 04 章节映射

## 用途
把模板的 23 个章节节点映射到原始报告标题：精确/互含 = high，领域同义词 = medium，
其余 = gap。medium 只是「提案」，必须在 `人工确认.yaml` 逐条拍板后才算落位；
未拍板 → exit 3，不静默采用。gap 是合法缺项结论，不阻塞。

## 单独运行
```powershell
python Skills/04-章节映射/scripts/make_mapping_skeleton.py ^
  --template ..\提资数据\模板\<模板>.docx --source 输入文件/<源报告>.docx ^
  --out 过程文件/04-章节映射/_骨架.yaml

python Skills/04-章节映射/scripts/fill_mapping.py ^
  --skeleton 过程文件/04-章节映射/_骨架.yaml ^
  --source-outline 过程文件/03-模板与源报告转储/源报告转储/<源报告>.outline.json ^
  --out 过程文件/04-章节映射/章节映射.yaml ^
  --confirm 过程文件/04-章节映射/人工确认.yaml ^
  --direction 老旧街区

python Skills/04-章节映射/scripts/make_mapping_skeleton.py --check 过程文件/04-章节映射/章节映射.yaml ^
  --template ..\提资数据\模板\<模板>.docx --source 输入文件/<源报告>.docx
```

## 输出
- `章节映射.yaml`：每节点 `status`（high/medium/gap）、`sources`、`confidence`、`notes`。
- `映射统计.json`：exact/alias/gap 计数、`pending_confirm`（待复核提案）、`mediumNodes`。
- `人工确认.yaml`（人工编写）：`nodes[].id` + `sources`。`sources` 非空 = 确认落位；
  `sources: []` = 复核后判为 gap（不作推断）；`note` 写复核理由。

## 退出码
0 完成；3 存在中等置信提案待人工复核；1 自检未通过；2 参数错误。

## 规则要点
- 归一化后精确相等 或 互相包含 → high。
- `ALIASES` 领域同义词（如“现状分析/现状详细分析”）→ medium（提案，待人复核）。
- 缺口比例 = gap / 总节点；只作提示，**不再自动进入缺口版**（v2 默认干净版，`--edition gap` 才留痕）。
- 映射结果必须能被内置 `_yaml_lite` 解析：输出使用缩进块序列（`IndentDumper`）。
