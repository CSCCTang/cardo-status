# 03 模板与源报告转储

## 用途
把标准模板与原始报告转成机器可读的 `outline.json / blocks.txt / tables.txt`，
并用 `章节契约.yaml` 校验模板指纹与节点数，产出章节骨架。

## 单独运行
```powershell
python Skills/03-模板与源报告转储/scripts/dump_docx_outline.py ..\提资数据\模板\<模板>.docx ^
  --out 过程文件/03-模板与源报告转储/模板转储 --tables
python Skills/03-模板与源报告转储/scripts/dump_docx_outline.py 输入文件/<源报告>.docx ^
  --out 过程文件/03-模板与源报告转储/源报告转储 --tables

python Skills/03-模板与源报告转储/scripts/build_skeleton.py ^
  --template ..\提资数据\模板\<模板>.docx ^
  --template-outline 过程文件/03-模板与源报告转储/模板转储/<模板>.outline.json ^
  --source-outline 过程文件/03-模板与源报告转储/源报告转储/<源报告>.outline.json ^
  --contract Skills/03-模板与源报告转储/规则/章节契约.yaml ^
  --out 过程文件/03-模板与源报告转储
```

## 输出
- `模板转储/*.outline.json`、`*.blocks.txt`、`*.tables.txt`。
- `源报告转储/*` 同上（事实与成文的唯一取料源）。
- `章节骨架.json`：模板节点（H1=6 / H2=17，共 23）与源报告标题清单。

## 退出码
0 通过；1 模板指纹不符或节点数不对（旧模板/误用方向模板）；2 参数错误。

## 门禁
- 模板 sha256 必须与 `章节契约.yaml` 的 `template.sha256` 一致。
- 节点数必须等于契约 `nodes` 数量；否则明确报“可能是旧模板”。
- 转储只读，不改源 docx。
