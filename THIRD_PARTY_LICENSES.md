# 第三方依赖与许可证清单（cardo-status）

只列本包 `requirements.txt` 实际声明的依赖。

| 依赖 | 用途 | 许可证 |
|---|---|---|
| python-docx | 报告 docx 生成 | MIT |
| openpyxl | 台账与 xlsx 读取 | MIT |
| PyYAML | 规则文件解析 | MIT |
| pymupdf | PDF 采样与有文本层 PDF 全文抽取 | **AGPL-3.0 / Artifex 商业双许可** |
| pypdf | pymupdf 不可用时的轻量兜底 | BSD-3-Clause |
| rapidocr-onnxruntime | 扫描页离线 OCR（结果只进归类依据） | Apache-2.0 |
| opencv-python-headless | 图像预处理 | Apache-2.0 |
| onnxruntime | OCR 推理运行时 | MIT |
| numpy | 数值计算 | BSD-3-Clause |

## 说明

- **PyMuPDF 是本包、也是四个包里唯一的 AGPL 依赖**；公开分发含 PyMuPDF 的组合按 AGPL 处理最稳妥，
  所以本仓库采用 AGPL-3.0-or-later。闭源商用需向 Artifex 取得商业许可。
- 网络提供服务时，须按 AGPL 第 13 条向使用者提供对应源码。
- 本包不包含任何第三方二进制、wheel 或 Python 运行时；依赖由使用者从 PyPI 安装。
- 不调用 OSGB → 3D Tiles 转换器。
