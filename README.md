# cardo-status

> 安装（二选一）：
> `dsh plugin --profile <profile> add github:CSCCTang/cardo-status#v0.1.1`
> `dsh plugin --profile <profile> add <本地路径>\cardo-status-0.1.1.tgz`

现状分析：给一批项目资料与现状底稿，产出《现状分析报告》《输入材料补齐确认单》《数据缺口清单》。

## 需要提供什么

| 输入 | 形态 | 最小样例 |
|---|---|---|
| 项目资料 | 批复、规划、立项等，放 `输入文件\` 任意层级 | 本包不附样例；有几份放几份 |
| 现状底稿 | docx / xlsx（甲方或项目组调研底稿） | 一份现状底稿 docx 即可起步 |
| 分析范围答复 | 运行时会问；也可一次答完（见下） | — |

支持 docx / xlsx / pdf（扫描件走离线识别，识别不可用时按文件名归类）。

## 装完怎么调

在 dsh 对话里说「用 cardo-status 跑，工作区是 `D:\proj\status-01`，分析范围是 XX 片区」。

工作区是独立空文件夹：`输入文件\` 放料、`过程文件\` 留痕、`成果文件\` 出活。

## 会得到什么

`成果文件\`：

- 《现状分析报告》（docx）
- 《输入材料补齐确认单》
- 《数据缺口清单》

## 许可证与版本

- 版本 `0.1.1`；许可证 **AGPL-3.0-or-later**，见 [LICENSE](LICENSE) 与 [NOTICE](NOTICE)。
- 第三方组件与许可证见 [THIRD_PARTY_LICENSES.md](THIRD_PARTY_LICENSES.md)。
- 运行环境：Windows；Python ≥ 3.11；dsh ≥ 0.2.2-rc.5。
- 更新记录见 [CHANGELOG.md](CHANGELOG.md)（最近更新 2026-11-24）。

## 文档

- [docs/01_使用流程.md](docs/01_使用流程.md)
- [docs/02_输入准备.md](docs/02_输入准备.md)
- [docs/03_常见失败.md](docs/03_常见失败.md)
