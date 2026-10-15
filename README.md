# 序衡 本地报表智能排序对齐

Python Windows 桌面应用。导入参考表和待排序表，自动识别编号列，按参考顺序重排完整业务行，导出新 Excel 和逐行审计文件。运行时不联网，也不需要 API Key。

## 快速运行

完整源码 ZIP 附带 `Windows/Xuheng.exe`。解压整个 ZIP，进入 Windows 目录双击运行，无需安装 Python；请保留同目录的 `_internal` 与 `samples`。打包为 Windows x64，已在本轮 Windows 11 启动验证。

源码运行需 Python 3.10+（开发验证使用 3.12.14，Tkinter 需包含在 Python 安装中）：

```powershell
py -3 -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python main.py
```

也可运行 `install.cmd` 后运行 `start.cmd`。命令行与桌面版使用同一核心：

```powershell
py -3 main.py --reference samples/模拟参考.xlsx --source samples/模拟待排序.xlsx --output output/对齐成果.xlsx
```

## 实际操作

1. 两个文件输入框分别选择参考文件、待排序文件；也可分别点“粘贴文本”，粘贴 Excel 区域、CSV 或单列编号。
2. 点击“导入并识别”。程序显示自动选择的工作表、列号和表头行；原表页签展示来源顺序。
3. 默认“未匹配置尾”，也可选“未匹配保留原位”。点“开始对齐”。
4. 核对结果中的原始行、输出行和全部参考候选。黄色为多候选，橙色为未匹配；多候选已包含在匹配成功数量内。
5. 在“保存到”输入本地路径，点击“导出成果”。自动创建目标目录，生成新 Excel、同名 `.audit.csv` 和 `.audit.json`。

“填入模拟样例”只填文件路径；真正读取发生在“导入并识别”，真正计算发生在“开始对齐”。样例清楚标注为模拟业务数据，本轮演示实际结果为 8 行、7 行匹配、1 行未匹配、1 行多候选。

## 匹配与文件规则

- 方向：**待排序编号包含在参考编号中**。完全匹配优先，否则单向子串匹配。空编号不匹配。
- NFKC 全半角、大小写、首尾空白归一；内部字符保留，不擅自删除 `-`、空格或修订号。
- 一对多列出所有参考行，排序使用最早位置；多对一稳定排序，不复制业务行。
- 默认未匹配项置尾、保留相对顺序。`--policy fixed` 固定未匹配位置，其他匹配项填回其余行槽。
- XLSX 保留原有列结构、样式引用、行高、列宽和其他工作簿部件；只重排选定工作表的行。审计作为独立文件，不向业务表追加列。
- XLS 值型工作簿采用 xlrd/xlutils/xlwt，输出仍为 `.xls`。公式或数据区跨行合并会友好阻止导出，避免无提示损坏；含复杂嵌入对象的旧 XLS 应先在 Excel 内转为值型 XLSX。
- XLSX 含移动区公式、跨行合并时也会明确提示先转换或拆分。保留其他部件字节并不等于已验证所有图表、批注、超链接与外部链接在行重排后的业务语义；建议值型明细作为排序输入。
- 不覆盖来源文件；日志保存在 `%LOCALAPPDATA%/Xuheng/logs`，按月追加、不自动删除，保留策略不少于 180 天。

## 测试与真实记录

```powershell
py -3 -m unittest discover -s tests -v
$env:PYTHONPATH = (Get-Location).Path
py -3 scripts/benchmark.py
py -3 scripts/soak.py --seconds 259200
```

本轮记录在 `validation/`：9 项独立预期与文件完整性回归通过；10 组附件实际输入共 1,018 行，逐行值、顺序规则和来源哈希检查通过；1,000 行固定随机种子模拟数据 5 次完整流程中位 0.291 秒；180.05 秒循环完成 3,477 次，无异常。短时循环不等于 72 小时验证。

附件没有官方独立预期输出，因此不将匹配行数当作官方匹配准确率，也不声称已经满足 98%/95% 的官方准确率门槛。千行计时描述本轮机器，不替代题目指定硬件环境。Windows 10 未在本机实测。

附件重跑需要原始授权附件置于 `.private/报表智能排序对齐算法/数据/`，然后运行 `scripts/validate_attachments.py`。原始业务资料不会公开；公共汇总仅使用样本序号，不含文件名与编号内容。

## 目录与版本

- `aligner/`：桌面界面、读取识别、匹配排序、保真导出与本地审计。
- `samples/`：模拟 Excel 与文本样例。
- `tests/`、`scripts/`：测试、计时、耐久、打包与交付工具。
- `docs/`：架构、需求验收映射和公开版申报书。
- `demo/`：本轮导出成果、中文解说词、SRT/ASS 字幕、视频实测记录。
- `validation/`：实际运行的汇总证据。

正式交付版本为 `1.0.0`。视频、申报书和附带 Windows 程序通过同一 Release 交付。`build_windows.cmd` 可重建 Windows 程序；`requirements-dev.txt` 包含开发交付工具，应用本身仅需 `requirements.txt`。

公开版本不包含参赛人的联系电话、完整身份资料、附件业务原件或密钥。本地完整版申报书单独保留。演示语音通过神经语音服务生成，仅发送公开旁白文字；应用处理图纸资料不使用该服务。

## 开源依赖

openpyxl（MIT）、xlrd（BSD）、xlwt（BSD）、xlutils（MIT）；Python/Tkinter 遵循对应发行许可。第三方许可清单见 `THIRD_PARTY_NOTICES.md`。本工程采用 MIT 许可。
