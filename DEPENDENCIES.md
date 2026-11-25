# 运行环境与依赖

## Python路径：校验、分析、打包

`run.py verify`、`analyze`、`pack` 与 `src/validate_workbook.py` 只使用Python标准库，不需要联网、不需要pip安装，也不需要Excel或WPS在后台运行。

建议Python 3.10或以上。本次实际环境为Linux、Python 3.12.14。Python脚本按跨平台路径编写，但Windows CMD包装器及Windows本机执行本次未实测。

## Excel重建路径

`src/rebuild.mjs` 使用 `@oai/artifact-tool` 进行工作簿导入、公式重算和导出。没有用openpyxl或xlsxwriter替代作者引擎。Python元数据步骤只保留原打印/字体属性及补齐图表缓存，不改学生值或公式。

本次环境检测：

| 项目 | 实际版本 |
| --- | --- |
| Node.js | v24.22.3 |
| @oai/artifact-tool | 2.8.169 |
| 主运行环境包 | 26.1117.11355 |
| Python | 3.12.14 |

默认优先使用当前工作环境提供的 `CODEX_PRIMARY_RUNTIME_NODE` 和 `CODEX_PRIMARY_RUNTIME_NODE_MODULES`，不硬编码原会话目录。对应变量未设置时，可用 `GRADE4_NODE` 和 `GRADE4_NODE_MODULES` 指向已安装的兼容环境；若要切换环境，应先取消相应的CODEX变量。

工程包没有复制整套Node运行时或node_modules，也没有把私有工具包伪称为公共npm依赖。新环境没有这个引擎时，`doctor` 会显示不能重建；此时仍能打开成品Excel、校验文件、提取JSON/CSV并独立分析。要完整自动重建，应在有该表格环境的ChatGPT工作区运行。

`rebuild` 在系统临时目录中链接宿主Node依赖，结束后清理，不在工程包内保存符号链接。完整重建的实际测试环境是Linux。

## 字体与资源

工作簿原有宋体、Arial等字体名称保留，原生图表及公式均已包含在XLSX内，没有外部图片链接需要另行下载。数值计算、文件校验和CSV提取不依赖字体。

字体文件未嵌入本包。Windows上的具体字形取决于本机已安装字体；PNG预览使用当前运行环境可用字体，新环境可能有字体替换或细微排版差异。这不影响原始分数及公式数据。不能把跨机器逐像素一致当作已经验证的结果。

## 离线与数据范围

运行脚本不发送学生信息、不调用搜索、不依赖账号令牌。包内只含本对话四年级成绩项目及同一工作簿内原有年级数据，不包含其他对话的工程或后台任务。
