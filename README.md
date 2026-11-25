# 四年级期中成绩核查续作工程

打包日期：2026-11-20。成绩所属月份：2026年11月。原分析日期：2026-11-06。

## 先打开哪个文件

当前完整、可编辑的最新工作簿在：

`inputs/2026.11期中考试四年级名单.xlsx`

它是已经完善的七工作表版本：一年级、二年级、三年级、四年级、五年级、四年级分析、数据核查。`inputs` 在这里表示“继续工作的基线”，不是声称它是最初上传的五工作表原件。

继续新对话时，请把整个ZIP上传，并让新对话先读 `START_HERE.txt`、`HANDOFF.md` 和 `STATUS.json`。不需要旧对话的临时目录、旧下载链接或旧 `/tmp` 文件。

## 三条常用命令

先解压ZIP，进入包含本README和 `run.py` 的目录。

```bash
python run.py verify
python run.py analyze
python run.py rebuild --render
```

部分系统使用 `python3`，Windows也可使用 `py -3`。提供了 `run_linux.sh` 和 `run_windows.cmd` 包装入口，例如 `run_windows.cmd verify`。

| 命令 | 用途 | 环境要求 |
| --- | --- | --- |
| `verify` | 按MANIFEST逐文件检查SHA-256及大小，并检查基线XLSX的ZIP校验 | Python标准库 |
| `analyze` | 从当前工作簿重新提取学生记录、独立复算、导出JSON和CSV | Python标准库 |
| `doctor` | 查看Python、Node及表格引擎是否可用，不安装任何软件 | Python标准库 |
| `rebuild --render` | 导入现有七页基线、重算公式、导出XLSX、修补原生元数据并验证，可另出PNG预览 | Python、Node及环境中已安装的 `@oai/artifact-tool` |
| `pack --destination 路径.zip` | 生成新的MANIFEST和续作ZIP | Python标准库 |

分析输出默认在 `outputs/analyze/`，重建输出默认在 `outputs/rebuild/`。入口不会覆盖 `inputs` 中的基线。重建完成的文件名为 `2026.11期中考试四年级名单_重建.xlsx`。

如在别处编辑了一份工作簿，可以显式指定输入：

```bash
python run.py analyze --input "修改后的工作簿.xlsx"
python run.py rebuild --input "修改后的工作簿.xlsx" --output-dir "outputs/本次修改" --render
```

本项目的严格验证器固定核对本次214名学生的历史快照。修改成绩或统计范围后，应先核实变更，再同步更新期望检查；不能为使测试通过而改回真实成绩。`audit.py` 不加 `--strict-baseline` 时可提取变更后的数据。若增删学生行或扩班，先更新 `src/audit.py` 的 `BLOCKS` 及Excel统计范围。

`rebuild` 会固定调用 `src/validate_workbook.py`。合法改分后，XLSX可能已成功导出，但命令仍会在最终历史快照校验处以非0状态退出；这表示旧快照的期望值需要按已核实的新依据更新，再重新验收。传入修改版作为 `--input` 或校验器的 `--reference` 不会关闭这些历史断言。`audit.py --strict-baseline` 只控制只读分析脚本，与重建的固定快照验证不是同一个开关。

## 目录用途

| 位置 | 内容 |
| --- | --- |
| `inputs/` | 最新完整Excel，仅保留一份基线 |
| `src/audit.py` | 便携只读分析脚本，Python标准库 |
| `src/rebuild.mjs` | 当前七工作表版本的重算与导出入口 |
| `src/preserve_native_metadata.py` | 保留其他年级打印设置、原字体charset，补齐图表引用缓存 |
| `src/validate_workbook.py` | 独立验证关键统计、真实0与空白、图表及重建前后的一致性 |
| `src/archive/` | 依据本对话恢复的历史生成器及其说明，仅作追溯 |
| `audit/` | 对本次恢复基线重新生成的数据快照与实际验证报告 |
| `HANDOFF.md` | 已完成工作、数据口径、待核实事项与下一步 |
| `STATUS.json` | 机器可读的接续状态、验证范围及未恢复项目 |
| `DEPENDENCIES.md` | 运行环境、字体与依赖边界 |
| `MANIFEST.json` | 包内文件大小与SHA-256 |

## 文件来源必须如实区分

1. **已实际恢复**：此前保存的最新版七工作表Excel，SHA-256为 `f1be87a06018d10b9763f0674f288710a3284a3c7c45e90e7fe5b796f4492307`。
2. **本次重新提取**：`audit/analysis.json`、`source_cells.json`、CSV等，来自上述已恢复基线，不冒称为旧临时目录里的原JSON。
3. **依据对话重建**：`src/archive/` 中的历史源码，包含对话中可见的最终修补，不能声称与旧磁盘文件逐字节相同。
4. **未取回**：最初上传的五工作表原件字节和已清空的临时文件。最新版保留了学生姓名、性别、基础分、附加分及原手填汇总；当前续作入口不依赖那份旧原件。

不要把当前七工作表Excel改名冒充原上传版，也不要直接交给旧五页生成器。归档生成器检测到已存在“分析”与“核查”页会明确停止。

## 如何查看验证证据

`STATUS.json` 记录哪些检查实际执行过；`audit/portable_audit_report.json` 是对恢复基线的独立复算；重建验证证据由最终打包过程加入 `audit/verification/`。

文件校验通过，表示包内字节与MANIFEST相符，不表示学校统计规则已经确认。自动重建通过，也不等同于在Windows Excel或WPS客户端中逐项做过界面验收。实际通过与尚未执行的检查分别记录，不混写。

## 再次打包

`pack` 会打包工程目录中符合规则的文件，但不会自动选择最近修改的Excel，也不会把 `--input` 指定过的外部文件自动带入包。`outputs/` 被排除，所以新成品只留在输出目录时，下一包仍会以旧工作簿为基线。

继续修改后，请按以下顺序交接：

1. 把本轮已验收、准备交接的Excel复制到 `inputs/` 中一个明确的新文件名；如需追溯，旧基线可保留在 `history/`，不要将其标成最初上传原件。
2. 更新 `project.json` 的 `baseline` 为新文件的包内相对路径，并更新其SHA-256、打包日期等元数据；同步 `STATUS.json` 中的路径、字节数、校验值、日期、统计数字、完成状态和待办。更新README与HANDOFF中会误导下一轮的旧基线名及旧数值。
3. 运行 `python run.py analyze --output-dir audit`，使包内JSON/CSV与新的当前基线一致。成绩或名单有合法变更时，先按新依据更新严格验证器的历史期望、公式范围和静态说明，再验收；不要仅为通过旧断言而回改成绩。
4. 保留本轮实际执行的验证证据，更新 `STATUS.json` 的证据路径和通过状态。旧报告若保留，应明确标为历史报告，不能拿它证明新数据已通过。
5. 运行下方打包命令，再将新ZIP解压到另一个目录执行 `verify` 和 `analyze`。工作文件修改后、重新打包前，旧MANIFEST校验不通过是预期现象；`pack` 会生成新MANIFEST。

```bash
python run.py pack --destination "../四年级成绩_下一版续作工程.zip"
```

临时输出、依赖目录、Python缓存和ZIP本身不会重复装入包中。保留成品清晰度依赖原生Excel内容，不做图片化压缩。
