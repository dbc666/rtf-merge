# CHANGELOG

## 1.1.2 — 2026-11-06

第二批：结构收敛。全量 60 文件 **RTF 36,086,787 B / SHA256 `39f04383…2ab999`
与 1.1.1 逐字节一致**，合并段零变化。

### 主链路收敛：ps2 整块降级为只读检测
- `step2_find_empty.ps1` 重写：删掉 `Remove-AllSpacesInEmptyPages`（源头治理后
  实测恒产 0 个空格）、交互删页整段（`MessageBox` 会让无人值守挂死，最长
  2700s）、`Update-TocBlue`，以及 `:204` 的 `SaveAs2`。
  现在只做一次全页扫描，输出 `EMPTY_PAGE=` / `EMPTY_PAGES_COUNT=` /
  `EMPTY_PAGES_SKIPPED=` / `EMPTY_PAGES_DELETED=0` / `SPACE_CHARS_DELETED=0` /
  `DOC_PAGES=` / `WORD_VERSION=`。
- 链路变成 `ps1（1 次 Word 写盘）→ python-docx 重写 → docx_xml（唯一最终写盘）
  → ps3（只读）→ 验收`，彻底消除"Word 重存摧毁 XML 字体/目录颜色"的顺序耦合。
- 保留原版 `Get-EmptyPages` 全页扫描（`$sec.Range.Paragraphs.Last` 版实测卡死）。
- **实测**：降级后 278 页、**0 空白页**，与降级前一致，无需回退。

### 目录蓝色迁到 XML 层（替代 ps2 的 Update-TocBlue）
- `docx_xml._fix_toc_styles()`：直接把 `styles.xml` 里 TOC 1~9 设为
  宋体 + Times New Roman、`sz=24`（小四）、`color=0000FF`（蓝）、`b val="0"`、
  移除 `w:autoRedefine`；**缺失的 TOC4~9 自动补建**（以后正文出现 H4~H9
  标题时条目直接继承蓝色，不会回落黑色）。
- 写完自检：样式缺失、颜色/字号/加粗不符即抛错。
- 效果：用户在 Word 里右键"更新目录"后条目仍为蓝色；实测 9 个样式全部
  `0000FF / 24 / b=0 / 宋体 / 无 autoRedefine`，Word 能正常打开（ps3 复检 278 页 / 0 空白页）。

### manifest 补齐后处理留痕
- 新增 `merge_core.update_manifest_postprocess()`：把 docx 的 SHA256/大小/页数/
  空白页/rFonts 覆盖率/目录域/TOC 样式数/Word 版本/命令行/mode/耗时写进清单
  `postprocess`；`validation` 段收敛为总验收结论 `all_ok`。原子写（`.tmp` + `os.replace`）。
- `word_post.fix_fonts_via_word` 返回 `PostResult`（dict 子类，`__str__` 仍是 docx
  路径，旧用法不变），承载 `step1` / `step2_detect` / `step3_recheck` / `fonts` 等统计；
  `_verify_final_docx` 改为返回统计 dict；`_recheck_empty_pages` 返回
  `(残留空白页数, 总页数)`。
- `cli.py` / `gui.py` 都在后处理结束后写清单（cli 带完整命令行、`mode=cli`；
  gui 带 `mode=gui`）。

### 结构校验进主流程
- `rtf_lint.check_structure_file()`：流式扫合并输出（不把 36MB 读成字符串），
  统计行数、转义花括号的开闭配平、`\rtf1` 次数、首末行。
  `\rtf1` 次数只记录不判错（60 份头的 SAS 原语义已确认保留）。
- 通过时打日志、失败时写 `warnings`；结果进 `result["rtf_validation"]` 与清单
  `rtf_validation`，`validation.rtf_structure_ok` 取其结论。
- **实测**：334300 行、花括号 47652/47652 配平、`\rtf1` 60 处。

### 顺手修
- ps1 之前先检查目标 docx 是否被占用：被上一轮残留的无窗口 Word 锁住时
  按基线差集清一次，清不掉再明确报错（不再把"文件已在别处打开"甩给用户猜）。
  本次即由此复现并解决了真实阻塞（PID 23232）。
- `step3_recheck.ps1` 的 `Documents.Open` 统一为 1 参（与 ps2 注释一致）。
- GUI 失败提示补 `order_incomplete` 条目。

### 验收数据（1.1.2 全量，两批改动的最终值）
- RTF 36,086,787 B / SHA256 `39f04383…2ab999`（与 1.1.0/1.1.1 逐字节一致，
  第三批同样复现该基线）
- docx 15,759,891 B（第三批比第二批 +137 B，来自标题段落样式化）
  / **278 页 / 0 空白页** / rFonts 88408（TNR 100%、宋体 100%）
- 101 单元格合并 / 63 表格 AutoFit / 230 中文引号 / 81 标题加粗 / TOC 1~9 全蓝
- 清单 `validation.all_ok = true`，后处理耗时 694.5s

### 第三批：XML 层与 ps1 副作用收敛（同日收口）

#### 字体：符号字体不再被误改成 Times New Roman
- `_fix_docx_fonts` 增加白名单 `SYMBOL_FONTS = {Symbol, Wingdings, Wingdings 2/3,
  Webdings, Marlett}`，检查 ascii/hAnsi/cs/eastAsia 四个属性；命中即原样保留。
  项目符号、复选框这类用符号字体渲染的字符不会因"统一字体"变成一个方框。
- `set_rpr_fonts` 改为返回"是否真的改写"，`fix_root` 只对真正改写的 rPr 计数，
  统计值不再虚高。

#### 中文引号：拆分支持 `<w:br>` / `<w:tab>` 原子段
- 原先按子元素切分时遇到换行/制表符会被整体跳过，导致这类 run 里的引号漏改。
  现在把 `w:br` / `w:tab` 作为原子段独立成 run（`deepcopy` 原元素），其余仍按
  引号切分；纯文本段首尾带空格时补 `xml:space="preserve"`。
- 抽取 `set_quote_hint()` 复用同一套 rPr 处理。本次实测引号数仍为 230（无回归）。

#### 标题加粗：保留原段落样式，不再一律改成 RTFHeading
- `_fix_toc_source_bold` 重写：只处理**run 本来带直接 `w:b`/`w:bCs`** 的大纲段落；
  段落已有非 Normal 的 `pStyle` 时**保留原 pStyle**，改为给该样式补 `b/bCs` 与
  统一 rFonts（`ensure_style_bold`）；仅无 pStyle 或 Normal 时才改用 `RTFHeading`。
- 无大纲标题时返回 0 而不是抛错。本次实测 81 个标题、与第二批一致。

#### ps1：单元格合并失败不再拖垮整篇
- `tbl.Cell(r, c+1)` / `Merge` 包 `try/catch`：垂直合并的表头等取不到右侧单元格
  时跳过并计数，输出 `MERGE_SKIPPED=`。本次实测 `MERGE_CELLS=101`、
  `MERGE_SKIPPED` 未触发。

#### ps1：目录已存在时改为 `Update()`，不再先 `Delete()` 再 `Add()`
- 重复跑同一份 RTF 时原目录域不会被删掉重建（保留用户可能的域设置与书签），
  仅在确实没有目录时才新建（新建分支保留点状引导符、`Update()`、分页符）。

#### GUI：新增"停止"按钮与关窗清理
- `word_post` 引入 `CANCEL_EVENT` / `TaskCancelled` / `check_cancel()`；
  `_run_ps` 由 `subprocess.run` 改为 `Popen` + 1 秒轮询，取消即杀当前 PowerShell
  子进程（保留原有 timeout 语义），并在 ps1 前 / ps2 前 / XML 修正前 / ps3 前
  设检查点。CLI 从不置位，行为不变。
- GUI 新增"停止"按钮（运行中可用，点击后置取消信号并显示"正在停止..."）；
  关窗确认退出时同样置取消、等线程 10 秒、调用 `_kill_headless_word()` 清理
  本次产生的无窗口 Word，再销毁窗口。
- 空白页选项由"自动跳过"改为纯"检测"：`EMPTY_PAGE_AUTO` 固定 `NO`（ps2 已只读，
  ASK 分支在 1.1.2 已删除）。

#### 归档
- v1.0 的两个单体实现（`rtf_merge_legacy.py`、`rtf_merge_gui_legacy.py`）移入
  `archive/`，并新增 `archive/README.md` 标明"仅历史参考、不应运行"。

### 第四批：第三方复审问题收敛（同日）

> 复审对象是第三批产物。以下按复审编号记录；未改项的取舍见文末。

#### M1（P1）输出文件名后缀统一——实锤缺陷
- 根因：`run_rtf_merge.bat` 传 `--outfile "%OUTDIR%\合并结果"` 不带后缀，而补
  `.rtf` 只写在 GUI 的 `start()` 里，CLI 与 bat 都没有。复现结果是产出无扩展名
  文件 `合并结果`，完成提示却列着 `合并结果.rtf`。
- 修法：核心层新增 `merge_core.normalize_outfile()`，`merge_rtf` 入口对无扩展名
  的输出路径统一补 `.rtf`，CLI / GUI / bat 三条入口行为一致；
  `dist\run_rtf_merge.bat` 同时把参数写成 `%OUTDIR%\合并结果.rtf`（双保险）。
- 实测：`--outfile "...\合并结果"` → 产出 `合并结果.rtf` / `.docx` /
  `_manifest.json`，清单 `outfile` 字段也是 `.rtf`。

#### M3（P1）回归门跨机器不再报 ERROR
- `test_drop_removes_one_line_per_file` 直接遍历 `self.samples`；样本目录缺失时
  列表为空，随后读 `out.rtf` 抛 `FileNotFoundError`，在别的机器上表现为 ERROR。
  现在入口先 `require_samples(1)`。
- 实测：`RTF_MERGE_TEST_INPUT` 指向不存在目录 → **13 tests / OK / skipped=9**，
  无 ERROR/FAIL（修复前为 1 ERROR + 8 skipped）。

#### M4（P2）清掉空白页选项的"死开关"宣传
- CLI `--empty-page-auto` 保留但 help 明确写"兼容保留，已无实际作用"；
  不再把取值写入 `EMPTY_PAGE_AUTO` 环境变量（ps2 已不读，留着反误导）。
- GUI 删除"空白页自动处理"复选框（`var_empty` / `empty_auto_no`），改为一行
  灰色说明"自动检测（只读统计，不删页、不弹窗）"；docstring 与"完整后处理"
  文案里的"删除空白页"统一改为"空白页检测"。
- README 标题版本号 v1.1.0 → v1.1.2。

#### M6（P2）总验收不再"带告警也全绿"
- `validation.all_ok` 追加两个条件：无产物级 warnings、`step1.merge_skipped == 0`；
  同时写入 `warning_count` / `merge_skipped`。"输出文件位于输入目录内"属于
  使用提示，显式排除在产物级告警之外。
- `word_post` 解析 `MERGE_SKIPPED=` 记入 `step1.merge_skipped`，非 0 时打印提示。
- 结构校验未通过时 CLI 直接返回 **4** 且不进入 Word 后处理（此前只记日志、
  返回 0，产物"看似成功"）。

#### M8（P2）GUI 清单可回答"当时选了什么"
- `command_line` 由写死的 `"gui"` 改为
  `gui: inpath=… outfile=… order=… pageyn=… fix_fonts=… delline=… encoding=…`，
  并把 `opts` 快照存进 `postprocess.options`。

#### M9（P2）Word 进程基线每次任务重拍
- `_snapshot_word_baseline()` 去掉"整个进程生命周期只拍一次"的早退：GUI 连跑
  第二个任务时重新拍快照，上一任务遗留的僵尸实例仍可清，用户在两次任务之间
  新开的 Word 也不会被算作"本次新增"。
- **更正（见第五批）**：这段因果写错了。重拍会把上一任务的残留并进基线，
  反而使它清不掉；第五批已回退为"每进程只拍一次"，并把真实因果写进注释。

#### M10（P2）清单一致性与实际编码
- 首份清单 `_write_manifest` 改为 `.tmp` + `os.replace` 原子写，与
  `update_manifest_postprocess` 一致。
- 新增 `encodings_used`（各文件实际解码编码集合），`encoding` 仍记参数值。

#### M5（P2）不再让人误读两个页数
- ps2/ps3 页数不同（262 / 278）是因为 ps2 测的是 XML 修正前的中间态。保留两次
  检测（ps2 承担"转换阶段是否引入空白页"的诊断价值），但清单里
  `step2_detect.stage="pre_xml_fix"`、`step3_recheck.stage="final"` 标明阶段；
  `all_ok` 只采用 ps3 结论（原本如此）。

#### 未改（取舍说明）
- **M7**：大纲标题在真实产物里 pStyle 全是 `RTFHeading`，`ensure_style_bold`
  分支未被真实数据覆盖；符号字体白名单命中 0 次。两者都是防御性代码，删掉会
  失去对真实文档的兜底，故保留并登记为待补单测项。
  （第五批已补 `tests/test_docx_xml.py`，6 个用例，见下。）

#### 验收（第四批小样本，重建后 exe）
- 输入 3 个 RTF，`--outfile "…\合并结果"`（不带后缀）
- 产物齐备：`合并结果.rtf` / `.docx` / `_manifest.json`
- 清单：`encodings_used=["utf-8"]`、`step1.merge_skipped=0`、`warning_count=0`、
  `all_ok=true`、`docx_pages=5 / empty=0 / rfonts=482`、
  `step2.stage=pre_xml_fix`、`step3.stage=final`

#### 验收（第四批全量，由交付版 exe 产出）
- 命令即清单 `command_line`：`F:\rtf_merge\dist\rtf_merge.exe --inpath …
  --outfile F:\rtf_merge\all_merged.rtf --fix-fonts --order 0 --pageyn 1`
  （验收数据第一次由交付 exe 而非源码入口产出，M2 的时序问题就此消除）
- RTF 36,086,787 B / SHA256 `39f04383…2ab999`，与基线逐字节一致；
  334,300 行、花括号 47652/47652、`\rtf1` 60 处
- docx 15,759,649 B / 278 页 / 0 空白页 / rFonts 88408（TNR 100%、宋体 100%）
- 101 单元格合并 / 63 表格 AutoFit / 230 中文引号 / 81 标题加粗 / TOC 1~9 全蓝
- 清单 `warnings=[]`、`warning_count=0`、`merge_skipped=0`、`all_ok=true`、
  `encodings_used=["utf-8"]`、Word 16.0、耗时 1269.5s

### 第五批：第二轮复审（Q1–Q6）收敛（同日）

> 复审对象是第四批产物。三个 P1（M1/M2/M3）已实证销项，但发现一处 P1 级
> CLI/GUI 行为分叉，以及数处文档与健壮性问题。

#### Q1（P1）结构校验闸门下沉到核心层——CLI/GUI/bat 一处判定
- 现象（已复现）：构造 4 开 3 闭的输入，`merge_rtf` 仍返回 `status="merged"`，
  RTF 与清单已落盘，`rtf_validation.issues` 非空；CLI 有自己的 issues 判断会
  返回 4 且不进 Word，而 GUI 只判 `status != "merged"`，于是一路跑完 Word
  后处理并弹"任务完成"。GUI 恰是 README 推荐的入口，这个分叉是交付级缺陷。
- 修法：`merge_rtf` 在结构校验失败时直接把状态置为 `structure_invalid`
  （文件与清单照常写出，便于排查）；CLI 识别该状态返回 4，GUI 的 hints 增加
  同名条目并提示"已停止 Word 后处理"，bat 走 CLI 自然一致。
- 回归：新增 `TestStructureGate` 两个用例——坏结构必须 `structure_invalid`
  且产物/清单仍存在；好结构必须仍是 `merged`。
- 实测：CLI `exit=4` 且日志停在"已停止 Word 后处理"；GUI `run_task` 返回
  error 文案，且 `fix_fonts_via_word` 一次都没被调用。

#### Q2（P2）`test_drop_removes_one_line_per_file` 的样本阈值
- `require_samples(1)` 不足：单文件时 `maxn==1`，首文件分支根本不删末行
  `}`（keep/drop 差 0），而 `expected` 仍按"末行为 `}` 的文件数"算 1。
  用"末行为 `}` 的单样本"实测得到 `AssertionError: 0 != 1`。
- 修法：改为 `require_samples(2)`；样本不足一律 skip。实测该场景从
  FAIL 变为 `OK (skipped=1)`。

#### Q3（P2）M9 的方向反了：回退为"每进程只拍一次"
- 上一批把"每次任务重拍基线"当成修复，因果恰好相反：重拍后上一任务的
  残留被并入基线，**反而永远清不掉**，连 pre-flight 也清不掉（同一个判据），
  只会落到"请关闭后重试"。沿用同一次快照时，本进程后续任务自己起的 Word
  始终在基线之外，才能被差集清掉。
- 修法：恢复 `_BASELINE_TAKEN` 早退（`_BASELINE_TAKEN` 也因此重新有意义），
  注释改写为真实因果，并注明代价（进程存活期间用户新开的无窗口 Word 可能
  被当作新增，但判据里的 `MainWindowTitle` 为空已把用户文档排除）。
- 配套：docx 被占用的报错文案补一句"若看不到 Word 窗口，多半是上次任务
  残留的无窗口实例，请在任务管理器里结束 WINWORD.EXE 后重试"。

#### Q4（P2）"输出在输入目录内"改用常量判定
- 原先两处各抄一遍中文前缀，改一个字就会让 `all_ok` 的口径静默变化。
  抽成 `INPUT_DIR_WARN_PREFIX` 常量，追加与过滤共用。

#### Q5（P2）耗时归因纠正：不成立，且已拿到反证
- 第四批把 1269.5s 归因为"ps2 → python-docx → ps3 三段冷启动"，但两版链路
  完全相同、step2/step3 页数也一致，多出的 575s 不可能是链路差异造成的。
- 反证：本轮同一台机器、同一链路、同一输入、ps2 未删，全量耗时 **659.7s**
  （与第三批 694.5s 同量级）。同一配置两次跑相差 610s，说明那个 1269.5s 是
  一次异常慢跑（冷启动/机器负载），不是链路特性。因此"删 ps2 能省 575s"
  这个说法没有依据，已从记录中删掉。

#### Q6（P2）残留文案清理
- `gui.py` 的 `run_task` docstring 删掉已不存在的 `empty_auto_no`。
- README 参数表把 `--empty-page-auto NO|YES|ASK` 改成空参数说明，只保留
  "兼容保留、无实际作用"，不再列出两个无效取值。

#### M7 补单测：`tests/test_docx_xml.py`
- 报告指出"登记为待补单测"只留在文字里，换人接手会丢。现落成 6 个合成
  docx 用例（不依赖真实样本、不启动 Word）：
  1) Symbol 字体的 run 在 `_fix_docx_fonts` 后原样保留，普通 run 变 TNR/宋体；
  2) 未修字体的 docx 被 `_verify_final_docx` 拦下（不再静默通过）；
  3) 带 `Heading 1` pStyle 的标题迁移后**保留原 pStyle**，加粗改到样式级；
  4) 同样式但原本无直接加粗的段落不受影响；
  5) 无大纲级别段落返回 0、不抛错；
  6) 中文引号被拆成 `hint=eastAsia` 的独立 run（4 个引号 → 4 个 run）。

#### ps1 任务级 PID 精确回收（决策1 的采纳方案）
- 三段 ps1 在 `New-Object -ComObject Word.Application` 前后各枚举一次
  WINWORD PID，差集以 `WORD_PID=37700` 形式输出；ps1 复用已有实例时差集为空，
  正好不记录（用户的 Word 不会被登记）。
- Python 侧 `_TASK_WORD_PIDS` 每次任务清空、逐段累加，`_kill_headless_word()`
  改为「先按 PID 精确回收（多一道 WINWORD 进程名校验防 PID 复用），再用基线
  差集兜底」；任务超时/失败/正常结束都会走到这里。
- 相对内存版 `_TOOL_PIDS` 的好处：不需要持久化、不受 PID 复用影响、不依赖
  窗口标题与全局基线、也不会扩大误杀面；跨进程（进程被强杀）遗留仍靠
  pre-flight 报错误导用户手动结束。
- 实测：单跑 step2 输出 `WORD_PID=37700`、`DOC_PAGES=5`；小样本端到端结束后
  `WINWORD=0`。

#### 分段计时（`duration_sec` 落地，替代"凭归因删 ps2"）
- `postprocess` 新增 `durations` 段，按 `step1_convert` / `docx_rewrite` /
  `step2_detect` / `xml_fix` / `step3_recheck` / `verify` 六段分别计时；
  `duration_sec` 仍为总耗时。
- 两次全量实测（同一版 exe、同一输入、产物 RTF 逐字节一致）：

  | 段 | 11-07（正常运行） | 11-08（机器偏慢） |
  |---|---|---|
  | step1_convert（RTF→docx + 合并单元格 + AutoFit + 目录） | 269.0 | 625.7 |
  | step2_detect（262 页空白页扫描，只读） | 174.9 | 261.3 |
  | xml_fix（字体/引号/标题/TOC 样式） | 18.3 | 42.6 |
  | step3_recheck（278 页独立实例复检，只读） | 197.0 | 233.7 |
  | docx_rewrite + verify | 4.1 | 5.0 |
  | **合计** | **663.3** | **1168.3** |

- 结论一：**真正的大头是 step1，其次是 step2 与 step3**。删 ps2 的理论上限
  只有 175s（26%，正常跑）/ 261s（22%，慢跑），与"删 ps2 能省 575s"相距甚远，
  该结论彻底作废；且 step3 比 step2 还慢（278 页 > 262 页 + 新实例冷启动），
  把两者合并只能省一次冷启动，却会丢掉"独立实例能正常打开最终文件"这项验收，
  故 ps2 保留不删。
- 结论二：**交付时长本身就有 1.7 倍抖动**（663.3s → 1168.3s），且各段同比例
  变慢（step1 2.33×、step2 1.49×、xml_fix 2.33×、step3 1.19×），产物却完全
  一致——这是 Word/机器负载造成的，与代码路径无关。任何"某一步耗时多少秒，
  所以删掉它"的归因都必须先排除这个抖动（这正是 Q5 那条错误结论的成因）。

#### 告警结构化（加法式演进，不动旧字段）
- 清单新增 `warnings_detail`：每条 `{kind, level, text}`；`warnings` 保持
  纯字符串数组，旧读取方（CLI/GUI/同事脚本）完全不受影响。
- 判定口径从"中文前缀 startswith"改为"level != info"：改造前唯一按前缀
  豁免的是「输出落在输入目录内」，改造后该类固定 `level=info`，其余 8 类
  （`order_contains_output` / `order_partial_allowed` / `resource_mismatch` /
  `encoding_fallback` / `rtf_header_invalid` / `rtf_tail_invalid` /
  `structure_invalid` / `manifest_write_failed`）全为 `warn`。
- 效果：`warning_count` 与 `all_ok` 的判定结果与改造前**逐例一致**（纯结构
  升级）；以后改文案不会让这条豁免静默失效。
- `update_manifest_postprocess` 读老清单（无 `warnings_detail`）时自动回退
  到前缀匹配，旧产物仍能算 `all_ok`。
- 新增 `TestWarningDetail` 三条：info 级提示不进 `warning_count` 且
  `all_ok` 仍 True、`warnings` 与 `warnings_detail` 文本逐条同源并双写进清单、
  输出目录外时两个字段都为空。

#### `--empty-page-auto` 改为"保留但显式提示"（选项 C）
- 不删除该参数：删了会让传过 `--empty-page-auto NO` 的老脚本在 argparse
  阶段直接 exit 2（连合并都跑不起来）。
- 改为：help 标"已废弃"，且传 `YES`/`ASK` 时打印
  "该参数已废弃，无实际作用；空白页只做只读检测，不会删除任何页面"。
  行为与 `NO` 完全相同，退出码不变。实测传 `YES` 合并正常完成。
- 顺手修正 README 的一处错误描述：原来写"不再接受 NO/YES/ASK"，
  实际 choices 仍接受三值，只是不再起作用。

#### `order_incomplete` 默认拒绝写盘：维持，并补文档
- 该闸门保护的是"静默漏合并"——60 个文件只填 1 行 order 会产出只含 1 个
  文件的成品，事后无法发现。故**不改成默认放行**。
- `--allow-partial-order` 仍是显式通道；放行后清单记 `order_partial_allowed`
  告警且 `all_ok=false`（放行≠变绿），这一设计不动。
- README 新增一段说明"这是防漏合并闸门，不是缺陷"，并给出补救步骤。

#### 本轮未做（明确记录）
- 无（前两项已完成；`--empty-page-auto` 与 `order_incomplete` 按上两节定格）。

#### 验收（第五批）
- 单测：**24 tests OK**（原 15 + `TestStructureGate` 2 + `test_docx_xml.py` 6 +
  `TestWarningDetail` 3）
- 无样本环境：13 OK / skipped=9（原 13 项不变）；单样本目录：该用例
  `OK (skipped=1)`，不再 FAIL
- 结构闸门：CLI `exit=4` 且日志"已停止 Word 后处理"；GUI `run_task` 返回
  error 文案且 `fix_fonts_via_word` 零调用
- ps1 语法：三段 `ERRORS=0`；单跑 step2 实测输出 `WORD_PID=37700`、
  `DOC_PAGES=5`（新 PID 枚举块生效）
- 小样本（重建后 exe，3 文件）：exit 0，产物三件齐全，`warnings=[]`、
  `warnings_detail=[]`、`all_ok=true`，任务结束后 `WINWORD=0`
- `--empty-page-auto`：源码入口传 `YES` 实测打印"已废弃参数…不会删除任何
  页面"、合并正常完成、退出码 0；`NO` 无提示、行为不变
- 全量（重建后 exe）：RTF 36,086,787 B / SHA `39f04383…2ab999` 与基线
  逐字节一致；334,300 行 / 花括号 47652:47652 / `\rtf1` 60；
  docx 15,760,004 B / SHA `c0196887…631269e` / 278 页 / 0 空白 /
  rFonts 88408（TNR 与宋体各 100%）/ 101 单元格 / 63 表格 / 230 引号 /
  81 标题 / TOC 1~9 全蓝；step2 262 页（pre_xml_fix）→ step3 278 页（final）；
  清单 `warnings=[]`、`warnings_detail=[]`、`warning_count=0`、
  `merge_skipped=0`、`all_ok=true`、`command_line` 指向
  `dist\rtf_merge.exe`、总耗时 1168.3s（分段见上，两次全量产物一致）；任务
  结束后 `WINWORD=0`（PID 精确回收生效）
- 打包 ps1 同步：`dist\rtf_merge_gui\_internal\rtfmerge\ps\` 三份与源码
  `same=True`（`781D44B3810C19CE` / `FA24BECE14DBB799` / `0C4FD6DD2EC06335`）
- docx 每次重建 SHA 都不同（zip 重打包 + Word 域结果），故只称"结构指标
  等价"，不称"内容等价"

### 第六批：GUI 默认输出名 / 总耗时日志 / 内存占用定位
#### GUI 默认输出文件名改为输入文件夹名
- `browse_in`：选中输入目录后，输出框自动填 `<输入文件夹名>.rtf`，且放在
  输入目录的**上一级**（如输入 `…\data\空腹_TF` → 默认输出
  `…\data\空腹_TF.rtf`）。放上一级是为了不把产物落进输入目录、被下次扫描
  当成输入重复合并。
- 用户仍可直接在输出框里改路径或文件名；盘根目录等取不到文件夹名的情形
  回退为 `合并结果.rtf`。路径末尾带 `\` 或 `/` 时归一化，不会得到空名。
- CLI 的 `--outfile` 语义不变（显式指定，不做推导）。

#### 运行日志输出耗时与内存
- GUI 日志新增三类行：`合并阶段耗时`、`后处理分段耗时`（六段 + 合计）、
  `总耗时`（完成/取消/出错都在 `finally` 里打印，便于事后核对"这次为什么慢"）。
- 同时打印 `进程内存：全程峰值 … MB，当前 … MB（起始 … MB）`，用
  `GetProcessMemoryInfo`（ctypes，零依赖）。这样"为什么还占 2GB"可以自助核对。
- CLI 侧把 `main()` 拆成包装 + `_main()`，在 `finally` 里统一记一条总耗时，
  与合并/后处理/结构失败各条返回路径无关。

#### 内存占用根因（实测数据）
| 阶段 | 工作集 | 峰值 |
|---|---|---|
| 进程基线（import docx 后） | 25 MB | 25 MB |
| 打开一次 15.7MB / 278 页 docx | 493 MB | 525 MB |
| 保存 | 495 MB | 558 MB |
| **save 后不释放主 DOM 就重新打开自检** | **957 MB** | **990 MB** |
| 两棵 DOM 都 del + gc | 31 MB | 990 MB |

- 结论：15.7MB 的 docx 展开成 lxml 元素树要 **约 470MB**（30 倍膨胀）。
  四个修正函数原本都是「save 后立刻重新打开自检」，而主 DOM 当时还没释放，
  于是每步同时存在两棵树 = **峰值约 1GB**；四次串行、加上 `save()` 内部的
  序列化/压缩临时，进程峰值就到 1.5~2GB。
- **【第七批更正】上一句之后原写"任务结束后 Windows 不会主动把 WorkingSet
  还给系统"，这条归因是错的**：同机实测 `del` 之后若**不跑 `gc.collect()`**，
  工作集确实停在 966MB 不动（lxml 元素树有循环引用，引用计数收不回），
  但补一次 `gc.collect()` **立刻回落到 37MB**。所以任务管理器长期显示 2GB
  ≠"Windows 不还内存"，而是**还有文档树活着**——第六批漏了 `word_post.py`
  里 `_rd` 那一棵（详见第七批）。
- 已做（低风险）：四个 `_fix_*` 在 `doc.save()` 后 `del doc; gc.collect()`
  再复检，复检完 `del check; gc.collect()`；`_verify_final_docx` 结束也释放
  （含它序列化出来的 `xml` 字符串）。消除"两棵树并存"这一步，约省 470MB。
- 未做（需重构，另议）：把四个修正步骤合并成「一次 load → 四次修改 →
  一次 save」，可再省三次加载/保存（`xml_fix` 42s 也会明显下降），但四个
  函数各自的独立自检要重新组织，风险高于本轮范围。

#### 验收（第六批）
- 单测：24 tests OK
- 默认输出名：五个分支实测——首次选目录、再选另一个目录（跟随改名）、
  手改后再选目录（保留手改值）、清空后再选（重填默认值）、盘根 `F:\`
  （回退为 `F:\合并结果.rtf`）；带尾部 `\` 与不带尾部两种写法结果相同
- GUI 日志实测（3 文件小样本 + 完整后处理）：`合并阶段耗时 0.2 秒`、
  `后处理分段耗时 …（合计 67.2 秒）`、`总耗时 76.4 秒（约 1.3 分钟）`、
  `进程内存：全程峰值 40 MB，当前 37 MB（起始 22 MB）`
- 内存实测：单次加载 493MB（峰值 525）→ save 后不释放再打开 957MB
  （峰值 990）→ 两棵 DOM 都释放后回到 31MB；改后四步串行峰值 1041MB
  （`xml_fix` 21.7s，docx 15,759,749 B）
- 全量：RTF 36,086,787 B / SHA `39f04383…2ab999` 逐字节一致；docx
  15,759,749 B / 278 页 / 0 空白 / rFonts 88408（TNR 与宋体各 100%）/
  101 单元格 / 63 表格 / 230 引号 / 81 标题；`all_ok=true`、`warning_count=0`；
  总耗时 1097.8s（step1 605.3 / docx_rewrite 2.8 / step2 242.4 /
  xml_fix 21.7 / step3 210.3 / verify 1.9）；`WINWORD=0`

### 第七批：内存驻留根因更正 + `_rd` 未释放（复审销项）
> 来源：外部复审（Work buddy）指出"任务结束后 GUI 仍占 2GB+ 不是 Windows 不还
> 内存，而是 `word_post.py` 的 docx 重写树 `_rd` 从未释放"。三条指控逐条复核：
> ① `_rd` 未释放——**成立，已实测证实并修复**；② 第六批"Windows 不还内存"的
> 归因——**成立，该表述已更正**；③ 交付 exe 带缺陷——成立（`word_post.py`
> 自第五批后未改，而第六批只改了 `docx_xml.py`）。

#### 遗漏点：`word_post.py` 的 `_rd` 活到函数返回
- `:354 _rd = _RewriteDoc(docx_out)` / `:355 _rd.save(docx_out)` 之后既无
  `del` 也无 `gc.collect()`，此前全仓 `gc.collect()` 只在 `docx_xml.py`。
- 后果：这棵约 470MB 的树从"第 1.5 步 docx 重写"一直活到 `fix_fonts_via_word`
  返回（10–20 分钟），成为后面 4 个 `_fix_*` 的**常驻底座**；函数返回后只是
  引用计数 −1，lxml 元素树的循环引用要靠 gc 才收得回，所以任务结束后那
  470MB 仍挂着（GUI 空闲时几乎不分配新对象，gen2 迟迟不触发）。

#### 机制实测（同机、当前 `all_merged.docx`、只读副本）
| 步骤 | 工作集 | 峰值 |
|---|---|---|
| 基线 | 26 MB | 26 MB |
| `_rd` 打开并保存后**不释放**（旧行为） | 496 MB | 559 MB |
| 旧行为跑完 4 步 `_fix_*` + verify | 966 MB | **1505 MB** |
| 旧行为：`del _rd` 但**不 gc** | **966 MB（纹丝不动）** | 1505 MB |
| 旧行为：补一次 `gc.collect()` | **37 MB** | 1505 MB |
| 修复后：重写后立即 `del _rd; gc.collect()` | 30 MB | 559 MB |
| 修复后跑完 4 步 + verify | 502 MB | **1043 MB** |
| 修复后收尾 + gc | 37 MB | 1043 MB |

- 结论一：`del` 不等于释放。lxml 元素树有循环引用，**不 gc 就不还内存**；
  gc 一次立刻从 966MB 掉到 37MB。故第六批"Windows 不会主动把 WorkingSet
  还给系统"不成立，看到 GB 级占用说明**还有对象活着**。
- 结论二：`_rd` 这一棵底座对峰值的贡献是实打实的 **462MB**（1505 → 1043，
  −31%），且它让"收尾常驻"从 966MB 变成 502MB（再 gc 后 37MB）。

#### 改动（A 级，3 处、不改产物字节）
- `word_post.py`：顶部 `import gc`；第 1.5 步 docx 重写改为
  `try/except/finally`，在 `finally` 里 `del _rd; gc.collect()`
  （重写失败时同样释放，用 `except UnboundLocalError` 兜住未定义的情况）。
- `gui.py`：`run_task` 的 `finally` 在 `_mem_mb()` **之前**先 `gc.collect()`，
  避免把残留报成"当前占用"。
- 未采纳：`SetProcessWorkingSetSize(-1,-1)` 主动裁剪工作集——只是把页换进
  pagefile，峰值不变、下次访问还要换回来，纯观感，不做默认行为。

#### 文档更正
- `README.md`【内存】段重写：峰值是"曾经达到过"（任务管理器"峰值工作集"
  同口径）；当前占用收尾会回落到几十 MB；若长期停在 GB 级说明有树没回收。
- 上文第六批"内存占用根因"里"Windows 不会主动把 WorkingSet 还给系统"一句
  已就地更正，保留原文并标注【第七批更正】，便于追溯当时的判断偏差。

#### 验收（第七批）
- 单测：24 tests OK（8.4s）
- 内存对比实测见上表（`mem_probe4.py` 旧行为 / `mem_probe5.py` 修复后）
- 产物：`_rd.save()` 复用同一对象、只多一次 `del`，不改变写盘内容；
  本轮小样本后处理 `rFonts 88408`、TNR 与宋体各 100%、目录域存在，与第六批
  指纹一致（全量逐字节复验留待下次全量运行时确认 `39f04383…2ab999`）

### 第七批补充：粘贴输入路径也自动带出默认输出名
- 反馈：默认名原先只在"浏览..."按钮的回调里推导，用户从资源管理器**手动
  粘贴**路径到输入框时不生效（输出框保持空白）。
- 改法：把推导抽成 `App.apply_default_out(inpath)`（单一实现，浏览与键入
  共用），并在 `var_in` 上挂 `trace_add` → `on_in_typed()`：内容变化且是
  有效目录才推导，因此钳住"逐字符键入"的大量中间态。
- 踩坑（自测发现）：`os.path.isdir("F:")` 在 Windows 上**为真**（盘符简写
  等价于该盘当前目录），`abspath("F:")` 会得到 `F:\<该盘当前目录>`，据此
  推导会填出 `F:\rtf_merge.rtf` 这种错误默认名。故加一道校验——只认
  "本身就是规范化绝对路径"的输入（`normcase(abspath(v).rstrip sep)` 必须
  等于 `normcase(v.rstrip sep)`），盘符简写、相对路径一律不取样。
- 实测（无头，8 个场景）：粘贴有效目录 ✓、带引号+尾部空格粘贴 ✓、不存在
  的目录不填 ✓、手改输出后再换目录仍保留手改值 ✓、键入 `F:` 与中间态不填
  ✓、键入到完整路径才填 ✓、盘根 `F:\` → `F:\合并结果.rtf` ✓。
- 回归：单测 24 tests OK；GUI exe 重建后 3.11s 出窗，标题正常。

### 第八批：四个 docx 修正合并为「一次 load → 多次改 → 一次 save → 一次自检」（C 级）

> 来源：第七批评估中列为 C 级的架构改造。第七批已把峰值从 1505MB 压到
> 1043MB，但 `xml_fix` 仍是四个函数各自「load → 改 → save → 重开自检」：
> 一次任务要解析整棵树 8 次、序列化 4 次。本批把这 4 条链合并成一条。

#### 改动
- `docx_xml.py` 拆成两层：
  - **纯修改层**（只改已打开的 doc、不落盘）：`_apply_fonts`、
    `_apply_quotes`、`_apply_toc_source_bold`、`_apply_toc_styles`；
  - **单步入口**（保留原语义 load→apply→save→自检）：`_fix_docx_fonts`、
    `_fix_chinese_quotes`、`_fix_toc_source_bold`、`_fix_toc_styles`，
    供旧调用路径与既有 6 个单测继续使用，不改签名。
- 新增 `_self_check(docx_path, stats)`：字体、中文引号、TOC 样式、源标题加粗
  四项检查合并为**一次重开**；检查判据与报错文案与原四个自检逐条等价。
- 新增 `apply_all_fixes(docx_path) -> dict`：一次 `Document()` → 按与原实现
  **逐字相同**的顺序修改（字体 → 中文引号 → 源标题加粗 → 目录样式）→
  唯一一次 `save()` → `del doc; gc.collect()` → `_self_check` → 再 `gc.collect()`
  → 返回 `{"rfonts", "chinese_quotes", "headings_bolded", "toc_styles"}`。
- `word_post.py` 第 3~5 步改为单次 `fixes = docx_xml.apply_all_fixes(docx_out)`，
  日志四条中文提示文案一字未改。

#### 等价性（同一份 docx、两种实现，逐 zip 成员 SHA256 比对）
| 样本 | 旧实现（四个单步） | 新实现（`apply_all_fixes`） | 成员差异 |
|---|---|---|---|
| `sample3_7b.docx`（5 页 / 19 成员） | 4.9s / 峰值 34MB | 3.8s / 峰值 30MB | **0** |
| `all_merged.docx`（278 页 / 433 成员） | **48.0s / 峰值 1028MB** | **19.0s / 峰值 610MB** | **0** |

- 比对排除 `docProps/`（`core.xml` 含 modified 时间戳，必然不同）。
- 结论：修改顺序不变 → 产出的每个 XML 部件逐字节一致；耗时 −60%
  （48.0→19.0s），峰值 −41%（1028→610MB）。
- 端到端（源码入口重跑 3 文件小样本，输入为**未迁移状态**，因此这次真正跑到了
  「源标题加粗」分支）：`EXIT=0`、5 页 / 0 空白 / rFonts 482 / TNR 与宋体各
  100% / TOC 9 样式 / 4 引号 / **5 源标题加粗** / 总耗时 77.1s
  （其中 `xml_fix` 由 1.0s 降到 **0.2s**）/ `warnings=0`。
  与旧实现产物做归一化比对（剔除 rsid、`w14:paraId`、`w14:textId`），差异只剩
  Word 自己生成的 TOC 书签名（`w:anchor="_Toc2450…"`）与 `w15:docId` GUID，
  属 Word 每次运行的随机量，与本次重构无关。

#### 顺手修掉的一处收尾内存
- 自检函数内部那一次 `gc.collect()` 收不掉东西：调用时自检自己的栈帧与循环
  变量还绑着元素。实测 `apply_all_fixes` 返回后工作集仍 498MB，补一次 gc
  立刻降到 34MB（回收 3697 个对象，第二次 gc 回收 0 个 = 确无遗留）。
- 故在 `_self_check()` 返回后再 `gc.collect()` 一次，并在自检内部把循环变量
  显式置 `None` 断开元素引用。
- 复测：`apply_all_fixes` 返回后 **35MB**、峰值 609MB；再 gc 无对象可回收。

#### 验收（第八批）
- 单测：24 tests OK（19.5s）
- 等价性 + 峰值：见上表
- 双 exe 重建：`dist\rtf_merge.exe` 31,714,880 B、
  `dist\rtf_merge_gui\rtf_merge_gui.exe` 5,982,836 B（同批重建，`--clean`）
- 全量（`dist\rtf_merge.exe`，`--inpath output_rtf_空腹_TF --order 0 --pageyn 1`，
  60 个 RTF）：
  - RTF 36,086,787 B / SHA256 `39f04383…2ab999`，与 SAS 基线**逐字节一致**；
    334,300 行 / 花括号 47652:47652 / `\rtf1` 60 处
  - docx 15,759,993 B / **278 页 / 0 空白** / rFonts 88408（Times New Roman 与
    宋体各 100%）/ 101 个关键词单元格 / 63 个正文表格 AutoFit / 230 中文引号 /
    81 源标题加粗 / TOC 1~9 全蓝
  - `warnings=[]`、`warnings_detail=[]`；任务结束后 `WINWORD=0`
  - 总耗时 **930.4 秒**（step1 376.6 / docx_rewrite 3.8 / step2 244.3 /
    **xml_fix 11.9** / step3 281.6 / verify 2.7）；`xml_fix` 由第六批的
    21.7s 降到 11.9s（−45%）。总耗时与 step1/step3 的绝对值受 Word 负载影响
    （本机历史有 1.7 倍抖动），只有 `xml_fix` 这一项是本批改动的直接结果。
  - 口径说明：上面的 930.4s 是**进程级**（CLI 日志"总耗时"，含 RTF 合并与
    清单写入）；清单里的 `postprocess.duration_sec` 是 **921.0s**（只含后处理
    段，分段合计 920.9s）。两者差 9.4s 属口径差，不是数据打架。

#### 第八批补充：新主路径入回归门 + 内存读数入清单 + 文档口径修正
> 来源：外部审计（Work buddy）提出的 S1~S4。逐条核实：S1（新主路径零单测）
> 成立、S2（内存数据没进清单）成立、S4（根目录散落旧产物）成立；
> S3 属口径表述问题——两个耗时数字各有其义，补说明即可，不改数。

- **S1（P2）新主路径没有回归门**：`tests/` 里对 `apply_all_fixes` /
  `_self_check` / `_apply_*` 的引用数为 0，24 项测的全是旧单步入口。本轮
  最大的架构改动只靠一次手工 A/B 证明等价，后续重构可能静默改坏主路径。
  已补 3 个用例到 `tests/test_docx_xml.py::TestApplyAllFixes`：
  ① 四键齐全 + 四项自检不误抛；② 无中文引号 / 无大纲标题时不误抛；
  ③ 同一输入分别走"四个单步入口"与 `apply_all_fixes`，除 `docProps/` 外
  各 zip 成员 SHA256 全等（把一次性手工 A/B 固化成可重复的门）。
  结果：`test_docx_xml` 由 6 项增至 **9 项全 OK**。
- **S2（P2）内存读数没进交付物**：`gui.py` 只把内存行写进 GUI 日志队列
  （不落盘），CLI 产出的清单里 19 个 postprocess 键无任何 memory 字段。
  已新增 `rtfmerge/meminfo.py`（单一 ctypes 实现，GUI 与清单共用同一口径）、
  `gui._mem_mb()` 改为调用它；`word_post` 在进入后处理时采样起点、在最终
  验收后先 `gc.collect()` 再采样，写入 `postprocess.memory =
  {start_mb, peak_mb, end_mb}`。
- **S3（P3）口径**：`README`【内存】段的"一次读入"已收窄为"**四个 XML
  修正**合并为一次读入"（Word 侧重写、最终验收、自检复检仍各需一次读入）；
  并在 README 与 CHANGELOG 注明"清单 duration_sec 只算后处理段，日志总耗时
  算整个进程"。
- **S4（P3）散落旧产物**：根目录确有 `output_rtf_空腹_TF.rtf`（36MB，与基线
  同名不同路径）、同名 `.docx`、`_manifest.json`，以及一个 Word 锁文件
  `~$l_merged.docx`。若把输入目录设为 `F:\rtf_merge`，`rglob("*.rtf")` 会
  把这些副本收进来（README 警告过的自包含风险）。已移入
  `archive\旧产物_2026-11-08\`；`合并单元格清单.csv` 是交付件，留在原位。

## 1.1.1 — 2026-11-06

针对第二轮复审（N1–N9）的第一批修复。

### 修复：清理残留 Word 会误杀用户已打开的文档（P1，不可逆风险）
- 原实现按 `MainWindowTitle -eq ''` 全量 `Stop-Process -Force`，无法区分
  本工具新建的 COM 实例与用户自己打开的 Word（最小化/后台时标题可能为空）。
- 改为**PID 基线差集**：任务开始时快照现有 WINWORD PID，清理时只杀
  「不在基线内 且 无主窗口」的进程；拿不到基线时**一律不杀**（宁可留僵尸
  也不误伤）。
- 实测：打开一个用户文档（PID 19060）+ 新建一个无窗口实例（PID 19668）→
  清理后 `user_word_alive=True`、`new_hidden_alive=False`。
- 顺手统一 PowerShell 解释器：清理不再硬编码 `powershell`，与 `_find_powershell()` 一致。

### 修复：排序表部分填写会静默漏合并（P1，申报可追溯）
- `order=1` 时把排序表与实际目录做**集合比对**：目录内有、排序表未列出的
  文件即"不会被合并"，默认**拒绝写盘**并列出文件名，返回新状态
  `order_incomplete`。
- 新增 `--allow-partial-order` 显式放行（只告警），不建议用于申报资料。
- 排序表中若列出输出文件本身，自动跳过并告警。

### 新增：`--empty-page-auto {NO,YES,ASK}`
- 此前只有 GUI 设置 `EMPTY_PAGE_AUTO=NO`，CLI 默认进入交互分支：一旦检出
  残留空白页会弹 `MessageBox` 阻塞（超时上限 2700s）。现有显式参数，
  **默认 NO**（不交互、不删整页），GUI 未勾选时映射为 `ASK`（两者行为统一）。
- 说明：ps2 里"删空格字符"那段仍在 `$auto` 判断之前无条件执行，
  第二批会把它一并降级（见后续版本）。

### 修复：`--outfile` 位于输入目录的告警
- 判定从"直接父目录相同"改为"位于输入目录子树内"（子目录也提醒）；
- 文案修正为"本次输出已排除，但改名输出时历史产物仍会被收集"
  （原文案称"下次运行一定会重复合并"，与实际排除逻辑不符）。

### 测试
- 修正 `test_drop_removes_one_line_per_file` 的断言公式：实现里**每个**末行为
  独立 `}` 的文件都会被丢弃该行（含最后一个，末尾再补回 `}`），故应按全量
  `files` 计数，原来用 `files[:-1]` 是错的。
- 新增 `test_drop_on_synthetic_bare_braces`：3 个合成 bare-brace 文件做黄金
  基线（keep − drop = 3），不依赖真实样本，能暴露上面那个断言错误。
- 新增 `test_partial_order_refused`：排序表只填 1 行必须拒绝，放行参数生效时
  只告警。
- `openpyxl` 相关用例加 `skipUnless`，docstring 改为"除 order 用例外为零依赖"；
  真实样本目录支持 `RTF_MERGE_TEST_INPUT` 覆盖，缺失时 skip 而不是假绿。
- 结果：13 项全部 OK。

### 修复：合并清单里的版本号写死
- `merge_core.py` 曾自带一份 `__version__ = "1.1.0"`，与包版本脱钩，导致
  1.1.1 的产物清单仍写 `"version": "1.1.0"`。改为统一取 `rtfmerge.__version__`。

### 交付物
- `dist\README.txt` 同步为本版本说明；`dist\run_rtf_merge.bat` 重写（原为
  11-5 的旧占位符）：支持双击后粘贴/拖入文件夹，或把文件夹直接拖到 .bat 上，
  结果输出到桌面「合并结果.rtf / .docx」，默认 `--empty-page-auto NO`。
- 全量复跑（60 个 RTF，v1.1.1 的 exe 直接运行）：输出 36,086,787 字节、
  SHA256 `39f04383…2ab999`，与 1.1.0 逐字节一致；rFonts 88370、TNR/宋体
  各 100%、0 空白页、101 单元格合并、63 表格 AutoFit、230 中文引号、
  81 标题加粗，与改动前完全一致。

### 文档
- `word_post.py` 模块说明改为与实际一致：**ps2 会写盘**（删空格 + TOC 更新 +
  `SaveAs2`），之所以安全只是因为它在字体修正之前；ps3 严格只读。
- `step3_recheck.ps1` 注释删去"交 Python 重做字体修正"，改为"仅告警、不自愈"。

## 1.1.0 — 2026-11-06

### 空白页：改为从源头消除（P2-8）
- 新增 `drop_sect_gap`（默认 `True`）：合并时直接丢弃「文件间分节符前的空行」
  （SAS `put` 把置空的末行写出 1 个空格，正是 Word 中空白页的来源）。
  文件内部 SAS 自带的 `\sect` **一律不动**，每个表格仍从新页开始。
  A/B 实测：与 SAS 原样相比仅少 46 行（46 个文件的末行 `}` 空行），
  `\rtf1` 仍 60、`\sect` 386、`\pgnrestart` 0、花括号 47652/47652 平衡。
- `--keep-sect-gap` 可切回 SAS 原行为（用于对照）。

### 修复：字体修正被最终复检的重存冲掉（P0）
- 根因：`ps3` 一旦删了空格就 `SaveAs2`，Word 重写全部 OOXML 把
  `w:ascii="Times New Roman"` / `w:eastAsia="宋体"` 冲成主题字体
  （实测 7658 个 `rFonts` 中显式字体 0 处），程序仍报成功。
- 改动：`ps3` 改为**只检测、不写盘**；新增 `_verify_final_docx()` 最终验收，
  写盘结束后重新解析 docx + 页眉页脚 + styles，要求中西文字体覆盖率 ≥95%
  且目录域存在，否则抛错。
- 全量实测：278 页、0 空白页、`rFonts` 88370、TNR/宋体覆盖 100%。

### 修复：排序表全空却写出 0 字节并报成功（P0）
- `order` 列全空 → 返回 `no_ordered_files`；合并结果为空 → `empty_output`，
  两种情况都拒绝写盘。

### 修复：输入收集与输出位置（P0）
- `collect_rtf_files` 排除 `~$*`（Word 锁文件）、`.` 隐藏文件、以及 outfile 自身
  （用 `Path.resolve()` 比对），杜绝二次运行把产物再合并、指数膨胀。
- GUI 默认输出改为输入目录的**上一级**。

### 修复：边界条件破坏 RTF 结构（P1）
- 末文件为空被 `continue` 跳过时不再丢失补 `}` 与页码修正（新增收尾兜底块）。
- `maxn==1`（单文件）也走页码修正分支。
- 首/末行 `{`/`}` 缺失时记 warnings（RTF 结构 lint），不再静默。
- `delline` 删掉首行后，不再无条件执行 `lines[0][1:]`。

### 新增：可追溯与原子写（P1）
- 写盘改为「临时文件 + `os.replace` 原子替换」，并校验非空。
- 同时生成 `<输出名>_manifest.json`：输入清单 + SHA256、合并顺序、全部参数、
  资源表指纹分组、时间戳、输出哈希、告警。
- 新增资源表一致性校验（`\fonttbl`/`\colortbl` 指纹分组），不一致时告警
  （只告警，不改变输出结构）。

### 架构重构（P2）
- 拆分为 `rtfmerge` 包：
  `merge_core`（纯文本合并，可单测）/ `word_post`（Word COM 编排）/
  `docx_xml`（python-docx 层修正与验收）/ `rtf_lint` / `cli` / `gui`。
- 三段内嵌 PowerShell 外置为 `rtfmerge/ps/*.ps1`（可 lint、可单独调试）。
- `rtf_merge.py` / `rtf_merge_gui.py` 保留为兼容入口，旧调用路径不变。
- 新旧实现合并输出**逐字节一致**（36,086,787 字节，SHA256 相同）。

### 环境
- `subprocess.run` 全部显式 `encoding="utf-8"`。
- 新增 `requirements.txt`、`CHANGELOG.md`、`tests/`。

## 1.0.0 — 2026-11-03

- 首版：按 SAS 宏 `%rtf_mergeV4_1` 逐行复刻合并逻辑，10 文件子集与 60 文件全量
  均与 SAS 输出逐字节一致。
- Word 后处理链：RTF→docx、关键词单元格合并、正文表格 AutoFit、插入目录、
  字体修正、中文引号、标题加粗迁移。
- 交付：GUI（onedir）+ 命令行（onefile）双 exe，自定义图标。
