$ErrorActionPreference = 'Stop'
# 记录本脚本创建的 Word 实例 PID：创建前后各枚举一次 WINWORD 取差集，交给
# Python 侧在任务结束/取消/异常时精确回收。不依赖窗口标题、不依赖进程基线，
# 也不会波及用户自己打开的 Word（复用已有实例时差集为空，正好不记录）。
$wordPidBefore = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
                   ForEach-Object { $_.Id })
$word = New-Object -ComObject Word.Application
Start-Sleep -Milliseconds 300
$wordPidAfter = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
                  ForEach-Object { $_.Id })
$wordNewPid = @($wordPidAfter | Where-Object { $wordPidBefore -notcontains $_ })
if ($wordNewPid.Count -gt 0) { Write-Output ("WORD_PID=" + ($wordNewPid -join ",")) }
$word.Visible = $false
$word.DisplayAlerts = 0        # wdAlertsNone：抑制一切弹窗（防隐藏实例挂死泄漏）
try {
    $doc = $word.Documents.Open($args[0])

    # ===== 第 1 步：MergeCellWithAdjacentEmptyRight（AutoFit 之前）=====
    # 对每个正文表格：单元格文本（去段落符/单元格结束符后 trim）精确等于
    # 关键词，且其右侧紧邻单元格为空时，把两者水平合并（对应追加的 VBA
    # 功能；$doc.Tables 只含正文表格，页眉/页脚表格不受影响）
    function Get-CellText([object]$c) {
        return ($c.Range.Text -replace [char]13, '' -replace [char]7, '').Trim()
    }
    $kwSet = @{}
    foreach ($kw in @('N','BQL','Mean','Std','Median','CV(%)','Min','Max',
                      'GeoMean','GeoCV(%)','Geo CV(%)','Q1','Q3','BLQ')) {
        $kwSet[$kw] = $true
    }
    $merged = 0
    $mergeSkipped = 0
    foreach ($tbl in $doc.Tables) {
        $hits = New-Object System.Collections.ArrayList
        foreach ($cell in $tbl.Range.Cells) {
            if ($kwSet.ContainsKey((Get-CellText $cell))) {
                [void]$hits.Add(@{ row = [int]$cell.RowIndex; col = [int]$cell.ColumnIndex })
            }
        }
        # 从右到左处理，避免合并改变后续列索引（忠实还原 VBA 顺序）
        for ($i = $hits.Count - 1; $i -ge 0; $i--) {
            $r = $hits[$i].row
            $c = $hits[$i].col
            if ($c -ge 1 -and $c -lt $tbl.Columns.Count) {
                # 表格含垂直合并表头时，Cell(r,c) 可能取不到目标单元格（COM 抛
                # 错），此时跳过该命中并计数，不让整篇转换因此失败。
                try {
                    $right = $tbl.Cell($r, $c + 1)
                    if ((Get-CellText $right) -eq '') {
                        $tbl.Cell($r, $c).Merge($right)
                        $merged++
                    }
                } catch {
                    $mergeSkipped++
                }
            }
        }
    }
    Write-Output ("MERGE_CELLS=" + $merged)
    Write-Output ("MERGE_SKIPPED=" + $mergeSkipped)

    # ===== 第 2 步：按内容并铺满窗口自动调整正文所有表格 =====
    # （wdAutoFitContent=1 后再 wdAutoFitWindow=2：先按内容分配列宽比例，
    # 再按页面宽度铺满，对应 AutoFitTablesContent + AutoFitTablesWindow
    # 两个 VBA 功能）；$doc.Tables 只含正文表格，页眉/页脚中的表格不在
    # 其中，天然不受影响
    $n = 0
    foreach ($tbl in $doc.Tables) {
        $tbl.AutoFitBehavior([int]1)
        $tbl.AutoFitBehavior([int]2)
        $n++
    }
    Write-Output ("AUTOFIT_TABLES=" + $n)

    # ===== 第 3 步：插入目录（对应 InsertTableOfContents VBA 功能）=====
    # 文档已有“目录”二字时在其后插入目录；否则在文档开头创建“目录”页：
    # “目录”标题（居中、宋体、16 磅加粗）+ 目录域 + 分页符（正文另起一页）。
    # 目录基于段落大纲级别抓取：包含所有非“正文文本”级别（Outline 1~9，
    # 对应 outlineLvl 0~8）的标题——不仅是 1~3 级，4 级及更深标题也进目录。
    # 目录内容格式：宋体 + Times New Roman + 小四(12 磅)
    # + 蓝色 + 不加粗 + 点状引导符 + 含页码。
    $tocRng = $null
    $created = $false
    # 定位“目录”页标题：兼容“目 录”（两字间可有任意个半角/全角空格或
    # 制表符）与紧邻的“目录”两种写法；都找不到才在文档开头新建目录页。
    $f1 = $doc.Content
    $f1.Find.Forward = $true
    $f1.Find.Wrap = 0              # wdFindStop
    $f1.Find.MatchWildcards = $true
    $f1.Find.Text = "目[ 　^t]@录"
    if ($f1.Find.Execute()) {
        $tocRng = $f1
        $tocRng.Collapse(0)        # 折叠到“目录”二字之后
    } else {
        $f2 = $doc.Content
        $f2.Find.Forward = $true
        $f2.Find.Wrap = 0
        $f2.Find.Text = "目录"
        if ($f2.Find.Execute()) {
            $tocRng = $f2
            $tocRng.Collapse(0)
        } else {
        # 无“目录”页：在文档开头创建。注意必须先 InsertParagraphBefore
        # 建新段落再写“目录”，否则文字会并入原第一段（其带大纲级别，
        # 会被目录自身抓取）
        $startRng = $doc.Range(0, 0)
        $startRng.InsertParagraphBefore()
        $titleRng = $doc.Range(0, 0)
        $titleRng.Text = "目录"
        $titleRng.ParagraphFormat.Alignment = 1   # wdAlignParagraphCenter
        $titleRng.ParagraphFormat.OutlineLevel = 10  # wdOutlineLevelBodyText
        $titleRng.Font.NameFarEast = "宋体"
        $titleRng.Font.NameAscii = "Times New Roman"
        $titleRng.Font.Size = 16
        $titleRng.Font.Bold = $true
        $titleRng.Collapse(0)
        $titleRng.InsertParagraphAfter()
        $titleRng.Collapse(0)
        $tocRng = $titleRng
        $created = $true
        }
    }
    if ($doc.TablesOfContents.Count -gt 0) {
        # 已有目录（输入文档自带 TOC）：保留原域与位置，只刷新内容，
        # 不再“删除重建”，避免丢失原有目录样式/分页设置。
        $toc = $doc.TablesOfContents.Item(1)
        $toc.Update()
    } else {
        $toc = $doc.TablesOfContents.Add(
            $tocRng, $true, 1, 9, $false, "", $true, $true, "", $true, $true, $true)
        $toc.TabLeader = 1          # wdTabLeaderDots
        $toc.Update()
    }
    $tocFont = $toc.Range.Font
    $tocFont.NameFarEast = "宋体"
    $tocFont.NameAscii = "Times New Roman"
    $tocFont.Size = 12
    $tocFont.Color = 16711680       # 蓝色：VBA RGB(0,0,255)=0xFF0000（COLORREF）
    $tocFont.Bold = $false
    if ($created) {
        $tocEnd = $toc.Range
        $tocEnd.Collapse(0)
        $tocEnd.InsertBreak(7)      # wdPageBreak：正文另起一页
    }
    Write-Output ("TOC_INSERTED=" + $doc.TablesOfContents.Count)
    $doc.SaveAs2($args[1], 16)          # wdFormatDocumentDefault (.docx)
    $doc.Close(0)
} finally {
    try { $doc.Close(0) } catch {}
    $word.Quit()
    try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null } catch {}
}
