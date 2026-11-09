$ErrorActionPreference = 'Stop'
# 记录本脚本创建的 Word 实例 PID（创建前后枚举差集），供 Python 侧精确回收。
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
Write-Output "STEP=word_started"
Start-Sleep -Seconds 3         # 等 Word COM 服务器就绪，防 RPC_E_CALL_REJECTED
try {
    # 1 参 Open（与 ps1 完全相同，PowerShell 5.1 多参 COM 绑定不稳——实测
    # 3 参/7 参 Open 会挂起或报“bool→Object”参数错误，1 参在 ps1 中稳定成功）。
    # 加 5 次重试兜底 RPC 拒绝。
    $doc = $null
    for ($i = 1; $i -le 5; $i++) {
        try {
            $doc = $word.Documents.Open($args[0])
            break
        } catch {
            Write-Output ("OPEN_RETRY_" + $i)
            if ($i -ge 5) { throw }
            Start-Sleep -Seconds 3
        }
    }
    Write-Output "STEP=doc_opened"

    # ===== 本脚本只检测、不写盘 =====
    # 历史上这里做过三件事，现已全部移出：
    #   1) Remove-AllSpacesInEmptyPages：删除空白页内“仅含空格”段落的空格。
    #      合并阶段 drop_sect_gap 已从 RTF 源头去掉该空行后，此函数实测产出
    #      0（SPACE_CHARS_DELETED=0），属空转；保留只会多一次全页扫描。
    #   2) 交互删除整页 / EMPTY_PAGE_AUTO=YES 自动删页：源头治理后已无空白页，
    #      且 MessageBox 会让无人值守运行挂死（最长 2700s）。
    #   3) Update-TocBlue + SaveAs2：TOC 1~9 的样式级蓝色已迁移到
    #      docx_xml._fix_toc_styles（直接改 styles.xml，不依赖 Word 写盘）。
    #      Word 在这里 SaveAs2 会重写全部 OOXML，把 python-docx 写入的显式
    #      字体冲成主题字体（实测 7658 个 rFonts 中 Times New Roman 变 0 处）。
    # 现在只做全页扫描并把页数/空白页数报给 Python 侧，供日志与清单留痕。

    function Get-EmptyPages([object]$d) {
        # 全页扫描（原版做法）。注意不要改成 Sections 遍历定位——
        # 实测 $sec.Range.Paragraphs.Last 会卡死数十分钟。
        $total = $d.ComputeStatistics(2)        # wdStatisticPages
        $list = New-Object System.Collections.ArrayList
        for ($pn = 1; $pn -le $total; $pn++) {
            $rg = $d.GoTo(1, 1, $pn)            # wdGoToPage, wdGoToAbsolute
            $nx = $d.GoTo(1, 1, $pn + 1)
            if ($nx.Start -gt $rg.Start) {
                $rg.End = $nx.Start - 1
            } else {
                $rg.End = $d.Content.End        # 最后一页：覆盖到文档末尾
            }
            if ($rg.InlineShapes.Count -gt 0) { continue }   # 含图页非空白
            $txt = ($rg.Text -replace [char]13, '' -replace [char]7, '' `
                    -replace [char]12, '' -replace ' ', '' -replace "`t", '').Trim()
            if ($txt -eq '') { [void]$list.Add([int]$pn) }
        }
        return $list
    }

    $empties = @(Get-EmptyPages $doc)
    foreach ($pn in $empties) { Write-Output ("EMPTY_PAGE=" + $pn) }
    Write-Output ("EMPTY_PAGES_COUNT=" + $empties.Count)
    # 兼容旧字段：检测到但未处理（本脚本从不删页）
    Write-Output ("EMPTY_PAGES_SKIPPED=" + $empties.Count)
    Write-Output ("EMPTY_PAGES_DELETED=0")
    Write-Output ("SPACE_CHARS_DELETED=0")
    Write-Output ("DOC_PAGES=" + $doc.ComputeStatistics(2))
    Write-Output ("WORD_VERSION=" + $word.Version)
    $doc.Close(0)
    Write-Output "STEP=closed"
} finally {
    try { $doc.Close(0) } catch {}
    $word.Quit()
    try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null } catch {}
}
