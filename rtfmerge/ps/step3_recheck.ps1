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
Start-Sleep -Seconds 3         # 等 Word COM 服务器就绪，防 RPC_E_CALL_REJECTED
try {
    $doc = $null
    for ($i = 1; $i -le 5; $i++) {
        try {
            # 1 参 Open，与 step1/step2 完全一致。PowerShell 5.1 下多参 COM
            # 绑定不稳：实测 3 参/7 参 Open 会挂起或报“bool→Object”参数错误，
            # 1 参稳定成功（本文件此前用 3 参，与 ps2 注释矛盾，已统一）。
            $doc = $word.Documents.Open($args[0])
            break
        } catch {
            if ($i -ge 5) { throw }
            Start-Sleep -Seconds 3
        }
    }

    function Get-EmptyPages([object]$d) {
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

    function Get-PageRange([object]$d, [int]$pn) {
        $total = $d.ComputeStatistics(2)
        if ($pn -lt 1 -or $pn -gt $total) { return $null }
        $rg = $d.GoTo(1, 1, $pn)
        $nx = $d.GoTo(1, 1, $pn + 1)
        if ($nx.Start -gt $rg.Start) {
            $rg.End = $nx.Start - 1
        } else {
            $rg.End = $d.Content.End
        }
        return $rg
    }

    # 只检测、不修改：合并阶段已在 RTF 源头去掉“文件间分节符前的空行”
    # （drop_sect_gap），空白页理论上不会再出现。此处仅统计并报告，绝不写盘
    # ——一旦在此 SaveAs2，Word 会重写全部 OOXML，把 python-docx 写入的
    # 显式字体（宋体/Times New Roman）冲成主题字体（实测：7658 个 rFonts 中
    # w:ascii="Times New Roman" 变成 0 处）。
    # 若确有残留：本脚本只输出 EMPTY_PAGES_COUNT 供 Python 侧告警，不会自愈，
    # 也不会重做字体修正；需人工在 Word 中核对或调整上游（drop_sect_gap / ps2）。
    $empties = @(Get-EmptyPages $doc)
    foreach ($pn in $empties) { Write-Output ("EMPTY_PAGE=" + $pn) }
    Write-Output ("EMPTY_PAGES_COUNT=" + $empties.Count)
    Write-Output ("DOC_PAGES=" + $doc.ComputeStatistics(2))
    Write-Output ("SPACE_CHARS_DELETED=0")
    $doc.Close(0)
} finally {
    try { $doc.Close(0) } catch {}
    $word.Quit()
    try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null } catch {}
}
