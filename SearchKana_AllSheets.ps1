# SearchKana_Fast.ps1
# 递归搜索 Excel：所有工作表的 C/D 列中查找包含「カナ」的单元格（数组扫描，速度快）
# 输出到 result.csv / result.txt （UTF-8 with BOM，日文不乱码）
# 需要本机安装 Microsoft Excel（COM）

param(
    [string]$Root = (Get-Location).Path,
    [string]$Keyword = "カナ",
    [string]$OutCsv = "result.csv",
    [string]$OutTxt = "result.txt"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$results = New-Object System.Collections.Generic.List[object]
$excelFiles = Get-ChildItem -Path $Root -Recurse -File -Include *.xlsx, *.xlsm, *.xls

if ($excelFiles.Count -eq 0) {
    Write-Host "未找到 Excel 文件：$Root"
    exit 0
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

function Release-ComObject([object]$o) {
    if ($null -ne $o) {
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($o) } catch {}
    }
}

try {
    foreach ($f in $excelFiles) {
        $wb = $null
        try {
            $wb = $excel.Workbooks.Open($f.FullName, 0, $true)  # ReadOnly

            foreach ($ws in $wb.Worksheets) {
                $sheetName = $ws.Name

                # 用 UsedRange 推断实际行范围（比扫整列快）
                $used = $null
                try { $used = $ws.UsedRange } catch { $used = $null }
                if ($null -eq $used) { Release-ComObject $ws; continue }

                $firstRow = $used.Row
                $rowCount = $used.Rows.Count
                if ($rowCount -lt 1) { Release-ComObject $used; Release-ComObject $ws; continue }

                $lastRow = $firstRow + $rowCount - 1

                # 一次性读 C:D（两列）
                $range = $ws.Range("C$firstRow:D$lastRow")
                $vals  = $range.Value2

                # Value2 可能返回：
                # - 2D object[,]（多行）
                # - scalar（只有一个单元格时）
                # 我们统一处理
                if ($vals -is [System.Array] -and $vals.Rank -eq 2) {
                    $rMin = $vals.GetLowerBound(0); $rMax = $vals.GetUpperBound(0)
                    $cMin = $vals.GetLowerBound(1); $cMax = $vals.GetUpperBound(1) # 理论上是 2 列

                    for ($ri = $rMin; $ri -le $rMax; $ri++) {
                        for ($ci = $cMin; $ci -le $cMax; $ci++) {
                            $v = $vals[$ri, $ci]
                            if ($null -ne $v) {
                                $s = [string]$v
                                if ($s -like "*$Keyword*") {
                                    $actualRow = $firstRow + ($ri - $rMin)
                                    $colLetter = if (($ci - $cMin) -eq 0) { "C" } else { "D" }

                                    $results.Add([pscustomobject]@{
                                        FileName = $f.Name
                                        FullPath = $f.FullName
                                        Sheet    = $sheetName
                                        Row      = $actualRow
                                        Column   = $colLetter
                                        Value    = $s
                                    })
                                }
                            }
                        }
                    }
                }
                else {
                    # 只有一个单元格的极端情况：它就是 C(firstRow)
                    if ($null -ne $vals) {
                        $s = [string]$vals
                        if ($s -like "*$Keyword*") {
                            $results.Add([pscustomobject]@{
                                FileName = $f.Name
                                FullPath = $f.FullName
                                Sheet    = $sheetName
                                Row      = $firstRow
                                Column   = "C"   # 单格情况下 Range 从 C:D 仍可能退化，按 C 记录
                                Value    = $s
                            })
                        }
                    }
                }

                Release-ComObject $range
                Release-ComObject $used
                Release-ComObject $ws
            }

            $wb.Close($false) | Out-Null
            Release-ComObject $wb
        }
        catch {
            $results.Add([pscustomobject]@{
                FileName = $f.Name
                FullPath = $f.FullName
                Sheet    = ""
                Row      = ""
                Column   = ""
                Value    = "ERROR: $($_.Exception.Message)"
            })

            if ($wb -ne $null) {
                try { $wb.Close($false) | Out-Null } catch {}
                Release-ComObject $wb
            }
        }
    }
}
finally {
    try { $excel.Quit() } catch {}
    Release-ComObject $excel
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

# 写文件：UTF-8 BOM（Excel/记事本都不乱码）
$utf8Bom = New-Object System.Text.UTF8Encoding($true)

# CSV（Excel 直接打开）
$csvText = $results | ConvertTo-Csv -NoTypeInformation
[System.IO.File]::WriteAllLines((Join-Path $Root $OutCsv), $csvText, $utf8Bom)

# TXT（制表符分隔）
$txtLines = New-Object System.Collections.Generic.List[string]
$txtLines.Add("FileName`tFullPath`tSheet`tRow`tColumn`tValue")
foreach ($x in $results) {
    $v = ($x.Value -replace "(\r\n|\n|\r)", " ")
    $txtLines.Add("$($x.FileName)`t$($x.FullPath)`t$($x.Sheet)`t$($x.Row)`t$($x.Column)`t$v")
}
[System.IO.File]::WriteAllLines((Join-Path $Root $OutTxt), $txtLines, $utf8Bom)

Write-Host "完成：命中 $($results.Count) 条"
Write-Host "CSV: $(Join-Path $Root $OutCsv)"
Write-Host "TXT: $(Join-Path $Root $OutTxt)"