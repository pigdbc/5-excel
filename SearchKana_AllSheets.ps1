# Export_CE_FromRow7_MergeAware.ps1
# 从每个工作表第7行开始导出：
# A=文件名 B=sheet名 C=源Sheet的C列(支持合并单元格按显示填充) D空 E=源Sheet的E列(同样支持合并)
# 输出 result.xlsx
# 需要本机安装 Microsoft Excel（COM）

param(
    [string]$Root = (Get-Location).Path,
    [string]$OutXlsx = "result.xlsx",
    [int]$StartRow = 7
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Release-ComObject([object]$o) {
    if ($null -ne $o) {
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($o) } catch {}
    }
}

# 找 Excel 文件（包含子文件夹，过滤临时锁文件）
$excelFiles = Get-ChildItem -Path $Root -Recurse -File -Include *.xlsx, *.xlsm, *.xls |
    Where-Object { $_.Name -notlike "~$*" }

Write-Host "扫描目录: $Root"
Write-Host "找到 Excel 文件数: $($excelFiles.Count)"

if ($excelFiles.Count -eq 0) {
    Write-Host "没有找到任何 Excel。"
    exit 0
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$outWb = $excel.Workbooks.Add()
$outWs = $outWb.Worksheets.Item(1)
$outWs.Name = "result"

# 表头
$outWs.Cells.Item(1,1).Value2 = "FileName"
$outWs.Cells.Item(1,2).Value2 = "Sheet"
$outWs.Cells.Item(1,3).Value2 = "ColC"
$outWs.Cells.Item(1,4).Value2 = ""      # D 空
$outWs.Cells.Item(1,5).Value2 = "ColE"

$outRow = 2

try {
    foreach ($f in $excelFiles) {
        $wb = $null
        try {
            $wb = $excel.Workbooks.Open($f.FullName, 0, $true)  # ReadOnly

            foreach ($ws in $wb.Worksheets) {
                $sheetName = $ws.Name

                # 找最后一个非空单元格，确定 lastRow
                $lastCell = $ws.Cells.Find(
                    "*",
                    $ws.Cells.Item(1,1),
                    -4163,   # xlValues
                    1,       # xlPart
                    1,       # xlByRows
                    2,       # xlPrevious
                    $false
                )

                if ($null -eq $lastCell) {
                    Release-ComObject $ws
                    continue
                }

                $lastRow = $lastCell.Row
                Release-ComObject $lastCell

                if ($lastRow -lt $StartRow) {
                    Release-ComObject $ws
                    continue
                }

                # 一次性读 C:E（3列），从第7行开始
                $rng = $ws.Range("C$StartRow:E$lastRow")
                $vals = $rng.Value2

                # 合并填充用的“区间缓存”（避免每行都去问 MergeArea）
                $carryCValue = ""
                $carryCUntil = 0
                $carryEValue = ""
                $carryEUntil = 0

                if ($vals -is [System.Array] -and $vals.Rank -eq 2) {
                    $rMin = $vals.GetLowerBound(0); $rMax = $vals.GetUpperBound(0)

                    for ($ri = $rMin; $ri -le $rMax; $ri++) {
                        $actualRow = $StartRow + ($ri - $rMin)

                        # C:E => 1=C, 2=D(忽略), 3=E
                        $cVal = $vals[$ri, 1]
                        $eVal = $vals[$ri, 3]

                        # --- 处理 C 列合并 ---
                        $cOut = ""
                        if ($null -ne $cVal -and [string]$cVal -ne "") {
                            $cOut = [string]$cVal
                            $carryCValue = $cOut
                            $carryCUntil = 0
                        }
                        elseif ($actualRow -le $carryCUntil -and $carryCValue -ne "") {
                            # 在已知合并区间内，填充
                            $cOut = $carryCValue
                        }
                        else {
                            # 检查这一行的 C 单元格是否处于合并区域
                            $cellC = $ws.Cells.Item($actualRow, 3)
                            if ($cellC.MergeCells) {
                                $area = $cellC.MergeArea
                                $topVal = $area.Cells.Item(1,1).Value2
                                $cOut = if ($null -ne $topVal) { [string]$topVal } else { "" }
                                $carryCValue = $cOut
                                $carryCUntil = $area.Row + $area.Rows.Count - 1
                                Release-ComObject $area
                            }
                            Release-ComObject $cellC
                        }

                        # --- 处理 E 列合并（可选但更稳）---
                        $eOut = ""
                        if ($null -ne $eVal -and [string]$eVal -ne "") {
                            $eOut = [string]$eVal
                            $carryEValue = $eOut
                            $carryEUntil = 0
                        }
                        elseif ($actualRow -le $carryEUntil -and $carryEValue -ne "") {
                            $eOut = $carryEValue
                        }
                        else {
                            $cellE = $ws.Cells.Item($actualRow, 5)
                            if ($cellE.MergeCells) {
                                $areaE = $cellE.MergeArea
                                $topValE = $areaE.Cells.Item(1,1).Value2
                                $eOut = if ($null -ne $topValE) { [string]$topValE } else { "" }
                                $carryEValue = $eOut
                                $carryEUntil = $areaE.Row + $areaE.Rows.Count - 1
                                Release-ComObject $areaE
                            }
                            Release-ComObject $cellE
                        }

                        # 写入结果
                        $outWs.Cells.Item($outRow, 1).Value2 = $f.Name
                        $outWs.Cells.Item($outRow, 2).Value2 = $sheetName
                        $outWs.Cells.Item($outRow, 3).Value2 = $cOut
                        # D 空着
                        $outWs.Cells.Item($outRow, 5).Value2 = $eOut
                        $outRow++
                    }
                }

                Release-ComObject $rng
                Release-ComObject $ws
            }

            $wb.Close($false) | Out-Null
            Release-ComObject $wb
        }
        catch {
            # 文件打不开/被锁/损坏：写一行错误
            $outWs.Cells.Item($outRow, 1).Value2 = $f.Name
            $outWs.Cells.Item($outRow, 2).Value2 = ""
            $outWs.Cells.Item($outRow, 3).Value2 = "ERROR: $($_.Exception.Message)"
            $outWs.Cells.Item($outRow, 5).Value2 = ""
            $outRow++

            if ($wb -ne $null) {
                try { $wb.Close($false) | Out-Null } catch {}
                Release-ComObject $wb
            }
        }
    }

    $outWs.Columns.AutoFit() | Out-Null

    $outPath = Join-Path $Root $OutXlsx
    if (Test-Path $outPath) { Remove-Item -Force $outPath }
    $outWb.SaveAs($outPath) | Out-Null

    Write-Host "完成：写入行数 = $($outRow - 2)"
    Write-Host "输出文件：$outPath"
}
finally {
    try { $outWb.Close($true) | Out-Null } catch {}
    Release-ComObject $outWs
    Release-ComObject $outWb

    try { $excel.Quit() } catch {}
    Release-ComObject $excel

    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}