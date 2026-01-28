# Export_CE_To_ResultExcel.ps1
# A=源文件名, B=Sheet名, C=源Sheet的C列(逐行), D空, E=源Sheet的E列(逐行)
# 输出 result.xlsx

param(
    [string]$Root = (Get-Location).Path,
    [string]$OutXlsx = "result.xlsx"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Release-ComObject([object]$o) {
    if ($null -ne $o) {
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($o) } catch {}
    }
}

# 1) 找Excel文件（过滤掉临时锁文件）
$excelFiles = Get-ChildItem -Path $Root -Recurse -File -Include *.xlsx, *.xlsm, *.xls |
    Where-Object { $_.Name -notlike "~$*" }

if ($excelFiles.Count -eq 0) {
    Write-Host "未找到 Excel 文件：$Root"
    exit 0
}

# 2) 启动Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 3) 新建结果工作簿
$outWb = $excel.Workbooks.Add()
$outWs = $outWb.Worksheets.Item(1)
$outWs.Name = "result"

# 表头（可删）
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

                # 用 UsedRange 决定要导出的行范围
                $used = $null
                try { $used = $ws.UsedRange } catch { $used = $null }
                if ($null -eq $used) { Release-ComObject $ws; continue }

                $firstRow = $used.Row
                $rowCount = $used.Rows.Count
                if ($rowCount -lt 1) { Release-ComObject $used; Release-ComObject $ws; continue }

                $lastRow = $firstRow + $rowCount - 1

                # 一次性读取 C 和 E 两列
                $rngC = $ws.Range("C$firstRow:C$lastRow")
                $rngE = $ws.Range("E$firstRow:E$lastRow")
                $valC = $rngC.Value2
                $valE = $rngE.Value2

                # 把 Range.Value2 统一转成“按行取值”的方式
                # 多行时是二维数组 [1..n,1..1]，单行时可能是 scalar
                if ($valC -is [System.Array] -and $valC.Rank -eq 2) {
                    $rMin = $valC.GetLowerBound(0); $rMax = $valC.GetUpperBound(0)
                    for ($ri = $rMin; $ri -le $rMax; $ri++) {
                        $cStr = $valC[$ri, 1]
                        $eStr = $null
                        if ($valE -is [System.Array] -and $valE.Rank -eq 2) {
                            $eStr = $valE[$ri, 1]
                        } else {
                            $eStr = $valE
                        }

                        # 可选：如果 C 和 E 都空，就不写（不想过滤就删掉这个 if）
                        if ($null -eq $cStr -and $null -eq $eStr) { continue }

                        $outWs.Cells.Item($outRow, 1).Value2 = $f.Name
                        $outWs.Cells.Item($outRow, 2).Value2 = $sheetName
                        $outWs.Cells.Item($outRow, 3).Value2 = if ($null -ne $cStr) { [string]$cStr } else { "" }
                        # D列留空
                        $outWs.Cells.Item($outRow, 5).Value2 = if ($null -ne $eStr) { [string]$eStr } else { "" }
                        $outRow++
                    }
                } else {
                    # 单行/单格退化情况
                    if ($null -ne $valC -or $null -ne $valE) {
                        $outWs.Cells.Item($outRow, 1).Value2 = $f.Name
                        $outWs.Cells.Item($outRow, 2).Value2 = $sheetName
                        $outWs.Cells.Item($outRow, 3).Value2 = if ($null -ne $valC) { [string]$valC } else { "" }
                        $outWs.Cells.Item($outRow, 5).Value2 = if ($null -ne $valE) { [string]$valE } else { "" }
                        $outRow++
                    }
                }

                Release-ComObject $rngC
                Release-ComObject $rngE
                Release-ComObject $used
                Release-ComObject $ws
            }

            $wb.Close($false) | Out-Null
            Release-ComObject $wb
        }
        catch {
            # 如果某个文件打不开，也写一行错误到结果里
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

    # 简单美化：自动列宽
    $outWs.Columns.AutoFit() | Out-Null

    # 保存
    $outPath = Join-Path $Root $OutXlsx
    # 如果已存在，先删掉避免 SaveAs 异常
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