$logFile = "C:\Users\emily.ott\Desktop\MacroRunLog.txt"
Start-Transcript -Path $logFile -Append
Write-Output "Task started at $(Get-Date)"


# RunAllMacros.ps1
# ------------------------------------------------------------
# Purpose: Open an Excel workbook and run multiple VBA macros
# ------------------------------------------------------------

# Define workbook and macro list
$excelFile = "C:\Users\emily.ott\My SecuriSync\Tier 1\Office Share\FRN\FRN_Tracking_2025.xlsm"

# The workbook name in each macro reference must match the actual file name ("FRN_Tracking_2025.xlsm")
$macros = @(
    "FRN_Tracking_2025.xlsm!Module1.SendEmailIfDeadlineIn7DaysOrLess",
    "FRN_Tracking_2025.xlsm!Module2.SendEmailIfDeadlineIn7DaysOrLess",
    "FRN_Tracking_2025.xlsm!Module3.SendEmailOnDeadlineReminder_7DaysOrLess",
    "FRN_Tracking_2025.xlsm!Module4.SendEmailIfDeadline2DaysOrLess",
    "FRN_Tracking_2025.xlsm!Module5.SendEmailOnDeadlineReminder_7DaysOrLess",
    "FRN_Tracking_2025.xlsm!Module6.SendEmailIfDeadlineIn7DaysOrLess",
    "FRN_Tracking_2025.xlsm!Module7.SendEmailIfDeadlineIn7DaysOrLess"
)

# Optional: path to a log file
$logFile = "C:\Scripts\MacroRunLog.txt"
Start-Transcript -Path $logFile -Append

try {
    Write-Host "Starting Excel..."
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false

    Write-Host "Opening workbook: $excelFile"
    $workbook = $excel.Workbooks.Open($excelFile)

    foreach ($macro in $macros) {
        Write-Host "Running macro: $macro"
        try {
            $excel.Run($macro)
            Write-Host "✅ Successfully ran $macro"
        }
        catch {
            # use ${macro} to avoid ':' parsing error
            Write-Warning "❌ Error running macro ${macro}: $($_.Exception.Message)"
        }
    }

    Write-Host "Saving and closing workbook..."
    $workbook.Close($true)

    Write-Host "Quitting Excel..."
    $excel.Quit()
}
catch {
    Write-Warning "❌ An unexpected error occurred: $($_.Exception.Message)"
}
finally {
    # Clean up COM objects
    if ($workbook) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null }
    if ($excel)    { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Stop-Transcript
}

Write-Host "🎉 All macros executed successfully."

Write-Output "Task finished at $(Get-Date)"
Stop-Transcript
