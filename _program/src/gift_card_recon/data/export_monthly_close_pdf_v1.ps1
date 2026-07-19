# GC Recon Excel PDF exporter, interface version 1.
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $true)]
    [string]$PdfPath,

    [Parameter(Mandatory = $false)]
    [string]$WorksheetName = "Monthly Close Report"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$excel = $null
$workbooks = $null
$workbook = $null
$worksheets = $null
$worksheet = $null
$exitCode = 0
$failureMessage = ""

function Release-ComObject {
    param([object]$ComObject)

    if (($null -ne $ComObject) -and [System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        try {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
        }
        catch {
            # Excel Close/Quit outcomes below determine process success. COM
            # release is a final best-effort cleanup after those calls.
        }
    }
}

try {
    $sourcePath = [System.IO.Path]::GetFullPath($WorkbookPath)
    $destinationPath = [System.IO.Path]::GetFullPath($PdfPath)

    if (-not [System.IO.File]::Exists($sourcePath)) {
        throw "Workbook not found: $sourcePath"
    }

    $destinationDirectory = [System.IO.Path]::GetDirectoryName($destinationPath)
    if (-not [string]::IsNullOrWhiteSpace($destinationDirectory)) {
        [System.IO.Directory]::CreateDirectory($destinationDirectory) | Out-Null
    }
    if ([System.IO.File]::Exists($destinationPath)) {
        [System.IO.File]::Delete($destinationPath)
    }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false
    # 3 = msoAutomationSecurityForceDisable; -4135 = xlCalculationManual.
    $excel.AutomationSecurity = 3
    $excel.Calculation = -4135

    $workbooks = $excel.Workbooks
    $workbook = $workbooks.Open($sourcePath, 0, $true)
    $worksheets = $workbook.Worksheets
    $worksheet = $worksheets.Item($WorksheetName)

    # 0 = xlTypePDF and xlQualityStandard. Calling ExportAsFixedFormat on the
    # worksheet (not the workbook) prevents support tabs from entering the PDF.
    $worksheet.ExportAsFixedFormat(0, $destinationPath, 0, $true, $false)

    if (-not [System.IO.File]::Exists($destinationPath)) {
        throw "Excel did not create the requested PDF: $destinationPath"
    }
    if ((Get-Item -LiteralPath $destinationPath).Length -le 0) {
        throw "Excel created an empty PDF: $destinationPath"
    }
}
catch {
    $exitCode = 1
    $failureMessage = $_.Exception.Message
}
finally {
    if ($null -ne $workbook) {
        try {
            $workbook.Close($false)
        }
        catch {
            if ($exitCode -eq 0) {
                $exitCode = 1
                $failureMessage = "Could not close the Excel workbook: $($_.Exception.Message)"
            }
        }
    }
    if ($null -ne $excel) {
        try {
            $excel.Quit()
        }
        catch {
            if ($exitCode -eq 0) {
                $exitCode = 1
                $failureMessage = "Could not close Excel: $($_.Exception.Message)"
            }
        }
    }

    Release-ComObject $worksheet
    Release-ComObject $worksheets
    Release-ComObject $workbook
    Release-ComObject $workbooks
    Release-ComObject $excel
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

if ($exitCode -ne 0) {
    [Console]::Error.WriteLine($failureMessage)
}
exit $exitCode
