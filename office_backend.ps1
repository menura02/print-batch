param([Parameter(Mandatory=$true)][string]$Request)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$job = Get-Content -LiteralPath $Request -Raw -Encoding UTF8 | ConvertFrom-Json
$app = $null
$document = $null
$collection = $null
try {
    if ($job.action -eq 'printers') {
        $printers = @(Get-Printer | Sort-Object Name | ForEach-Object { $_.Name })
        $default = (Get-CimInstance Win32_Printer | Where-Object Default | Select-Object -First 1).Name
        @{ printers = $printers; default = $default } | ConvertTo-Json -Compress | Write-Output
        exit 0
    }
    $source = [System.IO.Path]::GetFullPath($job.source)
    $target = [System.IO.Path]::GetFullPath($job.target)
    if (!(Test-Path -LiteralPath $source -PathType Leaf)) { throw 'Dosya bulunamadi.' }
    $ext = [System.IO.Path]::GetExtension($source).ToLowerInvariant()
    switch ($ext) {
        { $_ -in '.doc', '.docx' } {
            $app = New-Object -ComObject Word.Application
            $app.AutomationSecurity = 3
            $app.Visible = $false
            $app.DisplayAlerts = 0
            $collection = $app.Documents
            $document = $collection.Open($source, $false, $true, $false, '', '', $false, '', '', 0, 0, $false, $false, 0, $true)
            $document.ExportAsFixedFormat($target, 17)
        }
        { $_ -in '.xls', '.xlsx' } {
            $app = New-Object -ComObject Excel.Application
            $app.AutomationSecurity = 3
            $app.Visible = $false
            $app.DisplayAlerts = $false
            $app.AskToUpdateLinks = $false
            $collection = $app.Workbooks
            $document = $collection.Open($source, 0, $true, [Type]::Missing, '', '', $true)
            $document.ExportAsFixedFormat(0, $target, 0, $true, $false)
        }
        { $_ -in '.ppt', '.pptx' } {
            $app = New-Object -ComObject PowerPoint.Application
            $app.AutomationSecurity = 3
            $app.DisplayAlerts = 1
            $collection = $app.Presentations
            $document = $collection.Open($source, -1, 0, 0)
            $document.SaveAs($target, 32)
        }
        default { throw 'Desteklenmeyen belge turu.' }
    }
    if (!(Test-Path -LiteralPath $target) -or (Get-Item -LiteralPath $target).Length -eq 0) { throw 'PDF olusturulamadi.' }
    @{ ok = $true; target = $target } | ConvertTo-Json -Compress | Write-Output
}
catch {
    @{ ok = $false; error = $_.Exception.Message } | ConvertTo-Json -Compress | Write-Output
    exit 1
}
finally {
    if ($null -ne $document) {
        try { if ($ext -in '.ppt', '.pptx') { $document.Close() } else { $document.Close($false) } } catch {}
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
    }
    if ($null -ne $collection) { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($collection) }
    if ($null -ne $app) {
        try { $app.Quit() } catch {}
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
    }
}
