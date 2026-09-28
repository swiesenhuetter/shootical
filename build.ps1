# Builds the PSUE calendar (HTML page + iCalendar files) from the Excel
# Schiesskalender into .\output (gitignored), named as on the website.
#
#   .\build.ps1                                  # data\PSUE_Schiesskalender<this year>.xlsx
#   .\build.ps1 -Excel data\PSUE_Schiesskalender2027.xlsx
#   .\build.ps1 -Website ..\psue_website_new\calendar   # also copy to the website

param(
    [string]$Excel = "data\PSUE_Schiesskalender$((Get-Date).Year).xlsx",
    [string]$Website
)

$ErrorActionPreference = 'Stop'
Push-Location $PSScriptRoot  # cal2html.py reads the template from the current directory
try {

    function Invoke-Checked {
        $cmd, $rest = $args
        & $cmd @rest
        if ($LASTEXITCODE -ne 0) { throw "Failed ($LASTEXITCODE): $args" }
    }

    if (-not (Test-Path $Excel)) { throw "Excel file not found: $Excel" }
    Write-Host "Excel: $Excel"

    # Python environment (created once)
    $python = ".venv\Scripts\python.exe"
    if (-not (Test-Path $python)) {
        Write-Host "Creating .venv and installing requirements ..."
        Invoke-Checked py -m venv .venv
        Invoke-Checked $python -m pip install -q -r requirements.txt
    }

    # Build
    Invoke-Checked $python -W ignore cal2html.py $Excel | Select-Object -Last 1
    Invoke-Checked $python -W ignore extract_xlsx_cal.py $Excel

    # Collect the results in output\ under their website names
    $base = $Excel -replace '\.xlsx$', ''
    $outputs = [ordered]@{
        'psue_cal.html'         = 'psue_cal.html'
        "${base}_all.ics"       = 'psue_cal.ics'
        "${base}_training.ics"  = 'psue_cal_training.ics'
        "${base}_sonst.ics"     = 'psue_cal_sonst.ics'
    }
    New-Item -ItemType Directory -Force output | Out-Null
    foreach ($src in $outputs.Keys) {
        Move-Item $src (Join-Path output $outputs[$src]) -Force
    }
    Write-Host "Output:"
    Get-ChildItem output | ForEach-Object { Write-Host "  output\$($_.Name)" }

    if ($Website) {
        if (-not (Test-Path $Website)) { throw "Website calendar folder not found: $Website" }
        Copy-Item output\* $Website -Force
        Write-Host "Copied to $Website. Review with 'git diff' there, then commit and run the publish action."
    }
}
finally {
    Pop-Location
}
