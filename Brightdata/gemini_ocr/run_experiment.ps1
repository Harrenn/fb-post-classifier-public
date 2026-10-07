$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
$python = Join-Path $repoRoot "Brightdata\rapidocr\.venv\Scripts\python.exe"
$ocrScript = Join-Path $PSScriptRoot "process_brightdata.py"
$classificationScript = Join-Path $PSScriptRoot "classify_compare_to_baseline.py"

if (-not (Test-Path -LiteralPath $python)) {
    throw "Python environment not found: $python"
}

$secret = Read-Host "Gemini API key" -AsSecureString
$pointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secret)
try {
    $env:GEMINI_API_KEY = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer)
    if ([string]::IsNullOrWhiteSpace($env:GEMINI_API_KEY)) {
        throw "No API key was entered."
    }

    Push-Location $repoRoot
    try {
        & $python $ocrScript
        if ($LASTEXITCODE -ne 0) {
            throw "Gemini OCR stage exited with code $LASTEXITCODE. Completed image calls remain cached."
        }

        & $python $classificationScript
        if ($LASTEXITCODE -ne 0) {
            throw "Classification stage exited with code $LASTEXITCODE."
        }
    }
    finally {
        Pop-Location
    }
}
finally {
    Remove-Item Env:GEMINI_API_KEY -ErrorAction SilentlyContinue
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer)
    Remove-Variable secret, pointer -ErrorAction SilentlyContinue
}
