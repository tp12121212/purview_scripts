[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$UserPrincipalName,
    
    [Parameter(Mandatory = $false)]
    [string]$WinFile,
    
    [Parameter(Mandatory = $false)]
    [string]$MacFile,
    
    [Parameter(Mandatory = $false)]
    [string]$PythonScriptPath = "./keyword_extraction.py",
    
    [Parameter(Mandatory = $false)]
    [string]$PythonExecutable = "python3"
)

Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowBanner:$false -ErrorAction Stop

try {
    # Determine file path based on OS
    if ($IsWindows) {
        $FilePath = if ($WinFile) { $WinFile } else { $MacFile }
    }
    elseif ($IsMacOS) {
        $FilePath = if ($MacFile) { $MacFile } else { $WinFile }
    }
    else {
        throw "Unsupported OS"
    }

    if (-not (Test-Path -LiteralPath $FilePath)) {
        throw "File not found: $FilePath"
    }

    # Extract text
    $fileBytes = [System.IO.File]::ReadAllBytes($FilePath)
    $extractionResult = Test-TextExtraction -FileData $fileBytes

    # Create array of extracted texts with metadata
    $extractedStreams = @()
    foreach ($result in $extractionResult.ExtractedResults) {
        $extractedStreams += @{
            StreamName = $result.StreamName
            StreamId = $result.StreamId
            StreamTextLength = $result.StreamTextLength
            ExtractedStreamText = $result.ExtractedStreamText
        }
    }

    # Prefer local venv interpreter when available and default python3 was used.
    $ResolvedPythonExecutable = $PythonExecutable
    if ($PythonExecutable -eq "python3") {
        if ($IsMacOS -and (Test-Path -LiteralPath "./.venv/bin/python")) {
            $ResolvedPythonExecutable = "./.venv/bin/python"
        }
        elseif ($IsWindows -and (Test-Path -LiteralPath ".\.venv\Scripts\python.exe")) {
            $ResolvedPythonExecutable = ".\.venv\Scripts\python.exe"
        }
    }

    if (-not (Get-Command $ResolvedPythonExecutable -ErrorAction SilentlyContinue)) {
        throw "Python executable not found: $ResolvedPythonExecutable"
    }

    # Convert to JSON and pass to Python
    $jsonInput = $extractedStreams | ConvertTo-Json -Depth 10 -Compress

    # Call Python script with JSON input using specified Python executable
    Write-Host "Using Python: $ResolvedPythonExecutable" -ForegroundColor Cyan
    $pythonOutput = $jsonInput | & $ResolvedPythonExecutable $PythonScriptPath

    # Display results
    Write-Host "`n=== Python Analysis Results ===" -ForegroundColor Green
    $pythonOutput

}
catch {
    Write-Error $_
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}
