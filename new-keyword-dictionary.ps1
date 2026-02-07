<#
.SYNOPSIS
Creates a new Purview keyword dictionary from a CSV or TXT file.

.DESCRIPTION
Connects to the Purview compliance session (IPPS), reads a local file as bytes,
and creates a new keyword dictionary. The input file, dictionary name, description,
and match behavior can be provided as parameters or prompted interactively.

.PARAMETER UserPrincipalName
The Exchange Online sign-in identity (UPN), e.g. admin@contoso.com.

.PARAMETER FilePath
Path to the CSV or TXT file containing keywords.

.PARAMETER Name
Dictionary name.

.PARAMETER Description
Dictionary description.

.PARAMETER MatchType
Optional. Desired match behavior (e.g. Exact, Word, String). The script will
apply this value only if a compatible parameter exists on New-DlpKeywordDictionary.

.EXAMPLE
pwsh ./new-keyword-dictionary.ps1 -UserPrincipalName "admin@contoso.com" -FilePath "$HOME/Downloads/LOCALITY.csv" -Name "dic_locality_australia" -Description "All localities in Australia"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory, HelpMessage = "UPN used to authenticate to Purview (e.g. admin@contoso.com).")]
    [ValidateNotNullOrEmpty()]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $false, HelpMessage = "Path to the CSV/TXT file containing keywords.")]
    [ValidateNotNullOrEmpty()]
    [string]$FilePath,

    [Parameter(Mandatory = $false, HelpMessage = "Keyword dictionary name.")]
    [ValidateNotNullOrEmpty()]
    [string]$Name,

    [Parameter(Mandatory = $false, HelpMessage = "Keyword dictionary description.")]
    [ValidateNotNullOrEmpty()]
    [string]$Description,

    [Parameter(Mandatory = $false, HelpMessage = "Match behavior (e.g. Exact, Word, String).")]
    [ValidateNotNullOrEmpty()]
    [string]$MatchType
)

function Resolve-MatchParameter {
    param(
        [System.Management.Automation.CommandInfo]$Command,
        [string]$MatchType
    )

    if ([string]::IsNullOrWhiteSpace($MatchType)) {
        return $null
    }

    $candidates = @("ExactMatch", "IsExactMatch", "MatchType", "MatchMode", "WordMatch", "MatchBehavior")
    foreach ($name in $candidates) {
        if ($Command.Parameters.ContainsKey($name)) {
            $param = $Command.Parameters[$name]
            $type = $param.ParameterType

            if ($type -eq [switch] -or $type -eq [System.Management.Automation.SwitchParameter]) {
                $truthy = @("exact", "word", "exactmatch", "wordmatch", "true", "yes", "1")
                return @{
                    Name = $name
                    Value = $truthy -contains $MatchType.ToLowerInvariant()
                }
            }

            if ($type -eq [bool] -or $type -eq [System.Boolean]) {
                $truthy = @("exact", "word", "exactmatch", "wordmatch", "true", "yes", "1")
                return @{
                    Name = $name
                    Value = $truthy -contains $MatchType.ToLowerInvariant()
                }
            }

            return @{
                Name = $name
                Value = $MatchType
            }
        }
    }

    return $null
}

try {
    if (-not $FilePath) {
        $FilePath = Read-Host "Enter CSV/TXT file path"
    }
    if (-not $Name) {
        $Name = Read-Host "Enter dictionary name"
    }
    if (-not $Description) {
        $Description = Read-Host "Enter dictionary description"
    }
    if (-not $MatchType) {
        $MatchType = Read-Host "Enter match type (Exact/Word/String or leave blank)"
    }

    if (-not (Test-Path -LiteralPath $FilePath)) {
        throw "File not found: $FilePath"
    }

    Connect-IPPSSession -UserPrincipalName $UserPrincipalName -ShowBanner:$true -ErrorAction Stop

    $fileData = [System.IO.File]::ReadAllBytes($FilePath)
    $cmd = Get-Command New-DlpKeywordDictionary -ErrorAction Stop

    $params = @{
        Name = $Name
        Description = $Description
        FileData = $fileData
        ErrorAction = "Stop"
    }

    $matchParam = Resolve-MatchParameter -Command $cmd -MatchType $MatchType
    if ($matchParam) {
        $params[$matchParam.Name] = $matchParam.Value
    }

    New-DlpKeywordDictionary @params
}
catch {
    Write-Error $_
}
finally {
    if (Get-Command Disconnect-IPPSSession -ErrorAction SilentlyContinue) {
        Disconnect-IPPSSession -Confirm:$false -ErrorAction SilentlyContinue
    }
}
