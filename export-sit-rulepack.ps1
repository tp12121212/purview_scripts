<#
.SYNOPSIS
Lists and exports Microsoft Purview SIT rulepacks (XML) from a tenant.

.DESCRIPTION
Connects to the Purview compliance session (IPPS), lists available rulepacks,
prompts for a selection, and exports the chosen rulepack to an XML file named
after the rulepack.

.PARAMETER UserPrincipalName
The Exchange Online sign-in identity (UPN), e.g. admin@contoso.com.

.PARAMETER OutputDirectory
Optional. Directory to write the XML file. Defaults to the current directory.

.EXAMPLE
pwsh ./export-sit-rulepack.ps1 -UserPrincipalName "admin@contoso.com"

.EXAMPLE
pwsh ./export-sit-rulepack.ps1 -UserPrincipalName "admin@contoso.com" -OutputDirectory "$HOME/Downloads"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory, HelpMessage = "UPN used to authenticate to Purview (e.g. admin@contoso.com).")]
    [ValidateNotNullOrEmpty()]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $false, HelpMessage = "Directory for exported XML.")]
    [ValidateNotNullOrEmpty()]
    [string]$OutputDirectory = (Get-Location).Path
)

function Get-RulepackIdentity {
    param([object]$Rulepack)
    foreach ($prop in @("Identity", "Id", "Name", "DisplayName", "RulePackageId", "RulePackageID", "PackageId", "PackageID")) {
        if ($Rulepack.PSObject.Properties.Name -contains $prop) {
            $value = $Rulepack.$prop
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                return $value
            }
        }
    }
    return $null
}

function Get-RulepackName {
    param([object]$Rulepack)
    foreach ($prop in @("LocalizedName", "Localized Name", "Name", "DisplayName", "Identity", "Id", "RulePackageName")) {
        $property = $Rulepack.PSObject.Properties[$prop]
        if ($property) {
            $value = $property.Value
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                return $value
            }
        }
    }
    return "Rulepack"
}

try {
    Connect-IPPSSession -UserPrincipalName $UserPrincipalName -ShowBanner:$true -ErrorAction Stop

    $rulepacks = @(Get-DlpSensitiveInformationTypeRulePackage -ErrorAction Stop)
    if (-not $rulepacks -or $rulepacks.Count -eq 0) {
        throw "No rulepacks were returned for this tenant."
    }

    Write-Host ""
    Write-Host "Available rulepacks:" -ForegroundColor Cyan
    Write-Host ""

    $nameMap = @()
    for ($i = 0; $i -lt $rulepacks.Count; $i++) {
        $name = Get-RulepackName $rulepacks[$i]
        $nameMap += [pscustomobject]@{
            Index = $i
            Name = $name
        }
        Write-Host ("[{0}] {1}" -f ($i + 1), $name)
    }

    Write-Host ""
    $selection = Read-Host "Select a rulepack by name or index"
    if ($selection -as [int]) {
        $index = [int]$selection - 1
        if ($index -lt 0 -or $index -ge $rulepacks.Count) {
            throw "Selection is out of range."
        }
        $selected = $rulepacks[$index]
    }
    else {
        $matches = $nameMap | Where-Object { $_.Name -eq $selection }
        if ($matches.Count -eq 0) {
            throw "No rulepack name matched '$selection'."
        }
        if ($matches.Count -gt 1) {
            throw "Multiple rulepacks matched '$selection'. Use the index instead."
        }
        $selected = $rulepacks[$matches.Index]
    }
    $selectedName = Get-RulepackName $selected
    $selectedIdentity = Get-RulepackIdentity $selected
    if ([string]::IsNullOrWhiteSpace($selectedIdentity)) {
        $selectedIdentity = $selectedName
    }
    Write-Verbose ("Selected rulepack name: {0}" -f $selectedName)
    Write-Verbose ("Selected rulepack identity: {0}" -f $selectedIdentity)

    if (-not (Test-Path -LiteralPath $OutputDirectory)) {
        throw "OutputDirectory not found: $OutputDirectory"
    }

    $exported = Get-DlpSensitiveInformationTypeRulePackage -Identity $selectedIdentity -ErrorAction Stop
    $exportedName = Get-RulepackName $exported
    Write-Verbose ("Exported object type: {0}" -f $exported.GetType().FullName)
    Write-Verbose ("Exported properties: {0}" -f ($exported.PSObject.Properties.Name -join ", "))

    $nameForFile = if (-not [string]::IsNullOrWhiteSpace($exportedName)) { $exportedName } else { $selectedName }
    $safeName = ($nameForFile -replace '[^\w\-. ]', '_').Trim()
    if ([string]::IsNullOrWhiteSpace($safeName)) {
        $safeName = "Rulepack"
    }

    $outputPath = Join-Path -Path $OutputDirectory -ChildPath ($safeName + ".xml")
    $xmlContent = $null
    $rawBytes = $null

    foreach ($prop in @("RulePackage", "RulePackageXML", "RulePackageXml", "Xml", "XML", "FileData", "Data", "Content", "RulePackageFileData", "FileContent", "Binary")) {
        if ($exported.PSObject.Properties.Name -contains $prop) {
            $value = $exported.$prop
            if ($value -is [byte[]]) {
                $rawBytes = $value
                break
            }
            if ($value -is [xml] -or $value -is [System.Xml.XmlDocument]) {
                $xmlContent = $value.OuterXml
                break
            }
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                $xmlContent = $value
                break
            }
        }
    }

    if (-not $xmlContent -and -not $rawBytes) {
        if ($exported -is [byte[]]) {
            $rawBytes = $exported
        }
        elseif ($exported -is [string]) {
            $xmlContent = $exported
        }
    }

    if (-not $xmlContent -and -not $rawBytes) {
        foreach ($prop in $exported.PSObject.Properties) {
            if ($prop.Value -is [byte[]]) {
                $rawBytes = $prop.Value
                break
            }
            if ($prop.Value -is [xml] -or $prop.Value -is [System.Xml.XmlDocument]) {
                $xmlContent = $prop.Value.OuterXml
                break
            }
            if ($prop.Value -is [string] -and $prop.Value.TrimStart().StartsWith("<")) {
                $xmlContent = $prop.Value
                break
            }
        }
    }

    if (-not $xmlContent -and -not $rawBytes) {
        if (Get-Command Export-DlpSensitiveInformationTypeRulePackage -ErrorAction SilentlyContinue) {
            Write-Verbose "Falling back to Export-DlpSensitiveInformationTypeRulePackage..."
            $exported = Export-DlpSensitiveInformationTypeRulePackage -Identity $selectedIdentity -ErrorAction Stop
            Write-Verbose ("Exported (fallback) object type: {0}" -f $exported.GetType().FullName)
            Write-Verbose ("Exported (fallback) properties: {0}" -f ($exported.PSObject.Properties.Name -join ", "))

            foreach ($prop in @("RulePackage", "RulePackageXML", "RulePackageXml", "Xml", "XML", "FileData", "Data", "Content", "RulePackageFileData", "FileContent", "Binary")) {
                if ($exported.PSObject.Properties.Name -contains $prop) {
                    $value = $exported.$prop
                    if ($value -is [byte[]]) {
                        $rawBytes = $value
                        break
                    }
                    if ($value -is [xml] -or $value -is [System.Xml.XmlDocument]) {
                        $xmlContent = $value.OuterXml
                        break
                    }
                    if (-not [string]::IsNullOrWhiteSpace($value)) {
                        $xmlContent = $value
                        break
                    }
                }
            }

            if (-not $xmlContent -and -not $rawBytes) {
                if ($exported -is [byte[]]) {
                    $rawBytes = $exported
                }
                elseif ($exported -is [string]) {
                    $xmlContent = $exported
                }
            }
        }
    }

    if (-not $xmlContent -and -not $rawBytes) {
        throw "Could not locate XML content on the exported rulepack object. Try running with -Verbose to inspect properties."
    }

    if ($rawBytes) {
        [System.IO.File]::WriteAllBytes($outputPath, $rawBytes)
    }
    else {
        $xmlContent | Out-File -FilePath $outputPath -Encoding UTF8
    }

    Write-Host ""
    Write-Host ("Exported: {0}" -f $outputPath) -ForegroundColor Green
}
catch {
    Write-Error $_
}
finally {
    if (Get-Command Disconnect-IPPSSession -ErrorAction SilentlyContinue) {
        Disconnect-IPPSSession -Confirm:$false -ErrorAction SilentlyContinue
    }
}
