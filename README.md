# purview_scripts

Small PowerShell helpers for Microsoft Purview / Exchange Online text extraction and (optionally) data classification. The scripts connect to Exchange Online, submit a local file to `Test-TextExtraction`, and optionally run `Test-DataClassification` against extracted streams.

## What's included

- `textExctraction.ps1` — Connects to Exchange Online and runs `Test-TextExtraction` against a local file.
- `testDataclassification.ps1` — Runs `Test-TextExtraction` and, if requested, connects to the Purview compliance session (IPPS) and runs `Test-DataClassification` against extracted text streams.
- `export-sit-rulepack.ps1` — Lists available Purview SIT rulepacks and exports the selected rulepack as XML.

## Requirements

- PowerShell 7+ (`pwsh`) on macOS, PowerShell 5.1+ on Windows.
- Exchange Online PowerShell module (`ExchangeOnlineManagement`).
- Permissions to run `Test-TextExtraction` and, if using data classification, `Test-DataClassification` and `Get-DlpSensitiveInformationType` in the compliance session.
- A valid Exchange Online user principal name (UPN) for authentication.

### Install the Exchange Online module

Windows (PowerShell 5.1+ or 7+):

```powershell
Install-PSResource -Name ExchangeOnlineManagement -Scope CurrentUser
```

macOS (PowerShell 7+):

```powershell
pwsh -Command "Install-PSResource -Name ExchangeOnlineManagement -Scope CurrentUser"
```

If prompted, accept the PSGallery repository trust prompt.

## Usage

### Text extraction only

Windows:

```powershell
pwsh ./textExctraction.ps1 -UserPrincipalName "admin@contoso.com" -WinFile "C:\Temp\document.pdf"
```

macOS:

```powershell
pwsh ./textExctraction.ps1 -UserPrincipalName "admin@contoso.com" -MacFile "$HOME/temp/document.pdf"
```

### Extraction + data classification

```powershell
pwsh ./testDataclassification.ps1 -UserPrincipalName "admin@contoso.com" -WinFile "C:\Temp\document.msg" -DataClassification
```

Scope to specific Sensitive Information Types (display names or IDs):

```powershell
pwsh ./testDataclassification.ps1 -UserPrincipalName "admin@contoso.com" -WinFile "C:\Temp\document.pdf" -DataClassification -SensitiveInformationTypes "U.S. Social Security Number, Credit Card Number"
```

Run against all Sensitive Information Types:

```powershell
pwsh ./testDataclassification.ps1 -UserPrincipalName "admin@contoso.com" -WinFile "C:\Temp\document.pdf" -DataClassification -AllSensitiveInformationTypes
```

### Export a SIT rulepack (XML)

```powershell
pwsh ./export-sit-rulepack.ps1 -UserPrincipalName "admin@contoso.com"
```

Export to a specific directory:

```powershell
pwsh ./export-sit-rulepack.ps1 -UserPrincipalName "admin@contoso.com" -OutputDirectory "$HOME/Downloads"
```

The selection list and export filename prefer the rulepack `LocalizedName` when available.

### File path selection

Both scripts accept Windows and macOS paths and choose the one that matches the current OS.
If you provide both, the OS-appropriate path is used.

`testDataclassification.ps1` also accepts `-FilePath` as a direct override.

## Output

- `textExctraction.ps1` returns `ExtractedResults` as JSON.
- `testDataclassification.ps1` returns a JSON object with:
  - `SourceFile`
  - `Streams` (derived stream metadata)
  - `Extraction` (raw `ExtractedResults`)
  - `DataClassification` (per-stream results)

## Notes

- `textExctraction.ps1` is intentionally named with the original filename spelling.
- macOS paths must be full paths (do not use `~/...` in parameters).
- For email files (`.msg`, `.eml`), stream naming heuristics treat the main body as `Body` and attachments as `Attachment`.

## Troubleshooting

- **Authentication failures**: confirm your UPN has access to Exchange Online and that the `ExchangeOnlineManagement` module is installed.
- **`Test-TextExtraction` not found**: verify the Exchange Online module version and that your session is connected.
- **`Test-DataClassification` errors**: ensure you can connect to the IPPS session and have compliance permissions.

## License

Add your preferred license here.
