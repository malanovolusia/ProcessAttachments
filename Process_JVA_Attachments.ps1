#Requires -Version 5.1
<#
.SYNOPSIS
    Master script to process JVA attachments - download and create DIP files
    
.PARAMETER Mode
    Operation mode: 'Download', 'CreateDIP', or 'Both' (default)
    
.EXAMPLE
    .\Process_JVA_Attachments.ps1
    Runs both download and DIP creation
    
.EXAMPLE
    .\Process_JVA_Attachments.ps1 -Mode Download
    Only downloads attachments
    
.EXAMPLE
    .\Process_JVA_Attachments.ps1 -Mode CreateDIP
    Only creates DIP files from existing downloads
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [ValidateSet('Download', 'CreateDIP', 'Both')]
    [string]$Mode = 'Both'
)

function WriteLog {
    param ([String]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] $message"
}

WriteLog "========================================="
WriteLog "JVA Attachment Processing - Mode: $Mode"
WriteLog "========================================="

$scriptPath = $PSScriptRoot

try {
    if ($Mode -eq 'Download' -or $Mode -eq 'Both') {
        WriteLog ""
        WriteLog "Step 1: Downloading Attachments"
        WriteLog "========================================="
        
        $downloadScript = Join-Path $scriptPath "Download_JVA_Attachments.ps1"
        
        if (Test-Path $downloadScript) {
            & $downloadScript
            
            if ($LASTEXITCODE -ne 0 -and $null -ne $LASTEXITCODE) {
                WriteLog "ERROR: Download script failed with exit code $LASTEXITCODE"
                exit 1
            }
        } else {
            WriteLog "ERROR: Download script not found: $downloadScript"
            exit 1
        }
    }
    
    if ($Mode -eq 'CreateDIP' -or $Mode -eq 'Both') {
        WriteLog ""
        WriteLog "Step 2: Creating DIP Files"
        WriteLog "========================================="
        
        $dipScript = Join-Path $scriptPath "Create_JVA_DIP_Files.ps1"
        
        if (Test-Path $dipScript) {
            & $dipScript
            
            if ($LASTEXITCODE -ne 0 -and $null -ne $LASTEXITCODE) {
                WriteLog "ERROR: DIP creation script failed with exit code $LASTEXITCODE"
                exit 1
            }
        } else {
            WriteLog "ERROR: DIP creation script not found: $dipScript"
            exit 1
        }
    }
    
    WriteLog ""
    WriteLog "========================================="
    WriteLog "All operations completed successfully!"
    WriteLog "========================================="
    
} catch {
    WriteLog "ERROR: $_"
    WriteLog $_.ScriptStackTrace
    exit 1
}

