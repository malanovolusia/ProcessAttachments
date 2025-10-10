#Requires -Version 5.1
<#
.SYNOPSIS
    Download JVA (Journal Voucher) attachments from ERP database
    
.EXAMPLE
    .\Download_JVA_Attachments.ps1
#>

Import-Module "\\erp311script\Library\PSM1\ERP_mod_logging.psm1"
Import-Module "\\erp311script\Library\PSM1\ERP_mod_wmi.psm1"
Import-Module "\\erp311script\Library\PSM1\ERP_mod_database.psm1"
Import-Module "\\erp311script\Library\PSM1\ERP_mod_file.psm1"
Import-Module "\\erp311script\Library\PSM1\ERP_mod_notify.psm1"
Import-Module "\\erp311script\Library\PSM1\ERP_mod_datetime.psm1"
Import-Module "\\erp311script\Library\PSM1\ERP_mod_util.psm1"
Import-Module "\\erp311script\Library\PSM1\ERP_mod_string.psm1"
Import-Module "\\erp311script\Library\PSM1\ERP_mod_print.psm1"
Import-Module "\\erp311script\Library\PSM1\ERP_mod_hrm.psm1"
Import-Module "\\erp311script\Library\PSM1\ERP_mod_fin.psm1"
Import-Module "\\erp311script\Library\PSM1\ERP_mod_env.psm1"
Import-Module "\\erp311script\Library\PSM1\ERP_mod_exec.psm1"
Import-Module "\\erp311script\Library\PSM1\ERP_mod_interface.psm1"
Import-Module "\\erp311script\Library\PSM1\ERP_mod_reporting.psm1"

$env:ENVIRONMENT = "PROD"

if ($env:ENVIRONMENT -eq "PROD") {
    # Production logging handled by module
} else {
    function WriteLog {
        param ([String]$message)
        Write-Host "Log: $message"
    }
}

function Test-OnBaseAttachmentExists {
    param(
        [System.Data.SqlClient.SqlConnection]$Connection,
        [string]$AttachmentID
    )
    if (-not $CheckDuplicates) { return $false }
    try {
        $sql = "SELECT COUNT(*) AS NUM_FOUND
FROM hsi.keyitem481 k
INNER JOIN hsi.itemdata i ON i.itemnum = k.itemnum
WHERE k.keyvaluebig = $AttachmentID
AND i.itemtypenum IN (132, 547, 548, 549)"
        $cmd = New-Object System.Data.SqlClient.SqlCommand($sql, $Connection)
        $reader = $cmd.ExecuteReader()
        $count = 0
        if ($reader.Read()) { $count = $reader["NUM_FOUND"] }
        $reader.Close()
        return ($count -gt 0)
    } catch {
        WriteLog "Error checking OnBase: $_"
        return $false
    }
}


#region Main Processing

try {
    CreateGlobalVariables $PSCommandPath $PSScriptRoot

    $arr_emailList = @(
        "cbarber@volusia.gov",
        "mconway@volusia.gov",
        "sdesai@volusia.gov",
        "jveresciaka@volusia.gov",
        "mpeterka@volusia.gov",
        "malano@volusia.gov"
    )
    $emailString = $arr_emailList -join ";"
    
    $AttachmentsPath = "\\dlerp311birt\ProcessingCenter\Main\output\JETPDF\JVA\attachments"
    $IndexFilePath = Join-Path $AttachmentsPath "!JVA_attachment_index.txt"
    $DIPFilePath = Join-Path $AttachmentsPath "!JVA_attachment_indexes_DIP.txt"

    $SID = "ERP19PRO"
    
    # Database connection strings
    $erpConnString = GetOracleConnectString $SID "PDI_USER" $False

    $onbaseServer = if ($SID -eq "ERP19PRO") { "obdbprod" } else { "obdbtest" }
    $onbaseDatabase = if ($SID -eq "ERP19PRO") { "OnBase" } else { "OB15TEST" }
    $onbaseConnString = "Server=$onbaseServer;Database=$onbaseDatabase;User Id=onbase_db_readonly;Password=p0exV3XanGknDfFnBvMe;"

    $CheckDuplicates = $true

    # Check if index file exists
    if (-not (Test-Path $IndexFilePath)) {
        WriteLog "ERROR: Index file not found at: $IndexFilePath"
        exit 1
    }

    WriteLog "Starting OnBase validation"
    WriteLog "Index File: $IndexFilePath"
    WriteLog "OnBase Server: $onbaseServer"
    WriteLog "OnBase Database: $onbaseDatabase"

    # Connect to OnBase
    $onbaseConnection = New-Object System.Data.SqlClient.SqlConnection($onbaseConnString)
    $onbaseConnection.Open()
    WriteLog "Connected to OnBase database"

    # Read index file
    $indexLines = Get-Content $IndexFilePath
    if ($indexLines.Length -lt 2) {
        WriteLog "ERROR: Index file is empty or has no data rows"
        $onbaseConnection.Close()
        exit 1
    }

    $headerLine = $indexLines[0]
    $dataLines = $indexLines[1..($indexLines.Length - 1)]

    WriteLog "Found $($dataLines.Length) attachments to validate"

    # Validate each attachment
    $validatedCount = 0
    $foundCount = 0
    $notFoundCount = 0
    $errorCount = 0
    $missingAttachments = @()
    $errorAttachments = @()

    foreach ($line in $dataLines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        $fields = $line -split '\|'
        if ($fields.Length -lt 1) { continue }

        $OBJ_ATT_UNID = $fields[0]
        $DOC_ID = if ($fields.Length -gt 6) { $fields[6] } else { "" }
        $ORIGINAL_FILENAME = if ($fields.Length -gt 18) { $fields[18] } else { "" }
        $SEQ_NO = if ($fields.Length -gt 13) { $fields[13] } else { "" }

        $validatedCount++

        try {
            $existsInOnBase = Test-OnBaseAttachmentExists -Connection $onbaseConnection -AttachmentID $OBJ_ATT_UNID

            if ($existsInOnBase) {
                $foundCount++
                WriteLog "Found in OnBase: $OBJ_ATT_UNID"
            } else {
                $notFoundCount++
                WriteLog "NOT found in OnBase: $OBJ_ATT_UNID"

                # Store missing attachment details
                $missingAttachments += [PSCustomObject]@{
                    OBJ_ATT_UNID = $OBJ_ATT_UNID
                    DOC_ID = $DOC_ID
                    SEQ_NO = $SEQ_NO
                    ORIGINAL_FILENAME = $ORIGINAL_FILENAME
                }
            }
        } catch {
            $errorCount++
            WriteLog "ERROR validating $OBJ_ATT_UNID : $_"

            # Store error attachment details
            $errorAttachments += [PSCustomObject]@{
                OBJ_ATT_UNID = $OBJ_ATT_UNID
                DOC_ID = $DOC_ID
                SEQ_NO = $SEQ_NO
                ORIGINAL_FILENAME = $ORIGINAL_FILENAME
                ERROR = $_.Exception.Message
            }
        }

        if ($validatedCount % 10 -eq 0) {
            WriteLog "Progress: Validated $validatedCount of $($dataLines.Length) attachments..."
        }
    }

    # Close connection
    $onbaseConnection.Close()

    # Summary
    WriteLog ""
    WriteLog "========================================="
    WriteLog "Validation Summary"
    WriteLog "========================================="
    WriteLog "Total attachments validated: $validatedCount"
    WriteLog "Found in OnBase: $foundCount"
    WriteLog "NOT found in OnBase: $notFoundCount"
    WriteLog "Errors: $errorCount"
    WriteLog "========================================="

    if ($notFoundCount -gt 0 -or $errorCount -gt 0) {
        # Build detailed HTML email message
        $strMsg = @"
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; }
        .container { padding: 30px; max-width: 900px; margin: 0 auto; }
        h1 { border-bottom: 2px solid currentColor; padding-bottom: 10px; margin-top: 0; }
        h2 { margin-top: 30px; border-bottom: 1px solid currentColor; padding-bottom: 8px; }
        .summary { border: 1px solid currentColor; padding: 15px; margin: 20px 0; }
        .summary-item { margin: 8px 0; font-size: 16px; }
        .summary-item strong { display: inline-block; width: 250px; }
        .alert { border: 2px solid currentColor; padding: 15px; margin: 20px 0; }
        table { width: 100%; border-collapse: collapse; margin: 20px 0; border: 1px solid currentColor; }
        th { background-color: ButtonFace; padding: 12px; text-align: left; font-weight: 600; border: 1px solid currentColor; }
        td { padding: 10px 12px; border-bottom: 1px solid currentColor; }
        .footer { margin-top: 30px; padding-top: 20px; border-top: 1px solid currentColor; font-size: 14px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>JVA Attachment Validation Results</h1>

        <div class="alert">
            <strong>Issues Found:</strong> $notFoundCount attachment(s) missing, $errorCount error(s) occurred
        </div>

        <div class="summary">
            <div class="summary-item"><strong>Total Attachments Validated:</strong> $validatedCount</div>
            <div class="summary-item"><strong>Found in OnBase:</strong> $foundCount</div>
            <div class="summary-item"><strong>NOT Found in OnBase:</strong> $notFoundCount</div>
            <div class="summary-item"><strong>Validation Errors:</strong> $errorCount</div>
        </div>
"@

        # Add missing attachments table
        if ($notFoundCount -gt 0) {
            $strMsg += @"

        <h2>Missing Attachments ($notFoundCount)</h2>
        <table>
            <thead>
                <tr>
                    <th>Attachment ID</th>
                    <th>Document ID</th>
                    <th>Seq #</th>
                    <th>Filename</th>
                </tr>
            </thead>
            <tbody>
"@
            foreach ($missing in $missingAttachments) {
                $strMsg += @"
                <tr>
                    <td><strong>$($missing.OBJ_ATT_UNID)</strong></td>
                    <td>$($missing.DOC_ID)</td>
                    <td>$($missing.SEQ_NO)</td>
                    <td>$($missing.ORIGINAL_FILENAME)</td>
                </tr>
"@
            }
            $strMsg += @"
            </tbody>
        </table>
"@
        }

        # Add error attachments table
        if ($errorCount -gt 0) {
            $strMsg += @"

        <h2>Validation Errors ($errorCount)</h2>
        <table>
            <thead>
                <tr>
                    <th>Attachment ID</th>
                    <th>Document ID</th>
                    <th>Seq #</th>
                    <th>Filename</th>
                    <th>Error</th>
                </tr>
            </thead>
            <tbody>
"@
            foreach ($error in $errorAttachments) {
                $strMsg += @"
                <tr>
                    <td><strong>$($error.OBJ_ATT_UNID)</strong></td>
                    <td>$($error.DOC_ID)</td>
                    <td>$($error.SEQ_NO)</td>
                    <td>$($error.ORIGINAL_FILENAME)</td>
                    <td>$($error.ERROR)</td>
                </tr>
"@
            }
            $strMsg += @"
            </tbody>
        </table>
"@
        }

        $strMsg += @"

        <div class="footer">
            <p><strong>Log File:</strong> $global:logname</p>
            <p><em>Please review the log file for complete details.</em></p>
            <p style="color: #999; font-size: 12px;">Generated: $(Get-Date -Format 'MM/dd/yyyy HH:mm:ss')</p>
        </div>
    </div>
</body>
</html>
"@

        $strSub = "JVA Attachment Processing Validation - Issues Found ($notFoundCount missing, $errorCount errors)"

        NotifyCustom -Subject $strSub -Message $strMsg -Target $emailString -Attachment1 "" -Attachment2 ""
        WriteLog "WARNING: $notFoundCount attachments were not found in OnBase, $errorCount errors occurred"
    } else {
        # Send success email with HTML
        $strMsg = @"
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; }
        .container { padding: 30px; max-width: 900px; margin: 0 auto; }
        h1 { border-bottom: 2px solid currentColor; padding-bottom: 10px; margin-top: 0; }
        .success { border: 2px solid currentColor; padding: 20px; margin: 20px 0; }
        .success h2 { margin-top: 0; text-align: center; }
        .summary { border: 1px solid currentColor; padding: 20px; margin: 20px 0; }
        .summary-item { margin: 10px 0; font-size: 16px; }
        .summary-item strong { display: inline-block; width: 250px; }
        .footer { margin-top: 30px; padding-top: 20px; border-top: 1px solid currentColor; font-size: 14px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>JVA Attachment Validation Results</h1>

        <div class="success">
            <h2>All Attachments Validated Successfully!</h2>
        </div>

        <div class="summary">
            <div class="summary-item"><strong>Total Attachments Validated:</strong> $validatedCount</div>
            <div class="summary-item"><strong>Found in OnBase:</strong> $foundCount</div>
            <div class="summary-item"><strong>NOT Found in OnBase:</strong> 0</div>
            <div class="summary-item"><strong>Validation Errors:</strong> 0</div>
        </div>

        <p style="font-size: 16px; margin: 20px 0;">
            All JVA attachments have been successfully validated in OnBase. No issues were detected.
        </p>

        <div class="footer">
            <p><strong>Log File:</strong> $global:logname</p>
            <p style="font-size: 12px;">Generated: $(Get-Date -Format 'MM/dd/yyyy HH:mm:ss')</p>
        </div>
    </div>
</body>
</html>
"@

        $strSub = "JVA Attachment Processing Validation - SUCCESS ($validatedCount attachments validated)"

        NotifyCustom -Subject $strSub -Message $strMsg -Target $emailString -Attachment1 "" -Attachment2 ""
        WriteLog "SUCCESS: All attachments validated successfully"
    }

} catch {
    WriteLog "ERROR: $_"
    WriteLog $_.ScriptStackTrace
    exit 1
}

#endregion

