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
    $indexLines = @(Get-Content $IndexFilePath)
    if ($indexLines.Length -lt 2) {
        WriteLog "ERROR: Index file is empty or has no data rows"
        $onbaseConnection.Close()

        $strSub = "JVA Attachment Processing Validation (No new JVA Attachments found to validate) - INFO "
        $strMsg = "No new JVA Attachments found to validate in OnBase.<br><br>"

        NotifyCustom -Subject $strSub -Message $strMsg -Target $emailString -Attachment1 "" -Attachment2 ""

        exit 0
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
        # Build concise email message (limited to 1000 characters)
        $strMsg = "<b>JVA Attachment Validation - Issues Found</b><br><br>"
        $strMsg += "Total Validated: $validatedCount<br>"
        $strMsg += "Found in OnBase: $foundCount<br>"
        $strMsg += "NOT Found: <b>$notFoundCount</b><br>"
        $strMsg += "Errors: <b>$errorCount</b><br><br>"

        if ($notFoundCount -gt 0) {
            $strMsg += "<b>Missing Attachments:</b><br>"
            foreach ($m in $missingAttachments) {
                $strMsg += "- $($m.OBJ_ATT_UNID) (Doc: $($m.DOC_ID))<br>"
            }
            $strMsg += "<br>"
        }

        if ($errorCount -gt 0) {
            $strMsg += "<b>Validation Errors:</b><br>"
            foreach ($e in $errorAttachments) {
                $strMsg += "- $($e.OBJ_ATT_UNID): $($e.ERROR)<br>"
            }
        }

        $strMsg += "<br>See log for details: $global:logname"

        # Ensure message is under 1000 characters
        if ($strMsg.Length -gt 1000) {
            $strMsg = $strMsg.Substring(0, 997) + "..."
        }

        $strSub = "JVA Attachment Processing Validation - Issues Found ($notFoundCount missing, $errorCount errors)"

        NotifyCustom -Subject $strSub -Message $strMsg -Target $emailString -Attachment1 "" -Attachment2 ""
        WriteLog "WARNING: $notFoundCount attachments were not found in OnBase, $errorCount errors occurred"
    } else {
        # Send success email (limited to 1000 characters)
        $strMsg = "<b>JVA Attachment Validation - SUCCESS</b><br><br>"
        $strMsg += "All attachments validated successfully!<br><br>"
        $strMsg += "Total Validated: <b>$validatedCount</b><br>"
        $strMsg += "Found in OnBase: <b>$foundCount</b><br>"
        $strMsg += "NOT Found: 0<br>"
        $strMsg += "Errors: 0<br><br>"
        $strMsg += "All JVA attachments have been successfully validated in OnBase.<br><br>"
        $strMsg += "Log: $global:logname"

        # Ensure message is under 1000 characters
        if ($strMsg.Length -gt 1000) {
            $strMsg = $strMsg.Substring(0, 997) + "..."
        }

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

