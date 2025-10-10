# Reusable Download-BlobChunked Function

## Quick Copy-Paste

To use the fast chunked download in any script, copy this function:

```powershell
<#
.SYNOPSIS
    Downloads a single BLOB from Oracle using chunked streaming.
#>
function Download-BlobChunked {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Data.Odbc.OdbcConnection]$Connection,
        
        [Parameter(Mandatory=$true)]
        [long]$UNID,
        
        [Parameter(Mandatory=$true)]
        [string]$OutputFile,
        
        [Parameter(Mandatory=$false)]
        [long]$ExpectedSize = 0,
        
        [Parameter(Mandatory=$false)]
        [int]$ChunkSize = 16000
    )
    
    $result = @{
        Success = $false
        BytesDownloaded = 0
        ElapsedSeconds = 0
        SpeedKBps = 0
        Error = $null
    }
    
    $fileStream = $null
    $reader = $null
    $startTime = Get-Date
    
    try {
        $blobQuery = @"
SELECT b.OBJ_ATT_DATA
FROM O_FINPROD.IN_OBJ_ATT_STOR b
WHERE b.OBJ_ATT_UNID = $UNID
"@
        
        $blobCmd = New-Object System.Data.Odbc.OdbcCommand($blobQuery, $Connection)
        $blobCmd.CommandTimeout = 300
        $reader = $blobCmd.ExecuteReader([System.Data.CommandBehavior]::SequentialAccess)
        
        if ($reader.Read()) {
            if (-not $reader.IsDBNull(0)) {
                $fileStream = [System.IO.File]::Create($OutputFile)
                $buffer = New-Object byte[] $ChunkSize
                $fieldOffset = 0
                
                while ($true) {
                    $bytesRead = $reader.GetBytes(0, $fieldOffset, $buffer, 0, $ChunkSize)
                    if ($bytesRead -eq 0) { break }
                    
                    $fileStream.Write($buffer, 0, $bytesRead)
                    $result.BytesDownloaded += $bytesRead
                    $fieldOffset += $bytesRead
                }
                
                $result.Success = $true
            } else {
                $result.Error = "BLOB data is NULL"
            }
        } else {
            $result.Error = "No record found for UNID $UNID"
        }
    } catch {
        $result.Error = $_.Exception.Message
    } finally {
        if ($reader -ne $null) { $reader.Close() }
        if ($fileStream -ne $null) { 
            $fileStream.Close()
            $fileStream.Dispose()
        }
        
        $elapsed = (Get-Date) - $startTime
        $result.ElapsedSeconds = $elapsed.TotalSeconds
        $result.SpeedKBps = if ($elapsed.TotalSeconds -gt 0) { 
            [math]::Round(($result.BytesDownloaded / 1024) / $elapsed.TotalSeconds, 2) 
        } else { 0 }
    }
    
    return $result
}
```

---

## Usage Examples

### Example 1: Basic Download

```powershell
# Connect to database
$erpConnString = GetOracleConnectString "ERP19PRO" "PDI_USER" $False
$erpConnection = New-Object System.Data.Odbc.OdbcConnection($erpConnString)
$erpConnection.Open()

# Download a single attachment
$result = Download-BlobChunked `
    -Connection $erpConnection `
    -UNID 1685781 `
    -OutputFile "C:\Attachments\myfile.pdf"

if ($result.Success) {
    Write-Host "Downloaded $($result.BytesDownloaded) bytes"
    Write-Host "Speed: $($result.SpeedKBps) KB/s"
} else {
    Write-Host "Failed: $($result.Error)"
}

$erpConnection.Close()
```

### Example 2: Download Multiple Attachments

```powershell
$erpConnection = New-Object System.Data.Odbc.OdbcConnection($erpConnString)
$erpConnection.Open()

# Get list of UNIDs
$query = "SELECT OBJ_ATT_UNID, OBJ_ATT_NM FROM O_FINPROD.IN_OBJ_ATT_CTLG WHERE ..."
$cmd = New-Object System.Data.Odbc.OdbcCommand($query, $erpConnection)
$reader = $cmd.ExecuteReader()

$successCount = 0
$failCount = 0

while ($reader.Read()) {
    $unid = $reader["OBJ_ATT_UNID"]
    $filename = $reader["OBJ_ATT_NM"]
    
    $result = Download-BlobChunked `
        -Connection $erpConnection `
        -UNID $unid `
        -OutputFile "C:\Attachments\$filename"
    
    if ($result.Success) {
        $successCount++
        Write-Host "✓ $filename - $($result.SpeedKBps) KB/s"
    } else {
        $failCount++
        Write-Host "✗ $filename - $($result.Error)"
    }
}

$reader.Close()
Write-Host "Success: $successCount, Failed: $failCount"

$erpConnection.Close()
```

### Example 3: With Progress Tracking

```powershell
$erpConnection = New-Object System.Data.Odbc.OdbcConnection($erpConnString)
$erpConnection.Open()

$attachments = @(
    @{UNID=1685781; Name="file1.pdf"},
    @{UNID=1685782; Name="file2.pdf"},
    @{UNID=1685783; Name="file3.pdf"}
)

$totalFiles = $attachments.Count
$currentFile = 0

foreach ($att in $attachments) {
    $currentFile++
    Write-Host "[$currentFile/$totalFiles] Downloading $($att.Name)..."
    
    $result = Download-BlobChunked `
        -Connection $erpConnection `
        -UNID $att.UNID `
        -OutputFile "C:\Attachments\$($att.Name)"
    
    if ($result.Success) {
        $sizeKB = [math]::Round($result.BytesDownloaded / 1024, 2)
        Write-Host "  ✓ $sizeKB KB in $([math]::Round($result.ElapsedSeconds, 2))s"
    } else {
        Write-Host "  ✗ Error: $($result.Error)"
    }
}

$erpConnection.Close()
```

### Example 4: Custom Table/Schema

```powershell
# If your BLOB is in a different table, modify the function:

function Download-CustomBlob {
    param(
        [System.Data.Odbc.OdbcConnection]$Connection,
        [string]$TableName,
        [string]$BlobColumn,
        [string]$WhereClause,
        [string]$OutputFile
    )
    
    $result = @{Success = $false; BytesDownloaded = 0; Error = $null}
    $fileStream = $null
    $reader = $null
    
    try {
        $query = "SELECT $BlobColumn FROM $TableName WHERE $WhereClause"
        
        $cmd = New-Object System.Data.Odbc.OdbcCommand($query, $Connection)
        $cmd.CommandTimeout = 300
        $reader = $cmd.ExecuteReader([System.Data.CommandBehavior]::SequentialAccess)
        
        if ($reader.Read() -and -not $reader.IsDBNull(0)) {
            $fileStream = [System.IO.File]::Create($OutputFile)
            $buffer = New-Object byte[] 16000
            $offset = 0
            
            while ($true) {
                $bytesRead = $reader.GetBytes(0, $offset, $buffer, 0, 16000)
                if ($bytesRead -eq 0) { break }
                
                $fileStream.Write($buffer, 0, $bytesRead)
                $result.BytesDownloaded += $bytesRead
                $offset += $bytesRead
            }
            
            $result.Success = $true
        }
    } catch {
        $result.Error = $_.Exception.Message
    } finally {
        if ($reader) { $reader.Close() }
        if ($fileStream) { $fileStream.Close(); $fileStream.Dispose() }
    }
    
    return $result
}

# Usage:
$result = Download-CustomBlob `
    -Connection $conn `
    -TableName "MY_SCHEMA.MY_TABLE" `
    -BlobColumn "BLOB_DATA" `
    -WhereClause "ID = 12345" `
    -OutputFile "C:\output.dat"
```

---

## Integration with Existing Scripts

### Replace Old Method

**Find this pattern:**
```powershell
$cmd = New-Object System.Data.Odbc.OdbcCommand($query, $connection)
$reader = $cmd.ExecuteReader()
if ($reader.Read()) {
    $blobData = $reader["BLOB_COLUMN"]
    [System.IO.File]::WriteAllBytes($outputFile, $blobData)
}
$reader.Close()
```

**Replace with:**
```powershell
$result = Download-BlobChunked `
    -Connection $connection `
    -UNID $unid `
    -OutputFile $outputFile

if ($result.Success) {
    Write-Host "Downloaded successfully"
}
```

---

## Return Value Reference

The function returns a hashtable with these properties:

| Property | Type | Description |
|----------|------|-------------|
| `Success` | Boolean | `$true` if download succeeded, `$false` otherwise |
| `BytesDownloaded` | Long | Total bytes downloaded |
| `ElapsedSeconds` | Double | Time taken in seconds |
| `SpeedKBps` | Double | Download speed in KB/s |
| `Error` | String | Error message if failed, `$null` if succeeded |

### Example Usage of Return Values

```powershell
$result = Download-BlobChunked -Connection $conn -UNID 123 -OutputFile "file.pdf"

# Check success
if ($result.Success) {
    # Get file size
    $sizeMB = [math]::Round($result.BytesDownloaded / 1024 / 1024, 2)
    Write-Host "Downloaded $sizeMB MB"
    
    # Get speed
    Write-Host "Speed: $($result.SpeedKBps) KB/s"
    
    # Get time
    Write-Host "Time: $($result.ElapsedSeconds) seconds"
} else {
    # Handle error
    Write-Host "Error: $($result.Error)"
    
    # Log to file
    Add-Content "errors.log" "UNID 123: $($result.Error)"
}
```

---

## Performance Tuning

### Adjust Chunk Size

```powershell
# Slower network - use smaller chunks
$result = Download-BlobChunked -Connection $conn -UNID 123 -OutputFile "file.pdf" -ChunkSize 8000

# Faster network - use larger chunks (max 32000)
$result = Download-BlobChunked -Connection $conn -UNID 123 -OutputFile "file.pdf" -ChunkSize 32000

# Default (recommended)
$result = Download-BlobChunked -Connection $conn -UNID 123 -OutputFile "file.pdf" -ChunkSize 16000
```

### Parallel Downloads

```powershell
# Download multiple files in parallel using PowerShell jobs
$jobs = @()

foreach ($unid in $unidList) {
    $jobs += Start-Job -ScriptBlock {
        param($connString, $unid, $outputPath)
        
        # Each job needs its own connection
        $conn = New-Object System.Data.Odbc.OdbcConnection($connString)
        $conn.Open()
        
        # Call the function (must be defined in job scope)
        $result = Download-BlobChunked -Connection $conn -UNID $unid -OutputFile "$outputPath\$unid.dat"
        
        $conn.Close()
        return $result
    } -ArgumentList $erpConnString, $unid, $outputPath
}

# Wait for all jobs
$jobs | Wait-Job | Receive-Job
```

---

## Troubleshooting

### "No record found for UNID"
- UNID doesn't exist in `IN_OBJ_ATT_STOR` table
- Check the UNID value is correct
- Verify table name and schema

### "BLOB data is NULL"
- The `OBJ_ATT_DATA` column is NULL
- File was never uploaded or was deleted
- Check source data

### Slow downloads
- Increase chunk size to 32000
- Check network connection to database
- Verify database performance

### Out of memory
- Should not happen with chunked method
- Verify you're using `Download-BlobChunked` not old method
- Check chunk size is reasonable (16000-32000)

---

## Summary

✅ **Copy the function** to any script  
✅ **Call with UNID** and output file path  
✅ **Check result.Success** for status  
✅ **Use result properties** for metrics  
✅ **Adjust ChunkSize** for performance  

The function is self-contained and has no external dependencies!

