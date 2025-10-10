# Oracle BLOB Download Functions

## Overview

Two reusable PowerShell functions for downloading Oracle BLOB attachments using efficient chunked streaming via ODBC.

## Functions

### 1. `Download-OracleAttachments`

High-level function that queries attachment metadata and downloads all matching files.

#### Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `Connection` | OdbcConnection | Yes | - | Open ODBC connection to Oracle |
| `OutputPath` | string | Yes | - | Directory where files will be saved |
| `DocType` | string | Yes | - | Document type filter (e.g., 'JV') |
| `DocCode` | string | Yes | - | Document code filter (e.g., 'JVA') |
| `DocId` | string | Yes | - | Document ID filter (e.g., 'CV12251005') |
| `DocVersionNo` | int | No | 1 | Maximum document version number |
| `ChunkSize` | int | No | 16000 | Chunk size in bytes for streaming |

#### Returns

Hashtable with statistics:
```powershell
@{
    TotalFiles = 5
    SuccessCount = 5
    FailedCount = 0
    TotalBytes = 12345678
}
```

#### Example Usage

```powershell
# Basic usage
$stats = Download-OracleAttachments `
    -Connection $erpConnection `
    -OutputPath "C:\Attachments" `
    -DocType "JV" `
    -DocCode "JVA" `
    -DocId "CV12251005"

Write-Host "Downloaded $($stats.SuccessCount) of $($stats.TotalFiles) files"

# With custom chunk size
$stats = Download-OracleAttachments `
    -Connection $erpConnection `
    -OutputPath "C:\Attachments" `
    -DocType "JV" `
    -DocCode "JVA" `
    -DocId "CV12251005" `
    -DocVersionNo 2 `
    -ChunkSize 32000
```

---

### 2. `Download-BlobChunked`

Low-level function that downloads a single BLOB using chunked streaming.

#### Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `Connection` | OdbcConnection | Yes | - | Open ODBC connection to Oracle |
| `UNID` | long | Yes | - | OBJ_ATT_UNID identifying the attachment |
| `OutputFile` | string | Yes | - | Full path where file will be saved |
| `ExpectedSize` | long | No | 0 | Expected size in bytes (for progress) |
| `ChunkSize` | int | No | 16000 | Chunk size in bytes for streaming |

#### Returns

Hashtable with download result:
```powershell
@{
    Success = $true
    BytesDownloaded = 12345
    ElapsedSeconds = 1.23
    SpeedKBps = 123.45
    Error = $null  # or error message if failed
}
```

#### Example Usage

```powershell
# Download a specific attachment by UNID
$result = Download-BlobChunked `
    -Connection $erpConnection `
    -UNID 1685781 `
    -OutputFile "C:\Attachments\myfile.pdf" `
    -ExpectedSize 4057890 `
    -ChunkSize 16000

if ($result.Success) {
    Write-Host "Downloaded $($result.BytesDownloaded) bytes in $($result.ElapsedSeconds)s"
    Write-Host "Speed: $($result.SpeedKBps) KB/s"
} else {
    Write-Host "Failed: $($result.Error)"
}
```

---

## How It Works

### Chunked Streaming Process

1. **Metadata Query** (Fast)
   - Queries attachment catalog without BLOB data
   - Uses `DBMS_LOB.GETLENGTH()` to get file sizes
   - Filters out null/empty BLOBs

2. **Chunked Download** (Efficient)
   - Opens file stream on disk
   - Queries BLOB column with `SequentialAccess` mode
   - Uses `DataReader.GetBytes()` to read in chunks
   - Writes each chunk directly to disk
   - Shows real-time progress

### Why This Approach?

| Method | Speed | Memory | Limitations |
|--------|-------|--------|-------------|
| **ExecuteScalar()** | ❌ Slow | ❌ High | Loads entire BLOB into memory |
| **DBMS_LOB.SUBSTR()** | ❌ Fails | ✅ Low | Oracle RAW limit: 32767 bytes |
| **GetBytes() Streaming** | ✅ Fast | ✅ Low | ✅ No limits! |

### Technical Details

- **ODBC Sequential Access**: Optimized for large binary data
- **Chunk Size**: 16KB default (safe for all Oracle/ODBC versions)
- **Memory Usage**: Only one chunk in memory at a time
- **Progress Tracking**: Real-time percentage and speed display
- **Error Handling**: Per-file error handling, continues on failure

---

## Performance

### Typical Download Speeds

| File Size | Chunks | Time (approx) | Speed |
|-----------|--------|---------------|-------|
| 100 KB | 7 | 0.5s | 200 KB/s |
| 1 MB | 64 | 3s | 340 KB/s |
| 5 MB | 320 | 15s | 340 KB/s |
| 10 MB | 640 | 30s | 340 KB/s |

*Speeds vary based on network, database load, and disk I/O*

### Optimization Tips

1. **Chunk Size**
   - Increase for faster networks: 32000 bytes
   - Decrease for slow/unstable networks: 8000 bytes
   - Default 16000 is safe for most scenarios

2. **Parallel Downloads**
   - Use PowerShell jobs for multiple files
   - Each job needs its own database connection

3. **Network**
   - Ensure good connection to Oracle server
   - VPN can slow downloads significantly

---

## Integration Examples

### Example 1: Download All JVA Attachments

```powershell
# Connect to database
$erpConnString = GetOracleConnectString "ERP19PRO" "PDI_USER" $False
$erpConnection = New-Object System.Data.Odbc.OdbcConnection($erpConnString)
$erpConnection.Open()

# Download all attachments for a JVA document
$stats = Download-OracleAttachments `
    -Connection $erpConnection `
    -OutputPath "\\server\share\JVA\attachments" `
    -DocType "JV" `
    -DocCode "JVA" `
    -DocId "CV12251005"

Write-Host "Success: $($stats.SuccessCount), Failed: $($stats.FailedCount)"

$erpConnection.Close()
```

### Example 2: Download Multiple Documents

```powershell
$erpConnection = New-Object System.Data.Odbc.OdbcConnection($erpConnString)
$erpConnection.Open()

$documents = @("CV12251005", "CV12251006", "CV12251007")

foreach ($docId in $documents) {
    Write-Host "Processing $docId..."
    
    $stats = Download-OracleAttachments `
        -Connection $erpConnection `
        -OutputPath "C:\Attachments\$docId" `
        -DocType "JV" `
        -DocCode "JVA" `
        -DocId $docId
    
    Write-Host "  Downloaded: $($stats.SuccessCount) files"
}

$erpConnection.Close()
```

### Example 3: Custom Processing with Low-Level Function

```powershell
# Get list of UNIDs from custom query
$query = "SELECT OBJ_ATT_UNID FROM O_FINPROD.IN_OBJ_ATT_CTLG WHERE ..."
$cmd = New-Object System.Data.Odbc.OdbcCommand($query, $erpConnection)
$reader = $cmd.ExecuteReader()

while ($reader.Read()) {
    $unid = $reader["OBJ_ATT_UNID"]
    
    # Download each BLOB individually
    $result = Download-BlobChunked `
        -Connection $erpConnection `
        -UNID $unid `
        -OutputFile "C:\Attachments\$unid.dat"
    
    if (-not $result.Success) {
        Write-Host "Failed to download UNID $unid: $($result.Error)"
    }
}

$reader.Close()
```

---

## Troubleshooting

### "raw variable length too long"
- **Cause**: Chunk size too large for Oracle/ODBC
- **Solution**: Reduce chunk size to 8000 or 4000

### Slow downloads
- **Cause**: Network latency, small chunk size
- **Solution**: Increase chunk size to 32000 (if no errors)

### Out of memory
- **Cause**: Not using chunked streaming
- **Solution**: Ensure using `Download-BlobChunked` function

### Connection timeout
- **Cause**: Large file, slow network
- **Solution**: Increase `CommandTimeout` in function

---

## File Naming Convention

Downloaded files are named: `{UNID}_{SEQ_NO}_{FILENAME}`

Example: `1685781_1_IET_Conversion_-_Deductible_$2,500.00.pdf`

- **UNID**: Unique identifier from database
- **SEQ_NO**: Sequence number
- **FILENAME**: Original filename (sanitized)

Special characters in filenames are replaced with underscores.

---

## Summary

These functions provide a **production-ready, efficient solution** for downloading Oracle BLOBs via ODBC:

✅ **Fast**: Chunked streaming avoids memory bottlenecks  
✅ **Reliable**: Handles large files without errors  
✅ **Reusable**: Easy to integrate into any script  
✅ **Informative**: Progress tracking and statistics  
✅ **Safe**: Per-file error handling, continues on failure  

