# OnBase Connection Updated to SqlClient

## Changes Made

The OnBase database connection has been updated from **ODBC** to **System.Data.SqlClient** for better performance and native SQL Server support.

## What Changed

### 1. Connection String (Line 289)
**Before (ODBC):**
```powershell
$onbaseConnString = "DRIVER={SQL Server};SERVER=$onbaseServer;DATABASE=$onbaseDatabase;User Id=onbase_db_readonly;Password=p0exV3XanGknDfFnBvMe;"
```

**After (SqlClient):**
```powershell
$onbaseConnString = "Server=$onbaseServer;Database=$onbaseDatabase;User Id=onbase_db_readonly;Password=p0exV3XanGknDfFnBvMe;"
```

### 2. Connection Object (Line 299)
**Before (ODBC):**
```powershell
$onbaseConnection = New-Object System.Data.Odbc.OdbcConnection($onbaseConnString)
```

**After (SqlClient):**
```powershell
$onbaseConnection = New-Object System.Data.SqlClient.SqlConnection($onbaseConnString)
```

### 3. Test-OnBaseAttachmentExists Function (Lines 96-114)
**Before (ODBC):**
```powershell
function Test-OnBaseAttachmentExists {
    param(
        [System.Data.Odbc.OdbcConnection]$Connection,
        [string]$AttachmentID
    )
    # ...
    $cmd = New-Object System.Data.Odbc.OdbcCommand($sql, $Connection)
    # ...
}
```

**After (SqlClient):**
```powershell
function Test-OnBaseAttachmentExists {
    param(
        [System.Data.SqlClient.SqlConnection]$Connection,
        [string]$AttachmentID
    )
    # ...
    $cmd = New-Object System.Data.SqlClient.SqlCommand($sql, $Connection)
    # ...
}
```

### 4. Get-JVAAttachments Function Parameter (Line 119)
**Before (ODBC):**
```powershell
function Get-JVAAttachments {
    param(
        [System.Data.Odbc.OdbcConnection]$ERPConnection,
        [System.Data.Odbc.OdbcConnection]$OnBaseConnection,
        # ...
    )
}
```

**After (SqlClient):**
```powershell
function Get-JVAAttachments {
    param(
        [System.Data.Odbc.OdbcConnection]$ERPConnection,
        [System.Data.SqlClient.SqlConnection]$OnBaseConnection,
        # ...
    )
}
```

## Benefits of SqlClient

### Performance
- **Native SQL Server driver** - Optimized for SQL Server
- **Faster execution** - Direct communication without ODBC layer
- **Better connection pooling** - More efficient resource management

### Features
- **Better error messages** - More detailed SQL Server-specific errors
- **Advanced features** - Support for SQL Server-specific features
- **Async support** - Better async/await capabilities

### Reliability
- **No ODBC dependency** - One less layer to troubleshoot
- **Built into .NET** - No external drivers needed
- **Better maintained** - Part of .NET Framework/Core

## Connection Comparison

| Aspect | ODBC | SqlClient |
|--------|------|-----------|
| **Driver** | Generic ODBC driver | Native SQL Server driver |
| **Performance** | Good | Excellent |
| **Connection String** | `DRIVER={SQL Server};SERVER=...` | `Server=...;Database=...` |
| **Namespace** | System.Data.Odbc | System.Data.SqlClient |
| **Dependencies** | Requires ODBC driver | Built into .NET |
| **SQL Server Features** | Limited | Full support |

## ERP Connection (Still ODBC)

The ERP Oracle connection remains ODBC because:
- Oracle requires Oracle client drivers
- ODBC is the standard way to connect to Oracle from PowerShell
- No native .NET provider for Oracle in standard PowerShell

```powershell
# ERP Oracle connection (unchanged)
$erpConnection = New-Object System.Data.Odbc.OdbcConnection($erpConnString)
```

## Testing

After this change, test the script to ensure:

1. **OnBase connection works:**
   ```powershell
   # Should connect successfully
   $onbaseConnection = New-Object System.Data.SqlClient.SqlConnection($onbaseConnString)
   $onbaseConnection.Open()
   ```

2. **Duplicate checking works:**
   ```powershell
   # Should query OnBase successfully
   Test-OnBaseAttachmentExists -Connection $onbaseConnection -AttachmentID "12345"
   ```

3. **Full script runs:**
   ```powershell
   .\ERP_Process_JVA_Dip_Files.ps1
   ```

## Troubleshooting

### Connection Errors

**Error:** "A network-related or instance-specific error occurred"
- Check server name is correct
- Verify SQL Server is running
- Check firewall settings

**Error:** "Login failed for user 'onbase_db_readonly'"
- Verify credentials are correct
- Check user has permissions on OnBase database

### Performance Issues

If you experience slow queries:
- Check SQL Server performance
- Review query execution plans
- Consider adding indexes to OnBase tables

## Summary

The OnBase connection now uses **System.Data.SqlClient** instead of ODBC, providing:
- ✅ Better performance
- ✅ Native SQL Server support
- ✅ Simpler connection string
- ✅ No ODBC driver dependency
- ✅ Better error handling

The script is ready to use with the improved SqlClient connection!

