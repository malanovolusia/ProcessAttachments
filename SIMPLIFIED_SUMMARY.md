# Simplified JVA Attachment Processor

## What Changed

The script has been simplified to focus **only on attachment processing**, removing all non-essential infrastructure:

### Removed:
- ❌ Module imports (ERP_mod_logging, ERP_mod_database, etc.)
- ❌ CreateGlobalVariables function call
- ❌ Complex error handling wrappers
- ❌ Unnecessary helper functions (Write-DIPHeader, Write-DIPFooter, Write-IndexHeader, Write-IndexFooter)
- ❌ Extra documentation and comments
- ❌ Unused variables and parameters

### Kept:
✅ Core attachment retrieval logic  
✅ Database connections (ERP Oracle + OnBase SQL Server)  
✅ DIP file generation  
✅ Index file generation  
✅ Duplicate checking in OnBase  
✅ SHA-256 hash calculation  
✅ GUID-based unique filenames  

## File Size Comparison

| Version | Lines of Code | Focus |
|---------|---------------|-------|
| Original | 672 lines | Full-featured with modules |
| Simplified | 353 lines | Attachment processing only |
| **Reduction** | **47% smaller** | **Core functionality** |

## The Script Now Does:

1. **Connect to databases** (ERP Oracle + OnBase SQL Server)
2. **Query JVA documents** that have attachments
3. **For each JVA document:**
   - Query attachments from IN_OBJ_ATT_* tables
   - Check if already in OnBase (optional)
   - Save BLOB to disk with unique GUID filename
   - Calculate SHA-256 hash
   - Write DIP entry
   - Write index entry
4. **Close connections** and report statistics

## Usage

```powershell
# Basic usage
.\ERP_Process_JVA_Dip_Files.ps1

# Production
.\ERP_Process_JVA_Dip_Files.ps1 -SID "ERP19PRO" -OutputPath "\\server\share\JVA"

# Without duplicate checking
.\ERP_Process_JVA_Dip_Files.ps1 -CheckDuplicates $false
```

## Key Functions

### 1. Helper Functions (Lines 42-117)
- `Get-RandomHexValue` - Generate GUID for unique filenames
- `Get-FileTypeNum` - Determine OnBase file type from extension
- `Get-SHA256Hash` - Calculate file hash
- `Save-BinaryData` - Save BLOB to disk
- `Format-PadDate` - Format dates for DIP file
- `Test-OnBaseAttachmentExists` - Check for duplicates

### 2. Main Attachment Function (Lines 119-245)
- `Get-JVAAttachments` - Retrieve and process attachments for a JVA document

### 3. Main Processing (Lines 250-353)
- Direct execution flow
- No wrapper functions
- Simple and straightforward

## Database Connection

**Note:** You'll need to update line 262 with the correct Oracle password:

```powershell
$erpConnString = "Driver={Oracle in OraClient11g_home1};Dbq=$SID;Uid=PDI_USER;Pwd=your_password;"
```

Replace `your_password` with the actual PDI_USER password.

## Output Files

### DIP File: `!JVA_attachment_indexes_DIP.txt`
Contains OnBase import entries:
```
BEGIN:
>>Dummy Key: Document #1
>>DocTypeName: FIN - JVA Attachments
>>DocDate: 01/15/2024
Journal Voucher #: JV123456
Advantage Attachment ID: 12345
...
END:
```

### Index File: `!JVA_indexes.txt`
Pipe-delimited tracking file:
```
OBJ_ATT_UNID|ATT_DATE|GUID_FILENAME|FULL_PATH|USER_ID|DESCRIPTION|DOC_ID|DEPT_CD|DOC_ID_OUT|DOC_TYP|DOC_CD|VERS_NO|SHA256
12345|01/10/2024|invoice_[guid].pdf|C:\path\file.pdf|JSMITH|Invoice|JV123456|100|JV123456|JV|JVA|1|ABC123...
```

### Attachment Files
Individual files with GUID-based names:
```
invoice_[a1b2c3d4e5f6789012345678901234].pdf
receipt_[b2c3d4e5f6789012345678901234a1].jpg
```

## What's Different from Reference VBScript

| Aspect | VBScript (PO/DO/MA) | PowerShell (JVA) Simplified |
|--------|---------------------|----------------------------|
| **Lines** | ~2,143 | 353 |
| **Modules** | 14+ VBS modules | None (standalone) |
| **Doc Types** | PO, DO, MA, RQ | JVA only |
| **XML Processing** | Yes (BIRT) | No |
| **Printing** | Yes | No |
| **Attachments** | Yes | Yes (core focus) |
| **Complexity** | High | Low |

## Advantages of Simplified Version

1. **Easy to understand** - No module dependencies
2. **Easy to modify** - All code in one file
3. **Easy to debug** - Simple execution flow
4. **Portable** - Runs anywhere with PowerShell 5.1+
5. **Focused** - Does one thing well

## Next Steps

1. Update the Oracle password on line 262
2. Test with `-SID "ERP19TEST"` first
3. Verify DIP file format
4. Import to OnBase
5. Schedule if needed

## Troubleshooting

### Connection Issues
- Verify Oracle client is installed
- Check SID is correct
- Verify credentials

### No Attachments Found
- Check JVA documents have OBJ_ATT_PG_UNID populated
- Verify OBJ_ATT_PG_TOT > 0
- Check IN_OBJ_ATT_STOR has BLOB data

### Permission Errors
- Ensure write access to output path
- Check database permissions

## Summary

This simplified version removes all the non-attachment-related complexity while maintaining the core functionality needed to extract JVA attachments and prepare them for OnBase import. It's a clean, focused, and maintainable solution.

