# JVA Attachment Processing Script

## Overview

This PowerShell script processes Journal Voucher (JVA) document attachments from the ERP system and creates DIP (Document Import Package) files for OnBase import. It is based on the PO/DO/MA attachment processing logic from `ERP_ncp_purchasing_post_processor.wsf`.

## Features

- **Retrieves JVA attachments** from the ERP Oracle database
- **Checks for duplicates** in OnBase to avoid re-importing existing attachments
- **Generates unique filenames** using GUIDs to prevent conflicts
- **Creates DIP files** for OnBase import with proper metadata
- **Creates index files** for tracking and reference
- **Calculates SHA-256 hashes** for file integrity verification
- **Handles deleted attachments** appropriately
- **Comprehensive logging** for troubleshooting

## Requirements

- PowerShell 5.1 or higher
- Access to ERP Oracle database (PDI_USER)
- Access to OnBase SQL Server database (read-only)
- Required PowerShell modules from `\\erp311script\Library\PSM1\`:
  - ERP_mod_logging.psm1
  - ERP_mod_database.psm1
  - ERP_mod_file.psm1
  - And other ERP modules

## Usage

### Basic Usage

```powershell
.\ERP_Process_JVA_Dip_Files.ps1
```

This will use default parameters:
- SID: ERP19TEST
- OutputPath: .\Output\JVA\attachments
- CheckDuplicates: $true

### With Parameters

```powershell
# Production environment
.\ERP_Process_JVA_Dip_Files.ps1 -SID "ERP19PRO" -OutputPath "\\server\share\JVA\attachments"

# Test environment without duplicate checking
.\ERP_Process_JVA_Dip_Files.ps1 -SID "ERP19TEST" -OutputPath "C:\Temp\JVA" -CheckDuplicates $false
```

### Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| SID | string | No | ERP19TEST | Oracle SID (e.g., ERP19PRO, ERP19TEST) |
| OutputPath | string | No | .\Output\JVA\attachments | Path where attachments and DIP files will be saved |
| CheckDuplicates | bool | No | $true | Whether to check if attachments already exist in OnBase |

## Output Files

The script creates the following files in the output directory:

### 1. Attachment Files
Individual attachment files with GUID-based unique names:
- Format: `originalname_[guid].ext`
- Example: `invoice_[a1b2c3d4e5f6].pdf`

### 2. DIP File
`!JVA_attachment_indexes_DIP.txt` - Contains metadata for OnBase import:
```
BEGIN:
>>Dummy Key: Document #1
>>DocTypeName: FIN - JVA Attachments
>>DocDate: 01/15/2024
Journal Voucher #: JV123456
Advantage Attachment ID: 12345
Attachment Date: 01/10/2024
Filename: invoice.pdf
>>Dummy Key: hidden keywords begin here
Doc ID: JV123456
Version #: 1
Department #: 100
...
```

### 3. Index File
`!JVA_indexes.txt` - Pipe-delimited index for tracking:
```
OBJ_ATT_UNID|ATT_DATE|GUID_FILENAME|FULL_PATH|USER_ID|DESCRIPTION|DOC_ID|DEPT_CD|DOC_ID_OUT|DOC_TYP|DOC_CD|VERS_NO|SHA256
12345|01/10/2024|invoice_[guid].pdf|C:\path\to\file.pdf|JSMITH|Invoice for services|JV123456|100|JV123456|JV|JVA|1|ABC123...
```

## Database Tables Used

### ERP Oracle Database (O_FINPROD schema)

1. **JV_DOC_HDR** - Journal Voucher document headers
   - DOC_CD, DOC_DEPT_CD, DOC_ID, DOC_VERS_NO
   - OBJ_ATT_PG_UNID (attachment page unique ID)
   - OBJ_ATT_PG_TOT (total attachments)

2. **IN_OBJ_ATT_CTLG** - Attachment catalog (metadata)
   - OBJ_ATT_UNID, OBJ_ATT_NM, OBJ_ATT_DSCR
   - OBJ_ATT_DT, OBJ_ATT_USER_ID
   - OBJ_ATT_SEQ_NO, OBJ_ATT_ST, OBJ_ATT_TYP

3. **IN_OBJ_ATT_STOR** - Attachment storage (BLOB data)
   - OBJ_ATT_DATA (binary attachment data)

4. **IN_OBJ_ATT_DOC_REF** - Document-attachment references
   - Links attachments to documents via DOC_TYP, DOC_CD, DOC_ID

### OnBase SQL Server Database

1. **hsi.keyitem481** - OnBase key items for duplicate checking
2. **hsi.itemdata** - OnBase item data

## How It Works

1. **Initialize**: Sets up logging, database connections, and output files
2. **Query JVA Documents**: Retrieves all JVA documents that have attachments
3. **For Each JVA Document**:
   - Query attachments from IN_OBJ_ATT_* tables
   - Check if attachment already exists in OnBase (if enabled)
   - Save attachment BLOB to disk with unique filename
   - Calculate SHA-256 hash
   - Write DIP entry with metadata
   - Write index entry for tracking
4. **Finalize**: Close files, connections, and report statistics

## Key Differences from PO Processing

| Aspect | PO Processing | JVA Processing |
|--------|---------------|----------------|
| Document Type | PO, DO, MA | JVA |
| DOC_TYP | 'PO', 'MA', 'RQ' | 'JV' |
| Header Table | PO_DOC_HDR, MA_DOC_HDR | JV_DOC_HDR |
| OnBase Doc Type | PUR - Advantage Attachments | FIN - JVA Attachments |
| Related Docs | Retrieves RQS for POs | No related documents |

## Error Handling

- Database connection failures are logged and script exits with code 1
- Individual attachment errors are logged but processing continues
- Deleted attachments (NULL BLOB) are logged but not saved
- Missing file extensions are handled gracefully

## Logging

The script provides detailed logging including:
- Database connection status
- SQL queries executed
- Each document and attachment processed
- Duplicate detection results
- Error messages with stack traces
- Final statistics

## Troubleshooting

### Common Issues

1. **Database Connection Failed**
   - Verify SID parameter is correct
   - Check network connectivity to database servers
   - Verify credentials in connection string

2. **No Attachments Found**
   - Verify JVA documents have OBJ_ATT_PG_UNID populated
   - Check OBJ_ATT_PG_TOT > 0
   - Verify attachments exist in IN_OBJ_ATT_STOR

3. **Permission Denied on Output Path**
   - Ensure write permissions to output directory
   - Check if files are locked by another process

4. **Module Not Found**
   - Verify access to `\\erp311script\Library\PSM1\`
   - Check network connectivity to script library

## Customization

### Filtering JVA Documents

Modify the `$jvaQuery` in the `Process-JVADocuments` function to filter documents:

```powershell
# Only process documents from specific department
$jvaQuery = @"
SELECT DOC_CD, DOC_DEPT_CD, DOC_ID, DOC_VERS_NO, OBJ_ATT_PG_UNID, DOC_DSCR, DOC_REC_DT
FROM O_FINPROD.JV_DOC_HDR
WHERE OBJ_ATT_PG_UNID IS NOT NULL
  AND OBJ_ATT_PG_TOT > 0
  AND DOC_DEPT_CD = '100'
ORDER BY DOC_ID, DOC_VERS_NO
"@
```

### Changing OnBase Document Type

Modify line 398 in the `Get-JVAAttachments` function:

```powershell
$docTypeName = "FIN - JVA Attachments - Custom"
```

## Version History

- **v1.0** - Initial version based on ERP_ncp_purchasing_post_processor.wsf
  - Support for JVA document attachments
  - DIP and index file generation
  - Duplicate checking in OnBase
  - SHA-256 hash calculation

## Support

For issues or questions, contact the ERP development team.

## References

- Original VBScript: `References\ERP_ncp_purchasing_post_processor.wsf`
- ERP Module Library: `\\erp311script\Library\PSM1\`

