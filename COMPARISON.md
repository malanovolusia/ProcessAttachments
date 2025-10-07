# Comparison: VBScript PO Processor vs PowerShell JVA Processor

## Overview

This document compares the original VBScript purchasing post-processor with the new PowerShell JVA attachment processor.

## File Comparison

| Aspect | VBScript (Reference) | PowerShell (New) |
|--------|---------------------|------------------|
| **File** | ERP_ncp_purchasing_post_processor.wsf | ERP_Process_JVA_Dip_Files.ps1 |
| **Language** | VBScript | PowerShell 5.1+ |
| **Lines of Code** | ~2,143 | ~672 |
| **Document Types** | PO, DO, MA, RQ | JVA |
| **Complexity** | High (handles XML, BIRT, printing) | Focused (attachments only) |

## Key Functional Differences

### 1. Document Types Processed

**VBScript:**
- PO (Purchase Orders)
- DO (Delivery Orders)
- MA (Master Agreements)
- RQ (Requisitions - related to POs)

**PowerShell:**
- JVA (Journal Vouchers)

### 2. Database Tables

**VBScript:**
```sql
-- Header tables
PO_DOC_HDR
MA_DOC_HDR
RQ_DOC_HDR

-- Document type in attachment query
DOC_TYP = 'PO', 'MA', 'RQ'
```

**PowerShell:**
```sql
-- Header table
JV_DOC_HDR

-- Document type in attachment query
DOC_TYP = 'JV'
```

### 3. OnBase Document Type Names

**VBScript:**
- "PUR - Advantage Attachments"
- "PUR - Advantage Attachments - EL" (Amendment 10)
- "PUR - Advantage Attachments - PA" (Amendment 10)
- "PUR - Advantage Attachments - TC" (Amendment 10)
- "PUR - Purchase Orders"
- "PUR - Delivery Orders"
- "PUR - Master Agreements"

**PowerShell:**
- "FIN - JVA Attachments"

### 4. Processing Scope

**VBScript:**
- Processes XML files from staging directories
- Modifies XML for BIRT report generation
- Handles JSON print instruction files
- Manages receiving copies (RC)
- Handles print queues
- Calculates row heights for commodity lines
- Splits commodity lines for overflow
- Archives to 7z files
- Copies to OnBase watched folders

**PowerShell:**
- Focuses solely on attachment retrieval
- No XML/BIRT processing
- No printing functionality
- Direct database query approach
- Simpler workflow

## Code Structure Comparison

### VBScript Structure

```vbscript
' Main script flow
1. Initialize variables and connections
2. Clear staging/output directories
3. Loop through document types (DO, PO, MA)
4. For each XML file in staging:
   a. Parse XML
   b. Get attachments
   c. Get related RQ attachments (for POs)
   d. Modify XML/JSON
   e. Calculate row heights
   f. Generate PDFs via BIRT
5. Archive and cleanup
```

### PowerShell Structure

```powershell
# Main script flow
1. Initialize parameters and connections
2. Query JVA documents with attachments
3. For each JVA document:
   a. Get attachments
   b. Save to disk
   c. Write DIP entries
   d. Write index entries
4. Close files and connections
```

## Attachment Retrieval Logic

### Common Elements (Both Scripts)

Both scripts use the same core attachment tables:

```sql
FROM O_FINPROD.IN_OBJ_ATT_CTLG a,      -- Attachment catalog (metadata)
     O_FINPROD.IN_OBJ_ATT_STOR b,      -- Attachment storage (BLOB)
     O_FINPROD.IN_OBJ_ATT_DOC_REF c    -- Document references
WHERE a.OBJ_ATT_UNID = b.OBJ_ATT_UNID(+)
  AND a.OBJ_ATT_UNID = c.OBJ_ATT_UNID
  AND a.OBJ_ATT_PG_UNID = '[page_unid]'
  AND c.DOC_TYP = '[doc_type]'
  AND c.DOC_CD = '[doc_code]'
  AND c.DOC_ID = '[doc_id]'
  AND c.DOC_VERS_NO <= [version]
```

### VBScript GetAttachments Function

```vbscript
Function GetAttachments(pDOC_TYP, pPO_DOC_ID, pDOC_CD, pDOC_DEPT_CD, _
                       pDOC_ID, pDOC_VERS_NO, pOutPath, pOBJ_ATT_PG_UNID, _
                       pVEND_CUST_CD, pLGL_NM, pDOC_REC_DT_DC_FRMT)
  ' Retrieves attachments for PO, MA, or RQ documents
  ' Writes to separate DIP files based on document type
  ' Handles vendor information
End Function
```

### PowerShell Get-JVAAttachments Function

```powershell
function Get-JVAAttachments {
    param(
        [System.Data.Odbc.OdbcConnection]$ERPConnection,
        [System.Data.Odbc.OdbcConnection]$OnBaseConnection,
        [string]$DOC_CD,
        [string]$DOC_DEPT_CD,
        [string]$DOC_ID,
        [string]$DOC_VERS_NO,
        [string]$OBJ_ATT_PG_UNID,
        [string]$OutPath,
        [System.IO.StreamWriter]$DIPFileStream,
        [System.IO.StreamWriter]$IndexFileStream,
        [hashtable]$DocInfo
    )
    # Retrieves attachments for JVA documents
    # Writes to single DIP file
    # No vendor information needed
}
```

## DIP File Format Comparison

### VBScript DIP Entry (PO Attachment)

```
BEGIN:
>>Dummy Key: Document #1
>>DocTypeName: PUR - Advantage Attachments
>>DocDate: 01/15/2024
Purchase Order #: PO123456
Advantage Attachment ID: 12345
Attachment Date: 01/10/2024
Long Description: Invoice for services
Filename: invoice.pdf
>>Dummy Key: hidden keywords begin here
Doc ID: PO123456
Version #: 1
Department #: 100
Advantage Doc Type: PO
Advantage Doc Code: PO
GUID File Name: invoice_[guid].pdf
Vendor-Customer #: V12345
Vendor Name: ABC Company
Advantage Attachment Primary Group ID: 98765
...
```

### PowerShell DIP Entry (JVA Attachment)

```
BEGIN:
>>Dummy Key: Document #1
>>DocTypeName: FIN - JVA Attachments
>>DocDate: 01/15/2024
Journal Voucher #: JV123456
Advantage Attachment ID: 12345
Attachment Date: 01/10/2024
Long Description: Supporting documentation
Filename: receipt.pdf
>>Dummy Key: hidden keywords begin here
Doc ID: JV123456
Version #: 1
Department #: 100
Advantage Doc Type: JV
Advantage Doc Code: JVA
GUID File Name: receipt_[guid].pdf
Advantage Attachment Primary Group ID: 98765
...
```

### Key Differences in DIP Format

| Field | VBScript (PO) | PowerShell (JVA) |
|-------|---------------|------------------|
| Document Label | "Purchase Order #:" | "Journal Voucher #:" |
| Doc Type | PO, MA, RQ | JV |
| Doc Code | PO, DO, MA, RQS | JVA |
| Vendor Fields | Included | Not included |
| OnBase Doc Type | PUR - * | FIN - * |

## Duplicate Checking

Both scripts check OnBase to avoid re-importing existing attachments:

### VBScript
```vbscript
Function OnBaseAttachmentIDFound(pOBJ_ATT_UNID)
  strSQL = "SELECT COUNT(*) AS NUM_FOUND FROM hsi.keyitem481 " &_
           "WHERE hsi.keyitem481.keyvaluebig = " & pOBJ_ATT_UNID & " " &_
           "AND (SELECT itemtypenum FROM hsi.itemdata " &_
           "WHERE hsi.itemdata.itemnum = hsi.keyitem481.itemnum) = 267"
  ' Returns True if found
End Function
```

### PowerShell
```powershell
function Test-OnBaseAttachmentExists {
    param(
        [System.Data.Odbc.OdbcConnection]$Connection,
        [string]$AttachmentID
    )
    $sql = @"
SELECT COUNT(*) AS NUM_FOUND 
FROM hsi.keyitem481 
WHERE hsi.keyitem481.keyvaluebig = $AttachmentID 
AND (SELECT itemtypenum FROM hsi.itemdata 
     WHERE hsi.itemdata.itemnum = hsi.keyitem481.itemnum) = 267
"@
    # Returns $true if found
}
```

## File Naming

Both scripts use GUID-based unique filenames:

### VBScript
```vbscript
strFileNameGUID = Mid(strFileName, 1, (intPos - 1)) & _
                  "_[" & LCase(GetRandomHexValue) & "]." & strExt
```

### PowerShell
```powershell
$fileNameGUID = "${baseName}_[${guidHex}]${extension}"
```

Example: `invoice_[a1b2c3d4e5f6789012345678901234].pdf`

## Error Handling

### VBScript
```vbscript
On Error Resume Next
' ... code ...
If Err.Number <> 0 Then
  WriteLog(Err.Description)
  WScript.Quit(12)
End If
On Error Goto 0
```

### PowerShell
```powershell
try {
    # ... code ...
} catch {
    WriteLog "Error: $_"
    WriteLog $_.ScriptStackTrace
    throw
}
```

## Performance Considerations

| Aspect | VBScript | PowerShell |
|--------|----------|------------|
| **File I/O** | Multiple passes through XML files | Direct database queries |
| **Processing** | Sequential XML file processing | Batch database retrieval |
| **Memory** | Loads entire XML into memory | Streams data from database |
| **Complexity** | High (XML parsing, BIRT, printing) | Low (focused on attachments) |

## Migration Path

If you need to add features from the VBScript to the PowerShell version:

### 1. Add Related Document Support (like RQ for PO)
```powershell
# Query for related documents
$relatedQuery = @"
SELECT DISTINCT B.OBJ_ATT_PG_UNID AS RF_OBJ_ATT_PG_UNID,
       A.RF_DOC_CD, A.RF_DOC_DEPT_CD, A.RF_DOC_ID, A.RF_DOC_VERS_NO
FROM O_FINPROD.R_DOC_RF A, O_FINPROD.[RELATED_HDR] B
WHERE A.DOC_CD = 'JVA' AND A.RF_DOC_CD = '[RELATED_TYPE]'
  AND A.DOC_ID = '$DOC_ID'
  -- ... additional joins
"@
```

### 2. Add Department-Specific Document Types
```powershell
$docTypeName = "FIN - JVA Attachments"
switch ($DOC_DEPT_CD) {
    "040" { $docTypeName = "FIN - JVA Attachments - EL" }
    "060" { $docTypeName = "FIN - JVA Attachments - PA" }
    "805" { $docTypeName = "FIN - JVA Attachments - TC" }
}
```

### 3. Add Archiving
```powershell
function Archive-JVAFiles {
    param([string]$ArchiveFolder)
    # Use 7-Zip to archive processed files
    & "C:\Program Files\7-Zip\7z.exe" a -t7z "$ArchiveFolder\JVA_$(Get-Date -Format 'yyyyMMdd').7z" "$OutputPath\*"
}
```

## Advantages of PowerShell Version

1. **Simpler**: Focused on one task (attachments)
2. **Modern**: Uses PowerShell features and .NET classes
3. **Maintainable**: Clearer structure with functions and regions
4. **Flexible**: Easy to add parameters and customize
5. **Debuggable**: Better error messages and stack traces
6. **Portable**: Can run on any system with PowerShell 5.1+

## When to Use Each

### Use VBScript Version When:
- Processing PO/DO/MA documents
- Need BIRT report generation
- Need XML manipulation
- Need printing functionality
- Need receiving copies
- Full purchasing workflow required

### Use PowerShell Version When:
- Processing JVA documents
- Only need attachment extraction
- Want simpler, more maintainable code
- Need to integrate with other PowerShell scripts
- Want better error handling and logging

## Summary

The PowerShell JVA processor is a focused, streamlined version of the VBScript purchasing processor. It maintains the core attachment retrieval logic while removing the complexity of XML processing, BIRT integration, and printing workflows. This makes it ideal for JVA document processing where only attachment extraction is needed.

