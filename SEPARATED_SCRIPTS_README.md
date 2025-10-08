# JVA Attachment Processing - Separated Scripts

## Overview

The JVA attachment processing has been separated into two distinct scripts for better modularity and flexibility:

1. **Download_JVA_Attachments.ps1** - Downloads attachments from ERP database
2. **Create_JVA_DIP_Files.ps1** - Creates DIP files from downloaded attachments
3. **Process_JVA_Attachments.ps1** - Master script to run both

## Why Separate?

### Benefits

✅ **Modularity** - Run download and DIP creation independently  
✅ **Flexibility** - Re-create DIP files without re-downloading  
✅ **Debugging** - Easier to troubleshoot specific steps  
✅ **Performance** - Download once, create DIP multiple times if needed  
✅ **Testing** - Test DIP format changes without database queries  

### Use Cases

- **Download only** - Get attachments for backup or review
- **DIP only** - Regenerate DIP files with different settings
- **Both** - Complete end-to-end processing

## File Structure

```
ProcessAttachments/
├── Download_JVA_Attachments.ps1      ← Downloads attachments
├── Create_JVA_DIP_Files.ps1          ← Creates DIP files
├── Process_JVA_Attachments.ps1       ← Master script (runs both)
├── ERP_Process_JVA_Dip_Files.ps1     ← Original combined script (deprecated)
└── Attachments/
    ├── !JVA_attachment_index.txt     ← Index file (created by download)
    ├── !JVA_attachment_indexes_DIP.txt ← DIP file (created by DIP script)
    └── [attachment files]            ← Downloaded files
```

## Scripts

### 1. Download_JVA_Attachments.ps1

**Purpose:** Download attachments from ERP database

**What it does:**
- Connects to ERP Oracle database
- Connects to OnBase SQL Server database
- Queries JVA documents with attachments
- Downloads attachment BLOBs to disk
- Checks for duplicates in OnBase (optional)
- Creates index file with metadata

**Output:**
- `Attachments/!JVA_attachment_index.txt` - Index of downloaded files
- `Attachments/[filename]_[guid].[ext]` - Attachment files

**Index File Format:**
```
OBJ_ATT_UNID|ATT_DATE|GUID_FILENAME|FULL_PATH|USER_ID|DESCRIPTION|DOC_ID|DEPT_CD|DOC_ID_OUT|DOC_TYP|DOC_CD|VERS_NO|SG_UNID|SEQ_NO|STATUS|TYPE|COMP_NM|COMP_DESC|ORIGINAL_FILENAME
12345|01/10/2024|invoice_[guid].pdf|C:\path\file.pdf|JSMITH|Invoice|JV123456|100|JV123456|JV|JVA|1|67890|1|1|1|Component|Context|invoice.pdf
```

**Usage:**
```powershell
.\Download_JVA_Attachments.ps1
```

### 2. Create_JVA_DIP_Files.ps1

**Purpose:** Create DIP files for OnBase import

**What it does:**
- Reads the index file created by download script
- Calculates SHA-256 hashes for each file
- Generates DIP entries for OnBase import
- Creates properly formatted DIP file

**Input:**
- `Attachments/!JVA_attachment_index.txt` - Index file
- `Attachments/[files]` - Attachment files

**Output:**
- `Attachments/!JVA_attachment_indexes_DIP.txt` - DIP file for OnBase

**DIP File Format:**
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

**Usage:**
```powershell
.\Create_JVA_DIP_Files.ps1
```

### 3. Process_JVA_Attachments.ps1

**Purpose:** Master script to run both operations

**What it does:**
- Orchestrates download and DIP creation
- Provides mode selection
- Handles errors between steps

**Usage:**
```powershell
# Run both download and DIP creation
.\Process_JVA_Attachments.ps1

# Download only
.\Process_JVA_Attachments.ps1 -Mode Download

# Create DIP only (requires existing downloads)
.\Process_JVA_Attachments.ps1 -Mode CreateDIP

# Run both (explicit)
.\Process_JVA_Attachments.ps1 -Mode Both
```

## Workflow

### Complete Process (Both)

```
1. Download_JVA_Attachments.ps1
   ├─ Connect to databases
   ├─ Query JVA documents
   ├─ Download attachments
   └─ Create index file
   
2. Create_JVA_DIP_Files.ps1
   ├─ Read index file
   ├─ Calculate SHA-256 hashes
   ├─ Generate DIP entries
   └─ Create DIP file
```

### Download Only

```
1. Download_JVA_Attachments.ps1
   ├─ Connect to databases
   ├─ Query JVA documents
   ├─ Download attachments
   └─ Create index file
   
[Stop here - review files, backup, etc.]
```

### DIP Creation Only

```
[Assumes attachments already downloaded]

1. Create_JVA_DIP_Files.ps1
   ├─ Read existing index file
   ├─ Calculate SHA-256 hashes
   ├─ Generate DIP entries
   └─ Create DIP file
```

## Common Scenarios

### Scenario 1: First Time Processing

```powershell
# Run complete process
.\Process_JVA_Attachments.ps1
```

### Scenario 2: Re-create DIP with Different Settings

```powershell
# Attachments already downloaded, just recreate DIP
.\Process_JVA_Attachments.ps1 -Mode CreateDIP
```

### Scenario 3: Download for Backup

```powershell
# Just download, don't create DIP yet
.\Process_JVA_Attachments.ps1 -Mode Download
```

### Scenario 4: Test DIP Format Changes

```powershell
# 1. Download once
.\Process_JVA_Attachments.ps1 -Mode Download

# 2. Modify Create_JVA_DIP_Files.ps1 as needed

# 3. Re-create DIP (fast, no database queries)
.\Process_JVA_Attachments.ps1 -Mode CreateDIP

# 4. Repeat steps 2-3 until DIP format is correct
```

## Advantages Over Combined Script

| Aspect | Combined Script | Separated Scripts |
|--------|----------------|-------------------|
| **Flexibility** | Run all or nothing | Run steps independently |
| **Testing** | Re-download to test DIP | Test DIP without re-downloading |
| **Performance** | Always queries database | Query once, create DIP many times |
| **Debugging** | Hard to isolate issues | Easy to debug specific step |
| **Maintenance** | One large file | Smaller, focused files |
| **Reusability** | Limited | High - use download for other purposes |

## Migration from Original Script

If you were using `ERP_Process_JVA_Dip_Files.ps1`:

**Before:**
```powershell
.\ERP_Process_JVA_Dip_Files.ps1
```

**After (equivalent):**
```powershell
.\Process_JVA_Attachments.ps1
```

**Or run steps separately:**
```powershell
# Step 1
.\Download_JVA_Attachments.ps1

# Step 2
.\Create_JVA_DIP_Files.ps1
```

## File Sizes

| Script | Lines | Purpose |
|--------|-------|---------|
| Download_JVA_Attachments.ps1 | ~290 | Download attachments |
| Create_JVA_DIP_Files.ps1 | ~180 | Create DIP files |
| Process_JVA_Attachments.ps1 | ~90 | Master orchestrator |
| **Total** | **~560** | **All functionality** |

Compare to original: 386 lines (but less flexible)

## Error Handling

Each script handles errors independently:

- **Download script** - Exits if database connection fails
- **DIP script** - Exits if index file not found
- **Master script** - Stops if any step fails

## Next Steps

1. **Test download:** `.\Download_JVA_Attachments.ps1`
2. **Review files:** Check `Attachments/` folder
3. **Test DIP creation:** `.\Create_JVA_DIP_Files.ps1`
4. **Review DIP:** Check `!JVA_attachment_indexes_DIP.txt`
5. **Import to OnBase:** Use the DIP file

## Troubleshooting

### "Index file not found"
- Run download script first: `.\Download_JVA_Attachments.ps1`

### "No attachments downloaded"
- Check database connections
- Verify JVA documents have attachments
- Check `OBJ_ATT_PG_UNID` is populated

### "DIP file is empty"
- Check index file has data
- Verify attachment files exist
- Check file paths in index file

## Summary

The separated scripts provide:
- ✅ Better modularity
- ✅ Easier testing
- ✅ Faster iteration
- ✅ Independent execution
- ✅ Clearer responsibilities

Use the master script for convenience, or run individual scripts for more control!

