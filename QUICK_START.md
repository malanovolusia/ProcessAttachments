# Quick Start Guide - JVA Attachment Processing

## What This Script Does

Extracts attachments from Journal Voucher (JVA) documents in the ERP system and prepares them for import into OnBase.

## Prerequisites

✅ PowerShell 5.1 or higher  
✅ Access to ERP Oracle database  
✅ Access to OnBase SQL Server database  
✅ Network access to `\\erp311script\Library\PSM1\`

## Quick Start (3 Steps)

### Step 1: Open PowerShell

```powershell
cd "N:\Projects\35442 - JVA DIP\ProcessAttachments"
```

### Step 2: Run the Script

**For Testing:**
```powershell
.\ERP_Process_JVA_Dip_Files.ps1
```

**For Production:**
```powershell
.\ERP_Process_JVA_Dip_Files.ps1 -SID "ERP19PRO" -OutputPath "\\server\share\JVA\attachments"
```

### Step 3: Check the Output

Look in the output directory for:
- `!JVA_attachment_indexes_DIP.txt` - DIP file for OnBase import
- `!JVA_indexes.txt` - Index file for tracking
- Individual attachment files with GUID names

## What Happens When You Run It

```
1. Connects to ERP database
2. Connects to OnBase database
3. Queries for JVA documents with attachments
4. For each JVA document:
   ├─ Retrieves attachment metadata
   ├─ Checks if already in OnBase (optional)
   ├─ Downloads attachment BLOB
   ├─ Saves to disk with unique name
   ├─ Calculates SHA-256 hash
   ├─ Writes DIP entry
   └─ Writes index entry
5. Closes files and connections
6. Reports statistics
```

## Example Output

```
[2024-01-15 10:30:15] Starting JVA Attachment Processing
[2024-01-15 10:30:15] SID: ERP19TEST
[2024-01-15 10:30:15] Output Path: .\Output\JVA\attachments
[2024-01-15 10:30:16] Connected to ERP database
[2024-01-15 10:30:16] Connected to OnBase database
[2024-01-15 10:30:17] Processing JVA Document #1 : JVA 100 JV123456 v1
[2024-01-15 10:30:17] Retrieving attachments for JVA 100 JV123456 - 1
[2024-01-15 10:30:18] JVA 100 JV123456 [v1] Saving attachment: [1] [12345] invoice_[abc123].pdf
[2024-01-15 10:30:19] Processing Complete
[2024-01-15 10:30:19] Total JVA documents processed: 1
[2024-01-15 10:30:19] Total attachments found: 1
[2024-01-15 10:30:19] Total attachments saved: 1
```

## Common Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-SID` | ERP19TEST | Oracle database (ERP19PRO for production) |
| `-OutputPath` | .\Output\JVA\attachments | Where to save files |
| `-CheckDuplicates` | $true | Skip attachments already in OnBase |

## Troubleshooting

### "Cannot connect to database"
- Check VPN connection
- Verify SID parameter is correct
- Check database credentials

### "No attachments found"
- Verify JVA documents have attachments in ERP
- Check the SQL query in the script
- Verify OBJ_ATT_PG_UNID is populated

### "Access denied to output path"
- Check folder permissions
- Try a local path first: `-OutputPath "C:\Temp\JVA"`

### "Module not found"
- Verify network access to `\\erp311script\Library\PSM1\`
- Check if you're on the correct network

## Next Steps

1. **Review the output** - Check DIP and index files
2. **Test with OnBase** - Import the DIP file
3. **Verify attachments** - Confirm they appear in OnBase
4. **Schedule it** - Set up as a scheduled task if needed

## Need Help?

- See `README.md` for detailed documentation
- See `COMPARISON.md` for differences from PO processing
- See `Example_Usage.ps1` for more examples

## Files in This Project

```
ProcessAttachments/
├── ERP_Process_JVA_Dip_Files.ps1  ← Main script
├── README.md                       ← Full documentation
├── QUICK_START.md                  ← This file
├── COMPARISON.md                   ← VBScript vs PowerShell comparison
├── Example_Usage.ps1               ← Usage examples
└── References/
    └── ERP_ncp_purchasing_post_processor.wsf  ← Original VBScript reference
```

## Important Notes

⚠️ **Test First**: Always test with `-SID "ERP19TEST"` before running in production  
⚠️ **Check Output**: Review the DIP file before importing to OnBase  
⚠️ **Backup**: Keep a copy of attachments before importing  
⚠️ **Duplicates**: Use `-CheckDuplicates $false` only for testing  

## Quick Reference Commands

```powershell
# Test run (default settings)
.\ERP_Process_JVA_Dip_Files.ps1

# Production run
.\ERP_Process_JVA_Dip_Files.ps1 -SID "ERP19PRO" -OutputPath "\\server\share\JVA"

# Test without duplicate checking
.\ERP_Process_JVA_Dip_Files.ps1 -CheckDuplicates $false

# Get help
Get-Help .\ERP_Process_JVA_Dip_Files.ps1 -Full

# View the script
notepad .\ERP_Process_JVA_Dip_Files.ps1
```

---

**Ready to go?** Just run: `.\ERP_Process_JVA_Dip_Files.ps1`

