# Processor Script Upgrade - Start-Process Implementation

## Summary

The `ERP_JVA_Attachments_Processor.ps1` orchestration script has been upgraded to use `Start-Process` instead of the call operator (`&`) for executing child scripts. This provides better process isolation, error handling, and control.

## What Changed

### Before (Call Operator)

```powershell
try {
    & $downloadScript

    if ($LASTEXITCODE -ne 0 -and $null -ne $LASTEXITCODE) {
        throw "Download script failed with exit code: $LASTEXITCODE"
    }
} catch {
    WriteLog "ERROR in download step: $_"
    exit 1
}
```

### After (Start-Process)

```powershell
try {
    $downloadProcess = Start-Process -FilePath "powershell.exe" `
        -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$downloadScript`"" `
        -Wait `
        -PassThru `
        -NoNewWindow

    if ($downloadProcess.ExitCode -ne 0) {
        throw "Download script failed with exit code: $($downloadProcess.ExitCode)"
    }
} catch {
    WriteLog "ERROR in download step: $_"
    exit 1
}
```

---

## Benefits of Start-Process

### 1. **Process Isolation**
- Each script runs in its own PowerShell process
- Memory is cleaned up after each step
- No variable pollution between scripts
- Better resource management

### 2. **Better Error Handling**
- Explicit exit code checking via `$process.ExitCode`
- More reliable than `$LASTEXITCODE`
- Clear process success/failure status

### 3. **Execution Control**
- `-NoProfile` - Faster startup, no user profile loading
- `-ExecutionPolicy Bypass` - Ensures scripts run regardless of policy
- `-Wait` - Waits for completion before continuing
- `-PassThru` - Returns process object for exit code checking
- `-NoNewWindow` - Runs in same console (cleaner output)

### 4. **Logging Visibility**
- All output from child scripts appears in parent console
- Easier to track progress
- Better for scheduled task logging

---

## Scripts Updated

All 5 child scripts now run in separate processes:

1. **Cleanup Script** (`ERP_JVA_Attachments_CleanUp.ps1`)
   - Runs in isolated process
   - Cleans up old files before download

2. **Download Script** (`ERP_JVA_Attachments_Downloader.ps1`)
   - Runs in isolated process
   - Downloads JVA attachments from ERP
   - **Now uses fast chunked download!**

3. **DIP Script** (`ERP_JVA_Attachments_DIP.ps1`)
   - Runs in isolated process
   - Creates DIP files for OnBase import

4. **Validation Script** (`ERP_JVA_Attachments_Validation.ps1`)
   - Runs in isolated process
   - Validates attachments were imported to OnBase

5. **Archiver Script** (`ERP_JVA_Attachments_Archiver.ps1`)
   - Runs in isolated process
   - Archives processed files

---

## Process Flow

```
┌─────────────────────────────────────────────────────────────┐
│ ERP_JVA_Attachments_Processor.ps1 (Orchestrator)           │
└─────────────────────────────────────────────────────────────┘
                            │
                            ▼
        ┌───────────────────────────────────────┐
        │ STEP 0: Cleanup Old Files             │
        │ Start-Process → CleanUp.ps1           │
        │ Wait for completion                   │
        │ Check exit code                       │
        └───────────────────────────────────────┘
                            │
                            ▼
        ┌───────────────────────────────────────┐
        │ STEP 1: Download Attachments          │
        │ Start-Process → Downloader.ps1        │
        │ Wait for completion                   │
        │ Check exit code                       │
        └───────────────────────────────────────┘
                            │
                            ▼
        ┌───────────────────────────────────────┐
        │ STEP 2: Create DIP Files              │
        │ Start-Process → DIP.ps1               │
        │ Wait for completion                   │
        │ Check exit code                       │
        └───────────────────────────────────────┘
                            │
                            ▼
        ┌───────────────────────────────────────┐
        │ STEP 3: Wait 15 Minutes               │
        │ Allow OnBase to process DIP           │
        └───────────────────────────────────────┘
                            │
                            ▼
        ┌───────────────────────────────────────┐
        │ STEP 4: Validate in OnBase            │
        │ Start-Process → Validation.ps1        │
        │ Wait for completion                   │
        │ Check exit code                       │
        └───────────────────────────────────────┘
                            │
                            ▼
        ┌───────────────────────────────────────┐
        │ STEP 5: Archive Files                 │
        │ Start-Process → Archiver.ps1          │
        │ Wait for completion                   │
        │ Check exit code                       │
        └───────────────────────────────────────┘
                            │
                            ▼
        ┌───────────────────────────────────────┐
        │ Final Summary & Statistics            │
        └───────────────────────────────────────┘
```

---

## Start-Process Parameters Explained

### `-FilePath "powershell.exe"`
- Launches a new PowerShell process
- Uses system PowerShell (not PowerShell Core)

### `-ArgumentList`
Arguments passed to PowerShell:

| Argument | Purpose |
|----------|---------|
| `-NoProfile` | Skip loading user profile (faster startup) |
| `-ExecutionPolicy Bypass` | Ignore execution policy restrictions |
| `-File "script.ps1"` | Script to execute |

### `-Wait`
- Blocks until the process completes
- Ensures sequential execution
- Required for orchestration flow

### `-PassThru`
- Returns the process object
- Allows checking `$process.ExitCode`
- Essential for error detection

### `-NoNewWindow`
- Runs in same console window
- Output appears in parent console
- Better for logging and monitoring

---

## Error Handling

### Exit Code Checking

Each script's exit code is checked:

```powershell
if ($downloadProcess.ExitCode -ne 0) {
    throw "Download script failed with exit code: $($downloadProcess.ExitCode)"
}
```

**Exit Codes:**
- `0` = Success
- `1` = Error (standard error code)
- Other = Script-specific error codes

### Error Propagation

If any step fails:
1. Error is logged with `WriteLog`
2. Stack trace is logged
3. Processor exits with code `1`
4. Subsequent steps are skipped

---

## Comparison: Call Operator vs Start-Process

| Aspect | Call Operator (`&`) | Start-Process |
|--------|---------------------|---------------|
| **Process Isolation** | ❌ Same process | ✅ Separate process |
| **Memory Cleanup** | ❌ Manual | ✅ Automatic |
| **Exit Code** | `$LASTEXITCODE` (unreliable) | `$process.ExitCode` (reliable) |
| **Execution Policy** | ❌ Inherits | ✅ Can override |
| **Profile Loading** | ❌ Inherits | ✅ Can skip |
| **Output Control** | ✅ Direct | ✅ Configurable |
| **Error Handling** | ⚠️ Complex | ✅ Simple |
| **Resource Usage** | ✅ Lower | ⚠️ Slightly higher |

---

## Performance Impact

### Memory Usage
- **Before:** All scripts share same process memory
- **After:** Each script gets fresh process (better cleanup)

### Execution Time
- **Overhead:** ~100-200ms per script launch
- **Total overhead:** ~500-1000ms for 5 scripts
- **Negligible** compared to script execution time (minutes)

### Benefits Outweigh Costs
- Better reliability
- Cleaner execution
- Easier debugging
- Worth the minimal overhead

---

## Logging Output

### Example Console Output

```
=========================================
Starting JVA Attachment Processing Orchestration
=========================================

Scripts verified:
  - Cleanup: N:\...\ERP_JVA_Attachments_CleanUp.ps1
  - Download: N:\...\ERP_JVA_Attachments_Downloader.ps1
  - DIP: N:\...\ERP_JVA_Attachments_DIP.ps1
  - Validation: N:\...\ERP_JVA_Attachments_Validation.ps1
  - Archiver: N:\...\ERP_JVA_Attachments_Archiver.ps1

=========================================
STEP 0: Cleaning Up Old Files
=========================================
[Cleanup script output appears here...]

Cleanup completed successfully in 2.5 seconds

=========================================
STEP 1: Downloading JVA Attachments
=========================================
[Download script output appears here...]
JVA 100 CV12251005 [v1] Downloading attachment: [1] [1685781] file.pdf
JVA 100 CV12251005 [v1] SUCCESS: Downloaded 4057.89 KB in 12.5s (324.63 KB/s)

Download completed successfully in 125.3 seconds

=========================================
STEP 2: Creating DIP Files
=========================================
[DIP script output appears here...]

DIP file creation completed successfully in 5.2 seconds

=========================================
STEP 3: Waiting 15 minutes before validation
=========================================
Waiting to allow OnBase to process the DIP file...
Wait started at: 01/15/2025 14:30:00
Wait completed at: 01/15/2025 14:45:00

=========================================
STEP 4: Validating Attachments in OnBase
=========================================
[Validation script output appears here...]

Validation completed successfully in 15.8 seconds

=========================================
STEP 5: Archiving Processed Files
=========================================
[Archiver script output appears here...]

Archiving completed successfully in 8.3 seconds

=========================================
JVA Attachment Processing Complete
=========================================
Total execution time: 17.2 minutes
  - Cleanup: 2.5 seconds
  - Download: 125.3 seconds
  - DIP Creation: 5.2 seconds
  - Wait Time: 900 seconds (15 minutes)
  - Validation: 15.8 seconds
  - Archiving: 8.3 seconds
=========================================

Process completed successfully
```

---

## Troubleshooting

### Script Not Found Error

**Error:**
```
Start-Process : This command cannot be run due to the error: The system cannot find the file specified.
```

**Solution:**
- Verify script paths are correct
- Check scripts exist in expected location
- Ensure quotes around script path

### Exit Code Always 0

**Problem:** Script fails but exit code is 0

**Solution:**
- Ensure child scripts call `exit 1` on error
- Check child scripts have proper error handling
- Verify `try/catch` blocks in child scripts

### Output Not Visible

**Problem:** Can't see child script output

**Solution:**
- Ensure `-NoNewWindow` parameter is used
- Check child scripts use `WriteLog` or `Write-Host`
- Verify logging is configured correctly

### Execution Policy Error

**Error:**
```
File cannot be loaded because running scripts is disabled on this system
```

**Solution:**
- Already handled by `-ExecutionPolicy Bypass`
- If still occurs, check system-level restrictions
- May need administrator privileges

---

## Testing Recommendations

### 1. Test Individual Scripts
```powershell
# Test each script runs successfully in new process
Start-Process -FilePath "powershell.exe" `
    -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ".\ERP_JVA_Attachments_Downloader.ps1" `
    -Wait -PassThru -NoNewWindow
```

### 2. Test Error Handling
```powershell
# Temporarily modify a child script to exit with error
# Verify processor catches error and stops
```

### 3. Test Full Orchestration
```powershell
# Run full processor script
.\ERP_JVA_Attachments_Processor.ps1

# Verify all steps complete
# Check logs for any issues
```

### 4. Monitor Resource Usage
```powershell
# Watch memory usage during execution
# Verify processes are cleaned up after each step
```

---

## Backward Compatibility

✅ **Fully backward compatible**

- Child scripts unchanged (except Downloader has fast download)
- Same execution flow
- Same logging format
- Same error handling behavior
- Same output files

**No changes required** to child scripts or downstream processes.

---

## Summary

✅ **Better isolation** - Each script in separate process  
✅ **Reliable exit codes** - Direct process.ExitCode checking  
✅ **Cleaner execution** - No profile loading, bypass policy  
✅ **Better logging** - All output visible in console  
✅ **Easier debugging** - Clear process boundaries  
✅ **Minimal overhead** - ~1 second total for 5 scripts  

The upgrade provides better reliability and maintainability with negligible performance impact!

