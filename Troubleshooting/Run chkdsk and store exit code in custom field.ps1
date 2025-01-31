# Run as system
# Specify the target drive here - may need to modify if the device is not using C as its main system drive
$targetDrive = 'C:' 

# Run chkdsk to check the disk and capture the output
$chkdskOutput = chkdsk $targetDrive | Out-String

# Log the output to a file
$chkdskOutput | Out-File -FilePath "C:\chkdsk_log.txt"

# Create a custom object to store the exit code and corresponding status message
$result = New-Object -TypeName PSObject -Property @{
    ExitCode = $LASTEXITCODE
    Message = switch ($LASTEXITCODE) {
        0 {"No errors found."}
        1 {"Errors found and fixed."}
        2 {"Disk cleanup needed."}
        3 {"Errors found but not fixed."}
        default {"chkdsk ran with an unexpected exit code: $LASTEXITCODE"}
    }
    DetailedOutput = $chkdskOutput
}

# Output the result object
$result

Ninja-Property-Set chkdskStatus $result.Message