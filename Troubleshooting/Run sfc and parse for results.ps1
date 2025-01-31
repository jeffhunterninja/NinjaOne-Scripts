# This will run SFC in verifyonly mode and report the result into a custom field based off the presence of the keyword "protection"

$OriginalEncoding = [console]::OutputEncoding
[console]::OutputEncoding = [Text.Encoding]::Unicode
$SFC = & $env:SystemRoot\System32\sfc.exe /VerifyOnly
[console]::OutputEncoding = $OriginalEncoding
 
# Condensing output for readibility in activities feed
$SFC | Select-Object -First 10
Write-Host "[...]"
$SFC | Select-Object -Last 10 

# Looking for integrity as keyword
$integrityStatus = $SFC | Where-Object { $_ -match "protection" }

# Write status into custom field
Ninja-Property-Set sfc $integrityStatus

# When I was debugging this I wanted to see the output, but I don't think this is really necessary for most people.
# $SFC | Ninja-Property-Set-Piped sfcOutput