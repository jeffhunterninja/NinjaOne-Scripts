function Get-DeviceJoinState {
    $result = [ordered]@{
        AzureAdJoined   = $false
        DomainJoined    = $false
        WorkplaceJoined = $false
        JoinType        = "Unknown"
    }

    $out = & dsregcmd /status 2>$null
    if (-not $out) {
        $result.JoinType = "dsregcmd_unavailable"
        return [pscustomobject]$result
    }

    $result.AzureAdJoined   = ($out | Select-String -Pattern '^\s*AzureAdJoined\s*:\s*YES\s*$') -ne $null
    $result.DomainJoined    = ($out | Select-String -Pattern '^\s*DomainJoined\s*:\s*YES\s*$') -ne $null
    $result.WorkplaceJoined = ($out | Select-String -Pattern '^\s*WorkplaceJoined\s*:\s*YES\s*$') -ne $null

    if ($result.AzureAdJoined -and $result.DomainJoined) {
        $result.JoinType = "Hybrid_Entra+AD"
    } elseif ($result.AzureAdJoined) {
        $result.JoinType = "Entra_Joined"
    } elseif ($result.DomainJoined) {
        $result.JoinType = "AD_Domain_Joined"
    } elseif ($result.WorkplaceJoined) {
        $result.JoinType = "Workplace_Registered"
    } else {
        $result.JoinType = "Workgroup/Unknown"
    }

    [pscustomobject]$result
}

# Example
$join = Get-DeviceJoinState

Ninja-Property-Set deviceJoinState $($join.JoinType)