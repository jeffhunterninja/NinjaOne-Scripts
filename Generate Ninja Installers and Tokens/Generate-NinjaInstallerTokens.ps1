$ProgressPreference = 'SilentlyContinue'

$NinjaOneClientID = ''
$NinjaOneClientSecret = ''
$NinjaOneInstance = 'ca.ninjarmm.com'

$OutPath = 'c:\temp\InstallerTokens.csv'

# Authenticate and get bearer token
$AuthBody = @{
    'grant_type'    = 'client_credentials'
    'client_id'     = $NinjaOneClientID
    'client_secret' = $NinjaOneClientSecret
    'scope'         = 'monitoring management' 
}

$Result = Invoke-WebRequest -uri "https://$($NinjaOneInstance)/ws/oauth/token" -Method POST -Body $AuthBody -ContentType 'application/x-www-form-urlencoded'

$NinjaAuthHeader = @{
    'Authorization' = "Bearer $(($Result.content | ConvertFrom-Json).access_token)"
}

# Function to generate installer
Function Get-NinjaAdvancedInstaller {
    Param (
        $OrgID,
        $LocID,
        $Type,
        $NodeRoleId
    )

    $InstallerBody = (@{
            organizationId = $OrgID
            locationId     = $LocID
            installerType  = $Type
            content        = @{
                nodeRoleId = $NodeRoleId
            }
        }) | ConvertTo-Json -Depth 100

    $Result = (Invoke-WebRequest -uri "https://$($NinjaOneInstance)/api/v2/organization/generate-installer" `
        -Method POST -Headers $NinjaAuthHeader -ContentType 'application/json' -Body $InstallerBody).content | ConvertFrom-Json -depth 100
    Return $Result
}

# Get Ninja Organisations
$After = 0
$PageSize = 1000
$NinjaOrgs = do {
    $Result = (Invoke-WebRequest -uri "https://$($NinjaOneInstance)/api/v2/organizations?pageSize=$PageSize&after=$After" `
        -Method GET -Headers $NinjaAuthHeader -ContentType 'application/json').content | ConvertFrom-Json -depth 100
    $Result
    $ResultCount = ($Result.id | Measure-Object -Maximum)
    $After = $ResultCount.maximum
} while ($ResultCount.count -eq $PageSize)

# Get Ninja Locations
$After = 0
$PageSize = 1000
$NinjaLocs = do {
    $Result = (Invoke-WebRequest -uri "https://$($NinjaOneInstance)/api/v2/locations?pageSize=$PageSize&after=$After" `
        -Method GET -Headers $NinjaAuthHeader -ContentType 'application/json').content | ConvertFrom-Json -depth 100
    $Result
    $ResultCount = ($Result.id | Measure-Object -Maximum)
    $After = $ResultCount.maximum
} while ($ResultCount.count -eq $PageSize)

# Get Roles and filter for WINDOWS nodeClass
$Roles = (Invoke-WebRequest -uri "https://$($NinjaOneInstance)/api/v2/roles" `
    -Method GET -Headers $NinjaAuthHeader -ContentType 'application/json').content | ConvertFrom-Json -Depth 100

$WindowsRoles = $Roles | Where-Object { $_.nodeClass -match 'WINDOWS' }

# Generate Installers for each location and each Windows role
$Tokens = foreach ($Location in $NinjaLocs) {
    $Org = $NinjaOrgs | Where-Object { $_.id -eq $Location.organizationId }

    foreach ($Role in $WindowsRoles) {
        $Installer = Get-NinjaAdvancedInstaller -OrgID $Location.organizationId -LocID $Location.id -Type 'WINDOWS_MSI' -NodeRoleId $Role.id

        $Token = (($Installer.url -split 'NinjaOneAgent_')[1]) -replace '\.pkg','' -replace '\.msi',''

        [PSCustomObject]@{
            OrgName   = $Org.Name
            LocName   = $Location.Name
            OrgID     = $Location.organizationId
            LocID     = $Location.id
            RoleName  = $Role.name
            RoleID    = $Role.id
            Token     = $Token
        }
    }
}

$Tokens | Export-CSV $OutPath -NoTypeInformation
