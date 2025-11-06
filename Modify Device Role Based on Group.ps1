<#

This is provided as an educational example of how to interact with the NinjaAPI.
Any scripts should be evaluated and tested in a controlled setting before being utilized in production.
As this script is an educational example, further improvements and enhancement may be necessary to handle larger datasets.

#>

# Your NinjaRMM credentials
$NinjaOneInstance = 'ca.ninjarmm.com' # This varies depending on region or environment. For example, if you are in the US, this would be '$NinjaOneInstance'
$NinjaOneClientId = ''
$NinjaOneClientSecret = ''

# Add in group ID (all devices in group will be moved to specific organization and location)
$GroupID = '222'
$organizationId = '8'
$locationId = '18'

$body = @{
  grant_type = "client_credentials"
  client_id = $NinjaOneClientId
  client_secret = $NinjaOneClientSecret
  scope = "monitoring management"
}
    
$API_AuthHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$API_AuthHeaders.Add("accept", 'application/json')
$API_AuthHeaders.Add("Content-Type", 'application/x-www-form-urlencoded')
   
$auth_token = Invoke-RestMethod -Uri https://$NinjaOneInstance/oauth/token -Method POST -Headers $API_AuthHeaders -Body $body
$access_token = $auth_token | Select-Object -ExpandProperty 'access_token' -EA 0
   
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("accept", 'application/json')
$headers.Add("Authorization", "Bearer $access_token")
    

$groupdevices_url = "https://$NinjaOneInstance/v2/group/$GroupID/device-ids"

# Retrieves all devices IDs in that group
$groupdevices = Invoke-RestMethod -Uri $groupdevices_url -Method GET -Headers $headers
  

# For each device ID found in that group, an API call is made to move the device to specific organization and location
  foreach ($key in $groupdevices) {

  # define ninja urls
  $deviceupdates_url = "https://$NinjaOneInstance/api/v2/device/" + $key

  # define request body - need to find desired role ID as definied in https://$NinjaOneInstance/api/v2/roles
  $request_body = @{
    organizationId = $organizationId
    locationId = $locationId
  }

  # convert body to JSON
  $json = $request_body | ConvertTo-Json

  Write-Host "Assigning device role to" $key

  # Let's make the magic happen
  Invoke-RestMethod -Method 'Patch' -Uri $deviceupdates_url -Headers $headers -Body $json -ContentType "application/json" -Verbose
  }
