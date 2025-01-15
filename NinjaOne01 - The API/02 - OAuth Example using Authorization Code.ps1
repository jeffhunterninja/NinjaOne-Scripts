<#

This is provided as an educational example of how to interact with the NinjaAPI with the authorization code grant type and the "Web" application platform.
Any scripts should be evaluated and tested in a controlled setting before being utilized in production.
As this script is an educational example, further improvements and enhancement may be necessary to handle larger datasets.

# Attributions:
Ryan Southwell for the HTTP listener that retrieves the authorization code

# How this script works:
#   1) Define essential variables for OAuth (client_id, client_secret, redirect_uri, scope, etc.).
#   2) Load or add the System.Web assembly if not already loaded; required to parse URL query strings.
#   3) Start a local HTTP server that listens for the authorization code (callback from NinjaOne).
#   4) Open (launch) a browser window to NinjaOne's OAuth authorization page.
#   5) Once authorized, the script captures the "code" parameter from the redirected URL.
#   6) Use that authorization code, along with your client credentials, to request an access token.
#   7) Define headers with the new Bearer token.
#   8) Make an example API call (e.g. retrieve a list of organizations).
#   9) Output the results to the console.

#   DISCLAIMER: This script intentionally stores OAuth credentials in variables (for demonstration).
#   In a production environment, consider storing credentials more securely using:
#       - Windows Credential Manager
#       - Secret Management Module in PowerShell
#       - Azure Key Vault (in Azure environments)
#   and always restrict file permissions for any stored secrets.

#>

$NinjaOneInstance     = "app.ninjarmm.com"
$NinjaOneClientId     = "-"
$NinjaOneClientSecret = "-"
$redirect_uri = "http://localhost:8888/"

# OAuth2 scope and authorization URL
$scope = "monitoring management"
$auth_url = "https://$NinjaOneInstance/ws/oauth/authorize"

# This assembly is required to parse the query strings from the URL callback
Try {
    [System.Web.HttpUtility] | Out-Null
}
Catch {
    Add-Type -AssemblyName System.Web
}

# Start Local HTTP Server to Capture Auth Code
# The listener will respond to GET requests with the 'code' parameter in the query string
Write-Host "Starting HTTP server to listen for callback to $redirect_uri ..."
$httpServer = [System.Net.HttpListener]::new()
$httpServer.Prefixes.Add($redirect_uri)
$httpServer.Start()


# Launch Browser to NinjaOne OAuth Page
Try {
Write-Host "Launching NinjaOne API OAuth authorization page $auth_url ..."
# Build the full authorization URL with query parameters
$auth_redirect_url = $auth_url + "?response_type=code&client_id=" + $NinjaOneClientId + "&redirect_uri=" + $redirect_uri + "&state=custom_state&scope=monitoring%20management"
Start-Process $auth_redirect_url

Write-Host "Listening for authorization code from local callback to $redirect_uri ..."

# Listen for the Authorization Code
while ($httpServer.IsListening) {
    $httpContext   = $httpServer.GetContext()
    $httpRequest   = $httpContext.Request
    $httpResponse  = $httpContext.Response
    $httpRequestURL = [uri]($httpRequest.Url)

    if ($httpRequest.IsLocal) {
        Write-Host "Heard local request to $httpRequestURL ..."
        # Parse the query string to see if it contains the authorization code
        $httpRequestQuery = [System.Web.HttpUtility]::ParseQueryString($httpRequestURL.Query)

        if (-not [string]::IsNullOrEmpty($httpRequestQuery['code'])) {
            # Store the code if present
            $authorization_code = $httpRequestQuery['code']
            $httpResponse.StatusCode = 200

            # Simple HTML to display success message in the browser
            [string]$httpResponseContent = "<html><body>Authorized! You may now close this browser tab/window.</body></html>"
            $httpResponseBuffer = [System.Text.Encoding]::UTF8.GetBytes($httpResponseContent)
            $httpResponse.ContentLength64 = $httpResponseBuffer.Length
            $httpResponse.OutputStream.Write($httpResponseBuffer, 0, $httpResponse.ContentLength64)
        }
        else {
            Write-Host "HTTP 400: Missing 'code' parameter in URL query string."
            $httpResponse.StatusCode = 400
        }
    }
    else {
        # Reject any non-local request to our listener
        Write-Host "HTTP 403: Blocking remote request to $httpRequestURL ..."
        $httpResponse.StatusCode = 403
    }

    # Close the connection
    $httpResponse.Close()

    # Stop the server once we have the authorization code
    if (-not [string]::IsNullOrEmpty($authorization_code)) {
        $httpServer.Stop()
    }
}

Write-Host "Parsed authorization code: $authorization_code"
}
Catch {
    Write-Error "Failed to retrieve authorization code from NinjaOne API. Error: $_"
    exit
}

# Prepare headers for token request
$API_AuthHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$API_AuthHeaders.Add("accept", "application/json")
$API_AuthHeaders.Add("Content-Type", "application/x-www-form-urlencoded")

# Body for token request
$body = @{
    grant_type    = "authorization_code"
    client_id     = $NinjaOneClientId
    client_secret = $NinjaOneClientSecret
    redirect_uri  = $redirect_uri
    scope         = $scope
    code          = $authorization_code
}

try {
    Write-Host "Requesting access token from NinjaOne ..."
    $auth_token  = Invoke-RestMethod -Uri "https://$NinjaOneInstance/ws/oauth/token" -Method POST -Headers $API_AuthHeaders -Body $body
}
catch {
    Write-Error "Failed to retrieve access token from NinjaOne API. Error: $_"
    exit
}

# Extract the access token from the returned JSON
$access_token = $auth_token | Select-Object -ExpandProperty 'access_token' -EA 0
# Check if we successfully obtained an access token
if (-not $access_token) {
    Write-Host "Failed to obtain access token. Please check your client ID and client secret."
    exit
}
Write-Host "Retrieved access token: $access_token"
# Build headers with access token to make API calls
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("accept", "application/json")
$headers.Add("Authorization", "Bearer $access_token")

# Define Ninja URLs
$devices_url = "https://$NinjaOneInstance/v2/devices-detailed"
$organizations_url = "https://$NinjaOneInstance/v2/organizations-detailed"

# Call Ninja URLs to get data
try {
    $devices = Invoke-RestMethod -Uri $devices_url -Method GET -Headers $headers
    $organizations = Invoke-RestMethod -Uri $organizations_url -Method GET -Headers $headers
}
catch {
    Write-Error "Failed to retrieve organizations and devices from NinjaOne API. Error: $_"
    exit
}
# Extend organizations objects with additional properties to classify devices
Foreach ($organization in $organizations) {
    Add-Member -InputObject $organization -NotePropertyName "Workstations" -NotePropertyValue @()
    Add-Member -InputObject $organization -NotePropertyName "Servers" -NotePropertyValue @()
}

# Loop through all devices and copy each device to corresponding organization, with separate properties for storing servers and workstations
Foreach ($device in $devices) {
    $currentOrg = $organizations | Where-Object {$_.id -eq $device.organizationId}
    if ($device.nodeClass.EndsWith("_SERVER")) {
        $currentOrg.servers += $device.systemName
    } elseif ($device.nodeClass.EndsWith("_WORKSTATION") -or $device.nodeClass -eq "MAC") {
        $currentOrg.workstations += $device.systemName
    }
}

# Create and display a summary report of organizations and their device counts broken down by servers and workstations, plus total devices
$reportSummary = Foreach ($organization in $organizations) {
    [PSCustomObject]@{
        Name = $organization.Name
        Workstations = $organization.workstations.length
        Servers = $organization.servers.length
        TotalDevices = ($organization.workstations.length + $organization.servers.length)
    }
}

# Display the summary report in a table format
$reportSummary | Format-Table | Out-String

