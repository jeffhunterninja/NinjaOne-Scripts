<#
This is provided as an educational example of how to interact with the NinjaOne API using the authorization code and refresh token flow.
Any scripts should be evaluated and tested in a controlled setting before being utilized in production.
As this script is an educational example, further improvements and enhancement may be necessary to handle larger datasets.

Description:
    This script demonstrates the OAuth2 "Authorization Code" flow plus a token refresh to obtain a Bearer token
    for NinjaOne API calls. It spins up a local HTTP listener to intercept the "authorization code" from the
    NinjaOne OAuth callback, exchanges it for tokens, and then uses the refresh token to obtain a final access token.

How the script works:
    1) Define basic variables: $ClientID, $Secret, $Instance, $RedirectURL.
    2) Create a function that starts a local HTTP listener on $RedirectURL, launches a browser to NinjaOne's
       authorization page, and waits for the returned authorization code.
    3) Exchange the authorization code for an initial token (including refresh token).
    4) Use the refresh token to obtain the final Bearer token.
    5) Prepare an Authorization header with the Bearer token for subsequent NinjaOne API calls.

Security Note:
    For demonstration, this script stores client_secret in a plain text variable. In a production environment,
    store such secrets more securely (e.g., via the PowerShell Secret Management Module, Windows Credential Manager,
    or Azure Key Vault). Also protect file permissions carefully.

#>

$NinjaOneInstance     = "app.ninjarmm.com"
$NinjaOneClientId     = "-"
$NinjaOneClientSecret = "-"
$redirect_uri = "http://localhost:8888/"
$auth_url = "https://$NinjaOneInstance/ws/oauth/authorize"

# Ensure System.Web Assembly is loaded
# This assembly is required to parse the query strings from the URL callback
try {
    [System.Web.HttpUtility] | Out-Null
}
catch {
    Add-Type -AssemblyName System.Web
}


# Start Local HTTP Server to Capture Auth Code

try {
    # The listener will respond to GET requests with the 'code' parameter in the query string
    Write-Host "Starting HTTP server to listen for callback to $redirect_uri ..."
    $httpServer = [System.Net.HttpListener]::new()
    $httpServer.Prefixes.Add($redirect_uri)
    $httpServer.Start()

    # Launch Browser to NinjaOne OAuth Page
    Write-Host "Launching NinjaOne API OAuth authorization page $auth_url ..."
    # Build the full authorization URL with query parameters
    $AuthURL = "https://$Instance/ws/oauth/authorize?response_type=code&client_id=$NinjaOneClientId&client_secret=$NinjaOneSecret&redirect_uri=$redirect_uri&state=custom_state&scope=monitoring%20management%20offline_access"
    Start-Process $AuthURL

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
catch {
    Write-Error "Failed to retrieve authorization code from the NinjaOne API. Error: $_"
    exit
}

# Exchange Authorization Code for an Initial Token
Write-Host "Exchanging Authorization Code for tokens ..."

$AuthBody = @{
    'grant_type' = 'authorization_code'
    'client_id' = $NinjaOneClientID
    'client_secret' = $NinjaOneClientSecret
    'code' = $authorization_code
    'redirect_uri' = $redirect_uri
    'scope' = "monitoring management offline_access"
}

try {
    $Response = Invoke-WebRequest -Uri "https://$NinjaOneInstance/ws/oauth/token" -Method POST -Body $AuthBody -ContentType 'application/x-www-form-urlencoded'
}
catch {
    Write-Error "Failed to connect to NinjaOne API. Error: $_"
    exit
}
# Store the refresh token for subsequent requests
$RefreshToken = ($Response.Content | ConvertFrom-Json).refresh_token

Write-Host "Initial token obtained. Refresh token is:" $RefreshToken

# Build Authorization Header
$AccessToken = ($Response.Content | ConvertFrom-Json).access_token
$AuthHeader = @{
    'Authorization' = "Bearer $AccessToken"
}

Write-Host "`nFinal Access Token obtained. You can use '$($AuthHeader.Authorization)' in your API calls."
Write-Host "Done!"
