# Setup API Function
# Handles SharePoint configuration and setup wizard

using namespace System.Net

param($Request, $TriggerMetadata)

# Import required modules
Import-Module Az.KeyVault -Force

function Get-GraphToken {
    try {
        # Get credentials from Key Vault
        $keyVaultUrl = $env:KEY_VAULT_URL
        $keyVaultName = $keyVaultUrl.Replace('https://', '').Replace('.vault.azure.net/', '')
        
        $clientId = (Get-AzKeyVaultSecret -VaultName $keyVaultName -Name 'ClientId' -AsPlainText)
        $clientSecret = (Get-AzKeyVaultSecret -VaultName $keyVaultName -Name 'ClientSecret' -AsPlainText)
        $tenantId = (Get-AzKeyVaultSecret -VaultName $keyVaultName -Name 'TenantId' -AsPlainText)
        
        # Get access token
        $tokenUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
        $body = @{
            client_id = $clientId
            client_secret = $clientSecret
            scope = "https://graph.microsoft.com/.default"
            grant_type = "client_credentials"
        }
        
        $response = Invoke-RestMethod -Uri $tokenUrl -Method Post -ContentType "application/x-www-form-urlencoded" -Body $body
        return $response.access_token
        
    } catch {
        Write-Warning "Failed to get Graph token: $($_.Exception.Message)"
        return $null
    }
}

function Invoke-GraphRequest {
    param(
        [string]$Uri,
        [string]$Method = "GET",
        [object]$Body = $null
    )
    
    $token = Get-GraphToken
    if (-not $token) {
        throw "Unable to get Graph API token"
    }
    
    $headers = @{
        "Authorization" = "Bearer $token"
        "Content-Type" = "application/json"
    }
    
    try {
        if ($Body) {
            $jsonBody = $Body | ConvertTo-Json -Depth 10
            return Invoke-RestMethod -Uri $Uri -Method $Method -Headers $headers -Body $jsonBody
        } else {
            return Invoke-RestMethod -Uri $Uri -Method $Method -Headers $headers
        }
    } catch {
        Write-Warning "Graph API request failed: $($_.Exception.Message)"
        throw $_
    }
}

# Main function logic
$method = $Request.Method
$query = $Request.Query
$body = $Request.Body

Write-Host "Setup API called with method: $method, action: $($query.action)"

try {
    switch ($method) {
        "GET" {
            switch ($query.action) {
                "getConsentUrl" {
                    $keyVaultUrl = $env:KEY_VAULT_URL
                    $keyVaultName = $keyVaultUrl.Replace('https://', '').Replace('.vault.azure.net/', '')
                    
                    $adminConsentUrl = Get-AzKeyVaultSecret -VaultName $keyVaultName -Name 'AdminConsentUrl' -AsPlainText
                    
                    $response = @{
                        adminConsentUrl = $adminConsentUrl
                        status = "success"
                    }
                }
                
                "checkConsent" {
                    # Test if we can make a Graph API call
                    try {
                        $token = Get-GraphToken
                        if ($token) {
                            # Test call to verify permissions
                            Invoke-GraphRequest -Uri "https://graph.microsoft.com/v1.0/sites" -Method GET
                            $response = @{
                                consentGranted = $true
                                status = "success"
                            }
                        } else {
                            $response = @{
                                consentGranted = $false
                                status = "no_token"
                            }
                        }
                    } catch {
                        $response = @{
                            consentGranted = $false
                            status = "error"
                            error = $_.Exception.Message
                        }
                    }
                }
                
                "getSites" {
                    # Get SharePoint sites
                    $sites = Invoke-GraphRequest -Uri "https://graph.microsoft.com/v1.0/sites?search=*"
                    
                    $response = $sites.value | Select-Object @{
                        Name = 'id'
                        Expression = { $_.id }
                    }, @{
                        Name = 'displayName' 
                        Expression = { $_.displayName }
                    }, @{
                        Name = 'webUrl'
                        Expression = { $_.webUrl }
                    }, @{
                        Name = 'description'
                        Expression = { $_.description }
                    }
                }
                
                "getLists" {
                    if (-not $query.siteUrl) {
                        throw "siteUrl parameter required"
                    }
                    
                    # Get site ID from URL
                    $siteUrl = [System.Web.HttpUtility]::UrlDecode($query.siteUrl)
                    $hostname = ([Uri]$siteUrl).Host
                    $sitePath = ([Uri]$siteUrl).AbsolutePath
                    
                    # Get site
                    $site = Invoke-GraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/$hostname`:$sitePath"
                    
                    # Get lists from site
                    $lists = Invoke-GraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists"
                    
                    # Filter for relevant lists
                    $response = $lists.value | Where-Object { 
                        $_.name -notlike "_*" -and
                        $_.name -ne "Form Templates" -and
                        $_.name -ne "Site Assets"
                    } | Select-Object @{
                        Name = 'id'
                        Expression = { $_.id }
                    }, @{
                        Name = 'displayName'
                        Expression = { $_.displayName }
                    }, @{
                        Name = 'name'
                        Expression = { $_.name }
                    }
                }
                
                "analyzeFields" {
                    if (-not $query.siteUrl -or -not $query.listId) {
                        throw "siteUrl and listId parameters required"
                    }
                    
                    $siteUrl = [System.Web.HttpUtility]::UrlDecode($query.siteUrl)
                    $hostname = ([Uri]$siteUrl).Host
                    $sitePath = ([Uri]$siteUrl).AbsolutePath
                    
                    # Get site
                    $site = Invoke-GraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/$hostname`:$sitePath"
                    
                    # Get list columns
                    $columns = Invoke-GraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists/$($query.listId)/columns"
                    
                    # Filter and format columns
                    $response = $columns.value | Where-Object { 
                        $_.readOnly -eq $false -and
                        $_.name -ne "ContentType" -and
                        $_.name -ne "Attachments"
                    } | Select-Object @{
                        Name = 'name'
                        Expression = { $_.name }
                    }, @{
                        Name = 'displayName'
                        Expression = { $_.displayName }
                    }, @{
                        Name = 'type'
                        Expression = { 
                            if ($_.text) { "text" }
                            elseif ($_.number) { "number" }  
                            elseif ($_.dateTime) { "dateTime" }
                            else { "other" }
                        }
                    }
                }
                
                default {
                    # Return simple setup page
                    $htmlContent = @"
<!DOCTYPE html>
<html>
<head><title>SharePoint Contact Sync Setup</title></head>
<body>
<h1>SharePoint Contact Sync Setup</h1>
<p>API Endpunkte:</p>
<ul>
<li><a href="?action=checkConsent">Check Consent Status</a></li>
<li><a href="?action=getSites">Get SharePoint Sites</a></li>
</ul>
<p>Status: Function App läuft!</p>
</body>
</html>
"@
                    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
                        StatusCode = [HttpStatusCode]::OK
                        Headers = @{ 'Content-Type' = 'text/html; charset=utf-8' }
                        Body = $htmlContent
                    })
                    return
                }
            }
        }
        
        "POST" {
            $actionData = $body | ConvertFrom-Json
            
            switch ($actionData.action) {
                "saveConfig" {
                    # Save configuration to Key Vault
                    $config = $actionData.config
                    
                    try {
                        $keyVaultUrl = $env:KEY_VAULT_URL
                        $keyVaultName = $keyVaultUrl.Replace('https://', '').Replace('.vault.azure.net/', '')
                        
                        $sharePointConfig = @{
                            sharepoint_configs = @(
                                @{
                                    id = "config-1"
                                    name = "Primary Configuration"
                                    site_url = $config.siteUrl
                                    list_id = $config.listId
                                    list_name = $config.listName
                                    enabled = $true
                                    field_mapping = $config.fieldMappings
                                    last_sync = $null
                                }
                            )
                        }
                        
                        $configJson = $sharePointConfig | ConvertTo-Json -Depth 10 -Compress
                        Set-AzKeyVaultSecret -VaultName $keyVaultName -Name 'SharePointConfig' -SecretValue (ConvertTo-SecureString $configJson -AsPlainText -Force)
                        
                        $response = @{
                            success = $true
                            message = "Configuration saved successfully"
                        }
                        
                    } catch {
                        $response = @{
                            success = $false
                            error = $_.Exception.Message
                        }
                    }
                }
                
                default {
                    throw "Unknown POST action: $($actionData.action)"
                }
            }
        }
    }
    
    # Return JSON response
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Headers = @{ 'Content-Type' = 'application/json' }
        Body = ($response | ConvertTo-Json -Depth 10)
    })
    
} catch {
    Write-Error "Setup API error: $($_.Exception.Message)"
    
    $errorResponse = @{
        success = $false
        error = $_.Exception.Message
        timestamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
    }
    
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::BadRequest
        Headers = @{ 'Content-Type' = 'application/json' }
        Body = ($errorResponse | ConvertTo-Json -Depth 3)
    })
}