# SharePoint Contact Sync - Main Function
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
        Write-Error "Failed to get Graph token: $($_.Exception.Message)"
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
        Write-Error "Graph API request failed: $($_.Exception.Message)"
        throw $_
    }
}

function Get-SharePointContacts {
    param($Config)
    
    try {
        # Parse site URL
        $hostname = ([Uri]$Config.site_url).Host
        $sitePath = ([Uri]$Config.site_url).AbsolutePath
        
        # Get site
        $site = Invoke-GraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/$hostname`:$sitePath"
        
        # Get list items
        $items = Invoke-GraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists/$($Config.list_id)/items?expand=fields"
        
        $contacts = @()
        foreach ($item in $items.value) {
            $contact = @{}
            
            # Apply field mappings
            foreach ($mapping in $Config.field_mapping.PSObject.Properties) {
                $graphField = $mapping.Name -replace '_', '.'
                $spField = $mapping.Value
                
                if ($item.fields.$spField) {
                    switch ($graphField) {
                        "emailAddresses[0].address" {
                            $contact["emailAddresses"] = @(@{
                                address = $item.fields.$spField
                                name = "$($item.fields.($Config.field_mapping.givenName)) $($item.fields.($Config.field_mapping.surname))"
                            })
                        }
                        "businessPhones[0]" {
                            $contact["businessPhones"] = @($item.fields.$spField)
                        }
                        default {
                            $contact[$graphField] = $item.fields.$spField
                        }
                    }
                }
            }
            
            # Only add contacts with required fields
            if ($contact.givenName -and $contact.surname) {
                $contacts += $contact
            }
        }
        
        return $contacts
        
    } catch {
        Write-Error "Failed to get SharePoint contacts: $($_.Exception.Message)"
        throw $_
    }
}

function Sync-ContactsToUsers {
    param($Contacts, $TargetUsers = @())
    
    $syncResults = @{
        totalContacts = $Contacts.Count
        successfulSyncs = 0
        errors = @()
        users = @()
    }
    
    try {
        # If no target users specified, sync to all users (for demo)
        if ($TargetUsers.Count -eq 0) {
            Write-Host "Getting all users for demo sync..."
            $users = Invoke-GraphRequest -Uri "https://graph.microsoft.com/v1.0/users?`$top=10&`$select=id,userPrincipalName"
            $TargetUsers = $users.value | ForEach-Object { $_.userPrincipalName }
        }
        
        foreach ($userEmail in $TargetUsers) {
            try {
                Write-Host "Syncing $($Contacts.Count) contacts to user: $userEmail"
                
                # Get existing contacts for this user
                $existingContacts = Invoke-GraphRequest -Uri "https://graph.microsoft.com/v1.0/users/$userEmail/contacts"
                
                $userSyncResult = @{
                    user = $userEmail
                    added = 0
                    updated = 0
                    errors = @()
                }
                
                foreach ($contact in $Contacts) {
                    try {
                        # Check if contact already exists (by email)
                        $existingContact = $existingContacts.value | Where-Object { 
                            $_.emailAddresses[0].address -eq $contact.emailAddresses[0].address 
                        }
                        
                        if ($existingContact) {
                            # Update existing contact
                            $updateUri = "https://graph.microsoft.com/v1.0/users/$userEmail/contacts/$($existingContact.id)"
                            Invoke-GraphRequest -Uri $updateUri -Method PATCH -Body $contact
                            $userSyncResult.updated++
                            Write-Host "Updated contact: $($contact.givenName) $($contact.surname)"
                        } else {
                            # Create new contact
                            $createUri = "https://graph.microsoft.com/v1.0/users/$userEmail/contacts"
                            Invoke-GraphRequest -Uri $createUri -Method POST -Body $contact
                            $userSyncResult.added++
                            Write-Host "Added contact: $($contact.givenName) $($contact.surname)"
                        }
                        
                        $syncResults.successfulSyncs++
                        
                    } catch {
                        $error = "Failed to sync contact $($contact.givenName) $($contact.surname): $($_.Exception.Message)"
                        $userSyncResult.errors += $error
                        $syncResults.errors += $error
                        Write-Warning $error
                    }
                }
                
                $syncResults.users += $userSyncResult
                
            } catch {
                $error = "Failed to sync to user $userEmail`: $($_.Exception.Message)"
                $syncResults.errors += $error
                Write-Error $error
            }
        }
        
    } catch {
        $error = "Failed to sync contacts: $($_.Exception.Message)"
        $syncResults.errors += $error
        Write-Error $error
    }
    
    return $syncResults
}

# Main execution
try {
    Write-Host "SharePoint Contact Sync started"
    
    # Get configuration from Key Vault
    $keyVaultUrl = $env:KEY_VAULT_URL
    $keyVaultName = $keyVaultUrl.Replace('https://', '').Replace('.vault.azure.net/', '')
    
    $configJson = Get-AzKeyVaultSecret -VaultName $keyVaultName -Name 'SharePointConfig' -AsPlainText
    if (-not $configJson) {
        throw "SharePoint configuration not found. Please run setup first."
    }
    
    $config = $configJson | ConvertFrom-Json
    $activeConfigs = $config.sharepoint_configs | Where-Object { $_.enabled -eq $true }
    
    if ($activeConfigs.Count -eq 0) {
        throw "No active SharePoint configurations found."
    }
    
    $totalSyncResults = @{
        configurations = $activeConfigs.Count
        totalContacts = 0
        totalSyncs = 0
        errors = @()
        details = @()
    }
    
    foreach ($spConfig in $activeConfigs) {
        try {
            Write-Host "Processing configuration: $($spConfig.name)"
            
            # Get contacts from SharePoint
            $contacts = Get-SharePointContacts -Config $spConfig
            Write-Host "Found $($contacts.Count) contacts in SharePoint list: $($spConfig.list_name)"
            
            # Sync to target users
            $syncResult = Sync-ContactsToUsers -Contacts $contacts -TargetUsers $spConfig.target_users
            
            $configResult = @{
                config_name = $spConfig.name
                site_url = $spConfig.site_url
                list_name = $spConfig.list_name
                contacts_found = $contacts.Count
                sync_result = $syncResult
                timestamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
            }
            
            $totalSyncResults.details += $configResult
            $totalSyncResults.totalContacts += $contacts.Count
            $totalSyncResults.totalSyncs += $syncResult.successfulSyncs
            $totalSyncResults.errors += $syncResult.errors
            
            # Update last sync time
            $spConfig.last_sync = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
            
        } catch {
            $error = "Failed to process config $($spConfig.name): $($_.Exception.Message)"
            $totalSyncResults.errors += $error
            Write-Error $error
        }
    }
    
    # Update configuration with last sync times
    $updatedConfigJson = $config | ConvertTo-Json -Depth 10 -Compress
    Set-AzKeyVaultSecret -VaultName $keyVaultName -Name 'SharePointConfig' -SecretValue (ConvertTo-SecureString $updatedConfigJson -AsPlainText -Force)
    
    $response = @{
        status = "success"
        message = "SharePoint Contact Sync completed"
        sync_results = $totalSyncResults
        timestamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
    }
    
    Write-Host "SharePoint Contact Sync completed successfully"
    Write-Host "Total configurations processed: $($totalSyncResults.configurations)"
    Write-Host "Total contacts found: $($totalSyncResults.totalContacts)"
    Write-Host "Total successful syncs: $($totalSyncResults.totalSyncs)"
    Write-Host "Total errors: $($totalSyncResults.errors.Count)"

} catch {
    Write-Error "SharePoint Contact Sync failed: $($_.Exception.Message)"
    
    $response = @{
        status = "error"
        message = $_.Exception.Message
        timestamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
    }
}

# Return response
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Headers = @{ 'Content-Type' = 'application/json' }
    Body = ($response | ConvertTo-Json -Depth 10)
})