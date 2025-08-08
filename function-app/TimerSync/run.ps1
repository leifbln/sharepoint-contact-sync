# Timer-triggered SharePoint Contact Sync
# Runs every hour to keep contacts in sync

param($Timer)

Write-Host "Timer-triggered SharePoint Contact Sync started at: $(Get-Date)"

try {
    # Call the main SharePoint sync function
    $functionUrl = $env:WEBSITE_HOSTNAME
    $syncEndpoint = "https://$functionUrl/api/SharePointSync"
    
    $body = @{
        source = "timer"
        timestamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
    } | ConvertTo-Json
    
    # Invoke the sync function
    $response = Invoke-RestMethod -Uri $syncEndpoint -Method Post -ContentType "application/json" -Body $body
    
    Write-Host "Timer sync completed successfully"
    Write-Host "Response: $($response | ConvertTo-Json -Depth 3)"
    
    Write-Error "Timer sync failed: $($_.Exception.Message)"
    
    # Log error to Application Insights
    $errorData = @{
        timestamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
        source = "timer"
        error = $_.Exception.Message
        stackTrace = $_.ScriptStackTrace
    }
    
    Write-Host "Error details: $($errorData | ConvertTo-Json -Depth 3)"
}