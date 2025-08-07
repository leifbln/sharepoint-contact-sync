using namespace System.Net

param($Request, $TriggerMetadata)

Write-Host "SharePoint Contact Sync - GitHub Release Version"

try {
    $response = @{
        status = "success"
        message = "🎉 SharePoint Contact Sync via GitHub Release!"
        deployment_info = @{
            deployment_id = $env:DEPLOYMENT_ID
            customer_name = $env:CUSTOMER_NAME
            admin_email = $env:ADMIN_EMAIL
            timestamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
        }
        method = "GitHub Release ZIP"
        version = "1.0.0"
    }
    
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body = ($response | ConvertTo-Json -Depth 3)
        Headers = @{"Content-Type" = "application/json"}
    })
    
} catch {
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::InternalServerError
        Body = "Error: $($_.Exception.Message)"
    })
}