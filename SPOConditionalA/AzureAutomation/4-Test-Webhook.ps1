# Test the Azure Automation Webhook
# This simulates a Power Automate call to the webhook

param(
    [Parameter(Mandatory=$false)]
    [string]$ConfigPath = ".\automation-config.json",
    
    [Parameter(Mandatory=$false)]
    [string]$SiteUrl = "https://m365cpi90282478.sharepoint.com/sites/DeutschKurs",
    
    [Parameter(Mandatory=$false)]
    [string]$AuthContextName = "Sensitive Information - Guest terms of Use",
    
    [Parameter(Mandatory=$false)]
    [string]$RequestorEmail = "test@m365cpi90282478.onmicrosoft.com"
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Testing Azure Automation Webhook" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Load configuration
if (-not (Test-Path $ConfigPath)) {
    Write-Host "❌ Configuration file not found: $ConfigPath" -ForegroundColor Red
    exit 1
}

$config = Get-Content $ConfigPath | ConvertFrom-Json

if (-not $config.WebhookUrl) {
    Write-Host "❌ Webhook URL not found in config" -ForegroundColor Red
    Write-Host "Please run 3-Setup-Webhook.ps1 first" -ForegroundColor Yellow
    exit 1
}

Write-Host "Test Parameters:" -ForegroundColor Cyan
Write-Host "  Site URL: $SiteUrl" -ForegroundColor White
Write-Host "  Auth Context: $AuthContextName" -ForegroundColor White
Write-Host "  Requestor: $RequestorEmail" -ForegroundColor White

# Prepare webhook payload
$requestId = [guid]::NewGuid().ToString()
$payload = @{
    SiteUrl = $SiteUrl
    AuthenticationContextName = $AuthContextName
    RequestorEmail = $RequestorEmail
    RequestId = $requestId
} | ConvertTo-Json

Write-Host "`n📌 Sending request to webhook..." -ForegroundColor Cyan
Write-Host "Payload:" -ForegroundColor Gray
Write-Host $payload -ForegroundColor Gray

try {
    # Send POST request to webhook
    $response = Invoke-WebRequest `
        -Uri $config.WebhookUrl `
        -Method Post `
        -Body $payload `
        -ContentType "application/json" `
        -UseBasicParsing
    
    Write-Host "`n✅ Webhook accepted the request!" -ForegroundColor Green
    Write-Host "Status Code: $($response.StatusCode)" -ForegroundColor White
    Write-Host "Status: $($response.StatusDescription)" -ForegroundColor White
    
    Write-Host @"

⏳ The runbook is now executing in the background...

To check the job status:
1. Go to Azure Portal
2. Navigate to: Automation Account → $($config.AutomationAccountName) → Jobs
3. Look for the most recent job for runbook: $($config.RunbookName)
4. Click on the job to see detailed output

The runbook typically takes 30-60 seconds to complete.

"@ -ForegroundColor Yellow
    
    # Wait a bit and try to get job status
    Write-Host "Waiting 10 seconds before checking job status..." -ForegroundColor Gray
    Start-Sleep -Seconds 10
    
    Write-Host "`n📌 Checking recent jobs..." -ForegroundColor Cyan
    
    # Connect to Azure if not already connected
    $context = Get-AzContext
    if (-not $context) {
        Connect-AzAccount -SubscriptionId $config.SubscriptionId | Out-Null
    }
    
    # Get recent jobs
    $recentJobs = Get-AzAutomationJob `
        -ResourceGroupName $config.ResourceGroupName `
        -AutomationAccountName $config.AutomationAccountName `
        -RunbookName $config.RunbookName | 
        Sort-Object StartTime -Descending | 
        Select-Object -First 5
    
    Write-Host "`nRecent Jobs (last 5):" -ForegroundColor Cyan
    $recentJobs | Format-Table -Property @(
        @{Label="Start Time"; Expression={$_.StartTime}; Width=20}
        @{Label="Status"; Expression={$_.Status}; Width=15}
        @{Label="Job ID"; Expression={$_.JobId}; Width=36}
    ) -AutoSize
    
    # Try to get output from most recent job
    $latestJob = $recentJobs | Select-Object -First 1
    
    if ($latestJob -and ($latestJob.Status -eq "Completed" -or $latestJob.Status -eq "Failed")) {
        Write-Host "`n📌 Getting output from latest completed job..." -ForegroundColor Cyan
        
        $jobOutput = Get-AzAutomationJobOutput `
            -ResourceGroupName $config.ResourceGroupName `
            -AutomationAccountName $config.AutomationAccountName `
            -Id $latestJob.JobId `
            -Stream Output
        
        if ($jobOutput) {
            Write-Host "`nJob Output:" -ForegroundColor Cyan
            foreach ($output in $jobOutput) {
                $outputRecord = Get-AzAutomationJobOutputRecord `
                    -ResourceGroupName $config.ResourceGroupName `
                    -AutomationAccountName $config.AutomationAccountName `
                    -JobId $latestJob.JobId `
                    -Id $output.StreamRecordId
                
                Write-Host $outputRecord.Value.Value -ForegroundColor White
            }
        }
    } elseif ($latestJob) {
        Write-Host "`nLatest job is still running (Status: $($latestJob.Status))" -ForegroundColor Yellow
        Write-Host "Check Azure Portal for real-time progress" -ForegroundColor Yellow
    }
    
} catch {
    Write-Host "`n❌ Webhook request failed" -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    
    if ($_.Exception.Response) {
        Write-Host "Status Code: $($_.Exception.Response.StatusCode.value__)" -ForegroundColor Red
    }
    exit 1
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Test Complete" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
