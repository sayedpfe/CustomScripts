# Test webhook with detailed diagnostics
$config = Get-Content ".\automation-config.json" | ConvertFrom-Json

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Webhook Detailed Test" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Test data
$testData = @{
    SiteUrl = "https://m365cpi90282478.sharepoint.com/sites/DeutschKurs"
    AuthenticationContextName = "c1"
    RequestorEmail = "test@m365cpi90282478.onmicrosoft.com"
    RequestId = "TEST-DETAILED-001"
}

$body = $testData | ConvertTo-Json
Write-Host "`nRequest Body:" -ForegroundColor Yellow
Write-Host $body -ForegroundColor White

Write-Host "`nSending to webhook..." -ForegroundColor Yellow

try {
    $response = Invoke-WebRequest -Uri $config.WebhookUrl `
                                  -Method POST `
                                  -Body $body `
                                  -ContentType "application/json" `
                                  -UseBasicParsing

    Write-Host "`n✅ Response received:" -ForegroundColor Green
    Write-Host "Status: $($response.StatusCode) $($response.StatusDescription)" -ForegroundColor White
    Write-Host "Job ID: $($response.Headers['x-ms-request-id'])" -ForegroundColor White

} catch {
    Write-Host "`n❌ Error:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}

Write-Host "`nWaiting 60 seconds for job to complete..." -ForegroundColor Yellow
Start-Sleep -Seconds 60

# Now check the job
Write-Host "`nChecking for jobs created in the last 2 minutes..." -ForegroundColor Yellow
$recentJobs = Get-AzAutomationJob -ResourceGroupName $config.ResourceGroupName `
                                  -AutomationAccountName $config.AutomationAccountName `
                                  -RunbookName "applyconditionalaccesspolicy" |
              Where-Object { $_.CreationTime -gt (Get-Date).AddMinutes(-2) } |
              Sort-Object CreationTime -Descending

if ($recentJobs) {
    foreach ($job in $recentJobs) {
        Write-Host "`n========================================" -ForegroundColor Cyan
        Write-Host "Job: $($job.JobId)" -ForegroundColor Yellow
        Write-Host "Status: $($job.Status)" -ForegroundColor $(if($job.Status -eq 'Completed'){'Green'}elseif($job.Status -eq 'Failed'){'Red'}else{'Yellow'})
        Write-Host "Created: $($job.CreationTime)" -ForegroundColor White
        
        # Get full output
        $output = Get-AzAutomationJobOutput -ResourceGroupName $config.ResourceGroupName `
                                            -AutomationAccountName $config.AutomationAccountName `
                                            -Id $job.JobId `
                                            -Stream Any |
                  Sort-Object Time
        
        Write-Host "`nOutput:" -ForegroundColor Cyan
        foreach ($record in $output) {
            $detail = Get-AzAutomationJobOutputRecord -ResourceGroupName $config.ResourceGroupName `
                                                      -AutomationAccountName $config.AutomationAccountName `
                                                      -JobId $job.JobId `
                                                      -Id $record.StreamRecordId
            $color = switch ($record.StreamType) {
                'Error' { 'Red' }
                'Warning' { 'Yellow' }
                default { 'White' }
            }
            Write-Host $detail.Value.Value -ForegroundColor $color
        }
    }
} else {
    Write-Host "No recent jobs found!" -ForegroundColor Red
}
