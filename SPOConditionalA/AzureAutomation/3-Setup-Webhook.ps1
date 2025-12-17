# Setup Webhook for Power Automate Integration
# This creates a webhook URL that Power Automate will call

param(
    [Parameter(Mandatory=$false)]
    [string]$ConfigPath = ".\automation-config.json"
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Creating Webhook for Power Automate" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Load configuration
if (-not (Test-Path $ConfigPath)) {
    Write-Host "❌ Configuration file not found: $ConfigPath" -ForegroundColor Red
    Write-Host "Please run 1-Setup-AutomationAccount.ps1 and 2-Create-Runbook.ps1 first" -ForegroundColor Yellow
    exit 1
}

$config = Get-Content $ConfigPath | ConvertFrom-Json
Write-Host "Loaded configuration from: $ConfigPath" -ForegroundColor Gray

if (-not $config.RunbookName) {
    Write-Host "❌ Runbook name not found in config" -ForegroundColor Red
    Write-Host "Please run 2-Create-Runbook.ps1 first" -ForegroundColor Yellow
    exit 1
}

# Connect to Azure
Write-Host "`n📌 Connecting to Azure..." -ForegroundColor Cyan
Connect-AzAccount -SubscriptionId $config.SubscriptionId | Out-Null
Write-Host "   ✅ Connected" -ForegroundColor Green

# Create webhook
$webhookName = "PowerAutomate-ConditionalAccess-Webhook"
$expiryDate = (Get-Date).AddYears(2)

Write-Host "`n📌 Creating Webhook: $webhookName..." -ForegroundColor Cyan

try {
    # Check if webhook already exists
    $existingWebhook = Get-AzAutomationWebhook `
        -ResourceGroupName $config.ResourceGroupName `
        -AutomationAccountName $config.AutomationAccountName `
        -Name $webhookName `
        -ErrorAction SilentlyContinue
    
    if ($existingWebhook) {
        Write-Host "   ⚠️  Webhook already exists" -ForegroundColor Yellow
        Write-Host "   Removing old webhook and creating new one..." -ForegroundColor Yellow
        
        Remove-AzAutomationWebhook `
            -ResourceGroupName $config.ResourceGroupName `
            -AutomationAccountName $config.AutomationAccountName `
            -Name $webhookName `
            -Force | Out-Null
    }
    
    # Create new webhook
    $webhook = New-AzAutomationWebhook `
        -ResourceGroupName $config.ResourceGroupName `
        -AutomationAccountName $config.AutomationAccountName `
        -Name $webhookName `
        -RunbookName $config.RunbookName `
        -IsEnabled $true `
        -ExpiryTime $expiryDate `
        -Force
    
    Write-Host "   ✅ Webhook created successfully" -ForegroundColor Green
    
    # Display webhook URL (only available at creation time)
    $webhookUrl = $webhook.WebhookURI
    
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "✅ WEBHOOK CREATED!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Cyan
    
    Write-Host @"

Webhook Details:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Name:        $webhookName
Runbook:     $($config.RunbookName)
Expires:     $($expiryDate.ToString("yyyy-MM-dd"))
Status:      Enabled

⚠️  IMPORTANT - Webhook URL (shown only once):
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"@ -ForegroundColor White

    Write-Host $webhookUrl -ForegroundColor Yellow
    
    Write-Host @"

⚠️  SAVE THIS URL NOW! It cannot be retrieved later.

Expected Input JSON (from Power Automate):
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{
  "SiteUrl": "https://tenant.sharepoint.com/sites/sitename",
  "AuthenticationContextName": "Sensitive Information - Guest terms of Use",
  "RequestorEmail": "user@domain.com",
  "RequestId": "123"
}

"@ -ForegroundColor White
    
    # Save webhook URL to config (careful - this is sensitive)
    $config | Add-Member -NotePropertyName "WebhookUrl" -NotePropertyValue $webhookUrl -Force
    $config | Add-Member -NotePropertyName "WebhookName" -NotePropertyName $webhookName -Force
    $config | Add-Member -NotePropertyName "WebhookExpiry" -NotePropertyValue $expiryDate.ToString("yyyy-MM-dd") -Force
    $config | ConvertTo-Json | Out-File -FilePath $ConfigPath -Encoding UTF8
    
    # Also save to a separate secure file
    $webhookUrl | Out-File -FilePath ".\webhook-url-SECURE.txt" -Encoding UTF8
    Write-Host "⚠️  Webhook URL also saved to: webhook-url-SECURE.txt" -ForegroundColor Yellow
    Write-Host "   Keep this file secure - it contains authentication credentials!" -ForegroundColor Red
    
} catch {
    Write-Host "   ❌ Failed to create webhook" -ForegroundColor Red
    Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host @"

Next Steps:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1. ✅ Copy the webhook URL above

2. ⏭️  Run: .\4-Test-Webhook.ps1
   This will test the webhook with sample data

3. ⏭️  Create Power Automate Flow:
   - Use "HTTP" action
   - Method: POST
   - URI: <paste webhook URL>
   - Body: See expected JSON format above

4. ⏭️  Or import pre-built flow: PowerAutomate-Flow-Template.json

"@ -ForegroundColor White
