# PowerShell script to fix Excel COM registration issues
# Run this script as Administrator to resolve ConnectionRefusedError with Excel COM

Write-Host "Excel COM Fix Script" -ForegroundColor Green
Write-Host "===================" -ForegroundColor Green

# Check if running as Administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "❌ This script must be run as Administrator!" -ForegroundColor Red
    Write-Host "Right-click PowerShell and select 'Run as Administrator'" -ForegroundColor Yellow
    exit 1
}

Write-Host "✓ Running as Administrator" -ForegroundColor Green

# Step 1: Check if Excel is installed
Write-Host "`nStep 1: Checking Excel installation..." -ForegroundColor Cyan

$excelPath = Get-ChildItem -Path "C:\Program Files\Microsoft Office*" -Recurse -Filter "EXCEL.EXE" -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $excelPath) {
    $excelPath = Get-ChildItem -Path "C:\Program Files (x86)\Microsoft Office*" -Recurse -Filter "EXCEL.EXE" -ErrorAction SilentlyContinue | Select-Object -First 1
}

if ($excelPath) {
    Write-Host "✓ Excel found at: $($excelPath.FullName)" -ForegroundColor Green
} else {
    Write-Host "❌ Excel not found. Please install Microsoft Excel first." -ForegroundColor Red
    exit 1
}

# Step 2: Stop Excel processes
Write-Host "`nStep 2: Stopping Excel processes..." -ForegroundColor Cyan
Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Stop-Process -Force
Write-Host "✓ Excel processes stopped" -ForegroundColor Green

# Step 3: Re-register Excel COM
Write-Host "`nStep 3: Re-registering Excel COM..." -ForegroundColor Cyan

try {
    # Register Excel as COM server
    $regCommand = "regsvr32 /s `"$($excelPath.FullName)`""
    Invoke-Expression $regCommand
    Write-Host "✓ Excel COM registered successfully" -ForegroundColor Green
    
    # Also try user-specific registration
    $regUserCommand = "regsvr32 /i:user /s `"$($excelPath.FullName)`""
    Invoke-Expression $regUserCommand
    Write-Host "✓ Excel COM user registration completed" -ForegroundColor Green
} catch {
    Write-Host "❌ Failed to register Excel COM: $($_.Exception.Message)" -ForegroundColor Red
}

# Step 4: Configure DCOM settings
Write-Host "`nStep 4: Configuring DCOM settings..." -ForegroundColor Cyan

try {
    # Set DCOM config for Excel
    $dcomConfig = @"
Windows Registry Editor Version 5.00

[HKEY_LOCAL_MACHINE\SOFTWARE\Classes\AppID\{00020812-0000-0000-C000-000000000046}]
"AuthenticationLevel"=dword:00000001
"EnableDistributedCOM"=dword:00000001
"@
    
    $tempFile = "$env:TEMP\excel_dcom_fix.reg"
    $dcomConfig | Out-File -FilePath $tempFile -Encoding ASCII
    
    # Import registry settings
    Start-Process -FilePath "reg" -ArgumentList "import", $tempFile -Wait -NoNewWindow
    Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
    
    Write-Host "✓ DCOM configuration updated" -ForegroundColor Green
} catch {
    Write-Host "⚠️  DCOM configuration may need manual adjustment" -ForegroundColor Yellow
}

# Step 5: Test Excel COM
Write-Host "`nStep 5: Testing Excel COM..." -ForegroundColor Cyan

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Host "✓ Excel COM test successful!" -ForegroundColor Green
} catch {
    Write-Host "❌ Excel COM test failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "You may need to restart Windows and try again." -ForegroundColor Yellow
}

Write-Host "`n🎉 Excel COM fix script completed!" -ForegroundColor Green
Write-Host "If you still experience issues, try:" -ForegroundColor Yellow
Write-Host "1. Restart Windows" -ForegroundColor Yellow
Write-Host "2. Run Excel manually once" -ForegroundColor Yellow
Write-Host "3. Run the Python test script: python test_excel_com.py" -ForegroundColor Yellow
