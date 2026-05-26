# C:\temp\brother-mfc-j6730dw\gdi
# Brother MFC-J6730DW - Programing Printer reinstall script
# Installs Brother driver, removes bad WSD queue if present, creates LPR port, and adds printer system-wide.

$PrinterName = "Programing Printer"
$DriverName  = "Brother MFC-J6730DW Printer"
$PrinterIP   = "192.168.1.123"
$PortName    = "IP_192.168.1.123"
$LprQueue    = "BINARY_P1"
$DriverInf   = "C:\temp\brother-mfc-j6730dw\gdi\BRPRI16A.INF"

Write-Host "=== Brother Printer Install Started ==="

# Confirm driver INF exists
if (-not (Test-Path $DriverInf)) {
    Write-Host "[ERROR] Driver INF not found: $DriverInf"
    exit 1
}

# Remove existing printer queue if present
$ExistingPrinter = Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue
if ($ExistingPrinter) {
    Write-Host "Removing existing printer queue: $PrinterName"
    Remove-Printer -Name $PrinterName
}

# Add driver package to Windows Driver Store
Write-Host "Adding Brother driver package..."
pnputil /add-driver $DriverInf /install

# Find published OEM INF for Brother BRPRI16A driver
$BrotherDriver = Get-WindowsDriver -Online | Where-Object {
    $_.OriginalFileName -like "*BRPRI16A.INF*" -or $_.ProviderName -eq "Brother"
} | Sort-Object Date -Descending | Select-Object -First 1

if (-not $BrotherDriver) {
    Write-Host "[ERROR] Brother driver package was not found in Windows Driver Store."
    exit 1
}

$PublishedInf = "C:\Windows\INF\$($BrotherDriver.Driver)"
Write-Host "Using published INF: $PublishedInf"

# Register printer driver
Write-Host "Registering printer driver: $DriverName"
Add-PrinterDriver -Name $DriverName -InfPath $PublishedInf

# Confirm printer driver exists
$DriverCheck = Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue
if (-not $DriverCheck) {
    Write-Host "[ERROR] Printer driver did not register correctly: $DriverName"
    exit 1
}

# Create LPR printer port if missing
if (-not (Get-PrinterPort -Name $PortName -ErrorAction SilentlyContinue)) {
    Write-Host "Creating printer port: $PortName"
    Add-PrinterPort -Name $PortName `
        -LprHostAddress $PrinterIP `
        -LprQueueName $LprQueue
} else {
    Write-Host "Printer port already exists: $PortName"
}

# Add printer
Write-Host "Adding printer queue: $PrinterName"
Add-Printer -Name $PrinterName `
    -DriverName $DriverName `
    -PortName $PortName

# Confirm final result
Write-Host "`n=== Final Printer Configuration ==="
Get-Printer -Name $PrinterName |
    Select-Object Name, DriverName, PortName, PrinterStatus |
    Format-Table -AutoSize

Write-Host "=== Brother Printer Install Completed ==="
