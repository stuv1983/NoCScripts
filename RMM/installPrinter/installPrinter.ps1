# ============================================================
# Remote Printer Driver + Printer Installation via PowerShell
# ============================================================

# --- CONFIGURATION ---
$DriverSourcePath = "C:\temp\SH_D33_PCL6_PS_2508a_English_64bit"
$PrinterIP        = "192.168.15.200"   # <-- Replace with actual printer IP
$PortName         = "IP_$PrinterIP"
$PrinterName      = "Sharp D33 PCL6"  # <-- Customize as needed
$DriverName       = "SHARP D33 PCL6"  # <-- Must match exactly what's in the .inf

# --- STEP 1: Find the .inf file ---
$InfFile = Get-ChildItem -Path $DriverSourcePath -Filter "*.inf" -Recurse | Select-Object -First 1
if (-not $InfFile) {
    Write-Error "No .inf file found in $DriverSourcePath"; exit 1
}
Write-Host "Using INF: $($InfFile.FullName)"

# --- STEP 2: Stage the driver into the Windows driver store ---
Write-Host "Staging driver..."
pnputil.exe /add-driver "$($InfFile.FullName)" /install
if ($LASTEXITCODE -ne 0) {
    Write-Error "pnputil failed to stage driver (exit $LASTEXITCODE)"; exit 1
}

# --- STEP 3: Add the printer driver via Print Spooler ---
Write-Host "Adding printer driver to spooler..."
Add-PrinterDriver -Name $DriverName -ErrorAction Stop

# --- STEP 4: Create a TCP/IP port for the printer ---
if (-not (Get-PrinterPort -Name $PortName -ErrorAction SilentlyContinue)) {
    Write-Host "Creating printer port $PortName..."
    Add-PrinterPort -Name $PortName -PrinterHostAddress $PrinterIP
} else {
    Write-Host "Port $PortName already exists, skipping."
}

# --- STEP 5: Install the printer ---
if (-not (Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue)) {
    Write-Host "Adding printer $PrinterName..."
    Add-Printer -Name $PrinterName -DriverName $DriverName -PortName $PortName
} else {
    Write-Host "Printer $PrinterName already exists, skipping."
}

Write-Host "`n✅ Done! Printer '$PrinterName' installed on port $PortName"