$pc = $env:COMPUTERNAME

# Network (correct active IP)
$netConfig = Get-NetIPConfiguration | Where-Object {$_.IPv4DefaultGateway -ne $null}

$ip  = $netConfig.IPv4Address.IPAddress
$mac = $netConfig.NetAdapter.MacAddress

# System Info
$os = Get-CimInstance Win32_OperatingSystem
$cs = Get-CimInstance Win32_ComputerSystem
$cpu = Get-CimInstance Win32_Processor
$bios = Get-CimInstance Win32_BIOS

# RAM
$ramTotal = (Get-CimInstance Win32_PhysicalMemory | Measure-Object Capacity -Sum).Sum / 1GB

# Disk
$disk = Get-CimInstance Win32_DiskDrive
$diskModel = ($disk.Model -join " | ")
$diskSize  = [math]::Round(($disk.Size | Measure-Object -Sum).Sum / 1GB,2)

# Device Type
$deviceType = if ($cs.PCSystemType -eq 2) {"Laptop"} else {"Desktop"}

# Office Version
$office = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* |
Where-Object {$_.DisplayName -like "*Office*"} |
Select-Object -First 1 -ExpandProperty DisplayName

# Antivirus
$antivirus = (Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntivirusProduct |
Select-Object -ExpandProperty displayName) -join " | "

# Windows Activation
$activation = (Get-CimInstance SoftwareLicensingProduct |
Where-Object {$_.PartialProductKey -and $_.LicenseStatus -eq 1}).Name

# Build Data (matching your Excel)
$data = [PSCustomObject]@{
    "No."                       = ""
    "Employee Name"             = ""
    "Department"                = ""
    "System Name"               = $pc
    "Login ID"                  = $env:USERNAME
    "Password"                  = ""
    "WiFi SSID"                 = ""
    "WiFi Password"             = ""
    "Device Type"               = $deviceType
    "MAC Address"               = $mac
    "IP Address"                = $ip
    "Laptop Company"            = "$($cs.Manufacturer) $($cs.Model)"
    "RAM"                       = [math]::Round($ramTotal,2)
    "Processor"                 = $cpu.Name
    "Harddisk/SSD"              = $diskModel
    "Storage size"              = $diskSize
    "Microsoft Office Version"  = $office
    "Antivirus Installed"       = $antivirus
    "windows activated"         = if ($activation) {"Yes"} else {"No"}
    "Remark"                    = ""
}

$file = "inventory.csv"

if (!(Test-Path $file)) {
    $data | Export-Csv -Path $file -NoTypeInformation
}
else {
    $data | Export-Csv -Path $file -NoTypeInformation -Append
}

Write-Host "✅ Inventory added to inventory.csv"