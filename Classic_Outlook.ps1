# PowerShell script to set UseNewOutlook registry value to 0, restart Outlook
## Created by XGoodwil 09102024

# Define the registry path and value
$registryPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Preferences"
$valueName = "UseNewOutlook"

# Define the name of the Outlook process
$outlookProcessName = "OUTLOOK"

# Set the registry value to 0
Set-ItemProperty -Path $registryPath -Name $valueName -Value 0

# Function to close Outlook if it's running
function Stop-Outlook {
    $processes = Get-Process -Name $outlookProcessName -ErrorAction SilentlyContinue
    if ($processes) {
        Write-Output "Stopping Outlook..."
        Stop-Process -Name $outlookProcessName -Force
        Start-Sleep -Seconds 5 # Wait for 5 seconds to ensure Outlook is closed
    } else {
        Write-Output "Outlook is not running."
    }
}

# Function to start Outlook
function Start-Outlook {
    Write-Output "Starting Outlook..."
    Start-Process "OUTLOOK.EXE"
}

# Execute the functions
Stop-Outlook
Start-Outlook

Write-Output "Registry value set and Outlook restarted."