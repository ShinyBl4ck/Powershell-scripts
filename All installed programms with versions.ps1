# Retrieve computer information
$computer = Get-ComputerInfo

# Set file paths
$directoryPath = "C:\_Install"
$csvFile = Join-Path -Path $directoryPath -ChildPath ($computer.CsName +"-Programs.csv")

# Check if the directory exists, and create it if not
if (-not (Test-Path -Path $directoryPath)) {
    New-Item -Path $directoryPath -ItemType Directory
}

# Retrieve 32-bit and 64-bit programs and merge them into a single list
$program32 = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -ne $null }
$program64 = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -ne $null }
$allPrograms = $program32 + $program64 | Sort-Object DisplayName | Select-Object DisplayName, DisplayVersion

# Export results to a CSV file
$allPrograms | Export-Csv -LiteralPath $csvFile -Encoding UTF8 -UseCulture -NoTypeInformation

# Open the CSV file
Start-Process -FilePath $csvFile
