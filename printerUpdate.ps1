# The intended use of this script occurs after a physical printer survey has been completed
# The imported data sheet should be populated with 2 columns from the survey
# ----> Column 1: new (This is the new names for the printers following name convention)
# ----> Column 2: old (This is the printers from currently residing on the serverthat intend to change)
# Will need to be run as admin w/ proper credentials to reach servers (Especially if ran locally)

# Set the file path to the 2-column CSV containing new/old printer names 
# UPDATE FILEPATH BEFORE RUNNING
$filePath = "C:\Users\wadmin_jrkennelly\Documents\DV_EDU_PList.csv"

# Import printer data
$printer_list = Import-Csv  $filePath

# Optional - This is used if you are attempting 
# to run the script locally and not on the target print server
$prompt = read-Host "Enter the print server name: "
$server = "\\$($prompt)"

# Array to hold the successfully updated printers
$updatedPrinters = @()

# Iterate through the printer csv data
foreach ($printer in $printer_list) {
    # set a variable for each "new" printer for validation
    $printerExists = Get-Printer -ComputerName $server -Name $printer.new -ErrorAction SilentlyContinue
    # set vairable to hold the new port name (FQDN)
    $portName = "$($printer.new).pima.edu" 

    # check to see if "new" printer is already on the server
    # If it does not already exist, continue updating
    if (-not $printerExists) {
        # rename existing old printer name to the new printer name
        if (Get-Printer -ComputerName $server -Name $printer.old | Rename-Printer -NewName $printer.new) {
            Write-Host "$($printer.old) name changed to ---> $($printer.new)" -ForegroundColor Green
        }
        if (Set-Printer -ComputerName $server -Name $printer.new -ShareName $printer.new) {
            Write-Host "$($printer.old) share name changed to ---> $($printer.new)" -ForegroundColor Green
        }
        
        # add printer to updated list
        if ($updatedPrinters += $newPrinter) {
            Write-Host "$($printer.new) was added to the $($server) print server successfully!" -ForegroundColor Cyan
        }
    }
    else {
        Write-Host "<---------- Print Queue up to date! ---------->" -ForegroundColor Cyan
    }
} 
# Iterate through successfully updated printers and export info to a csv
foreach ($printer in $updatedPrinters) {
    if (Add-PrinterPort -ComputerName $server -Name $portName -PrinterHostAddress $portName -SNMP $true) {
        Write-Host "New port configured ---> $($portName)" -ForegroundColor Green
    }
    [PSCustomObject] @{
        Name = $printer.name
        ComputerName = $printer.ComputerName
        ShareName = $printer.ShareName
        DriverName = $printer.DriverName
        PortName = $printer.PortName
        # UPDATE FILEPATH BEFORE RUNNING
        } | Export-Csv "C:\Users\wadmin_jrkennelly\Documents\DV_EDU_PrintUpdate.csv" -notype -Append
}

