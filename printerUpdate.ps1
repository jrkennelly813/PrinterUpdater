# The intended use of this script occurs after a physical printer survey has been completed
# The imported data sheet should be populated with 2 columns from the survey
# ----> Column 1: new (This is the new names for the printers following name convention)
# ----> Column 2: old (This is the printers from currently residing on the serverthat intend to change)
# Will need to be run as admin w/ proper credentials to reach servers (Especially if ran locally)

# Set the file path to the 2-column CSV containing new/old printer names 
$filePath = "C:\Users\*USERNAME*\Documents\*FILENAME*.csv"

# Import printer data
$printer_list = Import-Csv  $filePath

# Optional - This is used if you are attempting 
# to run the script locally and not on the target print server
$server = read-Host "Enter the print server name: "

# Array to hold the successfully updated printers
$updatedPrinters = @()

# Iterate through the printer csv data
foreach ($printer in $printer_list) {
    # set a variable for each "new" printer for validation
    $printerExists = Get-Printer -ComputerName -Name $printer.new -ErrorAction SilentlyContinue

    # check to see if "new" printer is already on the server
    # If it does not already exist, continue creating
    if (-not $printerExists) {
        # rename existing old printer name to the new printer name
        Get-Printer -ComputerName $server -Name $printer.old | Rename-Printer -NewName $printer.new 
        # update existing printer object with new share name 
        Set-Printer -ComputerName $server -Name $printer.new -ShareName $printer.new
    }
    # set vairable to hold the new port name (FQDN)
    $portName = "{0}.pima.edu" -f $printer.new
    # set variable for port existence validation
    $portExists = Get-PrinterPort -ComputerName $server -Name $portName -PrinterName $printer.new -ErrorAction SilentlyContinue
    
    # Check if port already exists on the server
    if (-not $portExists) {
        # if port does not already exist, create it
        Add-PrinterPort -ComputerName $server -Name $portName -PrinterName $printer.new  
    }

    # set variable for updated printer
    $newPrinter = Get-Printer -ComputerName $server -Name $printer.new
    
    # validate the printer is updated after setting changes
    if ($newPrinter) {
        # add printer to updated list
        $updatedPrinters += $newPrinter
        Write-Host "{0} was added to the {0} print server successfully!" -ForegroundColor Green
    }
    else {
        # error if updated printer is not found
        Write-Host "The printer was not updated..." -ForegroundColor Red
    }
}

# Iterate through successfully updated printers and export info to a csv
foreach ($printer in $updatedPrinters) {
    [PSCustomObject] @{
        Name = $printer.name
        ComputerName = $printer.ComputerName
        ShareName = $printe.ShareName
        DriverName = $printer.DriverName
        PortName = $printer.PortName
        } | Export-Csv "C:\Users\*USERNAME*\Documents\*FILENAME*.csv" -notype -Append
}

