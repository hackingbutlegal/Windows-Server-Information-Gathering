# PowerShell Script that uses WMI to collect software information about machines (servers) on the network.
# Jackie Singh, 2011
#
# Expected format for Servers.txt is one hostname or IP per line in plaintext.
#

# Create a new Excel object using COM
$Excel = New-Object -ComObject Excel.Application
$Excel.visible = $True
$Excel.SheetsInNewWorkbook = @(get-content "C:\Servers.txt").count

# Counter variable for rows
$i = 1

# Read thru the contents of the Servers.txt file
foreach ($server in get-content "C:\Servers.txt")
{
    $Excel = $Excel.Workbooks.Add()
    $Sheet = $Excel.Worksheets.Item($i++)
    $Sheet.Name = $server
	$ping = ping

    $intRow = 1

       if ($ping = "ping") {

         #Create column headers
         $Sheet.Cells.Item($intRow,1) = "NAME:"
         $Sheet.Cells.Item($intRow,2) = $server.ToUpper()
         $Sheet.Cells.Item($intRow,1).Font.Bold = $True
         $Sheet.Cells.Item($intRow,2).Font.Bold = $True

         $intRow++

         $Sheet.Cells.Item($intRow,1) = "APPLICATION"
         $Sheet.Cells.Item($intRow,2) = "VERSION"

             #Format the column headers
             for ($col = 1; $col -le 2; $col++)
             {
                  $Sheet.Cells.Item($intRow,$col).Font.Bold = $True
                  $Sheet.Cells.Item($intRow,$col).Interior.ColorIndex = 48
                  $Sheet.Cells.Item($intRow,$col).Font.ColorIndex = 34
             }

             $intRow++

             $software = Get-WmiObject `
             -ComputerName $server -Class Win32_Product | Sort-Object Name 

             #Formatting using Excel

             foreach ($objItem in $software){
                $Sheet.Cells.Item($intRow, 1) = $objItem.Name
                $Sheet.Cells.Item($intRow, 2) = $objItem.Version

                   $intRow ++
             }

        $Sheet.UsedRange.EntireColumn.AutoFit()
    }
}

Clear
