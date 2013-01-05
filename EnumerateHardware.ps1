# PowerShell Script that uses WMI to enumerate hardware characteristics about machines (servers) on the network.
# Jackie Singh, 2011
#
# Expected format for Servers.txt is one hostname or IP per line in plaintext.
#

$Excel = New-Object -ComObject Excel.Application
$Excel.visible = $True
$Excel = $Excel.Workbooks.Add()

$i = 1
$intRow = 1
$Sheet = $Excel.Worksheets.Item($i++)

         $Sheet.Cells.Item($intRow,1).Font.Bold = $True
         $Sheet.Cells.Item($intRow,2).Font.Bold = $True
         $Sheet.Cells.Item($intRow,3).Font.Bold = $True
         $Sheet.Cells.Item($intRow,4).Font.Bold = $True
         $Sheet.Cells.Item($intRow,5).Font.Bold = $True
		 $Sheet.Cells.Item($intRow,6).Font.Bold = $True
		 $Sheet.Cells.Item($intRow,7).Font.Bold = $True
		 $Sheet.Cells.Item($intRow,8).Font.Bold = $True
		 $Sheet.Cells.Item($intRow,9).Font.Bold = $True
		 $Sheet.Cells.Item($intRow,10).Font.Bold = $True
		 $Sheet.Cells.Item($intRow,11).Font.Bold = $True
		 $Sheet.Cells.Item($intRow,12).Font.Bold = $True
		 
         $Sheet.Cells.Item($intRow,1) = "Server Name"
         $Sheet.Cells.Item($intRow,2) = "Service Tag"
		 $Sheet.Cells.Item($intRow,3) = "Vendor"
		 $Sheet.Cells.Item($intRow,4) = "Model"
		 $Sheet.Cells.Item($intRow,5) = "Processor"
		 $Sheet.Cells.Item($intRow,6) = "Processor Desc"
		 $Sheet.Cells.Item($intRow,7) = "IP Address"
		 $Sheet.Cells.Item($intRow,8) = "MAC Address"
		 $Sheet.Cells.Item($intRow,9) = "OS Name"
		 $Sheet.Cells.Item($intRow,10) = "OS Version"
		 $Sheet.Cells.Item($intRow,11) = "Processor Total"
		 $Sheet.Cells.Item($intRow,12) = "Physical Memory"
		 
		 
		              for ($col = 1; $col –le 12; $col++)
             {
                  $Sheet.Cells.Item($intRow,$col).Font.Bold = $True
                  $Sheet.Cells.Item($intRow,$col).Interior.ColorIndex = 48
                  $Sheet.Cells.Item($intRow,$col).Font.ColorIndex = 34
             }

# Place your list of machines to gather data from in Servers.txt, one on each line
foreach ($server in get-content "C:\Documents and Settings\Jacqueline.Singh\Desktop\Servers.txt")
{
		$intRow++
		$Sheet.Cells.Item($intRow,1) = $server

             $ComputerSystemProduct = Get-WmiObject `
             -ComputerName $server -Class Win32_ComputerSystemProduct | Sort-Object IdentifyingNumber,Name,Vendor 

			 foreach ($objItem in $ComputerSystemProduct){
                $CSPNumber = $objItem.IdentifyingNumber
                $CSPName = $objItem.Name
				$CSPVendor = $objItem.Vendor
                $Sheet.Cells.Item($intRow, 2) = $CSPNumber
				$Sheet.Cells.Item($intRow, 3) = $CSPVendor
				$Sheet.Cells.Item($intRow, 4) = $CSPName
				}	
			
			 $Win32Processor = Get-WmiObject `
             -ComputerName $server -Class Win32_Processor | Sort-Object Name,Description 
			 
			 foreach ($objItem in $Win32Processor){
                $Win32ProcName = ($objItem.Name)
                $Win32ProcDesc = $objItem.Description
 				$Sheet.Cells.Item($intRow, 5) = $Win32ProcName
                $Sheet.Cells.Item($intRow, 6) = $Win32ProcDesc
				}
				
			 $NetworkAdapterConfiguration = Get-WmiObject `
             -ComputerName $server -query "select * from win32_networkadapterconfiguration where IPEnabled='True'"
			 
			 foreach ($objItem in $NetworkAdapterConfiguration){
                $IPAddress = ($objItem.IPAddress)
				$NetAdapterServiceName = ($objItem.ServiceName)
 				$Sheet.Cells.Item($intRow, 7) = $IPAddress
				}
				
			 $NetworkAdapter = Get-WmiObject `
             -ComputerName $server -query "select * from win32_networkadapter where ServiceName='$NetAdapterServiceName'"
			 
			 foreach ($objItem in $NetworkAdapter){
                $MACAddress = ($objItem.MacAddress)
 				$Sheet.Cells.Item($intRow, 8) = $MACAddress
				}
			
			 $OperatingSystemDetails = (cmd /c systeminfo /s $server | `
			 findstr /B /C:"OS Name" /C:"OS Version" /C:"Processor(s)" /C:"Total Physical Memory")			 
				
			 foreach ($objItem in $OperatingSystemDetails){
				$Sheet.Cells.Item($intRow, 9) = $OperatingSystemDetails[0]
 				$Sheet.Cells.Item($intRow, 10) = $OperatingSystemDetails[1]
				$Sheet.Cells.Item($intRow, 11) = $OperatingSystemDetails[2]
				$Sheet.Cells.Item($intRow, 12) = $OperatingSystemDetails[3]
				}
				
        $Sheet.UsedRange.EntireColumn.AutoFit()
}
