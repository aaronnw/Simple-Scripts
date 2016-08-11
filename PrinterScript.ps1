$printers = Get-WmiObject -Class Win32_Printer -ComputerName c201-115
$ipregex = "^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$"
#$Output= $PSScriptRoot + "\printers.csv"
#$Output = "C:\printers.csv"
#$Output
#####################################
#Create excel COM object
$excel = New-Object -ComObject excel.application
$row = 1;
$column = 1;

#Make Visible
$excel.Visible = $True

#Add a workbook
$workbook = $excel.Workbooks.Add()

#Connect to first worksheet to rename and make active
$serverInfoSheet = $workbook.Worksheets.Item(1)
$serverInfoSheet.Name = 'Printers'
$serverInfoSheet.Activate() | Out-Null

#Create a header for Disk Space Report; set each cell to Bold and add a background color
$serverInfoSheet.Cells.Item($row,$column)= 'Type'
$serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
$serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
$Column++
$serverInfoSheet.Cells.Item($row,$column)= 'Description'
$serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
$serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
$Column++
$serverInfoSheet.Cells.Item($row,$column)= 'Location'
$serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
$serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
$Column++
$serverInfoSheet.Cells.Item($row,$column)= 'IP'
$serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
$serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
$Column++
$serverInfoSheet.Cells.Item($row,$column)= 'Toner'
$serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
$serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
$Column++
$serverInfoSheet.Cells.Item($row,$column)= 'Remaining'
$serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
$serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
$Column++
$serverInfoSheet.Cells.Item($row,$column)= 'Status'
$serverInfoSheet.Cells.Item($row,$column).Interior.ColorIndex =48
$serverInfoSheet.Cells.Item($row,$column).Font.Bold=$True
$column = 1;
$row++

$iplist = New-Object System.Collections.ArrayList
$iplist.Add("192.168.1.1")|out-null
$locationList = New-Object System.Collections.ArrayList
##################################################
foreach($printer in $printers){
      [string]$ip = $printer.Portname
      #Check if IP is the right form and is not a duplicate
      if(($ip -match $ipregex) -and ($iplist -notcontains $ip)){
            $iplist.Add($ip) |out-null
      ##################################################
      #Get the status
            $statuscode = $printer.DetectedErrorState
            if($statuscode -eq 0){
                $status = ""
            }elseif($statuscode -eq 9){
                $status = "Offline"
            }elseif($statuscode -eq 5){
                $status = "Low Toner"
            }elseif($statuscode -eq 6){
                $status = "No Toner"
            }else{
                $status = "Other"
            }
        ############################################
        #Get the location
            
            $ipPiece = $ip.Split(".")
            if($ipPiece[2] -eq 5){
                $location = "Plant 2"
            }elseif($ipPiece[2] -eq 7){
                $location = "Derby West"
            }elseif($ipPiece[2] -eq 10){
                $location = "Tyler"
            }elseif($ipPiece[2] -eq 13){
                $location = "Pawnee"
            }elseif($ipPiece[2] -eq 16){
                $location = "Tallgrass"
            }elseif($ipPiece[2] -eq 19){
                $location = "OKP"
            }elseif($ipPiece[2] -eq 22){
                $location = "Oliver"
            }elseif($ipPiece[2] -eq 28){
                $location = "Lakeside"
            }elseif($ipPiece[2] -eq 31){
                $location = "Derby East"
            }elseif($ipPiece[2] -eq 37){
                $location = "Maple"
            }elseif($ipPiece[2] -eq 40){
                $location = "Crosspointe"
            }elseif($ipPiece[2] -eq 66){
                $location = "Junction City"
            }elseif($ipPiece[2] -eq 69){
                $location = "Lawrence East"
            }elseif($ipPiece[2] -eq 72){
                $location = "Lawrence West"
            }elseif($ipPiece[2] -eq 75){
                $location = "Manhattan"
            }elseif(($ipPiece[2] -eq 90) -or ($ipPiece[2] -eq 91) -or ($ipPiece[2] -eq 92) -or ($ipPiece[2] -eq 93) -or ($ipPiece[2] -eq 201) ){
                $location = "Triad"
            }else{
                $location = "Other"
            }
            if(!$locationList.Contains($location)){
                $locationList.Add($location)|out-null
            }
        ##########################################
        #create the object with snmp if possible
            try{
            $SNMP = New-Object -ComObject olePrn.OleSNMP
            $SNMP.Open($ip, "public", 2, 300)
            $printertype = $snmp.Get(".1.3.6.1.2.1.25.3.2.1.3.1")
            $tonername = $snmp.get("43.11.1.1.6.1.1")
            $currentvolume = $snmp.get("43.11.1.1.9.1.1")
            $maxvolume = $snmp.get("43.11.1.1.8.1.1")
            
            if(($currentvolume -ge 0) -and ($currentvolume -le $maxvolume)){
                [int]$percent = ($currentvolume/$maxvolume)*100
                $obj = [pscustomobject]@{
                    "Type" = $printertype
                    "Description" = $printer.Caption
                    "Location" = $location
                    "IP" = $ip
                    "Toner" = $tonername
                    "Remaining" = $percent
                    "Status" = $status
                }
            }else{
                $obj = [pscustomobject]@{
                    "Type" = $printertype
                    "Description" = $printer.Caption
                    "Location" = $location
                    "IP" = $ip
                    "Toner" = $tonername
                    "Remaining" = "unknown"
                    "Status" = $status
                    }
            }
        #If snmp throws an error just use the built-in stuff
        }catch{
            $obj = [pscustomobject]@{
                    "Type" = "Unknown"
                    "Description" = $printer.Caption
                    "Location" = $location
                    "IP" = $ip
                    "Toner" = "unknown"
                    "Remaining" = "unknown"
                    "Status" = $status
                    }

        }
        #Old csv export
        #$obj | Export-CSV -Append -Path $Output -NoTypeInformation
        ############################################
        #output to excel
        $serverInfoSheet.Cells.Item($row,$column)= $obj.Type
        $Column++
        $serverInfoSheet.Cells.Item($row,$column)= $obj.Description
        $Column++
        $serverInfoSheet.Cells.Item($row,$column)= $obj.Location
        $Column++
        $serverInfoSheet.Cells.Item($row,$column)= $obj.IP
        $Column++
        $serverInfoSheet.Cells.Item($row,$column)= $obj.Toner
        $Column++
        $serverInfoSheet.Cells.Item($row,$column)= $obj.Remaining
        $Column++
        $serverInfoSheet.Cells.Item($row,$column)= $obj.Status
    
        #Check to see if space is near empty and use appropriate background colors
        $range = $serverInfoSheet.Range(("A{0}"  -f $row),("G{0}"  -f $row))
        $range.Select() | Out-Null
    
        #Add colors for low ink and errors
        If ($obj.Remaining -lt 10 -OR ($statuscode -eq 5) -OR ($statuscode -eq 6)) {
            #Low ink
            $range.Interior.ColorIndex = 3
        } ElseIf ($obj.Remaining -lt 30) {
            #Warning threshold 
            $range.Interior.ColorIndex = 6
        }
    
        #Increment to next row and reset Column to 1
        $Column = 1
        $row++

########################################################
        $printertype = "unknown"
        $tonername ="unknown"
        $currentvolume = "unknown"
        $maxvolume = "unknown"
        $percent = -1
        $statuscode = -1
        $status = "unknown"
      }
}