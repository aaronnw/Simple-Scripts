$printers = Get-WmiObject -Class Win32_Printer -ComputerName citrixprint
$ipregex = "^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$"
#$Output= $PSScriptRoot + "\printers.csv"
$Output = "C:\printers.csv"
$Output
foreach($printer in $printers){
      $ip = $printer.Portname
      if($ip -match $ipregex){
        $SNMP = New-Object -ComObject olePrn.OleSNMP
        $SNMP.Open($ip, "public", 2, 300);
        $printertype = $snmp.Get(".1.3.6.1.2.1.25.3.2.1.3.1")
        if($printer.Caption.contains("P75") -OR $printertype.Contains("Color")){
            $blacktonercapacity = $snmp.get("43.11.1.1.8.1.1")
            $blackcurrentvolume = $snmp.get("43.11.1.1.9.1.1")
            [int]$blackpercent = ($blackcurrentvolume / $blacktonercapacity) * 100 
            $cyantonervolume = $snmp.get("43.11.1.1.8.1.2")
            $cyancurrentvolume = $snmp.get("43.11.1.1.9.1.2")
            [int]$cyanpercent = ($cyancurrentvolume / $cyantonercapacity) * 100
            $magentatonervolume = $snmp.get("43.11.1.1.8.1.3")
            $magentacurrentvolume = $snmp.get("43.11.1.1.9.1.3")
            [int]$magentapercent = ($magentacurrentvolume / $magentatonercapacity) * 100
            $yellowtonervolume = $snmp.get("43.11.1.1.8.1.4")
            $yellowcurrentvolume = $snmp.get("43.11.1.1.9.1.4")
            [int]$yellowpercent = ($yellowcurrentvolume / $yellowtonercapacity) * 100
            $printer.DetectedErrorState
            $obj = [pscustomobject]@{
                "State" = $printer.DetectedErrorState
                "Description" = $printer.Caption;
                "Location" = $printer.Location;
                "IP" = $printer.portname;
                "Toner" = $tonername;
                "Capacity" = $blacktonercapacity;
                "Level" = $blackcurrentvolume;
                "Remaining" = $blackpercent;
                "Cyan" = $cyanpercent
                "Magenta" = $magentapercent
                "Yellow" = $yellowpercent
            }
            $obj | Export-CSV -Append -Path $Output -NoTypeInformation -Force

        }else{
            $tonername = $snmp.get("43.11.1.1.6.1.1")
            $currentvolume = $snmp.get("43.11.1.1.9.1.1")
            $maxvolume = $snmp.get("43.11.1.1.8.1.1")
            [int]$percent = ($currentvolume/$maxvolume)*100
            $obj = [pscustomobject]@{
                "State" = $printer.DetectedErrorState
                "Type" = $printertype
                "Description" = $printer.Caption;
                "Location" = $printer.Location;
                "IP" = $printer.portname;
                "Toner" = $tonername;
                "Capacity" = $maxvolume;
                "Level" = $currentvolume;
                "Remaining" = $percent;
            }
            $obj | Export-CSV -Append -Path $Output -NoTypeInformation
        }
        $printertype = "unknown"
        $tonername ="unknown"
        $currentvolume = "unknown"
        $maxvolume = "unknown"
        $percent = -1
      }
}