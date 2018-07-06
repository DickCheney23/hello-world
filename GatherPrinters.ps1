$strFilter = "(objectCategory=printQueue)"

$objDomain = New-Object System.DirectoryServices.DirectoryEntry

$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.SearchRoot = $objDomain
$objSearcher.PageSize = 1000
$objSearcher.Filter = $strFilter

$colResults = $objSearcher.FindAll()

$objPrinters = @()
foreach ($objResult in $colResults)
 {$objItem = $objResult.Properties;
 $objptr = New-Object system.object
 $objptr | Add-Member -type noteproperty -name Name -value $objItem.printername
 $objptr | Add-Member -type noteproperty -name Location -value $objItem.location
 $objptr | Add-Member -type noteproperty -name Driver -value $objItem.drivername
 $objptr | Add-Member -type noteproperty -name Server -value $objItem.servername
 $objptr | Add-Member -type noteproperty -name Description -value $objItem.description
 $objPrinters += $objptr
 }

#$objPrinters

ForEach ($Row in $objPrinters){
	echo $Row
	echo Next...
	}

echo Ver3
<#
$Printers = [pscustomobject]@{
	objPrinters = (@($objPrinters) -join ',')
	}

$Printers


$OutArray | Select-Object Name,@{Expression={$_.objPrinters.Name -join ';'}} | Export-Csv test.csv
echo Out...
$OutArray

$NumColsToExport = 4

#Create array to hold the data to be sent to the CSV file
$Printers=@()
$pNames=@("Name", "Driver", "Server", "Description")
ForEach ($Row in $objPrinters){
    $obj = New-Object PSObject
    For ($i=0;$i -lt $NumColsToExport; $i++){
              $obj | Add-Member -MemberType NoteProperty -Name $pNames[$i] -Value $Row[$i]
      echo $i
      echo $Row[$i]
      }
   $Printers+=$obj
   $obj=$Null
}

#$Printers

#Export the data to a CSV file
$Printers | export-csv Printers.csv -NTI

$Printers | ft name, driver, server, description

#Export the data to a CSV file
$Printers | export-csv Printers.csv -NTI
$Printers | ft name, driver, server, description
#>
