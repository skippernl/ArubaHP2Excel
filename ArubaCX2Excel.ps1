<#
.SYNOPSIS
ArubaCX2Excel parses the configuration from a HP/Aruba  device into a Excel file.
.DESCRIPTION
The ArubaCX2Excel reads a ArubaCX config file and pulls out the configuration into excel.
.PARAMETER ArubaCXconfig
[REQUIRED] This is the path to the HP/Aruba config/credential file
.PARAMETER SkipFilter 
[OPTIONAL] Set this value to $TRUE for not using Excel Filters.
.\ArubaCX2Excel.ps1 -HPConfig "c:\temp\config.conf"
    Parses a Aruba CX config file and places the Excel file in the same folder where the config was found.
.\ArubaCX2Excel.ps1 -HPConfig "c:\temp\config.conf" -SkipFilter:$true
    Parses a Aruba CX config file and places the Excel file in the same folder where the config was found.
    No filters will be auto applied.
.NOTES
Author: Xander Angenent (@XaAng70)
Last Modified: 20250625
#Uses Estimated completion time from http://mylifeismymessage.net/1672/
#Uses Posh-SSH https://github.com/darkoperator/Posh-SSH if reading directly from the firewall
#Uses Function that converts any Excel column number to A1 format https://gallery.technet.microsoft.com/office/Powershell-function-that-88f9f690
#>Param
(
    [Parameter(Mandatory = $true)]
    $ArubaCXconfig,
    [switch]$SkipFilter = $false
)
Function CleanSheetName ($CSName) {
    $CSName = $CSName.Replace("-","_")
    $CSName = $CSName.Replace(" ","_")
    $CSName = $CSName.Replace("\","_")
    $CSName = $CSName.Replace("/","_")
    $CSName = $CSName.Replace("[]","_")
    $CSName = $CSName.Replace("]","_")
    $CSName = $CSName.Replace("*","_")
    $CSName = $CSName.Replace("?","_")
    if ($CSName.Length -gt 32) {
        Write-output "Sheetname ($CSName) cannot be longer that 32 character shorting name to fit."
        $CSName = $CSName.Substring(0,31)
    }    

    return $CSName
}
Function InitInterface {
    $InitRule = New-Object System.Object;
    $InitRule | Add-Member -type NoteProperty -name access-vlan -Value "" 
    $InitRule | Add-Member -type NoteProperty -name address -Value ""  
    $InitRule | Add-Member -type NoteProperty -name Allowed-vlan -Value "" 
    $InitRule | Add-Member -type NoteProperty -name description -Value ""
    $InitRule | Add-Member -type NoteProperty -name Interface -Value ""    
    $InitRule | Add-Member -type NoteProperty -name lag -Value ""  
    $InitRule | Add-Member -type NoteProperty -name name -Value ""  
    $InitRule | Add-Member -type NoteProperty -name Native-Vlan -Value ""  
    $InitRule | Add-Member -type NoteProperty -name shutdown -Value "" 
    $InitRule | Add-Member -type NoteProperty -name speed-duplex -Value "Auto"    

    return $InitRule
}
Function InitSpanningTree {
    $InitRule = New-Object System.Object;
    $InitRule | Add-Member -type NoteProperty -name Interface -Value ""  
    $InitRule | Add-Member -type NoteProperty -name priority -Value ""
    $InitRule | Add-Member -type NoteProperty -name mode -Value ""

    return $InitRule
}
Function InitVlan {
    $InitRule = New-Object System.Object;
    $InitRule | Add-Member -type NoteProperty -name name -Value ""  
    $InitRule | Add-Member -type NoteProperty -name IPAddress -Value ""
    $InitRule | Add-Member -type NoteProperty -name Vlan -Value ""    

    return $InitRule
}
Function ChangeFontExcelCell ($ChangeFontExcelCellSheet, $ChangeFontExcelCellRow, $ChangeFontExcelCellColumn) {
    $ChangeFontExcelCellSheet.Cells.Item($ChangeFontExcelCellRow, $ChangeFontExcelCellColumn).HorizontalAlignment = -4108
    $ChangeFontExcelCellSheet.Cells.Item($ChangeFontExcelCellRow, $ChangeFontExcelCellColumn).Font.Size = 18
    $ChangeFontExcelCellSheet.Cells.Item($ChangeFontExcelCellRow, $ChangeFontExcelCellColumn).Font.Bold=$True
    $ChangeFontExcelCellSheet.Cells.Item($ChangeFontExcelCellRow, $ChangeFontExcelCellColumn).Font.Name = "Cambria"
    $ChangeFontExcelCellSheet.Cells.Item($ChangeFontExcelCellRow, $ChangeFontExcelCellColumn).Font.ThemeFont = 1
    $ChangeFontExcelCellSheet.Cells.Item($ChangeFontExcelCellRow, $ChangeFontExcelCellColumn).Font.ThemeColor = 4
    $ChangeFontExcelCellSheet.Cells.Item($ChangeFontExcelCellRow, $ChangeFontExcelCellColumn).Font.ColorIndex = 55
    $ChangeFontExcelCellSheet.Cells.Item($ChangeFontExcelCellRow, $ChangeFontExcelCellColumn).Font.Color = 8210719
}
Function CreateExcelSheet ($SheetName, $SheetArray) {
    if ($SheetArray) {
        $row = 1
        $Sheet = $workbook.Worksheets.Add()
        $Sheet.Name = $SheetName
        $Column=1
        $excel.cells.item($row,$Column) = $SheetName 
        ChangeFontExcelCell $Sheet $row $Column  
        $row++
        $NoteProperties = SkipEmptyNoteProperties $SheetArray
        foreach ($Noteproperty in $NoteProperties) {
            $excel.cells.item($row,$Column) = $Noteproperty.Name
            $Column++
        }
        $StartRow = $Row
        $row++
        foreach ($rule in $SheetArray) {
            $Column=1
            foreach ($Noteproperty in $NoteProperties) {
                $PropertyString = [string]$NoteProperty.Name
                $Value = $Rule.$PropertyString
                $excel.cells.item($row,$Column) = $Value
                $Column++
            }    
            $row++
        }    
        #No need to filer if there is only one row.
        if (!($SkipFilter) -and ($SheetArray.Count -gt 1)) {
            $RowCount =  $Sheet.UsedRange.Rows.Count
            $ColumCount =  $Sheet.UsedRange.Columns.Count
            $ColumExcel = Convert-NumberToA1 $ColumCount
            $Sheet.Range("A$($StartRow):$($ColumExcel)$($RowCount)").AutoFilter() | Out-Null
        }
        #Use autoFit to expand the colums
        $UsedRange = $Sheet.usedRange                  
        $UsedRange.EntireColumn.AutoFit() | Out-Null
    }
}
#Function from https://gallery.technet.microsoft.com/office/Powershell-function-that-88f9f690
Function Convert-NumberToA1 { 
    <# 
    .SYNOPSIS 
    This converts any integer into A1 format. 
    .DESCRIPTION 
    See synopsis. 
    .PARAMETER number 
    Any number between 1 and 2147483647 
    #> 

    Param([parameter(Mandatory=$true)] 
        [int]$number) 

    $a1Value = $null 
    While ($number -gt 0) { 
        $multiplier = [int][system.math]::Floor(($number / 26)) 
        $charNumber = $number - ($multiplier * 26) 
        If ($charNumber -eq 0) { $multiplier-- ; $charNumber = 26 } 
        $a1Value = [char]($charNumber + 64) + $a1Value 
        $number = $multiplier 
    } 
    Return $a1Value 
}
Function GetSubnetCIDR ([string]$Subnet,[IPAddress]$SubnetMask) {
    $binaryOctets = $SubnetMask.GetAddressBytes() | ForEach-Object { [Convert]::ToString($_, 2) }
    $SubnetCIDR = $Subnet + "/" + ($binaryOctets -join '').Trim('0').Length
    return $SubnetCIDR
}
#Function SkipEmptyNoteProperties ($SkipEmptyNotePropertiesArray)
#This function Loopt through all available noteproperties and checks if it is used.
#If it is not used the property will not be returned as it is not needed in the export.
Function SkipEmptyNoteProperties ($SkipEmptyNotePropertiesArray) {
    $ReturnNoteProperties = [System.Collections.ArrayList]@()
    $SkipNotePropertiesOrg = $SkipEmptyNotePropertiesArray | get-member -Type NoteProperty
    foreach ($SkipNotePropertieOrg in $SkipNotePropertiesOrg) {
        foreach ($SkipEmptyNotePropertiesMember in $SkipEmptyNotePropertiesArray) {
            $NotePropertyFound = $False
            $SkipNotePropertiePropertyString = [string]$SkipNotePropertieOrg.Name
            if ($SkipEmptyNotePropertiesMember.$SkipNotePropertiePropertyString) { 
                $NotePropertyFound = $True
                break;
            }
        }
        If ($NotePropertyFound) { $ReturnNoteProperties.Add($SkipNotePropertieOrg) | Out-Null  }
    }

    return $ReturnNoteProperties
}

$startTime = get-date 
$date = Get-Date -Format yyyyMMddHHmm
Clear-Host
Write-Output "Started script"
#Clear 5 additional lines for the progress bar
$I=0
DO {
    Write-output ""
    $I++
} While ($i -le 4)
If ($SkipFilter) {
    Write-Output "SkipFilter parmeter is set to True. Skipping filter function in Excel."
}
if (!(Test-Path $ArubaCXconfig)) {
    Write-Output "File $ArubaCXconfig not found. Aborting script."
    exit 1
}
$loadedConfig = Get-Content $ArubaCXconfig
$Counter=0
$workingFolder = Split-Path $ArubaCXconfig;
$fileName = Split-Path $ArubaCXconfig -Leaf;
$fileName = (Get-Item $ArubaCXconfig).Basename
$ExcelFullFilePad = "$workingFolder\$fileName"
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$workbook = $excel.Workbooks.Add()
$FirstSheet = $workbook.Worksheets.Item(1) 
$FirstSheet.Cells.Item(1,1)= 'Aruba CX configuration'
$MergeCells = $FirstSheet.Range("A1:G1")
$MergeCells.Select() | Out-Null
$MergeCells.MergeCells = $true
$FirstSheet.Cells(1, 1).HorizontalAlignment = -4108
$FirstSheet.Cells.Item(1,1).Font.Size = 18
$FirstSheet.Cells.Item(1,1).Font.Bold=$True
$FirstSheet.Cells.Item(1,1).Font.Name = "Cambria"
$FirstSheet.Cells.Item(1,1).Font.ThemeFont = 1
$FirstSheet.Cells.Item(1,1).Font.ThemeColor = 4
$FirstSheet.Cells.Item(1,1).Font.ColorIndex = 55
$FirstSheet.Cells.Item(1,1).Font.Color = 8210719
$DomainName = "Unknown"
$NameServer = "None"
$InterfaceSwitch=$False
$InterfaceList = [System.Collections.ArrayList]@()
$VlanConfig=$False
$VlanList = [System.Collections.ArrayList]@()
$MaxCounter=$loadedConfig.count
$RouterTable = [System.Collections.ArrayList]@()
$SpanningTreeList = [System.Collections.ArrayList]@()
$TrunkList = [System.Collections.ArrayList]@()
foreach ($Line in $loadedConfig) {
    $Proc = $Counter/$MaxCounter*100
    $elapsedTime = $(get-date) - $startTime 
    if ($Counter -eq 0) { $estimatedTotalSeconds = $MaxCounter/ 1 * $elapsedTime.TotalSecond }
    else { $estimatedTotalSeconds = $MaxCounter/ $counter * $elapsedTime.TotalSeconds }
    $estimatedTotalSecondsTS = New-TimeSpan -seconds $estimatedTotalSeconds
    $estimatedCompletionTime = $startTime + $estimatedTotalSecondsTS    
    Write-Progress -Activity "Parsing config file. Estimate completion time $estimatedCompletionTime" -PercentComplete ($Proc)
    $Counter++
    $Configline=$Line.Trim() -replace '\s+',' '
    $ConfigLineArray = $Configline.Split(" ")
    switch ($ConfigLineArray[0]) {
        "!Version" { #Version information for Aruba devices
            $SwitchType = $ConfigLineArray[1]
            $SwitchFirmware = $ConfigLineArray[2]
        }
        "description" {
            $Interface | Add-Member -MemberType NoteProperty -Name "description" -Value $ConfigLineArray[1] -force
        }        
        "hostname" {
            $Hostname = $ConfigLineArray[1]
        }
        "interface" {
            #Check if there is a Vlan that is not added to the list.
            if ($VlanConfig) { #we are already in a vlan config
                $VlanList.Add($Vlan) | Out-Null
                $VlanConfig = $false
            }
            switch ($ConfigLineArray[1]) {
                "lag" {
                    $InterfaceName = "lag-" + $ConfigLineArray[2]
                }
                "vlan" {
                    $InterfaceName = "vlan-" + $ConfigLineArray[2]
                }                
                default {
                    $InterfaceName = '"' + $ConfigLineArray[1] + '"'
                }
            }
            if ($InterfaceSwitch) { #we are already in an interface config
                $Interfacelist.Add($Interface) | Out-Null
                $Interface = InitInterface
                $Interface | Add-Member -MemberType NoteProperty -Name "Interface" -Value  $InterfaceName  -force
            }
            else {
                $Interface = InitInterface
                $Interface | Add-Member -MemberType NoteProperty -Name "Interface" -Value  $InterfaceName -force
                $InterfaceSwitch=$true
            }
        }
        "ip" {
            switch ($ConfigLineArray[1]) {
                "address" {
                    if ($InterfaceSwitch) { #we are already in an interface config
                        $Interface | Add-Member -MemberType NoteProperty -Name "address" -Value $ConfigLineArray[2] -force
                    }
                    else { #we are in a vlan config
                        $Vlan | Add-Member -MemberType NoteProperty -Name "IPAddress" -Value $ConfigLineArray[2] -force
                    }
                }
                "route" {
                        $Route = New-Object System.Object;
                        $Route | Add-Member -type NoteProperty -name Network -Value $ConfigLineArray[2]
                        $Route | Add-Member -type NoteProperty -name Gateway -Value $ConfigLineArray[3]
                        $RouterTable.Add($Route) | Out-Null
                }
            }
        }
        "lag" {
            $Interface | Add-Member -MemberType NoteProperty -Name "lag" -Value $ConfigLineArray[1] -force
        }
        "name" {
            if ($InterfaceSwitch) {
                $Interface | Add-Member -MemberType NoteProperty -Name "name" -Value $ConfigLineArray[1] -force
            }
            else {
                $Vlan | Add-Member -MemberType NoteProperty -Name "name" -Value $ConfigLineArray[1] -force
            }
        }
        "shutdown" {
            $Interface | Add-Member -MemberType NoteProperty -Name "shutdown" -Value "Down" -force
        }
        "vlan" {
            if ($VlanConfig) { #we are already in a vlan config
                $VlanList.Add($Vlan) | Out-Null
                $Vlan = InitVlan
                $Vlan | Add-Member -MemberType NoteProperty -Name "Vlan" -Value $ConfigLineArray[1] -force
            }
            else { #we are starting a new vlan config or are we in an interface config.
                if ($InterfaceSwitch) { #we are in an interface config
                    switch ($ConfigLineArray[1]) {
                        "access" {
                            $Interface | Add-Member -MemberType NoteProperty -Name "access-vlan" -Value $ConfigLineArray[2] -force
                        }
                        "trunk" {
                            switch ($ConfigLineArray[2]) {
                                "native" {
                                    $Interface | Add-Member -MemberType NoteProperty -Name "Native-Vlan" -Value $ConfigLineArray[3] -force
                                }
                                "allowed" {
                                    $Interface | Add-Member -MemberType NoteProperty -Name "Allowed-vlan" -Value $ConfigLineArray[3] -force
                                }
                            }
                        }
                    }
                }
                else { #we are in a new vlan config
                    $Vlan = InitVlan
                    $Vlan | Add-Member -MemberType NoteProperty -Name "Vlan" -Value $ConfigLineArray[1] -force
                    $VlanConfig = $true
                }
            }        
        }
    }
}
if ($InterfaceSwitch) { #add the last interface to the list.
    $Interfacelist.Add($Interface) | Out-Null
}
CreateExcelSheet "Interfaces" $Interfacelist 
CreateExcelSheet "RoutingTable" $RouterTable
CreateExcelSheet "Spanningtree" $SpanningTreeList
CreateExcelSheet "Trunk" $TrunkList
CreateExcelSheet "VLAN" $VlanList

#make sure that the first sheet that is opened by Excel is the global sheet.
$FirstSheet.Activate()
$FirstSheet.Cells.Item(2,1) = 'Excel Creation Date'
$FirstSheet.Cells.Item(2,2) = $Date
$FirstSheet.Cells.Item(2,2).numberformat = "00"
$FirstSheet.Cells.Item(3,1) = 'Switch Type'
$FirstSheet.Cells.Item(3,2) = $SwitchType
$FirstSheet.Cells.Item(4,1) = 'Switch Firmware'
$FirstSheet.Cells.Item(4,2) = $SwitchFirmware
#$FirstSheet.Cells.Item(4,2).numberformat = "00"
$FirstSheet.Cells.Item(5,1) = 'Hostname'
$FirstSheet.Cells.Item(5,2) = $Hostname
$FirstSheet.Cells.Item(8,1) = "Domainname"
$FirstSheet.Cells.Item(8,2) = $DomainName
$FirstSheet.Cells.Item(9,1) = "DNS Server(s)"
$FirstSheet.Cells.Item(9,2) = $NameServer
$FirstSheet.Name = $Hostname
$UsedRange = $FirstSheet.usedRange                  
$UsedRange.EntireColumn.AutoFit() | Out-Null
Write-Output "Writing Excelfile $ExcelFullFilePad.xls"
$workbook.SaveAs($ExcelFullFilePad)
$excel.Quit()
$elapsedTime = $(get-date) - $startTime
$Minutes = $elapsedTime.Minutes
$Seconds = $elapsedTime.Seconds
Write-Output "Script done in $Minutes Minute(s) and $Seconds Second(s)."