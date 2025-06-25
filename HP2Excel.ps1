<#
.SYNOPSIS
HP2Excel parses the configuration from a HP/Aruba  device into a Excel file.
.DESCRIPTION
The HP2Excel reads a HP/Aruba config file and pulls out the configuration into excel.
.PARAMETER HPConfig
[REQUIRED] This is the path to the HP/Aruba config/credential file
.PARAMETER SkipFilter 
[OPTIONAL] Set this value to $TRUE for not using Excel Filters.
.\HP2Excel.ps1 -HPConfig "c:\temp\config.conf"
    Parses a HP config file and places the Excel file in the same folder where the config was found.
.\HP2Excel.ps1 -HPConfig "c:\temp\config.conf" -SkipFilter:$true
    Parses a HP config file and places the Excel file in the same folder where the config was found.
    No filters will be auto applied.
.NOTES
Author: Xander Angenent (@XaAng70)
Last Modified: 20250624
#Uses Estimated completion time from http://mylifeismymessage.net/1672/
#Uses Posh-SSH https://github.com/darkoperator/Posh-SSH if reading directly from the firewall
#Uses Function that converts any Excel column number to A1 format https://gallery.technet.microsoft.com/office/Powershell-function-that-88f9f690
#>Param
(
    [Parameter(Mandatory = $true)]
    $HPConfig,
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
    $InitRule | Add-Member -type NoteProperty -name name -Value ""  
    $InitRule | Add-Member -type NoteProperty -name speed-duplex -Value "Auto" 
    $InitRule | Add-Member -type NoteProperty -name Interface -Value ""    

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
    $InitRule | Add-Member -type NoteProperty -name no_untagged -Value ""
    $InitRule | Add-Member -type NoteProperty -name tagged -Value ""
    $InitRule | Add-Member -type NoteProperty -name untagged -Value ""
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
if (!(Test-Path $HPConfig)) {
    Write-Output "File $HPConfig not found. Aborting script."
    exit 1
}
$loadedConfig = Get-Content $HPConfig
$Counter=0
$workingFolder = Split-Path $HPConfig;
$fileName = Split-Path $HPConfig -Leaf;
$fileName = (Get-Item $HPConfig).Basename
$ExcelFullFilePad = "$workingFolder\$fileName"
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$workbook = $excel.Workbooks.Add()
$FirstSheet = $workbook.Worksheets.Item(1) 
$FirstSheet.Cells.Item(1,1)= 'HP Configuration'
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
    if ($InterfaceSwitch -or $VlanConfig) {
        switch($ConfigLineArray[0]) {
            "exit" {
            if ($VlanConfig) {
                $VlanList.Add($Interface) | Out-Null
                $VlanConfig = $False
            }
            else {
                $InterfaceList.Add($Interface) | Out-Null
                $InterfaceSwitch=$false
            }
            }
            "ip" {
                if ($VlanConfig) {
                    switch($ConfigLineArray[1]) {
                        "address" {
                            if ($Interface.IPAddress -eq "") {
                                $Value = GetSubnetCIDR $ConfigLineArray[2] $ConfigLineArray[3] 
                            }
                            else {
                                $Value = $Interface.IPAddress + "," + (GetSubnetCIDR $ConfigLineArray[2] $ConfigLineArray[3]) 
                            }
                            $Interface | Add-Member -MemberType NoteProperty -Name IPAddress -Value $Value -force
                        }
                        "helper-address" {
                            if ($Interface.IPhelper -eq "") {
                                $Value = $ConfigLineArray[2]
                            }
                            else {
                                $Value = $Interface.IPhelper + "," + $ConfigLineArray[2]
                            }
                            $Interface | Add-Member -MemberType NoteProperty -Name IPhelper -Value $Value -force
                        }
                    }
                }
            }
            "name" {
                $Name = $configline.split('"')[1]
                $Interface | Add-Member -MemberType NoteProperty -Name $ConfigLineArray[0] -Value $Name -force
            }
            "no" {
                if ($ConfigLineArray[1] -eq "untagged") {
                    $Interface | Add-Member -MemberType NoteProperty -Name "no_untagged" -Value $ConfigLineArray[2] -force
                }
            }
            default {
                $Value = '"' + $ConfigLineArray[1] + '"'
                $Interface | Add-Member -MemberType NoteProperty -Name $ConfigLineArray[0] -Value $Value -force
            }
        }
        
    }
    else {
        switch($ConfigLineArray[0]) {
            ";" { #Version information for HP devices
                if ($ConfigLineArray[2] -eq "Configuration") {
                    $SwitchType = $ConfigLineArray[1]
                    $SwitchFirmware = $ConfigLineArray[7]
                }
            }
            "hostname" {
                $Hostname = $configline.split('"')[1]
            }
            "interface" {
                $Interface = InitInterface
                $Interface | Add-Member -MemberType NoteProperty -Name "Interface" -Value $ConfigLineArray[1] -force
                $InterfaceSwitch=$true
            }
            "ip" {
                switch ($ConfigLineArray[1]) {
                    "default-gateway" {
                        $Route = New-Object System.Object;
                        $Route | Add-Member -type NoteProperty -name Network -Value "0.0.0.0/0"
                        $Route | Add-Member -type NoteProperty -name Gateway -Value $ConfigLineArray[2]
                        $RouterTable.Add($Route) | Out-Null
                    }                    
                    "domain" {
                        $DomainName = $ConfigLineArray[3]
                    }
                    "name-server" {
                        $NameServer = $ConfigLineArray[2]
                        $Counter = 2
                        do {
                            $Counter++
                            $NameServer = $NameServer + " " + $ConfigLineArray[$Counter]                        
                        } While ($Counter -le $ConfigLineArray.Count)
                    }
                    "route" {
                        $Value = GetSubnetCIDR $ConfigLineArray[2] $ConfigLineArray[3]
                        $Route = New-Object System.Object;
                        $Route | Add-Member -type NoteProperty -name Network -Value $Value
                        $Route | Add-Member -type NoteProperty -name Gateway -Value $ConfigLineArray[4]
                        $RouterTable.Add($Route) | Out-Null
                    }                    
                }
            }
            "spanning-tree" {
                if ($ConfigLineArray.Count -gt 2) {
                    $SpanningTree = InitSpanningTree
                    $SpanningTree | Add-Member -MemberType NoteProperty -Name "Interface" -Value $ConfigLineArray[1] -force
                    if ($ConfigLineArray[2] -eq "priority") {
                        $SpanningTree | Add-Member -MemberType NoteProperty -Name "priority" -Value $ConfigLineArray[3] -force
                    }
                    else {
                        $SpanningTree | Add-Member -MemberType NoteProperty -Name "mode" -Value $ConfigLineArray[2] -force
                    }
                    $SpanningTreeList.Add($SpanningTree) | Out-Null
                }
            } 
            "trunk" {
                $Trunk = New-Object System.Object;
                $Trunk | Add-Member -MemberType NoteProperty -Name "Interface" -Value $ConfigLineArray[2] -force
                $Trunk | Add-Member -MemberType NoteProperty -Name "mode" -Value $ConfigLineArray[3] -force
                $Trunk | Add-Member -MemberType NoteProperty -Name "ports" -Value $ConfigLineArray[1] -force
                $TrunkList.Add($Trunk) | Out-Null
            }
            "vlan" {
                $Interface = InitVlan
                $Interface | Add-Member -MemberType NoteProperty -Name "Interface" -Value $ConfigLineArray[1] -force
                $VlanConfig = $true
            }        
            default {
            }
        }
    }
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