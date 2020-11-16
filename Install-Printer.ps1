Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework
[System.Windows.Forms.Application]::EnableVisualStyles()

$site = Get-ItemPropertyValue -Name Site-Name -Path "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine"
$servers = Import-Csv -Path "\\mainNetworkShare\servers.csv"
$languages = Import-Csv -Path "\\mainNetworkShare\languages.csv"
$iconPath = "\\mainNetworkShare\Install-Printer.ico"
$global:installedPrinters = Get-Printer
$global:oldPrinters = $installedPrinters.Name -match "oldPrintersModelsREGEX"

$language = Get-Culture
Switch ( $language.Name ) {

    "pl-PL" { $lng = "PL" }
    default { $lng = "EN" }

}

#Checks if machine is connected to VPN. Condition may differ in different companies.
<#
If ( $site -match "VPN" ) {

    [System.Windows.MessageBox]::Show($languages.$lng[7],$languages.$lng[8], "OK" ,"Error")
    exit
    
}
#>

$printersPath = $servers.$site[0] + "\printers.csv"
$printers = Import-Csv -Path $printersPath

$printingServer = $servers.$site[1]



$WindowX                                         = 350
$ButtonY                                         = 15
$groupHeight                                     = 45
$groupDrawingPointX                              = 10
$groupDrawingPointY                              = 10
$Spacer                                          = 10

$WindowY                                         = $printers.Count * ( $groupHeight + $Spacer ) + 10

If ( $WindowY -gt 800 ) {

    $WindowY = 800
    $EnableScroll = $true
    $WindowX += 10


}

$groupWidth                                      = $WindowX - $Spacer * 2
$buttonWidth                                     = ( $groupWidth - $Spacer * 4 ) / 3
$buttonHeight                                    = $groupHeight - $Spacer * 2
$installButtonX                                  = $Spacer
$configButtonX                                   = $Spacer * 2 + $buttonWidth
$removeButtonX                                   = $Spacer * 3  + $buttonWidth * 2

$groupDrawingPointSpacer                         = $groupHeight + $Spacer

$printerGroup                                    = @()

$installPrinterWindow                            = New-Object system.Windows.Forms.Form
$installPrinterWindow.ClientSize                 = "$WindowX,$WindowY"
$installPrinterWindow.text                       = $languages.$lng[0]
$installPrinterWindow.BackColor                  = "#ffffff"
$installPrinterWindow.TopMost                    = $false
$installPrinterWindow.icon						 = $iconPath
$installPrinterWindow.FormBorderStyle			 = "FixedSingle"
$installPrinterWindow.MaximizeBox                = $false

If ( $EnableScroll ) {

    $installPrinterWindow.AutoScroll             = $true

}

If ( $oldPrinters -ne $null ) {

    $WindowY                                      += 65
    $groupDrawingPointY                           += 65
    $installPrinterWindow.ClientSize              = "$WindowX,$WindowY"
    $removeOldPrintersButton                      = New-Object system.Windows.Forms.Button
    $removeOldPrintersButton.Text                 = $languages.$lng[9] 
    $removeOldPrintersButton.Width                = $WindowX - 2 * $Spacer
    $removeOldPrintersButton.Height               = $buttonHeight * 2
    $removeOldPrintersButton.BackColor            = "#FF0000"
    $removeOldPrintersButton.Font                 = 'Microsoft Sans Serif,20'
    $removeOldPrintersButton.Location             = New-Object System.Drawing.Point($groupDrawingPointX ,$ButtonY)
    $installPrinterWindow.controls.AddRange(@($removeOldPrintersButton))
    $removeOldPrintersButton.Add_Click( { removeOldPrinters } )
}

$printerCount = 0

ForEach ( $printerGroup in $printers) {
    
    $printerID = $printers[$printerCount].ID

    New-Variable -Name "printerGroup_$printerID" -Force -Value (New-Object System.Windows.Forms.GroupBox)
    $thisGroup               = Get-Variable -ValueOnly -Include "printerGroup_$printerID"
    $thisGroup.height        = $groupHeight
    $thisGroup.width         = $groupWidth
    $thisGroup.text          = $printers[$printerCount].DisplayName + "  -  " + "[" + $printers[$printerCount].Name + "]"
    $thisGroup.location      = New-Object System.Drawing.Point($groupDrawingPointX,$groupDrawingPointY)
    $installPrinterWindow.controls.Add($thisGroup)

    $groupDrawingPointY += $groupDrawingPointSpacer

    $printerCount++
}

$printerCount = 0
ForEach ( $button in $printers) {

    $printerName                               = $printers[$printerCount].Name
    $printerID                                 = $printers[$printerCount].ID

    New-Variable -Name "addButton_$printerID" -Force -Value (New-Object system.Windows.Forms.Button)
    $thisAddButton                             = Get-Variable -ValueOnly -Include "addButton_$printerID"

    New-Variable -Name "configButton_$printerID" -Force -Value (New-Object system.Windows.Forms.Button)
    $thisConfigButton                          = Get-Variable -ValueOnly -Include "configButton_$printerID"

    New-Variable -Name "removeButton_$printerID" -Force -Value (New-Object system.Windows.Forms.Button)
    $thisRemoveButton                          = Get-Variable -ValueOnly -Include "removeButton_$printerID"

    If ( Get-Printer -Name "*$printerName*" ) {

        $thisAddButton.text                    = $languages.$lng[3]
	    $thisAddButton.Font                    = 'Microsoft Sans Serif,8'
	    $thisAddButton.Enabled				   = $false
        $thisConfigButton.Enabled              = $true
        $thisRemoveButton.Enabled              = $true

    } else {

 	    $thisAddButton.text                    = $languages.$lng[1]
	    $thisAddButton.Font                    = 'Microsoft Sans Serif,10'
        $thisAddButton.Enabled                 = $true
        $thisConfigButton.Enabled              = $false
        $thisRemoveButton.Enabled              = $false

    }

    $thisAddButton.Width                       = $buttonWidth
    $thisAddButton.Height                      = $buttonHeight
    $thisAddButton.Location                    = New-Object System.Drawing.Point($installButtonX,$ButtonY)

    $thisConfigButton.Text                     = $languages.$lng[4]
    $thisConfigButton.Width                    = $buttonWidth
    $thisConfigButton.Height                   = $buttonHeight
    $thisConfigButton.Location                 = New-Object System.Drawing.Point($configButtonX,$ButtonY)

    $thisRemoveButton.Text                     = $languages.$lng[5]
    $thisRemoveButton.Width                    = $buttonWidth
    $thisRemoveButton.Height                   = $buttonHeight
    $thisRemoveButton.Location                 = New-Object System.Drawing.Point($removeButtonX,$ButtonY)

    $thisGroup                                 = Get-Variable -ValueOnly -Include "printerGroup_$printerID"
    $thisGroup.controls.AddRange(@($thisAddButton, $thisConfigButton, $thisRemoveButton))

    $thisAddButton.Add_Click( { installPrinter $printerName $printerID }.GetNewClosure() )
    $thisConfigButton.Add_Click( { configurePrinter $printerName $printerID }.GetNewClosure() )
    $thisRemoveButton.Add_Click( { removePrinter $printerName $printerID }.GetNewClosure() )

    $printerCount++

}



#LOGIC

function global:installPrinter ( [string]$printerName, [int]$printerID ) {

	$connectionName                            = "\\" + $printingServer + "\" + $printerName

    $thisAddButton                             = Get-Variable -ValueOnly -Include "addButton_$printerID"
    $thisConfigButton                          = Get-Variable -ValueOnly -Include "configButton_$printerID"
    $thisRemoveButton                          = Get-Variable -ValueOnly -Include "removeButton_$printerID"

	$thisAddButton.Enabled                     = $false
    $thisAddButton.text                        = $languages.$lng[2]
    
    Add-Printer -AsJob -ConnectionName $connectionName
    Start-Sleep 3
	
    $thisAddButton.text                        = $languages.$lng[3]
	$thisAddButton.Font                        = 'Microsoft Sans Serif,8'
	$thisConfigButton.Enabled                  = $true
    $thisRemoveButton.Enabled                  = $true
			
}
	

function global:configurePrinter ( [string]$printerName ) {

	$connectionName                            = "\\" + $printingServer + "\" + $printerName
	rundll32 printui.dll,PrintUIEntry /e /n $connectionName

}
	

function global:removePrinter ( [string]$printerName, [int]$printerID ) {

	$connectionName                             = "\\" + $printingServer + "\" + $printerName
    $thisAddButton                              = Get-Variable -ValueOnly -Include "addButton_$printerID"
    $thisConfigButton                           = Get-Variable -ValueOnly -Include "configButton_$printerID"
    $thisRemoveButton                           = Get-Variable -ValueOnly -Include "removeButton_$printerID"

	Remove-Printer -AsJob -Name $connectionName
    
    $thisAddButton.Enabled                      = $true
    $thisAddButton.Text                         = $languages.$lng[1]
    
	$thisConfigButton.Enabled                   = $false
    $thisRemoveButton.Enabled                   = $false

}

function global:removeOldPrinters () {

    $removeOldPrintersButton.Enabled            = $false
    $removeOldPrintersButton.BackColor          = "#FFA500"
    $removeOldPrintersButton.Font               = 'Microsoft Sans Serif,10'

    ForEach ( $oldPrinter in $oldPrinters ) {
        $removeOldPrintersButton.Text           = $languages.$lng[10] + " " + $oldPrinter
        Remove-Printer -AsJob -Name $oldPrinter 
        Start-Sleep 1
    }

    $removeOldPrintersButton.BackColor          = "#00FF00"
    $removeOldPrintersButton.Text               = $languages.$lng[11]

}


[void]$installPrinterWindow.ShowDialog()