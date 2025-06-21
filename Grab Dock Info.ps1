<#
    File: Get Dock SN on Dock.ps1
    Original Date: 03/13/25
    Version: 1.0.0
    
    Purpose: Pull the serial number of connected docking HP USB-C G5 Docking Station and save for a report.

#>

 # Uncomment below for no debug messages
 $DEBUGPREFVAL = "SilentlyContinue"

 # Uncomment below to Debug
 #$DEBUGPREFVAL = "Continue"
 
 $DebugPreference = $DEBUGPREFVAL

function Get-DockSerialNumber {
    $DebugPreference = $DEBUGPREFVAL
    
    try{
        $className = "HP_DockAccessory"
        $namespace = "ROOT\HP\InstrumentedServices\v1"

        $roomNumber = Read-Host "Enter the Room Number"
        $tagNumber = Read-Host "Please enter the Tag Number"

        Write-Host "Getting WMI Object - Please Wait . . ."
        $wmiObject = Get-WmiObject -Class $className -Namespace $namespace | Select-Object SerialNumber,FirmwarePackageVersion,ID
        
        $record = @{
            RoomNumber = $roomNumber
            TagNumber= $tagNumber
            SerialNumber = $wmiObject.SerialNumber
            FirmwarePackageVersion = $wmiObject.FirmwarePackageVersion
            ID = $wmiObject.ID
        }
    }
    catch {
        Write-Warning "$ComputerName cannot find a docking station"
        Write-Debug $_               
        return (, $null)
    }

    Write-Debug "Returning Dock Information"
    return (, $record)
  
}

function Display-DockInfo
{    
    param(
        [array]$DataObject
    )
    $DebugPreference = $DEBUGPREFVAL

    If ($DataObject) 
    {
        Write-Host "Dock Information"
        $DataObject | Format-Table -Property RoomNumber,TagNumber,SerialNumber,FirmwarePackageVersion,ID
    }
    else
    {
        Write-Warning "No Dock Information Recorded"
    }
}

$statusDataPath = "C:\temp\DockSerialNumbers_$(get-date -f yyyy-MM-dd_HHmmss).csv"
$choices = [System.Management.Automation.Host.ChoiceDescription[]] @("&Yes","&No")
$Results=@()


try {
    Do {
        # Loop until user is NOT connected to a dock
        $choice = $Host.UI.PromptForChoice("Connected to Dock? ","",$choices,0)
        if ( $choice -eq 0 ) {
            # Get dock information and add to array
            $info = Get-DockSerialNumber
            $Results += New-Object psobject -Property $info
        }
    } While ( $choice -eq 0 )

    # Display Array
    Display-DockInfo -DataObject $Results

    # Is array being saved to file?
    $savecsv = $Host.UI.PromptForChoice("Save results to file? ","",$choices,0)
}
finally {
    if ( $Results -and ($savecsv -eq 0 -or $savecsv -eq $null)) {
        Write-Host "Creating CSV Report" 

        # Save array to file
        $Results | Select-Object RoomNumber,TagNumber,SerialNumber,FirmwarePackageVersion,ID | Export-Csv -notypeinformation -Path $statusDataPath

        Write-Host "Report saved at $statusDataPath"
        Pause
    }
    else
    {
        Write-Warning "No Dock Information to Save!"
        Pause
    }
}


