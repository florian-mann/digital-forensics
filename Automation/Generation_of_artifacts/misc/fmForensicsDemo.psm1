function Start-fmForensicsDemo1{
# Opens Explorer with SourceFile selected, copies SourceFile to DestinationFolder, closes this Explorer Windows and reopens
# Explorer with Destination File selected and closes Explorer Windows again.
 
# Source: BackSlasher, Feb 19, 2013, https://blog.backslasher.net/copying-files-in-powershell-using-windows-explorer-ui.html
    # CopyHere Flags https://docs.microsoft.com/en-us/windows/win32/shell/folder-copyhere
        # "16" -> Yes to all
        # "80" -> Yes to all and Preserve undo information if possible
# Existing files are overwritten without question!

#Notes
    # Open Windows in PowerShell https://devblogs.microsoft.com/scripting/hey-scripting-guy-how-can-i-use-windows-powershell-to-get-a-list-of-all-the-open-windows-on-a-computer/
        #$a = New-Object -com "Shell.Application"; $b = $a.windows() | select-object LocationName; $b

#If Show FullPath in Window Title is not set to true -> localized Folders (Documents/Dokumente -> EN/DE) make problems closing windows by title


#$DestinationFolderPath = "C:\Users\fortrace\"
#Konsole nicht anzeigen wenn nicht in ISE ausgefuehrt
    if(!($PSISE))
    {
            Add-Type -Name win -MemberDefinition '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);' -Namespace native
            [native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle,0)
    }

[string]$tempip = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object {$_.IPAddress -like "192.168.103.*"}).IPAddress

#Identifiy Plugged in USB-Stick (only one plugged in suppoerted)
    $USBDriveLetter = (Get-WmiObject win32_volume | Where-Object {$_.DriveLetter -eq (Get-WmiObject -Query "Select * From Win32_LogicalDisk"`
    | Where-Object { $_.driveType -eq 2 }).DeviceID}).DriveLetter #Output for example: "E:"
    $DestinationFolderPath = "$($USBDriveLetter)\"
    #$USBDriveLetter = "E:\"

#Search for Searchstring via Windows Explorer in Thumbdrive
    Query-fmExplorerSearch -SearchString "customer" -SearchPath ”\\$($tempip)\Company_files\”
    Get-Process -Name explorer | Where-Object {$_.MainWindowTitle -eq ”customer - search-ms:displayname=Search%20Results%20in%20Company_files%20(%5C%5C$($tempip))”} | Stop-Process

################
    $SourceFilePath = "\\$($tempip)\Company_files\Customer\customer-list.xlsx"
    Copy-fmItemUsingExplorer -source $SourceFilePath -destination $DestinationFolderPath -CopyFlags 80
    
################
    $SourceFilePath = "\\$($tempip)\Company_files\Pictures\motor.JPG"
    Copy-fmItemUsingExplorer -source $SourceFilePath -destination $DestinationFolderPath -CopyFlags 80

################
    $SourceFilePath = "\\$($tempip)\Company_files\Product_Prices.docx"
#    Copy-fmItemUsingExplorer -source $SourceFilePath -destination $DestinationFolderPath -CopyFlags 80

################
    $SourceFilePath = "\\$($tempip)\Company_files\Passwords.xlsx"
#    Copy-fmItemUsingExplorer -source $SourceFilePath -destination $DestinationFolderPath -CopyFlags 80


################
    $SourceFilePath = "\\$($tempip)\Company_files\Pictures\Kikuzuki-blueprint.pdf"
    Copy-fmItemUsingExplorer -source $SourceFilePath -destination $DestinationFolderPath -CopyFlags 80

################


#Open Excel
    explorer.exe "$($USBDriveLetter)\customer-list.xlsx"
        Start-Sleep -Milliseconds 9000
        Get-Process -Name EXCEL | Stop-Process
        Start-Sleep -Milliseconds 2000

#Open Photo on USB Drive
    #Does not generate Recent items entry (C:\Users\USERNAME\AppData\Roaming\Microsoft\Windows\Recent)
    start "$($USBDriveLetter)\Motor.JPG"
        Start-Sleep -Milliseconds 10000
        Get-Process -Name Microsoft.Photos | Stop-Process

#Eject USB Drive
    Start-Sleep -Milliseconds 5000
    $Eject =  New-Object -ComObject Shell.Application
    $Eject.NameSpace(17).ParseName($USBDriveLetter).InvokeVerb("Eject")

}