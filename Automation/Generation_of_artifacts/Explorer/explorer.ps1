# Explorer COM examples (Search...)
# https://www.computerperformance.co.uk/powershell/com-shell/

#Files to interact with in Windows Explorer
    $SourceFilePath = "C:\Users\hanshack\Pictures\Product.jpg"
    #Identifiy Plugged in USB-Stick (only one plugged in suppoerted)
        $USBDriveLetter = (Get-WmiObject win32_volume | Where-Object {$_.DriveLetter -eq (Get-WmiObject -Query "Select * From Win32_LogicalDisk" | Where-Object { $_.driveType -eq 2 }).DeviceID}).DriveLetter #Output for example: "E:"
    $DestinationFolderPath = "$($USBDriveLetter)\test2\test3"

function Copy-fmItemUsingExplorer{
# Opens Explorer with SourceFile selected, copies SourceFile to DestinationFolder, closes this Explorer Windows and reopens Explorer with Destination File selected and closes Explorer Windows again.
 
# Source: BackSlasher, Feb 19, 2013, https://blog.backslasher.net/copying-files-in-powershell-using-windows-explorer-ui.html
    # CopyHere Flags https://docs.microsoft.com/en-us/windows/win32/shell/folder-copyhere
        # "16" -> Yes to all
        # "80" -> Yes to all and Preserve undo information if possible
# Existing files are overwritten without question?????
    param(
        [Parameter(Mandatory=$True)]
        [ValidateScript({Test-Path -Path $_})]
            [string]$SourceFilePath,

        [Parameter(Mandatory=$True)]
            [string]$DestinationFolderPath,

        [bool]$CreateDestination = $True,

        [ValidateSet("16", "80")]
            [int]$CopyFlags = 80
    )

    #load VB
        Add-Type -AssemblyName Microsoft.VisualBasic
        Add-Type -AssemblyName System.Windows.Forms

    #Create Destination Folder if not exists
    if(!(Test-Path $DestinationFolderPath) -and $CreateDestination)
    {
        New-Item -Path $DestinationFolderPath -ItemType Directory | Out-Null
    }
    if((Test-Path -Path $SourceFilePath) -and (Test-Path -Path $DestinationFolderPath))
    {
        #Top Foldername to identifiy Windows Explorer WindowTitle
            $SourceExplorerWindowsTitle = Split-Path -Path (Split-Path -Path $SourceFilePath) -Leaf
            $DestinationExplorerWindowsTitle = Split-Path -Path $DestinationFolderPath -Leaf

        #Open Explorer select file, to generate artifacts, for example Thumbnails
            #Parameters to launch explorer.exe with
            $Params = @()
                $Params += "/select,"
                $Params += "$($SourceFilePath)"
            Start-Process explorer.exe $Params #Start-Process explorer.exe -ArgumentList '/select, ""C:\Users\hanshack\Pictures\Product.bmp""'
            Start-Sleep -Milliseconds 3000
        #Focus Explorer Window
            $wshell = New-Object -ComObject wscript.shell
            $wshell.AppActivate($SourceExplorerWindowsTitle) | Out-Null
            Start-Sleep -Milliseconds 10000
        #Copy File with Explorer (COM) - not via Copy-Item, maybe this generates more User like artifacts...
            $objShell = New-Object -ComObject 'Shell.Application'    
            $objFolder = $objShell.NameSpace((Get-Item $DestinationFolderPath).FullName)
            $objFolder.CopyHere((Get-Item $SourceFilePath).FullName,$CopyFlags) #Blocks until finished
        #Close Explorer Window
            Get-Process -Name explorer | Where-Object {$_.MainWindowTitle -eq $SourceExplorerWindowsTitle} | Stop-Process

        #Open Explorer select file, to generate artifacts, for example Thumbnails
            #Parameters to launch explorer.exe with
            $Params = @()
                $Params += "/select,"
                $Params += "$($DestinationFolderPath)$(Split-Path -Path $SourceFilePath -Leaf)"
            Start-Process explorer.exe $Params #Start-Process explorer.exe -ArgumentList '/select, ""C:\Users\hanshack\Pictures\Product.bmp""'
            Start-Sleep -Milliseconds 3000
        #Focus Explorer Window
            $wshell = New-Object -ComObject wscript.shell
            $wshell.AppActivate($DestinationExplorerWindowsTitle) | Out-Null
            Start-Sleep -Milliseconds 10000
        #Close Explorer Window
            Get-Process -Name explorer | Where-Object {$_.MainWindowTitle -eq $DestinationExplorerWindowsTitle} | Stop-Process

        #validation and return values
        if(Test-Path -Path "$($DestinationFolderPath)\$(Split-Path -Path $SourceFilePath -Leaf)")
        { return $True }
        else
        { return $false }
    }
    else
    {
        #Return Error Source File or Destination Folder does not exist
        return $false
    }
}

Copy-fmItemUsingExplorer -source $SourceFilePath -destination $DestinationFolderPath -CopyFlags 80


#Eject USB Drive
    $Eject =  New-Object -ComObject Shell.Application
    $Eject.NameSpace(17).ParseName($USBDriveLetter).InvokeVerb("Eject")

