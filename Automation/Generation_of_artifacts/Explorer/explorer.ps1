#Konsole nicht anzeigen wenn nicht in ISE ausgefuehrt
    if(!($PSISE))
    {
            Add-Type -Name win -MemberDefinition '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);' -Namespace native
            [native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle,0)
    }
    
function Copy-fmItemUsingExplorer{
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

    #Set Show Full Path in Window Title true if not allready -> Put outside of the Funktion????
        $showfullpath = 0
        $showfullpath = (Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState").FullPath
        if($showfullpath -eq 0)
        { 
            Set-ItemProperty -Type DWord -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState" -Name "FullPath" -value "1"
            #Restart Explorer
            Get-Process explorer | Stop-Process
            Start-Sleep -Seconds 5        
        }

    #Create Destination Folder if not exists
    if(!(Test-Path $DestinationFolderPath) -and $CreateDestination)
    {
        New-Item -Path $DestinationFolderPath -ItemType Directory | Out-Null
    }
    if((Test-Path -Path $SourceFilePath) -and (Test-Path -Path $DestinationFolderPath))
    {
        #Top Foldername to identifiy Windows Explorer WindowTitle
            #If Show FullPath in Window Title is not set to true
                #$SourceExplorerWindowTitle = Split-Path -Path (Split-Path -Path $SourceFilePath) -Leaf
                #$DestinationExplorerWindowTitle = Split-Path -Path $DestinationFolderPath -Leaf
            #If Show FullPath in Window Title is set to true
                $SourceExplorerWindowTitle = Split-Path -Path $SourceFilePath
                $pathlength = 0
                $pathlength = $DestinationFolderPath.Length
                if($pathlength -gt 0 -and $pathlength -le 3)
                    { $DestinationExplorerWindowTitle = $DestinationFolderPath}
                else
                    { $DestinationExplorerWindowTitle = Split-Path -Path $DestinationFolderPath }
        #Open Explorer select file, to generate artifacts, for example Thumbnails
            #Parameters to launch explorer.exe with
            $Params = @()
                $Params += "/select,"
                $Params += "$($SourceFilePath)"
            Start-Process explorer.exe $Params #Start-Process explorer.exe -ArgumentList '/select, ""C:\Users\hanshack\Pictures\Product.bmp""'
            Start-Sleep -Milliseconds 2000
        #Copy File with Explorer (COM) - not via Copy-Item, maybe this generates more User like artifacts...
            $objShell = New-Object -ComObject 'Shell.Application'    
            $objFolder = $objShell.NameSpace((Get-Item $DestinationFolderPath).FullName)
            $objFolder.CopyHere((Get-Item $SourceFilePath).FullName,$CopyFlags) #Blocks until finished
            Start-Sleep -Milliseconds 1000
        #Close Explorer Window
            Get-Process -Name explorer | Where-Object {$_.MainWindowTitle -eq $SourceExplorerWindowTitle} | Stop-Process

        #Open Explorer select file, to generate artifacts, for example Thumbnails
            #Check if Path is Volume Root
                $volumelabel = ""
                $pathlength = 0
                $pathlength = $DestinationFolderPath.Length
                if($pathlength -gt 0 -and $pathlength -le 3)
                {
                    $volumedriveletter = $DestinationFolderPath[0]
                    $volumelabel = (Get-Volume $volumedriveletter).FileSystemLabel
                }
            #Parameters to launch explorer.exe with
            $Params = @()
                $Params += "/select,"
                $Params += "$($DestinationFolderPath)$(Split-Path -Path $SourceFilePath -Leaf)"
            Write-Host $Params
            Start-Sleep -Milliseconds 500
            Start-Process explorer.exe $Params #Start-Process explorer.exe -ArgumentList '/select, ""C:\Users\hanshack\Pictures\Product.bmp""'
           Start-Sleep -Milliseconds 2000

        #Focus Explorer Window
            #Check if DestinationFolderPath is Root of Volume
                #if($volumelabel -ne "")
                #{ $DestinationExplorerWindowTitle = "$($volumelabel) ($($DestinationFolderPath[0]):)" } # e.g.: "DATA (E:)"
            
            #$wshell = New-Object -ComObject wscript.shell
            #$wshell.AppActivate($DestinationExplorerWindowTitle) | Out-Null
                #$wshell = New-Object -ComObject wscript.shell
                #$wshell.AppActivate($DestinationExplorerWindowTitle) | Out-Null
                #Start-Sleep -Milliseconds 100
            [System.Windows.Forms.SendKeys]::SendWait(“%{TAB}”)
            Start-Sleep -Milliseconds 2000

            #Start Photos with Enter to generate Recent items entry (C:\Users\USERNAME\AppData\Roaming\Microsoft\Windows\Recent)
            if($SourceFilePath -eq "C:\Users\hanshack\Pictures\Motor.JPG")
            {
                [System.Windows.Forms.SendKeys]::SendWait(“{ENTER}”)
                Start-Sleep -Milliseconds 10000
                Get-Process -Name Microsoft.Photos | Stop-Process  # FileName is not in WindowsTitle!
            }

        #Close Explorer Window
            Get-Process -Name explorer | Where-Object {$_.MainWindowTitle -eq $DestinationExplorerWindowTitle} | Stop-Process

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

Start-Sleep -Seconds 5

#Identifiy Plugged in USB-Stick (only one plugged in suppoerted)
    $USBDriveLetter = (Get-WmiObject win32_volume | Where-Object {$_.DriveLetter -eq (Get-WmiObject -Query "Select * From Win32_LogicalDisk"`
    | Where-Object { $_.driveType -eq 2 }).DeviceID}).DriveLetter #Output for example: "E:"
    $DestinationFolderPath = "$($USBDriveLetter)\"

################
    $SourceFilePath = "Z:\Company_files\Product_Prices.docx"
    Copy-fmItemUsingExplorer -source $SourceFilePath -destination $DestinationFolderPath -CopyFlags 80

################
    $SourceFilePath = "Z:\Company_files\Passwords.xlsx"
    Copy-fmItemUsingExplorer -source $SourceFilePath -destination $DestinationFolderPath -CopyFlags 80

################
    $SourceFilePath = "C:\Users\hanshack\Documents\customer-list.xlsx"
    Copy-fmItemUsingExplorer -source $SourceFilePath -destination $DestinationFolderPath -CopyFlags 80

################
    $SourceFilePath = "C:\Users\hanshack\Documents\Electric_motors.pdf"
    Copy-fmItemUsingExplorer -source $SourceFilePath -destination $DestinationFolderPath -CopyFlags 80

################
    $SourceFilePath = "C:\Users\hanshack\Pictures\Motor.JPG"
    Copy-fmItemUsingExplorer -source $SourceFilePath -destination $DestinationFolderPath -CopyFlags 80
################

#Open Excel
    explorer.exe "$($USBDriveLetter)\customer-list.xlsx"
        Start-Sleep -Milliseconds 10000
        Get-Process -Name EXCEL | Stop-Process

#Open Photo 
    #Does not generate Recent items entry (C:\Users\USERNAME\AppData\Roaming\Microsoft\Windows\Recent)
    #explorer.exe "$($USBDriveLetter)\Motor.JPG"
        #Start-Sleep -Milliseconds 10000
        #Get-Process -Name Microsoft.Photos | Stop-Process

Eject USB Drive
    $Eject =  New-Object -ComObject Shell.Application
    $Eject.NameSpace(17).ParseName($USBDriveLetter).InvokeVerb("Eject")
