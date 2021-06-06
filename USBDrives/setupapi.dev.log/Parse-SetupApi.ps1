#requires -version 5.1
<# Parse-SetupApi

 Version 0.1, 06.06.2021, AUTHOR: Florian Mann
 
.SYNOPSYS
  Pase setupapi.dev.log for specific DeviceIDs

.NOTES

 Requirements: 
  - Windows 10
     
.SOURCES


.TODO
  - redo as PS-Module with parameters
  - identifiy relevant "* Driver Update]" lines
         
.CHANGELOG
  0.1
   - Inital release
#>

$SetupapiPath = "$($env:SystemRoot)\INF\setupapi.dev.log"
$DeviceSerialNumber = "1234567..." #DeviceID

if(Test-Path $SetupapiPath)
{
    $SetupapiContent = Get-Content $SetupapiPath
    [int]$SetupapiNumberofLines = $SetupapiContent.Count
    Write-Host "setupapi.dev.log found with $($SetupapiNumberofLines) lines..."
    Write-Host "Searching for $($DeviceSerialNumber) entrys. Please Wait..."

    #Arry for all Device Entrys
        $RelevantDeviceEntrys = @()
    #Array for Device lines
        $TempRelevantDeviceLines = @()
    #For current DeviceEntry to build Entrys Array
        class DeviceEntry {
            [string]$type
            [datetime]$time
            [System.Collections.ArrayList]$lines
        }
    #Detection of relevant Device Section Begin
        $RelevantEntryBegin = $false
        $regexEntryBegin =  "^>>>  \[(.)*$($DeviceSerialNumber)(.)*]$" #">>>  [ANYSTRING SERIALNUMBER ANYSTRING ]"
    #Detection of relevant Device Update Section Begin
        $RelevantEntryUpdateBegin = $false
        $regexEntryUpdateBegin =  "^>>>  \[(.)* Driver Updates]$"
    #Detection of relevant Device Section End
        $RelevantEntryEnd = $false
        $regexEntryEnd =  "^<<<  \[Exit status:(.)*]$"
    $i=1
    $linefortime = "0"
    foreach($line in $SetupapiContent)
    {
        #Find Begin
        if(!($RelevantEntryBegin) -and ($line -match $regexEntryBegin))
        { 
            $RelevantEntryBegin = $true
            $Type = (($line.Split("[")[1]).Split("-")[0]).Trim()
                #Begin types Format
                    #>>>  [Device Install (Hardware initiated)
                    #>>>  [Delete Device -
                    #>>>  [Device Uninstall (Device Manager)

            #Time on next Line
            $linefortime = $i + 1
        }
        #Find Begin Updates
        if(!($RelevantEntryBegin -and $RelevantEntryUpdateBegin) -and ($line -match $regexEntryUpdateBegin))
        { 
            $RelevantEntryUpdateBegin = $true
            $Type = (($line.Split("[")[1]).Split("]")[0]).Trim()
                #Begin types Format
                    #>>>  [Unstage Driver Updates]
                    #>>>  [Uninstall Driver Updates]
                    #>>>  [Install Driver Updates]
                             #sto:      USBSTOR\Disk&Ven_Kingston&Prod_DataTraveler_2.0&Rev_PMAP\00241D8CE459BFB0194935D4&0 -> GenDisk [disk_install.NT]
                    #>>>  [Stage Driver Updates]

            #Time on next Line
            $linefortime = $i + 1

        }
        #Find End
        elseif(($RelevantEntryBegin -or $RelevantEntryUpdateBegin) -and ($line -match $regexEntryEnd))
        { $RelevantEntryEnd = $true}

        #Find Line with Time
        if($linefortime -eq $i)
        {
            [DateTime]$Time = Get-Date $line.Substring(19,23)
            #Format
                #>>>  Section start 2021/06/05 22:47:55.279
        }

        if($RelevantEntryBegin -and !($RelevantEntryEnd))
        {
            #Add line to current Device Entry
            $TempRelevantDeviceLines += $line
        }
        elseif($RelevantEntryUpdateBegin -and !($RelevantEntryEnd))
        {
            #Search for relevant Update lines
            if($line -match $DeviceSerialNumber)
            {
                #Add line to current Device Entry
                $TempRelevantDeviceLines += $line
            }
        }
        elseif(($RelevantEntryBegin -or $RelevantEntryUpdateBegin) -and $RelevantEntryEnd)
        {
            if($RelevantEntryBegin)
            {
            #Add last line to current Device Entry
                $TempRelevantDeviceLines += $line
            }

            #If $TempRelevantDeviceLines not empty
            if($TempRelevantDeviceLines.Count -gt 0)
            {
                #Add Entry to Array of RelevantDeviceEntrys
                    $TempRelevantDeviceEntry = [DeviceEntry]::new()
                    $TempRelevantDeviceEntry.Type = $Type
                    $TempRelevantDeviceEntry.time = $Time
                    $TempRelevantDeviceEntry.lines = $TempRelevantDeviceLines

                    $RelevantDeviceEntrys += $TempRelevantDeviceEntry
            }
            #Prepare for search of next relevant entry
                $TempRelevantDeviceLines = @()
                $line = ""
                $RelevantEntryBegin = $false
                $RelevantEntryUpdateBegin = $false
                $RelevantEntryEnd = $false
        }
        Write-Progress -Activity "Searching Device Entrys" -Status "Finding SerialNumber:$($DeviceSerialNumber)" -PercentComplete($i/$SetupapiNumberofLines*100) 
        $i++
    }

    #Summary
        Write-Host "--------------------------------------------"
        Write-Host "$($RelevantDeviceEntrys.Count) Entrys found."
}
else
{ Write-Host "$($SetupapiPath) not found!" -ForegroundColor Red }
