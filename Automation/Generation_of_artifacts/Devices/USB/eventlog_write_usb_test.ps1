#Configuration
    $EventProvider = "Microsoft-Windows-DriverFrameworks-UserMode/Operational"
    $EventID = 10000
        $EventPayload_DeviceId = "SWD\WPDBUSENUM\_??_USBSTOR#DISK&VEN_KINGSTON&PROD_DATATRAVELER_2.0&REV_PMAP#00241D8CE459BFB0194935D4&0#{53F56307-B6BF-11D0-94F2-00A0C91EFB8B}"
        $EventPayload_FrameworkVersion = "2.31.0"
#Add Eventlog Entry
New-WinEvent -ProviderName $EventProvider -Id $EventID -Payload @($EventPayload_DeviceId, $EventPayload_FrameworkVersion)