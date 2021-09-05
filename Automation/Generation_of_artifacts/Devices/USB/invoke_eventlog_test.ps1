#Configuration
    #Path to PsExec
    $PsExecPath = "C:\tools\PsExec.exe"
    $ScriptAsSystemPath = "C:\Users\semda\Documents\eventlog_write_usb_test.ps1"

#Run Powershell-Script as SYSTEM
    $process = Start-Process -PassThru -FilePath cmd.exe -Verb Runas -ArgumentList "/k $($PsExecPath) -i -s powershell.exe -ExecutionPolicy Bypass -Command ""$($ScriptAsSystemPath)"""
    Start-Sleep -Seconds 5

# Stop all Processes (-Name cmd) which did not started before RunAs

    Stop-Process $process

