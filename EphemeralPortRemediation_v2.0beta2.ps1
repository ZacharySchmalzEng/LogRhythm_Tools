<####################################################################################################################
#Author: ZS (MgrPS)
#Version: 2.0
#Version changes:  v2.0 Introduced diagnostic and logging logic. v1.4.1 Modified wording of the Windows registry 
#reload prompt v1.4 Added a reload of the Windows Registry v1.3 Cleaner output v1.2 added elevation functions 
#v1.1 Fixed typos v1.0 Original release
#
#Description: 
#The LogRhythm application heavily utilizes ephemeral ports. Many third-party applications, especially newer endpoint
#protection solutions can cause excessive TIME_WAIT to form over the ephemeral ranges.
#
#This script is designed to expand the available range and make the OS more aggressive in truncating hung ports,
#thus minimizing the impacts. However, this script should only be considered a short-term solution while a partner
#or internal personnel seeks the root cause of the port exhaustion.  
#
#Disclaimer: 
#This script does not fully remediate port exhaustion; rather, it lessens the impacts. LogRhythm is not responsible
#for identifying the root causes of ephemeral port exhaustion. 
#
#See https://docs.microsoft.com/en-us/windows/client-management/troubleshoot-tcpip-port-exhaust for more information
####################################################################################################################>
# Start transcript logging to a file in the script's directory
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$logFilePath = Join-Path $scriptDirectory 'PortScript.log'
$PortActivityLogPath = Join-Path $scriptDirectory 'PortActivity.log'
Start-Transcript -Path $logFilePath -Append

function LogMessage {
    param (
        [string]$message
    )
    # Log message to console and transcript
    Write-Host $message
}

function IsAdministrator {
    $principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function ExpandEphemeralPorts {
    LogMessage "Expanding ephemeral port values"
    netsh int ipv4 set dynamicport tcp start=10000 num=55535 | Out-Null
    netsh int ipv4 set dynamicport udp start=10000 num=55535 | Out-Null
    LogMessage "Expanded ephemeral port values`n"
}

function UpdateRegistry {
    LogMessage "Updating registry keys"
    $path = 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters'
    
    New-ItemProperty -Path $path -Name 'TcpTimedWaitDelay' -Value 30 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $path -Name 'StrictTimeWaitSeqCheck' -Value 1 -PropertyType DWord -Force | Out-Null
    
    LogMessage "Updated HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\ with the following values:`nTcpTimedWaitDelay to 30 `nStrictTimeWaitSeqCheck to 1`n"
}

function RebootPrompt {
    $title = 'A reboot is required to complete the remediation.'
    $question = 'Would you like to reboot now?'

    $choices = @('&Yes', '&No')
    $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)

    if ($decision -eq 0) {
        LogMessage "`nRebooting in one minute.`n"
        shutdown /r /t 60
    } else {
        RestartExplorerPrompt
    }
}

function RestartExplorerPrompt {
    $title = "Reloading the Registry by restarting the explorer.exe process is recommended.`n"
    $question = 'Would you like to restart explorer.exe now?'

    $choices = @('&Yes', '&No')
    $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)

    if ($decision -eq 0) {
        taskkill /f /im explorer.exe
        Start-Process explorer.exe
    } else {
        LogMessage "Please complete remediation during the next maintenance window."
        Pause
        exit
    }
}

function Main {
	try {
		StartElevatedScript
		ExpandEphemeralPorts
		UpdateRegistry
		RebootPrompt
	} catch {
		LogMessage "An error occurred: $_"
	}
}

function StartElevatedScript {
    LogMessage "Elevating script..."
    Start-Process -FilePath "powershell" -ArgumentList "$('-File ""')$(Get-Location)$('\')$($MyInvocation.MyCommand.Name)$('""')" -Verb runAs
}

# Main execution starts here

# Redirect the output to a variable to check the count later
$connectionInfo = Get-NetTCPConnection | Group-Object -Property State, OwningProcess | 
    Select-Object Count, Name, @{Name="ProcessName";Expression={(Get-Process -PID ($_.Name.Split(',')[-1].Trim(' '))).Name}}, Group | 
    Sort-Object Count -Descending

# Log the content of $connectionInfo to a separate file
$connectionInfo | Out-File -FilePath $PortActivityLogPath 

if ($connectionInfo.Count -gt 3000) {
    LogMessage "Single process port count is greater than 3000...`nThis indicates an port exhaustion issue.`nBeginning remediation"
    Main
} else {
    LogMessage "Single process port count is not greater than 3000. Port exhaustion not detected."
}

# Stop transcript logging
Stop-Transcript
Pause