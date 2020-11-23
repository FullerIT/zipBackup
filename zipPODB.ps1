<#Begin Header#>
<#
 Notes:
    Either set "Set-ExecutionPolicy Unrestricted" or run the script with 
 
 
 
#>
<#End Header#>

<#Common Starter & Code Blocks#>

#Debugging
#$DebugPreference = "Continue"
#Logging feature
#$ErrorActionPreference="SilentlyContinue"
try { Stop-Transcript | out-null } catch { }

#start a transcript file
try { Start-Transcript -path $scriptLog } catch { }

# TimeStamps
$Time = Get-Date -format "yyyy-MM-dd-hh-mm-ss"

#Path variables https://docs.microsoft.com/en-us/dotnet/api/system.environment.getfolderpath?view=net-5.0
$DesktopPath = [Environment]::GetFolderPath("Desktop")
$AllUsersDesktopPath = [Environment]::GetFolderPath("CommonDesktopDirectory")

#current script directory
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
#current script name
$path = Get-Location
$scriptName = $MyInvocation.MyCommand.Name
$scriptLog = "$scriptPath\log\$scriptName.log"
#start a transcript file
Start-Transcript -path $scriptLog

<#Pause#>
#Pause
Function Pause($M="Press any key to continue . . . "){If($psISE){$S=New-Object -ComObject "WScript.Shell";$B=$S.Popup("Click OK to continue.",0,"Script Paused",0);Return};Write-Host -NoNewline $M;$I=16,17,18,20,91,92,93,144,145,166,167,168,169,170,171,172,173,174,175,176,177,178,179,180,181,182,183;While($K.VirtualKeyCode -Eq $Null -Or $I -Contains $K.VirtualKeyCode){$K=$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")};Write-Host}

<# Begin Program #>
#Write-Host "Make sure you've closed the PO Database file before running the backup"
#Pause
Add-Type -AssemblyName PresentationCore,PresentationFramework
$msgBoxInput =  [System.Windows.MessageBox]::Show('Close the access database before running this script. Is it closed?','Is Access Closed?','YesNoCancel','Error')

switch ($msgBoxInput) {

  'Yes' {
    $zipfile = "$DesktopPath\$($Time)-PODB.zip"
    $zipfile2 = "C:\PO Backups\$($Time)-PODB.zip"
    Compress-Archive -Path 'C:\PO Orders\*' -DestinationPath $zipfile -CompressionLevel Optimal
    Compress-Archive -Path 'C:\PO Orders\*' -DestinationPath $zipfile2 -CompressionLevel Optimal
  }

  'No' {

    Exit

  }
  
  'Cancel' {

    Exit

  }
}
  
<# End Program #>

<#Begin Footer#>
#Close all open sessions
try
{
	Remove-PSSession $Session
}
catch
{
   #Just suppressing Error Dialogs
}

Get-PSSession | Remove-PSSession
#Close Transcript log
Stop-Transcript
<#End Footer#>