### Setting parameters
param
(
    [Parameter(Mandatory=$false)][Switch]$Install,
    [Parameter(Mandatory=$false)][Switch]$Uninstall,
    [Parameter(ValueFromRemainingArguments=$true)] $args
)

# LOGGING INITIALISATION
$logSource = "OneDrive Per-Machine Deployment"
if (![System.Diagnostics.EventLog]::SourceExists($logSource))
{
        New-Eventlog -LogName Application -Source $logSource
}
$installationSource = "\\Shared\PathTo\OneDriveFolder"
$destinationPath = "C:\Program Files\Microsoft OneDrive"
if (Test-Path -Path "$destinationPath\OneDrive.exe" -PathType Leaf)
{
    $installedVersion = (Get-Command "$destinationPath\OneDrive.exe").FileVersionInfo.FileVersion
}
else
{
    Write-Eventlog -LogName Application -Source $logSource -EntryType Information -EventId 800 -Message "Can't locate installed OneDrive application on the default path."
    $installedVersion = 0
}

### Check the installed OneDrive's version
try
{
    $ErrorActionPreference = "Stop"
    $targetVersion = (Get-Command "$installationSource\OneDriveSetup.exe").FileVersionInfo.FileVersion
} 
catch
{
    Write-Eventlog -LogName Application -Source $logSource -EntryType Error -EventId 900 -Message "Unable to determine target OneDrive version - check network connectivity or existence of deployment files."
    Exit
}

### Installation
If ($Install)
{
    try 
    {
        if ($targetVersion -gt $installedVersion)
            {
            Write-Eventlog -LogName Application -Source $logSource -EntryType Information -EventId 1 -Message "Microsoft OneDrive not installed or out of date. Installed version: $installedVersion; target version: $targetVersion. Installation starting..."
                if (Test-Path ($destinationPath))
                    {
                    Remove-Item $destinationPath -recurse
                    Write-Eventlog -LogName Application -Source $logSource -EntryType Information -EventId 2 -Message "Existing OneDrive installation removed"
                    }
            Start-Process -FilePath "$installationSource\OneDriveSetup.exe" -ArgumentList "/Allusers /firstSetup"
            Write-Eventlog -LogName Application -Source $logSource -EntryType Information -EventId 5 -Message "OneDrive installation complete"
            }
        else
        {
            Write-Eventlog -LogName Application -Source $logSource -EntryType Information -EventId 1 -Message "Microsoft OneDrive is installed on the computer. Installed version: $installedVersion"
        }
	} 
    catch [exception] 
    {
        Write-Eventlog -LogName Application -Source $logSource -EntryType Information -EventId 1 -Message "An error occured while installing OneDrive client"
	Exit
    }
}
### Uninstallation
If ($Uninstall)
{
    try 
    {
        if (Test-Path ($destinationPath))
            {
            $filePath = "$destinationPath\$installedVersion\OneDriveSetup.exe"
            Start-Process -FilePath $filePath -ArgumentList "/uninstall /allusers" -Wait
            Remove-Item $destinationPath -recurse
            Write-Eventlog -LogName Application -Source $logSource -EntryType Information -EventId 2 -Message "The OneDrive has been sucessfully removed."
            }
	} 
    catch [exception] 
    {
        Write-Eventlog -LogName Application -Source $logSource -EntryType Information -EventId 1 -Message "An error occured while uninstalling OneDrive client."
	Exit
    }
}
