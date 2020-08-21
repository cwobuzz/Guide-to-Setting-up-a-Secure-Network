function Install-WinlogBeatGPO {
    <#
    .SYNOPSIS 
    Uses existing PS-Session to your DC to install WinLogBeat GPO. Downloads all software and helps you configure it.
    .DESCRIPTION
    This command  downloads the latest WinlogBeat files and builds a Batch file. If there is a PSSession to your DC it will push the files to the DC. 
    If there is no PSSession just make your connection and rerun the script. It will not download or create anything new if it has the files inplace.
    .PARAMETER FullyQualifiedDomainNameofDC
    The fully qualified Domain name of your DC. Example: dc.corp.local
    .PARAMETER FullyQualifiedDomain
    The fully qualifed Domain. Example: corp.local
    .PARAMETER IPofElasticServer
    The IP address of your Elastic Server. Example: 192.168.5.50
    .PARAMETER WinlogBeatPort
    The port you want WinlogBeat to use. Example: 9700
    #>
        [CmdletBinding()]
        Param(
        #Set the Sysmon File Path
            [Parameter(Mandatory=$True,
            ValueFromPipelineByPropertyName=$True)]
            [ValidateNotNull()]
            [String]$FullyQualifiedDomainNameofDC,

            [Parameter(Mandatory=$True,
            ValueFromPipelineByPropertyName=$True)]
            [ValidateNotNull()]
            [String]$FullyQualifiedDomain

        ) #End of Params
            
            
#Create Folder and sets current directory there
$dir = "$env:HOMEPATH\Desktop\Host_Tools\"
if(-not $(Test-Path $dir)){ New-Item -ItemType Directory -Path $dir -InformationAction SilentlyContinue }
Write-Host "Creating a new folder here: $dir"
Set-Location $dir -InformationAction SilentlyContinue

#Downloads current Winlogbeat to $dir
$ProgressPreference = 'SilentlyContinue'
$ElasticDownloadURL = (Invoke-WebRequest -Uri "https://www.elastic.co/downloads/beats/winlogbeat" -UseBasicParsing).Links.Href
$32WinlogbeatURL = $ElasticDownloadURL -clike "*x86.zip" | Out-String -Stream
$64WinlogbeatURL = $ElasticDownloadURL -clike "*_64.zip" | Out-String -Stream
$32WinlogbeatZip = "32Winlogbeat.zip"
$64WinlogbeatZIP = "64Winlogbeat.zip"
$32WinlogbeatOutPath = "$($dir)$32WinlogbeatZip"
$64WinlogbeatOutPath = "$($dir)$64WinlogbeatZip"
if(!$(Test-Path $32WinlogbeatOutPath)){
    Invoke-WebRequest $32WinlogbeatURL -OutFile $32WinlogbeatOutPath    
#Unzip Winlogbeat
Expand-Archive -Path $32WinlogbeatOutPath -DestinationPath $dir
   } #End 32 Winlogbeat Download
if(!$(Test-Path $64WinlogbeatOutPath)){
    Invoke-WebRequest $64WinlogbeatURL -OutFile $64WinlogbeatOutPath    
    #Unzip Winlogbeat
    Expand-Archive -Path $64WinlogbeatOutPath -DestinationPath $dir
    Write-Host -BackgroundColor Red "Configure Both 32bit and 64bit WinLogbeat configuration files to work on your network, then rerun this script for it to finish installing!!!"    
    Break
    } #End 64 Winlogbeat Download

#Builds batch file
$batchfile = @"

@echo off
:: Author: CWObuzz
:: Twitter: @CWObuzz
:: Version: 1.0
:: Credits: Credit to Ryan Watson @gentlemanwatson for his work on sysmon version of this which I am modifying 
::
:: Note: It is recommended that the Winlogbeat binaries and the Winlogbeat config file
:: be placed in the sysvol folder on the Domain Controller. The goal
:: being that the computers can read from the folder, but no one except for
:: domain admins have the ability to write to the folder hosting the files. 
:: Otherwise this will be a great way for attackers to escalate privs 
:: in the domain. You have been warned. 

SET DC=$FullyQualifiedDomainNameofDC
SET FQDN=$FullyQualifiedDomain

:: Determine architecture to set Arch Type for the Winlogbeat Binary

IF EXIST "C:\Program Files (x86)" (
SET BINARCH=sysmon64.exe
SET SERVBINARCH=Sysmon64
) ELSE (
SET BINARCH=sysmon.exe
SET SERVBINARCH=Sysmon
)

SET SYSMONDIR=C:\windows\sysmon
SET SYSMONBIN=%SYSMONDIR%\%BINARCH%
SET SYSMONCONFIG=%SYSMONDIR%\sysmonconfig.xml

:: Checks file path for %FQDN% 
IF EXIST \\%DC%\sysvol\%FQDN% (
SET GLBSYSMONBIN=\\%DC%\sysvol\%FQDN%\Sysmon\%BINARCH%   
SET GLBSYSMONCONFIG=\\%DC%\sysvol\%FQDN%\Sysmon\sysmonconfig.xml
) ELSE (
SET GLBSYSMONBIN=\\%DC%\sysvol\Sysmon\%BINARCH%
SET GLBSYSMONCONFIG=\\%DC%\sysvol\Sysmon\sysmonconfig.xml
) 

sc query "%SERVBINARCH%" | Find "RUNNING"
If "%ERRORLEVEL%" EQU "1" (
goto startsysmon
)
  
:installsysmon
IF Not EXIST %SYSMONDIR% (
mkdir %SYSMONDIR%
)
xcopy %GLBSYSMONBIN% %SYSMONDIR% /y
xcopy %GLBSYSMONCONFIG% %SYSMONDIR% /y
chdir %SYSMONDIR%
%SYSMONBIN% -i %SYSMONCONFIG% -accepteula -h md5,sha256 -n -l
sc config %SERVBINARCH% start= auto
  
:updateconfig

fc  %SYSMONCONFIG% %GLBSYSMONCONFIG% > nul
If “%ERRORLEVEL%” EQU “1” (
xcopy %GLBSYSMONCONFIG% %SYSMONCONFIG% /y
chdir %SYSMONDIR%
%SYSMONBIN% -c %SYSMONCONFIG%
EXIT /B 0
)
  
:startsysmon
sc start %SERVBINARCH%
If "%ERRORLEVEL%" EQU "1060" (
goto installsysmon
) ELSE (
goto updateconfig
)
"@
$sysmonbatPath = $($dir + "SysmonInstall.bat")
Out-file -FilePath $sysmonbatPath -InputObject $batchfile -Encoding ascii

#Create Folder and sets current directory there
$FileStagingdir = "$env:HOMEPATH\Desktop\Host_Tools\Sysmon\"
if(-not $(Test-Path $FileStagingdir)){ New-Item -ItemType Directory -Path $FileStagingdir -InformationAction SilentlyContinue}
Write-Information -InformationAction Continue "Creating a new folder here: $FileStagingdir"
Copy-Item -Path $dir\sysmon.exe -Destination $FileStagingdir
Copy-Item -Path $dir\sysmon64.exe -Destination $FileStagingdir
Copy-Item -Path $dir\SysmonInstall.bat -Destination $FileStagingdir
Copy-Item -Path $dir\sysmon-modular-master\sysmonconfig.xml -Destination $FileStagingdir

#Checks for PSsession to DC
if(!$(Get-PSSession -ComputerName $FullyQualifiedDomainNameofDC -ErrorAction SilentlyContinue)) {
    Write-Host "Please make a PSession with your DC as Domain Admin and rerun script, it will skip downloading if the files are inplace" -BackgroundColor Red   
    } else {
        $SysVolPath = "\\$FullyQualifiedDomainNameofDC\sysvol\"
        Copy-Item $FileStagingdir -Destination $SysVolPath -Force -Recurse
    }
    Write-Information -InformationAction Continue -Tags "GPOInstruction" -MessageData "Now that we have all the files within the Sysmon folder within SYSVOL, we can now create the GPO to perform the deployment. 
    Take the following steps to create the GPO:

    Create a new GPO and title it SYSMON Deploy
    Navigate to Computer Configuration –> Policies –> Windows Settings –> Scripts (Startup/Shutdown)
    Right-click on top of Startup and select Properties.
    In the Startup Properties window, click on Add, then on Browser and navigate to the SysmonStartup.bat
    Click the OK buttons to save and close.
    Lastly, linked the GPO to all the OUs you wish to deploy Sysmon to.
"
} #End of Function