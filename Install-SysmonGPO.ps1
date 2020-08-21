function Install-SysmonGPO {
    <#
    .SYNOPSIS 
    Uses existing PS-Session to your DC to install a Sysmon GPO
    .DESCRIPTION
    This command  downloads the latest Sysmon and Sysmon Modular Config files and builds a GPO to install Sysmon with the config. If there is no PSSession just make your connection and rerun the script. It will not download or create anything new if it has the files inplace.
    .PARAMETER FullyQualifiedDomainNameofDC
    The fully qualified Domain name of your DC. Example: dc.corp.local
    .PARAMETER FullyQualifiedDomain
    The fully qualifed Domain. Example: corp.local
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
if(-not $(Test-Path $dir)){ New-Item -ItemType Directory -Path $dir }
Write-Host "Creating a new folder here: $dir"
Set-Location $dir

#Downloads current Sysmon to $dir
$SysinternalsDownloadURL = "https://download.sysinternals.com/files/"
$sysmonZip = "Sysmon.zip"
$dlSysmonUrl = "$($SysinternalsDownloadURL)$sysmonZip"
$dlSysmonPath = "$($dir)$sysmonZip"
if(!$(Test-Path $dlSysmonPath)){
    Invoke-WebRequest $dlSysmonUrl -OutFile $dlSysmonPath
    
#Unzip Sysmon
Expand-Archive -Path $dlSysmonPath -DestinationPath $dir
   } #End Sysmon Download

#Download https://github.com/olafhartong/sysmon-modular/archive/master.zip

$sysmonmodularZIP = "sysmon-modular.zip" 
$dlsysmonmodularPath = "$($dir)$sysmonmodularZIP"
if(!$(Test-Path $dlsysmonmodularPath)) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest "https://github.com/olafhartong/sysmon-modular/archive/master.zip" -OutFile $dlsysmonmodularPath
    Write-Information -MessageData "Please take the time to build your Sysmon Config using the files downloaded from olafhartong/sysmon-modular. It will help to tune your sysmon to your envirment" -InformationAction Continue
    #unzip sysmon-modular
    Expand-Archive -Path $dlsysmonmodularPath -DestinationPath $dir
    } #End Sysmon Modular Download
#Builds batch file
$batchfile = @"

@echo off
:: Author: Ryan Watson 
:: Edited for Install-SysmonGPO by @CWObuzz
:: Twitter: @gentlemanwatson
:: Version: 2.0
:: Credits: Credit to Syspanda.com and their Sysmon GPO article for the kick off point 
:: https://www.syspanda.com/index.php/2017/02/28/deploying-sysmon-through-gpo/

:: ** IMPORTANT **
:: 1) Create a Sysmon folder with the SYSVOL share on your domain controller
:: 2) Download Sysmon from Microsoft and place both sysmon.exe and sysmon64.exe in
::    newly created Sysmon folder
:: 3) Download a sample sysmon config from SwiftOnSecurity, rename the file to
::    sysmonConfig.xml and place it within the Sysmon folder
:: 4) Enter the appropriate values for your DC and FQDN below.
:: 5) Create a GPO that will launch this batch file on startup.
:: 6) Apply the GPO to your specified OUs. 
::
:: Note: It is recommended that the Sysmon binaries and the Sysmon config file
:: be placed in the sysvol folder on the Domain Controller. The goal
:: being that the computers can read from the folder, but no one except for
:: domain admins have the ability to write to the folder hosting the files. 
:: Otherwise this will be a great way for attackers to escalate privs 
:: in the domain. You have been warned. 


SET DC=$FullyQualifiedDomainNameofDC
SET FQDN=$FullyQualifiedDomain

:: Determine architecture to set Arch Type for the SYSMON Binary

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
if(-not $(Test-Path $FileStagingdir)){ New-Item -ItemType Directory -Path $FileStagingdir }
Write-Host "Creating a new folder here: $FileStagingdir"
Copy-Item -Path $dir\sysmon.exe -Destination $FileStagingdir
Copy-Item -Path $dir\sysmon64.exe -Destination $FileStagingdir
Copy-Item -Path $dir\SysmonInstall.bat -Destination $FileStagingdir
Copy-Item -Path $dir\sysmon-modular-master\sysmonconfig.xml -Destination $FileStagingdir

#Checks for PSsession to DC
if(!$(Get-PSSession -ComputerName $FullyQualifiedDomainNameofDC)) {
    Write-Host "Please make a PSession with your DC as Domain Admin and rerun script, it will skip downloading if files are inplace"    
    } else {
        $SysVolPath = "\\$FullyQualifiedDomainNameofDC\sysvol\"
        Copy-Item $FileStagingdir -Destination $SysVolPath -Force -Recurse
    }
    Write-Host "Now that we have all the files within the Sysmon folder within SYSVOL, we can now create the GPO to perform the deployment. Take the following steps to create the GPO:

    Create a new GPO and title it SYSMON Deploy
    Navigate to Computer Configuration –> Policies –> Windows Settings –> Scripts (Startup/Shutdown)
    Right-click on top of Startup and select Properties.
    In the Startup Properties window, click on Add, then on Browser and navigate to the SysmonStartup.bat
    Click the OK buttons to save and close.
    Lastly, linked the GPO to all the OUs you wish to deploy Sysmon to.
"
} #End of Function