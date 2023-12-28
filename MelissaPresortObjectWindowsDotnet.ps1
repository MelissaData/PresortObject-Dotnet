# Name:    MelissaPresortObjectWindowsDotnet
# Purpose: Use the Melissa Updater to make the MelissaPresortObjectWindowsDotnet code usable


######################### Parameters ##########################

param($file ='""', $license = '', [switch]$quiet = $false )

######################### Classes ##########################

class FileConfig {
  [string] $FileName;
  [string] $ReleaseVersion;
  [string] $OS;
  [string] $Compiler;
  [string] $Architecture;
  [string] $Type;
}

######################### Config ###########################

$RELEASE_VERSION = '2023.11'
$ProductName = "presort_data"

# Uses the location of the .ps1 file 
# Modify this if you want to use 
$CurrentPath = $PSScriptRoot
Set-Location $CurrentPath
$ProjectPath = "$CurrentPath\MelissaPresortObjectWindowsDotnet"
$DataPath = "$ProjectPath\Data"
$BuildPath = "$ProjectPath\Build"

If (!(Test-Path $DataPath)) {
  New-Item -Path $ProjectPath -Name 'Data' -ItemType "directory"
}

If (!(Test-Path $BuildPath)) {
  New-Item -Path $ProjectPath -Name 'Build' -ItemType "directory"
}


$DLLs = @(
  [FileConfig]@{
    FileName       = "mdPresort.dll";
    ReleaseVersion = $RELEASE_VERSION;
    OS             = "WINDOWS";
    Compiler       = "DLL";
    Architecture   = "64BIT";
    Type           = "BINARY";
  }
)

$Wrapper          = [FileConfig]@{
  FileName        = "mdPresort_cSharpCode.cs";
  ReleaseVersion  = $RELEASE_VERSION;
  OS              = "ANY";
  Compiler        = "NET";
  Architecture    = "ANY" ;
  Type            = "INTERFACE"
}

######################## Functions #########################

function DownloadDataFiles([string] $license) {
  $DataProg = 0
  Write-Host "=========================== MELISSA UPDATER =========================="
  Write-Host "MELISSA UPDATER IS DOWNLOADING DATA FILE(S)..."

  .\MelissaUpdater\MelissaUpdater.exe manifest -p $ProductName -r $RELEASE_VERSION -l $license -t $DataPath 
  if($? -eq $False ) {
    Write-Host "`nCannot run Melissa Updater. Please check your license string!"
    Exit
  }     
  Write-Host "Melissa Updater finished downloading data file(s)!"

}

function DownloadDLLs() {
  Write-Host "MELISSA UPDATER IS DOWNLOADING DLL(s)..."
  $DLLProg = 0
  foreach ($DLL in $DLLs) {
    Write-Progress -Activity "Downloading DLL(s)" -Status "$([math]::round($DLLProg / $DLLs.Count * 100, 2))% Complete:"  -PercentComplete ($DLLProg / $DLLs.Count * 100)

    # Check for quiet mode
    if ($quiet) {
      .\MelissaUpdater\MelissaUpdater.exe file --filename $DLL.FileName --release_version $DLL.ReleaseVersion --license $LICENSE --os $DLL.OS --compiler $DLL.Compiler --architecture $DLL.Architecture --type $DLL.Type --target_directory $BuildPath > $null
      if(($?) -eq $False) {
          Write-Host "`nCannot run Melissa Updater. Please check your license string!"
          Exit
      }
    }
    else {
      .\MelissaUpdater\MelissaUpdater.exe file --filename $DLL.FileName --release_version $DLL.ReleaseVersion --license $LICENSE --os $DLL.OS --compiler $DLL.Compiler --architecture $DLL.Architecture --type $DLL.Type --target_directory $BuildPath 
      if(($?) -eq $False) {
          Write-Host "`nCannot run Melissa Updater. Please check your license string!"
          Exit
      }
    }
    
    Write-Host "Melissa Updater finished downloading " $DLL.FileName "!"
    $DLLProg++
  }
}

function DownloadWrapper() {
  Write-Host "MELISSA UPDATER IS DOWNLOADING WRAPPER(S)..."

  # Check for quiet mode
  if ($quiet) {
    .\MelissaUpdater\MelissaUpdater.exe file --filename $Wrapper.FileName --release_version $Wrapper.ReleaseVersion --license $LICENSE --os $Wrapper.OS --compiler $Wrapper.Compiler --architecture $Wrapper.Architecture --type $Wrapper.Type --target_directory $ProjectPath > $null
    if(($?) -eq $False) {
        Write-Host "`nCannot run Melissa Updater. Please check your license string!"
        Exit
    }
  }
  else {
    .\MelissaUpdater\MelissaUpdater.exe file --filename $Wrapper.FileName --release_version $Wrapper.ReleaseVersion --license $LICENSE --os $Wrapper.OS --compiler $Wrapper.Compiler --architecture $Wrapper.Architecture --type $Wrapper.Type --target_directory $ProjectPath 
    if(($?) -eq $False) {
        Write-Host "`nCannot run Melissa Updater. Please check your license string!"
        Exit
    }
  }

  Write-Host "Melissa Updater finished downloading " $Wrapper.FileName "!"
}

function CheckDLLs() {
  Write-Host "`nDouble checking dll(s) were downloaded...`n"
  $FileMissing = $false 
  if (!(Test-Path ("$BuildPath\mdPresort.dll"))) {
    Write-Host "mdPresort.dll not found." 
    $FileMissing = $true
  }
  if ($FileMissing) {
    Write-Host "`nMissing the above data file(s).  Please check that your license string and directory are correct."
    return $false
  }
  else {
    return $true
  }
}


########################## Main ############################

Write-Host "`n====================== Melissa Presort Object ======================`n                     [ .NET | Windows | 64BIT ]`n"

# Get license (either from parameters or user input)
if ([string]::IsNullOrEmpty($license) ) {
  $License = Read-Host "Please enter your license string"
}

# Check for License from Environment Variables 
if ([string]::IsNullOrEmpty($License) ) {
  $License = $env:MD_LICENSE # Get-ChildItem -Path Env:\MD_LICENSE   #[System.Environment]::GetEnvironmentVariable('MD_LICENSE')
}

if ([string]::IsNullOrEmpty($License)) {
  Write-Host "`nLicense String is invalid!"
  Exit
}

# Use Melissa Updater to download data file(s) 
# Download data file(s) 
DownloadDataFiles -license $License      # comment out this line if using DQS Release

# Set data file(s) path
#$DataPath = "C:\Program Files\Melissa DATA\DQT\Data"      # uncomment this line and change to your DQS Release data file(s) directory 

# Download dll(s)
DownloadDlls -license $License

# Download wrapper(s)
DownloadWrapper -license $License

# Check if all dll(s) have been downloaded. Exit script if missing
$DLLsAreDownloaded = CheckDLLs

if (!$DLLsAreDownloaded) {
  Write-Host "`nAborting program, see above.  Press any button to exit."
  $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
  exit
}

Write-Host "All file(s) have been downloaded/updated! "

# Start Program
# Build project
Write-Host "`n=========================== BUILD PROJECT =========================="

dotnet publish -f="net7.0" -c Release -o $BuildPath MelissaPresortObjectWindowsDotnet\MelissaPresortObjectWindowsDotnet.csproj

# Run project
if ([string]::IsNullOrEmpty($file)) {
  Push-Location MelissaPresortObjectWindowsDotnet
  dotnet $BuildPath\MelissaPresortObjectWindowsDotnet.dll --license $License  --dataPath $DataPath
  Pop-Location
}
else {
  Push-Location MelissaPresortObjectWindowsDotnet
  dotnet $BuildPath\MelissaPresortObjectWindowsDotnet.dll --license $License  --dataPath $DataPath --file $file
  Pop-Location
}
