# Melissa - Presort Object Windows Dotnet

## Purpose
This code showcases the Melissa Presort Object using C#.

Please feel free to copy or embed this code to your own project. Happy coding!

For the latest Melissa Presort Object release notes, please visit https://releasenotes.melissa.com/on-premise-api/presort-object/

The console will ask the user for:

- A csv file that contains addresses you would like to presort

And return 

- A csv file of presorted addresses

## Tested Environments

- Windows 64-bit .NET 7.0, .NET 6.0, .NET 5.0, .NET Core 3.1
- Powershell 5.1
- Melissa data files for 2023-03

## Required File(s) and Programs

#### mdPresort.dll
This is the c++ code of the Melissa Data Object.

#### Data File(s) 
- mdL004A.dat
- mdL004B.dat
- mdL004C.dat
- mdL005.dat
- mdL006A.dat
- mdL006B.dat
- mdL007.dat
- mdL008.dat
- mdL009A.dat
- mdL009B.dat
- mdL010.dat
- mdL011A.dat
- mdL011B.dat
- mdL012.dat
- mdL201A.dat
- mdL201B.dat
- mdL601.dat
- mdL602.dat
- mdL801A.dat
- mdL801B.dat
- mdPresort.lic
- mdValid5.dat
- PostalRatesFile.txt
- 3600r_T.pdf
- 3602c_T.pdf
- 3602n_T.pdf
- 3602r_T.pdf
- 8125_T.pdf
- mdCityState.dat
- mdDSF.DAT
- mdL001.dat
- mdL002A.dat
- mdL002B1.dat
- mdL002B2.dat
- mdL002C.dat

## Getting Started
These instructions will get you a copy of the project up and running on your local machine for development and testing purposes.

This project is compatible with .NET 7.0, .NET 6.0, .NET 5.0, and .NET Core 3.1. If you would like to run this project for any other version besides .NET 7.0, proceed with the following procedures but check for and download your desired .NET version.

#### Install the Dotnet Core SDK
Before starting, make sure that the .NET 7.0 SDK has been correctly installed on your machine (If you have Visual Studio installed, you most likely have it already). If you are unsure, you can check by opening a command prompt window and typing the following:

`dotnet --list-sdks`

If the .NET 7.0 SDK is already installed, you should see it in the following list:

![alt text](/screenshots/dotnet_output.png)

As long as the above list contains version `7.0.xxx` (underlined in red), then you can skip to the next step. If your list does not contain version 7.0, or you get any kind of error message, then you will need to download and install the .NET 7.0 SDK from the Microsoft website.

To download, follow this link: https://dotnet.microsoft.com/en-us/download/dotnet

Select `.NET 7.0` and you will be navigated to the download page.

Click and download the `x64` SDK installer for your operating system.

(IMPORTANT: Make sure you download the SDK, NOT the runtime. the SDK contains both the runtime as well as the tools needed to build the project.)

![alt text](/screenshots/net7.png)

Once clicked, your web browser will begin downloading an installer for the SDK. Run the installer and follow all of the prompts to complete the installation (your computer may ask you to restart before you can continue). Once all of that is done, you should be able to verify that the SDK is installed with the `dotnet --list-sdks` command.

#### Set up Powershell settings
If running Powershell for the first time, you will need to run this command in the Powershell console: `Set-ExecutionPolicy RemoteSigned`.
The console will then prompt you with the following warning shown in the image below. 
 - Enter `'A'`. 
 	- If successful, the console will not output any messages. (You may need to run Powershell as administrator to enforce this setting).
	
 ![alt text](/screenshots/powershell_executionpolicy.png)

----------------------------------------

#### Download this project
```
$ git clone https://github.com/MelissaData/PresortObject-Dotnet
$ cd PresortObject-Dotnet
```

#### Set up Melissa Updater 
Melissa Updater is a CLI application allowing the user to update their Melissa applications/data. 

- Download Melissa Updater here: <https://releases.melissadata.net/Download/Library/WINDOWS/NET/ANY/latest/MelissaUpdater.exe>
- Create a folder within the cloned repo called `MelissaUpdater`.
- Put `MelissaUpdater.exe` in `MelissaUpdater` folder you just created.

----------------------------------------

#### Different ways to get data file(s)
1.  Using Melissa Updater
	- It will handle all of the data download/path and dll(s) for you. 
2.  If you already have the latest DQS Release (ZIP), you can find the data file(s) and dll(s) in there
	- Use the location of where you copied/installed the data and update the "$DataPath" variable in the powershell script.
	- Copy all the dll(s) mentioned above into the `MelissaPresortObjectWindowsDotnet` project folder.
	
----------------------------------------

#### Configure Target Framework
Depending on your target .NET framework, you may need to configure the powershell script. In order to do this, open up the `MelissaPresortObjectWindowsDotnet.ps1` for editing, proceed to the bottom of the script where you will find this section of code.

Default set for .NET 7.0
```
dotnet publish -f="net7.0" -c Release -o $BuildPath MelissaPresortObjectWindowsDotnet\MelissaPresortObjectWindowsDotnet.csproj
#dotnet publish -f="net6.0" -c Release -o $BuildPath MelissaPresortObjectWindowsDotnet\MelissaPresortObjectWindowsDotnet.csproj
#dotnet publish -f="net5.0" -c Release -o $BuildPath MelissaPresortObjectWindowsDotnet\MelissaPresortObjectWindowsDotnet.csproj
#dotnet publish -f="netcoreapp3.1" -c Release -o $BuildPath MelissaPresortObjectWindowsDotnet\MelissaPresortObjectWindowsDotnet.csproj
```
The target framework is specified with the -f flag found in the command line. If you wish to use any version besides .NET 7.0, please uncomment the line containing that framework and comment out the line containing the .NET 7.0 framework (# to comment).

## Run Powershell Script
Parameters:
- -file: a test file
 	
  This is convenient when you want to get results for a specific file in one run instead of testing multiple files in interactive mode.

- -license (optional): a license string to test the Presort Object
- -quiet (optional): add to the command if you do not want to get any console output from the Melissa Updater

When you have modified the script to match your data location, let's run the script. There are two modes:
- Interactive 

    The script will prompt the user for a file path, then use the provided file to test Presort Object.  For example:
    ```
    $ .\MelissaPresortObjectWindowsDotnet.ps1
    ```
    For quiet mode:
    ```
    $ .\MelissaPresortObjectWindowsDotnet.ps1 -quiet
    ```
- Command Line 

    You can pass a file path in ```-file``` parameter and a license string in ```-license``` parameter to test Presort Object. For example:
    ```
    $ .\MelissaPresortObjectWindowsDotnet.ps1 -file ".\MelissaPresortObjectWindowsDotnet\MelissaPresortObjectSampleInput.csv" 
    $ .\MelissaPresortObjectWindowsDotnet.ps1 -file ".\MelissaPresortObjectWindowsDotnet\MelissaPresortObjectSampleInput.csv" -license "<your_license_string>"
    ```
    For quiet mode:
    ```
    $ .\MelissaPresortObjectWindowsDotnet.ps1 -file ".\MelissaPresortObjectWindowsDotnet\MelissaPresortObjectSampleInput.csv" -quiet
    $ .\MelissaPresortObjectWindowsDotnet.ps1 -file ".\MelissaPresortObjectWindowsDotnet\MelissaPresortObjectSampleInput.csv" -license "<your_license_string>" -quiet
    ```
This is the expected output from a successful setup for interactive mode:

![alt text](/screenshots/output.png)

    
## Troubleshooting

Troubleshooting for errors found while running your program.

### C# Errors:

| Error      | Description |
| ----------- | ----------- |
| ErrorRequiredFileNotFound      | Program is missing a required file. Please check your Data folder and refer to the list of required files above. If you are unable to obtain all required files through the Melissa Updater, please contact technical support below. |
| ErrorDatabaseExpired   | .db file(s) are expired. Please make sure you are downloading and using the latest release version. (If using the Melissa Updater, check powershell script for '$RELEASE_VERSION = {version}'  and change the release version if you are using an out of date release).     |
| ErrorFoundOldFile   | File(s) are out of date. Please make sure you are downloading and using the latest release version. (If using the Melissa Updater, check powershell script for '$RELEASE_VERSION = {version}'  and change the release version if you are using an out of date release).    |
| ErrorLicenseExpired   | Expired license string. Please contact technical support below. |


## Contact Us

For free technical support, please call us at 800-MELISSA ext. 4 (800-635-4772 ext. 4) or email us at tech@melissa.com.

To purchase this product, contact the Melissa sales department at 800-MELISSA ext. 3 (800-635-4772 ext. 3).
