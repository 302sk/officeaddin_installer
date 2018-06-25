# Office add-in installer
This project is for side-load office add-in by PowerShell script. Or you can convert this PowerShell script to EXE file to simplify the process.

# PowerShell Script Usage

## Step 0x01
Checkout this project using following command (**git required**)  

```git clone https://github.com/302sk/officeaddin_installer.git```

## Step 0x02
Right click the file "loadaddin.ps1" and choose "Run with PowerShell"  
This command will create a folder called `testfolder` in `C:\ `and make it shared with share name `sharedtestfolder`. And copy the add-in xml file `addinexample.xml` into the share folder, then add the share folder path to MSWord's `Trusted add-in Catalogs` .


# Convert PowerShell Script to EXE
Run following command in Windows PowerShell 

```.\ps2exe.ps1 -inputFile .\loadaddin.ps1 .\addin_installer.exe```

# How to use the EXE installer

1. Make sure these  two file `addin_installer.exe` and `addinexample.xml` in the same folder
2. Right click addin_installer.exe and choose "Run as administrator" option.

# How to publish the add-in installer package
1. Copy the addin_installer.exe and addinexample.xml to a folder.
2. Zip the folder and send it to users.
3. Tell users how to unzip the package and [How to use the installer]()


# NOTE
You can modify the loadaddin.ps1 file to specify the add-in xml file name,  the folder which you want the add-in  xml file store in and shared. And don't forget replace the add-in xml file with your add-in manifest xml file.