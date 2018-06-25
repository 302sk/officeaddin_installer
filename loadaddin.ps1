<# 
            " Sideload office add-in"     
             
            Author  :  shnkn
            E-Mail  :  302sk@163.com 
          
 
            Date    : 2018-06-18 
            File    : sideload-office-addin 
            Purpose : sideload office addin 
             
            Version : 1 
 
            
#> 
 

#-------------------------------------------------- 
# Global Variables Section 
#-------------------------------------------------- 
 

#Create share folder for save add-in xml file
$trustFolder = 'c:\testfolder'
if((Test-Path $trustFolder) -eq $false){
    Write-Host "==>创建add-in xml目录"
    mkdir c:\testfolder
}
else{
   Write-Host "==>add-in xml目录已存在"
}

Write-Host "==>拷贝add-in xml文件"
copy .\addinexample.xml C:\testfolder -Force # copy add-in xml file to testfolder

if (!(Get-SmbShare -Name sharedtestfolder -ea 0)) {
    net share sharedtestfolder=c:\testfolder  #require administrator permission
    #Write-Host "==>目录共享成功"
}
else {
    Write-Host "==>共享目录已存在"
}

#Date 
# 
$setDATE = (Get-Date).ToString() 
# 
# trusted Location to save add-in xml file
# 
$trustlocation = '\\localhost\sharedtestfolder'    # <---- Change this as per your requirement 
# 
# Note I am using static guid for key name, and this guid could probably never conflict with existed key
# 
$guid = "396f6026-26d4-4f20-bf58-47d789e12345"  #(New-Guid).ToString()
$guidstr = "{$guid}"
$regWORD = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{$guid}"

# Testing if Location guid Key Exists in word. 
$testWORDPAth = Test-Path $regWORD 
# 
# 
# 
$setDesription = "Trusted location for the office add-in xml file"   #<----------- You can change the Description 
# 
 
#-------------------------------------------------- 
# Script Section 
#-------------------------------------------------- 
 
# 
# For Word 
# 

# if that key exist then remove it
if ($testWORDPAth -eq $true){
    Remove-Item $regWORD -Recurse
}
 
# create registry key for trusted add-in folder
Write-Host "==> Adding Trusted Location for Microsoft Word add-in" -ForegroundColor 'Green' 
New-Item -Path $regWORD -type Directory -Force | Out-Null 
#New-ItemProperty -Path $regWORD -Name "Date" -Value $setDATE | Out-Null 
New-ItemProperty -Path $regWORD -Name "Url" -Value $trustlocation | Out-Null 
New-ItemProperty -Path $regWORD -Name "Id" -Value $guidstr | Out-Null 
New-ItemProperty -Path $regWORD -Type DWord  -Name "Flags" -Value 1 | Out-Null 
         
    
# pause for check debug info
cmd /c pause | out-null

#-------------------------------------------------- 
# Script Section Ends Here || 
#-------------------------------------------------- 
 
############### shnkn 