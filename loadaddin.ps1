<# 
            " Satnaam WaheGuru Ji"     
             
            Author  :  Aman Dhally 
            E-Mail  :  amandhally@gmail.com 
            website :  www.amandhally.net 
            twitter : https://twitter.com/#!/AmanDhally 
            facebook: http://www.facebook.com/groups/254997707860848/ 
            Linkedin: http://www.linkedin.com/profile/view?id=23651495 
 
            Date    : 31-August-2012 
            File    : Office_2010_Trusted_Locations 
            Purpose : Set trusted locations for Office 2010 Templates 
             
            Version : 1 
 
            my Spider runned Away :(  
 
 
#> 
 
#-------------------------------------------------- 
# Helpful Weblinks for script 
#-------------------------------------------------- 
 
# http://technet.microsoft.com/en-us/library/cc179039.aspx#implementlocations 
# http://technet.microsoft.com/en-us/library/dd315394.aspx 
# 
 
 
 
#-------------------------------------------------- 
# Global Variables Section 
#-------------------------------------------------- 
 
#Date 
# 
$setDATE = (Get-Date).ToString() 
# 
# trusted Location to Add Addd Your templates Path or location path 
# 
$trustlocation = '\\localhost\sharedtestfolder'    # <---- Change this as per your requirement 
# 
# Note I am using Location99 as a KEY so that i can test it and i know that is created by me. 
# 
$guid = (New-Guid).ToString()
$guidstr = "{$guid}"
$regWORD = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{$guid}"
# HKCU:\Software\Microsoft\Office\14.0\Word\Security\Trusted Locations\Location99' 
# Testing if Location 9 Key Exists in word. 
$testWORDPAth = Test-Path $regWORD 
# 
# 
$regEXCEL = 'HKCU:\Software\Microsoft\Office\14.0\Excel\Security\Trusted Locations\Location99' 
# Testing if Location 9 Key Exists in Excel registry Key. 
$testEXCELPATH = Test-Path $regEXCEL 
# 
# 
$regPPT = 'HKCU:\Software\Microsoft\Office\14.0\Powerpoint\Security\Trusted Locations\Location99' 
# Testing if Location 9 Key Exists in PowerPoint. 
$testPPTPATH = Test-Path $regPPT 
# 
# 
$setDesription = "Trusted location for YOUR COMPANY in-house Templates"   #<----------- You can change the Description 
# 
# Registry Keys Need to be Create 
# 
# AllowSubfolders = 1 
# Date = todays Date 
# Description = setDescription 
# Path = location 
 
#-------------------------------------------------- 
# Global Variables Section Ends Here |  
#-------------------------------------------------- 
 
# You can comment this if you want :) this is just a quote. 
write-host "Love me when I least deserve it, because that’s when I really need it." -ForegroundColor 'Yellow' 
Write-Host "                                           - Sri Guru Granth Sahib Ji " -ForegroundColor 'Green' 
"`n" 
# 
 
#-------------------------------------------------- 
# Script Section 
#-------------------------------------------------- 
 
# 
# For Word 
# 
 
if ( $testWORDPAth -eq $false) { 
     
    Write-Host "==> Adding Trusted Location for Microsoft Word" -ForegroundColor 'Green' 
    New-Item -Path $regWORD -type Directory -Force | Out-Null 
    #New-ItemProperty -Path $regWORD -Name "Date" -Value $setDATE | Out-Null 
    New-ItemProperty -Path $regWORD -Name "Url" -Value $trustlocation | Out-Null 
    New-ItemProperty -Path $regWORD -Name "Id" -Value $guidstr | Out-Null 
    New-ItemProperty -Path $regWORD -Type DWord  -Name "Flags" -Value 0 | Out-Null 
         
    } else {  
    Write-Host "==>  Word Templated are already added in Trusted Locations."  -ForegroundColor 'Green' 
    } 

#-------------------------------------------------- 
# Script Section Ends Here || 
#-------------------------------------------------- 
 
############### A m a n   D h a l l y .. 