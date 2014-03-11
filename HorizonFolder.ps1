<#--------------------------------------------------------------------------------- 
This script will discover your VMware Horizon Data Folder and add it as a save
location in Microsoft Office.

Tested on Windows 7 / Office 2013
Should work with Windows 7
Should work with Windows 8

Office 2010 and earlier does not support Cloud Storage intergration

v1.0    11 March 2014:   First!

Cobbled together by Josh Townsend

All Code provided as is and used at your own risk.
#---------------------------------------------------------------------------------#>

#requires -Version 3.0

$Shell = New-Object -ComObject Shell.Application
$Desktop = $Shell.NameSpace(0X0)
$WshShell = New-Object -comObject WScript.Shell

#Get UserName and check default location of folder
[string]$username = [System.Environment]::UserName
[string]$horizondir = get-itemproperty -path "HKCU:\Software\VMware, Inc.\Horizon Data\" |% {$_.Folder}

#Ask for input if folder can't be found in default location
    If(!(Test-Path -Path "$($horizondir)"))
    {
        [string]$horizondir = Read-Host 'Could not find Horizon. Enter path to folder manually: '
    }
    
#Add registry values
New-Item -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\8583bc37-7a65-4c06-903c-fef75f3f08ed'
New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\8583bc37-7a65-4c06-903c-fef75f3f08ed' -Name DisplayName -PropertyType String -Value 'Horizon'
New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\8583bc37-7a65-4c06-903c-fef75f3f08ed' -Name Description -PropertyType String -Value 'VMware® Horizon Workspace™ provides an easy way to access apps and files on any device, while enabling IT to centrally deliver, manage and secure these assets.'
New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\8583bc37-7a65-4c06-903c-fef75f3f08ed' -Name Url48x48 -PropertyType String -Value http://vmtoday.com/wp-content/horizon/icons/Horizon_48x48.png
New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\8583bc37-7a65-4c06-903c-fef75f3f08ed' -Name LearnMoreURL -PropertyType String -Value https://www.vmware.com/
New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\8583bc37-7a65-4c06-903c-fef75f3f08ed' -Name ManageURL -PropertyType String -Value https://horizonworkspace.vmware.com/hc/login/
New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\8583bc37-7a65-4c06-903c-fef75f3f08ed' -Name LocalFolderRoot -PropertyType String -Value $horizondir

New-Item -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\8583bc37-7a65-4c06-903c-fef75f3f08ed\Thumbnails'
New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\8583bc37-7a65-4c06-903c-fef75f3f08ed\Thumbnails' -Name Url256x256 -PropertyType String -Value http://vmtoday.com/wp-content/horizon/icons/Horizon_256x256.png
New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\8583bc37-7a65-4c06-903c-fef75f3f08ed\Thumbnails' -Name Url128x128 -PropertyType String -Value http://vmtoday.com/wp-content/horizon/icons/Horizon_128x128.png
New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\8583bc37-7a65-4c06-903c-fef75f3f08ed\Thumbnails' -Name Url96x96 -PropertyType String -Value http://vmtoday.com/wp-content/horizon/icons/Horizon_96x96.png
New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\8583bc37-7a65-4c06-903c-fef75f3f08ed\Thumbnails' -Name Url64x64 -PropertyType String -Value http://vmtoday.com/wp-content/horizon/icons/Horizon_64x64.png
New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\8583bc37-7a65-4c06-903c-fef75f3f08ed\Thumbnails' -Name Url48x48 -PropertyType String -Value http://vmtoday.com/wp-content/horizon/icons/Horizon_48x48.png
New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\8583bc37-7a65-4c06-903c-fef75f3f08ed\Thumbnails' -Name Url40x40 -PropertyType String -Value http://vmtoday.com/wp-content/horizon/icons/Horizon_40x40.png
New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\8583bc37-7a65-4c06-903c-fef75f3f08ed\Thumbnails' -Name Url32x32 -PropertyType String -Value http://vmtoday.com/wp-content/horizon/icons/Horizon_32x32.png
New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\8583bc37-7a65-4c06-903c-fef75f3f08ed\Thumbnails' -Name Url24x24 -PropertyType String -Value http://vmtoday.com/wp-content/horizon/icons/Horizon_24x24.png
New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\8583bc37-7a65-4c06-903c-fef75f3f08ed\Thumbnails' -Name Url20x20 -PropertyType String -Value http://vmtoday.com/wp-content/horizon/icons/Horizon_20x20.png
New-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\Common\Cloud Storage\8583bc37-7a65-4c06-903c-fef75f3f08ed\Thumbnails' -Name Url16x16 -PropertyType String -Value http://vmtoday.com/wp-content/horizon/icons/Horizon_16x16.png