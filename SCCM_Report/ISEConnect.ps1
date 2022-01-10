$SCCMsitecode = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\SMS\DP" 
# Site configuration 

$SiteCode = $SCCMsitecode.SiteCode # Site code  

$ProviderMachineName = $SCCMsitecode.Server # SMS Provider machine name 

# Customizations 

$initParams = @{} 

# Import the ConfigurationManager.psd1 module  
if((Get-Module ConfigurationManager) -eq $null) { 

    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams  

} 

# Connect to the site's drive if it is not already present 

if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) { 

    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams 

} 
# Set the current location to be the site code. 

Set-Location "$($SiteCode):\" @initParams 

 