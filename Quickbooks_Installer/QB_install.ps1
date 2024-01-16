<#

    .SYNOPSIS
        Installs Quickbooks Desktop

    .DESCRIPTION
        Installs all specified versions of Quickbooks Desktop that exist in the .xlsx file under the "company" tab

    .NOTES
        Author: Roy Smith
        Origin Credit: Aaron J. Stevenson - https://github.com/wise-io/scripts/blob/main/scripts/InstallQuickBooks.ps1

#>

########################
# Variable Declaration #
########################
$Company = "company" #Worksheet name on QuickbooksKeys.xlsx where company keys are stored


function confirm-SystemCheck{
    $currentUserSID = ([System.Security.Principal.WindowsIdentity]::GetCurrent()).User.Value
    if($currentUserSID -eq "S-1-5-18"){
        Write-Warning "This script cannot be ran as SYSTEM. Please run as Admin"

        exit 1
    }
}

function install-XPSDocumentWriter{
    $XPSFeature = Get-WindowsOptionalFeature -Online | Where-Object {$_.FeatureName -ieq "Printing-XPSServices-Features"}
    if($XPSFeature.State -eq "Disabled"){
        try{
            Write-Output "Installing required PDF components (MS XPS Document Writer)...."
            Enable-WindowsOptionalFeature -Online -FeatureName "Printing-XPSServices-Features" -All -NoRestart | Out-Null
        }
        catch{
            Write-Warning "Unable to install MS XPS Document Writer feature"
            Write-Warning "$_"
        }
    }
    elseif($XPSFeature.State -eq "Enabled"){
        Write-Output "MS XPS Document Writer is already enabled in Optional Features"
    }
}

#imports all of the clients data from the .xlxs file into PS
function import-QBkeys{
    $QBKeys = Join-Path $PSScriptRoot "QuickbooksKeys.xlsx"
    
    if(Get-Module -Name "ImportExcel" -ListAvailable){
        Write-Output "The Module ImportExcel is already installed on session"
    }
    else{
        Install-Module -Name ImportExcel -Force -AllowClobber
    }

    $global:wwlData = Import-Excel -Path $QBKeys -WorksheetName $Company
}

#converts the imported data from the .xlsx file into usable variables throughout the script
function convert-QBkeys{ 
    $companyQBEditions = @()
    

    foreach($item in $wwlData){
        $edition = $item.Edition
        $editionYear = $item.Year
        $editionMerged = "$edition$editionYear"
        $productID = $item.'Product #'
        $licenseKey = $item.'License #'
        $downloadURL = $item.'Download URL'

        $editionHash = @{
            Edition = $editionMerged
            ProductID = $productID
            LicenseKey = $licenseKey
            URL = $downloadURL
        }
        
        $companyQBEditions += $editionHash
    }

    
    $global:uniqueEditions = $companyQBEditions.Edition | Sort-Object -Unique
    foreach($uniqueEd in $uniqueEditions){
        New-Variable -Name $uniqueEd -Value @() -Scope Global
    }

    #this ensures that you do not have any duplicate productIDs, LicenseKeys or URL's inside the edition variable
    foreach($item in $companyQBEditions){
        $editionVar = $item.Edition
        $tempVar = Get-Variable -Name $editionVar
        if($tempVar.Value -notcontains $item.ProductID){
            $tempVar.Value += @($item.ProductID)
        }
        if($tempVar.Value -notcontains $item.URL){
            $tempVar.Value += @($item.URL)
        }
        if($tempVar.Value -notcontains $item.LicenseKey){
            $tempVar.Value += @($item.LicenseKey)
        }        
                                
    }
    
}

#This checks for the Temp folder on C: and then adds the QBInstallers directory inside of C:\Temp\
function add-InstallerDir{
    $pathVar = @("C:\temp","C:\temp\QBInstallers")
    $scripthPath = Join-Path $PSScriptRoot "Installers\$item.exe"

    foreach($path in $pathVar){
        if(!(Test-Path $path -PathType Container)){
            New-Item -ItemType Directory -Path $path
        }
    }
    #this checks to see if you have a copy of the installer.exe, if not then it downloads to the script path and then copies it to the C:\Temp\QBinstallers
    foreach($item in $uniqueEditions){
        $scripthPath = Join-Path $PSScriptRoot "Installers\$item.exe"
        if(Test-Path $scripthPath){
            $checkPath = Join-Path $pathVar[1] "$item.exe"
            if(!(Test-Path $checkPath)){
                Copy-Item $scripthPath $pathVar[1]
                Write-Output "$item.exe copied to "$pathVar[1]
            }
            else{
                Write-Output "A copy of $item.exe is already in "$pathVar[1]
            }
        }
        else{
            $tempVar = Get-Variable -Name $item
            $downloadURL = $tempVar.Value[1]
            try{Invoke-WebRequest -Uri $downloadURL -OutFile $scripthPath}
            catch{Write-Warning $_.Exception}
        }
    }
}

#this installs quickbooks
function install-Quckbooks{
    $installerPath = "C:\temp\QBInstallers"
    foreach($item in $uniqueEditions){
        $installer = Join-Path $installerPath "$item.exe"
        $editionVar = Get-Variable -Name $item
        $productKey = $editionVar.Value[0]
        #Some clients have more than 1 key per edition, this picks a random key and uses that for installation
        $countKeys = $editionVar.Value.Count 
        $getRandomNumber = Get-Random -Minimum 2 -Maximum $countKeys
        $licenseKey = $editionVar.Value[$getRandomNumber]
        
        try{
            Write-Output "Installing $item"
            Start-Process -Wait -NoNewWindow -FilePath $installer -ArgumentList "-s -a QBMIGRATOR=1 MSICOMMAND=/s QB_PRODUCTNUM=$($productKey) QB_LICENSENUM=$licenseKey"
            Write-Output "$item has been installed"
        }
        catch{
            Write-Warning "Error Installing $item"
            Write-Warning $_
        }
    }
}

import-QBkeys
convert-QBkeys
add-InstallerDir
install-Quckbooks

#removes the C:\Temp\QBInstallers directory after all versions have been installed
Remove-Item -Path C:\temp\QBInstallers -Recurse -Force


