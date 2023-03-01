<#
.NOTES
    *****************************************************************************
    ETML
    Name:	Install-Requirements.ps1
    Author:	Sylvain Philipona
    Date:	23.02.2023
 	*****************************************************************************
    Modifications
 	Date  : 01.03.2023
 	Author: Sylvain Philipona
 	Reason: Changement de l'encodage du fichier en UFT-8 with BOM
 	*****************************************************************************
.SYNOPSIS
    Installe les exigences d'installation

.DESCRIPTION
    Installe les modules et packages necessaires au bon fonctionnement des scripts

.EXAMPLE
    .\Install-Requirements.ps1

    Installation de NuGet
    Ajout du répertoire PSGallery en répertoire de confiance
    Installation de ImportExcel
#>

#Install the NuGet package
Write-Host "Installation de NuGet" -ForegroundColor Green
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.208 -Scope CurrentUser -Force -Confirm:$false | Out-Null

#Set PSGallery repo to trusted -> For the ImportExcel installation
if((Get-PSRepository -Name "PSGallery").InstallationPolicy -ne "Trusted"){
    Write-Host "Ajout du répertoire PSGallery en répertoire de confiance" -ForegroundColor Green
    Set-PSRepository -name  "PSGallery" -InstallationPolicy Trusted
}

#Install the module ImportExcel
if(!(Get-Module -ListAvailable -name ImportExcel)){
    Write-Host "Installation de ImportExcel" -ForegroundColor Green
    Install-Module ImportExcel -Scope CurrentUser -Confirm:$false #https://github.com/dfinke/ImportExcel
}