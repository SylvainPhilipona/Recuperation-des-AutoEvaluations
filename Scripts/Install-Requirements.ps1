#Install the NuGet package
Write-Host "Installation de NuGet" -ForegroundColor Green
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.208 -Scope CurrentUser -Force -Confirm:$false | Out-Null

if((Get-PSRepository -Name "PSGallery").InstallationPolicy -ne "Trusted"){
    Write-Host "Ajout du répertoire PSGallery en répertoire de confiance" -ForegroundColor Green
    Set-PSRepository -name  "PSGallery" -InstallationPolicy Trusted
}

if(!(Get-Module -ListAvailable -name ImportExcel)){
    Write-Host "Installation de ImportExcel" -ForegroundColor Green
    Install-Module ImportExcel -Scope CurrentUser -Confirm:$false #https://github.com/dfinke/ImportExcel
}