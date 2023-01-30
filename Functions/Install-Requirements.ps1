function Install-Requirements{

    #Install the NuGet package
    Write-Host "Installing NuGet" -ForegroundColor Green
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.208 -Scope CurrentUser -Force -Confirm:$false | Out-Null
    
    if((Get-PSRepository -Name "PSGallery").InstallationPolicy -ne "Trusted"){
        Write-Host "Setting PSGallery repo to Trusted..."
        Set-PSRepository -name  "PSGallery" -InstallationPolicy Trusted
    }

    if(!(Get-Module -ListAvailable -name ImportExcel)){
        Write-Host "Instaling ImportExcel"
        Install-Module ImportExcel -Scope CurrentUser -Confirm:$false #https://github.com/dfinke/ImportExcel
    }
}