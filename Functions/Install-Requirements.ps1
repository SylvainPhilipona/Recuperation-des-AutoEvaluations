function Install-Requirements{

    #Install the NuGet package
    Write-Verbose "Installing NuGet..."
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.208 -Scope CurrentUser -Force -Confirm:$false
    
    if((Get-PSRepository -Name "PSGallery").InstallationPolicy -ne "Trusted"){
        Write-Verbose "Setting PSGallery repo to Trusted..."
        Set-PSRepository -name  "PSGallery" -InstallationPolicy Trusted
    }

    if(!(Get-Module -ListAvailable -name ImportExcel)){
        Write-Verbose "Instaling ImportExcel"
        Install-Module ImportExcel -Scope CurrentUser -Confirm:$false #https://github.com/dfinke/ImportExcel
    }
}