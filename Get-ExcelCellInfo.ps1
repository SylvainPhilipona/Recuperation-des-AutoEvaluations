function Get-ExcelCellInfo {

    param (
        [string]$ModelPath = "$($PSScriptRoot)\DataFiles\Modele_AutoEval.xlsx",
        [string]$InputsPath = "$($PSScriptRoot)\DataFiles\Inputs_AutoEval.xlsx"
    )
    
    #Load the required functions
    . .\Manage-Functions.ps1
    . Manage-Functions


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


    #Import the inputs
    $Inputs = (Import-Excel -Path $InputsPath -WorksheetName "inputs")
    $Students = (Import-Excel -Path $InputsPath -WorksheetName "students")

    #Convert to hash table
    $hash = @{}
    foreach($input in $Inputs){
        $hash.Add($input.Champs, $input.Valeurs)
    }

    foreach($student in  $students){

        #Import the model file
        $excel = New-Object -ComObject excel.application
        $excel.visible = $false
        $workbook = $excel.Workbooks.Open("$($PSScriptRoot)\Modeles\Modele_AutoEval.xlsx")
        $Sheet1 = $workbook.worksheets.item(1)

        #Unprotect the sheet
        $Sheet1.Unprotect()  

        #Replace the cells with the incoming datas
        $Sheet1.cells.find("[NAME]") = "$($student.Prenom) $($student.Nom)"
        $Sheet1.cells.find("[CLASSE]") = $hash["Classe"]
        $Sheet1.cells.find("[TEACHER]") = $hash["Enseignant"]
        $Sheet1.cells.find("[PROJECTNAME]") = $hash["Nom du projet"]
        $Sheet1.cells.find("[NBWEEKS]") = $hash["Nbr de semaines"]
        $Sheet1.cells.find("[DATES]") = "$($hash["Date debut"].ToString("yyyy/MM/dd"))-$($hash["Date fin"].ToString("yyyy/MM/dd"))"

        #Protect the sheet
        $Sheet1.Protect()
        
        #Save the file
        $workbook.Saveas("$($PSScriptRoot)\Output\AutoEval-$($student.Prenom + "-" + $student.Nom).xlsx")

        
        #Close the object
        $excel.workbooks.Close()
    }


   




























# # "test" | Out-File "./DataFiles/Test.txt" -Verbose

    # UnZip-File "./DataFiles/Test.docx" "./DataFiles/Temp"

    # Remove-Item "./DataFiles/Ninino.docx" -Force -ErrorAction SilentlyContinue
    # Zip-File "./DataFiles/Ninino.docx" "./DataFiles/Temp"





    #Unloading the functions
    . Manage-Functions -remove
}