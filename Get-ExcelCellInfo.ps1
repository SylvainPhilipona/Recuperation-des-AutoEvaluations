function Get-ExcelCellInfo {

    param (
        [string]$ModelPath = "$($PSScriptRoot)\DataFiles\Modele_AutoEval.xlsx",
        [string]$InputsPath = "$($PSScriptRoot)\DataFiles\Inputs_AutoEval.xlsx"
    )
    
    #Load the required functions
    . .\Manage-Functions.ps1
    . Manage-Functions

    # Set-PSRepository -name  "PSGallery" -InstallationPolicy Trusted
    # Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser
    # Install-Module ImportExcel -Scope CurrentUser -Confirm:$false #https://github.com/dfinke/ImportExcel

    # https://powershell.one/tricks/parsing/excel


    #Import the inputs
    $Inputs = (Import-Excel -Path $InputsPath -WorksheetName "inputs") #| Select-Object "Champs","Valeurs"
    $Students = (Import-Excel -Path $InputsPath -WorksheetName "students") #| Select-Object "Nom","Prenom"

    foreach($student in  $students){

        #Import the model file
        $excel = New-Object -ComObject excel.application
        $excel.visible = $false
        $workbook = $excel.Workbooks.Open("$($PSScriptRoot)\DataFiles\Modele_AutoEval.xlsx")
        $Sheet1 = $workbook.worksheets.item(1)

        #Replace the cells with the incoming datas
        $Sheet1.cells.find("[NAME]") = "$($student.Prenom) $($student.Nom)"

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