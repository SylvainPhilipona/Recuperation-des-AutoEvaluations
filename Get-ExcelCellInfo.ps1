function Get-ExcelCellInfo {
    
    #Load the required functions
    . .\Manage-Functions.ps1
    . Manage-Functions

    # Set-PSRepository -name  "PSGallery" -InstallationPolicy Trusted
    # Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser
    # Install-Module ImportExcel -Scope CurrentUser -Confirm:$false #https://github.com/dfinke/ImportExcel


    # #Get all excel files from the repo
    # $AllFiles = Get-ChildItem -Path ".\DataFiles\*.xlsx"
    

    # foreach($file in $AllFiles){
        
    #     Import-Excel -Path ".\DataFiles\$($file.name)"

    #     "------------------------------------------------------"
    # }



    # https://powershell.one/tricks/parsing/excel


    $students = @(
        "Joca Bolli",
        "Nolan Praz",
        "Dorian Capelli"
    )


    foreach($student in  $students){

        #Import the model file
        $excel = New-Object -ComObject excel.application
        $excel.visible = $false
        $workbook = $excel.Workbooks.Open("$($PSScriptRoot)\DataFiles\Modele_AutoEval.xlsx")
        $Sheet1 = $workbook.worksheets.item(1)

        #Replace the cells with the incoming datas
        $Sheet1.cells.find("[NAME]") = $student


        $workbook.Saveas("$($PSScriptRoot)\Output\AutoEval-$($student.replace(' ', '-')).xlsx")

        

        $excel.workbooks.Close()

    }


   





























# # "test" | Out-File "./DataFiles/Test.txt" -Verbose

    # UnZip-File "./DataFiles/Test.docx" "./DataFiles/Temp"

    # Remove-Item "./DataFiles/Ninino.docx" -Force -ErrorAction SilentlyContinue
    # Zip-File "./DataFiles/Ninino.docx" "./DataFiles/Temp"





    #Unloading the functions
    . Manage-Functions -remove
}