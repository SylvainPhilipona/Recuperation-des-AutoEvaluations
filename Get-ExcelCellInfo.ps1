function Get-ExcelCellInfo {

    param (
        [string]$InputsPath = "$($PSScriptRoot)\DataFiles\01-configs-auto-eval.xlsx",
        [string]$ModelPath = "$($PSScriptRoot)\DataFiles\02-modele-auto-eval.xlsx"
    )
    
    #Load the required functions
    . .\Manage-Functions.ps1
    . Manage-Functions

    #Load the data file
    $data = Import-LocalizedData -BaseDirectory ./DataFiles -FileName Inputs.psd1

    #Install all requirements for the script to run
    Install-Requirements

    #Import the inputs
    $Inputs = (Import-Excel -Path $InputsPath -WorksheetName "inputs")
    $Students = (Import-Excel -Path $InputsPath -WorksheetName "students")

    #Check that the Config file contains all the required inputs
    foreach($input in $data.RequiredInputs.GetEnumerator()){
        if(!($Inputs.Champs.contains($input.Value))){
            #Unloading the functions
            . Manage-Functions -remove

            throw "Il manque un champ wesh"
        }
    }

    #Convert Inputs to hash table
    $hash = @{}
    foreach($input in $Inputs){
        $hash.Add($input.Champs, $input.Valeurs)
    }

    foreach($student in  $students){

        #Import the model file
        $excel = New-Object -ComObject excel.application
        $excel.visible = $false
        $workbook = $excel.Workbooks.Open($ModelPath)
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

        
        #Close the object https://social.technet.microsoft.com/Forums/office/en-US/e5d8594b-b14b-4a54-913c-61089b5d9ab4/release-or-delete-a-com-object-from-powershell?forum=winserverpowershell
        $excel.workbooks.Close()
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    }


   




























# # "test" | Out-File "./DataFiles/Test.txt" -Verbose

    # UnZip-File "./DataFiles/Test.docx" "./DataFiles/Temp"

    # Remove-Item "./DataFiles/Ninino.docx" -Force -ErrorAction SilentlyContinue
    # Zip-File "./DataFiles/Ninino.docx" "./DataFiles/Temp"





    #Unloading the functions
    . Manage-Functions -remove
}