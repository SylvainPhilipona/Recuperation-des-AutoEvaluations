function Create-AutoEvals {

    param (
        [string]$ConfigsPath = "$($PSScriptRoot)\DataFiles\01-configs-auto-eval.xlsx",
        [string]$ModelPath = "$($PSScriptRoot)\DataFiles\02-modele-auto-eval.xlsx"
    )

    Start-Transcript -Path "./Output/Output.log" -Append -Force
    
    #Load the required functions
    . .\Manage-Functions.ps1
    . Manage-Functions

    #Install all requirements for the script to run
    Install-Requirements

    #Load the data file
    $data = Import-LocalizedData -BaseDirectory ./DataFiles -FileName Inputs.psd1
    Write-Host "Loading inputs file" -ForegroundColor Green

    #Import the configs and students inputs
    $Configs = (Import-Excel -Path $ConfigsPath -WorksheetName $data.ConfigFile.ConfigSheet)
    $Students = (Import-Excel -Path $ConfigsPath -WorksheetName $data.ConfigFile.StudentsSheet)
    Write-Host "Loading configs file" -ForegroundColor Green

    # Check that the Config file contains all the required inputs.
    # The inputs are specified in the Inputs.psd1 file
    foreach($input in $data.RequiredInputs.GetEnumerator()){
        if(!($Configs.Champs.contains($input.Value))){
            Stop-Program -errorMessage "Il manque un champ wesh"
        }
    }

    #Convert Configs to ConfigsHash table
    $ConfigsHash = @{}
    foreach($config in $Configs){
        $ConfigsHash.Add($config.Champs, $config.Valeurs)
    }

    foreach($student in  $students){

        Write-Host "Creating $($student.Prenom) $($student.Nom) AutoEval"

        #Import the model file
        try{
            $excel = New-Object -ComObject excel.application
        }
        catch [System.Runtime.InteropServices.COMException] {
           Stop-Program -errorMessage "Excel n'est pas installé. Veuillez l'installer et recomencer !"
        }
        $excel = New-Object -ComObject excel.application
        $excel.visible = $false
        $workbook = $excel.Workbooks.Open($ModelPath)
        $Sheet1 = $workbook.worksheets.item(1)

        #Unprotect the sheet
        $Sheet1.Unprotect()
        
        #Replace the cells with the configs datas
        $Sheet1.cells.find("[NAME]") = "$($student.Prenom) $($student.Nom)"
        $Sheet1.cells.find("[CLASSE]") = $ConfigsHash[$data.RequiredInputs.CLASSE]
        $Sheet1.cells.find("[TEACHER]") = $ConfigsHash[$data.RequiredInputs.TEACHER]
        $Sheet1.cells.find("[PROJECTNAME]") = $ConfigsHash[$data.RequiredInputs.PROJECTNAME]
        $Sheet1.cells.find("[NBWEEKS]") = $ConfigsHash[$data.RequiredInputs.NBWEEKS]
        $Sheet1.cells.find("[DATES]") = "$($ConfigsHash[$data.RequiredInputs.DATES].ToString("yyyy/MM/dd"))-$($ConfigsHash["Date fin"].ToString("yyyy/MM/dd"))"

        #Protect the sheet
        $Sheet1.Protect()
        
        #Save the new file as the student name (Overwrite if the file exists)
        $filename = "$($PSScriptRoot)\Output\AutoEval-$($student.Prenom + "-" + $student.Nom).xlsx"
        Remove-Item -Path $filename -Force -Confirm:$false -ErrorAction SilentlyContinue
        $workbook.Saveas($filename)
        Write-Host "    --> Saving $filename"

        
        #Close the object
        $excel.workbooks.Close()
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }

    #Unloading the functions
    . Manage-Functions -remove

    Stop-Transcript
}