param (
    [string]$ConfigsPath, # = "$($PSScriptRoot)\DataFiles\01-configs-auto-eval.xlsx",
    [string]$ModelPath # = "$($PSScriptRoot)\DataFiles\02-modele-auto-eval.xlsx"
)

Start-Transcript -Path "./Output/Output.log" -Append -Force

#Install all requirements for the script to run
.\Install-Requirements.ps1

#Load the data file
$data = Import-LocalizedData -BaseDirectory ./DataFiles -FileName Inputs.psd1
Write-Host "Loading inputs file" -ForegroundColor Green

#Verify if the config and model file exists
$testPaths = ./Test-Paths.ps1 -paths $ConfigsPath, $ModelPath
if(!$testPaths.count -eq 0){
    #Dispaly the missing paths
    $errorMessage = "Les fichiers suivants n'existent pas : `n`r"
    foreach($path in $testPaths){
        $errorMessage += " $path `n`r"
    }
    .\Stop-Program.ps1 -errorMessage $errorMessage
}

#Import the configs and students inputs
$Configs = (Import-Excel -Path $ConfigsPath -WorksheetName $data.ConfigFile.ConfigSheet)
$Students = (Import-Excel -Path $ConfigsPath -WorksheetName $data.ConfigFile.StudentsSheet)
Write-Host "Loading configs file" -ForegroundColor Green

# Check that the Config file contains all the required inputs.
# The inputs are specified in the Inputs.psd1 file
foreach($input in $data.RequiredInputs.GetEnumerator()){
    if(!($Configs.Champs.contains($input.Value))){
        .\Stop-Program.ps1 -errorMessage "Il manque un champ wesh"
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
        Stop-Program.ps1 -errorMessage "Excel n'est pas installé. Veuillez l'installer et recomencer !"
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

    #Set the sheet name
    $Sheet1.Name = "$($student.Prenom) $($student.Nom)"

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

Stop-Transcript