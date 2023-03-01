<#
.NOTES
    *****************************************************************************
    ETML
    Name:	Create-AutoEvals.ps1
    Author:	Sylvain Philipona
    Date:	23.02.2023
 	*****************************************************************************
    Modifications
 	Date  : 27.02.2023
 	Author: Sylvain Philipona
 	Reason: Modification des messages d'erreurs et amélioration de la gestion des erreurs
 	*****************************************************************************
.SYNOPSIS
    Création des auto-évaluations selon une liste d'élèves 
 	
.DESCRIPTION
    Ce script crée les auto-évaluations d'une liste d'élèves selon un modèle
    Un fichier est crée par élève, avec les informations fournies. (Nom projet, nom prof, nom élève...)
    Ces fichiers sont ensuite envoyés aux éleves pour remplisage
  	
.PARAMETER ConfigsPath
    Ceci est le chemin du fichier XLSX contenant Les configurations et la liste des élèves
 	
.PARAMETER ModelPath
    Ceci est le chemin du fichier XLSX Modèle d'auto-évaluations
 	
.PARAMETER OutputPath
    Ceci est le chemin où les fichiers des auto-évaluations personels des élèves vont êtres crées

.OUTPUTS
	- Les fichiers des auto-évaluations personels des élèves (se crée dans le même répertoire)
	- Un fichier log avec les actions éffectuées

.EXAMPLE
    .\Create-AutoEvals.ps1 -ConfigsPath ".\DataFiles\01-configs-auto-eval.xlsx" -ModelPath ".\DataFiles\02-modele-auto-eval.xlsx" -OutputPath ".\Output"

    Installation de NuGet
    Chargement du fichier d'inputs
    Chargement du fichier de configurations
    Création de Dorian Capelli AutoEval
        --> Enregistrement de E:\09-P_Appro\PS-Eval\Scripts\Output\AutoEval-Dorian-Capelli.xlsx
    Création de Nolan Praz AutoEval
        --> Enregistrement de E:\09-P_Appro\PS-Eval\Scripts\Output\AutoEval-Nolan-Praz.xlsx
    Création de Joca Bolli AutoEval
        --> Enregistrement de E:\09-P_Appro\PS-Eval\Scripts\Output\AutoEval-Joca-Bolli.xlsx
    Création de Sayeh Younes AutoEval
        --> Enregistrement de E:\09-P_Appro\PS-Eval\Scripts\Output\AutoEval-Sayeh-Younes.xlsx
    Création de Sylvain Philipona AutoEval
        --> Enregistrement de E:\09-P_Appro\PS-Eval\Scripts\Output\AutoEval-Sylvain-Philipona.xlsx
 	
.LINK
    Install-Requirements.ps1
    Test-Paths.ps1
    Stop-Program.ps1
#>

param (
    [string]$ConfigsPath,
    [string]$ModelPath,
    [string]$OutputPath
)

if(!(Test-Path -Path $OutputPath -PathType Container)){
    New-Item -Path $OutputPath -ItemType Directory -Force -Confirm:$false
}

Start-Transcript -Path "$OutputPath/Output.log" -Append -Force

#Install all requirements for the script to run
.\Install-Requirements.ps1

#Load the data file
$data = Import-LocalizedData -BaseDirectory "$($PSScriptRoot)\01-config" -FileName Inputs.psd1
Write-Host "Chargement du fichier d'inputs" -ForegroundColor Green

#Verify if the config and model files exists
$testPaths = .\Test-Paths.ps1 -paths $ConfigsPath, $ModelPath
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
Write-Host "Chargement du fichier de configurations" -ForegroundColor Green

# Check that the Config file contains all the required inputs.
# The inputs are specified in the Inputs.psd1 file
foreach($input in $data.RequiredInputs.GetEnumerator()){
    if(!($Configs.Champs.contains($input.Value))){
        .\Stop-Program.ps1 -errorMessage "Veuillez remplir tout les champs de configurations"
    }
}

#Convert Configs to ConfigsHash table
$ConfigsHash = @{}
foreach($config in $Configs){
    $ConfigsHash.Add($config.Champs, $config.Valeurs)
}

foreach($student in  $students){

    Write-Host "Création de $($student.Prenom) $($student.Nom) AutoEval"

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
    $Sheet1.cells.find("[DATES]") = "$($ConfigsHash[$data.RequiredInputs.DATES].ToString("yyyy/MM/dd"))-$($ConfigsHash[$data.RequiredInputs.DATEEND].ToString("yyyy/MM/dd"))"

    #Set the sheet name
    $Sheet1.Name = "$($student.Prenom) $($student.Nom)"

    #Protect the sheet
    $Sheet1.Protect()
    
    #Save the new file as the student name (Overwrite if the file exists)
    $filename = "$OutputPath\AutoEval-$($student.Prenom + "-" + $student.Nom).xlsx"
    Remove-Item -Path $filename -Force -Confirm:$false -ErrorAction SilentlyContinue
    $workbook.Saveas($filename)
    Write-Host "    --> Enregistrement de $filename"

    
    #Close the object
    $excel.workbooks.Close()
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

Stop-Transcript