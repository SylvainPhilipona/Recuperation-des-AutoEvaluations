<#
.NOTES
    *****************************************************************************
    ETML
    Name:	Create-AutoEvals.ps1
    Author:	Sylvain Philipona
    Date:	23.02.2023
 	*****************************************************************************
    Modifications
 	Date  : 06.03.2023
 	Author: Sylvain Philipona
 	Reason: Ajout d'une notif toast
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
    - Tout les fichiers sont crées dans le répertoire défini dans 'OutputPath'

.EXAMPLE
    .\Create-AutoEvals.ps1 -ConfigsPath "./01-config\01-configs-auto-eval.xlsx" -ModelPath "./01-config\02-modele-auto-eval.xlsx" -OutputPath "./02-evaluations"

    Installation de NuGet
    Chargement du fichier de configurations
    Création de Dorian Capelli AutoEval
        --> Enregistrement de E:\09-P_Appro\PS-Eval\Scripts\02-evaluations\AutoEval-Dorian-Capelli.xlsx
    Création de Nolan Praz AutoEval
        --> Enregistrement de E:\09-P_Appro\PS-Eval\Scripts\02-evaluations\AutoEval-Nolan-Praz.xlsx
    Création de Joca Bolli AutoEval
        --> Enregistrement de E:\09-P_Appro\PS-Eval\Scripts\02-evaluations\AutoEval-Joca-Bolli.xlsx
    Création de Sayeh Younes AutoEval
        --> Enregistrement de E:\09-P_Appro\PS-Eval\Scripts\02-evaluations\AutoEval-Sayeh-Younes.xlsx
    Création de Sylvain Philipona AutoEval
        --> Enregistrement de E:\09-P_Appro\PS-Eval\Scripts\02-evaluations\AutoEval-Sylvain-Philipona.xlsx
 	
.LINK
    Get-Constants.ps1
    Install-Requirements.ps1
    Test-Paths.ps1
    Stop-Program.ps1
#>

param (
    [string]$ConfigsPath,
    [string]$ModelPath,
    [string]$OutputPath
)

#####################  Constants  #####################

$constants = .\Get-Constants.ps1
$ConfigSheet = $constants.ConfigFile.ConfigSheet
$StudentsSheet = $constants.ConfigFile.StudentsSheet
$CLASSE = $constants.RequiredInputs.CLASSE
$TEACHER = $constants.RequiredInputs.TEACHER
$PROJECTNAME = $constants.RequiredInputs.PROJECTNAME
$NBWEEKS = $constants.RequiredInputs.NBWEEKS
$DATES = $constants.RequiredInputs.DATES
$DATEEND = $constants.RequiredInputs.DATEEND



if(!(Test-Path -Path $OutputPath -PathType Container)){
    New-Item -Path $OutputPath -ItemType Directory -Force -Confirm:$false
}

Start-Transcript -Path "$OutputPath/Output.log" -Append -Force

#Install all requirements for the script to run
.\Install-Requirements.ps1

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
$Configs = (Import-Excel -Path $ConfigsPath -WorksheetName $ConfigSheet)
$Students = (Import-Excel -Path $ConfigsPath -WorksheetName $StudentsSheet)
Write-Host "Chargement du fichier de configurations" -ForegroundColor Green

# Check that the Config file contains all the required inputs.
# The inputs are specified in the constants file
foreach($input in $constants.RequiredInputs.GetEnumerator()){
    if(!($Configs.Champs.contains($input.Value))){
        .\Stop-Program.ps1 -errorMessage "Veuillez remplir tout les champs du fichier de configurations $ConfigsPath"
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
        .\Stop-Program.ps1 -errorMessage "Excel n'est pas installé. Veuillez l'installer et recomencer !"
    }
    catch{
        .\Stop-Program.ps1 -errorMessage "Une erreur est survenue. Verifiez que Excel est bien installé et configuré !"
    }

    $excel = New-Object -ComObject excel.application
    $excel.visible = $false
    $workbook = $excel.Workbooks.Open($ModelPath)
    $Sheet1 = $workbook.worksheets.item(1)

    #Unprotect the sheet
    $Sheet1.Unprotect()
    
    #Replace the cells with the configs datas
    $Sheet1.cells.find("[NAME]") = "$($student.Prenom) $($student.Nom)"
    $Sheet1.cells.find("[CLASSE]") = $ConfigsHash[$CLASSE]
    $Sheet1.cells.find("[TEACHER]") = $ConfigsHash[$TEACHER]
    $Sheet1.cells.find("[PROJECTNAME]") = $ConfigsHash[$PROJECTNAME]
    $Sheet1.cells.find("[NBWEEKS]") = $ConfigsHash[$NBWEEKS]
    $Sheet1.cells.find("[DATES]") = "$($ConfigsHash[$DATES].ToString("yyyy/MM/dd"))-$($ConfigsHash[$DATEEND].ToString("yyyy/MM/dd"))"

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