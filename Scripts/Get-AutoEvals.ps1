<#
.NOTES
    *****************************************************************************
    ETML
    Name:	Get-AutoEvals.ps1
    Author:	Sylvain Philipona
    Date:	23.02.2023
 	*****************************************************************************
    Modifications
 	Date  : 27.02.2023
 	Author: Sylvain Philipona
 	Reason: Modification des messages d'erreurs et amélioration de la gestion des erreurs
 	*****************************************************************************
.SYNOPSIS
    Récupération des auto-évaluations des élèves
 	
.DESCRIPTION
    Ce script récupère les auto-évaluations remplies par les élèves et mises en commun dans un répertoire.
    Les auto-évaluations sont ensuite rapatriées dans un fichier de synthèse (SynthesisModelPath)
  	
.PARAMETER ConfigsPath
    Ceci est le chemin du fichier XLSX contenant Les configurations et la liste des élèves
 	
.PARAMETER SynthesisModelPath
    Ceci est le chemin du fichier XLSX Modèle de la synthèse des auto-évaluations
 	
.PARAMETER FilesPath
    Ceci est le chemin où les fichiers des auto-évaluations personels des élèves vont êtres récupérés

.OUTPUTS
	
.EXAMPLE
    .\Create-AutoEvals.ps1 -ConfigsPath ".\DataFiles\01-configs-auto-eval.xlsx" -SynthesisModelPath ".\DataFiles\03-synthese-auto-eval.xlsm" -FilesPath ".\Output"
 	
    Installation de NuGet
    Chargement du fichier d'inputs
    Chargement du fichier de configurations
    Importation de E:\09-P_Appro\PS-Eval\Scripts\Output\AutoEval-Dorian-Capelli.xlsx
    Importation de E:\09-P_Appro\PS-Eval\Scripts\Output\AutoEval-Joca-Bolli.xlsx
    Importation de E:\09-P_Appro\PS-Eval\Scripts\Output\AutoEval-Nolan-Praz.xlsx
    Importation de E:\09-P_Appro\PS-Eval\Scripts\Output\AutoEval-Sayeh-Younes.xlsx
    Importation de E:\09-P_Appro\PS-Eval\Scripts\Output\AutoEval-Sylvain-Philipona.xlsx
.LINK
    Install-Requirements.ps1
    Stop-Program.ps1
#>

param (
    [string]$ConfigsPath,
    [string]$SynthesisModelPath,
    [string]$FilesPath
)

if(!(Test-Path -Path $FilesPath -PathType Container)){
    .\Stop-Program.ps1 -errorMessage "Le dossier $FilesPath n'existe pas"
}

Start-Transcript -Path "$FilesPath/Output.log" -Append -Force

#Install all requirements for the script to run
.\Install-Requirements.ps1

#Load the data file
$data = Import-LocalizedData -BaseDirectory ./DataFiles -FileName Inputs.psd1
Write-Host "Chargement du fichier d'inputs" -ForegroundColor Green

#Import the configs and students inputs
$Configs = (Import-Excel -Path $ConfigsPath -WorksheetName $data.ConfigFile.ConfigSheet)
Write-Host "Chargement du fichier de configurations" -ForegroundColor Green

# Verify that the folder exists
if(!(Test-Path $FilesPath -PathType Container)){
    .\Stop-Program.ps1 -errorMessage "Le dossier '$FilesPath' n'existe pas"
}

#Get all excel files in the path
$AutoEvals = Get-ChildItem -Path $FilesPath -recurse -File -Include *.xlsx

# Verify that the path contains at least 1 AutoEval
if($AutoEvals.Length -lt 1){
    .\Stop-Program.ps1 -errorMessage "Le dossier '$FilesPath' ne contient pas d'auto évaluations"
}

#Create the COM object
try{
    $excel = New-Object -ComObject excel.application
    $excel.visible = $false
}
catch [System.Runtime.InteropServices.COMException] {
    .\Stop-Program.ps1 -errorMessage "Excel n'est pas installé. Veuillez l'installer et recomencer !"
}

$WorkbooxSynthesis = $excel.workbooks.Open($SynthesisModelPath)

# Recover all evals in the folder
foreach($eval in $AutoEvals){

    Write-Host "Importation de $($eval.FullName)"

    #Open the auto eval
    $WorkbookEval = $excel.Workbooks.Open($eval.FullName)
    $SheetEval = $WorkbookEval.worksheets.item(1)

    #Copy the auto eval in the synthesis file
    $SheetEval.copy($WorkbooxSynthesis.sheets.item(1))
    $WorkbookEval.Close()
}

#Convert Configs to ConfigsHash table
$ConfigsHash = @{}
foreach($config in $Configs){
    $ConfigsHash.Add($config.Champs, $config.Valeurs)
}


#Save and close the object
# AutoEvals-ProjectName-Classe-Prof-01.xlsm
$ExcelFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbookMacroEnabled
$FileName = "$FilesPath\AutoEvals-$($ConfigsHash[$data.RequiredInputs.PROJECTNAME])-$($ConfigsHash[$data.RequiredInputs.CLASSE])-$($ConfigsHash[$data.RequiredInputs.VISA])-1.xlsm"
$WorkbooxSynthesis.Saveas($FileName,$ExcelFixedFormat)
$excel.workbooks.Close()
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
"Enregistrement de $filename"

Stop-Transcript