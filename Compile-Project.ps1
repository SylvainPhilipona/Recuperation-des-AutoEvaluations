<#
.NOTES
    *****************************************************************************
    ETML
    Name:	Compile-Project.ps1
    Author:	Sylvain Philipona
    Date:	23.02.2023
 	*****************************************************************************
    Modifications
 	Date  : 02.03.2023
 	Author: Sylvain Philipona
 	Reason: Ajout de l'entête
 	*****************************************************************************
.SYNOPSIS
    Compile tout les fichiers de scripts en 1 seul
 	
.DESCRIPTION
    Récupère tout les scripts .ps1 d'un dossier, les compile en 1 seul et crée un fichier .bat de lancement
    Permet de générer un fichier simple à uttiliser pour l'utilisateur final
  	
.PARAMETER compiled
    Le nom du fichier script de sortie qui contient tout les autres scripts

.PARAMETER mainScript
    Le script de départ (GUI)

.PARAMETER configsPath
    Le nom du dossier de configurations

.PARAMETER scriptsPath
    Chemin du dossier contenant les scripts

.PARAMETER outputPath
    Le chemin de sortie du script compilé

.OUTPUTS
	- Un dossier contenant les éléments suivant
    - Un dossier avec les fichiers de config
    - Un fichier .ps1 contenant tout les scripts
    - Un fichier .bat permettant de lancer le programme

.EXAMPLE
    .\Compile-Project.ps1 -compiled "app-eval-projets.ps1" -mainScript "PS-Eval.ps1" -configsPath "01-config" -scriptsPath "./Scripts" -outputPath "./Program"
 	
#>

param (
    [string]$compiled = "app-eval-projets.ps1",
    [string]$mainScript = "PS-Eval.ps1",
    [string]$constantsScript = "Get-Constants.ps1",
    [string]$configsPath = "01-config",
    [string]$scriptsPath = "./Scripts",
    [string]$outputPath = "./Recuperation-des-AutoEvaluations"
)
$scripts = Get-ChildItem -Path $scriptsPath -Filter *.ps1 -Exclude @($mainScript, $constantsScript) -Recurse

# Create the Output folder and move to it
New-Item $outputPath -ItemType Directory -Force | Out-Null
Set-Location $outputPath

# Get all files content and wrap it in a 'Function'
# Remove all ".\" and ".ps1". Because in the files the scripts are called like this ".\My-Function.ps1" and when wrapped like this "My-Function"
# The result will be output in a file 
$wrapContent = ""
$wrapContent += "Function $($constantsScript.Replace('.ps1', '')) {"
$wrapContent += [IO.File]::ReadAllText("$scriptsPath\$constantsScript").Replace(".\", "").Replace(".ps1", "")
$wrapContent += "}"
foreach($script in $scripts){
    $wrapContent += "Function $($script.Name.Replace('.ps1', '')) {"
    $wrapContent +=     [IO.File]::ReadAllText($script.FullName).Replace(".\", "").Replace(".ps1", "")
    $wrapContent += "}"
}
$wrapContent += [IO.File]::ReadAllText("$scriptsPath\$mainScript").Replace(".\", "").Replace(".ps1", "")
$wrapContent >> $compiled

# Copy the configs folder
Copy-Item "..\$scriptsPath\$configsPath" -force -Recurse

# Create the launch file
New-Item -Path . -Name "start.bat" -Force | Out-Null
Set-Content "start.bat" "powershell -executionPolicy bypass -file $compiled"

# Go back in the main folder
Set-Location "../"


# Install-Module ps2exe -Scope CurrentUser
# Import-Module ps2exe -UseWindowsPowerShell
# Invoke-ps2exe $compiled .\JOCA.exe -UNICODEEncoding

# Remove-Item $compiled -Force
