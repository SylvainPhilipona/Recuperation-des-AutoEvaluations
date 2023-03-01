<#
.NOTES
    *****************************************************************************
    ETML
    Name:	Compile-Project.ps1
    Author:	Sylvain Philipona
    Date:	23.02.2023
 	*****************************************************************************
    Modifications
 	Date  : 01.03.2023
 	Author: Sylvain Philipona
 	Reason: Ajout de l'entête
 	*****************************************************************************
.SYNOPSIS
    
 	
.DESCRIPTION
    
  	
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
	

.EXAMPLE
    .\Compile-Project.ps1 -compiled "app-eval-projets.ps1" -mainScript "PS-Eval.ps1" -configsPath "01-config" -scriptsPath "./Scripts" -outputPath "./Program"
 	
#>

param (
    [string]$compiled = "app-eval-projets.ps1",
    [string]$mainScript = "PS-Eval.ps1",
    [string]$configsPath = "01-config",
    [string]$scriptsPath = "./Scripts",
    [string]$outputPath = "./Program"
)
$scripts = Get-ChildItem -Path $scriptsPath -Filter *.ps1 -Exclude $mainScript -Recurse

# Create the Output folder and move to it
New-Item $outputPath -ItemType Directory -Force | Out-Null
Set-Location $outputPath

# Get all files content and wrap it in a 'Function'
# Remove all ".\" and ".ps1". Because in the files the scripts are called like this ".\My-Function.ps1" and when wrapped like this "My-Function"
# The result will be output in a file 
$wrapContent = ""
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
