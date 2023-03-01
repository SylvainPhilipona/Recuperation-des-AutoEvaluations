<#
.NOTES
    *****************************************************************************
    ETML
    Name:	Test-Paths.ps1
    Author:	Sylvain Philipona
    Date:	23.02.2023
 	*****************************************************************************
    Modifications
 	Date  : 01.03.2023
 	Author: Sylvain Philipona
 	Reason: Changement de l'encodage du fichier en UFT-8 with BOM
 	*****************************************************************************
.SYNOPSIS
    Test l'existance de plusieurs fichiers / dossiers
 	
.DESCRIPTION
    Effectue un Test-Path sur tout les fichiers / dossiers fournis en paramètres et retourne un tableau de tout les elements qui n'existent pas
  	
.PARAMETER paths
    Ceci est la liste des fichiers / dossier à tester l'existance

.OUTPUTS
    - Un tableau de tout les elements qui n'existent pas
	
.EXAMPLE
    .\Test-Paths.ps1 -paths ".\01-config\01-infos-proj-eleves.xlsx", ".\01-config\02-modele-grille.xlsx", ".\01-config" 

    .\01-config\01-infos-proj-eleves.xlsx
    .\01-config\02-modele-grille.xlsx
    .\01-config
#>

param (
    [String[]]$paths
)

$notExistingPaths = @()

foreach($path in $paths){
    if(!(Test-path $path)){
        $notExistingPaths += $path
    }
}

return $notExistingPaths