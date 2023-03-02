<#
.NOTES
    *****************************************************************************
    ETML
    Name:	Stop-Program.ps1
    Author:	Sylvain Philipona
    Date:	23.02.2023
 	*****************************************************************************
    Modifications
 	Date  : 01.03.2023
 	Author: Sylvain Philipona
 	Reason: Changement de l'encodage du fichier en UFT-8 with BOM
 	*****************************************************************************
.SYNOPSIS
    Arrete le programme
 	
.DESCRIPTION
    Arrete la transcription des logs
    Arrete le programme avec un 'throw' et affiche un message d'erreur optionel
  	
.PARAMETER errorMessage
    Ceci est le message d'erreur optipnel qui sera affiché lors de l'arret du programme 

.OUTPUTS
    - Une erreur PowerShell avec un message d'erreur optionel
	
.EXAMPLE
    .\Stop-Program.ps1 -errorMessage "Raison de l'arret"

    Raison de l'arret
    Au caractère E:\09-P_Appro\PS-Eval\Scripts\Stop-Program.ps1:40 : 1
    + throw $errorMessage
    + ~~~~~~~~~~~~~~~~~~~
        + CategoryInfo          : OperationStopped: (Raison de l'arret:String) [], RuntimeException
        + FullyQualifiedErrorId : Raison de l'arret
#>

param (
    [string]$errorMessage
)

try{
    Stop-Transcript | out-null
}
catch{}

throw $errorMessage