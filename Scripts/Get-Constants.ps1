<#
.NOTES
    *****************************************************************************
    ETML
    Name:	Get-Constants.ps1
    Author:	Sylvain Philipona
    Date:	02.03.2023
 	*****************************************************************************
    Modifications
 	Date  : 02.03.2023
 	Author: Sylvain Philipona
 	Reason: Ajout de constantes
 	*****************************************************************************
.SYNOPSIS
    Fichier de constantes
 	
.DESCRIPTION
    Ce fichier contient toutes les constantes nécessaires au bon fonctionement des scripts

.OUTPUTS
	- Un PSCustomObject avec toutes les constantes

.EXAMPLE
    .\Get-Constants.ps1

    form
    ----
    @{FormWidth=400; FormHeight=170; FormHeightAdvanced=380; InputsWidth=350; InputsHeight=30; LabelsHeight=15; SpaceBetweenInputs=15}
#>

return [PSCustomObject]@{
    form = [PSCustomObject]@{
        FormWidth = 400
        FormHeight = 170
        FormHeightAdvanced = 380
        InputsWidth = 350
        InputsHeight = 30
        LabelsHeight = 15
        SpaceBetweenInputs = 15
    }
}