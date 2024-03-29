Function Get-Constants {<#
.NOTES
    *****************************************************************************
    ETML
    Name:	Get-Constants
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
    Get-Constants

    form
    ----
    @{FormWidth=400; FormHeight=170; FormHeightAdvanced=380; InputsWidth=350; InputsHeight=30; LabelsHeight=15; SpaceBetweenInputs=15}
#>

return [PSCustomObject]@{
    form = @{
        FormWidth = 400
        FormHeight = 170
        FormHeightAdvanced = 380
        InputsWidth = 350
        InputsHeight = 30
        LabelsHeight = 15
        SpaceBetweenInputs = 15
    }

    ConfigFile = @{
        ConfigSheet = 'configs'
        StudentsSheet = 'students'
    }

    RequiredInputs = @{
        CLASSROOM = 'Classe'
        TEACHER = 'Enseignant'
        PROJECTNAME = 'Nom du projet'
        NBWEEKS = 'Nbr de semaines'
        DATESSTART = 'Date debut'
        DATEEND = 'Date fin'
        VISA = 'Visa Enseignant'
    } 
}}Function Create-AutoEvals {<#
.NOTES
    *****************************************************************************
    ETML
    Name:	Create-AutoEvals
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
    Create-AutoEvals -ConfigsPath "./01-config\01-configs-auto-eval.xlsx" -ModelPath "./01-config\02-modele-auto-eval.xlsx" -OutputPath "./02-evaluations"

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
    Get-Constants
    Install-Requirements
    Test-Paths
    Stop-Program
#>

param (
    [string]$ConfigsPath,
    [string]$ModelPath,
    [string]$OutputPath
)

#####################  Constants  #####################

$constants = Get-Constants
$ConfigSheet = $constants.ConfigFile.ConfigSheet
$StudentsSheet = $constants.ConfigFile.StudentsSheet
$CLASSROOM = $constants.RequiredInputs.CLASSROOM
$TEACHER = $constants.RequiredInputs.TEACHER
$PROJECTNAME = $constants.RequiredInputs.PROJECTNAME
$NBWEEKS = $constants.RequiredInputs.NBWEEKS
$DATESSTART = $constants.RequiredInputs.DATESSTART
$DATEEND = $constants.RequiredInputs.DATEEND


# Create the output path if not exists
if(!(Test-Path -Path $OutputPath -PathType Container)){
    New-Item -Path $OutputPath -ItemType Directory -Force -Confirm:$false
}

# Start the transcription
Start-Transcript -Path "$OutputPath/Output.log" -Append -Force

#Install all requirements for the script to run
Install-Requirements

#Verify if the config and model files exists
$testPaths = Test-Paths -paths $ConfigsPath, $ModelPath
if(!$testPaths.count -eq 0){
    #Dispaly the missing paths
    $errorMessage = "Les fichiers suivants n'existent pas : `n`r"
    foreach($path in $testPaths){
        $errorMessage += " $path `n`r"
    }
    Stop-Program -errorMessage $errorMessage
}

#Import the configs and students inputs
$Configs = (Import-Excel -Path $ConfigsPath -WorksheetName $ConfigSheet)
$Students = (Import-Excel -Path $ConfigsPath -WorksheetName $StudentsSheet)
Write-Host "Chargement du fichier de configurations" -ForegroundColor Green

# Check that the Config file contains all the required inputs.
# The inputs are specified in the constants file
foreach($input in $constants.RequiredInputs.GetEnumerator()){
    if(!($Configs.Champs.contains($input.Value))){
        Stop-Program -errorMessage "Veuillez remplir tout les champs du fichier de configurations $ConfigsPath"
    }
}

#Convert Configs to ConfigsHash table
$ConfigsHash = @{}
foreach($config in $Configs){
    $ConfigsHash.Add($config.Champs, $config.Valeurs)
}

foreach($student in  $students){

    Write-Host "Création de $($student.Prenom) $($student.Nom) AutoEval"

    try{
        $excel = New-Object -ComObject excel.application
    }
    catch [System.Runtime.InteropServices.COMException] {
        Stop-Program -errorMessage "Excel n'est pas installé. Veuillez l'installer et recomencer !"
    }
    catch{
        Stop-Program -errorMessage "Une erreur est survenue. Verifiez que Excel est bien installé et configuré !"
    }

    #Import the model file
    $excel.visible = $false
    $workbook = $excel.Workbooks.Open($ModelPath)
    $Sheet1 = $workbook.worksheets.item(1)

    #Unprotect the sheet
    $Sheet1.Unprotect()
    
    #Replace the cells with the configs datas
    $Sheet1.cells.find("NAME") = "$($student.Prenom) $($student.Nom)"
    $Sheet1.cells.find("CLASSROOM") = $ConfigsHash[$CLASSROOM]
    $Sheet1.cells.find("TEACHER") = $ConfigsHash[$TEACHER]
    $Sheet1.cells.find("PROJECTNAME") = $ConfigsHash[$PROJECTNAME]
    $Sheet1.cells.find("NBWEEKS") = $ConfigsHash[$NBWEEKS]
    $Sheet1.cells.find("DATESMERGED") = "$($ConfigsHash[$DATESSTART].ToString("yyyy/MM/dd"))-$($ConfigsHash[$DATEEND].ToString("yyyy/MM/dd"))"

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

Stop-Transcript}Function Get-AutoEvals {<#
.NOTES
    *****************************************************************************
    ETML
    Name:	Get-AutoEvals
    Author:	Sylvain Philipona
    Date:	23.02.2023
 	*****************************************************************************
    Modifications
 	Date  : 06.03.2023
 	Author: Sylvain Philipona
 	Reason: Ajout d'une notif toast
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
    - Un fichier synthèse 'AutoEvals-ProjectName-Classe-Prof-01.xlsm' englobant toutes les auto-évaluations des élèves
    - Un fichier log avec les actions éffectuées
	
.EXAMPLE
    Create-AutoEvals -ConfigsPath "./01-config\01-configs-auto-eval.xlsx" -SynthesisModelPath "./01-config\03-synthese-auto-eval.xlsm" -FilesPath "./02-evaluations"
 	
    Installation de NuGet
    Chargement du fichier de configurations
    Importation de E:\09-P_Appro\PS-Eval\Scripts\02-evaluations\AutoEval-Dorian-Capelli.xlsx
    Importation de E:\09-P_Appro\PS-Eval\Scripts\02-evaluations\AutoEval-Joca-Bolli.xlsx
    Importation de E:\09-P_Appro\PS-Eval\Scripts\02-evaluations\AutoEval-Nolan-Praz.xlsx
    Importation de E:\09-P_Appro\PS-Eval\Scripts\02-evaluations\AutoEval-Sayeh-Younes.xlsx
    Importation de E:\09-P_Appro\PS-Eval\Scripts\02-evaluations\AutoEval-Sylvain-Philipona.xlsx
    Enregistrement de E:\09-P_Appro\PS-Eval\Scripts\02-evaluations\AutoEvals-P_Appro-CIN4b-GGZ-1.xlsm
.LINK
    Get-Constants
    Install-Requirements
    Stop-Program
#>

param (
    [string]$ConfigsPath,
    [string]$SynthesisModelPath,
    [string]$FilesPath
)

#####################  Constants  #####################

$constants = Get-Constants
$ConfigSheet = $constants.ConfigFile.ConfigSheet
$CLASSROOM = $constants.RequiredInputs.CLASSROOM
$PROJECTNAME = $constants.RequiredInputs.PROJECTNAME
$VISA = $constants.RequiredInputs.VISA

# Test if the auto-evals path exists
if(!(Test-Path -Path $FilesPath -PathType Container)){
    Stop-Program -errorMessage "Le dossier $FilesPath n'existe pas"
}

# Start the transcription
Start-Transcript -Path "$FilesPath/Output.log" -Append -Force

#Install all requirements for the script to run
Install-Requirements

#Verify if the config and model files exists
$testPaths = Test-Paths -paths $ConfigsPath, $SynthesisModelPath
if(!$testPaths.count -eq 0){
    #Dispaly the missing paths
    $errorMessage = "Les fichiers suivants n'existent pas : `n`r"
    foreach($path in $testPaths){
        $errorMessage += " $path `n`r"
    }
    Stop-Program -errorMessage $errorMessage
}

#Import the configs and students inputs
$Configs = (Import-Excel -Path $ConfigsPath -WorksheetName $ConfigSheet)
Write-Host "Chargement du fichier de configurations" -ForegroundColor Green

#Get all excel files in the path
$AutoEvals = Get-ChildItem -Path $FilesPath -recurse -File -Include *.xlsx

# Verify that the path contains at least 1 AutoEval
if($AutoEvals.Length -lt 1){
    Stop-Program -errorMessage "Le dossier '$FilesPath' ne contient pas d'auto évaluations"
}

#Create the COM object
try{
    $excel = New-Object -ComObject excel.application
    $excel.visible = $false
}
catch [System.Runtime.InteropServices.COMException] {
    Stop-Program -errorMessage "Excel n'est pas installé. Veuillez l'installer et recomencer !"
}
catch{
    Stop-Program -errorMessage "Une erreur est survenue. Verifiez que Excel est bien installé et configuré !"
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
$FileName = "$FilesPath\AutoEvals-$($ConfigsHash[$PROJECTNAME])-$($ConfigsHash[$CLASSROOM])-$($ConfigsHash[$VISA])-1.xlsm"
$WorkbooxSynthesis.Saveas($FileName,$ExcelFixedFormat)
$excel.workbooks.Close()
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "Enregistrement de $filename"

Stop-Transcript}Function Install-Requirements {<#
.NOTES
    *****************************************************************************
    ETML
    Name:	Install-Requirements
    Author:	Sylvain Philipona
    Date:	23.02.2023
 	*****************************************************************************
    Modifications
 	Date  : 01.03.2023
 	Author: Sylvain Philipona
 	Reason: Changement de l'encodage du fichier en UFT-8 with BOM
 	*****************************************************************************
.SYNOPSIS
    Installe les exigences d'installation

.DESCRIPTION
    Installe les modules et packages necessaires au bon fonctionnement des scripts

.EXAMPLE
    Install-Requirements

    Installation de NuGet
    Ajout du répertoire PSGallery en répertoire de confiance
    Installation de ImportExcel
#>

# Install the NuGet package
Write-Host "Installation de NuGet" -ForegroundColor Green
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.208 -Scope CurrentUser -Force -Confirm:$false | Out-Null

# Set PSGallery repo to trusted -> For the ImportExcel installation
if((Get-PSRepository -Name "PSGallery").InstallationPolicy -ne "Trusted"){
    Write-Host "Ajout du répertoire PSGallery en répertoire de confiance" -ForegroundColor Green
    Set-PSRepository -name  "PSGallery" -InstallationPolicy Trusted
}

# Install the module ImportExcel
if(!(Get-Module -ListAvailable -name ImportExcel)){
    Write-Host "Installation de ImportExcel" -ForegroundColor Green
    Install-Module ImportExcel -Scope CurrentUser -Confirm:$false #https://github.com/dfinke/ImportExcel
}}Function Show-Notification {<#
.NOTES
    *****************************************************************************
    ETML
    Name:	Show-Notification
    Author:	Sylvain Philipona
    Date:	06.02.2023
 	*****************************************************************************
    Modifications
 	Date  : 
 	Author: 
 	Reason: 
 	*****************************************************************************
.SYNOPSIS
    Affiche une notification toast
 	
.DESCRIPTION
    Affiche une notification toast avec le text fourni en paramètres
  	
.PARAMETER ToastTitle
    Le titre de la notification
 	
.PARAMETER ToastText
    Le contenu de la notification
	
.EXAMPLE
    Show-Notification -ToastTitle "Recuperation-des-AutoEvaluations" -ToastText "Lancement de la création des auto-évaluations. Cela peux prendre 1-2 minutes"
    
.LINK
    https://den.dev/blog/powershell-windows-notification/
#>

[cmdletbinding()]
Param (
    [string]$ToastTitle,
    [string]$ToastText
)

# Load the Windows.UI.Notifications namespace and get the template for a ToastText02 style notification
[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null
$Template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02)

# Append the title and text to the XML document using the AppendChild method
# Use where-object cmdlet to filter the elements with the 'id' attribute equal to 1 or 2
$RawXml = [xml] $Template.GetXml()
($RawXml.toast.visual.binding.text|Where-Object {$_.id -eq "1"}).AppendChild($RawXml.CreateTextNode($ToastTitle)) > $null
($RawXml.toast.visual.binding.text|Where-Object {$_.id -eq "2"}).AppendChild($RawXml.CreateTextNode($ToastText)) > $null

# Create a new XML document object using the deserialized version of the modified XML string
$SerializedXml = New-Object Windows.Data.Xml.Dom.XmlDocument
$SerializedXml.LoadXml($RawXml.OuterXml)

# Create a new ToastNotification object and set its properties for the tag, group, and expiration time
$Toast = [Windows.UI.Notifications.ToastNotification]::new($SerializedXml)
$Toast.Tag = "PowerShell"
$Toast.Group = "PowerShell"
$Toast.ExpirationTime = [DateTimeOffset]::Now.AddMinutes(1)

# Create a ToastNotifier object for the PowerShell app and call its Show method with the ToastNotification object to display the toast notification on the Windows desktop
$Notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("PowerShell")
$Notifier.Show($Toast);}Function Stop-Program {<#
.NOTES
    *****************************************************************************
    ETML
    Name:	Stop-Program
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
    Stop-Program -errorMessage "Raison de l'arret"

    Raison de l'arret
    Au caractère E:\09-P_Appro\PS-Eval\Scripts\Stop-Program:40 : 1
    + throw $errorMessage
    + ~~~~~~~~~~~~~~~~~~~
        + CategoryInfo          : OperationStopped: (Raison de l'arret:String) [], RuntimeException
        + FullyQualifiedErrorId : Raison de l'arret
#>

param (
    [string]$errorMessage
)

try{
    # Stop the transcription
    Stop-Transcript | out-null
}
catch{}

# Throw the custom error message
throw $errorMessage}Function Test-Paths {<#
.NOTES
    *****************************************************************************
    ETML
    Name:	Test-Paths
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
    Test-Paths -paths "./01-config\01-infos-proj-eleves.xlsx", "./01-config\02-modele-grille.xlsx", "./01-config" 

    ./01-config\01-infos-proj-eleves.xlsx
    ./01-config\02-modele-grille.xlsx
    ./01-config
#>

param (
    [String[]]$paths
)

$notExistingPaths = @()

# Test all paths
foreach($path in $paths){
    if(!(Test-path $path)){
        $notExistingPaths += $path
    }
}

return $notExistingPaths}<#
.NOTES
    *****************************************************************************
    ETML
    Name:	PS-Eval
    Author:	Sylvain Philipona
    Date:	24.02.2023
 	*****************************************************************************
    Modifications
 	Date  : 01.03.2023
 	Author: Sylvain Philipona
 	Reason: Ajout de constantes
 	*****************************************************************************
.SYNOPSIS
    Permet le lancement des scripts de création / récupération des auto-évaluations
 	
.DESCRIPTION
    Ce script affiche un formulaire pour lancer la création / récupération des auto-évaluations
    Un mode avancé permet de changer les fichiers de modèle, config et le dossier de sortie

.OUTPUTSFormWidth
	Affiche un formulaire pour lancer la création / récupération des auto-évaluations

.EXAMPLE
    PS-Eval
 	
.LINK
    Get-Constants
    Create-AutoEvals
    Get-AutoEvals 
#>

#####################  Constants  #####################

$constants = Get-Constants
$FormWidth = $constants.form.FormWidth
$FormHeight = $constants.form.FormHeight
$FormHeightAdvanced = $constants.form.FormHeightAdvanced
$InputsWidth = $constants.form.InputsWidth
$InputsHeight = $constants.form.InputsHeight
$LabelsHeight = $constants.form.LabelsHeight
$SpaceBetweenInputs = $constants.form.SpaceBetweenInputs

#####################  No console display #####################

# .Net methods for hiding/showing the console in the background
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

function Hide-Console
{
    $consolePtr = [Console.Window]::GetConsoleWindow()
    #0 = hide
    [Console.Window]::ShowWindow($consolePtr, 0) | Out-Null
} 
Hide-Console

#####################  Config form  #####################

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global variable because its value is changed in an event
$global:advancedMode = $false

# Main form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Récupération des AutoEvaluations'
$form.Size = New-Object System.Drawing.Size($FormWidth,$FormHeight)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedSingle'
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.Topmost = $false

# Create auto-evals button
$createEvalsButton = New-Object System.Windows.Forms.Button
$createEvalsButton.Size = New-Object System.Drawing.Size($InputsWidth,$InputsHeight)
$createEvalsButton.Left = ($form.ClientSize.Width - $createEvalsButton.Width) / 2 ;
$createEvalsButton.Top = $SpaceBetweenInputs 
$createEvalsButton.Text = 'Créer les auto-évaluations'
$form.Controls.Add($createEvalsButton)

# Get auto-evals button
$getEvalsButton = New-Object System.Windows.Forms.Button
$getEvalsButton.Size = New-Object System.Drawing.Size($InputsWidth,$InputsHeight)
$getEvalsButton.Left = ($form.ClientSize.Width - $getEvalsButton.Width) / 2 ;
$getEvalsButton.Top = $createEvalsButton.Bottom + $SpaceBetweenInputs 
$getEvalsButton.Text = 'Rapatrier les auto-évaluations'
$form.Controls.Add($getEvalsButton)

# Advanced Mode link
$advancedModeLink = New-Object System.Windows.Forms.LinkLabel
$advancedModeLink.Size = New-Object System.Drawing.Size($InputsWidth,$InputsHeight)
$advancedModeLink.Left = ($form.ClientSize.Width - $advancedModeLink.Width) / 2 ;
$advancedModeLink.Top = $getEvalsButton.Bottom + $SpaceBetweenInputs 
$advancedModeLink.LinkColor = "BLUE"
$advancedModeLink.ActiveLinkColor = "RED"
$advancedModeLink.Text = "Mode avancé"
$form.Controls.Add($advancedModeLink)


# Config path
$configPathLabel = New-Object System.Windows.Forms.Label
$configPathLabel.Size = New-Object System.Drawing.Size($InputsWidth,$LabelsHeight)
$configPathLabel.Left = ($form.ClientSize.Width - $configPathLabel.Width) / 2 ;
$configPathLabel.Top = $getEvalsButton.Bottom + 50 
$configPathLabel.Text = "Fichier de configurations"

$configPathInput = New-Object System.Windows.Forms.TextBox
$configPathInput.Size = New-Object System.Drawing.Size($InputsWidth,$InputsHeight)
$configPathInput.Left = ($form.ClientSize.Width - $configPathInput.Width) / 2 ;
$configPathInput.Top = $configPathLabel.Bottom 
$configPathInput.Text = "$($PSScriptRoot)\01-config\01-infos-proj-eleves.xlsx"
$configPathInput.Enabled = $false
$form.Controls.Add($configPathLabel)
$form.Controls.Add($configPathInput)


# Model Path
$modelPathLabel = New-Object System.Windows.Forms.Label
$modelPathLabel.Size = New-Object System.Drawing.Size($InputsWidth,$LabelsHeight)
$modelPathLabel.Left = ($form.ClientSize.Width - $modelPathLabel.Width) / 2 ;
$modelPathLabel.Top = $configPathInput.Bottom + $SpaceBetweenInputs 
$modelPathLabel.Text = "Fichier model d'auto-évaluations"

$modelPathInput = New-Object System.Windows.Forms.TextBox
$modelPathInput.Size = New-Object System.Drawing.Size($InputsWidth,$InputsHeight)
$modelPathInput.Left = ($form.ClientSize.Width - $modelPathInput.Width) / 2 ;
$modelPathInput.Top = $modelPathLabel.Bottom
$modelPathInput.Text = "$($PSScriptRoot)\01-config\02-modele-grille.xlsx"
$modelPathInput.Enabled = $false
$form.Controls.Add($modelPathLabel)
$form.Controls.Add($modelPathInput)


# Synthesis Path
$synthesisPathLabel = New-Object System.Windows.Forms.Label
$synthesisPathLabel.Size = New-Object System.Drawing.Size($InputsWidth,$LabelsHeight)
$synthesisPathLabel.Left = ($form.ClientSize.Width - $synthesisPathLabel.Width) / 2 ;
$synthesisPathLabel.Top = $modelPathInput.Bottom + $SpaceBetweenInputs 
$synthesisPathLabel.Text = "Fichier model de synthèse"

$synthesisPathInput = New-Object System.Windows.Forms.TextBox
$synthesisPathInput.Size = New-Object System.Drawing.Size($InputsWidth,$InputsHeight)
$synthesisPathInput.Left = ($form.ClientSize.Width - $synthesisPathInput.Width) / 2 ;
$synthesisPathInput.Top = $synthesisPathLabel.Bottom
$synthesisPathInput.Text = "$($PSScriptRoot)\01-config\03-synthese-eval.xlsm"
$synthesisPathInput.Enabled = $false
$form.Controls.Add($synthesisPathLabel)
$form.Controls.Add($synthesisPathInput)


# Output Path
$outputPathLabel = New-Object System.Windows.Forms.Label
$outputPathLabel.Size = New-Object System.Drawing.Size($InputsWidth,$LabelsHeight)
$outputPathLabel.Left = ($form.ClientSize.Width - $outputPathLabel.Width) / 2 ;
$outputPathLabel.Top = $synthesisPathInput.Bottom + $SpaceBetweenInputs 
$outputPathLabel.Text = "Dossier de sortie"

$outputPathInput = New-Object System.Windows.Forms.TextBox
$outputPathInput.Size = New-Object System.Drawing.Size($InputsWidth,$InputsHeight)
$outputPathInput.Left = ($form.ClientSize.Width - $outputPathInput.Width) / 2 ;
$outputPathInput.Top = $outputPathLabel.Bottom
$outputPathInput.Text = "$($PSScriptRoot)\02-evaluations"
$outputPathInput.Enabled = $false
$form.Controls.Add($outputPathLabel)
$form.Controls.Add($outputPathInput)


# Create auto-evals Button event
$createEvalsButton.Add_Click(
    {    
        # Lock the form and buttons
        $form.Enabled = $false

        # Trim the Output Path 
        $outputPathInput.Text = $outputPathInput.Text.TrimEnd(' ')
        $outputPathInput.Text = $outputPathInput.Text.TrimEnd('\')

        try{
            # Display the toast notif
            Show-Notification -ToastTitle "Recuperation des auto-évaluations" -ToastText "Lancement de la création des auto-évaluations. Cela peux prendre 1-2 minutes"

            # Start the creation
            Create-AutoEvals -ConfigsPath $configPathInput.Text -ModelPath $modelPathInput.Text -OutputPath $outputPathInput.Text

            [System.Windows.Forms.MessageBox]::Show("Les auto-évaluations ont étés crées avec succes !" , "Création réussie")
        }
        catch{
            #Display the error message
            try{Stop-Transcript}catch{}
            [System.Windows.Forms.MessageBox]::Show($_ , "Erreur d'execution")
        }
        
        # Unlock the form and buttons
        $form.Enabled = $true
    }
);

# Get auto-evals Button event
$getEvalsButton.Add_Click(
    {    
        # Lock the form and buttons
        $form.Enabled = $false

        # Trim the Output Path 
        $outputPathInput.Text = $outputPathInput.Text.TrimEnd(' ')
        $outputPathInput.Text = $outputPathInput.Text.TrimEnd('\')
        
        try{
            # Dispaly the toast notif
            Show-Notification -ToastTitle "Recuperation des auto-évaluations" -ToastText "Lancement de la récupération des auto-évaluations. Cela peux prendre 1-2 minutes"

            # Start the creation
            Get-AutoEvals -ConfigsPath $configPathInput.Text -SynthesisModelPath $synthesisPathInput.Text -FilesPath $outputPathInput.Text
            [System.Windows.Forms.MessageBox]::Show("Les auto-évaluations ont étés récupérées avec succes !" , "Récupération réussie")

        }
        catch{
            #Display the error message
            try{Stop-Transcript}catch{}
            [System.Windows.Forms.MessageBox]::Show($_ , "Erreur d'execution")
        }
        
        # Unlock the form and buttons
        $form.Enabled = $true
    }
);

$advancedModeLink.add_Click(
    {
        if(!($global:advancedMode)){
            $form.Size = New-Object System.Drawing.Size($FormWidth,$FormHeightAdvanced)
            $configPathInput.Enabled = $true
            $modelPathInput.Enabled = $true
            $synthesisPathInput.Enabled = $true
            $outputPathInput.Enabled = $true
            $advancedModeLink.Text = "Masquer le mode avancé"
            $global:advancedMode = $true
        }
        else{
            $form.Size = New-Object System.Drawing.Size($FormWidth,$FormHeight)
            $configPathInput.Enabled = $false
            $modelPathInput.Enabled = $false
            $synthesisPathInput.Enabled = $false
            $outputPathInput.Enabled = $false
            $advancedModeLink.Text = "Mode avancé"
            $global:advancedMode = $false
        }
    }
);

# Show the form
$form.ShowDialog()
