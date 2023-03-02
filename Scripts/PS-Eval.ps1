<#
.NOTES
    *****************************************************************************
    ETML
    Name:	PS-Eval.ps1
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
    .\PS-Eval.ps1
 	
.LINK
    Get-Constants.ps1
    Create-AutoEvals.ps1
    Get-AutoEvals.ps1 
#>
#####################  Constants  #####################

$constants = .\Get-Constants.ps1
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
$getEvalsButton.Text = 'Rappatrier les auto-évaluations'
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
            # Start the creation
            .\Create-AutoEvals.ps1 -ConfigsPath $configPathInput.Text -ModelPath $modelPathInput.Text -OutputPath $outputPathInput.Text

            [System.Windows.Forms.MessageBox]::Show("Tout bon" , "My Dialog Box")
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
            # Start the creation
            .\Get-AutoEvals.ps1 -ConfigsPath $configPathInput.Text -SynthesisModelPath $synthesisPathInput.Text -FilesPath $outputPathInput.Text
            [System.Windows.Forms.MessageBox]::Show("Tout bon" , "My Dialog Box")
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