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
 	Reason: Ajout du mode avancé
 	*****************************************************************************
.SYNOPSIS
    Permet le lancement des scripts de création / récupération des auto-évaluations
 	
.DESCRIPTION
    Ce script affiche un formulaire pour lancer la création / récupération des auto-évaluations
    Un mode avancé permet de changer les fichiers de modèle, config et le dossier de sortie

.OUTPUTS
	Affiche un formulaire pour lancer la création / récupération des auto-évaluations

.EXAMPLE
    .\PS-Eval.ps1
 	
.LINK
    Create-AutoEvals.ps1
    Get-AutoEvals.ps1
#>


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

#####################  POPUP  #####################

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Main form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Récupération des AutoEvaluations'
$form.Size = New-Object System.Drawing.Size(300,170)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedSingle'
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.Topmost = $false

# Create auto-evals button
$createEvalsButton = New-Object System.Windows.Forms.Button
$createEvalsButton.Size = New-Object System.Drawing.Size(($form.Size.Width - 50),30)
$createEvalsButton.Left = ($form.ClientSize.Width - $createEvalsButton.Width) / 2 ;
$createEvalsButton.Top = 15 
$createEvalsButton.Text = 'Créer les auto-évaluations'
$form.Controls.Add($createEvalsButton)

# Get auto-evals button
$getEvalsButton = New-Object System.Windows.Forms.Button
$getEvalsButton.Size = New-Object System.Drawing.Size(($form.Size.Width - 50),30)
$getEvalsButton.Left = ($form.ClientSize.Width - $getEvalsButton.Width) / 2 ;
$getEvalsButton.Top = $createEvalsButton.Bottom + 15 
$getEvalsButton.Text = 'Rappatrier les auto-évaluations'
$form.Controls.Add($getEvalsButton)

# Advanced Mode link
$advancedModeLink = New-Object System.Windows.Forms.LinkLabel
$advancedModeLink.Size = New-Object System.Drawing.Size(($form.Size.Width - 50),30)
$advancedModeLink.Left = ($form.ClientSize.Width - $advancedModeLink.Width) / 2 ;
$advancedModeLink.Top = $getEvalsButton.Bottom + 15 
$advancedModeLink.LinkColor = "BLUE"
$advancedModeLink.ActiveLinkColor = "RED"
$advancedModeLink.Text = "Mode avancé"
$form.Controls.Add($advancedModeLink)

# Config path
$configPathInput = New-Object System.Windows.Forms.TextBox
$configPathInput.Size = New-Object System.Drawing.Size(($form.Size.Width - 50),30)
$configPathInput.Left = ($form.ClientSize.Width - $configPathInput.Width) / 2 ;
$configPathInput.Top = $getEvalsButton.Bottom + 50 
$configPathInput.Text = "$($PSScriptRoot)\01-config\01-infos-proj-eleves.xlsx"
$configPathInput.Enabled = $false
$form.Controls.Add($configPathInput)

# Model Path
$modelPathInput = New-Object System.Windows.Forms.TextBox
$modelPathInput.Size = New-Object System.Drawing.Size(($form.Size.Width - 50),30)
$modelPathInput.Left = ($form.ClientSize.Width - $modelPathInput.Width) / 2 ;
$modelPathInput.Top = $configPathInput.Bottom + 15 
$modelPathInput.Text = "$($PSScriptRoot)\01-config\02-modele-grille.xlsx"
$modelPathInput.Enabled = $false
$form.Controls.Add($modelPathInput)

# Synthesis Path
$synthesisPathInput = New-Object System.Windows.Forms.TextBox
$synthesisPathInput.Size = New-Object System.Drawing.Size(($form.Size.Width - 50),30)
$synthesisPathInput.Left = ($form.ClientSize.Width - $synthesisPathInput.Width) / 2 ;
$synthesisPathInput.Top = $modelPathInput.Bottom + 15 
$synthesisPathInput.Text = "$($PSScriptRoot)\01-config\03-synthese-eval.xlsm"
$synthesisPathInput.Enabled = $false
$form.Controls.Add($synthesisPathInput)

# Output Path
$outputPathInput = New-Object System.Windows.Forms.TextBox
$outputPathInput.Size = New-Object System.Drawing.Size(($form.Size.Width - 50),30)
$outputPathInput.Left = ($form.ClientSize.Width - $outputPathInput.Width) / 2 ;
$outputPathInput.Top = $synthesisPathInput.Bottom + 15 
$outputPathInput.Text = "$($PSScriptRoot)\02-evaluations"
$outputPathInput.Enabled = $false
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
            Stop-Transcript
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
        $form.Size = New-Object System.Drawing.Size(300,335)
        $configPathInput.Enabled = $true
        $modelPathInput.Enabled = $true
        $synthesisPathInput.Enabled = $true
        $outputPathInput.Enabled = $true
    }
);

# Show the form
$form.ShowDialog()