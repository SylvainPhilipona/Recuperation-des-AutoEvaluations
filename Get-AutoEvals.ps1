function Get-AutoEvals {
    param (
        [string]$ConfigsPath = "$($PSScriptRoot)\DataFiles\01-configs-auto-eval.xlsx",
        [string]$SynthesisModelPath = "$($PSScriptRoot)\DataFiles\03-synthese-auto-eval.xlsm",
        [string]$FilesPath = "$($PSScriptRoot)\Output"
    )

    Start-Transcript -Path "./Output/Output.log" -Append -Force

    #Load the required functions
    . .\Manage-Functions.ps1
    . Manage-Functions

    # #Install all requirements for the script to run
    # Install-Requirements

    #Load the data file
    $data = Import-LocalizedData -BaseDirectory ./DataFiles -FileName Inputs.psd1
    Write-Host "Loading inputs file" -ForegroundColor Green

    #Import the configs and students inputs
    $Configs = (Import-Excel -Path $ConfigsPath -WorksheetName $data.ConfigFile.ConfigSheet)
    Write-Host "Loading configs file" -ForegroundColor Green

    # Verify that the folder exists
    if(!(Test-Path $FilesPath -PathType Container)){
        Stop-Program -errorMessage "Le dossier '$FilesPath' n'existe pas"
    }
    
    #Get all excel files in the path
    $AutoEvals = Get-ChildItem -Path $FilesPath -recurse -File -Include *.xlsx

    # Verify that the path contains at least 1 AutoEval
    if($AutoEvals.Length -lt 1){
        Stop-Program -errorMessage "Le dossier '$FilesPath' ne contient pas d'auto évaluations"
    }

    #Create the com object
    try{
        $excel = New-Object -ComObject excel.application
        $excel.visible = $false
    }
    catch [System.Runtime.InteropServices.COMException] {
       Stop-Program -errorMessage "Excel n'est pas installé. Veuillez l'installer et recomencer !"
    }

    $WorkbooxSynthesis = $excel.workbooks.Add()
    $WorkbookSynthesisModel = $excel.workbooks.Open($SynthesisModelPath)

    # Recover all evals in the folder
    foreach($eval in $AutoEvals){

        Write-Host "Loading $($eval.FullName)"

        #Import the model file
        $WorkbookEval = $excel.Workbooks.Open($eval.FullName)
        $SheetEval = $WorkbookEval.worksheets.item(1)

        #Get all cells with the needed content
        $NAME = Find-CellByName -Sheet $SheetEval -name "NAME"
        $CLASSE = Find-CellByName -Sheet $SheetEval -name "CLASSE"
        $TEACHER = Find-CellByName -Sheet $SheetEval -name "TEACHER"
        $PROJECTNAME = Find-CellByName -Sheet $SheetEval -name "PROJECTNAME"
        $NBWEEKS = Find-CellByName -Sheet $SheetEval -name "NBWEEKS"
        $DATES = Find-CellByName -Sheet $SheetEval -name "DATES"
        $FINALNOTE = Find-CellByName -Sheet $SheetEval -name "FINALNOTE"
        
        write-host "Importing data :" -ForegroundColor Green
        "   " + $NAME.Text
        "   " + $CLASSE.Text
        "   " + $TEACHER.Text
        "   " + $PROJECTNAME.Text
        "   " + $NBWEEKS.Text
        "   " + $DATES.Text
        "   " + $FINALNOTE.Text

        #Copy the auto eval in the synthesis file
        $SheetEval.copy($WorkbooxSynthesis.sheets.item(1))
        $WorkbookEval.Close()

        
    }

    $WorkbookSynthesisModel | Get-Member

    $WorkbookSynthesisModel.sheets(1).Copy($WorkbooxSynthesis.sheets.item(1))
    $WorkbookSynthesisModel.close()

    #Convert Configs to ConfigsHash table
    $ConfigsHash = @{}
    foreach($config in $Configs){
        $ConfigsHash.Add($config.Champs, $config.Valeurs)
    }
    

    #Save and close the object
    # AutoEvals-ProjectName-Classe-Prof-01.xlsm
    $ExcelFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbookMacroEnabled
    $FileName = "$($PSScriptRoot)\Output\AutoEvals-$($ConfigsHash[$data.RequiredInputs.PROJECTNAME])-$($ConfigsHash[$data.RequiredInputs.CLASSE])-$($ConfigsHash[$data.RequiredInputs.VISA])-1.xlsm"
    $WorkbooxSynthesis.Saveas($FileName,$ExcelFixedFormat)
    $excel.workbooks.Close()
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    "Saving $filename"


    #Unloading the functions
    . Manage-Functions -remove

    Stop-Transcript
}