function Get-AutoEvals {
    param (
        [string]$FilesPath = "$($PSScriptRoot)\Output"
    )

    Start-Transcript -Path "./Output/Output.log" -Append -Force

    #Load the required functions
    . .\Manage-Functions.ps1
    . Manage-Functions

    # #Install all requirements for the script to run
    # Install-Requirements

    # Verify that the folder exists
    if(!(Test-Path $FilesPath -PathType Container)){
        Stop-Program -errorMessage "Le dossier '$FilesPath' n'existe pas"
    }
    
    #Get all excel files in the path
    $AutoEvals = Get-ChildItem -Path $FilesPath -recurse -File -Include test.xlsx

    # Verify that the path contains at least 1 AutoEval
    if($AutoEvals.Length -lt 1){
        Stop-Program -errorMessage "Le dossier '$FilesPath' ne contient pas d'auto évaluations"
    }

    # Recover all evals in the folder
    foreach($eval in $AutoEvals){

        Write-Host "Loading $($eval.FullName)"

        #Import the model file
        try{
            $excel = New-Object -ComObject excel.application
        }
        catch [System.Runtime.InteropServices.COMException] {
            Stop-Program -errorMessage "Excel n'est pas installé. Veuillez l'installer et recomencer !"
        }

        $excel = New-Object -ComObject excel.application
        $excel.visible = $false
        $workbook = $excel.Workbooks.Open($eval.FullName)
        $Sheet1 = $workbook.worksheets.item(1)

        $result = Find-CellByName -Sheet $Sheet1 -name "Dodo"
        $result

        

        #Close the object
        $excel.workbooks.Close()
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        
    }








    #Unloading the functions
    . Manage-Functions -remove

    Stop-Transcript
}