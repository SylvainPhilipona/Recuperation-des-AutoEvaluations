$scripts = Get-ChildItem -Path .\Scripts -Filter *.ps1 -Exclude "PS-Eval.ps1" -Recurse

$compiled = "./app-eval-projets.ps1"

New-Item "./Program" -ItemType Directory -Force
Set-Location "./Program"

foreach($script in $scripts){
    "Function $($script.Name.Replace('.ps1', '')) {" >> $compiled
        (Get-Content $script.FullName).Replace(".\", "").Replace(".ps1", "") >> $compiled
    "}" >> $compiled
}

(Get-Content "..\Scripts\PS-Eval.ps1").Replace(".\", "").Replace(".ps1", "") >> $compiled
Copy-Item ..\Scripts\01-config -Recurse

New-Item -Path . -Name "start.bat"
Set-Content "start.bat" "powershell -executionPolicy bypass -noexit -file $compiled"

Set-Location "../"


# Install-Module ps2exe -Scope CurrentUser
# Import-Module ps2exe -UseWindowsPowerShell
# Invoke-ps2exe $compiled .\JOCA.exe -UNICODEEncoding

# Remove-Item $compiled -Force
