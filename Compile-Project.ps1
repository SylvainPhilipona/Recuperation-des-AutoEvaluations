$scripts = Get-ChildItem -Path .\Scripts -Filter *.ps1 -Exclude "PS-Eval.ps1" -Recurse

$compile = "app-eval-projets.ps1"

foreach($script in $scripts){
    "Function $($script.Name.Replace('.ps1', '')) {" >> $compile
        (Get-Content $script.FullName).Replace(".\", "").Replace(".ps1", "") >> $compile
    "}" >> $compile
}

(Get-Content ".\Scripts\PS-Eval.ps1").Replace(".\", "").Replace(".ps1", "") >> $compile

New-Item -Path . -Name "start.bat"
Set-Content "start.bat" "powershell -executionPolicy bypass -noexit -file $compile"


# Install-Module ps2exe -Scope CurrentUser
# Import-Module ps2exe -UseWindowsPowerShell
# Invoke-ps2exe $compile .\JOCA.exe -UNICODEEncoding

# Remove-Item $compile -Force
