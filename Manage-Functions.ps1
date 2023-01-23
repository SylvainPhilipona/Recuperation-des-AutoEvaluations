function Manage-Functions {
    param (
        [switch]$remove
    )

    $ScriptsPath = "Functions"
    $scripts = Get-ChildItem -Recurse $ScriptsPath -in *.ps1

    foreach($item in $scripts){
        if($remove){
            write-host "Removing $($item.Name)" -ForegroundColor Red
            Remove-Module $item.Name.Replace(".ps1","") -ErrorAction SilentlyContinue
        }
        else{
            write-host "Loading $($item.Name)" -ForegroundColor Green
            Import-module $item.FullName -Scope Local
        }
    }
}

