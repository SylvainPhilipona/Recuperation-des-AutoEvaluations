function Find-CellByName {
    param (
        $Sheet,
        [string]$Name
    )

    $i = 0
    foreach($cell in $Sheet.Range("A1", "Z50")){
        if($cell.Name.Name -eq $Name){
            Write-Host $Name -ForegroundColor Green
            return $cell
        }

        Write-Host $i -ForegroundColor Red
        $i++
    }
    
    return $null
}