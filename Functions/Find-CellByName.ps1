function Find-CellByName {
    param (
        [System.MarshalByRefObject]$Sheet, #ComObject type
        [string]$Name
    )

    foreach($cell in $Sheet.Range("A1", "Z50")){
        if($cell.Name.Name -eq $Name){
            return $cell
        }
    }
    return $null
}