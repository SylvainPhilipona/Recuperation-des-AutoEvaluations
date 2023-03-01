param (
    [String[]]$paths
)

$notExistingPaths = @()

foreach($path in $paths){
    if(!(Test-path $path)){
        $notExistingPaths += $path
    }
}

return $notExistingPaths