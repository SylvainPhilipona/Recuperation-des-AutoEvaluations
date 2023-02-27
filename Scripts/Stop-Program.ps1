param (
    [string]$errorMessage
)


try{
    Stop-Transcript | out-null
}
catch{}

throw $errorMessage