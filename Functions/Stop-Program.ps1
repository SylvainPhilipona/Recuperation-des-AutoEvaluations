function Stop-Program {
    param (
        [string]$errorMessage
    )
    
    #Unloading the functions
    . Manage-Functions -remove

    #Display the optional error message
    if($errorMessage){
        Write-Error $errorMessage
    }
    
    #Stop the transcription
    Stop-Transcript

    exit
} 