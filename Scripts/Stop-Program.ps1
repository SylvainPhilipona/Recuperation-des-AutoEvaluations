param (
    [string]$errorMessage
)

#Unloading the functions
. Manage-Functions -remove

Stop-Transcript

throw $errorMessage