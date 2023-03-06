<#
.NOTES
    *****************************************************************************
    ETML
    Name:	Show-Notification.ps1
    Author:	Sylvain Philipona
    Date:	06.02.2023
 	*****************************************************************************
    Modifications
 	Date  : 
 	Author: 
 	Reason: 
 	*****************************************************************************
.SYNOPSIS
    Affiche une notification toast
 	
.DESCRIPTION
    Affiche une notification toast avec le text fourni en paramètres
  	
.PARAMETER ToastTitle
    Le titre de la notification
 	
.PARAMETER ToastText
    Le contenu de la notification
	
.EXAMPLE
    .\Show-Notification.ps1 -ToastTitle "Recuperation-des-AutoEvaluations" -ToastText "Lancement de la création des auto-évaluations. Cela peux prendre 1-2 minutes"
    
.LINK
    https://den.dev/blog/powershell-windows-notification/
#>

[cmdletbinding()]
Param (
    [string]$ToastTitle,
    [string]$ToastText
)

# Load the Windows.UI.Notifications namespace and get the template for a ToastText02 style notification
[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null
$Template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02)

# Append the title and text to the XML document using the AppendChild method
# Use where-object cmdlet to filter the elements with the 'id' attribute equal to 1 or 2
$RawXml = [xml] $Template.GetXml()
($RawXml.toast.visual.binding.text|Where-Object {$_.id -eq "1"}).AppendChild($RawXml.CreateTextNode($ToastTitle)) > $null
($RawXml.toast.visual.binding.text|Where-Object {$_.id -eq "2"}).AppendChild($RawXml.CreateTextNode($ToastText)) > $null

# Create a new XML document object using the deserialized version of the modified XML string
$SerializedXml = New-Object Windows.Data.Xml.Dom.XmlDocument
$SerializedXml.LoadXml($RawXml.OuterXml)

# Create a new ToastNotification object and set its properties for the tag, group, and expiration time
$Toast = [Windows.UI.Notifications.ToastNotification]::new($SerializedXml)
$Toast.Tag = "PowerShell"
$Toast.Group = "PowerShell"
$Toast.ExpirationTime = [DateTimeOffset]::Now.AddMinutes(1)

# Create a ToastNotifier object for the PowerShell app and call its Show method with the ToastNotification object to display the toast notification on the Windows desktop
$Notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("PowerShell")
$Notifier.Show($Toast);