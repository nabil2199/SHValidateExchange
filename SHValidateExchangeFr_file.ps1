<#
Surface Hub on premises exchange account validator
Verion 0.1
Nabil LAKHNACHFI
OCWS
#>

#Write-Host $strLyncIdentity
$Global:iFileText=$null
#file path
$date=(Get-Date).ToString('yyyy-MM-dd_HH-mm-ss')
$logFilePath=$PSScriptRoot+"\LogSHValidateSfBfr-"+$date+".txt"

function CleanupAndFail($strMsg)
{
    if ($strMsg)
    {
        PrintError($strMsg);
    }
    Cleanup
    exit 1
}

$strUpn = Read-Host "Quel est l'adresse mail qui sera utilisee par la Surface HUB?"
if (!$strUpn.Contains('@'))
{
    CleanupAndFail "$strUpn n'est pas une adresse mail valide"
}

$mailbox = $null
$mailbox = Get-Mailbox -Identity $strUpn

if (!$mailbox)
{
    "Impossible de trouver la boite au lettre  $strUpn"
}

$exchange = Get-ExchangeServer
if (!$exchange.IsE14OrLater)
{
    CleanupAndFail "La version d'exchange est trop ancienne, veuillez utiliser exchange 2010 ou anterieur."
}


$Global:iTotalFailures = 0
$global:iTotalWarnings = 0
$Global:iTotalPasses = 0

function Validate()
{
    Param(
        [string]$Test,
        [bool]  $Condition,
        [string]$FailureMsg,
        [switch]$WarningOnly
    )

    Write-Host -NoNewline -ForegroundColor White $Test.PadRight(100,'.')
    if ($Condition)
    {
        Write-Host -ForegroundColor Green "Reussi"
        Out-file -filePath $logFilePath -append -InputObject("Reussi")
        $global:iTotalPasses++
    }
    else
    {
        if ($WarningOnly)
        {
            Write-Host -ForegroundColor Yellow ("Avertissement: "+$FailureMsg)
            Out-file -filePath $logFilePath -append -InputObject("Avertissement: "+$FailureMsg)
            $global:iTotalWarnings++
        }
        else
        {
            Write-Host -ForegroundColor Red ("Echec: "+$FailureMsg)
            Out-file -filePath $logFilePath -append -InputObject("Echec: "+$FailureMsg)
            $global:iTotalFailures++
        }
    }
}



## Exchange ##

Validate -WarningOnly -Test "La boite aux lettres $strUpn est configuree en tant que compte salle de reunion" -Condition ($mailbox.RoomMailboxAccountEnabled -eq $True) -FailureMsg "Propriete:RoomMailboxEnabled - La Surface HUB ne pourra pas utiliser certaines fonctionnalitees. Ceci ne concerne qu'Exchange 2016"
$calendarProcessing = Get-CalendarProcessing -Identity $strUpn -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
Validate -Test "La boite aux lettres $strUpn est configuree pour accepter les reunions" -Condition ($calendarProcessing -ne $null -and $calendarProcessing.AutomateProcessing -eq 'AutoAccept') -FailureMsg "Propriete:AutomateProcessing - La Surface Hub ne pourra pas envoyer de courier ni synchoniser son calendrier."
Validate -WarningOnly -Test "La boite aux lettres $strUpn ne supprimera pas les commentaires" -Condition ($calendarProcessing -ne $null -and !$calendarProcessing.DeleteComments) -FailureMsg "Propriete:DeleteComments - La Surface Hub presentra toutes les informations relatives aux reunions."
Validate -WarningOnly -Test "La boite aux lettres $strUpn gardera le statut prive des reunions" -Condition ($calendarProcessing -ne $null -and !$calendarProcessing.RemovePrivateProperty) -FailureMsg "Propriete:RemovePrivateProperty - La Surface Hub supprimera le flag private."
Validate -Test "La boite aux lettres $strUpn gardera les objets des reunions" -Condition ($calendarProcessing -ne $null -and !$calendarProcessing.DeleteSubject) -FailureMsg "Propriete:DeleteSubject - La Surface Hub supprimera l'objet de la reunion."
Validate -WarningOnly -Test "La boite aux lettres $strUpn n'ajoute pas le nom de l'organisateur a l'objet de la reunion" -Condition ($calendarProcessing -ne $null -and !$calendarProcessing.AddOrganizerToSubject) -FailureMsg "Propriete:AddOrganizerToSubject - the Surface Hub will not display meeting subjects as intended."


#ActiveSync
$casMailbox = Get-Casmailbox $strUpn -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
Validate -Test "La boite aux lettres $strUpn a une strategie Casmailbox" -Condition ($casMailbox -ne $null) -FailureMsg "Commandlet Get-Casmailbox  - La Surface Hub ne pourra pas envoyer de courier ni synchoniser son calendrier."

if ($casMailbox)
{
    $policy = $null
    if ($exchange.IsE15OrLater)
    {
        $strPolicy = $casMailbox.ActiveSyncMailboxPolicy
        $policy = Get-MobileDeviceMailboxPolicy -Identity $strPolicy -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
        Validate -Test "La strategie $strPolicy ne requiere pas de mot de passe" -Condition ($policy.PasswordEnabled -ne $True) -FailureMsg "Propriete:PasswordEnabled - La strategie requiere un mot de passe - La Surface Hub ne pourra pas envoyer de courier ni synchoniser son calendrier."
    }
    else
    {
        $strPolicy = $casMailbox.ActiveSyncMailboxPolicy
        $policy = Get-ActiveSyncMailboxPolicy -Identity $strPolicy -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
        Validate -Test "La strategie $strPolicy ne requiere pas de mot de passe" -Condition ($policy.DevicePasswordEnabled -ne $True) -FailureMsg "Propriete:DevicePasswordEnabled - La strategie requiere un mot de passe - La Surface Hub ne pourra pas envoyer de courier ni synchoniser son calendrier."
    }
    
    if ($policy -ne $null)
    {
        Validate -Test "La strategie $strPolicy autorise tous les peripheriques a se connnecter a exchange" -Condition ($policy.AllowNonProvisionableDevices -eq $null -or $policy.AllowNonProvisionableDevices -eq $true) -FailureMsg "Propriete:AllowNonProvisionableDevices - Exchange n'autorisera pas la Surface Hub a synchoniser"
    }
    
}

# Check the default access level
$orgSettings = Get-ActiveSyncOrganizationSettings
$strDefaultAccessLevel = $orgSettings.DefaultAccessLevel
Validate -Test "La connexion des peripherique ActiveSync est autorisee" -Condition ($strDefaultAccessLevel -eq 'Allow') -FailureMsg "Propriete Get-ActiveSyncOrganizationSettings.DefaultAccessLevel Les peripherique de type Windows Mail ne sont pas autorises par defaut - La Surface Hub ne pourra pas envoyer de courier ni synchoniser son calendrier."

# Check if there exists a device access rule that bans the device type Windows Mail
$blockingRules = Get-ActiveSyncDeviceAccessRule | where {($_.AccessLevel -eq 'Block' -or $_.AccessLevel -eq 'Quarantine') -and $_.Characteristic -eq 'DeviceType'-and $_.QueryString -eq 'WindowsMail'}
Validate -Test "Les peripheriques de type Windows mail ne sont pas bloques ni sous quarantaine" -Condition ($blockingRules -eq $null -or $blockingRules.Length -eq 0) -FailureMsg "Propriete Get-ActiveSyncOrganizationSettings.AccessLevel - Les peripherique de type Windows Mail sont bloques ou sous quarantaine - La Surface Hub ne pourra pas envoyer de courier ni synchoniser son calendrier."

## End Exchange ##

## Summary ##

$global:iTotalTests = ($global:iTotalFailures + $global:iTotalPasses + $global:iTotalWarnings)
Out-file -filePath $logFilePath -append -InputObject("")
Out-file -filePath $logFilePath -append -InputObject("")
Write-Host -NoNewline $global:iTotalTests "tests realises: "
Out-file -filePath $logFilePath -append -InputObject("tests realises: " + $global:iTotalTests )
Write-Host -NoNewline -ForegroundColor Red $Global:iTotalFailures "echecs "
Out-file -filePath $logFilePath -append -InputObject("echecs: " + $Global:iTotalFailures )
Write-Host -NoNewline -ForegroundColor Yellow $Global:iTotalWarnings "avertissements "
Out-file -filePath $logFilePath -append -InputObject("avertissements: " + $Global:iTotalWarnings  )
Write-Host -ForegroundColor Green $Global:iTotalPasses "reussis "
Out-file -filePath $logFilePath -append -InputObject("reussis: " + $Global:iTotalPasses )

## End Summary ##