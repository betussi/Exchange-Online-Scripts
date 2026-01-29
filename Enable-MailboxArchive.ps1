$usersToActive = Import-Csv -Path "C:\temp\E3_Mailboxes.csv" | Select-Object -ExpandProperty DisplayName
 
foreach ($user in $usersToActive) {
    Enable-Mailbox -Identity $user -Archive
    Set-Mailbox -Identity $user -RetentionPolicy "PoliticaRetencao6Meses"
}