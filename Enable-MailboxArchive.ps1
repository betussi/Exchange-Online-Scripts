<# 
PRÉ-REQUISITOS:

1) Criar um arquivo CSV no caminho:
   C:\temp\E3_Mailboxes.csv

2) O arquivo CSV deve conter, no mínimo, a coluna abaixo:
   DisplayName

   Exemplo de conteúdo do CSV:
   --------------------------------
   DisplayName
   João Silva
   Maria Oliveira
   Carlos Pereira
   --------------------------------

3) Antes de executar este script, é obrigatório conectar-se ao Exchange Online.
   Exemplo de conexão:

   Import-Module ExchangeOnlineManagement
   Connect-ExchangeOnline -UserPrincipalName admin@seudominio.com

#>

# Importa os usuários do CSV (coluna DisplayName)
$usersToActive = Import-Csv -Path "C:\temp\E3_Mailboxes.csv" | Select-Object -ExpandProperty DisplayName

# Habilita o Archive Mailbox e aplica a política de retenção
foreach ($user in $usersToActive) {
    Enable-Mailbox -Identity $user -Archive
    Set-Mailbox -Identity $user -RetentionPolicy "PoliticaRetencao6Meses"
}
Write-Host "Archive Mailbox habilitado e política de retenção aplicada para os usuários listados no CSV."
