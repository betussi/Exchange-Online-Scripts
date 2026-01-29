<# 
PRÉ-REQUISITOS:

1) Este script deve ser executado APÓS a conexão com o Exchange Online.

2) Para conectar ao Exchange Online, utilize:
   
   Import-Module ExchangeOnlineManagement
   Connect-ExchangeOnline -UserPrincipalName admin@seudominio.com

3) O script irá:
   - Verificar todas as mailboxes do Exchange Online
   - Ignorar mailboxes com quota ilimitada
   - Calcular o percentual de uso da mailbox
   - Gerar um relatório apenas das mailboxes que atingirem ou ultrapassarem
     o percentual definido na variável $threshold
   - Exportar o relatório para um arquivo CSV

4) O relatório será salvo em:
   C:\Temp\Mailboxes_75PercentOrMore.csv

#>

# Percentual mínimo de uso da mailbox para entrar no relatório
$threshold = 75

# Caminho do arquivo CSV de saída
$csvPath   = "C:\Temp\Mailboxes_75PercentOrMore.csv"

$result = Get-Mailbox -ResultSize Unlimited | ForEach-Object {

    $mbx = $_

    # Ignorar quota ilimitada
    if ($mbx.ProhibitSendReceiveQuota -eq "Unlimited") { return }

    # Estatísticas da mailbox usando GUID (mais confiável)
    $stats = Get-MailboxStatistics -Identity $mbx.ExchangeGuid -ErrorAction SilentlyContinue
    if (-not $stats) { return }

    # ===== USO (BYTES) =====
    $usedBytes = (
        $stats.TotalItemSize.ToString() -replace '.*\(| bytes\)',''
    ) -replace '[^\d]',''

    $usedMB = [math]::Round(([double]$usedBytes / 1MB), 2)

    # ===== QUOTA (BYTES) =====
    $quotaBytes = (
        $mbx.ProhibitSendReceiveQuota.ToString() -replace '.*\(| bytes\)',''
    ) -replace '[^\d]',''

    $quotaMB = [math]::Round(([double]$quotaBytes / 1MB), 2)

    if ($quotaMB -gt 0) {
        $percentUsed = [math]::Round(($usedMB / $quotaMB) * 100, 2)

        if ($percentUsed -ge $threshold) {
            [PSCustomObject]@{
                DisplayName    = $mbx.DisplayName
                UPN            = $mbx.UserPrincipalName
                UsedMB         = $usedMB
                QuotaMB        = $quotaMB
                PercentUsed    = $percentUsed
                ArchiveEnabled = if ($mbx.ArchiveStatus -eq "Active") { "Yes" } else { "No" }
            }
        }
    }
}

# Ordenar por percentual de uso (decrescente) e exportar
$result |
Sort-Object PercentUsed -Descending |
Export-Csv $csvPath -NoTypeInformation -Encoding UTF8

Write-Host "Relatório exportado para $csvPath"
