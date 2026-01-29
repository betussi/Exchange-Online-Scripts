$threshold = 75
$csvPath   = "C:\Temp\Mailboxes_75PercentOrMore.csv"

$result = Get-Mailbox -ResultSize Unlimited | ForEach-Object {

    $mbx = $_

    # Ignorar quota ilimitada
    if ($mbx.ProhibitSendReceiveQuota -eq "Unlimited") { return }

    # Estatísticas com ID único
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

# Ordenar e exportar
$result |
Sort-Object PercentUsed -Descending |
Export-Csv $csvPath -NoTypeInformation -Encoding UTF8

Write-Host "Relatório exportado para $csvPath"