<#
  Desfaz renomeacoes/movimentos registrados em registro.csv (do mais novo para o mais antigo).
  - Sem -Aplicar = so mostra o que faria.
  - Nunca sobrescreve: se o nome original estiver ocupado, restaura como "<nome>-restaurado-2" etc.
  - Arquivos desfeitos entram no registro como DESFEITO e o renomeador nao mexe neles de novo.
  - -Desde '2026-11-25 10:00' desfaz so o que veio depois desse horario.
#>
[CmdletBinding()]
param(
    [switch]$Aplicar,
    [string]$Desde
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$registro = Join-Path $PSScriptRoot 'registro.csv'
if (-not (Test-Path -LiteralPath $registro)) { Write-Host 'Nada no registro.'; return }

function Registrar($acao, $orig, $novo, $motivo) {
    [pscustomobject]@{
        DataHora = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Acao = $acao; Original = $orig; Novo = $novo; Motivo = $motivo
    } | Export-Csv -LiteralPath $registro -Append -NoTypeInformation -Encoding UTF8
}

$linhas = @(Import-Csv -LiteralPath $registro -Encoding UTF8 | Where-Object { $_.Acao -in 'RENOMEAR', 'MOVER' })
if ($Desde) { $corte = [datetime]$Desde; $linhas = @($linhas | Where-Object { [datetime]$_.DataHora -ge $corte }) }
[array]::Reverse($linhas)

foreach ($l in $linhas) {
    if (-not (Test-Path -LiteralPath $l.Novo)) { Write-Host "pulado (nao encontrado): $($l.Novo)"; continue }

    $destino = $l.Original
    if (Test-Path -LiteralPath $destino) {
        $dir = Split-Path $destino -Parent
        $b = [IO.Path]::GetFileNameWithoutExtension($destino); $e = [IO.Path]::GetExtension($destino)
        $n = 2
        $destino = Join-Path $dir ('{0}-restaurado-{1}{2}' -f $b, $n, $e)
        while (Test-Path -LiteralPath $destino) { $n++; $destino = Join-Path $dir ('{0}-restaurado-{1}{2}' -f $b, $n, $e) }
    }
    Write-Host ("{0}  ->  {1}" -f $l.Novo, $destino)
    if ($Aplicar) {
        [IO.File]::Move($l.Novo, $destino)
        Registrar 'DESFEITO' $l.Novo $destino 'desfeito pelo usuario'
    }
}
if (-not $Aplicar) { Write-Host "`n(simulacao: rode com -Aplicar para desfazer de verdade)" }
