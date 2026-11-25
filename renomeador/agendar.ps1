<#
  Agenda o renomeador no Agendador de Tarefas do Windows (roda com -Aplicar).
  So rode depois de ter conferido a simulacao.
    .\agendar.ps1                  -> a cada 15 minutos
    .\agendar.ps1 -CadaMinutos 30
    .\agendar.ps1 -Remover         -> cancela o agendamento
#>
[CmdletBinding()]
param(
    [int]$CadaMinutos = 15,
    [switch]$Remover
)

$nome = 'Renomeador Downloads'
if ($Remover) {
    Unregister-ScheduledTask -TaskName $nome -Confirm:$false
    Write-Host 'Agendamento removido.'
    return
}

$script = Join-Path $PSScriptRoot 'renomear-downloads.ps1'
$acao = New-ScheduledTaskAction -Execute 'powershell.exe' `
    -Argument ('-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}" -Aplicar' -f $script)
$gatilho = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) `
    -RepetitionInterval (New-TimeSpan -Minutes $CadaMinutos) -RepetitionDuration (New-TimeSpan -Days 3650)
$ajustes = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 10)

Register-ScheduledTask -TaskName $nome -Action $acao -Trigger $gatilho -Settings $ajustes `
    -Description 'Renomeia arquivos novos da pasta Downloads pelo conteudo' -Force | Out-Null
Write-Host "Agendado: a cada $CadaMinutos min. Relatorios em $(Join-Path $PSScriptRoot 'relatorios')"
