<#
  Renomeia arquivos da pasta Downloads pelo CONTEUDO:  "Nome - Tipo - Data.ext"

  - Sem parametros = SIMULACAO: nada e alterado, so mostra o que faria.
  - Com -Aplicar   = renomeia de verdade.
  - Nunca apaga. Nunca sobrescreve (acrescenta -2, -3...).
  - Ignora downloads incompletos (.crdownload, .tmp, .part), instaladores, pastas e arquivos em uso.
  - O que nao for entendido vai para Downloads\Conferir.
  - Tudo fica em registro.csv (original -> novo) para o desfazer.ps1.
#>
[CmdletBinding()]
param(
    [switch]$Aplicar,
    [string]$Pasta,
    [string]$Config
)

$ErrorActionPreference = 'Stop'
$raiz = $PSScriptRoot
if (-not $Config) { $Config = Join-Path $raiz 'regras.json' }
if (-not $Pasta)  { $Pasta  = Join-Path $env:USERPROFILE 'Downloads' }
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName System.IO.Compression.FileSystem

$cfg           = Get-Content -LiteralPath $Config -Raw -Encoding UTF8 | ConvertFrom-Json
$pastaConferir = Join-Path $Pasta 'Conferir'
$registro      = Join-Path $raiz 'registro.csv'
$pastaRel      = Join-Path $raiz 'relatorios'

# ---------- utilitarios ----------
function Normaliza-Leve([string]$s) {
    if ([string]::IsNullOrEmpty($s)) { return '' }
    $d = $s.ToLowerInvariant().Normalize([Text.NormalizationForm]::FormD)
    [regex]::Replace($d, '\p{Mn}', '')
}

function Simplifica([string]$s) {
    $d = Normaliza-Leve $s
    $d = [regex]::Replace($d, '[_\-\.]+', ' ')
    [regex]::Replace($d, '\s+', ' ')
}

function Limpa-Nome([string]$s) {
    $s = [regex]::Replace($s, '[\\/:*?"<>|]', ' ')
    $s = [regex]::Replace($s, '\s+', ' ')
    $s.Trim(' ', '.')
}

function Ler-ZipTexto([string]$caminho, [string]$entrada) {
    $zip = [IO.Compression.ZipFile]::OpenRead($caminho)
    try {
        $e = $zip.Entries | Where-Object { $_.FullName -eq $entrada } | Select-Object -First 1
        if (-not $e) { return '' }
        $sr = New-Object IO.StreamReader($e.Open(), [Text.Encoding]::UTF8)
        try { $sr.ReadToEnd() } finally { $sr.Dispose() }
    } finally { $zip.Dispose() }
}

function Xml-Texto([string]$xml) {
    $t = [regex]::Replace($xml, '<[^>]+>', ' ')
    [System.Net.WebUtility]::HtmlDecode($t)
}

# ---------- leitor de PDF (poppler / pdftotext) ----------
$script:pdftotext = $null
$cmd = Get-Command pdftotext.exe -ErrorAction SilentlyContinue
if ($cmd) { $script:pdftotext = $cmd.Source }
elseif ($cfg.caminho_pdftotext -and (Test-Path -LiteralPath $cfg.caminho_pdftotext)) { $script:pdftotext = $cfg.caminho_pdftotext }

function Obter-Texto([IO.FileInfo]$f) {
    $ext = $f.Extension.ToLowerInvariant()
    $t = ''
    try {
        if ($ext -eq '.pdf') {
            $t = (& $script:pdftotext -l 3 -enc UTF-8 $f.FullName - 2>$null) -join "`n"
        }
        elseif ($ext -eq '.docx') {
            $t = Xml-Texto (Ler-ZipTexto $f.FullName 'word/document.xml')
        }
        elseif ($ext -eq '.xlsx') {
            $t = Xml-Texto (Ler-ZipTexto $f.FullName 'xl/sharedStrings.xml')
            $wb = Ler-ZipTexto $f.FullName 'xl/workbook.xml'
            foreach ($m in [regex]::Matches($wb, '<sheet [^>]*name="([^"]+)"')) { $t += ' ' + $m.Groups[1].Value }
        }
        elseif ($ext -eq '.txt' -or $ext -eq '.csv') {
            $t = (Get-Content -LiteralPath $f.FullName -TotalCount 80 -Encoding UTF8) -join "`n"
        }
    } catch { $t = '' }
    if ($null -eq $t) { $t = '' }
    if ($t.Length -gt 8000) { $t = $t.Substring(0, 8000) }
    $t
}

# ---------- data ----------
$mesNum = @{
    janeiro = 1; fevereiro = 2; marco = 3; abril = 4; maio = 5; junho = 6
    julho = 7; agosto = 8; setembro = 9; outubro = 10; novembro = 11; dezembro = 12
    jan = 1; fev = 2; mar = 3; abr = 4; mai = 5; jun = 6; jul = 7; ago = 8; set = 9; out = 10; nov = 11; dez = 12
}

function Achar-Data([string]$bruto) {
    $s = Normaliza-Leve $bruto
    $achados = New-Object System.Collections.Generic.List[object]

    foreach ($m in [regex]::Matches($s, '(?<!\d)(\d{1,2})[/\-.](\d{1,2})[/\-.](20\d{2})(?!\d)')) {
        $d = [int]$m.Groups[1].Value; $mo = [int]$m.Groups[2].Value; $a = [int]$m.Groups[3].Value
        if ($d -ge 1 -and $d -le 31 -and $mo -ge 1 -and $mo -le 12) {
            $achados.Add([pscustomobject]@{ I = $m.Index; V = ('{0:0000}-{1:00}-{2:00}' -f $a, $mo, $d) })
        }
    }
    foreach ($m in [regex]::Matches($s, '(?<!\d)(20\d{2})-(\d{2})-(\d{2})(?!\d)')) {
        $a = [int]$m.Groups[1].Value; $mo = [int]$m.Groups[2].Value; $d = [int]$m.Groups[3].Value
        if ($d -ge 1 -and $d -le 31 -and $mo -ge 1 -and $mo -le 12) {
            $achados.Add([pscustomobject]@{ I = $m.Index; V = ('{0:0000}-{1:00}-{2:00}' -f $a, $mo, $d) })
        }
    }
    $nomes = ($mesNum.Keys | Sort-Object { $_.Length } -Descending) -join '|'
    foreach ($m in [regex]::Matches($s, "\b($nomes)\.?\s*(?:de|/|-)?\s*(20\d{2})(?!\d)")) {
        $mo = $mesNum[$m.Groups[1].Value]; $a = [int]$m.Groups[2].Value
        $achados.Add([pscustomobject]@{ I = $m.Index; V = ('{0:0000}-{1:00}' -f $a, $mo) })
    }
    foreach ($m in [regex]::Matches($s, '(?<!\d)(0?[1-9]|1[0-2])/(20\d{2})(?!\d)')) {
        $mo = [int]$m.Groups[1].Value; $a = [int]$m.Groups[2].Value
        $achados.Add([pscustomobject]@{ I = $m.Index; V = ('{0:0000}-{1:00}' -f $a, $mo) })
    }
    if ($achados.Count -eq 0) { return $null }
    ($achados | Sort-Object I | Select-Object -First 1).V
}

# ---------- nome sem colisao ----------
$reservados = New-Object 'System.Collections.Generic.HashSet[string]'
function Caminho-Livre([string]$dir, [string]$base, [string]$ext) {
    $p = Join-Path $dir ($base + $ext)
    $n = 2
    while ((Test-Path -LiteralPath $p) -or $reservados.Contains($p.ToLowerInvariant())) {
        $p = Join-Path $dir ('{0}-{1}{2}' -f $base, $n, $ext)
        $n++
    }
    [void]$reservados.Add($p.ToLowerInvariant())
    $p
}

function Registrar($acao, $orig, $novo, $motivo) {
    [pscustomobject]@{
        DataHora = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Acao = $acao; Original = $orig; Novo = $novo; Motivo = $motivo
    } | Export-Csv -LiteralPath $registro -Append -NoTypeInformation -Encoding UTF8
}

function Em-Uso([string]$caminho) {
    try {
        $fs = [IO.File]::Open($caminho, 'Open', 'Read', 'None')
        $fs.Dispose(); $false
    } catch { $true }
}

# ---------- arquivos que o usuario desfez: nao mexer de novo ----------
$naoMexer = @{}
if (Test-Path -LiteralPath $registro) {
    foreach ($r in (Import-Csv -LiteralPath $registro -Encoding UTF8)) {
        if ($r.Acao -eq 'DESFEITO') { $naoMexer[$r.Novo.ToLowerInvariant()] = $true }
    }
}

# ---------- principal ----------
if (-not (Test-Path -LiteralPath $Pasta)) { throw "Pasta nao encontrada: $Pasta" }

$ignorarExt = @($cfg.ignorar_extensoes | ForEach-Object { $_.ToLowerInvariant() })
$ocultos = [IO.FileAttributes]::Hidden -bor [IO.FileAttributes]::System
$agora = Get-Date
$linhas = New-Object System.Collections.Generic.List[object]
$ignorados = @{}
function Ignora($motivo) { if ($ignorados.ContainsKey($motivo)) { $ignorados[$motivo]++ } else { $ignorados[$motivo] = 1 } }

$arquivos = Get-ChildItem -LiteralPath $Pasta -File -Force | Sort-Object LastWriteTime

foreach ($f in $arquivos) {
    try {
        $ext = $f.Extension.ToLowerInvariant()
        $base = [IO.Path]::GetFileNameWithoutExtension($f.Name)

        if ($f.Attributes -band $ocultos)                    { Ignora 'oculto/sistema'; continue }
        if ($ignorarExt -contains $ext)                      { Ignora 'extensao ignorada (incompleto/instalador/compactado)'; continue }
        $pad = $false
        foreach ($p in $cfg.ignorar_padroes) { if ($f.Name -like $p) { $pad = $true; break } }
        if ($pad)                                            { Ignora 'nome ja padronizado (regra do usuario)'; continue }
        if ($base -match '^.+ - .+ - \d{4}(-\d{2}){1,2}(-\d+)?$') { Ignora 'ja esta no padrao'; continue }
        if ($naoMexer.ContainsKey($f.FullName.ToLowerInvariant())) { Ignora 'desfeito pelo usuario'; continue }

        $ultima = if ($f.CreationTime -gt $f.LastWriteTime) { $f.CreationTime } else { $f.LastWriteTime }
        if (($agora - $ultima).TotalMinutes -lt [double]$cfg.minutos_minimos) { Ignora 'muito recente (pode estar baixando)'; continue }
        if (Em-Uso $f.FullName)                              { Ignora 'em uso por outro programa'; continue }
        if ($ext -eq '.pdf' -and -not $script:pdftotext)     { Ignora 'PDF sem leitor (instale o poppler)'; continue }

        $texto = Obter-Texto $f
        $busca = Simplifica ($f.Name + ' ' + $texto)

        # pessoal -> Conferir
        $motivoConferir = $null
        foreach ($w in $cfg.pessoal) { if ($busca.Contains($w)) { $motivoConferir = "possivel arquivo pessoal ('$w')"; break } }

        $entidade = $null; $tipo = $null; $data = $null
        if (-not $motivoConferir) {
            foreach ($e in $cfg.entidades) {
                foreach ($w in $e.palavras) { if ($busca.Contains($w)) { $entidade = $e.nome; break } }
                if ($entidade) { break }
            }
            foreach ($t in $cfg.tipos) {
                foreach ($w in $t.palavras) { if ($busca.Contains($w)) { $tipo = $t.tipo; break } }
                if ($tipo) { break }
            }
            $data = Achar-Data ($f.Name + "`n" + $texto)

            $falta = @()
            if (-not $entidade) { $falta += 'nome' }
            if (-not $tipo)     { $falta += 'tipo' }
            if (-not $data)     { $falta += 'data' }
            if ($falta.Count -gt 0) {
                if ([string]::IsNullOrWhiteSpace($texto)) { $motivoConferir = 'sem texto legivel (imagem/PDF escaneado?)' }
                else { $motivoConferir = 'nao identificado: ' + ($falta -join ', ') }
            }
        }

        if ($motivoConferir) {
            $novo = Caminho-Livre $pastaConferir ([IO.Path]::GetFileNameWithoutExtension($f.Name)) $f.Extension
            $acao = 'CONFERIR'; $motivo = $motivoConferir
            if ($Aplicar) {
                if (-not (Test-Path -LiteralPath $pastaConferir)) { New-Item -ItemType Directory -Path $pastaConferir | Out-Null }
                [IO.File]::Move($f.FullName, $novo)
                Registrar 'MOVER' $f.FullName $novo $motivo
            }
        } else {
            $novoBase = Limpa-Nome ("$entidade - $tipo - $data")
            if ($novoBase.Length -gt 120) { $novoBase = $novoBase.Substring(0, 120).Trim() }
            $novo = Caminho-Livre $Pasta $novoBase $ext
            $acao = 'RENOMEAR'; $motivo = ''
            if ($Aplicar) {
                [IO.File]::Move($f.FullName, $novo)
                Registrar 'RENOMEAR' $f.FullName $novo ''
            }
        }
        $linhas.Add([pscustomobject]@{ Acao = $acao; Original = $f.Name; Novo = (Split-Path $novo -Leaf); Pasta = (Split-Path $novo -Parent | Split-Path -Leaf); Motivo = $motivo })
    }
    catch {
        $linhas.Add([pscustomobject]@{ Acao = 'ERRO'; Original = $f.Name; Novo = ''; Pasta = ''; Motivo = $_.Exception.Message })
        if ($Aplicar) { Registrar 'ERRO' $f.FullName '' $_.Exception.Message }
    }
}

# ---------- relatorio ----------
$modo = if ($Aplicar) { 'RODADA (aplicada)' } else { 'SIMULACAO (nada foi alterado)' }
$out = New-Object System.Collections.Generic.List[string]
$out.Add("$modo - $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))")
$out.Add("Pasta: $Pasta")
$out.Add(("Renomeados: {0} | Para Conferir: {1} | Erros: {2}" -f
    @($linhas | Where-Object Acao -eq 'RENOMEAR').Count,
    @($linhas | Where-Object Acao -eq 'CONFERIR').Count,
    @($linhas | Where-Object Acao -eq 'ERRO').Count))
if ($ignorados.Count -gt 0) {
    $out.Add('Ignorados:')
    foreach ($k in $ignorados.Keys) { $out.Add(("  {0} x {1}" -f $ignorados[$k], $k)) }
}
if ($linhas.Count -gt 0) {
    $out.Add('')
    $out.Add(($linhas | Format-Table Acao, Original, Novo, Pasta, Motivo -AutoSize -Wrap | Out-String -Width 250))
}
if (-not $Aplicar) {
    $ex = @($linhas | Where-Object Acao -eq 'RENOMEAR' | Select-Object -First 5)
    $out.Add('5 EXEMPLOS de renomeacao:')
    if ($ex.Count -eq 0) { $out.Add('  (nenhum arquivo teria nome identificado - ajuste regras.json)') }
    foreach ($x in $ex) { $out.Add(("  {0}  ->  {1}" -f $x.Original, $x.Novo)) }
}

$texto = $out -join "`r`n"
Write-Host $texto

if (-not (Test-Path -LiteralPath $pastaRel)) { New-Item -ItemType Directory -Path $pastaRel | Out-Null }
if (-not $Aplicar) {
    Set-Content -LiteralPath (Join-Path $pastaRel 'simulacao-ultima.txt') -Value $texto -Encoding UTF8
} elseif ($linhas.Count -gt 0) {
    $nomeRel = 'rodada-{0}.txt' -f (Get-Date).ToString('yyyyMMdd-HHmmss')
    Set-Content -LiteralPath (Join-Path $pastaRel $nomeRel) -Value $texto -Encoding UTF8
}
