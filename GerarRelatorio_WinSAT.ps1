# ==============================================================================
# Gerador de Relatório WinSAT + Battery + Software - Versão: 1.0
# Autor: Robson Nunes - Cyber Security
# Função: Executa teste, analisa XMLs, battery report e gera dashboard HTML.
# ==============================================================================

# 1. Verificação de Privilégios
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "❌ ERRO: Este script precisa ser executado como ADMINISTRADOR." -ForegroundColor Red
    Write-Host "💡 Clique com o botão direito no PowerShell e escolha 'Executar como Administrador'." -ForegroundColor Yellow
    Start-Sleep -Seconds 5
    exit
}

# ✅ Limpar tela e configurar modo silencioso
Clear-Host
$ErrorActionPreference = 'SilentlyContinue'

# 2. Configurações Iniciais
# ✅ Caminho fixo em C: com criação automática e tratamento de erro
$folderPath = "C:\Relatorios"
if (-not (Test-Path $folderPath)) { 
    try {
        New-Item -ItemType Directory -Path $folderPath -Force -ErrorAction Stop | Out-Null 
        Write-Host "📁 Pasta de saída criada: $folderPath" -ForegroundColor Cyan
    } catch {
        Write-Host "❌ Erro: Não foi possível criar a pasta $folderPath. Verifique as permissões." -ForegroundColor Red
        exit
    }
}

$outputPath = Join-Path $folderPath "Relatorio_Completo_Performance.html"
$basePath = Join-Path $env:SystemRoot "Performance\WinSAT\DataStore"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$batteryReportPath = Join-Path $folderPath "BatteryReport_$timestamp.html"

$outputPath = Join-Path $folderPath "Relatorio_Completo_Performance.html"
# Caminho portável usando variável de ambiente
$basePath = Join-Path $env:SystemRoot "Performance\WinSAT\DataStore"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$batteryReportPath = Join-Path $folderPath "BatteryReport_$timestamp.html"

# 3. Execução do WinSAT Formal
Write-Host "🚀 Iniciando avaliação oficial do Windows (WinSAT Formal)..." -ForegroundColor Green
Write-Host "⏳ Isso levará aproximadamente 2 minutos. Por favor, aguarde..." -ForegroundColor Yellow

# Validar se executável WinSAT existe
$winsatPath = "$env:SystemRoot\System32\winsat.exe"
if (-not (Test-Path $winsatPath)) {
    Write-Host "❌ WinSAT não encontrado neste sistema." -ForegroundColor Red
    exit
}

try {
    # Caminho explícito para evitar dependência do PATH
    $process = Start-Process -FilePath $winsatPath -ArgumentList "formal" -Wait -PassThru -NoNewWindow
    
    if ($process.ExitCode -ne 0) {
        throw "O teste WinSAT falhou com código de erro $($process.ExitCode)."
    }
    Write-Host "✅ Teste WinSAT concluído com sucesso!" -ForegroundColor Green
} catch {
    Write-Host "❌ Erro ao executar winsat formal: $_" -ForegroundColor Red
    exit
}

# ✅ AJUSTE: Pequena pausa para garantir que XMLs foram gravados
Write-Host "⏳ Aguardando finalização da gravação dos arquivos..." -ForegroundColor Cyan
Start-Sleep -Seconds 2

# Verificar se arquivos XML foram gerados após execução
Write-Host "🔍 Verificando arquivos gerados pelo WinSAT..." -ForegroundColor Cyan
if (-not (Test-Path $basePath)) {
    Write-Host "❌ Erro: diretório WinSAT não encontrado após execução: $basePath" -ForegroundColor Red
    exit
}

$xmlCheck = Get-ChildItem "$basePath\*.xml" -ErrorAction SilentlyContinue | 
    Where-Object { $_.Name -like "*Formal*" } | 
    Select-Object -First 1

if (-not $xmlCheck) {
    Write-Host "❌ Erro: WinSAT executou mas não gerou arquivo Formal.Assessment." -ForegroundColor Red
    exit
}

# 3.1 - Geração do Battery Report
Write-Host "🔋 Gerando relatório de bateria (powercfg)..." -ForegroundColor Cyan
try {
    if (Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue) {
        powercfg /batteryreport /output "$batteryReportPath" | Out-Null
        Write-Host "✅ Battery report gerado: $batteryReportPath" -ForegroundColor Green
        $hasBattery = $true
    } else {
        Write-Host "⚠️ Nenhuma bateria detectada (Desktop?). Pulando battery report." -ForegroundColor Yellow
        $hasBattery = $false
    }
} catch {
    Write-Host "⚠️ Não foi possível gerar battery report: $_" -ForegroundColor Yellow
    $hasBattery = $false
}

# 3.2 - Inventário de Software Instalado
Write-Host "📦 Coletando inventário de software instalado..." -ForegroundColor Cyan

# ErrorAction SilentlyContinue para evitar erro em chaves corrompidas
# Conversão de InstallDate para formato legível
# Inclusão de InstallDate na seleção
$softwareList = Get-ItemProperty `
    HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*, `
    HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* `
    -ErrorAction SilentlyContinue | 
    Where-Object { $_.DisplayName } | 
    Select-Object DisplayName, DisplayVersion, Publisher, 
        @{Name="InstallDate";Expression={
            if ($_.InstallDate) {
                try {
                    [datetime]::ParseExact($_.InstallDate, 'yyyyMMdd', $null).ToString("dd/MM/yyyy")
                } catch { $_.InstallDate }
            }
        }} |
    Sort-Object DisplayName -Unique  # Remove duplicatas

# ✅ AJUSTE: Limitar inventário de software para evitar HTML muito grande
if ($softwareList.Count -gt 500) {
    Write-Host "⚠️ Inventário grande ($($softwareList.Count) itens). Limitando a 500 para melhor performance." -ForegroundColor Yellow
    $softwareList = $softwareList | Select-Object -First 500
    $softwareCountLimited = $true
} else {
    $softwareCountLimited = $false
}

# Contagem segura com Measure-Object
$softwareCount = ($softwareList | Measure-Object).Count
Write-Host "✅ $softwareCount aplicativos catalogados" -ForegroundColor Green

# 4. Localização dos Arquivos XML
Write-Host "🔍 Lendo arquivos de dados gerados..." -ForegroundColor Cyan

# Leitura defensiva com -ErrorAction SilentlyContinue
$xmlFiles = Get-ChildItem "$basePath\*.xml" -ErrorAction SilentlyContinue | 
    Sort-Object LastWriteTime -Descending | 
    Select-Object -First 20

# Valida se arquivos XML foram encontrados
if (-not $xmlFiles) {
    Write-Host "❌ Erro: nenhum arquivo XML do WinSAT encontrado." -ForegroundColor Red
    exit
}

$formalFile = $xmlFiles | Where-Object { $_.Name -like "*Formal*" } | Select-Object -First 1
$dwmFile = $xmlFiles | Where-Object { $_.Name -like "*DWM*" } | Select-Object -First 1

if (-not $formalFile) {
    Write-Host "❌ Erro crítico: Arquivo Formal.Assessment não encontrado após o teste." -ForegroundColor Red
    exit
}

# Proteção adicional ao carregar XML com try/catch
try {
    [xml]$formalXml = Get-Content $formalFile.FullName -Raw -ErrorAction Stop
} catch {
    Write-Host "❌ Erro ao carregar XML Formal.Assessment" -ForegroundColor Red
    exit
}

[xml]$dwmXml = $null
if ($dwmFile) {
    try {
        [xml]$dwmXml = Get-Content $dwmFile.FullName -Raw -ErrorAction Stop
    } catch {
        Write-Host "⚠️ Aviso: não foi possível carregar XML DWM" -ForegroundColor Yellow
    }
}

# 5. Função de Extração (XML já carregado - otimizado)
function Get-XmlValueFromXml($xml, $tagName) {
    try {
        if ($xml) {
            $node = $xml.SelectSingleNode("//$tagName")
            if ($node -and $node.InnerText) {
                return [double]$node.InnerText
            }
        }
    }
    catch {
        Write-Host "⚠️ Aviso: erro ao extrair valor $tagName do XML" -ForegroundColor Yellow
    }
    
    return $null
}

function Get-ModeloDisco($xml) {
    try {
        if ($xml) {
            # Tenta encontrar via XPath direto
            $modelNode = $xml.SelectSingleNode("//Model")
            if ($modelNode -and $modelNode.InnerText) {
                return $modelNode.InnerText.Trim()
            }
        }
    } catch {
        Write-Host "⚠️ Aviso: erro ao ler modelo do disco" -ForegroundColor Yellow
    }
    return "Não detectado"
}

# 6. Extração dos Dados
Write-Host "📊 Processando métricas..." -ForegroundColor Gray

# Scores Gerais (Verificação explícita contra $null)
$sysScoreRaw = Get-XmlValueFromXml $formalXml "SystemScore"
$sysScore = if ($sysScoreRaw -ne $null) { [math]::Round($sysScoreRaw, 2) } else { "N/A" }

$cpuScoreRaw = Get-XmlValueFromXml $formalXml "CpuScore"
$cpuScore = if ($cpuScoreRaw -ne $null) { [math]::Round($cpuScoreRaw, 2) } else { "N/A" }

$memScoreRaw = Get-XmlValueFromXml $formalXml "MemoryScore"
$memScore = if ($memScoreRaw -ne $null) { [math]::Round($memScoreRaw, 2) } else { "N/A" }

$gfxScoreRaw = Get-XmlValueFromXml $formalXml "GraphicsScore"
$gfxScore = if ($gfxScoreRaw -ne $null) { [math]::Round($gfxScoreRaw, 2) } else { "N/A" }

$diskScoreRaw = Get-XmlValueFromXml $formalXml "DiskScore"
$diskScore = if ($diskScoreRaw -ne $null) { [math]::Round($diskScoreRaw, 2) } else { "N/A" }

# Detalhes do Disco
$diskSeqReadRaw = Get-XmlValueFromXml $formalXml "AvgThroughput"
if ($diskSeqReadRaw -ne $null -and $diskSeqReadRaw -gt 0 -and $diskSeqReadRaw -lt 2000) {
    $diskSeqRead = [math]::Round($diskSeqReadRaw, 2)
} else {
    $diskSeqRead = "N/A"
}

$diskRandRead = "N/A"
$diskModelRaw = Get-ModeloDisco $formalXml
# Escapar modelo do disco no HTML
$diskModel = ConvertTo-HtmlEncode -Text $diskModelRaw

# Detalhes da Memória
$memBandwidthRaw = Get-XmlValueFromXml $formalXml "Bandwidth"
if ($memBandwidthRaw -ne $null -and $memBandwidthRaw -gt 0) {
    $memBandwidth = [math]::Round($memBandwidthRaw, 2)
} else {
    $memBandwidth = "N/A"
}

# Detalhes da CPU
$cpuEncryptRaw = Get-XmlValueFromXml $formalXml "EncryptionMetric"
$cpuCompressRaw = Get-XmlValueFromXml $formalXml "CompressionMetric"
$cpuEncrypt = if ($cpuEncryptRaw -ne $null) { [math]::Round($cpuEncryptRaw, 2) } else { "N/A" }
$cpuCompress = if ($cpuCompressRaw -ne $null) { [math]::Round($cpuCompressRaw, 2) } else { "N/A" }

# Detalhes de Gráficos
$gfxFps = "N/A"
$gfxBandwidth = "N/A"
if ($dwmXml) {
    $gfxFpsRaw = Get-XmlValueFromXml $dwmXml "FPS"
    $gfxBandwidthRaw = Get-XmlValueFromXml $dwmXml "MbVideoMemPerSecond"
    if ($gfxFpsRaw -ne $null) { $gfxFps = [math]::Round($gfxFpsRaw, 2) }
    if ($gfxBandwidthRaw -ne $null) { $gfxBandwidth = [math]::Round($gfxBandwidthRaw, 2) }
}

# 6.1 - Informações do Sistema
Write-Host "🖥️ Coletando informações do sistema..." -ForegroundColor Cyan

# Proteção adicional para consultas CIM de hardware
$osInfo = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
if (-not $osInfo) {
    Write-Host "⚠️ Aviso: não foi possível coletar informações do SO" -ForegroundColor Yellow
}

# Garante apenas um processador (multi-socket/virtualização)
$cpuInfo = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
# Verificação defensiva para CPU
$cpuName = if ($cpuInfo) { ConvertTo-HtmlEncode -Text $cpuInfo.Name } else { "Não detectado" }

# ✅ AJUSTE: Proteção adicional para cálculo de memória RAM com fallback explícito
$ramModules = Get-CimInstance Win32_PhysicalMemory -ErrorAction SilentlyContinue
$ramSum = if ($ramModules) { ($ramModules | Measure-Object Capacity -Sum).Sum } else { 0 }
$ramInfo = if ($ramSum -gt 0) {
    [math]::Round($ramSum / 1GB, 2)
} else {
    "N/A"
}

# ✅ AJUSTE: Verificação defensiva para modelo do computador
$computerSystem = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
$computerModel = if ($computerSystem) { $computerSystem.Model } else { "Não detectado" }
$computerModelEscaped = ConvertTo-HtmlEncode -Text $computerModel

# Verificação defensiva para BIOS
$biosInfo = Get-CimInstance Win32_BIOS -ErrorAction SilentlyContinue | Select-Object -First 1
$biosVersion = if ($biosInfo) { $biosInfo.SMBIOSBIOSVersion } else { "Não detectado" }

# Verificação defensiva para GPU
$gpuInfo = Get-CimInstance Win32_VideoController -ErrorAction SilentlyContinue | Select-Object -First 1
if ($gpuInfo) {
    $gpuName = ConvertTo-HtmlEncode -Text $gpuInfo.Name
} else {
    $gpuName = "Não detectado"
}

$tpmInfo = Get-CimInstance -Namespace "root\cimv2\Security\MicrosoftTpm" -ClassName Win32_Tpm -ErrorAction SilentlyContinue

# Tempo de atividade (uptime)
if ($osInfo -and $osInfo.LastBootUpTime) {
    $uptime = (Get-Date) - $osInfo.LastBootUpTime
    $uptimeString = "{0} dias, {1} horas, {2} minutos" -f $uptime.Days, $uptime.Hours, $uptime.Minutes
} else {
    $uptimeString = "N/A"
}

# ✅ AJUSTE: Fallback seguro para valores do sistema operacional
$osCaption = if ($osInfo -and $osInfo.Caption) { ConvertTo-HtmlEncode -Text $osInfo.Caption } else { "Windows" }
$osArchitecture = if ($osInfo -and $osInfo.OSArchitecture) { $osInfo.OSArchitecture } else { "Desconhecida" }
$osVersion = if ($osInfo -and $osInfo.Version) { $osInfo.Version } else { "N/A" }
$osBuildNumber = if ($osInfo -and $osInfo.BuildNumber) { $osInfo.BuildNumber } else { "N/A" }

Write-Host "✅ Dados extraídos:" -ForegroundColor Green
Write-Host "   Sistema: $sysScore | Disco: $diskScore ($diskSeqRead MB/s)" -ForegroundColor White

# 6.2 - Coleta de Informações Adicionais
Write-Host "🔍 Coletando informações adicionais..." -ForegroundColor Cyan

# Status do BitLocker
$bitlockerText = "❓ Não disponível"
try {
    $bitlocker = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction SilentlyContinue
    if ($bitlocker) {
        $bitlockerStatus = $bitlocker.ProtectionStatus
        if ($bitlockerStatus -eq 1) { $bitlockerText = "✅ Ativado" }
        elseif ($bitlockerStatus -eq 0) { $bitlockerText = "⚠️ Desativado" }
    }
} catch {
    $bitlockerText = "❓ Não disponível (Windows Home?)"
}

# Status do TPM
if ($tpmInfo) {
    $tpmText = "✅ TPM $(if($tpmInfo.SpecVersion) { $tpmInfo.SpecVersion } else { 'Detectado' })"
} else {
    $tpmText = "⚠️ Não detectado"
}

# Contagem segura de drivers com erro
$driverErrors = (Get-CimInstance Win32_PnPEntity -Filter "ConfigManagerErrorCode <> 0" -ErrorAction SilentlyContinue | Measure-Object).Count
if ($driverErrors -gt 0) {
    $driverText = "❌ $driverErrors com erro"
} else {
    $driverText = "✅ Todos OK"
}

# 7. Lógica de Cores
function Get-Color($s) { 
    if ($s -eq $null -or $s -eq "N/A") { return "#cccccc" }
    
    # Conversão explícita para valor numérico
    $value = [double]$s
    
    if ($value -lt 4.0) { return "#dc3545" } 
    elseif ($value -lt 6.0) { return "#ffc107" } 
    else { return "#28a745" } 
}

$overallColor = Get-Color $sysScore
$diskColor = Get-Color $diskScore

# 8. Geração do HTML COMPLETO
$html = @"
<!DOCTYPE html>
<html lang='pt-br'>
<head>
    <meta charset='UTF-8'>
    <title>Relatório Completo de Diagnóstico</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f4f6f9; color: #333; margin: 0; padding: 20px; }
        .container { max-width: 1100px; margin: 0 auto; background: #fff; padding: 40px; border-radius: 12px; box-shadow: 0 5px 20px rgba(0,0,0,0.08); }
        h1 { text-align: center; color: #2c3e50; margin-bottom: 10px; }
        h2 { color: #34495e; border-bottom: 2px solid #3498db; padding-bottom: 10px; margin-top: 40px; }
        .subtitle { text-align: center; color: #7f8c8d; margin-bottom: 30px; font-size: 0.9em; }
        
        .scores-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 20px; margin-bottom: 40px; }
        .score-card { background: #f8f9fa; border-radius: 10px; padding: 20px; text-align: center; border: 1px solid #e9ecef; transition: transform 0.2s; }
        .score-card:hover { transform: translateY(-3px); box-shadow: 0 5px 15px rgba(0,0,0,0.1); }
        .score-val { font-size: 2.5em; font-weight: 700; display: block; margin: 10px 0; }
        .score-label { text-transform: uppercase; font-size: 0.75em; letter-spacing: 1px; color: #6c757d; font-weight: 600; }

        .section { margin-bottom: 30px; border: 1px solid #eee; border-radius: 8px; overflow: hidden; }
        .section-header { background: #34495e; color: #fff; padding: 15px 20px; font-size: 1.1em; font-weight: 600; display: flex; justify-content: space-between; align-items: center; }
        .section-body { padding: 20px; background: #fff; }
        
        table { width: 100%; border-collapse: collapse; }
        td { padding: 12px 15px; border-bottom: 1px solid #f1f1f1; }
        tr:last-child td { border-bottom: none; }
        .label { font-weight: 600; color: #555; width: 45%; }
        .value { font-family: 'Consolas', monospace; color: #2c3e50; font-weight: 500; text-align: right; }
        .unit { color: #95a5a6; font-size: 0.85em; margin-left: 5px; }

        .alert { padding: 20px; border-radius: 8px; margin-top: 20px; border-left: 6px solid; line-height: 1.6; }
        .alert-crit { background: #fdecea; color: #c0392b; border-color: #e74c3c; }
        .alert-warn { background: #fef9e7; color: #d35400; border-color: #f39c12; }
        .alert-ok { background: #eafaf1; color: #27ae60; border-color: #2ecc71; }

        .software-table { max-height: 400px; overflow-y: auto; }
        .software-table table { font-size: 0.85em; }
        
        .footer { text-align: center; margin-top: 40px; font-size: 0.8em; color: #bdc3c7; border-top: 1px solid #eee; padding-top: 20px; }
        
        .info-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; margin-bottom: 20px; }
        .info-box { background: #f8f9fa; padding: 15px; border-radius: 8px; border-left: 4px solid #3498db; }
        .info-box strong { color: #2c3e50; display: block; margin-bottom: 5px; }
        
        .security-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin-bottom: 20px; }
        .security-box { background: #f8f9fa; padding: 15px; border-radius: 8px; text-align: center; }
        .security-box .status { font-size: 1.2em; font-weight: 600; margin-top: 5px; }
        
        .uptime-badge { background: #3498db; color: #fff; padding: 5px 15px; border-radius: 20px; font-size: 0.85em; display: inline-block; margin-top: 10px; }
    </style>
</head>
<body>
    <div class='container'>
        <h1>📊 Relatório Completo de Diagnóstico</h1>
        <div class='subtitle'>Computador: $($env:COMPUTERNAME) | Modelo: $computerModelEscaped<br>Data: $(Get-Date -Format "dd/MM/yyyy HH:mm")<br>SO: $osCaption $osArchitecture | Versão: $osVersion (Build $osBuildNumber)</div>

        <!-- Informações do Sistema -->
        <h2>💻 Informações do Sistema</h2>
        <div class='info-grid'>
            <div class='info-box'>
                <strong>Processador</strong>
                $cpuName
            </div>
            <div class='info-box'>
                <strong>Memória RAM</strong>
                $ramInfo GB
            </div>
            <div class='info-box'>
                <strong>GPU</strong>
                $gpuName
            </div>
            <div class='info-box'>
                <strong>BIOS</strong>
                $biosVersion
            </div>
        </div>
        <div style='text-align: center;'>
            <div class='uptime-badge'>⏱️ Uptime: $uptimeString</div>
        </div>

        <!-- Status de Segurança -->
        <h2>🔒 Status de Segurança</h2>
        <div class='security-grid'>
            <div class='security-box'>
                <div>BitLocker</div>
                <div class='status'>$bitlockerText</div>
            </div>
            <div class='security-box'>
                <div>TPM</div>
                <div class='status'>$tpmText</div>
            </div>
            <div class='security-box'>
                <div>Drivers</div>
                <div class='status'>$driverText</div>
            </div>
        </div>

        <!-- Scores Principais -->
        <h2>📈 Windows System Assessment Tool (WinSAT)</h2>
        <div class='scores-grid'>
            <div class='score-card'>
                <span class='score-label'>Geral (Gargalo)</span>
                <span class='score-val' style='color:$overallColor'>$sysScore</span>
            </div>
            <div class='score-card'>
                <span class='score-label'>Processador</span>
                <span class='score-val' style='color:$(Get-Color $cpuScore)'>$cpuScore</span>
            </div>
            <div class='score-card'>
                <span class='score-label'>Memória</span>
                <span class='score-val' style='color:$(Get-Color $memScore)'>$memScore</span>
            </div>
            <div class='score-card'>
                <span class='score-label'>Gráficos</span>
                <span class='score-val' style='color:$(Get-Color $gfxScore)'>$gfxScore</span>
            </div>
            <div class='score-card'>
                <span class='score-label'>Disco</span>
                <span class='score-val' style='color:$diskColor'>$diskScore</span>
            </div>
        </div>

        <!-- Detalhe: Disco -->
        <div class='section'>
            <div class='section-header'>
                <span>💾 Subsistema de Disco (Armazenamento)</span>
                <span>Score: $diskScore</span>
            </div>
            <div class='section-body'>
                <table>
                    <tr><td class='label'>Modelo Detectado</td><td class='value'>$diskModel</td></tr>
                    <tr><td class='label'>Leitura Sequencial</td><td class='value'>$diskSeqRead <span class='unit'>MB/s</span></td></tr>
                    <tr><td class='label'>Leitura Aleatória</td><td class='value'>$diskRandRead <span class='unit'>MB/s</span></td></tr>
                </table>
                $(if($diskScore -ne "N/A" -and $diskScore -lt 6.0) {
                    "<div class='alert alert-crit'><strong>⚠️ ATENÇÃO:</strong> O desempenho do disco está abaixo do ideal para padrões modernos.</div>"
                } elseif($diskScore -ne "N/A") {
                    "<div class='alert alert-ok'><strong>✅ Ótimo:</strong> Desempenho de disco adequado para operações de segurança e multitarefa.</div>"
                } else {
                    "<div class='alert alert-warn'><strong>⚠️:</strong> Não foi possível obter métricas detalhadas do disco.</div>"
                })
            </div>
        </div>

        <!-- Detalhe: Memória -->
        <div class='section'>
            <div class='section-header'>
                <span>🧠 Subsistema de Memória (RAM)</span>
                <span>Score: $memScore</span>
            </div>
            <div class='section-body'>
                <table>
                    <tr><td class='label'>Largura de Banda</td><td class='value'>$memBandwidth <span class='unit'>MB/s</span></td></tr>
                </table>
                <p style='font-size:0.9em; color:#666; margin-top:10px;'>Uma largura de banda alta garante que o processador não fique ocioso aguardando dados.</p>
            </div>
        </div>

        <!-- Detalhe: Processador -->
        <div class='section'>
            <div class='section-header'>
                <span>⚙️ Subsistema de Processador (CPU)</span>
                <span>Score: $cpuScore</span>
            </div>
            <div class='section-body'>
                <table>
                    <tr><td class='label'>Criptografia</td><td class='value'>$cpuEncrypt <span class='unit'>MB/s</span></td></tr>
                    <tr><td class='label'>Compressão</td><td class='value'>$cpuCompress <span class='unit'>MB/s</span></td></tr>
                </table>
            </div>
        </div>

        <!-- Detalhe: Gráficos -->
        <div class='section'>
            <div class='section-header'>
                <span>🎮 Subsistema de Gráficos (Desktop/DWM)</span>
                <span>Score: $gfxScore</span>
            </div>
            <div class='section-body'>
                <table>
                    <tr><td class='label'>FPS na Área de Trabalho</td><td class='value'>$gfxFps <span class='unit'>fps</span></td></tr>
                    <tr><td class='label'>Banda de Memória de Vídeo</td><td class='value'>$gfxBandwidth <span class='unit'>MB/s</span></td></tr>
                </table>
                <p style='font-size:0.85em; color:#999; margin-top:8px;'>*Nota: O teste 3D tradicional foi descontinuado no Windows 11. Estes dados refletem o desempenho real da interface gráfica (DWM).</p>
            </div>
        </div>

        <!-- Battery Report Section -->
        $(if($hasBattery) {
        "<h2>🔋 Relatório de Bateria</h2>
        <div class='section'>
            <div class='section-header'>
                <span>Status da Bateria</span>
            </div>
            <div class='section-body'>
                <p>Um relatório detalhado de bateria foi gerado e salvo em:</p>
                <p style='background: #f8f9fa; padding: 10px; border-radius: 5px; font-family: Consolas; word-break: break-all;'>$batteryReportPath</p>
                <p style='font-size: 0.9em; color: #666; margin-top: 10px;'>Este relatório contém informações sobre:</p>
                <ul style='font-size: 0.9em; color: #666;'>
                    <li>Capacidade de design vs capacidade atual</li>
                    <li>Histórico de uso e duração</li>
                    <li>Ciclos de carga/descarga</li>
                    <li>Estimativa de vida útil</li>
                </ul>
            </div>
        </div>"
        })

 
        <!-- Software Instalado -->

<h2>📦 Inventário de Software ($softwareCount aplicativos)$(if($softwareCountLimited){' (limitado)'})</h2>
<div class='section'>
    <div class='section-header'>
        <span>Programas Instalados</span>
        <span>Total: $softwareCount</span>
    </div>
    <div class='section-body software-table'>
        <table>
            <thead>
                <tr style='background: #f8f9fa; position: sticky; top: 0;'>
                    <th style='text-align: left; padding: 10px;'>Aplicativo</th>
                    <th style='text-align: left; padding: 10px;'>Versão</th>
                    <th style='text-align: left; padding: 10px;'>Fabricante</th>
                    <th style='text-align: left; padding: 10px;'>Data Instalação</th>
                </tr>
            </thead>
            <tbody>
$($softwareList | ForEach-Object {
# Escapar caracteres especiais no HTML usando função nativa
$displayName = ConvertTo-HtmlEncode -Text $_.DisplayName
$displayVersion = ConvertTo-HtmlEncode -Text $_.DisplayVersion
$publisher = ConvertTo-HtmlEncode -Text $_.Publisher
$installDate = ConvertTo-HtmlEncode -Text $_.InstallDate

"                        <tr>
                            <td style='padding: 8px;'>$displayName</td>
                            <td style='padding: 8px;'>$displayVersion</td>
                            <td style='padding: 8px;'>$publisher</td>
                            <td style='padding: 8px;'>$installDate</td>
                        </tr>"
})
            </tbody>
        </table>
    </div>
$(if($softwareCountLimited) {
    "<p style='font-size:0.85em; color:#999; margin-top:10px; text-align:center;'>⚠️ Inventário limitado a 500 itens para melhor performance. Execute consulta direta ao registro para lista completa.</p>"
})
</div>

        <!-- Conclusão -->
        <h2>💡 Veredito Técnico</h2>
        $(if($sysScore -ne "N/A" -and $sysScore -lt 5.0) {
            "<div class='alert alert-crit'><strong>DIAGNÓSTICO CRÍTICO:</strong> O sistema possui limitações severas de hardware. A lentidão relatada é esperada e não culpa do software de segurança.</div>"
        } elseif ($diskScore -ne "N/A" -and $diskScore -lt 7.0) {
            "<div class='alert alert-warn'><strong>DIAGNÓSTICO:</strong> O sistema é funcional, mas o disco pode ser um gargalo em tarefas intensivas de I/O (como varreduras completas). Considere otimizações ou upgrade futuro.</div>"
        } else {
            "<div class='alert alert-ok'><strong>DIAGNÓSTICO POSITIVO:</strong> O hardware está em boas condições. Qualquer lentidão deve ser investigada em configurações de software, drivers ou malware ativo, pois não é falta de capacidade bruta.</div>"
        })

        <div class='footer'>
            <strong>🛠️ Suporte Técnico</strong><br>
            Script Version: 1.0 | Desenvolvido por: Robson Nunes - Cyber Security<br>
            Gerado automaticamente via PowerShell Script | WinSAT + Battery Report + Inventário de Software + Segurança
        </div>
    </div>
</body>
</html>
"@

# 9. Salvar e Abrir
try {
    $html | Out-File -FilePath $outputPath -Encoding UTF8 -ErrorAction Stop
    Write-Host "✅ Relatório gerado com SUCESSO!" -ForegroundColor Green
    Write-Host "📂 Local: $outputPath" -ForegroundColor Cyan
    
    Write-Host "🌐 Abrindo no navegador..." -ForegroundColor Yellow
    Start-Process $outputPath
} catch {
    Write-Host "❌ Erro ao salvar o arquivo: $_" -ForegroundColor Red
}

Write-Host "`n🎉 Processo finalizado!" -ForegroundColor Green
Write-Host "📋 Resumo dos arquivos gerados:" -ForegroundColor Cyan
Write-Host "   1. Relatório Principal: $outputPath" -ForegroundColor White
if ($hasBattery) {
    Write-Host "   2. Battery Report: $batteryReportPath" -ForegroundColor White
}
Write-Host "   3. Inventário: $softwareCount softwares catalogados" -ForegroundColor White
