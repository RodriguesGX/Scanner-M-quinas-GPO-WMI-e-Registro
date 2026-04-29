<#
.SYNOPSIS
    Script para coleta - VERSÃO FINAL HÍBRIDA
.DESCRIPTION
    Modo mínimo para PS 2.0 (Windows 7) + Modo completo para PS 5.0+.
.NOTES
    Versao: 13.0 - FINAL HÍBRIDA
    Destino: colocar destino dos arquivos csv final
#>

$SCRIPT_VERSION = "2.0" #sempre colocar a versão no arquivo vbs também

# Configuração mínima
$ErrorActionPreference = "SilentlyContinue"

$dadosPath = "colocar destino dos arquivos csv final"
$computerName = $env:COMPUTERNAME
$localLogPath = "C:\ProgramData\ScannerMaquinas\Logs"

# Detectar versão do PowerShell
$psVersion = $PSVersionTable.PSVersion.Major
$isPS5 = ($psVersion -ge 5)

# Criar pasta de logs
if (-not (Test-Path $localLogPath)) {
    New-Item -ItemType Directory -Path $localLogPath -Force | Out-Null
}

# Função de log
$logFile = "$localLogPath\debug_${computerName}.log"
function Write-Log($msg) {
    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $msg" | Out-File -FilePath $logFile -Append
}

Write-Log "========== V:$SCRIPT_VERSION (PS$psVersion) INICIADA =========="
Write-Log "Computador: $computerName"

# Verificar destino
if (-not (Test-Path $dadosPath)) {
    $dadosPath = "C:\ProgramData\ScannerMaquinas\CSV"
    if (-not (Test-Path $dadosPath)) {
        New-Item -ItemType Directory -Path $dadosPath -Force | Out-Null
    }
    Write-Log "Fallback local: $dadosPath"
}

# ============================================================
# COLETA DE DADOS DA MÁQUINA
# ============================================================

# Campos base (PS 2.0)
$dados = New-Object PSObject
$camposBase = @("Computador","Usuario","DataColeta","Sistema","VersaoSO","Arquitetura",
                "ChaveWindows_Registro","TipoLicenca","Processador","ProcessadorQtd",
                "MemoriaTotalGB","MemoriaDisponivelGB","MemoriaPctUso",
                "HD_TotalGB","HD_LivreGB","HD_PctUso",
                "ServiceTag","BiosFabricante","BiosVersao","ModeloMaquina",
                "TemGPU","GPU_Modelo","GPU_RAM_MB",
                "IP_Principal","MAC_Principal","Dominio",
                "Office2010_Chave","Office2010_Status",
                "Office2013_Chave","Office2013_Status",
                "Office2016_Chave","Office2016_Status",
                "Office365_Chave","Office365_Status",
                "PSVersion","DiasDesdeUltimaColeta")

# Campos extras (PS 5.0+)
$camposExtras = @("QtdMonitores","QtdDiscos","Discos_Detalhes","RAM_Slots","RAM_Modulos")

if ($isPS5) {
    $todosCampos = $camposBase + $camposExtras
} else {
    $todosCampos = $camposBase
}

foreach ($c in $todosCampos) { $dados | Add-Member -MemberType NoteProperty -Name $c -Value "" }

$dados.Computador = $computerName
$dados.Usuario = $env:USERNAME
$dados.DataColeta = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
$dados.PSVersion = "$psVersion.0"
$dados.DiasDesdeUltimaColeta = "999"

# ============================================================
# COLETAS VIA REGISTRO (Funciona em qualquer versão)
# ============================================================
Write-Log "Coletando via Registro..."

# Sistema
$reg = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
$dados.Sistema = $reg.ProductName
$dados.VersaoSO = $reg.CurrentVersion + "." + $reg.CurrentBuildNumber
$dados.Arquitetura = if ([Environment]::Is64BitOperatingSystem) { "64-bit" } else { "32-bit" }

# Processador
$cpu = Get-ItemProperty "HKLM:\HARDWARE\DESCRIPTION\System\CentralProcessor\0"
$dados.Processador = $cpu.ProcessorNameString
$dados.ProcessadorQtd = [Environment]::ProcessorCount

# BIOS
$bios = Get-ItemProperty "HKLM:\HARDWARE\DESCRIPTION\System\BIOS"
$dados.ServiceTag = $bios.SystemSerialNumber
$dados.BiosFabricante = $bios.BiosVendor
$dados.BiosVersao = $bios.BIOSVersion
$dados.ModeloMaquina = $bios.SystemProductName

# Chave Windows
$map = "BCDFGHJKMPQRTVWXY2346789"
$key = ""
$digitalProductId = $reg.DigitalProductId
if ($digitalProductId) {
    $isWin8 = [Math]::Floor(($digitalProductId[66] -band 0x06) / 2)
    $digitalProductId[66] = ($digitalProductId[66] -band 0xF7) -bor ($isWin8 * 4)
    for ($i = 24; $i -ge 0; $i--) {
        $current = 0
        for ($j = 14; $j -ge 0; $j--) {
            $current = $current * 256
            $current = $digitalProductId[$j + 52] + $current
            $digitalProductId[$j + 52] = [Math]::Floor($current / 24)
            $current = $current % 24
        }
        $key = $map[$current] + $key
    }
    $dados.ChaveWindows_Registro = $key -replace '(.{5})(.{5})(.{5})(.{5})(.{5})', '$1-$2-$3-$4-$5'
    $dados.TipoLicenca = "Registro (Instalacao)"
}

# ============================================================
# COLETAS VIA WMI (com fallback)
# ============================================================
Write-Log "Coletando via WMI..."

# Memória
try {
    $os = Get-WmiObject Win32_OperatingSystem -ErrorAction Stop
    if ($os) {
        $totalMem = [math]::Round($os.TotalVisibleMemorySize / 1MB, 2)
        $livreMem = [math]::Round($os.FreePhysicalMemory / 1MB, 2)
        $dados.MemoriaTotalGB = "$totalMem GB"
        $dados.MemoriaDisponivelGB = "$livreMem GB"
        if ($totalMem -gt 0) {
            $dados.MemoriaPctUso = [math]::Round((($totalMem - $livreMem) / $totalMem) * 100, 2).ToString() + "%"
        }
    }
} catch { Write-Log "Memória indisponível" }

# HD
try {
    $disk = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction Stop
    if ($disk -and $disk.Size) {
        $totalHD = [math]::Round($disk.Size / 1GB, 2)
        $livreHD = [math]::Round($disk.FreeSpace / 1GB, 2)
        $dados.HD_TotalGB = "$totalHD GB"
        $dados.HD_LivreGB = "$livreHD GB"
        if ($totalHD -gt 0) {
            $dados.HD_PctUso = [math]::Round((($totalHD - $livreHD) / $totalHD) * 100, 2).ToString() + "%"
        }
    }
} catch { Write-Log "HD indisponível" }

# Rede
try {
    $net = Get-WmiObject Win32_NetworkAdapterConfiguration -ErrorAction Stop | Where-Object { $_.IPEnabled -eq $true }
    if ($net) {
        if ($net.IPAddress -is [array]) { $dados.IP_Principal = $net.IPAddress[0] }
        else { $dados.IP_Principal = $net.IPAddress }
        if ($net.MACAddress -is [array]) { $dados.MAC_Principal = $net.MACAddress[0] }
        else { $dados.MAC_Principal = $net.MACAddress }
    }
} catch { Write-Log "Rede indisponível" }

# Domínio
try {
    $cs = Get-WmiObject Win32_ComputerSystem -ErrorAction Stop
    if ($cs) { $dados.Dominio = $cs.Domain }
} catch { }

# GPU
try {
    $gpus = Get-WmiObject Win32_VideoController -ErrorAction Stop
    if ($gpus) {
        $dados.TemGPU = "Sim"
        $nomes = ""
        $ramTotal = 0
        foreach ($g in $gpus) {
            if ($g.Name) {
                if ($nomes -ne "") { $nomes = $nomes + " | " }
                $nomes = $nomes + $g.Name
            }
            if ($g.AdapterRAM) { $ramTotal = $ramTotal + [math]::Round($g.AdapterRAM / 1MB, 0) }
        }
        $dados.GPU_Modelo = $nomes
        if ($ramTotal -gt 0) { $dados.GPU_RAM_MB = $ramTotal }
    }
} catch { }

# ============================================================
# COLETAS EXTRAS (PS 5.0+)
# ============================================================
if ($isPS5) {
    Write-Log "Coletando extras (PS 5.0+)..."
    
    # Monitores
    try {
        $mons = Get-CimInstance -ClassName Win32_DesktopMonitor -ErrorAction Stop | Where-Object { $_.Name -notlike "*Default*" }
        if ($mons) {
            $count = 0
            foreach ($m in $mons) { $count = $count + 1 }
            $dados.QtdMonitores = $count
        }
    } catch { }
    
    # Discos
    try {
        $allDisks = Get-CimInstance -ClassName Win32_LogicalDisk -ErrorAction Stop | Where-Object { $_.DriveType -eq 3 }
        if ($allDisks) {
            $count = 0
            $detalhes = ""
            foreach ($d in $allDisks) {
                $count = $count + 1
                if ($d.Size) {
                    $t = [math]::Round($d.Size / 1GB, 2)
                    $l = [math]::Round($d.FreeSpace / 1GB, 2)
                    $info = "$($d.DeviceID): $t GB (Livre: $l GB)"
                } else {
                    $info = "$($d.DeviceID): (indisponível)"
                }
                if ($detalhes -ne "") { $detalhes = $detalhes + " | " }
                $detalhes = $detalhes + $info
            }
            $dados.QtdDiscos = $count
            $dados.Discos_Detalhes = $detalhes
        }
    } catch { }
}

# ============================================================
# OFFICE
# ============================================================
Write-Log "Verificando Office..."
$officePaths = @(
    @{Path="C:\Program Files (x86)\Microsoft Office\Office14\OSPP.VBS"; Versao="Office2010"},
    @{Path="C:\Program Files\Microsoft Office\Office14\OSPP.VBS"; Versao="Office2010"},
    @{Path="C:\Program Files (x86)\Microsoft Office\Office15\OSPP.VBS"; Versao="Office2013"},
    @{Path="C:\Program Files\Microsoft Office\Office15\OSPP.VBS"; Versao="Office2013"},
    @{Path="C:\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS"; Versao="Office2016"},
    @{Path="C:\Program Files\Microsoft Office\Office16\OSPP.VBS"; Versao="Office2016"},
    @{Path="C:\Program Files\Microsoft Office\root\Office16\OSPP.VBS"; Versao="Office365"},
    @{Path="C:\Program Files (x86)\Microsoft Office\root\Office16\OSPP.VBS"; Versao="Office365"}
)
foreach ($office in $officePaths) {
    if (Test-Path $office.Path) {
        Write-Log "Office: $($office.Versao)"
        $output = & cscript $office.Path /dstatus //nologo 2>$null
        if ($output) {
            $chave = ""; $status = ""
            foreach ($line in ($output -split "`r`n")) {
                if ($line -match 'Last 5 characters.*:\s*(.*)') { $chave = $matches[1] }
                if ($line -match 'LICENSE STATUS.*:\s*(.*)') { $status = $matches[1] }
            }
            switch ($office.Versao) {
                "Office2010" { $dados.Office2010_Chave = $chave; $dados.Office2010_Status = $status }
                "Office2013" { $dados.Office2013_Chave = $chave; $dados.Office2013_Status = $status }
                "Office2016" { $dados.Office2016_Chave = $chave; $dados.Office2016_Status = $status }
                "Office365"  { $dados.Office365_Chave = $chave; $dados.Office365_Status = $status }
            }
        }
    }
}

# ============================================================
# SALVAR DADOS
# ============================================================
Write-Log "Salvando dados..."
$dadosFile = "$dadosPath\Dados_${computerName}.csv"
try {
    $dados | Select-Object $todosCampos | Export-Csv -Path $dadosFile -NoTypeInformation -Encoding UTF8 -Force
    Write-Log "CSV salvo: $dadosFile"
} catch {
    Write-Log "ERRO ao salvar: $_"
}

# ============================================================
# PROGRAMAS INSTALADOS
# ============================================================
Write-Log "Coletando programas..."

$programas = @()
$paths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

foreach ($path in $paths) {
    $progs = Get-ItemProperty $path -ErrorAction SilentlyContinue | 
             Where-Object { $_.DisplayName -and $_.DisplayName -notlike "*Update*" -and $_.DisplayName -notlike "*Security Update*" }
    foreach ($p in $progs) {
        $obj = New-Object PSObject
        
        # Informações detalhadas (disponíveis em todas as versões)
        $obj | Add-Member -MemberType NoteProperty -Name "Nome" -Value $p.DisplayName
        $obj | Add-Member -MemberType NoteProperty -Name "Versao" -Value $p.DisplayVersion
        $obj | Add-Member -MemberType NoteProperty -Name "Fabricante" -Value $p.Publisher
        
        # Data de instalação
        $dataInst = ""
        if ($p.InstallDate) { 
            try { $dataInst = [datetime]::ParseExact($p.InstallDate, "yyyyMMdd", $null).ToString("dd/MM/yyyy") } catch { }
        }
        $obj | Add-Member -MemberType NoteProperty -Name "DataInstalacao" -Value $dataInst
        
        # Tamanho
        $tamMB = ""
        if ($p.EstimatedSize) { $tamMB = [math]::Round($p.EstimatedSize / 1024, 2) }
        $obj | Add-Member -MemberType NoteProperty -Name "TamanhoMB" -Value $tamMB
        
        $programas += $obj
    }
}

Write-Log "Programas encontrados: $($programas.Count)"

# Remover duplicatas (por Nome + Versão)
$unicos = @()
$vistos = @{}
foreach ($p in $programas) {
    $k = $p.Nome + "|||" + $p.Versao
    if (-not $vistos.ContainsKey($k)) {
        $vistos[$k] = $true
        $unicos += $p
    }
}
Write-Log "Programas únicos: $($unicos.Count)"

# Ordenar por nome
$unicos = $unicos | Sort-Object Nome

# Salvar programas
$programasFile = "$dadosPath\Programas_${computerName}.csv"
try {
    $unicos | Export-Csv -Path $programasFile -NoTypeInformation -Encoding UTF8 -Force
    Write-Log "CSV de programas salvo"
} catch {
    Write-Log "ERRO ao salvar programas: $_"
}

Write-Log "========== FINALIZADO =========="
# Fim do script