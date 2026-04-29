' ===================================================
' SCANNER MAQUINAS - WRAPPER PARA GPO (STARTUP)
' Versao: 5.0 - Estrutura unificada SRV-AD02 + Anti-Fantasma
' ===================================================
Option Explicit

Dim FSO, WshShell, objNetwork
Dim SERVER_SCRIPT, LOCAL_FOLDER, LOCAL_SCRIPT
Dim LOG_FILE, ERROR_LOG, COMPUTERNAME, VERSION_FILE
Dim strPSCommand, expectedVersion

' Inicializa objetos
Set FSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")
Set objNetwork = CreateObject("WScript.Network")

' ===================================================
' CONFIGURACOES (Alterar apenas aqui quando necessario)
' ===================================================
COMPUTERNAME = objNetwork.ComputerName
SERVER_SCRIPT = "Colocar caminho do arquivo ps1 recomendado usar servidor com NETLOGON"
LOCAL_FOLDER = "C:\ProgramData\ScannerMaquinas"
LOCAL_SCRIPT = LOCAL_FOLDER & "\ColetaDados_Silencioso.ps1"
VERSION_FILE = LOCAL_FOLDER & "\script_version.txt"
LOG_FILE = LOCAL_FOLDER & "\execution.log"
ERROR_LOG = LOCAL_FOLDER & "\error.log"
expectedVersion = "2.0"

' ===================================================
' FUNCAO PARA ESCREVER LOG
' ===================================================
Sub WriteLog(strFile, strMessage)
    On Error Resume Next
    Dim objLogFile
    Set objLogFile = FSO.OpenTextFile(strFile, 8, True)
    objLogFile.WriteLine Now & " - [" & COMPUTERNAME & "] - " & strMessage
    objLogFile.Close
    On Error GoTo 0
End Sub

' ===================================================
' FUNCAO PARA EXECUTAR COMANDO COM SYSTEM (Metodo original que funcionava)
' ===================================================
Function RunAsSystem(strCmd)
    On Error Resume Next
    Dim objProcess
    Set objProcess = WshShell.Exec(strCmd)
    Do While objProcess.Status = 0
        WScript.Sleep 100
    Loop
    RunAsSystem = objProcess.ExitCode
    On Error GoTo 0
End Function

' ===================================================
' INICIO DA EXECUCAO
' ===================================================

WriteLog LOG_FILE, "========== WRAPPER 5.0 INICIADO =========="
WriteLog LOG_FILE, "Destino CSVs: colocar um servidor de destino dos arquivos"

' Criar pasta local com permissoes para SYSTEM
If Not FSO.FolderExists(LOCAL_FOLDER) Then
    On Error Resume Next
    FSO.CreateFolder LOCAL_FOLDER
    If Err.Number <> 0 Then
        LOCAL_FOLDER = "C:\Windows\Temp\ScannerMaquinas"
        FSO.CreateFolder LOCAL_FOLDER
        LOCAL_SCRIPT = LOCAL_FOLDER & "\ColetaDados_Silencioso.ps1"
        VERSION_FILE = LOCAL_FOLDER & "\script_version.txt"
        LOG_FILE = LOCAL_FOLDER & "\execution.log"
        ERROR_LOG = LOCAL_FOLDER & "\error.log"
    End If
    On Error GoTo 0
End If

' ===================================================
' VERIFICACAO DE VERSAO (Anti-Fantasma)
' ===================================================
Dim currentVersion, precisaAtualizar
currentVersion = ""
precisaAtualizar = True

If FSO.FileExists(VERSION_FILE) Then
    On Error Resume Next
    Dim verFile
    Set verFile = FSO.OpenTextFile(VERSION_FILE, 1)
    currentVersion = verFile.ReadLine()
    verFile.Close
    On Error GoTo 0
End If

WriteLog LOG_FILE, "Versao em cache: [" & currentVersion & "] | Esperada: [" & expectedVersion & "]"

If currentVersion = expectedVersion And FSO.FileExists(LOCAL_SCRIPT) Then
    WriteLog LOG_FILE, "Cache OK - Usando versao local"
    precisaAtualizar = False
End If

' ===================================================
' ATUALIZACAO DO SCRIPT (Se necessario)
' ===================================================
If precisaAtualizar Then
    WriteLog LOG_FILE, "Atualizando script..."
    
    ' Remove versao antiga
    If FSO.FileExists(LOCAL_SCRIPT) Then
        On Error Resume Next
        FSO.DeleteFile LOCAL_SCRIPT, True
        WriteLog LOG_FILE, "Versao antiga removida"
        On Error GoTo 0
    End If
    
    ' Remove logs antigos para debug limpo
    Dim debugLog
    debugLog = LOCAL_FOLDER & "\Logs\debug_" & COMPUTERNAME & ".log"
    If FSO.FileExists(debugLog) Then
        FSO.DeleteFile debugLog, True
        WriteLog LOG_FILE, "Log de debug antigo removido"
    End If
    
    ' Remove CSVs antigos do fallback local
    Dim oldCsvPath, oldCsv
    oldCsvPath = LOCAL_FOLDER & "\CSV"
    If FSO.FolderExists(oldCsvPath) Then
        On Error Resume Next
        oldCsv = oldCsvPath & "\Dados_" & COMPUTERNAME & ".csv"
        If FSO.FileExists(oldCsv) Then FSO.DeleteFile oldCsv, True
        oldCsv = oldCsvPath & "\Programas_" & COMPUTERNAME & ".csv"
        If FSO.FileExists(oldCsv) Then FSO.DeleteFile oldCsv, True
        On Error GoTo 0
    End If
    
    ' Verificar acesso ao servidor de scripts
    If Not FSO.FileExists(SERVER_SCRIPT) Then
        WriteLog ERROR_LOG, "ERRO: Nao foi possivel acessar: " & SERVER_SCRIPT
        WriteLog ERROR_LOG, "Verifique permissao de Domain Computers no compartilhamento"
        WScript.Quit 1
    End If
    
    ' Copiar script do servidor (com retry)
    Dim copySuccess, retryCount
    copySuccess = False
    retryCount = 0
    
    Do While copySuccess = False And retryCount < 3
        On Error Resume Next
        FSO.CopyFile SERVER_SCRIPT, LOCAL_SCRIPT, True
        If Err.Number = 0 Then
            copySuccess = True
        Else
            retryCount = retryCount + 1
            WriteLog LOG_FILE, "Tentativa " & retryCount & " falhou: " & Err.Description
            WScript.Sleep 2000
        End If
        On Error GoTo 0
    Loop
    
    If copySuccess = False Then
        WriteLog ERROR_LOG, "FALHA CRITICA: Nao foi possivel copiar o script apos 3 tentativas"
        WriteLog ERROR_LOG, "Origem: " & SERVER_SCRIPT
        WriteLog ERROR_LOG, "Destino: " & LOCAL_SCRIPT
        WScript.Quit 1
    End If
    
    ' Salva a versao atual
    On Error Resume Next
    Dim outFile
    Set outFile = FSO.CreateTextFile(VERSION_FILE, True)
    outFile.WriteLine expectedVersion
    outFile.Close
    On Error GoTo 0
    
    WriteLog LOG_FILE, "Script atualizado para versao " & expectedVersion
End If

' ===================================================
' VERIFICACAO FINAL DO SCRIPT LOCAL
' ===================================================
If Not FSO.FileExists(LOCAL_SCRIPT) Then
    WriteLog ERROR_LOG, "ERRO: Script local nao encontrado: " & LOCAL_SCRIPT
    WScript.Quit 1
End If

' ===================================================
' EXECUCAO DO POWERSHELL
' ===================================================
WriteLog LOG_FILE, "Executando PowerShell..."

' Metodo original que funcionava (RunAsSystem com WshShell.Exec)
strPSCommand = "powershell.exe -NoProfile -ExecutionPolicy Bypass -NonInteractive -File """ & LOCAL_SCRIPT & """"

Dim exitCode
exitCode = RunAsSystem(strPSCommand)

WriteLog LOG_FILE, "PowerShell finalizado - ExitCode: " & exitCode

' ===================================================
' FINALIZACAO
' ===================================================
WriteLog LOG_FILE, "========== WRAPPER 5.0 FINALIZADO =========="
WScript.Quit 0