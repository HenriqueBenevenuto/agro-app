' ===================================================================
'  AVES AGRO - LANCADOR + AUTO-ATUALIZACAO
' ===================================================================
'  CONFIGURACAO: de preferencia edite o arquivo config.txt (na mesma
'  pasta) em vez destas linhas -- ele tambem e usado pela pagina do
'  app pra checar atualizacoes sozinha. Se config.txt nao existir,
'  usa os valores abaixo.
' ===================================================================
GITHUB_USER_PADRAO = "SEU-USUARIO-AQUI"
GITHUB_REPO_PADRAO = "aves-agro-app"
' ===================================================================

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)

' -------------------------------------------------------------
' 0. SE JA ESTIVER RODANDO, NAO FAZ NADA (evita abrir duas janelas
'    caso o atalho seja clicado duas vezes, ou disparado por engano
'    mais de uma vez)
' -------------------------------------------------------------
Function JaRodando()
    JaRodando = False
    On Error Resume Next
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set colProcessos = wmi.ExecQuery("Select CommandLine from Win32_Process WHERE Name='powershell.exe'")
    For Each objProcesso In colProcessos
        If Not IsNull(objProcesso.CommandLine) Then
            If InStr(1, objProcesso.CommandLine, "servidor.ps1", 1) > 0 Then
                JaRodando = True
                Exit For
            End If
        End If
    Next
    On Error Goto 0
End Function

If JaRodando() Then
    WScript.Quit
End If

GITHUB_USER = GITHUB_USER_PADRAO
GITHUB_REPO = GITHUB_REPO_PADRAO
configPath = scriptDir & "\config.txt"
If fso.FileExists(configPath) Then
    Set cf = fso.OpenTextFile(configPath, 1)
    Do Until cf.AtEndOfStream
        linha = cf.ReadLine
        pos = InStr(linha, "=")
        If pos > 0 Then
            chave = Trim(Left(linha, pos - 1))
            valor = Trim(Mid(linha, pos + 1))
            If chave = "GITHUB_USER" And valor <> "" Then GITHUB_USER = valor
            If chave = "GITHUB_REPO" And valor <> "" Then GITHUB_REPO = valor
        End If
    Loop
    cf.Close
End If

BASE_URL = "https://raw.githubusercontent.com/" & GITHUB_USER & "/" & GITHUB_REPO & "/main/"

' -------------------------------------------------------------
' 1. VERIFICA ATUALIZACAO (silenciosamente; se nao tiver internet
'    ou o link nao estiver configurado ainda, so ignora e abre)
' -------------------------------------------------------------
On Error Resume Next

localVersionPath = scriptDir & "\version.txt"
localVersion = 0
If fso.FileExists(localVersionPath) Then
    Set f = fso.OpenTextFile(localVersionPath, 1)
    localVersion = ParseVersao(Trim(f.ReadAll))
    f.Close
End If

remoteVersionText = HttpGetText(BASE_URL & "version.txt")
If Err.Number = 0 And remoteVersionText <> "" Then
    remoteVersion = ParseVersao(Trim(remoteVersionText))
    If remoteVersion > localVersion Then
        changelog = HttpGetText(BASE_URL & "changelog.txt")
        msg = "Uma nova versao do Aves Agro esta disponivel." & vbCrLf & vbCrLf
        If changelog <> "" Then msg = msg & "Novidades:" & vbCrLf & changelog & vbCrLf & vbCrLf
        msg = msg & "Deseja atualizar agora?"
        resposta = MsgBox(msg, vbYesNo + vbQuestion, "Aves Agro - Atualizacao disponivel")
        If resposta = vbYes Then
            ok1 = AtualizarArquivo("aves-vivas.html")
            ok2 = AtualizarArquivo("servidor.ps1")
            ok3 = AtualizarArquivo("agro-benevenuto.ico")
            If ok1 And ok2 And ok3 Then
                Set f = fso.CreateTextFile(localVersionPath, True)
                f.Write Trim(remoteVersionText)
                f.Close
                MsgBox "Atualizado! Abrindo o Aves Agro...", vbInformation, "Aves Agro"
            Else
                MsgBox "Não foi possível baixar algum dos arquivos (falha de conexão). O Aves Agro vai abrir na versão atual, e vai tentar atualizar de novo na próxima vez que for aberto.", vbExclamation, "Aves Agro - Falha na atualização"
            End If
        End If
    End If
End If

Err.Clear
On Error Goto 0

' -------------------------------------------------------------
' 2. ABRE O APP (servidor local oculto -- ele mesmo abre a janela
'    do navegador, ja centralizada na tela)
' -------------------------------------------------------------
shell.Run "powershell -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File """ & scriptDir & "\servidor.ps1""", 0, False

' -------------------------------------------------------------
' FUNCOES AUXILIARES
' -------------------------------------------------------------
' Converte "2", "2.2", "2.10" etc num numero comparavel, sem depender da
' configuracao regional do Windows (evita bug de virgula x ponto decimal).
Function ParseVersao(txt)
    Dim partes, maior, menor
    partes = Split(Trim(txt), ".")
    maior = CIntSeguro(partes(0))
    menor = 0
    If UBound(partes) >= 1 Then menor = CIntSeguro(partes(1))
    ParseVersao = maior * 1000 + menor
End Function

Function CIntSeguro(s)
    On Error Resume Next
    CIntSeguro = CInt(s)
    If Err.Number <> 0 Then CIntSeguro = 0
    Err.Clear
    On Error Goto 0
End Function

Function HttpGetText(url)
    On Error Resume Next
    Dim http
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.SetTimeouts 4000, 4000, 4000, 4000
    http.Open "GET", url & "?t=" & Timer, False
    http.Send
    If http.Status = 200 Then
        HttpGetText = http.ResponseText
    Else
        HttpGetText = ""
    End If
    On Error Goto 0
End Function

Function AtualizarArquivo(nomeArquivo)
    On Error Resume Next
    AtualizarArquivo = False
    Dim http, stream
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.SetTimeouts 8000, 8000, 15000, 15000
    http.Open "GET", BASE_URL & nomeArquivo & "?t=" & Timer, False
    http.Send
    If Err.Number = 0 And http.Status = 200 Then
        Dim corpo
        corpo = http.ResponseBody
        ' checa se realmente veio conteudo (evita gravar arquivo vazio por cima do bom)
        If Not IsNull(corpo) Then
            Set stream = CreateObject("ADODB.Stream")
            stream.Type = 1 ' binario, evita qualquer problema de acentos/emojis
            stream.Open
            stream.Write corpo
            If stream.Size > 0 Then
                stream.SaveToFile scriptDir & "\" & nomeArquivo, 2 ' sobrescreve
                If Err.Number = 0 Then AtualizarArquivo = True
            End If
            stream.Close
        End If
    End If
    Err.Clear
    On Error Goto 0
End Function