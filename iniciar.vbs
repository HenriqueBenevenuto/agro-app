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
' 2. ABRE O APP (servidor local oculto + janela estilo aplicativo)
' -------------------------------------------------------------
shell.Run "powershell -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File """ & scriptDir & "\servidor.ps1""", 0, False

WScript.Sleep 1500

appUrl = "http://localhost:5502/aves-vivas.html"

edge1 = shell.ExpandEnvironmentStrings("%ProgramFiles(x86)%") & "\Microsoft\Edge\Application\msedge.exe"
edge2 = shell.ExpandEnvironmentStrings("%ProgramFiles%") & "\Microsoft\Edge\Application\msedge.exe"
chrome1 = shell.ExpandEnvironmentStrings("%ProgramFiles%") & "\Google\Chrome\Application\chrome.exe"
chrome2 = shell.ExpandEnvironmentStrings("%ProgramFiles(x86)%") & "\Google\Chrome\Application\chrome.exe"

If fso.FileExists(edge1) Then
    shell.Run """" & edge1 & """ --app=" & appUrl & " --window-size=1280,800", 1, False
ElseIf fso.FileExists(edge2) Then
    shell.Run """" & edge2 & """ --app=" & appUrl & " --window-size=1280,800", 1, False
ElseIf fso.FileExists(chrome1) Then
    shell.Run """" & chrome1 & """ --app=" & appUrl & " --window-size=1280,800", 1, False
ElseIf fso.FileExists(chrome2) Then
    shell.Run """" & chrome2 & """ --app=" & appUrl & " --window-size=1280,800", 1, False
Else
    shell.Run appUrl, 1, False
End If

' -------------------------------------------------------------
' 3. FORCA O ICONE DA JANELA (o Chrome/Edge no modo --app nem sempre
'    usa o favicon da pagina como icone da barra de tarefas). Gera um
'    script PowerShell temporario soh pra isso -- fica na pasta TEMP,
'    nao precisa manter mais nenhum arquivo junto do app.
' -------------------------------------------------------------
On Error Resume Next
psLinha1 = "$hwnd=[IntPtr]::Zero;$tent=0"
psLinha2 = "Add-Type -Name W -Namespace A -MemberDefinition '[DllImport(""user32.dll"",CharSet=CharSet.Auto)] public static extern IntPtr FindWindow(string c, string n); [DllImport(""user32.dll"")] public static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, IntPtr l); [DllImport(""user32.dll"")] public static extern bool ShowWindow(IntPtr h, int n);'"
psLinha3 = "while($hwnd -eq [IntPtr]::Zero -and $tent -lt 25){Start-Sleep -Milliseconds 300;$hwnd=[A.W]::FindWindow($null,'Agro Benevenuto - Encomendas');$tent++}"
psLinha4 = "if($hwnd -ne [IntPtr]::Zero){[A.W]::ShowWindow($hwnd,3)|Out-Null;Add-Type -AssemblyName System.Drawing;$ico=New-Object System.Drawing.Icon('" & scriptDir & "\agro-benevenuto.ico');[A.W]::SendMessage($hwnd,0x80,[IntPtr]0,$ico.Handle)|Out-Null;[A.W]::SendMessage($hwnd,0x80,[IntPtr]1,$ico.Handle)|Out-Null}"

tempPs = shell.ExpandEnvironmentStrings("%TEMP%") & "\avesagro_icone.ps1"
Set tf = fso.CreateTextFile(tempPs, True)
tf.WriteLine psLinha1
tf.WriteLine psLinha2
tf.WriteLine psLinha3
tf.WriteLine psLinha4
tf.Close

shell.Run "powershell -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File """ & tempPs & """", 0, False
Err.Clear
On Error Goto 0

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