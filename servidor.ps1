$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$port = 5502

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")
$listener.Start()

Write-Host ""
Write-Host "  Aves Vivas rodando em: http://localhost:$port/aves-vivas.html"
Write-Host "  Deixe esta janela aberta enquanto usa o app."
Write-Host "  Para encerrar, feche esta janela."
Write-Host ""

while ($listener.IsListening) {
    $context = $listener.GetContext()
    $request = $context.Request
    $response = $context.Response
    try {
        $localPath = $request.Url.LocalPath.TrimStart("/")
        if ([string]::IsNullOrEmpty($localPath)) { $localPath = "aves-vivas.html" }

        if ($localPath -eq "__abrir__") {
            # Abre um link usando o programa padrao do sistema (navegador
            # preferido, ou o WhatsApp Desktop se estiver instalado e for
            # o handler registrado) -- em vez de abrir dentro da propria
            # janela do app.
            try {
                $urlPedida = $request.QueryString["url"]
                if (-not $urlPedida) { throw "Nenhum link informado." }
                Start-Process -FilePath $urlPedida
                $corpo = '{"ok":true}'
            } catch {
                $erro = $_.Exception.Message -replace '"', "'"
                $corpo = '{"ok":false,"error":"' + $erro + '"}'
            }
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($corpo)
            $response.ContentType = "application/json; charset=utf-8"
            $response.ContentLength64 = $bytes.Length
            $response.OutputStream.Write($bytes, 0, $bytes.Length)
        } elseif ($localPath -eq "__atualizar__") {
            # Baixa os arquivos mais novos do GitHub e sobrescreve os locais.
            # A pagina so precisa recarregar depois disso -- nao precisa reiniciar nada.
            try {
                $cfgPath = Join-Path $scriptDir "config.txt"
                $cfg = @{}
                if (Test-Path $cfgPath) {
                    Get-Content $cfgPath | ForEach-Object {
                        if ($_ -match "^\s*([^=]+)=(.*)$") { $cfg[$matches[1].Trim()] = $matches[2].Trim() }
                    }
                }
                if (-not $cfg["GITHUB_USER"] -or -not $cfg["GITHUB_REPO"]) {
                    throw "config.txt nao encontrado ou incompleto (precisa de GITHUB_USER e GITHUB_REPO)."
                }
                $baseUrl = "https://raw.githubusercontent.com/$($cfg['GITHUB_USER'])/$($cfg['GITHUB_REPO'])/main/"
                $arquivos = @("aves-vivas.html", "servidor.ps1", "agro-benevenuto.ico")
                foreach ($arq in $arquivos) {
                    $destino = Join-Path $scriptDir $arq
                    Invoke-WebRequest -Uri ($baseUrl + $arq + "?t=" + (Get-Random)) -OutFile $destino -UseBasicParsing
                }
                $verResp = Invoke-WebRequest -Uri ($baseUrl + "version.txt?t=" + (Get-Random)) -UseBasicParsing
                Set-Content -Path (Join-Path $scriptDir "version.txt") -Value $verResp.Content.Trim() -NoNewline -Encoding UTF8
                $corpo = '{"ok":true}'
            } catch {
                $erro = $_.Exception.Message -replace '"', "'"
                $corpo = '{"ok":false,"error":"' + $erro + '"}'
            }
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($corpo)
            $response.ContentType = "application/json; charset=utf-8"
            $response.ContentLength64 = $bytes.Length
            $response.OutputStream.Write($bytes, 0, $bytes.Length)
        } else {

        $filePath = Join-Path $scriptDir $localPath

        if (Test-Path $filePath -PathType Leaf) {
            $ext = [System.IO.Path]::GetExtension($filePath)
            $contentType = switch ($ext) {
                ".html" { "text/html; charset=utf-8" }
                ".js"   { "application/javascript; charset=utf-8" }
                ".css"  { "text/css; charset=utf-8" }
                ".txt"  { "text/plain; charset=utf-8" }
                default { "application/octet-stream" }
            }
            $bytes = [System.IO.File]::ReadAllBytes($filePath)
            $response.ContentType = $contentType
            $response.ContentLength64 = $bytes.Length
            $response.OutputStream.Write($bytes, 0, $bytes.Length)
        } else {
            $response.StatusCode = 404
            $msg = [System.Text.Encoding]::UTF8.GetBytes("Arquivo nao encontrado: $localPath")
            $response.OutputStream.Write($msg, 0, $msg.Length)
        }
        }
    } catch {
        $response.StatusCode = 500
    } finally {
        $response.OutputStream.Close()
    }
}