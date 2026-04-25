param(
    [int]$Port = 3000,
    [string]$HostName = "localhost",
    [int]$RequestTimeoutMs = 5000,
    [switch]$SkipCertInstall
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
Set-Location $repoRoot

$certDir = Join-Path $repoRoot ".dev-certs"
$pfxPath = Join-Path $certDir "localhost.pfx"
$cerPath = Join-Path $certDir "localhost.cer"
$certPasswordText = "local-dev-only"
$certPassword = ConvertTo-SecureString $certPasswordText -AsPlainText -Force
$certProvider = "Cert:" + [IO.Path]::DirectorySeparatorChar
$currentUserMyStore = $certProvider + "CurrentUser\My"
$currentUserRootStore = $certProvider + "CurrentUser\Root"

function Ensure-LocalCertificate {
    if (-not (Test-Path $certDir)) {
        New-Item -ItemType Directory -Path $certDir | Out-Null
    }

    if ((Test-Path $pfxPath) -and (Test-Path $cerPath)) {
        if (-not $SkipCertInstall) {
            Import-Certificate -FilePath $cerPath -CertStoreLocation $currentUserRootStore | Out-Null
        }
        return
    }

    if ($SkipCertInstall) {
        throw "Missing local certificate files under .dev-certs. Run this script without -SkipCertInstall."
    }

    Write-Host "==> creating trusted local certificate for $HostName"
    $cert = New-SelfSignedCertificate `
        -DnsName $HostName `
        -CertStoreLocation $currentUserMyStore `
        -FriendlyName "Governed Excel Formula Modules Local Add-in" `
        -KeyExportPolicy Exportable `
        -KeySpec KeyExchange `
        -NotAfter (Get-Date).AddYears(1)

    Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $certPassword | Out-Null
    Export-Certificate -Cert $cert -FilePath $cerPath | Out-Null
    Import-Certificate -FilePath $cerPath -CertStoreLocation $currentUserRootStore | Out-Null
}

function Get-ContentType {
    param([string]$Path)

    switch ([IO.Path]::GetExtension($Path).ToLowerInvariant()) {
        ".css" { return "text/css; charset=utf-8" }
        ".html" { return "text/html; charset=utf-8" }
        ".js" { return "text/javascript; charset=utf-8" }
        ".json" { return "application/json; charset=utf-8" }
        ".svg" { return "image/svg+xml" }
        ".tsv" { return "text/tab-separated-values; charset=utf-8" }
        ".txt" { return "text/plain; charset=utf-8" }
        ".xml" { return "application/xml; charset=utf-8" }
        default { return "application/octet-stream" }
    }
}

function Write-Response {
    param(
        [System.Net.Security.SslStream]$Stream,
        [int]$StatusCode,
        [string]$StatusText,
        [string]$ContentType,
        [byte[]]$Body
    )

    $header = "HTTP/1.1 $StatusCode $StatusText`r`n" +
        "Content-Type: $ContentType`r`n" +
        "Content-Length: $($Body.Length)`r`n" +
        "Cache-Control: no-store`r`n" +
        "Access-Control-Allow-Origin: *`r`n" +
        "Connection: close`r`n`r`n"
    $headerBytes = [Text.Encoding]::ASCII.GetBytes($header)
    $Stream.Write($headerBytes, 0, $headerBytes.Length)
    if ($Body.Length -gt 0) {
        $Stream.Write($Body, 0, $Body.Length)
    }
}

function Resolve-RequestPath {
    param([string]$Target)

    $pathOnly = ($Target -split "\?", 2)[0]
    $relative = [Uri]::UnescapeDataString($pathOnly).TrimStart("/")
    if ([string]::IsNullOrWhiteSpace($relative)) {
        $relative = "addin/taskpane.html"
    }
    $relative = $relative -replace "/", [IO.Path]::DirectorySeparatorChar
    return [IO.Path]::GetFullPath((Join-Path $repoRoot $relative))
}

Ensure-LocalCertificate

$portInUse = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue
if ($portInUse) {
    throw "Port $Port is already in use. Stop the existing process or choose another port."
}

$certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($pfxPath, $certPasswordText)
$listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, $Port)
$rootFull = [IO.Path]::GetFullPath($repoRoot)
if (-not $rootFull.EndsWith([IO.Path]::DirectorySeparatorChar)) {
    $rootFull = $rootFull + [IO.Path]::DirectorySeparatorChar
}

Write-Host "==> serving repo root: $repoRoot"
Write-Host "==> add-in manifest: addin\manifest.xml"
Write-Host "==> local HTTPS host: $HostName on port $Port"
Write-Host "==> press Ctrl+C to stop the local add-in server"

$listener.Start()
try {
    while ($true) {
        $client = $listener.AcceptTcpClient()
        $client.ReceiveTimeout = $RequestTimeoutMs
        $client.SendTimeout = $RequestTimeoutMs
        $ssl = $null
        try {
            $ssl = New-Object System.Net.Security.SslStream($client.GetStream(), $false)
            $ssl.ReadTimeout = $RequestTimeoutMs
            $ssl.WriteTimeout = $RequestTimeoutMs
            $authTask = $ssl.AuthenticateAsServerAsync(
                $certificate,
                $false,
                [System.Security.Authentication.SslProtocols]::Tls12,
                $false
            )
            if (-not $authTask.Wait($RequestTimeoutMs)) {
                throw "TLS handshake timed out after $RequestTimeoutMs ms"
            }

            $reader = New-Object System.IO.StreamReader($ssl, [Text.Encoding]::ASCII, $false, 1024, $true)
            $requestLine = $reader.ReadLine()
            if ([string]::IsNullOrWhiteSpace($requestLine)) {
                continue
            }
            do {
                $line = $reader.ReadLine()
            } while ($line -ne $null -and $line -ne "")

            $parts = $requestLine -split " "
            if ($parts.Length -lt 2 -or $parts[0] -notin @("GET", "HEAD")) {
                $body = [Text.Encoding]::UTF8.GetBytes("Method not allowed")
                Write-Response $ssl 405 "Method Not Allowed" "text/plain; charset=utf-8" $body
                continue
            }

            $filePath = Resolve-RequestPath $parts[1]
            if (-not $filePath.StartsWith($rootFull, [StringComparison]::OrdinalIgnoreCase)) {
                $body = [Text.Encoding]::UTF8.GetBytes("Forbidden")
                Write-Response $ssl 403 "Forbidden" "text/plain; charset=utf-8" $body
                continue
            }
            if (-not (Test-Path $filePath -PathType Leaf)) {
                $body = [Text.Encoding]::UTF8.GetBytes("Not found")
                Write-Response $ssl 404 "Not Found" "text/plain; charset=utf-8" $body
                continue
            }

            if ($parts[0] -eq "HEAD") {
                $body = [byte[]]::new(0)
            } else {
                $body = [IO.File]::ReadAllBytes($filePath)
            }
            Write-Response $ssl 200 "OK" (Get-ContentType $filePath) $body
            Write-Host "200 $($parts[0]) $($parts[1])"
        } catch {
            Write-Warning $_.Exception.Message
        } finally {
            if ($ssl) {
                $ssl.Dispose()
            }
            $client.Close()
        }
    }
} finally {
    $listener.Stop()
}
