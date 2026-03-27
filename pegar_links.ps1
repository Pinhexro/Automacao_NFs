$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$nomePastaAlvo = "NFs_Pendentes"

# Função para buscar a pasta em qualquer nível (Recursiva)
function Buscar-Pasta {
    param($pastas, $nome)
    foreach ($f in $pastas) {
        if ($f.Name -eq $nome) { return $f }
        $sub = Buscar-Pasta $f.Folders $nome
        if ($sub) { return $sub }
    }
    return $null
}

# Busca a pasta em toda a sua conta da Diálogo
$pastaEncontrada = Buscar-Pasta $namespace.Folders $nomePastaAlvo

if ($null -eq $pastaEncontrada) {
    Write-Host "ERRO: A pasta '$nomePastaAlvo' nao foi encontrada no Outlook." -ForegroundColor Red
    exit
}

Write-Host "Pasta '$nomePastaAlvo' conectada!" -ForegroundColor Cyan

$links = @()
$emails = $pastaEncontrada.Items | Where-Object {$_.Unread -eq $true}

foreach ($mail in $emails) {
    # Regex para o link da prefeitura de SP
    if ($mail.Body -match "https://nfe\.sf\.prefeitura\.sp\.gov\.br/nfe\.aspx\?[\w=&]+") {
        $link = $matches[0]
        # Regex para o número da nota
        if ($mail.Body -match "NFS-e No\.\s+(\d+)") {
            $num = $matches[1]
            $links += "$num|$link"
            $mail.Unread = $false 
            Write-Host "Nota $num capturada." -ForegroundColor Green
        }
    }
}

# Salva a fila para o Python
if (!(Test-Path "C:\Notas_Fiscais_SP")) { New-Item -ItemType Directory -Path "C:\Notas_Fiscais_SP" }
$links | Out-File -FilePath "C:\Notas_Fiscais_SP\fila_processamento.txt" -Encoding utf8