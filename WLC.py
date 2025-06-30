# Variáveis de configuração
$wlcHost = "10.0.90.2"              # IP da WLC
$Username = "decunify"                    # Usuário da WLC
$Password = "drttech2K"            # Senha da WLC

# Comandos para backup
$Commands = @(
    "transfer upload mode tftp",
    "transfer upload datatype config",
    "transfer upload filename backup-config.cfg",
    "transfer upload serverip 10.0.90.228"
)

# Converte senha em formato seguro
$SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential ($Username, $SecurePassword)

# Envia comandos via SSH
foreach ($Command in $Commands) {
    ssh $Credential@$WlcHost $Command
}
