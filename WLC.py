# Variáveis de configuração
$wlcHost = "xxx.xxx.xxx.xxx"              # IP da WLC
$Username = "xxxx"                    # Usuário da WLC
$Password = "xxxx"            # Senha da WLC

# Comandos para backup
$Commands = @(
    "transfer upload mode tftp",
    "transfer upload datatype config",
    "transfer upload filename backup-config.cfg",
    "transfer upload serverip xxx.xxx.xxx.xxx"
)

# Converte senha em formato seguro
$SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential ($Username, $SecurePassword)

# Envia comandos via SSH
foreach ($Command in $Commands) {
    ssh $Credential@$WlcHost $Command
}
