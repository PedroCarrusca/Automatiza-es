import requests
import json
import browser_cookie3  # Captura cookies do navegador

# Coleta cookies do Chrome (ou use .firefox(), .edge(), etc.)
cookies = browser_cookie3.chrome(domain_name='servicos.grupodecsis.eu')

# URL da API do iTop
itop_url = "https://servicos.grupodecsis.eu/webservices/rest.php?version=1.3"

# Dados do ticket
ticket_data = {
    "operation": "core/create",
    "class": "UserRequest",
    "fields": {
        "title": "Disco acima da percentagem! no dispositivo FLATLANTIC-srvntilocls02",
        "description": "Equipamento em questão possui o disco c: acima da percentagem",
        "org_id": "Flatlantic",         # Substituir por ID se necessário
        "caller_id": "Nuno Cunha",      # Substituir por ID se necessário
        "impact": "2",
        "urgency": "2",
        "origin": "monitoramento",
        "contract_id": "308",           # <-- Substituir pelo ID real
        "service_id": "308"             # <-- Substituir pelo ID real
    }
}

headers = {
    "Content-Type": "application/json"
}

# Envio da requisição com os cookies capturados do navegador
response = requests.post(itop_url, headers=headers, cookies=cookies, json=ticket_data)

# Exibir resposta
if response.status_code == 200:
    print("Ticket criado com sucesso:")
    print(json.dumps(response.json(), indent=2))
else:
    print("Erro ao criar ticket:")
    print(response.status_code)
    print(response.text)
