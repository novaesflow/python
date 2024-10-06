import subprocess
import re

# Lista completa de IPs permitidos
allowed_ips = [
    "13.85.17.227", "13.84.200.155", "52.183.248.213", "20.118.127.19",
    "20.118.126.206", "20.118.127.88", "20.118.127.89", "20.118.125.255",
    "20.118.127.27", "20.118.120.187", "20.118.121.134", "20.118.123.237",
    "20.118.126.181", "20.118.125.250", "20.118.127.90", "191.232.176.16",
    "191.232.180.88", "191.232.177.130", "191.232.180.187", "191.232.179.236",
    "191.232.177.40", "191.232.180.122", "200.201.179.74", "191.238.209.93",
    "191.238.220.161", "191.238.220.174", "191.238.221.17", "191.233.203.32",
    "191.238.220.179", "191.238.208.208", "191.238.220.127", "191.232.68.163",
    "191.232.69.157", "191.232.71.210", "20.226.184.65", "20.226.184.164",
    "20.226.186.98", "20.226.186.251", "20.226.187.52", "20.226.187.160",
    "191.232.66.85", "20.226.189.21", "20.226.190.40"
]

def discover_local_ips():
    """ Descobre IPs na rede local usando o comando arp -a e retorna uma lista de IPs. """
    local_ips = []
    result = subprocess.run("arp -a", capture_output=True, text=True, shell=True)
    # Usando uma expressão regular para capturar IPs válidos do output do comando arp
    ip_addresses = re.findall(r'\d+\.\d+\.\d+\.\d+', result.stdout)
    for ip in ip_addresses:
        if ip not in local_ips:
            local_ips.append(ip)
    return local_ips

def setup_firewall(port):
    print("Configurando o Firewall do Windows...")
    # Primeiro, removemos a regra, se já existir (para evitar duplicatas)
    subprocess.run("netsh advfirewall firewall delete rule name=\"SQL Server 1433\"", shell=True)

    # Bloquear a porta 1433 para todos
    subprocess.run(f"netsh advfirewall firewall add rule name=\"SQL Server 1433 Block\" dir=in action=block protocol=TCP localport={port}", shell=True)

    # Descobrir IPs locais dinâmicos
    local_ips = discover_local_ips()
    all_ips = set(allowed_ips + local_ips)  # Unir IPs permitidos com IPs locais descobertos

    # Adicionar regra para permitir cada IP
    for ip in all_ips:
        subprocess.run(f"netsh advfirewall firewall add rule name=\"SQL Server 1433 Allow {ip}\" dir=in action=allow protocol=TCP localport={port} remoteip={ip}", shell=True)

    print("Regras de firewall configuradas.")

if __name__ == "__main__":
    setup_firewall(1433)
