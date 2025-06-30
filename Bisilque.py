import time
from pyVim.connect import SmartConnect, Disconnect
from pyVmomi import vim
import ssl

# Função para conectar ao vCenter
def connect_to_vcenter(vcenter_ip, username, password):
    context = ssl._create_unverified_context()
    si = SmartConnect(host=vcenter_ip, user=username, pwd=password, sslContext=context)
    return si

# Função para criar um snapshot no vSphere
def create_snapshot(vm, snapshot_name):
    print(f"Criando snapshot para a VM '{vm.name}' com o nome '{snapshot_name}'...")
    snapshot_task = vm.CreateSnapshot_Task(name=snapshot_name, memory=False, quiesce=False)
    snapshot_task_result = snapshot_task.info.state
    while snapshot_task.info.state not in [vim.TaskInfo.State.success, vim.TaskInfo.State.error]:
        time.sleep(2)
    if snapshot_task.info.state == vim.TaskInfo.State.success:
        print(f"Snapshot {snapshot_name} criado com sucesso!")
    else:
        print(f"Erro ao criar o snapshot {snapshot_name}")

# Função para buscar a VM no vCenter
def find_vm_in_vcenter(content, vm_name):
    for datacenter in content.rootFolder.childEntity:
        if isinstance(datacenter, vim.Datacenter):
            vm_folder = datacenter.vmFolder
            vm = find_vm_in_folder(vm_folder, vm_name)
            if vm:
                return vm
    return None

# Função recursiva para buscar a VM em subpastas
def find_vm_in_folder(folder, vm_name):
    for vm_entity in folder.childEntity:
        if isinstance(vm_entity, vim.VirtualMachine) and vm_entity.name == vm_name:
            return vm_entity
        elif isinstance(vm_entity, vim.Folder):
            # Chama a função recursivamente para subpastas
            vm = find_vm_in_folder(vm_entity, vm_name)
            if vm:
                return vm
    return None

# Parâmetros de conexão
vcenter_ips = [
    "192.168.20.26", 
    "192.168.20.27", 
    "192.168.20.29"
]  # Lista de IPs de vCenters

vcenter_username = "root"
vcenter_password = "1Bisilque23"

# Lista de VMs a processar
vm_names = [
    "BIS-PT-APL02",  # Adicione os nomes das VMs aqui
    "BIS-PT-DC00",
    "BIS-PT-ERP01",
    "BIS-PT-FG02",
    "BIS-PT-MNT01",
    "VM-PT-ESM-001",
    "VM-PT-ESM-002"    
    
]

# Itera sobre cada vCenter
for vcenter_ip in vcenter_ips:
    print(f"Conectando-se ao vCenter: {vcenter_ip}")
    
    # Conectar ao vCenter
    si = connect_to_vcenter(vcenter_ip, vcenter_username, vcenter_password)

    # Localizar as VMs no vCenter
    content = si.RetrieveContent()
    
    # Itera sobre os nomes das VMs
    for vm_name in vm_names:
        print(f"Procurando pela VM: {vm_name}")
        
        # Buscar VM
        vm = find_vm_in_vcenter(content, vm_name)

        # Criar Snapshot se a VM for encontrada
        if vm:
            snapshot_name = f"Snapshot-Updates Semanais-{time.strftime('%Y/%m/%d')}"
            create_snapshot(vm, snapshot_name)
        else:
            print(f"VM '{vm_name}' não encontrada no vCenter '{vcenter_ip}'!")

    # Desconectar do vCenter
    Disconnect(si)
