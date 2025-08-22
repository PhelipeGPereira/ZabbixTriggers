from pyzabbix import ZabbixAPI
import pandas as pd
import getpass
 
# Entradas do usuário
zurl = input("Zabbix URL (ex: http://seu_zabbix/zabbix): ").strip()
zuser = input("Zabbix User: ").strip()
zpass = getpass.getpass("Zabbix Password: ").strip()
output_name = input("Nome do arquivo Excel (ex: metrics.xlsx): ").strip()
group_id = input("Digite o ID do grupo de hosts: ").strip()
 
# Conecta ao Zabbix
api = ZabbixAPI(zurl)
try:
    api.login(zuser, zpass)
except Exception as e:
    print("Erro ao conectar/login no Zabbix:", e)
    exit()
 
# Busca os hosts do grupo e templates vinculadas
hosts = api.host.get(
    output=['hostid', 'name'],
    groupids=group_id,
    selectParentTemplates=['templateid', 'name']
)
 
if not hosts:
    print(f"Nenhum host encontrado no grupo ID {group_id}.")
    api.user.logout()
    exit()
 
# Busca macros globais (opcional)
global_macros = api.usermacro.get(output=['macro', 'value'], globalmacro=True)
global_macros_map = {m['macro'].upper(): m.get('value', 'N/A') for m in global_macros}
 
data = []
 
for host in hosts:
    hid = host['hostid']
    hname = host['name']
 
    # Macros do host
    host_macros = api.usermacro.get(hostids=hid, output=['macro', 'value'], inherited=False)
    host_macros_map = {m['macro'].upper(): m.get('value', 'N/A') for m in host_macros}
 
    # Macros das templates vinculadas
    templates = host.get('parentTemplates', [])
    templates_macros_map = {}
 
    for tpl in templates:
        tplid = tpl['templateid']
        tpl_macros = api.usermacro.get(hostids=tplid, output=['macro', 'value'], inherited=False)
        for m in tpl_macros:
            macros_key = m['macro'].upper()
            # Só adiciona se não tiver no mapa das templates (evitar sobrescrever)
            if macros_key not in templates_macros_map:
                templates_macros_map[macros_key] = m.get('value', 'N/A')
 
    # Junta as macros com prioridade:
    # host macros > template macros > global macros
    macros_map = {}
    macros_map.update(global_macros_map)
    macros_map.update(templates_macros_map)
    macros_map.update(host_macros_map)
 
    # Consumo CPU e Memória atuais
    cpu_items = api.item.get(hostids=hid, search={'key_': 'system.cpu.util'}, output=['lastvalue'])
    mem_items = api.item.get(hostids=hid, search={'key_': 'vm.memory.size[pused]'}, output=['lastvalue'])
 
    cpu_val = cpu_items[0]['lastvalue'] if cpu_items else 'N/A'
    mem_val = mem_items[0]['lastvalue'] if mem_items else 'N/A'
 
    # Pega as macros específicas que você pediu
    cpu_warn = macros_map.get('{$CPU.UTIL.WARN}', 'N/A')
    cpu_crit = macros_map.get('{$CPU.UTIL.CRIT}', 'N/A')
    mem_warn = macros_map.get('{$MEMORY.UTIL.WARN}', 'N/A')
    mem_max = macros_map.get('{$MEMORY.UTIL.MAX}', 'N/A')
 
    # Adiciona os dados para o DataFrame
    data.append({
        'Host': hname,
        'CPU - Usage (%)': cpu_val,
        'CPU - Macro WARN (%)': cpu_warn,
        'CPU - Macro CRIT (%)': cpu_crit,
        'Memory - Usage (%)': mem_val,
        'Memory - Macro WARN (%)': mem_warn,
        'Memory - Macro MAX (%)': mem_max
    })
 
if not data:
    print("Nenhum dado coletado dos hosts.")
    api.user.logout()
    exit()
 
# Criação do DataFrame
df = pd.DataFrame(data)
 
# Garante extensão .xlsx
if not output_name.lower().endswith(".xlsx"):
    output_name += ".xlsx"
 
# Salva o Excel corretamente
df.to_excel(output_name, index=False, engine='openpyxl')
 
print(f"✅ Dados salvos com sucesso em '{output_name}'.")
 
# Logout da API
api.user.logout()