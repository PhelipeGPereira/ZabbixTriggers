# -*- coding: utf-8 -*-
# Importa as bibliotecas necessárias
import tkinter as tk
from tkinter import messagebox
from pyzabbix import ZabbixAPI
import pandas as pd
import getpass # Although a standard library, it's good practice to list it
import sys

# Define a classe principal da aplicação GUI
class ZabbixReportApp:
    """
    Uma aplicação GUI para gerar relatórios do Zabbix.
    Permite que o usuário insira as credenciais e o ID do grupo de hosts
    e salva os dados em um arquivo Excel.
    """
    def __init__(self, master):
        self.master = master
        master.title("Gerador de Relatórios Zabbix")
        master.geometry("500x400")
        master.configure(bg="#f0f0f0")

        # Define as variáveis de estado
        self.zabbix_url = tk.StringVar()
        self.zabbix_user = tk.StringVar()
        self.zabbix_pass = tk.StringVar()
        self.output_name = tk.StringVar()
        self.group_id = tk.StringVar()
        
        # Define um rótulo de status para feedback do usuário
        self.status_label = tk.Label(master, text="", bg="#f0f0f0", fg="blue", font=("Arial", 10, "italic"))
        self.status_label.pack(pady=5)

        # Configura a interface gráfica (widgets)
        self.create_widgets(master)

    def create_widgets(self, master):
        """
        Cria todos os widgets da interface (rótulos, campos de entrada, botões).
        """
        frame = tk.Frame(master, padx=10, pady=10, bg="#f0f0f0")
        frame.pack(expand=True)

        # URL do Zabbix
        tk.Label(frame, text="Zabbix URL:", bg="#f0f0f0", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w", pady=5)
        url_entry = tk.Entry(frame, textvariable=self.zabbix_url, width=50)
        url_entry.grid(row=0, column=1)

        # Usuário do Zabbix
        tk.Label(frame, text="Usuário Zabbix:", bg="#f0f0f0", font=("Arial", 10, "bold")).grid(row=1, column=0, sticky="w", pady=5)
        user_entry = tk.Entry(frame, textvariable=self.zabbix_user, width=50)
        user_entry.grid(row=1, column=1)

        # Senha do Zabbix
        tk.Label(frame, text="Senha Zabbix:", bg="#f0f0f0", font=("Arial", 10, "bold")).grid(row=2, column=0, sticky="w", pady=5)
        pass_entry = tk.Entry(frame, textvariable=self.zabbix_pass, show="*", width=50)
        pass_entry.grid(row=2, column=1)

        # Nome do arquivo Excel
        tk.Label(frame, text="Nome do Arquivo Excel:", bg="#f0f0f0", font=("Arial", 10, "bold")).grid(row=3, column=0, sticky="w", pady=5)
        output_entry = tk.Entry(frame, textvariable=self.output_name, width=50)
        output_entry.grid(row=3, column=1)
        output_entry.insert(0, "metrics.xlsx")

        # ID do Grupo de Hosts
        tk.Label(frame, text="ID do Grupo de Hosts:", bg="#f0f0f0", font=("Arial", 10, "bold")).grid(row=4, column=0, sticky="w", pady=5)
        group_entry = tk.Entry(frame, textvariable=self.group_id, width=50)
        group_entry.grid(row=4, column=1)

        # Botão para gerar o relatório
        generate_button = tk.Button(master, text="Gerar Relatório", command=self.generate_report,
                                     bg="#007bff", fg="white", font=("Arial", 12, "bold"),
                                     bd=0, relief="raised", padx=20, pady=10)
        generate_button.pack(pady=20)
        generate_button.bind("<Enter>", lambda e: generate_button.config(bg="#0056b3"))
        generate_button.bind("<Leave>", lambda e: generate_button.config(bg="#007bff"))

    def generate_report(self):
        """
        Executa a lógica principal do script original.
        Conecta à API do Zabbix e gera o arquivo Excel.
        """
        zurl = self.zabbix_url.get().strip()
        zuser = self.zabbix_user.get().strip()
        zpass = self.zabbix_pass.get().strip()
        output_name = self.output_name.get().strip()
        group_id = self.group_id.get().strip()

        # Validação simples de entrada
        if not all([zurl, zuser, zpass, output_name, group_id]):
            messagebox.showerror("Erro de Entrada", "Por favor, preencha todos os campos.")
            return

        self.status_label.config(text="Iniciando a conexão com o Zabbix...", fg="blue")
        self.master.update_idletasks()
        
        # Conecta ao Zabbix
        api = ZabbixAPI(zurl)
        try:
            api.login(zuser, zpass)
        except Exception as e:
            self.status_label.config(text=f"Erro de login: {e}", fg="red")
            messagebox.showerror("Erro de Conexão", f"Erro ao conectar/login no Zabbix: {e}")
            return
        
        self.status_label.config(text="Login bem-sucedido. Buscando hosts...", fg="blue")
        self.master.update_idletasks()

        # Busca os hosts do grupo e templates vinculadas
        try:
            hosts = api.host.get(
                output=['hostid', 'name'],
                groupids=group_id,
                selectParentTemplates=['templateid', 'name']
            )
        except Exception as e:
            self.status_label.config(text=f"Erro ao buscar hosts: {e}", fg="red")
            messagebox.showerror("Erro da API", f"Erro ao buscar hosts no Zabbix: {e}")
            api.user.logout()
            return

        if not hosts:
            self.status_label.config(text=f"Nenhum host encontrado no grupo ID {group_id}.", fg="orange")
            messagebox.showinfo("Nenhum Host", f"Nenhum host encontrado no grupo ID {group_id}.")
            api.user.logout()
            return
            
        self.status_label.config(text="Coletando dados dos hosts...", fg="blue")
        self.master.update_idletasks()
        
        # Busca macros globais (opcional)
        try:
            global_macros = api.usermacro.get(output=['macro', 'value'], globalmacro=True)
            global_macros_map = {m['macro'].upper(): m.get('value', 'N/A') for m in global_macros}
        except Exception as e:
            self.status_label.config(text=f"Erro ao buscar macros globais: {e}", fg="red")
            messagebox.showerror("Erro da API", f"Erro ao buscar macros globais: {e}")
            api.user.logout()
            return

        data = []
        
        # Este loop é a parte principal do script original
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
                    if macros_key not in templates_macros_map:
                        templates_macros_map[macros_key] = m.get('value', 'N/A')

            # Junta as macros com prioridade: host > template > global
            macros_map = {}
            macros_map.update(global_macros_map)
            macros_map.update(templates_macros_map)
            macros_map.update(host_macros_map)

            # Consumo CPU e Memória atuais
            cpu_items = api.item.get(hostids=hid, search={'key_': 'system.cpu.util'}, output=['lastvalue'])
            mem_items = api.item.get(hostids=hid, search={'key_': 'vm.memory.size[pused]'}, output=['lastvalue'])

            cpu_val = cpu_items[0]['lastvalue'] if cpu_items else 'N/A'
            mem_val = mem_items[0]['lastvalue'] if mem_items else 'N/A'

            # Pega as macros específicas
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
            self.status_label.config(text="Nenhum dado coletado dos hosts.", fg="orange")
            messagebox.showinfo("Nenhum Dado", "Nenhum dado coletado dos hosts.")
            api.user.logout()
            return

        # Criação do DataFrame
        df = pd.DataFrame(data)

        # Garante extensão .xlsx
        if not output_name.lower().endswith(".xlsx"):
            output_name += ".xlsx"

        # Salva o Excel
        try:
            df.to_excel(output_name, index=False, engine='openpyxl')
            self.status_label.config(text=f"✅ Dados salvos com sucesso em '{output_name}'.", fg="green")
            messagebox.showinfo("Sucesso", f"Dados salvos com sucesso em '{output_name}'.")
        except Exception as e:
            self.status_label.config(text=f"Erro ao salvar arquivo: {e}", fg="red")
            messagebox.showerror("Erro ao Salvar", f"Erro ao salvar o arquivo Excel: {e}")
        finally:
            # Logout da API, mesmo que ocorra um erro
            api.user.logout()

# Inicia a aplicação
if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = ZabbixReportApp(root)
        root.mainloop()
    except Exception as e:
        print(f"Ocorreu um erro fatal: {e}")
        sys.exit(1)

# Logout da API

api.user.logout()
