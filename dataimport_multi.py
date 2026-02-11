import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

# ==================== VARIÁVEIS GLOBAIS ====================

caminho_arquivo = None
pasta_saida = None
processado = False
versao = "1.6.1 - Multi"

# ==================== FUNÇÕES DE UTILIDADE ====================

def resetar_pos_digitacao():
    global processado
    limpar_csvs()
    processado = False
    for w in frame_botoes_csv.winfo_children():
        w.destroy()


def gerar_soql_por_cpf():
    if not caminho_arquivo:
        return

    try:
        xls = pd.ExcelFile(caminho_arquivo)

        aba_alvo = None
        for nome in xls.sheet_names:
            if "account" in nome.lower() or "clientes" in nome.lower():
                aba_alvo = nome
                break

        if not aba_alvo:
            messagebox.showerror(
                "Erro",
                "Nenhuma aba contendo 'Account' ou 'Clientes' foi encontrada."
            )
            return

        df = pd.read_excel(xls, sheet_name=aba_alvo)

        if "CPF__pc" not in df.columns:
            messagebox.showerror(
                "Erro",
                "A coluna CPF__pc não foi encontrada na aba selecionada."
            )
            return

        cpfs = (
            df["CPF__pc"]
            .dropna()
            .astype(str)
            .str.replace(r"\D", "", regex=True)
            .apply(lambda x: x.zfill(11))
            .unique()
            .tolist()
        )

        if not cpfs:
            messagebox.showerror("Erro", "Nenhum CPF válido encontrado.")
            return

        lista_cpfs = ", ".join(f"'{c}'" for c in cpfs)

        soql = (
            "SELECT Id, Name, CPF__pc\n"
            "FROM Account\n"
            f"WHERE CPF__pc IN ({lista_cpfs})\n"
            "ORDER BY Name"
        )

        root.clipboard_clear()
        root.clipboard_append(soql)
        root.update()

        messagebox.showinfo(
            "SOQL copiado",
            f"SOQL gerado com {len(cpfs)} CPF(s) e copiado para a área de transferência."
        )

    except Exception as e:
        messagebox.showerror("Erro ao gerar SOQL", str(e))


def atualizar_botao_soql():
    if "btn_soql_cpf" not in globals():
        return

    ids_preenchidos = any(
        l.strip() for l in text_contas.get("1.0", tk.END).splitlines()
    )

    if caminho_arquivo and not ids_preenchidos:
        if not btn_soql_cpf.winfo_ismapped():
            btn_soql_cpf.pack(side=tk.LEFT, padx=5)
    else:
        if btn_soql_cpf.winfo_ismapped():
            btn_soql_cpf.pack_forget()


def atualizar_contador_ids(event=None):
    root.after(1, lambda: _contar_ids())

def _contar_ids():
    linhas = text_contas.get("1.0", tk.END).splitlines()
    total = sum(1 for l in linhas if l.strip())
    label_contador.config(text=str(total))

def centralizar_janela(root, largura, altura):
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (largura // 2)
    y = (root.winfo_screenheight() // 2) - (altura // 2)
    root.geometry(f"{largura}x{altura}+{x}+{y}")

def limpar_csvs():
    if not pasta_saida:
        return
    for nome in ["Account.csv", "Contract.csv", "Asset.csv"]:
        caminho = os.path.join(pasta_saida, nome)
        if os.path.exists(caminho):
            try:
                os.remove(caminho)
            except:
                pass

def resetar_interface(*args):
    global caminho_arquivo, pasta_saida, processado
    limpar_csvs()
    caminho_arquivo = None
    pasta_saida = None
    processado = False
    label_status.config(text="Nenhum arquivo anexado", fg="red")
    label_contador.config(text="0")
    for w in frame_botoes_csv.winfo_children():
        w.destroy()
    atualizar_botao_soql()

def ao_fechar():
    limpar_csvs()
    root.destroy()

# ==================== FUNÇÕES PRINCIPAIS ====================

def anexar_arquivo():
    global caminho_arquivo
    limpar_csvs()
    caminho = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    if caminho:
        caminho_arquivo = caminho
        label_status.config(text="✅ OK", fg="green")
        atualizar_botao_soql()

def abrir_csv(nome):
    caminho = os.path.join(pasta_saida, nome)
    with open(caminho, "r", encoding="utf-8-sig") as f:
        root.clipboard_clear()
        root.clipboard_append(f.read())
        root.update()
    messagebox.showinfo("Copiado", f"{nome} copiado para a área de transferência.")

# ==================== PROCESSAMENTO ====================

def processar_planilha():
    global pasta_saida, processado

    ids = [i.strip() for i in text_contas.get("1.0", tk.END).splitlines() if i.strip()]
    if not ids:
        messagebox.showerror("Erro", "Informe ao menos um Account ID.")
        return

    if not caminho_arquivo:
        messagebox.showerror("Erro", "Anexe o arquivo Excel.")
        return

    try:
        xls = pd.ExcelFile(caminho_arquivo)

        abas = xls.sheet_names

        for aba in ["Account", "Contract", "Ativo"]:
            if not any(aba.lower() in nome.lower() for nome in abas):
                messagebox.showerror(
                    "Erro",
                    f"Aba '{aba}' não foi encontrada no arquivo Excel."
                )
                return


        mapa_renomeacao = {
            "ContractNumber": "_ContractNumber",
            "Account.Name": "_Account.Name",
            "IDExternoAX__c": "_IDExternoAX__c",
            "EndDate": "_EndDate",
            "RecordType.DeveloperName": "_RecordType.DeveloperName",
            # INSERIDO ABAIXO POR CAUSA DESTES CAMPOS TRAVAREM A ATUALIZAÇÃO, PORÉM VIA UPDATE POSTERIOR SE RESOLVE.
            "IRIS_Codigo_Status_do_Tanque__c": "_IRIS_Codigo_Status_do_Tanque__c",
            "IRIS_Codigo_Situacao_do_Agendamento__c": "_IRIS_Codigo_Situacao_do_Agendamento__c"
            
        }

        # ================= ACCOUNT =================
        df_account = pd.read_excel(
            xls, sheet_name=[s for s in xls.sheet_names if "account" in s.lower()][0]
        )

        df_account["Id"] = ids
        df_account["Id"] = df_account["Id"].astype(str)


        df_account.rename(
            columns=lambda c: "Email__c" if isinstance(c, str) and c.strip().lower() == "email" else c,
            inplace=True
        )
        
        if len(df_account) != len(ids):
            messagebox.showerror(
            "Erro",
            "Quantidade de Account IDs diferente da quantidade de registros da aba Account."
            )
            return


        if "CPF__pc" in df_account.columns:
            df_account["CPF__pc"] = df_account["CPF__pc"].astype(str).str.strip()
            df_account["CPF__pc"] = df_account["CPF__pc"].apply(
                lambda x: ("0" + x) if len(x) == 10 else x
            )

        df_account["RecordTypeId"] = "0125A0000013RxeQAE"

        if "AreaNegocio__c" in df_account.columns:
            df_account["AreaNegocio__c"] = "Leves"

        df_account.rename(columns=mapa_renomeacao, inplace=True)

        # ================= CONTRACT =================
        df_contract = pd.read_excel(
            xls, sheet_name=[s for s in xls.sheet_names if "contract" in s.lower()][0]
        )

        df_contract["AccountId"] = ids
        df_contract.drop(columns=["Id"], inplace=True, errors="ignore")
        df_contract = df_contract[["AccountId"] + [c for c in df_contract.columns if c != "AccountId"]]

        df_contract["AccountId"] = df_contract["AccountId"].astype(str)

        df_contract["Status"] = "Draft"
        df_contract["IRIS_Categoria_Contrato__c"] = "2"

        if len(df_contract) != len(ids):
            messagebox.showerror(
            "Erro",
            "Quantidade de Account IDs diferente da quantidade de registros da aba Contract."
            )
            return

        for col in df_contract.columns:
            if "date" in col.lower() or "data" in col.lower():
                df_contract[col] = pd.to_datetime(
                    df_contract[col], errors="coerce"
                ).dt.strftime("%Y-%m-%d")

        for campo in [
            "IRIS_CapturaReservaPrimeiraParcela__c",
            "IRIS_ReservaPrimeiraParcela__c"
        ]:
            if campo in df_contract.columns:
                df_contract[campo] = "TRUE"

        df_contract["RecordTypeId"] = "0125A0000013RxeQAE"

        df_contract.rename(columns=mapa_renomeacao, inplace=True)

        # ================= ASSET =================
        df_asset = pd.read_excel(
            xls, sheet_name=[s for s in xls.sheet_names if "ativo" in s.lower()][0]
        )

        df_asset["AccountId"] = ids
        df_asset.drop(columns=["Id"], inplace=True, errors="ignore")
        df_asset = df_asset[["AccountId"] + [c for c in df_asset.columns if c != "AccountId"]]

        df_asset["AccountId"] = df_asset["AccountId"].astype(str)

        if len(df_asset) != len(ids):
            messagebox.showerror(
            "Erro",
            "Quantidade de Account IDs diferente da quantidade de registros da aba Assets."
            )
            return

        for col in df_asset.columns:
            if "date" in col.lower() or "data" in col.lower():
                df_asset[col] = pd.to_datetime(
                    df_asset[col], errors="coerce"
                ).dt.strftime("%Y-%m-%d")

        if "RecordType.Name" in df_asset.columns:
            df_asset.rename(columns={"RecordType.Name": "RecordTypeId"}, inplace=True)

        if "RecordTypeId" in df_asset.columns:
            df_asset["RecordTypeId"] = "012HY0000004NyFYAU"

        df_asset.rename(columns=mapa_renomeacao, inplace=True)

        # ================= SALVAR =================
        pasta_saida = os.path.dirname(caminho_arquivo)

        df_account.to_csv(os.path.join(pasta_saida, "Account.csv"), index=False, encoding="utf-8-sig")
        df_contract.to_csv(os.path.join(pasta_saida, "Contract.csv"), index=False, encoding="utf-8-sig")
        df_asset.to_csv(os.path.join(pasta_saida, "Asset.csv"), index=False, encoding="utf-8-sig")

        processado = True
        
        messagebox.showinfo("Sucesso", "Processamento concluído e CSVs gerados com sucesso!")

        for w in frame_botoes_csv.winfo_children():
            w.destroy()

        for i in range(3):
            frame_botoes_csv.grid_columnconfigure(i, weight=1)

        tk.Button(frame_botoes_csv, text="Account",
                  command=lambda: abrir_csv("Account.csv"),
                  bg="#2196F3", fg="white").grid(row=0, column=0, padx=10)

        tk.Button(frame_botoes_csv, text="Contract",
                  command=lambda: abrir_csv("Contract.csv"),
                  bg="#2196F3", fg="white").grid(row=0, column=1, padx=10)

        tk.Button(frame_botoes_csv, text="Asset",
                  command=lambda: abrir_csv("Asset.csv"),
                  bg="#2196F3", fg="white").grid(row=0, column=2, padx=10)
        
    except Exception as e:
        messagebox.showerror("Erro ao processar", str(e))

# ==================== INTERFACE ====================

root = tk.Tk()
root.title("Conversor de Planilha para Importação via CSV "+versao+" - Aggrandize - João Márcio Bicalho Andrade")
centralizar_janela(root, 700, 450)
root.protocol("WM_DELETE_WINDOW", ao_fechar)

tk.Label(root, text="Insira os Account IDs (1 por linha):").pack(pady=5)

frame_ids = tk.Frame(root)
frame_ids.pack(pady=5)

# ----- Frame do Text + Scroll -----
frame_text = tk.Frame(frame_ids)
frame_text.pack(side=tk.LEFT)

scroll_ids = tk.Scrollbar(frame_text)
scroll_ids.pack(side=tk.RIGHT, fill=tk.Y)

text_contas = tk.Text(
    frame_text,
    height=6,
    width=60,
    yscrollcommand=scroll_ids.set
)
text_contas.pack(side=tk.LEFT)

scroll_ids.config(command=text_contas.yview)

# ----- Contador -----
label_contador = tk.Label(
    frame_ids,
    text="0",
    font=("Verdana", 10, "bold"),
    width=4
)
label_contador.pack(side=tk.LEFT, padx=6)

text_contas.bind(
    "<Control-v>",
    lambda e: (atualizar_contador_ids(e), text_contas.after(1, lambda: text_contas.see(tk.END)))
)
text_contas.bind(
    "<Control-V>",
    lambda e: (atualizar_contador_ids(e), text_contas.after(1, lambda: text_contas.see(tk.END)))
)

text_contas.bind("<Return>", atualizar_contador_ids)

def on_keypress(event):
    resetar_pos_digitacao()
    atualizar_botao_soql()

text_contas.bind("<KeyPress>", on_keypress)

frame_anexo = tk.Frame(root)
frame_anexo.pack(pady=5)

btn_anexar = tk.Button(
    frame_anexo,
    text="Anexar Arquivo",
    command=anexar_arquivo
)
btn_anexar.pack(side=tk.LEFT, padx=5)

btn_soql_cpf = tk.Button(
    frame_anexo,
    text="Gerar SOQL por CPF",
    command=gerar_soql_por_cpf,
    bg="#673AB7",
    fg="white"
)

label_status = tk.Label(root, text="Nenhum arquivo anexado", fg="red")
label_status.pack(pady=5)

tk.Label(root, text="Observações:").pack(pady=5)

text_obs = tk.Text(root, height=5, width=70)
text_obs.pack(pady=5)
text_obs.insert(
    tk.END,
    "1- Account  -> UPSERT -> CPF__pc\n"
    "2- Contract -> INSERT\n"
    "3- Asset    -> UPSERT -> Placa__c\n\n"
    "Versão Multi-ID: Precisa organizar o XLSX manualmente, por enquanto."
)
text_obs.config(state=tk.DISABLED)

tk.Button(
    root,
    text="Processar e Salvar CSV",
    bg="#4CAF50",
    fg="white",
    command=processar_planilha
).pack(pady=15)

frame_botoes_csv = tk.Frame(root)
frame_botoes_csv.pack(pady=10, fill="x")

root.mainloop()
