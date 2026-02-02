import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

# ==================== VARIÁVEIS GLOBAIS ====================

caminho_arquivo = None
pasta_saida = None
processado = False
versao = "1.6 - Multi"

# ==================== FUNÇÕES DE UTILIDADE ====================

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
    for w in frame_botoes_csv.winfo_children():
        w.destroy()

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

        mapa_renomeacao = {
            "ContractNumber": "_ContractNumber",
            "Account.Name": "_Account.Name",
            "IDExternoAX__c": "_IDExternoAX__c",
            "EndDate": "_EndDate",
            "RecordType.DeveloperName": "_RecordType.DeveloperName"
        }

        # ================= ACCOUNT =================
        df_account = pd.read_excel(
            xls, sheet_name=[s for s in xls.sheet_names if "account" in s.lower()][0]
        )

        df_account["Id"] = ids

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

        df_contract.rename(columns=mapa_renomeacao, inplace=True)

        # ================= ASSET =================
        df_asset = pd.read_excel(
            xls, sheet_name=[s for s in xls.sheet_names if "ativo" in s.lower()][0]
        )

        df_asset["AccountId"] = ids
        df_asset.drop(columns=["Id"], inplace=True, errors="ignore")
        df_asset = df_asset[["AccountId"] + [c for c in df_asset.columns if c != "AccountId"]]

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

text_contas = tk.Text(root, height=6, width=70)
text_contas.pack(pady=5)
text_contas.bind("<KeyPress>", resetar_interface)

tk.Button(root, text="Anexar Arquivo", command=anexar_arquivo).pack(pady=5)

label_status = tk.Label(root, text="Nenhum arquivo anexado", fg="red")
label_status.pack(pady=5)

tk.Label(root, text="Observações:").pack(pady=5)

text_obs = tk.Text(root, height=5, width=70)
text_obs.pack(pady=5)
text_obs.insert(
    tk.END,
    "Account  -> UPSERT -> CPF__pc\n"
    "Contract -> INSERT\n"
    "Asset    -> UPSERT -> Placa__c"
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
