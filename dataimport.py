import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

# ==================== VARIÁVEIS GLOBAIS ====================

caminho_arquivo = None
pasta_saida = None
processado = False

# ==================== FUNÇÕES DE UTILIDADE ====================

def centralizar_janela(root, largura, altura):
    root.update_idletasks()

    tela_largura = root.winfo_screenwidth()
    tela_altura = root.winfo_screenheight()

    x = (tela_largura // 2) - (largura // 2)
    y = (tela_altura // 2) - (altura // 2)

    root.geometry(f"{largura}x{altura}+{x}+{y}")


def limpar_csvs():
    global pasta_saida
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

    for widget in frame_botoes_csv.winfo_children():
        widget.destroy()


def ao_fechar():
    limpar_csvs()
    root.destroy()

# ==================== FUNÇÕES PRINCIPAIS ====================

def anexar_arquivo():
    global caminho_arquivo, processado

    limpar_csvs()
    processado = False

    caminho = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )

    if caminho:
        caminho_arquivo = caminho
        label_status.config(text="✅ OK", fg="green")
    else:
        label_status.config(text="Nenhum arquivo anexado", fg="red")


def abrir_csv(nome_arquivo):
    try:
        caminho_completo = os.path.join(pasta_saida, nome_arquivo)

        if not os.path.exists(caminho_completo):
            messagebox.showerror("Erro", f"Arquivo {nome_arquivo} não encontrado.")
            return

        with open(caminho_completo, "r", encoding="utf-8-sig") as f:
            conteudo = f.read()

        root.clipboard_clear()
        root.clipboard_append(conteudo)
        root.update()

        messagebox.showinfo(
            "Copiado!",
            f"O conteúdo de '{nome_arquivo}' foi copiado para a área de transferência."
        )

    except Exception as e:
        messagebox.showerror("Erro ao copiar CSV", str(e))


def processar_planilha():
    global pasta_saida, processado

    conta_id = entry_conta.get().strip()

    if not conta_id:
        messagebox.showerror("Erro", "Digite o Account ID antes de continuar.")
        return

    if not caminho_arquivo:
        messagebox.showerror("Erro", "Anexe o arquivo Excel antes de continuar.")
        return

    try:
        xls = pd.ExcelFile(caminho_arquivo)
        abas = xls.sheet_names

        for aba in ["Account", "Contract", "Ativo"]:
            if not any(aba.lower() in nome.lower() for nome in abas):
                messagebox.showerror("Erro", f"Aba '{aba}' não foi encontrada.")
                return

        mapa_renomeacao = {
            "ContractNumber": "_ContractNumber",
            "Account.Name": "_Account.Name",
            "IDExternoAX__c": "_IDExternoAX__c",
            "EndDate": "_EndDate",
            "RecordType.DeveloperName": "_RecordType.DeveloperName"
        }

        # ================= ACCOUNT =================
        aba_account = [a for a in abas if "account" in a.lower()][0]
        df_account = pd.read_excel(xls, sheet_name=aba_account)

        if "Id" in df_account.columns:
            df_account["Id"] = conta_id
        else:
            df_account.insert(0, "Id", conta_id)

        df_account.rename(
            columns=lambda c: "Email__c" if isinstance(c, str) and c.strip().lower() == "email" else c,
            inplace=True
        )

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
        aba_contract = [a for a in abas if "contract" in a.lower()][0]
        df_contract = pd.read_excel(xls, sheet_name=aba_contract)

        df_contract.rename(
            columns=lambda c: "AccountId" if isinstance(c, str) and c.lower() == "id" else c,
            inplace=True
        )

        if "AccountId" in df_contract.columns:
            df_contract["AccountId"] = conta_id

        df_contract["Status"] = "Draft"
        df_contract["IRIS_Categoria_Contrato__c"] = "2"

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

        # ================= ATIVO =================
        aba_ativo = [a for a in abas if "ativo" in a.lower()][0]
        df_ativo = pd.read_excel(xls, sheet_name=aba_ativo)

        df_ativo.rename(
            columns=lambda c: "AccountId" if isinstance(c, str) and c.lower() == "id" else c,
            inplace=True
        )

        if "AccountId" in df_ativo.columns:
            df_ativo["AccountId"] = conta_id

        for col in df_ativo.columns:
            if "date" in col.lower() or "data" in col.lower():
                df_ativo[col] = pd.to_datetime(
                    df_ativo[col], errors="coerce"
                ).dt.strftime("%Y-%m-%d")

        df_ativo.rename(
            columns=lambda c: "RecordTypeId" if c == "RecordType.Name" else c,
            inplace=True
        )

        if "RecordTypeId" in df_ativo.columns:
            df_ativo["RecordTypeId"] = "012HY0000004NyFYAU"

        df_ativo.rename(columns=mapa_renomeacao, inplace=True)

        # ================= SALVAR CSV =================
        pasta_saida = os.path.dirname(caminho_arquivo)

        df_account.to_csv(os.path.join(pasta_saida, "Account.csv"), index=False, encoding="utf-8-sig")
        df_contract.to_csv(os.path.join(pasta_saida, "Contract.csv"), index=False, encoding="utf-8-sig")
        df_ativo.to_csv(os.path.join(pasta_saida, "Asset.csv"), index=False, encoding="utf-8-sig")

        processado = True

        messagebox.showinfo("Sucesso", "Processamento concluído e CSVs gerados com sucesso!")

        for widget in frame_botoes_csv.winfo_children():
            widget.destroy()

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
root.title("Conversor de Planilha para Importação via CSV 1.6 - Aggrandize - João Márcio Bicalho Andrade")
centralizar_janela(root, 700, 450)
root.protocol("WM_DELETE_WINDOW", ao_fechar)

tk.Label(root, text="Insira o Account ID:").pack(pady=5)

entry_conta = tk.Entry(root, width=50)
entry_conta.pack(pady=5)
entry_conta.bind("<KeyPress>", resetar_interface)

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
