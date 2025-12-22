import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import subprocess
import datetime

# ==================== FUNÇÕES DE LIMPEZA ====================

def limpar_csvs():
    global pasta_saida
    if not pasta_saida:
        return

    arquivos = [
        "Account.csv",
        "Contract.csv",
        "Asset.csv"
    ]

    for nome in arquivos:
        caminho = os.path.join(pasta_saida, nome)
        if os.path.exists(caminho):
            try:
                os.remove(caminho)
            except:
                pass  # evita erros caso o arquivo esteja aberto


def resetar_interface(*args):
    limpar_csvs()

    # reseta exibição
    label_status.config(text="Nenhum arquivo anexado", fg="red")

    # limpa botões
    for widget in frame_botoes_csv.winfo_children():
        widget.destroy()

    # reseta variáveis globais
    global caminho_arquivo, pasta_saida, processado
    caminho_arquivo = None
    pasta_saida = None
    processado = False


# ==================== FUNÇÕES PRINCIPAIS ====================

def anexar_arquivo():
    global caminho_arquivo
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

        # Lê o conteúdo do CSV
        with open(caminho_completo, "r", encoding="utf-8-sig") as f:
            conteudo = f.read()

        # Copia para a área de transferência
        root.clipboard_clear()
        root.clipboard_append(conteudo)
        root.update()  # garante que a cópia funcione em todas as plataformas

        messagebox.showinfo("Copiado!",
                            f"O conteúdo de '{nome_arquivo}' foi copiado para a área de transferência.")

    except Exception as e:
        messagebox.showerror("Erro ao copiar CSV", str(e))



def processar_planilha():
    global pasta_saida, processado
    conta_id = entry_conta.get().strip()

    if not conta_id:
        messagebox.showerror("Erro", "Digite o ID da conta antes de continuar.")
        return

    if not caminho_arquivo:
        messagebox.showerror("Erro", "Anexe o arquivo Excel antes de continuar.")
        return

    try:
        xls = pd.ExcelFile(caminho_arquivo)
        abas = xls.sheet_names

        obrigatorias = ["Account", "Contract", "Ativo"]
        for aba in obrigatorias:
            if not any(aba.lower() in nome.lower() for nome in abas):
                messagebox.showerror("Erro", f"Aba '{aba}' não foi encontrada.")
                return

        # === Aba Account ===
        aba_account = [a for a in abas if "account" in a.lower()][0]
        df_account = pd.read_excel(xls, sheet_name=aba_account)

        if "Id" in df_account.columns:
            df_account["Id"].fillna(conta_id, inplace=True)
        else:
            df_account.insert(0, "Id", conta_id)

        df_account.rename(columns=lambda c: "Email__c" if isinstance(c, str) and c.strip().lower() == "email" else c, inplace=True)

        if "CPF__pc" in df_account.columns:
            df_account["CPF__pc"] = df_account["CPF__pc"].astype(str).str.strip()
            df_account["CPF__pc"] = df_account["CPF__pc"].apply(lambda x: ("0" + x) if len(x) == 10 else x)

        df_account["RecordTypeId"] = "0125A0000013RxeQAE"
        if "AreaNegocio__c" in df_account.columns:
            df_account["AreaNegocio__c"] = "Leves"


        # === Aba Contract ===
        aba_contract = [a for a in abas if "contract" in a.lower()][0]
        df_contract = pd.read_excel(xls, sheet_name=aba_contract)

        df_contract.rename(columns=lambda c: "AccountId" if isinstance(c, str) and c.lower() == "id" else c, inplace=True)
        if "AccountId" in df_contract.columns:
            df_contract["AccountId"] = conta_id

        df_contract["Status"] = "Draft"
        df_contract["IRIS_Categoria_Contrato__c"] = "2"

        for col in df_contract.columns:
            if "date" in col.lower() or "data" in col.lower():
                df_contract[col] = pd.to_datetime(df_contract[col], errors="coerce").dt.strftime("%Y-%m-%d")

        for campo in ["IRIS_CapturaReservaPrimeiraParcela__c", "IRIS_ReservaPrimeiraParcela__c"]:
            if campo in df_contract.columns:
                df_contract[campo] = "TRUE"


        # === Aba Ativo ===
        aba_ativo = [a for a in abas if "ativo" in a.lower()][0]
        df_ativo = pd.read_excel(xls, sheet_name=aba_ativo)

        df_ativo.rename(columns=lambda c: "AccountId" if isinstance(c, str) and c.lower() == "id" else c, inplace=True)
        if "AccountId" in df_ativo.columns:
            df_ativo["AccountId"] = conta_id

        for col in df_ativo.columns:
            if "date" in col.lower() or "data" in col.lower():
                df_ativo[col] = pd.to_datetime(df_ativo[col], errors="coerce").dt.strftime("%Y-%m-%d")

        df_ativo.rename(columns=lambda c: "RecordTypeId" if c == "RecordType.Name" else c, inplace=True)
        if "RecordTypeId" in df_ativo.columns:
            df_ativo["RecordTypeId"] = "012HY0000004NyFYAU"


        # === Salvar CSV ===
        pasta_saida = os.path.dirname(caminho_arquivo)
        arquivos_csv = {
            "Account.csv": df_account,
            "Contract.csv": df_contract,
            "Asset.csv": df_ativo
        }

        for nome, df in arquivos_csv.items():
            df.to_csv(os.path.join(pasta_saida, nome), index=False, encoding="utf-8-sig")

        # Marca que o processamento foi concluído
        processado = True

        messagebox.showinfo("Sucesso", "Processamento concluído e arquivos CSV gerados com sucesso!")

        # === Criar botões ===
        for widget in frame_botoes_csv.winfo_children():
            widget.destroy()

        frame_botoes_csv.grid_columnconfigure(0, weight=1)
        frame_botoes_csv.grid_columnconfigure(1, weight=1)
        frame_botoes_csv.grid_columnconfigure(2, weight=1)

        btn_account = tk.Button(frame_botoes_csv, text="Account", command=lambda: abrir_csv("Account.csv"), bg="#2196F3", fg="white")
        btn_account.grid(row=0, column=0, padx=10, pady=5)

        btn_contract = tk.Button(frame_botoes_csv, text="Contract", command=lambda: abrir_csv("Contract.csv"), bg="#2196F3", fg="white")
        btn_contract.grid(row=0, column=1, padx=10, pady=5)

        btn_ativo = tk.Button(frame_botoes_csv, text="Asset", command=lambda: abrir_csv("Asset.csv"), bg="#2196F3", fg="white")
        btn_ativo.grid(row=0, column=2, padx=10, pady=5)


    except Exception as e:
        messagebox.showerror("Erro ao processar", str(e))



# ==================== INTERFACE ====================

root = tk.Tk()
root.title("Conversor de Planilha para Importação via CSV 1.6 - Aggrandize - João Márcio Bicalho Andrade")
root.geometry("700x450")

caminho_arquivo = None
pasta_saida = None
processado = False  # indicador novo

label_conta = tk.Label(root, text="Insira o Account ID:")
label_conta.pack(pady=5)

entry_conta = tk.Entry(root, width=50)
entry_conta.pack(pady=5)

# dispara reset quando o usuário altera o Account ID
entry_conta.bind("<KeyPress>", resetar_interface)

btn_anexar = tk.Button(root, text="Anexar Arquivo", command=anexar_arquivo)
btn_anexar.pack(pady=5)

label_status = tk.Label(root, text="Nenhum arquivo anexado", fg="red")
label_status.pack(pady=5)

label_obs = tk.Label(root, text="Observações:")
label_obs.pack(pady=5)

text_obs = tk.Text(root, height=5, width=70)
text_obs.pack(pady=5)
text_obs.insert(tk.END,
    "Account       -> UPSERT -> Caso a conta exista, usar o campo CPF__pc\n"
    "Contract      -> INSERT -> Caso o contrato NÃO exista\n"
    "Asset         -> UPSERT -> Usar o campo Placa__c"
)
text_obs.config(state=tk.DISABLED)

btn_processar = tk.Button(root, text="Processar e Salvar CSV", bg="#4CAF50", fg="white", command=processar_planilha)
btn_processar.pack(pady=15)

frame_botoes_csv = tk.Frame(root)
frame_botoes_csv.pack(pady=10, fill="x")

limpar_csvs()  # regras atuais continuam
root.mainloop()
