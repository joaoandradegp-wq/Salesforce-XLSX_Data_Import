import tkinter as tk
from tkinter import messagebox
import re
import math

soql_partes = []
botoes_partes = []

def gerar_partes():
    global soql_partes, botoes_partes
    soql_partes.clear()
    botoes_partes.clear()

    texto = txt_ids.get("1.0", tk.END).strip()

    if not texto:
        messagebox.showwarning("Aviso", "Cole pelo menos um ID.")
        return

    try:
        limite = int(entry_limite.get())
        if limite <= 0:
            raise ValueError
    except:
        messagebox.showwarning("Aviso", "Informe um número válido de IDs por parte.")
        return

    ids = re.split(r'[\s,;]+', texto)
    ids = [i.strip() for i in ids if i.strip()]

    if not ids:
        messagebox.showwarning("Aviso", "Nenhum ID válido encontrado.")
        return

    total_partes = math.ceil(len(ids) / limite)

    for i in range(total_partes):
        inicio = i * limite
        fim = inicio + limite
        parte_ids = ids[inicio:fim]

        ids_formatados = ",".join(f"'{x}'" for x in parte_ids)

        soql = (
            "SELECT Id, IRIS_Opportunity__r.Id "
            "FROM Contract "
            f"WHERE IRIS_Opportunity__r.Id IN ({ids_formatados})"
        )

        soql_partes.append(soql)

    criar_botoes()
    messagebox.showinfo("Sucesso", f"{total_partes} parte(s) criada(s)!")

def criar_botoes():
    for widget in frame_botoes.winfo_children():
        widget.destroy()

    botoes_partes.clear()

    for idx, _ in enumerate(soql_partes):
        btn = tk.Button(
            frame_botoes,
            text=f"PARTE {idx+1}",
            bg="#6A0DAD",
            fg="white",
            activebackground="#5A0099",
            width=15,
            height=2,
            relief="raised",
            command=lambda i=idx: copiar_parte(i)
        )
        btn.grid(row=idx//4, column=idx%4, padx=5, pady=5)
        botoes_partes.append(btn)

def copiar_parte(indice):
    # Copia SOQL
    root.clipboard_clear()
    root.clipboard_append(soql_partes[indice])
    root.update()

    # 🔁 Reset visual completo de todos
    for btn in botoes_partes:
        btn.config(
            bg="#6A0DAD",
            activebackground="#5A0099",
            relief="raised"
        )

    # 🟢 Destacar o atual
    botoes_partes[indice].config(
        bg="#2ECC71",
        activebackground="#27AE60",
        relief="sunken"
    )

    # 🔥 Remove foco do botão (isso resolve o bug visual do Windows)
    root.focus()

    root.update_idletasks()

    messagebox.showinfo("Copiado", f"SOQL da PARTE {indice+1} copiado!")

# =============================
# Interface
# =============================

root = tk.Tk()
root.title("Gerador de Pesquisa SOQL 1.0 - Oportunidades - Aggrandize - João Márcio Bicalho Andrade")
root.geometry("750x550")

tk.Label(root, text="Cole os IDs das Oportunidades:").pack(pady=5)

frame_texto = tk.Frame(root)
frame_texto.pack(fill="x", padx=15, pady=5)

scrollbar = tk.Scrollbar(frame_texto)
scrollbar.pack(side="right", fill="y")

txt_ids = tk.Text(
    frame_texto,
    height=15,
    yscrollcommand=scrollbar.set
)
txt_ids.pack(side="left", fill="both", expand=True)

scrollbar.config(command=txt_ids.yview)

frame_config = tk.Frame(root)
frame_config.pack(pady=10)

tk.Label(frame_config, text="Quantidade máxima por parte:").pack(side="left", padx=5)

entry_limite = tk.Entry(frame_config, width=10)
entry_limite.insert(0, "500")
entry_limite.pack(side="left")

btn_gerar = tk.Button(
    root,
    text="GERAR PARTES",
    command=gerar_partes,
    bg="#6A0DAD",
    fg="white",
    activebackground="#5A0099",
    height=2
)
btn_gerar.pack(pady=10)

frame_botoes = tk.Frame(root)
frame_botoes.pack(pady=10)

root.mainloop()
