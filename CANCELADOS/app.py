import pandas as pd
import pyodbc
import os
from datetime import datetime
from dotenv import load_dotenv
import tkinter as tk
from tkinter import messagebox
from dados import QUERY_PURA_PARA_RELATORIO, QUERY_PURA_PARA_RELATORIO_FEIRA

load_dotenv()

# =========================
# Paleta de cores
# =========================
COR_BG       = "#1e1e2e"
COR_CARD     = "#2a2a3d"
COR_BORDA    = "#3b3b55"
COR_ACCENT   = "#7c6af7"
COR_HOVER    = "#6355d4"
COR_TEXTO    = "#e0e0f0"
COR_SUBTEXTO = "#9090b0"
COR_ENTRY    = "#33334a"
COR_SEL      = "#4e4e78"

# =========================
# Conexão SQL
# =========================
def get_connection():
    return pyodbc.connect(
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={os.getenv('SERVER2')};"
        f"DATABASE={os.getenv('DB_NAME2')};"
        f"UID={os.getenv('DB_USER')};"
        f"PWD={os.getenv('DB_PASS')};"
    )

# =========================
# Gerar relatório
# =========================
def gerar_relatorio():
    try:
        data_inicio = datetime.strptime(entry_inicio.get(), "%d/%m/%Y")
        data_fim    = datetime.strptime(entry_fim.get(),    "%d/%m/%Y")

        data_inicio_sql = data_inicio.strftime("%Y%m%d")
        data_fim_sql    = data_fim.strftime("%Y%m%d")

        cidade  = cidade_var.get()
        query   = QUERY_PURA_PARA_RELATORIO if cidade == "Salvador" else QUERY_PURA_PARA_RELATORIO_FEIRA
        prefixo = "salvador" if cidade == "Salvador" else "feira"

        conn = get_connection()
        dados_excel = pd.read_sql_query(query, conn, params=[data_inicio_sql, data_fim_sql])
        nome_arquivo = f"cancelados_{prefixo}_{data_inicio_sql}_{data_fim_sql}.xlsx"
        dados_excel.to_excel(nome_arquivo, index=False)
        conn.close()

        messagebox.showinfo("✅ Sucesso", f"Relatório gerado:\n{nome_arquivo}")

    except Exception as e:
        messagebox.showerror("❌ Erro", str(e))

# =========================
# Helpers de estilo
# =========================
def on_enter(btn):
    btn.config(bg=COR_HOVER)

def on_leave(btn):
    btn.config(bg=COR_ACCENT)

def make_toggle(parent, text, value):
    """Radiobutton customizado como toggle pill."""
    rb = tk.Radiobutton(
        parent,
        text=text,
        variable=cidade_var,
        value=value,
        bg=COR_CARD,
        fg=COR_SUBTEXTO,
        activebackground=COR_SEL,
        activeforeground=COR_TEXTO,
        selectcolor=COR_SEL,
        font=("Segoe UI", 10, "bold"),
        indicatoron=False,          # aparência de botão sólido
        relief="flat",
        bd=0,
        padx=18,
        pady=6,
        cursor="hand2",
        command=lambda: atualizar_toggles(),
    )
    rb.config(fg=COR_TEXTO if cidade_var.get() == value else COR_SUBTEXTO)
    return rb

def atualizar_toggles():
    for rb, val in toggle_refs:
        if cidade_var.get() == val:
            rb.config(fg=COR_TEXTO, bg=COR_SEL)
        else:
            rb.config(fg=COR_SUBTEXTO, bg=COR_CARD)

def make_entry(parent, placeholder=""):
    frame = tk.Frame(parent, bg=COR_ENTRY, bd=0, highlightthickness=1,
                     highlightbackground=COR_BORDA, highlightcolor=COR_ACCENT)
    entry = tk.Entry(
        frame,
        font=("Segoe UI", 11),
        bg=COR_ENTRY,
        fg=COR_TEXTO,
        insertbackground=COR_TEXTO,
        relief="flat",
        bd=6,
    )
    entry.insert(0, placeholder)
    entry.pack(fill="x")

    def on_focus_in(e):
        if entry.get() == placeholder:
            entry.delete(0, "end")
            entry.config(fg=COR_TEXTO)
        frame.config(highlightbackground=COR_ACCENT)

    def on_focus_out(e):
        if not entry.get():
            entry.insert(0, placeholder)
            entry.config(fg=COR_SUBTEXTO)
        frame.config(highlightbackground=COR_BORDA)

    entry.config(fg=COR_SUBTEXTO)
    entry.bind("<FocusIn>",  on_focus_in)
    entry.bind("<FocusOut>", on_focus_out)
    return frame, entry

# =========================
# Janela principal
# =========================
janela = tk.Tk()
janela.title("Relatório de Cancelados")
janela.geometry("360x460")
janela.resizable(False, False)
janela.config(bg=COR_BG)

# --- Header ---
header = tk.Frame(janela, bg=COR_BG)
header.pack(fill="x", pady=(24, 0))

tk.Label(header, text="📋", font=("Segoe UI Emoji", 28),
         bg=COR_BG, fg=COR_ACCENT).pack()
tk.Label(header, text="Relatório de Cancelados",
         font=("Segoe UI", 14, "bold"), bg=COR_BG, fg=COR_TEXTO).pack()
tk.Label(header, text="Selecione a cidade e o período desejado",
         font=("Segoe UI", 9), bg=COR_BG, fg=COR_SUBTEXTO).pack(pady=(2, 0))

# --- Card central ---
card = tk.Frame(janela, bg=COR_CARD, bd=0,
                highlightthickness=1, highlightbackground=COR_BORDA)
card.pack(fill="both", padx=24, pady=20, expand=True)

# --- Toggle cidade ---
tk.Label(card, text="Cidade", font=("Segoe UI", 9, "bold"),
         bg=COR_CARD, fg=COR_SUBTEXTO).pack(anchor="w", padx=16, pady=(16, 4))

cidade_var  = tk.StringVar(value="Salvador")
toggle_refs = []  # (widget, value)

toggle_frame = tk.Frame(card, bg=COR_CARD)
toggle_frame.pack(padx=16, fill="x")

for cidade_nome in ("Salvador", "Feira de Santana"):
    rb = make_toggle(toggle_frame, cidade_nome, cidade_nome)
    rb.pack(side="left", padx=(0, 6))
    toggle_refs.append((rb, cidade_nome))

atualizar_toggles()  # aplica estilo inicial

# --- Separador ---
tk.Frame(card, bg=COR_BORDA, height=1).pack(fill="x", padx=16, pady=14)

# --- Data Início ---
tk.Label(card, text="Data Início", font=("Segoe UI", 9, "bold"),
         bg=COR_CARD, fg=COR_SUBTEXTO).pack(anchor="w", padx=16)
frame_ini, entry_inicio = make_entry(card, "dd/mm/aaaa")
frame_ini.pack(fill="x", padx=16, pady=(4, 12))

# --- Data Fim ---
tk.Label(card, text="Data Fim", font=("Segoe UI", 9, "bold"),
         bg=COR_CARD, fg=COR_SUBTEXTO).pack(anchor="w", padx=16)
frame_fim, entry_fim = make_entry(card, "dd/mm/aaaa")
frame_fim.pack(fill="x", padx=16, pady=(4, 20))

# --- Botão ---
btn = tk.Button(
    card,
    text="⬇  Gerar Relatório",
    command=gerar_relatorio,
    font=("Segoe UI", 11, "bold"),
    bg=COR_ACCENT,
    fg="#ffffff",
    activebackground=COR_HOVER,
    activeforeground="#ffffff",
    relief="flat",
    bd=0,
    padx=10,
    pady=10,
    cursor="hand2",
)
btn.pack(fill="x", padx=16, pady=(0, 16))
btn.bind("<Enter>", lambda e: on_enter(btn))
btn.bind("<Leave>", lambda e: on_leave(btn))

janela.mainloop()