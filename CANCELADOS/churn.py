import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
from sqlalchemy import create_engine
from dotenv import load_dotenv
import os, urllib

load_dotenv()

def gerar_engine():
    conn_str = (
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={os.getenv('SERVER2')};DATABASE={os.getenv('DB_NAME2')};"
        f"UID={os.getenv('DB_USER')};PWD={os.getenv('DB_PASS')};"
    )
    return create_engine(f"mssql+pyodbc:///?odbc_connect={urllib.parse.quote_plus(conn_str)}")

# ── Queries ────────────────────────────────────────────────
CTE_BASE = """
WITH UltimoBloqueio AS (
    SELECT BCA_MATRIC, BCA_TIPREG, MAX(R_E_C_N_O_) AS UltimoRecno
    FROM BCA010
    WHERE D_E_L_E_T_ = '' AND BCA_TIPO = '0'
      AND BCA_DATA BETWEEN ? AND ?
    GROUP BY BCA_MATRIC, BCA_TIPREG
),
Base_Calculo AS (
    SELECT
        BA1.BA1_MATVID,
        BA1.BA1_VTCLAS,
        DATEDIFF(MONTH, CONVERT(DATE, BA1.BA1_DATINC), CONVERT(DATE, BCA.BCA_DATA)) AS Meses,
        DATEDIFF(YEAR,  CONVERT(DATE, BA1.BA1_DATNAS), GETDATE()) AS Idade
    FROM BA1010 BA1
    INNER JOIN UltimoBloqueio UB ON (
        UB.BCA_MATRIC = BA1.BA1_CODINT + BA1.BA1_CODEMP + BA1.BA1_MATRIC
        AND UB.BCA_TIPREG = BA1.BA1_TIPREG
    )
    INNER JOIN BCA010 BCA ON (BCA.R_E_C_N_O_ = UB.UltimoRecno AND BCA.D_E_L_E_T_ = '')
    INNER JOIN BI3010 BI3 ON (BI3.BI3_CODIGO = BA1.BA1_CODPLA AND BI3.D_E_L_E_T_ = '')
    INNER JOIN BG3010 BG3 ON (BG3.BG3_CODBLO = BCA.BCA_MOTBLO AND BG3.D_E_L_E_T_ = '')
    WHERE BA1.D_E_L_E_T_ = ''
      AND BA1.BA1_CODEMP <> '0003'
      AND BI3.BI3_VTMOD IN ('1', '2')
      AND BG3.BG3_DESBLO IN (
          'MOTIVO REVERSSIVEIS', 'CANC.AUT. FALTA DE PAGAMENTO',
          'DIFICULDADE FINANCEIRA', 'DESINTERESSE PELO SERVICO',
          'POSSUI OUTROS PLANOS', 'MOTIVOS PARTICULARES',
          'BLOQUEIO CHAPP', 'INSATISFACAO C/ O SERVICO'
      )
)
"""

query_tempo_casa = CTE_BASE + """
, Categorias AS (
    SELECT BA1_MATVID,
        CASE
            WHEN Meses <=   6 THEN '1. Até 6 meses'
            WHEN Meses <=  12 THEN '2. 7 meses a 1 ano'
            WHEN Meses <=  24 THEN '3. 13 meses a 2 anos'
            WHEN Meses <=  60 THEN '4. 2 anos e 1 mês a 5 anos'
            WHEN Meses <= 120 THEN '5. 5 anos e 1 mês a 10 anos'
            WHEN Meses <= 240 THEN '6. 10 anos 1 mês a 20 anos'
            ELSE                   '7. Acima de 20 anos'
        END AS Tempo_de_casa
    FROM Base_Calculo
)
SELECT ISNULL(Tempo_de_casa, 'TOTAL') AS Tempo_de_casa, COUNT(*) AS Qtd
FROM Categorias
GROUP BY ROLLUP (Tempo_de_casa)
ORDER BY CASE WHEN Tempo_de_casa IS NULL THEN 2 ELSE 1 END, Tempo_de_casa;
"""

query_faixa_etaria = CTE_BASE + """
, Categorias AS (
    SELECT BA1_MATVID,
        CASE
            WHEN Idade BETWEEN  0 AND  9 THEN '01. 0 a 9 anos'
            WHEN Idade BETWEEN 10 AND 19 THEN '02. 10 a 19 anos'
            WHEN Idade BETWEEN 20 AND 29 THEN '03. 20 a 29 anos'
            WHEN Idade BETWEEN 30 AND 39 THEN '04. 30 a 39 anos'
            WHEN Idade BETWEEN 40 AND 49 THEN '05. 40 a 49 anos'
            WHEN Idade BETWEEN 50 AND 59 THEN '06. 50 a 59 anos'
            WHEN Idade BETWEEN 60 AND 69 THEN '07. 60 a 69 anos'
            ELSE                               '08. 70 anos ou mais'
        END AS Faixa_Etaria
    FROM Base_Calculo
)
SELECT ISNULL(Faixa_Etaria, 'TOTAL') AS Faixa_Etaria, COUNT(*) AS Qtd
FROM Categorias
GROUP BY ROLLUP (Faixa_Etaria)
ORDER BY CASE WHEN Faixa_Etaria IS NULL THEN 2 ELSE 1 END, Faixa_Etaria;
"""

query_classificacao = CTE_BASE + """
, Categorias AS (
    SELECT
        CASE LTRIM(RTRIM(BA1_VTCLAS))
            WHEN '1'   THEN '1 - Classe 1'
            WHEN '2'   THEN '2 - Classe 2'
            WHEN '3'   THEN '3 - Classe 3'
            WHEN '4'   THEN '4 - Classe 4'
            WHEN '5'   THEN '5 - Classe 5'
            WHEN 'IND' THEN 'Individual'
            WHEN 'PRE' THEN 'Premium'
            WHEN 'VER' THEN 'VIP'
            ELSE BA1_VTCLAS
        END AS ClassificacaoPlano
    FROM Base_Calculo
)
SELECT ISNULL(ClassificacaoPlano, 'TOTAL') AS ClassificacaoPlano, COUNT(*) AS Qtd
FROM Categorias
GROUP BY ROLLUP (ClassificacaoPlano)
ORDER BY CASE WHEN ClassificacaoPlano IS NULL THEN 2 ELSE 1 END, ClassificacaoPlano;
"""

RELATORIOS = [
    (query_tempo_casa,    'Tempo de Casa'),
    (query_faixa_etaria,  'Faixa Etária'),
    (query_classificacao, 'Classificação de Plano'),
]

# ── Cores & Tema ───────────────────────────────────────────
BG          = "#1e1e2e"
CARD        = "#2a2a3e"
ACCENT      = "#7c6af7"
ACCENT_HOV  = "#6a58e0"
TEXT        = "#e0e0f0"
TEXT_MUTED  = "#7070a0"
INPUT_BG    = "#12121e"
SUCCESS     = "#4ade80"
ERROR       = "#f87171"
BORDER      = "#3a3a5e"
FONT        = "Segoe UI"

# ── Lógica ─────────────────────────────────────────────────
def rodar_relatorios():
    data_ini, data_fim = entry_inicio.get().strip(), entry_fim.get().strip()

    if data_ini == "AAAAMMDD" or data_fim == "AAAAMMDD" or \
       len(data_ini) != 8 or len(data_fim) != 8:
        flash_status("⚠  Use o formato AAAAMMDD (ex: 20250101)", ERROR)
        return

    set_loading(True)

    try:
        engine = gerar_engine()
        nome_arquivo = f"Churn_Completo_{data_ini}_{data_fim}.xlsx"

        with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
            for i, (query, aba) in enumerate(RELATORIOS):
                flash_status(f"⏳  Processando: {aba}…", ACCENT)
                root.update()
                df = pd.read_sql(query, engine, params=(data_ini, data_fim))
                df.to_excel(writer, sheet_name=aba, index=False)

        flash_status(f"✔  {nome_arquivo} gerado com sucesso!", SUCCESS)

    except Exception as e:
        flash_status(f"✖  Erro: {str(e)[:80]}", ERROR)
        messagebox.showerror("Erro", str(e))
    finally:
        set_loading(False)
        if 'engine' in locals():
            engine.dispose()

def set_loading(state):
    if state:
        btn.config(state="disabled", text="  Processando…")
        progress.start(12)
    else:
        btn.config(state="normal", text="  Gerar Relatórios")
        progress.stop()
        progress_var.set(0)

def flash_status(msg, color):
    lbl_status.config(text=msg, fg=color)
    root.update()

def on_focus_in(entry, placeholder):
    if entry.get() == placeholder:
        entry.delete(0, tk.END)
        entry.config(fg=TEXT)

def on_focus_out(entry, placeholder):
    if not entry.get():
        entry.insert(0, placeholder)
        entry.config(fg=TEXT_MUTED)

# ── Janela ─────────────────────────────────────────────────
root = tk.Tk()
root.title("VitalMed BI churn reversível - v1.3")
root.geometry("420x480")
root.resizable(False, False)
root.configure(bg=BG)

# Centraliza na tela
root.update_idletasks()
x = (root.winfo_screenwidth()  - 420) // 2
y = (root.winfo_screenheight() - 480) // 2
root.geometry(f"420x480+{x}+{y}")

# ── Card central ───────────────────────────────────────────
card = tk.Frame(root, bg=CARD, bd=0, highlightthickness=1,
                highlightbackground=BORDER)
card.place(relx=0.5, rely=0.5, anchor="center", width=360, height=420)

# Faixa superior colorida
tk.Frame(card, bg=ACCENT, height=4).pack(fill="x")

# Ícone + Título
tk.Label(card, text="📊", font=(FONT, 26), bg=CARD, fg=TEXT).pack(pady=(18, 0))
tk.Label(card, text="Relatório de Churn", font=(FONT, 14, "bold"),
         bg=CARD, fg=TEXT).pack()
tk.Label(card, text="Motivos Reversíveis", font=(FONT, 9),
         bg=CARD, fg=TEXT_MUTED).pack(pady=(2, 18))

# Separador
tk.Frame(card, bg=BORDER, height=1).pack(fill="x", padx=20)

# ── Campos de data ─────────────────────────────────────────
def campo(parent, label, placeholder):
    wrap = tk.Frame(parent, bg=CARD)
    wrap.pack(fill="x", padx=24, pady=(14, 0))

    tk.Label(wrap, text=label, font=(FONT, 8, "bold"),
             bg=CARD, fg=TEXT_MUTED).pack(anchor="w")

    row = tk.Frame(wrap, bg=INPUT_BG, highlightthickness=1,
                   highlightbackground=BORDER)
    row.pack(fill="x", pady=(4, 0))

    tk.Label(row, text="📅", bg=INPUT_BG, fg=TEXT_MUTED,
             font=(FONT, 10)).pack(side="left", padx=(8, 0))

    e = tk.Entry(row, bg=INPUT_BG, fg=TEXT_MUTED, insertbackground=TEXT,
                 relief="flat", font=("Consolas", 11), bd=4,
                 highlightthickness=0)
    e.pack(side="left", fill="x", expand=True)
    e.insert(0, placeholder)

    e.bind("<FocusIn>",  lambda ev: on_focus_in(e, placeholder))
    e.bind("<FocusOut>", lambda ev: on_focus_out(e, placeholder))

    # highlight ao focar no frame
    e.bind("<FocusIn>",  lambda ev: (on_focus_in(e, placeholder),
                                     row.config(highlightbackground=ACCENT)), "+")
    e.bind("<FocusOut>", lambda ev: (on_focus_out(e, placeholder),
                                     row.config(highlightbackground=BORDER)), "+")
    return e

entry_inicio = campo(card, "DATA INÍCIO", "AAAAMMDD")
entry_fim    = campo(card, "DATA FIM",    "AAAAMMDD")

# ── Botão ──────────────────────────────────────────────────
btn_frame = tk.Frame(card, bg=CARD)
btn_frame.pack(fill="x", padx=24, pady=(22, 0))

btn = tk.Button(
    btn_frame,
    text="  Gerar Relatórios",
    font=(FONT, 10, "bold"),
    bg=ACCENT, fg="white",
    activebackground=ACCENT_HOV, activeforeground="white",
    relief="flat", bd=0, cursor="hand2",
    pady=10, command=rodar_relatorios
)
btn.pack(fill="x")
btn.bind("<Enter>", lambda e: btn.config(bg=ACCENT_HOV))
btn.bind("<Leave>", lambda e: btn.config(bg=ACCENT))

# ── Progress bar ───────────────────────────────────────────
progress_var = tk.DoubleVar()
progress = ttk.Progressbar(card, mode="indeterminate",
                            variable=progress_var, length=312)

style = ttk.Style()
style.theme_use("clam")
style.configure("TProgressbar", troughcolor=INPUT_BG,
                bordercolor=INPUT_BG, background=ACCENT, lightcolor=ACCENT,
                darkcolor=ACCENT)
progress.pack(padx=24, pady=(10, 0), fill="x")

# ── Status ─────────────────────────────────────────────────
lbl_status = tk.Label(card, text="", font=(FONT, 8),
                      bg=CARD, fg=TEXT_MUTED, wraplength=310)
lbl_status.pack(pady=(8, 0))

# ── Rodapé ─────────────────────────────────────────────────
tk.Frame(card, bg=BORDER, height=1).pack(fill="x", padx=20, pady=(10, 0))
abas_str = "  ·  ".join(aba for _, aba in RELATORIOS)
tk.Label(card, text=f"v1.3  ·  {abas_str}", font=(FONT, 7),
         bg=CARD, fg=TEXT_MUTED).pack(pady=(6, 12))

root.mainloop()