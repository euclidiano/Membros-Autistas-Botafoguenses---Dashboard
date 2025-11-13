import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import os
import threading
from queue import Queue
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.ticker as ticker

root = tk.Tk()
root.withdraw()
root.title("Membros Autistas Botafoguenses - Dashboard")
root.geometry("1920x1080")

CSV_FILE = "dados.csv"
URL_SHEET = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSdZFeAdOutdhsgJqs34JK0KA9egnlqstBHBx-KH-R2z8EhV_-uJbgHLLzJz5r5E_3eCfiqpi7ZhE79/pub?output=csv"

COLUNAS_PRINCIPAIS = [
    "ID","Nome","Nascimento","Telefone","CPF","Email","Profissão",
    "Carteirinha","Ajuda","PCD","TEA","CEP",
]

sync_queue = Queue()
INTERVALO_SYNC_MS = 300000

def sincronizar_dados(mostrar_popup=True):
    df_local = pd.DataFrame(columns=COLUNAS_PRINCIPAIS)
    if os.path.exists(CSV_FILE):
        try:
            df_local = pd.read_csv(CSV_FILE)
            if "Email" in df_local.columns:
                df_local["Email"] = df_local["Email"].astype(str).str.lower().str.strip()
        except pd.errors.EmptyDataError:
            pass
        except Exception as e:
            print(f"Erro ao ler CSV local: {e}")
    try:
        df_remoto = pd.read_csv(URL_SHEET)
        df_remoto.columns = df_remoto.columns.str.replace(r'\s+', ' ', regex=True).str.strip()
        mapa_colunas = {
            "Nome completo": "Nome",
            "Endereço de e-mail": "Email",
            "Deseja receber a carteirinha ?": "Carteirinha",
            "É PCD - se sim, descreva": "PCD",
            "Telefone para contato Ex: 21 9 9999-9999": "Telefone",
            "Trabalha com o público TEA ?": "TEA",
            "Data de nascimento": "Nascimento",
            "Deseja nos ajudar pagando apenas 15 R$ mensais e obter descontos exlusivos ?": "Ajuda",
            "CPF ( Obrigatório somente para membros que escolherem pagar mensalmente )": "CPF",
            "Carimbo de data/hora": "DataEntrada",
            "Endereço: CEP": "CEP"
        }
        df_remoto.rename(columns=mapa_colunas, inplace=True)
        if "DataEntrada" in df_remoto.columns:
            df_remoto["DataEntrada"] = pd.to_datetime(df_remoto["DataEntrada"], errors="coerce")
        colunas_para_manter = [col for col in COLUNAS_PRINCIPAIS if col in df_remoto.columns]
        df_remoto = df_remoto[colunas_para_manter]
        if "Email" in df_remoto.columns:
            df_remoto["Email"] = df_remoto["Email"].astype(str).str.lower().str.strip()
    except Exception as e:
        print(f"AVISO: Não foi possível carregar dados da planilha. Erro: {e}")
        if mostrar_popup:
            messagebox.showwarning("Sem Conexão", "Não foi possível sincronizar. Usando dados locais.")
        return df_local
    df_combinado = pd.concat([df_local, df_remoto], ignore_index=True)
    try:
        df_combinado.to_csv(CSV_FILE, index=False)
        if mostrar_popup:
            messagebox.showinfo("Sincronização", f"Dados sincronizados!\nTotal de {len(df_combinado)} registros carregados.")
    except Exception:
        if mostrar_popup:
            messagebox.showerror("Erro de Arquivo", "Não foi possível salvar o dados.csv.")
    df_combinado = df_combinado.reset_index(drop=True)
    df_combinado["ID"] = df_combinado.index + 1
    return df_combinado.reset_index(drop=True)

def sincronizar_em_background():
    print("Iniciando sincronização em background...")
    try:
        df_novo = sincronizar_dados(mostrar_popup=False)
        if df_novo is not None:
            sync_queue.put(df_novo)
            print("Sincronização em background concluída.")
    except Exception as e:
        print(f"Erro durante a sincronização em background: {e}")

def agendar_sincronizacao_periodica():
    threading.Thread(target=sincronizar_em_background, daemon=True).start()
    root.after(INTERVALO_SYNC_MS, agendar_sincronizacao_periodica)

def verificar_fila_e_atualizar_ui():
    global df
    try:
        df_novo = sync_queue.get(block=False)
        if not df.equals(df_novo):
            print("Novos dados detectados! Atualizando a tabela.")
            df = df_novo
            atualizar_tabela(df)
            coluna_menu['values'] = df.columns.tolist() if not df.empty else COLUNAS_PRINCIPAIS
        else:
            print("Sincronização verificada, sem mudanças.")
    except Exception:
        pass
    root.after(200, verificar_fila_e_atualizar_ui)

df = sincronizar_dados(mostrar_popup=True)

def salvar_registro():
    global df
    novo = {
        "Nome": entry_nome.get(),
        "Nascimento": entry_nascimento.get(),
        "Email": entry_email.get(),
        "Telefone": entry_telefone.get(),
        "CPF": entry_cpf.get(),
        "Profissão": entry_profissao.get(),
        "Carteirinha": entry_carteirinha.get(),
        "Ajuda": entry_ajuda.get(),
        "PCD": entry_pcd.get(),
        "TEA": entry_tea.get(),
        "CEP":entry_cep.get()
    }
    novo["Email"] = novo["Email"].lower().strip()
    if not novo["Email"] or novo["Email"] not in df["Email"].values:
        df = pd.concat([df, pd.DataFrame([novo])], ignore_index=True)
        df.to_csv(CSV_FILE, index=False)
        messagebox.showinfo("Sucesso", "Registro adicionado!")
        atualizar_tabela(df)
    else:
        messagebox.showwarning("Erro", "Um registro com este email já existe.")

frame_cadastro = tk.LabelFrame(root, text="Cadastro")
frame_cadastro.pack(fill="x", padx=9, pady=5)

entry_nome = ttk.Entry(frame_cadastro)
entry_nascimento = ttk.Entry(frame_cadastro)
entry_email = ttk.Entry(frame_cadastro)
entry_telefone = ttk.Entry(frame_cadastro)
entry_cpf = ttk.Entry(frame_cadastro)
entry_profissao = ttk.Entry(frame_cadastro)
entry_carteirinha = ttk.Entry(frame_cadastro)
entry_ajuda = ttk.Entry(frame_cadastro)
entry_pcd = ttk.Entry(frame_cadastro)
entry_tea = ttk.Entry(frame_cadastro)
entry_cep = ttk.Entry(frame_cadastro)

labels = ["Nome", "Nascimento", "Email", "Telefone", "CPF","Profissão","Carteirinha","Ajuda","PCD","TEA","CEP"]
entries = [entry_nome, entry_nascimento, entry_email, entry_telefone, entry_cpf,entry_profissao,entry_carteirinha,entry_ajuda,entry_pcd,entry_tea,entry_cep]

for i, (label, entry) in enumerate(zip(labels, entries)):
    ttk.Label(frame_cadastro, text=label).grid(row=0, column=i, padx=5, pady=5)
    entry.grid(row=1, column=i, padx=5, pady=5)

ttk.Button(frame_cadastro, text="Adicionar", command=salvar_registro).grid(row=1, column=len(entries), padx=5)

def pesquisar():
    coluna = coluna_var.get()
    valor = entrada_valor.get()
    if coluna and valor:
        resultado = df[df[coluna].astype(str).str.contains(valor, case=False, na=False)]
        atualizar_tabela(resultado)
    else:
        atualizar_tabela(df)

def atualizar_tabela(dados):
    dados_display = dados.copy()
    cols_para_exibir = [c for c in COLUNAS_PRINCIPAIS if c in dados_display.columns]
    for c in COLUNAS_PRINCIPAIS:
        if c not in dados_display.columns:
            dados_display[c] = ""
    if "DataEntrada" in dados_display.columns:
        dados_display["DataEntrada"] = dados_display["DataEntrada"].fillna("").astype(str)
    for i in tabela.get_children():
        tabela.delete(i)
    for idx, row in dados_display.iterrows():
        values = [row[col] for col in COLUNAS_PRINCIPAIS if col in dados_display.columns]
        tabela.insert("", "end", iid=idx, values=values)

def excluir():
    selected_items = tabela.selection()
    if not selected_items:
        messagebox.showwarning("Atenção", "Selecione um registro para excluir")
        return
    global df
    try:
        indices_para_excluir = [int(item) for item in selected_items]
        df = df.drop(indices_para_excluir).reset_index(drop=True)
        df.to_csv(CSV_FILE, index=False)
        atualizar_tabela(df)
        messagebox.showinfo("Sucesso", "Registro(s) excluído(s).")
    except Exception as e:
        print(f"Erro ao excluir: {e}")
        messagebox.showerror("Erro", "Não foi possível excluir o registro.")

janela_graf_cart = None
janela_graf_tea = None
janela_graf_crescimento = None
janela_graf_ajuda = None

def abrir_grafico_carteirinha():
    global janela_graf_cart
    if "Carteirinha" not in df.columns or df["Carteirinha"].isnull().all():
        messagebox.showwarning("Sem Dados", "Não há dados suficientes de 'Carteirinha' para gerar o gráfico.")
        return
    if janela_graf_cart is not None:
        try:
            janela_graf_cart.destroy()
        except tk.TclError:
            pass
    contagem = df["Carteirinha"].value_counts().sort_index()
    janela_graf_cart = tk.Toplevel(root)
    janela_graf_cart.title("Gráfico de Carteirinha")
    fig = Figure(figsize=(7, 5), dpi=100)
    ax = fig.add_subplot(111)
    barras = ax.bar(contagem.index, contagem.values, color='skyblue')
    ax.set_title("Pessoas que desejam receber a carteirinha")
    ax.set_xlabel("Pessoas")
    ax.set_ylabel("Quantidade")
    ax.bar_label(barras, fmt='%d')
    canvas = FigureCanvasTkAgg(fig, master=janela_graf_cart)
    canvas.draw()
    canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)

def abrir_grafico_tea():
    global janela_graf_tea
    if "TEA" not in df.columns or df["TEA"].isnull().all():
        messagebox.showwarning("Sem Dados", "Não há dados suficientes de 'TEA' para gerar o gráfico.")
        return
    if janela_graf_tea is not None:
        try:
            janela_graf_tea.destroy()
        except tk.TclError:
            pass
    contagem = df["TEA"].value_counts()
    janela_graf_tea = tk.Toplevel(root)
    janela_graf_tea.title("Gráfico de Pessoas que Trabalham com Público TEA")
    fig = Figure(figsize=(5, 4), dpi=100)
    ax = fig.add_subplot(111)
    ax.pie(contagem, labels=contagem.index, autopct='%1.1f%%', startangle=90)
    ax.axis('equal')
    ax.set_title("Trabalha com Público TEA")
    canvas = FigureCanvasTkAgg(fig, master=janela_graf_tea)
    canvas.draw()
    canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)

def abrir_grafico_ajuda():
    global janela_graf_ajuda
    if "Ajuda" not in df.columns or df["Ajuda"].isnull().all():
        messagebox.showwarning("Sem Dados", "Não há dados suficientes de 'Ajuda' para gerar o gráfico.")
        return
    if janela_graf_ajuda is not None:
        try:
            janela_graf_ajuda.destroy()
        except tk.TclError:
            pass
    contagem = df["Ajuda"].value_counts()
    janela_graf_ajuda = tk.Toplevel(root)
    janela_graf_ajuda.title("Gráfico de Pessoas que Desejam Ajudar")
    fig = Figure(figsize=(5, 4), dpi=100)
    ax = fig.add_subplot(111)
    ax.pie(contagem, labels=contagem.index, autopct='%1.1f%%', startangle=90, colors=['#66b3ff', '#ff9999'])
    ax.axis('equal')
    ax.set_title("Deseja Ajudar")
    canvas = FigureCanvasTkAgg(fig, master=janela_graf_ajuda)
    canvas.draw()
    canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)

def abrir_grafico_crescimento():
    global janela_graf_crescimento
    if janela_graf_crescimento is not None:
        try:
            janela_graf_crescimento.destroy()
        except tk.TclError:
            pass
    if "DataEntrada" not in df.columns:
        if "Carimbo de data/hora" in df.columns:
            df.rename(columns={"Carimbo de data/hora": "DataEntrada"}, inplace=True)
        else:
            df["DataEntrada"] = pd.Timestamp.now()
    df["DataEntrada"] = pd.to_datetime(df["DataEntrada"], errors="coerce")
    if df["DataEntrada"].isnull().all():
        df["DataEntrada"] = pd.Timestamp.now()
    df["Ano"] = df["DataEntrada"].dt.year
    crescimento = df.groupby("Ano").size().reset_index(name="Quantidade")
    ano_inicial = 2024
    ano_final = max(2025, crescimento["Ano"].max())
    anos_completos = pd.DataFrame({"Ano": range(ano_inicial, ano_final + 1)})
    crescimento = pd.merge(anos_completos, crescimento, on="Ano", how="left").fillna(0)
    crescimento["Quantidade"] = crescimento["Quantidade"].astype(int)
    crescimento["Crescimento Acumulado"] = crescimento["Quantidade"].cumsum()
    janela_graf_crescimento = tk.Toplevel(root)
    janela_graf_crescimento.title("Gráfico de Crescimento")
    fig = Figure(figsize=(6, 4), dpi=100)
    ax = fig.add_subplot(111)
    ax.bar(crescimento["Ano"], crescimento["Quantidade"], color='skyblue', alpha=0.5, label="Novos Cadastros")
    ax.plot(crescimento["Ano"], crescimento["Crescimento Acumulado"], marker='o', color='blue', label="Crescimento Acumulado")
    ax.set_title("Crescimento de Membros (2024 em diante)")
    ax.set_xlabel("Ano")
    ax.set_ylabel("Quantidade de Cadastros")
    ax.legend()
    ax.xaxis.set_major_locator(ticker.MultipleLocator(1))
    ax.set_ylim(bottom=0)
    canvas = FigureCanvasTkAgg(fig, master=janela_graf_crescimento)
    canvas.draw()
    canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)

frame_pesquisa = tk.LabelFrame(root, text="Pesquisa")
frame_pesquisa.pack(fill="x", padx=10, pady=5)

frame_graficos = tk.LabelFrame(root, text="Gráficos")
frame_graficos.pack(fill="x", padx=10, pady=5)
ttk.Button(frame_graficos, text="Gráfico da Carteirinha", command=abrir_grafico_carteirinha).pack(side="left", padx=5)
ttk.Button(frame_graficos, text="Gráfico TEA", command=abrir_grafico_tea).pack(side="left", padx=5)
ttk.Button(frame_graficos, text="Gráfico Crescimento", command=abrir_grafico_crescimento).pack(side="left", padx=5)
ttk.Button(frame_graficos, text="Gráfico Ajuda", command=abrir_grafico_ajuda).pack(side="left", padx=5)

coluna_var = tk.StringVar()
coluna_menu = ttk.Combobox(frame_pesquisa, textvariable=coluna_var)
coluna_menu['values'] = df.columns.tolist() if not df.empty else COLUNAS_PRINCIPAIS
coluna_menu.grid(row=0, column=0, padx=5, pady=5)
entrada_valor = ttk.Entry(frame_pesquisa)
entrada_valor.grid(row=0, column=1, padx=5, pady=5)
ttk.Button(frame_pesquisa, text="Pesquisar", command=pesquisar).grid(row=0, column=2, padx=5)

colunas_para_mostrar = df.columns.tolist()
tabela = ttk.Treeview(root, columns=colunas_para_mostrar, show="headings")
for col in colunas_para_mostrar:
    tabela.heading(col, text=col)
    tabela.column(col, width=120, minwidth=80)

tabela.pack(expand=True, fill='both', padx=10, pady=5)
ttk.Button(root, text="Excluir Selecionado", command=excluir).pack(pady=5)

atualizar_tabela(df)

def ajustar_colunas():
    total_colunas = len(COLUNAS_PRINCIPAIS)
    largura_janela = root.winfo_width() - 40
    largura_coluna = int(largura_janela / total_colunas)
    for col in COLUNAS_PRINCIPAIS:
        tabela.column(col, width=largura_coluna, minwidth=80)

atualizar_tabela(df)
ajustar_colunas()

def ao_redimensionar(event):
    ajustar_colunas()

root.bind("<Configure>", ao_redimensionar)

def ao_fechar():
    try:
        if os.path.exists(CSV_FILE):
            os.remove(CSV_FILE)
            print("Arquivo .csv removido com sucesso.")
    except Exception as e:
        print(f"Erro ao apagar o CSV: {e}")
    finally:
        root.destroy()

root.protocol("WM_DELETE_WINDOW", ao_fechar)
root.deiconify()
verificar_fila_e_atualizar_ui()
root.after(INTERVALO_SYNC_MS, agendar_sincronizacao_periodica)
print(f"Sincronização automática agendada para cada {INTERVALO_SYNC_MS / 1000 / 60} minutos.")
root.mainloop()
