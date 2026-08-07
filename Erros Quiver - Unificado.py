import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import datetime
import threading
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

# ------------------------------
# Utilidades de UI (thread-safe)
# ------------------------------

def set_progress_safe(root, canvas, value):
    try:
        root.after(0, lambda: desenhar_foguete(canvas, value))
    except Exception:
        pass

def show_info_safe(root, title, msg):
    root.after(0, lambda: messagebox.showinfo(title, msg))

def show_error_safe(root, title, msg):
    root.after(0, lambda: messagebox.showerror(title, msg))

def update_status_safe(root, status_label, msg, color="#A6ACCD"):
    try:
        root.after(0, lambda: status_label.configure(text=msg, foreground=color))
    except Exception:
        pass

# ------------------------------
# Animação do Foguete (Loading)
# ------------------------------
def desenhar_foguete(canvas, percent):
    canvas.delete("all")
    w = canvas.winfo_width()
    h = canvas.winfo_height()
    if w <= 1: w = 400
    if h <= 1: h = 40
    
    # Pista de fundo escurecida
    canvas.create_rectangle(15, h/2 - 4, w - 15, h/2 + 4, fill="#1A1B26", outline="", width=0)
    
    # Calculando a posição
    x_pos = 15 + (percent / 100.0) * (w - 50)
    
    # Rastro do Foguete (Gás / Fogo)
    if percent > 0:
        canvas.create_rectangle(15, h/2 - 4, x_pos, h/2 + 4, fill="#F39C12", outline="")
        
    # O Foguete!
    canvas.create_text(x_pos + 15, h/2, text="🚀", font=("Segoe UI", 16))

# -------------------------------------------------------------
# Formatação Premium do Excel
# -------------------------------------------------------------
def formatar_planilha_saida(caminho):
    try:
        wb = load_workbook(caminho)
        ws = wb.active
        
        # Cores e Fontes do Cabeçalho
        fill_header = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
        font_header = Font(color="FFFFFF", bold=True)
        align_center = Alignment(horizontal="center", vertical="center")
        
        # Aplicar no cabeçalho (Primeira linha)
        for cell in ws[1]:
            cell.fill = fill_header
            cell.font = font_header
            cell.alignment = align_center
            
        # Adicionar Filtro na primeira linha
        ws.auto_filter.ref = ws.dimensions
        
        # Congelar a primeira linha
        ws.freeze_panes = "A2"
        
        # Auto-ajuste da largura das colunas
        for col in ws.columns:
            max_length = 0
            column_letter = col[0].column_letter
            for cell in col:
                try:
                    if cell.value:
                        # Limita a largura máxima para não ficar bizarro em textos muito longos
                        max_length = min(max(max_length, len(str(cell.value))), 60)
                except:
                    pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column_letter].width = adjusted_width
            
        wb.save(caminho)
    except Exception as e:
        print(f"Erro ao formatar Excel: {e}")

# -------------------------------------------------------------
# 1) Garantias para Locação 
# -------------------------------------------------------------
def gpl_comparar_ponto_venda_vendedor(row):
    ponto_venda = str(row.get('PTO_NOME', '')).strip()
    vendedor = str(row.get('VENDEDOR', '')).strip()
    ignore_list = ['Jetimob', 'Renata guedes', 'Comercial totumseg', 'matriz', 'marco zero empreendimentos']
    if ponto_venda in ignore_list or vendedor in ignore_list: return ''
    if ponto_venda.lower().startswith(vendedor.lower()) or vendedor.lower().startswith(ponto_venda.lower()): return ''
    return f"Erro: Ponto de venda {ponto_venda} não coincide com vendedor {vendedor}"

def gpl_verificar_ausencia_vendedor(row):
    if pd.isna(row.get('VENDEDOR')) or str(row.get('VENDEDOR', '')).strip() == '': return "Erro: Vendedor ausente"
    return ''

def gpl_verificar_vigencia_zerada(row):
    try:
        if pd.to_numeric(row.get('FINAL_VIGENCIA'), errors='coerce') - pd.to_numeric(row.get('INICIO_VIGENCIA'), errors='coerce') == 0:
            return "Erro: Vigência zerada"
    except Exception as e: return f"Erro na vigência: {str(e)}"
    return ''

def pipeline_garantias_locacao(df, root, canvas_foguete):
    total = len(df)
    for i, row in df.iterrows():
        erro = ' | '.join(filter(None, [
            gpl_comparar_ponto_venda_vendedor(row),
            gpl_verificar_ausencia_vendedor(row),
            gpl_verificar_vigencia_zerada(row),
        ]))
        df.at[i, 'Erro'] = erro if erro else "Erros não localizados"
        set_progress_safe(root, canvas_foguete, ((i + 1) / max(total, 1)) * 100)
    return df

# -------------------------------------------------------------
# 2) Seguros Imobiliários
# -------------------------------------------------------------
def pipeline_seguros_imobiliarios(df, root, canvas_foguete):
    total = len(df)
    for i, row in df.iterrows():
        erro = ' | '.join(filter(None, [
            gpl_verificar_ausencia_vendedor(row), # Reaproveitando regra
            gpl_verificar_vigencia_zerada(row),
        ]))
        df.at[i, 'Erro'] = erro if erro else "Erros não localizados"
        set_progress_safe(root, canvas_foguete, ((i + 1) / max(total, 1)) * 100)
    return df

# -------------------------------------------------------------
# 3) Faturamento Educacional
# -------------------------------------------------------------
def pipeline_faturamento_educacional(df, root, canvas_foguete):
    total = len(df)
    for i, row in df.iterrows():
        df.at[i, 'Error'] = "Regras processadas..."
        set_progress_safe(root, canvas_foguete, ((i + 1) / max(total, 1)) * 100)
    return df

# -------------------------------------------------------------
# 4) Sinistro Educacional
# -------------------------------------------------------------
def pipeline_sinistro_educacional(df, root, canvas_foguete):
    total = len(df)
    for i, row in df.iterrows():
        df.at[i, 'Error'] = "Regras processadas..."
        set_progress_safe(root, canvas_foguete, ((i + 1) / max(total, 1)) * 100)
    return df

# -------------------------------------------------------------
# Orquestração
# -------------------------------------------------------------
CATEGORIAS = [
    'Garantias para Locação (Não utilizado)',
    'Seguros Imobiliários',
    'Faturamento Educacional',
    'Sinistro Educacional',
]

def processar_arquivo(root, canvas_foguete, status_label, categoria, input_path, output_path):
    try:
        update_status_safe(root, status_label, "Lendo arquivo Excel...", "#F1C40F")
        df = pd.read_excel(input_path)

        update_status_safe(root, status_label, "Decolando! Analisando regras...", "#3498DB")
        
        if categoria == 'Garantias para Locação (Não utilizado)':
            df_out = pipeline_garantias_locacao(df, root, canvas_foguete)
        elif categoria == 'Seguros Imobiliários':
            df_out = pipeline_seguros_imobiliarios(df, root, canvas_foguete)
        elif categoria == 'Faturamento Educacional':
            df_out = pipeline_faturamento_educacional(df, root, canvas_foguete)
        elif categoria == 'Sinistro Educacional':
            df_out = pipeline_sinistro_educacional(df, root, canvas_foguete)
        else:
            raise ValueError('Categoria inválida')

        update_status_safe(root, status_label, "Gerando planilha base...", "#F1C40F")
        df_out.to_excel(output_path, index=False)
        
        update_status_safe(root, status_label, "Formatando Layout VIP no Excel...", "#9B59B6")
        formatar_planilha_saida(output_path)
        
        update_status_safe(root, status_label, "✔ Pronto! Missão concluída.", "#2ECC71")
        show_info_safe(root, 'Sucesso', 'Arquivo gerado e formatado com sucesso!')
        
    except Exception as e:
        update_status_safe(root, status_label, "✖ O foguete caiu (Erro)", "#E74C3C")
        show_error_safe(root, 'Erro', f'Ocorreu um erro: {e}')

def escolher_e_executar(root, combo, canvas_foguete, status_label):
    categoria = combo.get().strip()
    if not categoria:
        messagebox.showwarning('Aviso', 'Selecione uma categoria primeiro!')
        return

    input_path = filedialog.askopenfilename(title='Selecione o arquivo Excel', filetypes=[('Excel Files', '*.xlsx;*.xls')])
    if not input_path:
        return

    output_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel Files', '*.xlsx')],
                                               title='Salvar Planilha VIP')
    if not output_path:
        return

    desenhar_foguete(canvas_foguete, 0)
    update_status_safe(root, status_label, "Iniciando motores...", "#3498DB")

    t = threading.Thread(target=processar_arquivo, args=(root, canvas_foguete, status_label, categoria, input_path, output_path), daemon=True)
    t.start()

# -------------------------------------------------------------
# GUI PRINCIPAL VIP - MACIA E ELEGANTE
# -------------------------------------------------------------
def criar_interface():
    root = tk.Tk()
    root.title('Analisador VIP - Quiver')
    root.geometry('550x380')  # Tamanho menor e mais elegante
    
    # Cores suaves e profissionais (Catppuccin Macchiato inspired)
    bg_main = "#24273A"      # Fundo principal suave
    bg_card = "#363A4F"      # Fundo do card
    text_color = "#CAD3F5"   # Texto claro azulado
    accent = "#8AADF4"       # Azul pastel para botões
    
    root.configure(bg=bg_main)
    style = ttk.Style(root)
    style.theme_use("clam")

    # Fontes mais delicadas
    fonte_titulo = ("Segoe UI", 14, "bold")
    fonte_sub = ("Segoe UI", 10)

    # Configurando o tema
    style.configure("Card.TFrame", background=bg_card)
    style.configure("Title.TLabel", background=bg_main, font=fonte_titulo, foreground=accent)
    style.configure("Texto.TLabel", background=bg_card, font=fonte_sub, foreground=text_color)
    style.configure("Status.TLabel", background=bg_card, font=("Segoe UI", 10, "italic"), foreground="#A6ACCD")
    
    # Arrumando o horror do Combobox branco
    style.configure("TCombobox", 
                    fieldbackground="#24273A", 
                    background="#494D64", 
                    foreground="#FFFFFF", 
                    arrowcolor="#FFFFFF",
                    bordercolor=bg_card,
                    lightcolor=bg_card,
                    darkcolor=bg_card)

    # Botão mais elegante
    style.configure("Action.TButton", font=("Segoe UI", 10, "bold"), background=accent, foreground="#24273A", padding=8, borderwidth=0)
    style.map("Action.TButton", background=[("active", "#B7BDF8")])

    # Layout Principal
    lbl_titulo = ttk.Label(root, text="🛸 Analisador Unificado VIP", style="Title.TLabel")
    lbl_titulo.pack(pady=(20, 10))

    card = ttk.Frame(root, style="Card.TFrame")
    card.pack(fill='both', expand=True, padx=30, pady=(0, 25))

    # Conteúdo do Card
    ttk.Label(card, text="Módulo de Diagnóstico:", style="Texto.TLabel").pack(anchor='w', padx=20, pady=(20, 5))
    
    combo = ttk.Combobox(card, values=CATEGORIAS, state='readonly', style="TCombobox")
    combo.pack(fill='x', padx=20, pady=(0, 15))
    if CATEGORIAS: combo.current(1)

    ttk.Label(card, text="Progresso:", style="Texto.TLabel").pack(anchor='w', padx=20, pady=(5, 5))
    
    # O CANVAS DO FOGUETE
    canvas_foguete = tk.Canvas(card, height=40, bg=bg_card, highlightthickness=0)
    canvas_foguete.pack(fill='x', padx=20, pady=(0, 5))
    desenhar_foguete(canvas_foguete, 0) # Desenha inicial (vazio)

    status_label = ttk.Label(card, text="Aguardando inicialização...", style="Status.TLabel")
    status_label.pack(anchor='w', padx=20, pady=(0, 15))

    btn = ttk.Button(card, text='SELECIONAR ARQUIVO E EXECUTAR', style="Action.TButton",
                     command=lambda: escolher_e_executar(root, combo, canvas_foguete, status_label))
    btn.pack(fill='x', padx=20, pady=(0, 20))

    root.mainloop()

if __name__ == '__main__':
    criar_interface()
