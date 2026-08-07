import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import datetime
import threading

# ------------------------------
# Utilidades de UI (thread-safe)
# ------------------------------

def set_progress_safe(root, progress_bar, value):
    try:
        root.after(0, lambda: progress_bar.configure(value=value))
    except Exception:
        pass

def set_max_progress_safe(root, progress_bar, maximum):
    try:
        root.after(0, lambda: progress_bar.configure(maximum=maximum))
    except Exception:
        pass

def show_info_safe(root, title, msg):
    root.after(0, lambda: messagebox.showinfo(title, msg))

def show_error_safe(root, title, msg):
    root.after(0, lambda: messagebox.showerror(title, msg))

def update_status_safe(root, status_label, msg, color="#00E5FF"):
    try:
        root.after(0, lambda: status_label.configure(text=msg, foreground=color))
    except Exception:
        pass

# -------------------------------------------------------------
# 1) Garantias para Locação — funções preservadas
# -------------------------------------------------------------

def gpl_comparar_ponto_venda_vendedor(row):
    ponto_venda = str(row.get('PTO_NOME', '')).strip()
    vendedor = str(row.get('VENDEDOR', '')).strip()
    ignore_list = ['Jetimob', 'Renata guedes', 'Comercial totumseg', 'matriz', 'marco zero empreendimentos']
    if ponto_venda in ignore_list or vendedor in ignore_list:
        return ''
    if ponto_venda.lower().startswith(vendedor.lower()) or vendedor.lower().startswith(ponto_venda.lower()):
        return ''
    return f"Erro: Ponto de venda {ponto_venda} não coincide com vendedor {vendedor}"

def gpl_verificar_ausencia_vendedor(row):
    if pd.isna(row.get('VENDEDOR')) or str(row.get('VENDEDOR', '')).strip() == '':
        return "Erro: Vendedor ausente"
    return ''

def gpl_verificar_vigencia_zerada(row):
    try:
        if pd.to_numeric(row.get('FINAL_VIGENCIA'), errors='coerce') - pd.to_numeric(row.get('INICIO_VIGENCIA'), errors='coerce') == 0:
            return "Erro: Vigência zerada"
    except Exception as e:
        return f"Erro na vigência: {str(e)}"
    return ''

def gpl_verificar_produto_divergente(row):
    seguradora = str(row.get('SEGURADORA', '')).strip()
    produto = str(row.get('PRODUTO', '')).strip()
    produtos_aceitos = {
        'BrasilCap': ['Título de capitalização'],
        'MAPFRECAP': ['FIANCA LOCATICIA TAXA FIXA MAPFRE', 'FIANCA LOCATICIA MAPFRE CONNECT'],
        'Porto Capitalização': ['FIANCA LOCATICIA', 'FIANCA PORTO ESSENCIAL'],
        'ICATU': ['Título de capitalização'],
        'SUL AMERICA': ['Título de capitalização'],
        'Liberty': ['Fiança Locatícia', 'Fiança Liberty Online'],
        'Yelum': ['Fiança Locatícia', 'Fiança Liberty Online'],
    }
    if seguradora in produtos_aceitos and produto not in produtos_aceitos[seguradora]:
        return f"Erro: Produto {produto} divergente para a seguradora {seguradora}"
    return ''

def gpl_verificar_proposta_cia_vazia(row):
    if pd.isna(row.get('PROPOSTA_CIA')) and row.get('TIPO_DOCUMENTO') != 'cancelamento' and row.get('QTDE_SINISTROS_DOCUMENTO', 0) != 0:
        return "Erro: Proposta_CIA vazia"
    return ''

def gpl_verificar_numero_arquivo_vazio(row):
    if pd.isna(row.get('NUMERO_ARQUIVO')) or str(row.get('NUMERO_ARQUIVO', '')).strip() == '':
        return "Erro: Número de arquivo vazio"
    return ''

def gpl_verificar_data_envio_post_emissao(row):
    try:
        formato = "%d/%m/%Y"
        data_envio = pd.to_datetime(row.get('DATA_ENVIO_PROPOSTA'), format=formato, errors='coerce')
        data_emissao = pd.to_datetime(row.get('DATA_EMISSAO'), format=formato, errors='coerce')
        if pd.notna(data_envio) and pd.notna(data_emissao) and data_envio.year >= 2023 and data_envio > data_emissao:
            return "Erro: Data de envio posterior à data de emissão"
    except Exception as e:
        return f"Erro na data: {str(e)}"
    return ''

def gpl_verificar_placa_matricula(row):
    placa_valida = ['PF RESIDENCIAL', 'PF NÃO RESIDENCIAL', 'PJ RESIDENCIAL', 'PJ NÃO RESIDENCIAL']
    val = str(row.get('PLACA_MATRICULA', '')).strip()
    if val == '':
        return "Erro: Placa da matrícula vazia"
    if val not in placa_valida:
        return f"Erro: Placa da matrícula {val} diferente do padrão"
    return ''

def gpl_verificar_ocupacao_vazia(row):
    if pd.isna(row.get('OCUPACAO')) or str(row.get('OCUPACAO', '')).strip() == '':
        return "Erro: Ocupação vazia"
    return ''

def gpl_verificar_descricao_bem_vazia(row):
    if pd.isna(row.get('DESCRICAO_BEM')) or str(row.get('DESCRICAO_BEM', '')).strip() == '':
        return "Erro: Descrição do bem vazia"
    return ''

def gpl_verificar_apolice_asterisco(row):
    if '*' in str(row.get('APOLICE', '')):
        return ''
    return ''

def pipeline_garantias_locacao(df, root, progress_bar):
    total = len(df)
    set_max_progress_safe(root, progress_bar, 100)
    for i, row in df.iterrows():
        erro = ' | '.join(filter(None, [
            gpl_comparar_ponto_venda_vendedor(row),
            gpl_verificar_ausencia_vendedor(row),
            gpl_verificar_vigencia_zerada(row),
            gpl_verificar_produto_divergente(row),
            gpl_verificar_proposta_cia_vazia(row),
            gpl_verificar_numero_arquivo_vazio(row),
            gpl_verificar_data_envio_post_emissao(row),
            gpl_verificar_placa_matricula(row),
            gpl_verificar_ocupacao_vazia(row),
            gpl_verificar_descricao_bem_vazia(row),
            gpl_verificar_apolice_asterisco(row),
        ]))
        df.at[i, 'Erro'] = erro
        set_progress_safe(root, progress_bar, (i + 1) / max(total, 1) * 100)

    colunas_ordem = [
        'PTO_NOME', 'VENDEDOR', 'APOLICE', 'CLIENTE', 'INICIO_VIGENCIA', 'FINAL_VIGENCIA', 'SEGURADORA',
        'PRODUTO', 'PROPOSTA_CIA', 'NUMERO_ARQUIVO', 'ENDOSSO', 'TIPO_MOVIMENTO', 'TIPO_DOCUMENTO',
        'PREMIO_PARCELAS', 'PREMIO_TOTAL', 'QTDE_SINISTROS_DOCUMENTO', 'SITUACAO_APOLICE', 'DATA_EMISSAO',
        'QTD_PARCELAS', 'DATA_ENVIO_PROPOSTA', 'DATA_ENTRADA', 'DATA_PROPOSTA', 'VALOR_REPASSES',
        'PERCENTUAL_REPASSE', 'PLACA_MATRICULA', 'OCUPACAO', 'DESCRICAO_BEM', 'PERC_COMISSAO',
        'AUD_INCLUSAO_USR', 'AUD_ALTERACAO_USR', 'AUD_INCLUSAO_DATA', 'Erro'
    ]
    existentes = [c for c in colunas_ordem if c in df.columns]
    return df[existentes]

# -------------------------------------------------------------
# 2) Seguros Imobiliários — funções preservadas
# -------------------------------------------------------------

def sim_comparar_ponto_venda_vendedor(row):
    ponto_venda = str(row.get('PTO_NOME', '')).strip()
    vendedor = str(row.get('VENDEDOR', '')).strip()
    ignore_list = ['Jetimob', 'Renata guedes', 'Comercial totumseg', 'matriz', 'marco zero empreendimentos',
                   'JETIMOB - ORUS TECNOLOGIA LTDA - ME']
    if ponto_venda in ignore_list or vendedor in ignore_list:
        return ''
    if vendedor == "Crédito Real" and (ponto_venda in ["Pontal", "SP ADM E LOCADORA"]):
        return ''
    if vendedor == "DEMBOSKI" and ponto_venda == "ANAGE":
        return ''
    if vendedor == "Jetimob" and ponto_venda:
        return ''
    if ponto_venda.lower().startswith(vendedor.lower()) or vendedor.lower().startswith(ponto_venda.lower()):
        return ''
    return f"Erro: Ponto de venda {ponto_venda} não coincide com vendedor {vendedor}"

def sim_verificar_ausencia_vendedor(row):
    if pd.isna(row.get('VENDEDOR')) or str(row.get('VENDEDOR', '')).strip() == '':
        return "Erro: Vendedor ausente"
    return ''

def sim_verificar_vigencia_zerada(row):
    try:
        if pd.to_numeric(row.get('FINAL_VIGENCIA'), errors='coerce') - pd.to_numeric(row.get('INICIO_VIGENCIA'), errors='coerce') == 0:
            return "Erro: Vigência zerada"
    except Exception as e:
        return f"Erro na vigência: {str(e)}"
    return ''

def sim_verificar_proposta_cia_vazia(row):
    produtos_permitidos = ['VIDA', 'PRÓ-TRABALHO', 'VIDA INDIVIDUAL (PREMIO TOTAL MENSAL)', 'VIDA EM GRUPO', 'VIDA INDIVIDUAL (PREMIO INTEGRAL)']
    if pd.isna(row.get('PROPOSTA_CIA')) and row.get('PRODUTO') not in produtos_permitidos:
        return "Erro: Proposta_CIA vazia"
    return ''

def sim_verificar_numero_arquivo_vazio(row):
    if pd.isna(row.get('NUMERO_ARQUIVO')) or str(row.get('NUMERO_ARQUIVO', '')).strip() == '':
        return "Erro: Número de arquivo vazio"
    return ''

def sim_verificar_data_envio_post_emissao(row):
    try:
        formato = "%d/%m/%Y"
        data_envio = pd.to_datetime(row.get('DATA_ENVIO_PROPOSTA'), format=formato, errors='coerce')
        data_emissao = pd.to_datetime(row.get('DATA_EMISSAO'), format=formato, errors='coerce')
        if pd.notna(data_envio) and pd.notna(data_emissao) and data_envio.year >= 2023 and data_envio > data_emissao:
            return "Erro: Data de envio posterior à data de emissão"
    except Exception as e:
        return f"Erro na data: {str(e)}"
    return ''

def sim_verificar_apolice_asterisco(row):
    if '*' in str(row.get('APOLICE', '')):
        return ''
    return ''

def pipeline_seguros_imobiliarios(df, root, progress_bar):
    total = len(df)
    set_max_progress_safe(root, progress_bar, 100)
    for i, row in df.iterrows():
        erro = ' | '.join(filter(None, [
            sim_verificar_apolice_asterisco(row),
            sim_comparar_ponto_venda_vendedor(row),
            sim_verificar_ausencia_vendedor(row),
            sim_verificar_vigencia_zerada(row),
            sim_verificar_proposta_cia_vazia(row),
            sim_verificar_numero_arquivo_vazio(row),
            sim_verificar_data_envio_post_emissao(row),
        ]))
        if not erro:
            erro = "Erros não localizados"
        df.at[i, 'Erro'] = erro
        set_progress_safe(root, progress_bar, (i + 1) / max(total, 1) * 100)

    colunas_ordem = [
        'PTO_NOME', 'VENDEDOR', 'APOLICE', 'CLIENTE', 'INICIO_VIGENCIA', 'FINAL_VIGENCIA', 'SEGURADORA',
        'PRODUTO', 'PROPOSTA_CIA', 'NUMERO_ARQUIVO', 'ENDOSSO', 'TIPO_MOVIMENTO', 'TIPO_DOCUMENTO',
        'PREMIO_PARCELAS', 'PREMIO_TOTAL', 'QTDE_SINISTROS_DOCUMENTO', 'SITUACAO_APOLICE', 'DATA_EMISSAO',
        'QTD_PARCELAS', 'DATA_ENVIO_PROPOSTA', 'DATA_ENTRADA', 'DATA_PROPOSTA', 'VALOR_REPASSES',
        'PERCENTUAL_REPASSE', 'PERC_COMISSAO', 'AUD_INCLUSAO_USR', 'AUD_ALTERACAO_USR', 'AUD_INCLUSAO_DATA', 'Erro'
    ]
    existentes = [c for c in colunas_ordem if c in df.columns]
    return df[existentes]

# -------------------------------------------------------------
# 3) Faturamento Educacional — funções preservadas
# -------------------------------------------------------------

def fe_verificar_vigencia_zerada(row):
    try:
        if pd.to_numeric(row.get('FINAL_VIGENCIA'), errors='coerce') - pd.to_numeric(row.get('INICIO_VIGENCIA'), errors='coerce') == 0:
            if row.get('TIPO_MOVIMENTO') in ['Cancelamento', 'Fatura Complementar', 'Não Emitida', 'Informativo']:
                return ''
            return "Erro: Vigência zerada"
    except Exception as e:
        return f"Erro na vigência: {str(e)}"
    return ''

def fe_verificar_seguradora(row):
    seguradora = str(row.get('SEGURADORA', '')).strip()
    if seguradora not in ["MAPFRE SEG GERAIS", "MAPFRE VIDA S.A."]:
        return "Erro: Seguradora inválida"
    return ''

def fe_verificar_produto(row):
    produto = str(row.get('PRODUTO', '')).strip()
    if produto not in ['2172 MAPFRE PROTEÇÃO ESCOLAR MULTIFLEX', '2173 MAPFRE PROTEÇÃO ESCOLAR MULTIFLEX', '2172 MAPFRE PROTEÇÃO EDUC MULTIFLEX']:
        return "Erro: Produto inválido"
    return ''

def fe_verificar_numero_arquivo(row):
    if pd.isna(row.get('NUMERO_ARQUIVO')) or str(row.get('NUMERO_ARQUIVO', '')).strip() == '':
        if pd.isna(row.get('PROPOSTA_CIA')) or str(row.get('PROPOSTA_CIA', '')).strip() == '':
            return "Erro: Número de arquivo vazio sem justificativa"
    return ''

def fe_verificar_endosso(row):
    if pd.isna(row.get('ENDOSSO')) or str(row.get('ENDOSSO', '')).strip() == '':
        if row.get('TIPO_MOVIMENTO') in ['Fatura', 'END. SEM. MOVIMENTO']:
            data_atual = datetime.datetime.now()
            try:
                data_emissao = pd.to_datetime(row.get('DATA_EMISSAO'), format="%d/%m/%Y", errors='coerce')
            except Exception:
                data_emissao = pd.NaT
            if pd.notna(data_emissao) and data_emissao > data_atual - datetime.timedelta(days=30):
                return ''
            if pd.isna(row.get('PROPOSTA_CIA')) or str(row.get('PROPOSTA_CIA', '')).strip() == '':
                return "Erro: ENDOSSO vazio sem justificativa"
    endosso = str(row.get('ENDOSSO', '')).strip().lower()
    if 'cancelamento' in endosso and 'cancela' not in endosso:
        return "Erro: Problema de padronização ou erro de português no ENDOSSO"
    return ''

def fe_verificar_premio_apolice_cancelamento(row):
    if row.get('TIPO_MOVIMENTO') in ['Apólice', 'Cancelamento']:
        if pd.notna(row.get('PREMIO_PARCELAS')) or pd.notna(row.get('PREMIO_TOTAL')):
            return "Erro: PRÊMIO_PARCELAS ou PRÊMIO_TOTAL não deve ser preenchido para apólice ou cancelamento"
    return ''

def fe_verificar_data_envio_proposta(row):
    if row.get('TIPO_DOCUMENTO') == 'Fatura':
        try:
            data_envio = pd.to_datetime(row.get('DATA_ENVIO_PROPOSTA'), format="%d/%m/%Y", errors='coerce')
            data_emissao = pd.to_datetime(row.get('DATA_EMISSAO'), format="%d/%m/%Y", errors='coerce')
            if pd.notna(data_envio) and pd.notna(data_emissao) and data_envio > data_emissao:
                return "Erro: DATA_ENVIO_PROPOSTA posterior a DATA_EMISSAO para fatura"
        except Exception as e:
            return f"Erro na data: {str(e)}"
    return ''

def pipeline_faturamento_educacional(df, root, progress_bar):
    total = len(df)
    set_max_progress_safe(root, progress_bar, 100)
    for i, row in df.iterrows():
        erro = ' | '.join(filter(None, [
            fe_verificar_vigencia_zerada(row),
            fe_verificar_seguradora(row),
            fe_verificar_produto(row),
            fe_verificar_numero_arquivo(row),
            fe_verificar_endosso(row),
            fe_verificar_premio_apolice_cancelamento(row),
            fe_verificar_data_envio_proposta(row),
        ]))
        if not erro:
            erro = 'Erros não localizados'
        df.at[i, 'Error'] = erro
        set_progress_safe(root, progress_bar, (i + 1) / max(total, 1) * 100)

    colunas_ordem = [
        'PTO_NOME', 'APOLICE', 'CLIENTE', 'INICIO_VIGENCIA', 'FINAL_VIGENCIA', 'SEGURADORA', 'PRODUTO',
        'PROPOSTA_CIA', 'NUMERO_ARQUIVO', 'ENDOSSO', 'TIPO_MOVIMENTO', 'TIPO_DOCUMENTO', 'PREMIO_PARCELAS',
        'PREMIO_TOTAL', 'QTDE_SINISTROS_DOCUMENTO', 'AUD_INCLUSAO_DATA', 'SITUACAO_APOLICE', 'DATA_EMISSAO', 'QTD_PARCELAS',
        'DATA_ENVIO_PROPOSTA', 'DATA_ENTRADA', 'DATA_PROPOSTA', 'AUD_INCLUSAO_USR','AUD_ALTERACAO_USR',
        'VALOR_REPASSES', 'PERC_COMISSAO', 'Error'
    ]
    existentes = [c for c in colunas_ordem if c in df.columns]
    return df[existentes]

# -------------------------------------------------------------
# 4) Sinistro Educacional — funções preservadas
# -------------------------------------------------------------

def sin_converte_data(data):
    try:
        return pd.to_datetime(data, format='%d/%m/%Y', errors='coerce')
    except Exception:
        return pd.NaT

def pipeline_sinistro_educacional(df, root, progress_bar):
    valid_columns = ['CLIENTE', 'CPF_CNPJ', 'SIN_NUMERO', 'SINSIT_DESCRICAO', 'BEM_SINISTRADO', 'CODIGO_CONTROLE',
                     'TIPO_SINISTRO', 'APOLICE', 'VALOR_RECLAMADO', 'VALOR_INDENIZADO', 'VALOR_LIBERADO',
                     'DATA_SINISTRO', 'DATA_AVISO', 'DATA_VISTORIA', 'DATA_LIBERACAO', 'DATA_LIQUIDACAO', 'LOCAL_SINISTRO']
    faltantes = [c for c in valid_columns if c not in df.columns]
    if faltantes:
        raise KeyError(f"Colunas faltantes na planilha de Sinistros: {faltantes}")

    df = df[valid_columns].copy()
    total = len(df)
    set_max_progress_safe(root, progress_bar, total if total > 0 else 1)

    for i, row in df.iterrows():
        error_message = []
        codigo = row.get('CODIGO_CONTROLE')
        if pd.isna(codigo) or codigo == '':
            error_message.append("Código de controle vazio")
        else:
            sinistro_desc = row.get('SINSIT_DESCRICAO')
            if sinistro_desc in ['LIQUIDADO', 'ANUIDADE EM LIQUIDAÇÃO', 'PERDA DE RENDA EM LIQUIDACAO']:
                if 'P' not in str(codigo) and 'T' not in str(codigo):
                    if 'TRONADOR' not in str(row.get('LOCAL_SINISTRO', '')).upper():
                        error_message.append("Código de controle sem P ou T e não é TRONADOR")

        if pd.isna(row.get('VALOR_RECLAMADO')) or row.get('VALOR_RECLAMADO') == 0:
            if pd.isna(row.get('LOCAL_SINISTRO')) or row.get('LOCAL_SINISTRO') == '':
                error_message.append("Sinistro sem valor reclamado e sem LOCAL_SINISTRO preenchido")

        try:
            data_sinistro = sin_converte_data(row.get('DATA_SINISTRO'))
            data_aviso = sin_converte_data(row.get('DATA_AVISO'))
            data_vistoria = sin_converte_data(row.get('DATA_VISTORIA'))
            data_liberacao = sin_converte_data(row.get('DATA_LIBERACAO'))
            data_liquidacao = sin_converte_data(row.get('DATA_LIQUIDACAO'))

            if pd.notna(data_sinistro):
                if (pd.notna(data_aviso) and data_sinistro > data_aviso) or \
                   (pd.notna(data_vistoria) and data_sinistro > data_vistoria) or \
                   (pd.notna(data_liberacao) and data_sinistro > data_liberacao) or \
                   (pd.notna(data_liquidacao) and data_sinistro > data_liquidacao):
                    error_message.append("Data do sinistro deve ser anterior ou igual a todas as outras datas")

            if pd.notna(data_aviso):
                if (pd.notna(data_vistoria) and data_aviso > data_vistoria) or \
                   (pd.notna(data_liberacao) and data_aviso > data_liberacao) or \
                   (pd.notna(data_liquidacao) and data_aviso > data_liquidacao):
                    error_message.append("Data do aviso deve ser anterior ou igual a data da vistoria, liberação e liquidação")

            if pd.notna(data_vistoria):
                if (pd.notna(data_liberacao) and data_vistoria > data_liberacao) or \
                   (pd.notna(data_liquidacao) and data_vistoria > data_liquidacao):
                    error_message.append("Data da vistoria deve ser anterior ou igual a data da liberação e da liquidação")

            if pd.notna(data_liberacao) and pd.notna(data_liquidacao) and data_liberacao > data_liquidacao:
                error_message.append("Data da liberação deve ser anterior ou igual a data da liquidação")
        except Exception as e:
            error_message.append(f"Erro ao comparar as datas: {e}")

        df.at[i, 'Error'] = ', '.join(error_message) if error_message else 'Erros não localizados'
        set_progress_safe(root, progress_bar, i + 1)

    return df

# -------------------------------------------------------------
# Orquestração
# -------------------------------------------------------------

# Texto alterado conforme solicitado
CATEGORIAS = [
    'Garantias para Locação (Não utilizado)',
    'Seguros Imobiliários',
    'Faturamento Educacional',
    'Sinistro Educacional',
]

def processar_arquivo(root, progress_bar, status_label, categoria, input_path, output_path):
    try:
        update_status_safe(root, status_label, "Lendo arquivo Excel...", "#FFC107")
        df = pd.read_excel(input_path)

        update_status_safe(root, status_label, "Analisando dados (Aguarde)...", "#00E5FF")
        
        # Mapeamento atualizado para refletir a nova string
        if categoria == 'Garantias para Locação (Não utilizado)':
            df_out = pipeline_garantias_locacao(df, root, progress_bar)
        elif categoria == 'Seguros Imobiliários':
            df_out = pipeline_seguros_imobiliarios(df, root, progress_bar)
        elif categoria == 'Faturamento Educacional':
            df_out = pipeline_faturamento_educacional(df, root, progress_bar)
        elif categoria == 'Sinistro Educacional':
            df_out = pipeline_sinistro_educacional(df, root, progress_bar)
        else:
            raise ValueError('Categoria inválida')

        update_status_safe(root, status_label, "Gerando arquivo de saída...", "#FFC107")
        df_out.to_excel(output_path, index=False)
        
        update_status_safe(root, status_label, "✔ Concluído com Sucesso!", "#00E676")
        show_info_safe(root, 'Sucesso', 'Arquivo com erros salvo com sucesso!')
    except Exception as e:
        update_status_safe(root, status_label, "✖ Erro no processamento", "#FF1744")
        show_error_safe(root, 'Erro', f'Ocorreu um erro: {e}')

def escolher_e_executar(root, combo, progress_bar, status_label):
    categoria = combo.get().strip()
    if not categoria:
        messagebox.showwarning('Aviso', 'Selecione uma categoria primeiro!')
        return

    input_path = filedialog.askopenfilename(title='Selecione o arquivo Excel', filetypes=[('Excel Files', '*.xlsx;*.xls')])
    if not input_path:
        return

    output_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel Files', '*.xlsx')],
                                               title='Salvar Planilha Com Erros')
    if not output_path:
        return

    progress_bar.configure(value=0)
    update_status_safe(root, status_label, "Iniciando processamento...", "#00E5FF")

    t = threading.Thread(target=processar_arquivo, args=(root, progress_bar, status_label, categoria, input_path, output_path), daemon=True)
    t.start()

# -------------------------------------------------------------
# GUI PRINCIPAL VIP - DARK PREMIUM
# -------------------------------------------------------------

def criar_interface():
    root = tk.Tk()
    root.title('Analisador de Erros Quiver - PRO EDITION')
    root.geometry('650x450')
    root.configure(bg="#0F0F13") # Fundo escuro profundo

    style = ttk.Style(root)
    style.theme_use("clam")

    # Cores Premium VIP
    bg_color = "#0F0F13"
    card_color = "#1C1C22"
    text_color = "#FFFFFF"
    accent_color = "#00E5FF" # Cyan vibrante
    accent_hover = "#00B3CC"

    fonte_titulo = ("Segoe UI", 16, "bold")
    fonte_sub = ("Segoe UI", 10)
    fonte_btn = ("Segoe UI", 11, "bold")

    # Customizando estilos do ttk
    style.configure("Card.TFrame", background=card_color)
    style.configure("Title.TLabel", background=bg_color, font=fonte_titulo, foreground=accent_color)
    style.configure("Texto.TLabel", background=card_color, font=fonte_sub, foreground=text_color)
    style.configure("Status.TLabel", background=card_color, font=("Segoe UI", 10, "italic"), foreground="#888888")
    
    style.configure("TCombobox", fieldbackground="#2A2A35", background="#1C1C22", foreground="white", font=fonte_sub, arrowcolor=accent_color)
    
    # Botão Premium
    style.configure("Action.TButton", font=fonte_btn, background=accent_color, foreground=bg_color, padding=10, borderwidth=0)
    style.map("Action.TButton",
              background=[("active", accent_hover)],
              foreground=[("active", bg_color)])

    # Barra de progresso tunada (Slim & Neon effect)
    style.configure("Neon.Horizontal.TProgressbar", thickness=15, troughcolor="#2A2A35", background=accent_color, bordercolor="#1C1C22", lightcolor=accent_color, darkcolor=accent_color)

    # Layout Principal
    lbl_titulo = ttk.Label(root, text="⚡ Analisador Unificado VIP", style="Title.TLabel")
    lbl_titulo.pack(pady=(30, 10))

    # Card central para agrupar os elementos
    card = ttk.Frame(root, style="Card.TFrame", padding=30)
    card.pack(fill='both', expand=True, padx=40, pady=(0, 40))

    ttk.Label(card, text="Selecione o Módulo de Diagnóstico:", style="Texto.TLabel").pack(anchor='w', pady=(0, 5))
    
    combo = ttk.Combobox(card, values=CATEGORIAS, state='readonly', style="TCombobox", width=40)
    combo.pack(fill='x', pady=(0, 20))
    if CATEGORIAS:
        combo.current(1) # Deixa a segunda opção selecionada por padrão

    ttk.Label(card, text="Progresso da Análise:", style="Texto.TLabel").pack(anchor='w', pady=(10, 5))
    
    progress_bar = ttk.Progressbar(card, mode='determinate', style="Neon.Horizontal.TProgressbar")
    progress_bar.pack(fill='x', pady=(0, 10))

    status_label = ttk.Label(card, text="Aguardando inicialização...", style="Status.TLabel")
    status_label.pack(anchor='w', pady=(0, 20))

    btn = ttk.Button(card, text='🚀 SELECIONAR ARQUIVO E EXECUTAR', style="Action.TButton",
                     command=lambda: escolher_e_executar(root, combo, progress_bar, status_label))
    btn.pack(fill='x', pady=(10, 0))

    root.mainloop()

if __name__ == '__main__':
    criar_interface()
