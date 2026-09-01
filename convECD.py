import streamlit as st
import pandas as pd
from thefuzz import process, fuzz
import os
import io
import json
import re
from datetime import datetime, timedelta

st.set_page_config(page_title="Conversor de Lançamentos Domínio", layout="wide")

st.markdown("<style>.cont-row {border-bottom: 1px solid #f0f2f6; padding: 15px 0px;}</style>", unsafe_allow_html=True)

st.title("🛠️ Conversor de Lançamentos - Formato Domínio Sistemas")
st.info("Geração de arquivos de implantação (Saldos Iniciais e Lançamentos em Lote) no padrão Domínio.")

# --- INICIALIZAÇÃO DO ESTADO ---
if 'de_para_map' not in st.session_state:
    st.session_state.de_para_map = {}

if 'conferidos' not in st.session_state:
    st.session_state.conferidos = {}

# --- CALLBACKS INDIVIDUAIS REATIVOS ---
def on_selectbox_change(cod_conta):
    key = f"sel_widget_{cod_conta}"
    if key in st.session_state:
        valor = st.session_state[key]
        if valor and valor != "-- SELECIONE --" and "📝" not in valor:
            cod_reduzido = valor.split(" | ")[0]
            st.session_state.de_para_map[str(cod_conta)] = str(cod_reduzido)
        elif valor == "📝 -- DIGITAR MANUALMENTE --":
            # Aguarda a digitação no campo manual, sem apagar o dicionário
            pass
        elif valor == "-- SELECIONE --":
            if str(cod_conta) in st.session_state.de_para_map:
                del st.session_state.de_para_map[str(cod_conta)]

def on_manual_change(cod_conta):
    key = f"in_{cod_conta}"
    if key in st.session_state:
        valor = st.session_state[key]
        if valor:
            st.session_state.de_para_map[str(cod_conta)] = str(valor)

def on_confirm_change(cod_conta):
    key = f"conf_{cod_conta}"
    if key in st.session_state:
        st.session_state.conferidos[str(cod_conta)] = bool(st.session_state[key])

def limpar_estados_widgets():
    """Remove cache visual do Streamlit para forçar a reconstrução limpa a partir do backup"""
    keys_to_clear = [k for k in list(st.session_state.keys()) if k.startswith("sel_widget_") or k.startswith("in_") or k.startswith("conf_")]
    for k in keys_to_clear:
        del st.session_state[k]

def limpar_nome_arquivo(nome):
    nome_limpo = re.sub(r'[\\/*?:"<>|]', "", nome)
    return nome_limpo.strip()

def format_moeda(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def ler_arquivo_texto_seguro(file):
    raw_data = file.getvalue()
    try:
        content = raw_data.decode("latin-1")
    except UnicodeError:
        content = raw_data.decode("cp1252", errors="ignore")
    return [linha.strip('\r\n') for linha_crua in content.splitlines() if (linha := linha_crua.strip())]

# --- SIDEBAR: CONFIGURAÇÕES ---
st.sidebar.header("1. Plano de Contas Destino")
usar_padrao = st.sidebar.checkbox("Usar Plano de Contas Padrão?", value=True)

df_novo = None
plano_carregado = False

if usar_padrao:
    caminho_padrao = "plano_padrao.xlsx"
    if os.path.exists(caminho_padrao):
        try:
            df_raw = pd.read_excel(caminho_padrao, header=None)
            if df_raw.shape[1] >= 4:
                df_novo = df_raw.iloc[:, [0, 1, 2, 3]]
                df_novo.columns = ['Código', 'Classificação', 'Nome', 'Tipo']
            else:
                df_novo = df_raw.iloc[:, [0, 1, 2]]
                df_novo.columns = ['Código', 'Classificação', 'Nome']
                df_novo['Tipo'] = 'A'
            plano_carregado = True
        except:
            st.sidebar.error("Erro ao ler plano_padrao.xlsx")
    else:
        st.sidebar.warning("Arquivo 'plano_padrao.xlsx' não encontrado.")
else:
    file_excel = st.sidebar.file_uploader("Subir Novo Plano (Excel)", type=["xlsx"])
    if file_excel:
        try:
            df_raw = pd.read_excel(file_excel, header=None)
            if df_raw.shape[1] >= 4:
                df_novo = df_raw.iloc[:, [0, 1, 2, 3]]
                df_novo.columns = ['Código', 'Classificação', 'Nome', 'Tipo']
            else:
                df_novo = df_raw.iloc[:, [0, 1, 2]]
                df_novo.columns = ['Código', 'Classificação', 'Nome']
                df_novo['Tipo'] = 'A'
            plano_carregado = True
        except Exception as e:
            st.sidebar.error(f"Erro ao ler arquivo Excel: {e}")

# --- CONTROLE DE SENSIBILIDADE DO ROBÔ ---
st.sidebar.divider()
st.sidebar.header("⚙️ Configurações da IA")
sensibilidade_ia = st.sidebar.slider(
    "Sensibilidade do Mapeamento Automático", 
    min_value=50, 
    max_value=95, 
    value=65, 
    step=5,
    help="Valores menores sugerem mais contas, valores maiores exigem mais exatidão do nome."
)

# --- SIDEBAR: UP DO SPED ---
file_sped = None
if plano_carregado:
    st.sidebar.divider()
    st.sidebar.header("2. Upload do Arquivo SPED")
    file_sped = st.sidebar.file_uploader("Subir Arquivo SPED (TXT)", type=["txt"])
else:
    st.sidebar.divider()
    st.sidebar.info("Defina o Plano de Contas para liberar a importação do SPED.")

# --- SEÇÃO BACKUPS E MODELO COMPARTILHADO ---
st.sidebar.divider()
st.sidebar.header("💾 Backup e Modelos")

# 1. Backup Completo (Progresso do Trabalho)
arquivo_backup = st.sidebar.file_uploader("Carregar Progresso Salvo (.json)", type=["json"], key="backup_upload")
if arquivo_backup is not None:
    try:
        file_id = f"{arquivo_backup.name}_{arquivo_backup.size}"
        if st.session_state.get("backup_id") != file_id:
            dados = json.load(arquivo_backup)
            mapa_carregado = dados.get("de_para_map", {}) if isinstance(dados, dict) else dados
            conferidos_carregados = dados.get("conferidos", {}) if isinstance(dados, dict) else {}
            
            # Limpa estados de widgets antigos para forçar reconstrução a partir do backup
            limpar_estados_widgets()
            
            st.session_state.de_para_map = {str(k): str(v) for k, v in mapa_carregado.items()}
            st.session_state.conferidos = {str(k): bool(v) for k, v in conferidos_carregados.items()}
            
            st.session_state["backup_id"] = file_id
            st.sidebar.success("Progresso total carregado!")
            st.rerun()
    except Exception as e:
        st.sidebar.error(f"Erro no backup: {e}")

# 2. Modelo de Mapeamento Puro (Multi-Empresas)
arquivo_modelo = st.sidebar.file_uploader("Carregar Modelo de DE/PARA (.json)", type=["json"], key="modelo_upload")
if arquivo_modelo is not None:
    try:
        file_id = f"mod_{arquivo_modelo.name}_{arquivo_modelo.size}"
        if st.session_state.get("modelo_id") != file_id:
            dados = json.load(arquivo_modelo)
            mapa_modelo = dados.get("de_para_map", {}) if isinstance(dados, dict) else dados
            
            # Limpa estados de widgets antigos
            limpar_estados_widgets()
            
            st.session_state.de_para_map = {str(k): str(v) for k, v in mapa_modelo.items()}
            st.session_state.conferidos = {}
            
            st.session_state["modelo_id"] = file_id
            st.sidebar.success("Modelo aplicado com sucesso!")
            st.rerun()
    except Exception as e:
        st.sidebar.error(f"Erro no modelo: {e}")

placeholder_botao_salvar = st.sidebar.empty()
placeholder_botao_modelo = st.sidebar.empty()

# --- FILTROS DE TELA ---
st.sidebar.divider()
st.sidebar.header("Filtros de Tela")
ocultar_mapeadas = st.sidebar.checkbox("Ocultar contas já mapeadas?", value=False)
ocultar_conferidas = st.sidebar.checkbox("Ocultar contas já conferidas?", value=False)

# --- PROCESSAMENTO PRINCIPAL ---
if file_sped and df_novo is not None:
    df_novo = df_novo.astype(str)
    if 'Tipo' in df_novo.columns:
        df_novo['Tipo'] = df_novo['Tipo'].str.strip().str.upper()
        df_novo = df_novo[~df_novo['Tipo'].str.startswith(('S', 'SIN'))]
        
    df_novo['Display'] = df_novo['Código'] + " | " + df_novo['Classificação'] + " - " + df_novo['Nome']
    df_novo['Grupo'] = df_novo['Classificação'].str.strip().str.lstrip('0').str[0].fillna("0")

    content_sped = ler_arquivo_texto_seguro(file_sped)
    
    # Validação do SPED
    e_sped_valido = any(line.startswith("|0000|") for line in content_sped)
    tem_registro_contabil = any(line.startswith("|I010|") for line in content_sped) or any(line.startswith("|I200|") for line in content_sped)

    if not e_sped_valido or not tem_registro_contabil:
        st.error("❌ O arquivo enviado não é um SPED ECD válido ou não possui dados contábeis.")
        st.stop()

    # Identificadores de Empresa
    nome_empresa = "EMPRESA"
    cnpj_empresa = "00000000000000"
    dt_inicial_sped = None
    dt_final_sped = None
    
    for line in content_sped:
        if line.startswith("|0000|"):
            parts = line.split("|")
            if len(parts) > 5:
                nome_empresa = limpar_nome_arquivo(parts[5])
            cnpj_match = re.search(r'\d{14}', line)
            if cnpj_match:
                cnpj_empresa = cnpj_match.group(0)
            if len(parts) > 3:
                try: dt_inicial_sped = datetime.strptime(parts[3], "%d%m%Y").date()
                except: pass
            if len(parts) > 4:
                try: dt_final_sped = datetime.strptime(parts[4], "%d%m%Y").date()
                except: pass
            break
            
    # Coleta de Saldo e Movimento
    initial_balances = {}
    final_balances = {}
    contas_com_movimento = set()
    
    for line in content_sped:
        if line.startswith("|I155|"):
            reg = line.split("|")
            if len(reg) >= 10:
                cod = reg[2].strip()
                initial_balances[cod] = (reg[4].strip(), reg[5].strip())
                final_balances[cod] = (reg[8].strip(), reg[9].strip())
                contas_com_movimento.add(cod)
        elif line.startswith("|I250|"):
            reg = line.split("|")
            if len(reg) > 2:
                contas_com_movimento.add(reg[2].strip())

    # LEITURA PRECISA UTILIZANDO O CAMPO "COD_NAT"
    contas_origem_data = []
    for line in content_sped:
        if line.startswith("|I050|"):
            reg = line.split("|")
            if len(reg) > 8:
                cod_cta_fixo = reg[6].strip()
                if cod_cta_fixo in contas_com_movimento:
                    cod_nat = reg[3].strip().lstrip('0')
                    nome_conta = reg[8].strip()
                    classif_raw = reg[6].strip()
                    
                    if cod_nat == "1":
                        grupo_aba = "1"
                    elif cod_nat in ["2", "3"]:
                        grupo_aba = "2"
                    elif cod_nat == "4":
                        grupo_aba = "3"
                    else:
                        grupo_aba = "9"
                    
                    contas_origem_data.append({
                        "cod": cod_cta_fixo, 
                        "classif": classif_raw, 
                        "nome": nome_conta, 
                        "grupo": grupo_aba
                    })
    
    df_origem = pd.DataFrame(contas_origem_data).drop_duplicates(subset=['cod'])

    if not df_origem.empty:
        total_mapeadas_count = 0
        map_final_para_geracao = st.session_state.de_para_map.copy()
        process_data = []

        # Algoritmo Fuzzy Match
        for idx, row in df_origem.iterrows():
            cod_atual = str(row['cod'])
            grupo_atual = row['grupo']
            
            df_filtrado = df_novo[df_novo['Grupo'] == grupo_atual]
            df_busca = df_filtrado if not df_filtrado.empty else df_novo
            
            lista_nomes = df_busca['Nome'].tolist()
            candidatos = process.extract(row['nome'], lista_nomes, scorer=fuzz.token_set_ratio, limit=5)
            melhor_match, melhor_score_final = None, -1
            for nome_cand, score_flexivel in candidatos:
                score_grid = fuzz.token_sort_ratio(row['nome'], nome_cand)
                media = (score_flexivel + score_grid) / 2
                if media > melhor_score_final:
                    melhor_score_final, melhor_match = media, nome_cand
            
            score = int(melhor_score_final)
            cod_sugerido_ia, display_sugerido_ia = None, None
            if score >= sensibilidade_ia:
                match_row = df_busca[df_busca['Nome'] == melhor_match]
                if not match_row.empty:
                    cod_sugerido_ia = match_row.iloc[0]['Código']
                    display_sugerido_ia = match_row.iloc[0]['Display']
            
            esta_no_mapa = cod_atual in st.session_state.de_para_map
            valor_no_mapa = str(st.session_state.de_para_map.get(cod_atual, ""))
            
            resolvida = esta_no_mapa or (score >= sensibilidade_ia)
            if resolvida and not esta_no_mapa:
                map_final_para_geracao[cod_atual] = cod_sugerido_ia

            if resolvida:
                total_mapeadas_count += 1
            
            process_data.append({
                "row": row, "df_busca": df_busca, "score": score,
                "cod_sugerido_ia": cod_sugerido_ia, "display_sugerido_ia": display_sugerido_ia,
                "resolvida": resolvida, "esta_no_mapa": esta_no_mapa, "valor_no_mapa": valor_no_mapa
            })

        # Painel de Progresso
        total_contas = len(df_origem)
        conferidas_count = sum(1 for k, v in st.session_state.conferidos.items() if v)
        perc_mapeamento = (total_mapeadas_count / total_contas) if total_contas > 0 else 0.0
        perc_conferencia = (conferidas_count / total_contas) if total_contas > 0 else 0.0

        st.subheader("📊 Progresso Geral")
        col_pb1, col_pb2 = st.columns(2)
        col_pb1.progress(perc_mapeamento, text=f"**Mapeamento:** {total_mapeadas_count}/{total_contas} ({perc_mapeamento * 100:.1f}%)")
        col_pb2.progress(perc_conferencia, text=f"**Conferência Realizada:** {conferidas_count}/{total_contas} ({perc_conferencia * 100:.1f}%)")
        
        # --- SEPARAÇÃO POR ABAS PATRIMONIAIS ---
        st.subheader("🔗 Mapeador por Abas Patrimoniais")
        tab_ativo, tab_passivo, tab_dre, tab_outros = st.tabs(["🔵 Ativo (Grupo 1)", "🟡 Passivo (Grupo 2)", "🟢 DRE (Grupos 3+)", "Outros (Compensação/Outros)"])

        dados_ativos = [x for x in process_data if x['row']['grupo'] == '1']
        dados_passivos = [x for x in process_data if x['row']['grupo'] == '2']
        dados_dre = [x for x in process_data if x['row']['grupo'] == '3']
        dados_outros = [x for x in process_data if x['row']['grupo'] == '9']

        def renderizar_lista_contas(dados_grupo, id_aba):
            busca = st.text_input("🔍 Filtrar contas desta aba:", "", key=f"busca_{id_aba}").strip().lower()
            
            dados_filtrados = []
            for item in dados_grupo:
                row = item['row']
                cod_atual = str(row['cod'])
                conferida = st.session_state.conferidos.get(cod_atual, False)
                resolvida = item['resolvida']

                if busca and not (busca in row['nome'].lower() or busca in cod_atual.lower() or busca in row['classif'].lower()):
                    continue
                if ocultar_mapeadas and resolvida: continue
                if ocultar_conferidas and conferida: continue
                dados_filtrados.append(item)

            if not dados_filtrados:
                st.info("Nenhuma conta localizada nesta aba com os filtros ativos.")
                return

            # Paginação interna por aba
            tamanho_pagina = 15
            total_paginas = max(1, (len(dados_filtrados) + tamanho_pagina - 1) // tamanho_pagina)
            
            chave_pag = f"pag_{id_aba}"
            if chave_pag not in st.session_state:
                st.session_state[chave_pag] = 1
                
            col_pag1, col_pag2, col_pag3 = st.columns([1, 2, 1])
            with col_pag2:
                pag_sel = st.selectbox(f"Página (Aba {id_aba})", range(1, total_paginas + 1), index=min(st.session_state[chave_pag] - 1, total_paginas - 1), key=f"sel_pag_{id_aba}")
                st.session_state[chave_pag] = pag_sel

            offset = (st.session_state[chave_pag] - 1) * tamanho_pagina
            pagina_itens = dados_filtrados[offset : offset + tamanho_pagina]

            st.write(f"Exibindo {offset+1} - {min(offset+tamanho_pagina, len(dados_filtrados))} de {len(dados_filtrados)} contas.")
            st.divider()

            for item in pagina_itens:
                row = item['row']
                cod_atual = str(row['cod'])
                resolvida = item['resolvida']
                esta_no_mapa = cod_atual in st.session_state.de_para_map
                conferida = st.session_state.conferidos.get(cod_atual, False)

                # --- CONTROLE DINÂMICO DE CHAVES DE MEMÓRIA (WIDGETS) ---
                key_widget = f"sel_widget_{cod_atual}"
                key_manual = f"in_{cod_atual}"
                key_conf = f"conf_{cod_atual}"

                # Reconstrói dinamicamente os estados visuais na mudança de aba/página
                if key_widget not in st.session_state:
                    if esta_no_mapa:
                        val_map = st.session_state.de_para_map[cod_atual]
                        match_row = df_novo[df_novo['Código'] == val_map]
                        if not match_row.empty:
                            st.session_state[key_widget] = match_row.iloc[0]['Display']
                        else:
                            st.session_state[key_widget] = "📝 -- DIGITAR MANUALMENTE --"
                            st.session_state[key_manual] = val_map
                    elif item['display_sugerido_ia']:
                        st.session_state[key_widget] = item['display_sugerido_ia']
                    else:
                        st.session_state[key_widget] = "-- SELECIONE --"

                if key_conf not in st.session_state:
                    st.session_state[key_conf] = conferida if resolvida else False

                with st.container():
                    col_origem, col_destino, col_conferido = st.columns([1, 1, 0.4])
                    with col_origem:
                        if conferida:
                            st.markdown(f"~~{row['nome']}~~ ✅")
                        else:
                            st.markdown(f"**{row['nome']}**")
                        st.caption(f"Código SPED: {cod_atual} | Classif: {row['classif']}")
                    
                    with col_destino:
                        opcoes = ["-- SELECIONE --", "📝 -- DIGITAR MANUALMENTE --"] + df_novo['Display'].tolist()
                        
                        # Injeta a conta de backup se ela não estiver na lista padrão
                        if esta_no_mapa:
                            val_map = st.session_state.de_para_map[cod_atual]
                            match_row = df_novo[df_novo['Código'] == val_map]
                            if not match_row.empty:
                                display_str = match_row.iloc[0]['Display']
                                if display_str not in opcoes:
                                    opcoes.insert(2, display_str)

                        if esta_no_mapa: 
                            st.info("📌 Mapeado pelo Usuário")
                        elif item['score'] >= 85: 
                            st.success(f"🟢 Alta Confiança ({item['score']}% - IA)")
                        elif item['score'] >= sensibilidade_ia: 
                            st.warning(f"🟡 Média Confiança ({item['score']}% - IA)")
                        else: 
                            st.error(f"🔴 Não mapeada (Baixa Confiança)")

                        # Selectbox agora usa "on_change" com salvamento síncrono dedicado
                        st.selectbox(
                            label=f"sel_{cod_atual}", 
                            options=opcoes, 
                            key=key_widget, 
                            label_visibility="collapsed",
                            on_change=on_selectbox_change,
                            args=(cod_atual,)
                        )

                        if st.session_state.get(key_widget) == "📝 -- DIGITAR MANUALMENTE --":
                            if key_manual not in st.session_state:
                                st.session_state[key_manual] = st.session_state.de_para_map.get(cod_atual, "")
                            st.text_input(
                                f"Código Manual para {cod_atual}:", 
                                key=key_manual,
                                on_change=on_manual_change,
                                args=(cod_atual,)
                            )
                    
                    with col_conferido:
                        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
                        st.checkbox(
                            "Conferido",
                            key=key_conf,
                            on_change=on_confirm_change,
                            args=(cod_atual,),
                            disabled=not resolvida
                        )
                    st.markdown("---")

        with tab_ativo: renderizar_lista_contas(dados_ativos, "ativo")
        with tab_passivo: renderizar_lista_contas(dados_passivos, "passivo")
        with tab_dre: renderizar_lista_contas(dados_dre, "dre")
        with tab_outros: renderizar_lista_contas(dados_outros, "outros")

        # --- EXPORTAÇÕES ---
        st.divider()
        st.subheader("📂 Exportação para a Domínio Sistemas")
        col1, col2 = st.columns(2)

        # Reconstrução Corrigida das Transações (Lançamentos Diários)
        transactions = []
        current_i200_date = None
        debits = []
        credits = []
        
        for line in content_sped:
            if line.startswith("|I200|"):
                if (debits or credits) and current_i200_date:
                    transactions.append((current_i200_date, debits.copy(), credits.copy()))
                debits = []
                credits = []
                reg = line.split("|")
                if len(reg) > 3:
                    dt_str = reg[3].strip()
                    try:
                        dt_obj = datetime.strptime(dt_str, "%d%m%Y")
                        current_i200_date = dt_obj.strftime("%d/%m/%Y")
                    except Exception as e:
                        current_i200_date = None
            elif line.startswith("|I250|") and current_i200_date:
                reg = line.split("|")
                if len(reg) > 5:
                    cod_cta = reg[2].strip()
                    val_str = reg[4].strip()
                    ind_dc = reg[5].strip()
                    cod_hist = reg[6].strip() if len(reg) > 6 else "1"
                    hist = reg[7].strip() if (len(reg) > 7 and reg[7].strip()) else "LANCAMENTO IMPORTADO"
                    
                    try: val_float = float(val_str.replace(",", "."))
                    except: val_float = 0.0
                    
                    if ind_dc == "D":
                        debits.append((cod_cta, val_float, hist, cod_hist))
                    else:
                        credits.append((cod_cta, val_float, hist, cod_hist))
                        
        if (debits or credits) and current_i200_date:
            transactions.append((current_i200_date, debits, credits))

        pendentes = len(df_origem) - total_mapeadas_count

        with col1:
            st.markdown("### 💾 Arquivos de Saída (Lotes Domínio)")
            
            if pendentes > 0:
                st.warning(f"⚠️ Mapeie todas as {pendentes} contas restantes para liberar os downloads.")
            else:
                # -------------------------------------------------------------
                # DOWNLOAD 1: LANÇAMENTOS DO PERÍODO (MOVIMENTAÇÃO)
                # -------------------------------------------------------------
                if st.button("🚀 Processar Lançamentos Diários", use_container_width=True):
                    dominio_lines = []
                    dominio_lines.append(f"|0000|{cnpj_empresa}|")
                    
                    for dt, deb_list, cred_list in transactions:
                        debs = [(map_final_para_geracao.get(c, c), v, h, ch) for c, v, h, ch in deb_list]
                        creds = [(map_final_para_geracao.get(c, c), v, h, ch) for c, v, h, ch in cred_list]
                        
                        while debs and creds:
                            d_acc, d_val, d_hist, d_cod = debs[0]
                            c_acc, c_val, c_hist, c_cod = creds[0]
                            val = min(d_val, c_val)
                            
                            val_str = f"{val:.2f}".replace(".", ",")
                            hist_limpo = re.sub(r'[|]', '', d_hist or c_hist or "LANCAMENTO")
                            cod_hist_limpo = d_cod or c_cod or "1"
                            
                            dominio_lines.append("|6000|X||||")
                            dominio_lines.append(f"|6100|{dt}|{d_acc}|{c_acc}|{val_str}|{cod_hist_limpo}|{hist_limpo}|GERENTE|||")
                            
                            if d_val == c_val:
                                debs.pop(0)
                                creds.pop(0)
                            elif d_val > c_val:
                                debs[0] = (d_acc, d_val - val, d_hist, d_cod)
                                creds.pop(0)
                            else:
                                creds[0] = (c_acc, c_val - val, c_hist, c_cod)
                                debs.pop(0)
                                
                    dominio_buffer = "\r\n".join(dominio_lines).encode("latin-1", errors="replace")
                    st.session_state.dominio_arquivo_pronto = dominio_buffer
                    st.success("✅ Lançamentos Diários prontos para download!")
                    
                if "dominio_arquivo_pronto" in st.session_state:
                    st.download_button(
                        label="⬇️ Baixar Lançamentos Diários (.txt)",
                        data=st.session_state.dominio_arquivo_pronto,
                        file_name=f"LAN_DIARIOS_DOMINIO_{nome_empresa}.txt",
                        mime="text/plain",
                        use_container_width=True
                    )
                
                st.markdown("---")
                
                # -------------------------------------------------------------
                # DOWNLOAD 2: SALDOS INICIAIS (ABERTURA CONTÁBIL NO DOMÍNIO)
                # -------------------------------------------------------------
                st.markdown("**Parâmetros do Saldo Inicial**")
                
                data_padrao_saldo = datetime.today()
                if dt_inicial_sped:
                    data_padrao_saldo = dt_inicial_sped - timedelta(days=1)
                
                data_balanco = st.date_input("Data de Implantação do Saldo:", data_padrao_saldo, format="DD/MM/YYYY")
                dt_fmt = data_balanco.strftime("%d/%m/%Y")
                
                if st.button("🚀 Processar Saldos Iniciais", use_container_width=True):
                    balanco_lines = [f"|0000|{cnpj_empresa}|"]
                    balanco_lines.append("|6000|V||||")
                    
                    has_balanco = False
                    for cod_antigo in map_final_para_geracao:
                        novo = map_final_para_geracao[cod_antigo].replace("|", "")
                        val_str, dc = initial_balances.get(cod_antigo, ("0,00", "D"))
                        
                        try: val_float = float(val_str.replace(",", "."))
                        except: val_float = 0.0
                        
                        if val_float > 0:
                            if dc == 'D':
                                linha = f"|6100|{dt_fmt}|{novo}||{val_str}|1|SALDO INICIAL {dt_fmt}|GERENTE|||"
                            else:
                                linha = f"|6100|{dt_fmt}||{novo}|{val_str}|1|SALDO INICIAL {dt_fmt}|GERENTE|||"
                            
                            balanco_lines.append(linha)
                            has_balanco = True
                    
                    if has_balanco:
                        st.session_state.dominio_saldos_pronto = "\r\n".join(balanco_lines).encode("latin-1", errors="replace")
                        st.success("✅ Saldos Iniciais processados com sucesso!")
                    else:
                        st.warning("⚠️ Não foram detectados saldos iniciais neste arquivo.")
                
                if "dominio_saldos_pronto" in st.session_state:
                    st.download_button(
                        label="⬇️ Baixar Saldos Iniciais (.txt)",
                        data=st.session_state.dominio_saldos_pronto,
                        file_name=f"SALDOS_INICIAIS_DOMINIO_{nome_empresa}.txt",
                        mime="text/plain",
                        use_container_width=True
                    )

        # Consistência de Saldos (Partidas Dobradas)
        with col2:
            st.markdown("### 🧮 Consistência dos Saldos (Partidas Dobradas)")
            total_d, total_c = 0.0, 0.0
            
            for cod_antigo in map_final_para_geracao:
                val_str, dc = final_balances.get(cod_antigo, ("0,00", "D"))
                try: val_float = float(val_str.replace(",", "."))
                except: val_float = 0.0
                
                if val_float > 0:
                    if dc == 'D': total_d += val_float
                    else: total_c += val_float
            
            diferenca = total_d - total_c
            col_d1, col_d2 = st.columns(2)
            col_d1.metric("Soma de Saldos Devedores", format_moeda(total_d))
            col_d2.metric("Soma de Saldos Credores", format_moeda(total_c))
            
            if abs(diferenca) > 0.01:
                st.error(f"⚠️ **Atenção:** Existe uma divergência de fechamento patrimonial de **{format_moeda(diferenca)}**.")
            else:
                st.success("✅ Partidas dobradas íntegras! Diferença de R$ 0,00.")

    else: 
        st.error("Nenhuma conta com movimento detectada.")

# --- GERENCIADOR DE BACKUPS NA SIDEBAR ---
if 'de_para_map' in st.session_state and len(st.session_state.de_para_map) > 0:
    with placeholder_botao_salvar:
        backup_data = {
            "de_para_map": st.session_state.de_para_map,
            "conferidos": st.session_state.conferidos
        }
        st.sidebar.download_button(
            "⬇️ Exportar Backup Total", 
            json.dumps(backup_data, indent=4), 
            "backup_mapeamento_dominio.json", 
            "application/json", 
            use_container_width=True,
            help="Exporta o progresso completo + histórico de checagens (específico desta empresa)."
        )

    with placeholder_botao_modelo:
        modelo_data = {
            "de_para_map": st.session_state.de_para_map
        }
        st.sidebar.download_button(
            "⬇️ Exportar Apenas Modelo (DE/PARA)", 
            json.dumps(modelo_data, indent=4), 
            "modelo_de_para_compartilhado.json", 
            "application/json", 
            use_container_width=True,
            help="Exporta somente os mapeamentos para aplicar em outras empresas com o mesmo plano destino."
        )
else: 
    st.info("Carregue o Plano de Contas e o SPED para iniciar os trabalhos.")
