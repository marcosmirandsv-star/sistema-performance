import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import io
import json
import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from supabase import create_client, Client
import hashlib
import requests

# ============================================
# CONFIGURAÇÕES INICIAIS
# ============================================

st.set_page_config(
    page_title="Sistema de Performance",
    page_icon="📊",
    layout="wide"
)

# ============================================
# SUPABASE CONFIG
# ============================================

def init_supabase():
    """Inicializa conexão com Supabase"""
    try:
        url = st.secrets["SUPABASE_URL"]
        key = st.secrets["SUPABASE_KEY"]
        return create_client(url, key)
    except:
        url = os.environ.get("SUPABASE_URL")
        key = os.environ.get("SUPABASE_KEY")
        if url and key:
            return create_client(url, key)
        return None

# ============================================
# LEITURA ROBUSTA DE ARQUIVOS
# ============================================

def normalizar_nome_coluna(nome):
    """
    Normaliza o nome da coluna para comparação
    Remove espaços extras, caracteres especiais e converte para minúsculo
    """
    import re
    # Converter para minúsculo
    nome = nome.lower().strip()
    # Remover caracteres especiais e espaços extras
    nome = re.sub(r'[^a-záéíóúãõç0-9\s]', '', nome)
    # Remover espaços múltiplos
    nome = re.sub(r'\s+', ' ', nome).strip()
    return nome

def identificar_coluna_flexivel(df, padroes):
    """
    Identifica uma coluna baseada em padrões de nome
    """
    colunas = df.columns.tolist()
    
    for col in colunas:
        col_normalizado = normalizar_nome_coluna(col)
        for padrao in padroes:
            padrao_normalizado = normalizar_nome_coluna(padrao)
            if padrao_normalizado in col_normalizado or col_normalizado in padrao_normalizado:
                return col
    return None

def carregar_arquivo_satisfacao(arquivo):
    """
    Carrega o arquivo de satisfação de forma robusta
    """
    try:
        df = pd.read_excel(arquivo)
    except Exception as e:
        try:
            df = pd.read_csv(arquivo, encoding='utf-8')
        except:
            try:
                df = pd.read_csv(arquivo, encoding='latin-1')
            except:
                st.error(f"❌ Não foi possível ler o arquivo de satisfação: {str(e)}")
                st.info("Verifique se é um arquivo Excel (.xlsx) ou CSV válido.")
                return None
    
    colunas = df.columns.tolist()
    st.info(f"🔍 Colunas encontradas no arquivo de satisfação: {', '.join(colunas)}")
    
    # Identificar cada coluna
    col_id = identificar_coluna_flexivel(df, [
        'ID do ticket', 'Ticket ID', 'ID', 'ticket_id', 'id_ticket'
    ])
    
    col_satisfacao = identificar_coluna_flexivel(df, [
        'Índice de satisfação do ticket', 'Satisfaction', 'Índice de satisfação',
        'Ticket Satisfaction', 'satisfacao', 'status_satisfacao',
        'índice de satisfação do ticket', 'satisfação'
    ])
    
    col_atribuido = identificar_coluna_flexivel(df, [
        'Nome do atribuído', 'Assignee', 'Atribuído', 'Responsável',
        'assignee_name', 'nome_atribuido', 'nome do atribuído'
    ])
    
    # Verificar se encontrou todas
    if not col_id:
        st.error("❌ Coluna 'ID do ticket' não encontrada no arquivo de satisfação.")
        st.info(f"Colunas disponíveis: {', '.join(colunas)}")
        return None
    
    if not col_satisfacao:
        st.error("❌ Coluna de satisfação não encontrada no arquivo de satisfação.")
        st.info(f"Colunas disponíveis: {', '.join(colunas)}")
        return None
    
    if not col_atribuido:
        st.error("❌ Coluna 'Nome do atribuído' não encontrada no arquivo de satisfação.")
        st.info(f"Colunas disponíveis: {', '.join(colunas)}")
        return None
    
    # Renomear colunas para o padrão esperado
    df_renomeado = df.rename(columns={
        col_id: 'ID do ticket',
        col_satisfacao: 'Índice de satisfação do ticket',
        col_atribuido: 'Nome do atribuído'
    })
    
    # Manter apenas as colunas necessárias
    colunas_necessarias = ['ID do ticket', 'Índice de satisfação do ticket', 'Nome do atribuído']
    df_final = df_renomeado[colunas_necessarias].copy()
    
    st.success(f"✅ Arquivo de satisfação carregado! {len(df_final)} registros encontrados.")
    
    with st.expander("📊 Amostra dos dados de satisfação"):
        st.dataframe(df_final.head(10))
    
    return df_final

def carregar_arquivo_inativos(arquivo):
    """
    Carrega o arquivo de inativos de forma robusta
    """
    try:
        df = pd.read_excel(arquivo)
    except Exception as e:
        try:
            df = pd.read_csv(arquivo, encoding='utf-8')
        except:
            try:
                df = pd.read_csv(arquivo, encoding='latin-1')
            except:
                st.error(f"❌ Não foi possível ler o arquivo de inativos: {str(e)}")
                st.info("Verifique se é um arquivo Excel (.xlsx) ou CSV válido.")
                return None
    
    colunas = df.columns.tolist()
    st.info(f"🔍 Colunas encontradas no arquivo de inativos: {', '.join(colunas)}")
    
    # Identificar cada coluna
    col_id = identificar_coluna_flexivel(df, [
        'ID do ticket', 'Ticket ID', 'ID', 'ticket_id', 'id_ticket'
    ])
    
    col_atribuido = identificar_coluna_flexivel(df, [
        'Nome do atribuído', 'Assignee', 'Atribuído', 'Responsável',
        'assignee_name', 'nome_atribuido', 'nome do atribuído'
    ])
    
    # Verificar se encontrou todas
    if not col_id:
        st.error("❌ Coluna 'ID do ticket' não encontrada no arquivo de inativos.")
        st.info(f"Colunas disponíveis: {', '.join(colunas)}")
        return None
    
    if not col_atribuido:
        st.error("❌ Coluna 'Nome do atribuído' não encontrada no arquivo de inativos.")
        st.info(f"Colunas disponíveis: {', '.join(colunas)}")
        return None
    
    # Renomear colunas para o padrão esperado
    df_renomeado = df.rename(columns={
        col_id: 'ID do ticket',
        col_atribuido: 'Nome do atribuído'
    })
    
    # Manter apenas as colunas necessárias
    colunas_necessarias = ['ID do ticket', 'Nome do atribuído']
    df_final = df_renomeado[colunas_necessarias].copy()
    
    st.success(f"✅ Arquivo de inativos carregado! {len(df_final)} registros encontrados.")
    
    with st.expander("📊 Amostra dos dados de inativos"):
        st.dataframe(df_final.head(10))
    
    return df_final

# ============================================
# FUNÇÕES EXISTENTES (mantidas)
# ============================================

def salvar_historico(supabase, dados, mes_ano, gestor):
    """Salva os dados processados no Supabase"""
    if not supabase:
        return False, "Supabase não configurado"
    
    try:
        registros = []
        for analista, d in dados.items():
            registros.append({
                'mes_ano': mes_ano,
                'analista': analista,
                'gestor': gestor,
                'csat': d['csat'],
                'perc_avaliacoes': d['perc_avaliacoes'],
                'perc_envio': d['perc_envio'],
                'total_atendimentos': d['total_atendimentos'],
                'total_inativos': d['total_inativos'],
                'validos': d['validos'],
                'avaliacoes': d['avaliacoes'],
                'positivos': d['positivos'],
                'negativos': d['negativos'],
                'meta_csat': d['meta_csat'],
                'delta_csat': d['delta_csat'],
                'meta_geral': d['meta_geral'],
                'status': d['status']
            })
        
        # Verificar se já existe registro para o período
        existing = supabase.table('historico_performance').select('*').eq('mes_ano', mes_ano).eq('gestor', gestor).execute()
        
        if existing.data:
            return False, f"Já existe um relatório importado para o período {mes_ano}."
        
        for registro in registros:
            supabase.table('historico_performance').insert(registro).execute()
        
        return True, "Salvo com sucesso"
    except Exception as e:
        return False, str(e)

def substituir_historico(supabase, dados, mes_ano, gestor):
    """Substitui os dados existentes no Supabase"""
    if not supabase:
        return False, "Supabase não configurado"
    
    try:
        # Deletar registros existentes
        supabase.table('historico_performance').delete().eq('mes_ano', mes_ano).eq('gestor', gestor).execute()
        
        # Inserir novos registros
        registros = []
        for analista, d in dados.items():
            registros.append({
                'mes_ano': mes_ano,
                'analista': analista,
                'gestor': gestor,
                'csat': d['csat'],
                'perc_avaliacoes': d['perc_avaliacoes'],
                'perc_envio': d['perc_envio'],
                'total_atendimentos': d['total_atendimentos'],
                'total_inativos': d['total_inativos'],
                'validos': d['validos'],
                'avaliacoes': d['avaliacoes'],
                'positivos': d['positivos'],
                'negativos': d['negativos'],
                'meta_csat': d['meta_csat'],
                'delta_csat': d['delta_csat'],
                'meta_geral': d['meta_geral'],
                'status': d['status']
            })
        
        for registro in registros:
            supabase.table('historico_performance').insert(registro).execute()
        
        return True, "Substituído com sucesso"
    except Exception as e:
        return False, str(e)

def carregar_historico(supabase, mes_ano=None, gestor=None, analista=None):
    """Carrega histórico do Supabase"""
    if not supabase:
        return None
    
    try:
        query = supabase.table('historico_performance').select('*')
        
        if mes_ano:
            query = query.eq('mes_ano', mes_ano)
        if gestor:
            query = query.eq('gestor', gestor)
        if analista:
            query = query.eq('analista', analista)
        
        response = query.execute()
        return pd.DataFrame(response.data)
    except Exception as e:
        st.error(f"Erro ao carregar histórico: {str(e)}")
        return None

def salvar_podio_manual(supabase, mes_ano, gestor, podio):
    """Salva edição manual do pódio"""
    if not supabase:
        return False, "Supabase não configurado"
    
    try:
        supabase.table('podio_manual').delete().eq('mes_ano', mes_ano).eq('gestor', gestor).execute()
        
        for i, (nome, csat, atendimentos, perc_avaliacoes) in enumerate(podio, 1):
            supabase.table('podio_manual').insert({
                'mes_ano': mes_ano,
                'gestor': gestor,
                'posicao': i,
                'analista': nome,
                'csat': csat,
                'atendimentos': atendimentos
            }).execute()
        
        return True, "Salvo com sucesso"
    except Exception as e:
        return False, str(e)

def carregar_podio_manual(supabase, mes_ano, gestor):
    """Carrega pódio manual do Supabase"""
    if not supabase:
        return None
    
    try:
        response = supabase.table('podio_manual').select('*').eq('mes_ano', mes_ano).eq('gestor', gestor).order('posicao').execute()
        if response.data:
            return [(d['analista'], d['csat'], d['atendimentos'], 0) for d in response.data]
        return None
    except Exception as e:
        return None

def listar_periodos(supabase, gestor=None):
    """Lista todos os períodos disponíveis"""
    if not supabase:
        return []
    
    try:
        query = supabase.table('historico_performance').select('mes_ano, gestor, data_criacao')
        if gestor:
            query = query.eq('gestor', gestor)
        
        response = query.execute()
        
        periodos = {}
        for item in response.data:
            mes_ano = item['mes_ano']
            if mes_ano not in periodos:
                periodos[mes_ano] = {
                    'mes_ano': mes_ano,
                    'gestor': item['gestor'],
                    'data_criacao': item.get('data_criacao', datetime.now().isoformat())
                }
        
        return list(periodos.values())
    except Exception as e:
        st.error(f"Erro ao listar períodos: {str(e)}")
        return []

def excluir_periodo(supabase, mes_ano, gestor):
    """Exclui todos os registros de um período"""
    if not supabase:
        return False, "Supabase não configurado"
    
    try:
        supabase.table('historico_performance').delete().eq('mes_ano', mes_ano).eq('gestor', gestor).execute()
        supabase.table('podio_manual').delete().eq('mes_ano', mes_ano).eq('gestor', gestor).execute()
        return True, "Período excluído com sucesso"
    except Exception as e:
        return False, str(e)

def verificar_periodo_existente(supabase, mes_ano, gestor):
    """Verifica se um período já existe no Supabase"""
    if not supabase:
        return False
    
    try:
        response = supabase.table('historico_performance').select('*').eq('mes_ano', mes_ano).eq('gestor', gestor).execute()
        return len(response.data) > 0
    except:
        return False

# ============================================
# GERENCIAMENTO DE USUÁRIOS
# ============================================

def hash_senha(senha):
    """Cria hash da senha"""
    return hashlib.sha256(senha.encode()).hexdigest()

def carregar_usuarios():
    """Carrega usuários do Supabase ou arquivo local"""
    supabase = init_supabase()
    if supabase:
        try:
            response = supabase.table('usuarios').select('*').execute()
            usuarios = {}
            for u in response.data:
                usuarios[u['usuario']] = {
                    'senha': u['senha'],
                    'nome': u['nome'],
                    'gestor': u['gestor'],
                    'acesso_total': u.get('acesso_total', False)
                }
            return usuarios
        except:
            pass
    
    try:
        if os.path.exists('usuarios.json'):
            with open('usuarios.json', 'r', encoding='utf-8') as f:
                return json.load(f)
    except:
        pass
    
    usuarios = {
        "marcos": {
            "senha": hash_senha("marcos2026"),
            "nome": "Marcos Miranda",
            "gestor": "Marcos Miranda - Chat Notas",
            "acesso_total": True
        },
        "polyana": {
            "senha": hash_senha("polyana2026"),
            "nome": "Polyana Ventura",
            "gestor": "Polyana Ventura - Chat Outros",
            "acesso_total": False
        },
        "carine": {
            "senha": hash_senha("carine2026"),
            "nome": "Carine Melo",
            "gestor": "Marcos Miranda - Chat Notas",
            "acesso_total": True
        }
    }
    
    try:
        with open('usuarios.json', 'w', encoding='utf-8') as f:
            json.dump(usuarios, f, ensure_ascii=False, indent=2)
    except:
        pass
    
    return usuarios

def salvar_usuario_supabase(supabase, usuario, nome, senha_hash, gestor, acesso_total=False):
    """Salva usuário no Supabase"""
    if not supabase:
        return False
    
    try:
        supabase.table('usuarios').insert({
            'usuario': usuario,
            'nome': nome,
            'senha': senha_hash,
            'gestor': gestor,
            'acesso_total': acesso_total
        }).execute()
        return True
    except:
        return False

# ============================================
# GERENCIAMENTO DE ANALISTAS
# ============================================

def carregar_analistas():
    """Carrega a lista de analistas do arquivo"""
    try:
        if os.path.exists('analistas.json'):
            with open('analistas.json', 'r', encoding='utf-8') as f:
                return json.load(f)
    except:
        pass
    
    config = {
        "Marcos Miranda - Chat Notas": {
            "gestor": "Marcos Miranda",
            "meta_geral_avaliacoes": 25,
            "membros": {
                "Ana Claudia Corrêa": {"meta_csat": 90, "ativo": True},
                "João Vitor Almeida": {"meta_csat": 90, "ativo": True},
                "João Pedro Vianey": {"meta_csat": 90, "ativo": True},
                "Lorena Almeida": {"meta_csat": 86, "ativo": True},
                "Paulo Victor": {"meta_csat": 86, "ativo": True},
                "Rayane Nunes": {"meta_csat": 86, "ativo": True},
                "Thiago Reis": {"meta_csat": 90, "ativo": True},
                "Vanessa Silva": {"meta_csat": 86, "ativo": True}
            }
        },
        "Polyana Ventura - Chat Outros": {
            "gestor": "Polyana Ventura",
            "meta_geral_avaliacoes": 25,
            "membros": {
                "Christian Matozinho": {"meta_csat": 86, "ativo": True},
                "Diego Machado": {"meta_csat": 86, "ativo": True},
                "Igor Siqueira": {"meta_csat": 86, "ativo": True},
                "Ismael Chagas Bessa": {"meta_csat": 86, "ativo": True},
                "João Pedro Santana": {"meta_csat": 90, "ativo": True},
                "Karolyne Moreira": {"meta_csat": 86, "ativo": True},
                "Luan Pereira": {"meta_csat": 86, "ativo": True},
                "Mario Junior": {"meta_csat": 90, "ativo": True},
                "Maycon Oliveira": {"meta_csat": 86, "ativo": True},
                "Miguel Augusto": {"meta_csat": 86, "ativo": True},
                "Polliana Santana": {"meta_csat": 86, "ativo": True}
            }
        }
    }
    
    try:
        with open('analistas.json', 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except:
        pass
    
    return config

def salvar_analistas(analistas):
    """Salva a lista de analistas no arquivo"""
    try:
        with open('analistas.json', 'w', encoding='utf-8') as f:
            json.dump(analistas, f, ensure_ascii=False, indent=2)
        return True
    except:
        return False

# ============================================
# LOGIN
# ============================================

def fazer_login():
    """Tela de login"""
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔐 Login")
    
    usuarios = carregar_usuarios()
    
    usuario = st.sidebar.text_input("Usuário")
    senha = st.sidebar.text_input("Senha", type="password")
    
    if st.sidebar.button("Entrar", use_container_width=True):
        if usuario in usuarios and usuarios[usuario]["senha"] == hash_senha(senha):
            st.session_state.logado = True
            st.session_state.usuario = usuario
            st.session_state.nome_usuario = usuarios[usuario]["nome"]
            st.session_state.gestor = usuarios[usuario]["gestor"]
            st.session_state.acesso_total = usuarios[usuario].get("acesso_total", False)
            st.rerun()
        else:
            st.sidebar.error("❌ Usuário ou senha inválidos!")
    
    if st.session_state.get('logado', False):
        st.sidebar.success(f"✅ Logado como {st.session_state.nome_usuario}")
        
        if st.session_state.get('acesso_total', False):
            st.sidebar.info("🔑 Acesso Total - Todos os times")
        else:
            st.sidebar.info(f"👥 Time: {st.session_state.gestor}")
        
        if st.session_state.get('acesso_total', False):
            if st.sidebar.button("👤 Cadastrar Usuário", use_container_width=True):
                st.session_state.cadastrar_usuario = True
        
        if st.sidebar.button("Sair", use_container_width=True):
            st.session_state.logado = False
            st.session_state.usuario = None
            st.session_state.nome_usuario = None
            st.session_state.gestor = None
            st.session_state.acesso_total = False
            st.rerun()
        return True
    return False

def cadastrar_usuario():
    """Interface para cadastrar novos usuários"""
    st.header("👤 Cadastrar Novo Usuário")
    
    supabase = init_supabase()
    usuarios = carregar_usuarios()
    
    col1, col2 = st.columns(2)
    
    with col1:
        novo_usuario = st.text_input("Nome de usuário (login)")
        novo_nome = st.text_input("Nome completo")
    
    with col2:
        nova_senha = st.text_input("Senha", type="password")
        confirma_senha = st.text_input("Confirmar senha", type="password")
        gestor_usuario = st.selectbox(
            "Gestor",
            ["Marcos Miranda - Chat Notas", "Polyana Ventura - Chat Outros"]
        )
        acesso_total = st.checkbox("🔑 Acesso Total ao Sistema", value=False)
    
    if st.button("✅ Cadastrar", use_container_width=True):
        if not novo_usuario or not novo_nome or not nova_senha:
            st.error("❌ Preencha todos os campos!")
        elif novo_usuario in usuarios:
            st.error("❌ Usuário já existe!")
        elif nova_senha != confirma_senha:
            st.error("❌ Senhas não conferem!")
        elif len(nova_senha) < 6:
            st.error("❌ Senha deve ter no mínimo 6 caracteres!")
        else:
            senha_hash = hash_senha(nova_senha)
            
            if supabase:
                if salvar_usuario_supabase(supabase, novo_usuario, novo_nome, senha_hash, gestor_usuario, acesso_total):
                    st.success(f"✅ Usuário {novo_usuario} cadastrado com sucesso!")
                    st.rerun()
                else:
                    st.error("❌ Erro ao salvar no Supabase!")
            else:
                usuarios[novo_usuario] = {
                    "senha": senha_hash,
                    "nome": novo_nome,
                    "gestor": gestor_usuario,
                    "acesso_total": acesso_total
                }
                try:
                    with open('usuarios.json', 'w', encoding='utf-8') as f:
                        json.dump(usuarios, f, ensure_ascii=False, indent=2)
                    st.success(f"✅ Usuário {novo_usuario} cadastrado com sucesso!")
                    st.rerun()
                except:
                    st.error("❌ Erro ao salvar usuário!")
    
    if st.button("🔙 Voltar"):
        st.session_state.cadastrar_usuario = False
        st.rerun()

# ============================================
# FUNÇÕES AUXILIARES
# ============================================

def get_genero_neutro(nome):
    """Determina o gênero baseado no nome"""
    nomes_femininos = [
        'Ana', 'Maria', 'Paula', 'Lorena', 'Vanessa', 'Rayane', 
        'Karolyne', 'Polliana', 'Polyana', 'Ariane', 'Julia',
        'Claudia', 'Almeida', 'Carine'
    ]
    
    for nome_feminino in nomes_femininos:
        if nome.startswith(nome_feminino):
            return "colaboradora"
    
    if nome.endswith('a') and not nome.endswith('as'):
        return "colaboradora"
    
    return "colaborador"

def extrair_periodo(df_satisfacao):
    """Extrai o período dos dados"""
    colunas = df_satisfacao.columns.tolist()
    
    for col in colunas:
        if 'data' in col.lower() or 'período' in col.lower() or 'periodo' in col.lower():
            try:
                datas = pd.to_datetime(df_satisfacao[col])
                mes = datas.dt.month_name().iloc[0]
                ano = datas.dt.year.iloc[0]
                return f"{mes} {ano}"
            except:
                pass
    
    return datetime.now().strftime('%B %Y')

def processar_dados(df_satisfacao, df_inativos, analistas_config):
    """Processa os dados e calcula todas as métricas"""
    resultados = {}
    
    analistas_ativos = []
    for gestor, config in analistas_config.items():
        for analista, dados in config['membros'].items():
            if dados['ativo']:
                analistas_ativos.append(analista)
    
    for analista in analistas_ativos:
        tickets_analista = df_satisfacao[df_satisfacao['Nome do atribuído'] == analista]
        total_atendimentos = len(tickets_analista)
        
        if total_atendimentos == 0:
            continue
        
        inativos_analista = df_inativos[df_inativos['Nome do atribuído'] == analista]
        total_inativos = len(inativos_analista)
        validos = total_atendimentos - total_inativos
        
        avaliacoes = len(tickets_analista[tickets_analista['Índice de satisfação do ticket'].isin(['Good', 'Bad'])])
        positivos = len(tickets_analista[tickets_analista['Índice de satisfação do ticket'] == 'Good'])
        negativos = len(tickets_analista[tickets_analista['Índice de satisfação do ticket'] == 'Bad'])
        
        perc_avaliacoes = (avaliacoes / validos * 100) if validos > 0 else 0
        csat = (positivos / avaliacoes * 100) if avaliacoes > 0 else 0
        
        nao_avaliaram = validos - avaliacoes
        perc_envio = (nao_avaliaram / validos * 100) if validos > 0 else 0
        
        gestor = "Outros"
        meta_csat = 86
        meta_geral = 25
        
        for gestor_nome, config in analistas_config.items():
            if analista in config['membros']:
                gestor = gestor_nome
                meta_csat = config['membros'][analista]['meta_csat']
                meta_geral = config.get('meta_geral_avaliacoes', 25)
                break
        
        delta_csat = csat - meta_csat
        
        if csat >= meta_csat and perc_avaliacoes >= meta_geral:
            status = "🟢 Meta Superada"
        elif csat >= meta_csat or perc_avaliacoes >= meta_geral:
            status = "🟡 Atenção"
        else:
            status = "🔴 Crítico"
        
        resultados[analista] = {
            'total_atendimentos': total_atendimentos,
            'total_inativos': total_inativos,
            'validos': validos,
            'avaliacoes': avaliacoes,
            'positivos': positivos,
            'negativos': negativos,
            'perc_avaliacoes': round(perc_avaliacoes, 2),
            'perc_envio': round(perc_envio, 2),
            'csat': round(csat, 2),
            'meta_csat': meta_csat,
            'delta_csat': round(delta_csat, 2),
            'meta_geral': meta_geral,
            'status': status,
            'gestor': gestor
        }
    
    return resultados

def calcular_media_operacao(resultados, gestor=None):
    """Calcula a média de atendimentos por agente"""
    if gestor:
        dados_gestor = [d for a, d in resultados.items() if d['gestor'] == gestor]
    else:
        dados_gestor = list(resultados.values())
    
    total_atendimentos = sum([d['total_atendimentos'] for d in dados_gestor])
    total_analistas = len(dados_gestor)
    return round(total_atendimentos / total_analistas) if total_analistas > 0 else 0

def calcular_podio(resultados, gestor=None, media_atendimentos=None, limiar_csat=90):
    """Calcula o pódio baseado em CSAT ≥ 90% e Atendimentos ≥ média"""
    if gestor:
        dados_filtrados = {k: v for k, v in resultados.items() 
                          if v['gestor'] == gestor and v['avaliacoes'] > 0}
    else:
        dados_filtrados = {k: v for k, v in resultados.items() if v['avaliacoes'] > 0}
    
    if media_atendimentos is None:
        media_atendimentos = calcular_media_operacao(resultados, gestor)
    
    dados_validos = {k: v for k, v in dados_filtrados.items() 
                     if v['csat'] >= limiar_csat and v['total_atendimentos'] >= media_atendimentos}
    
    sorted_analistas = sorted(dados_validos.items(), key=lambda x: x[1]['csat'], reverse=True)
    top_3 = sorted_analistas[:3]
    
    return [(nome, dados['csat'], dados['total_atendimentos'], dados['perc_avaliacoes']) for nome, dados in top_3]

# ============================================
# FUNÇÕES DE VISUALIZAÇÃO - PAINEL DO ANALISTA
# ============================================

def criar_painel_analista(analista, dados, media_operacao, podio):
    """Cria o painel visual do analista com métricas e gráficos"""
    
    posicao_podio = None
    for i, (nome, _, _, _) in enumerate(podio, 1):
        if nome == analista:
            posicao_podio = i
            break
    
    st.subheader(f"📊 Painel de Performance - {analista}")
    
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        delta_csat = dados['delta_csat']
        cor_delta = "green" if delta_csat >= 0 else "red"
        st.markdown(f"""
        <div style="background: #f8f9fa; padding: 15px; border-radius: 10px; text-align: center; border-left: 4px solid {'#28a745' if dados['csat'] >= dados['meta_csat'] else '#dc3545'};">
            <p style="font-size: 12px; color: #666; margin: 0;">⭐ CSAT</p>
            <p style="font-size: 28px; font-weight: bold; margin: 5px 0; color: {'#28a745' if dados['csat'] >= dados['meta_csat'] else '#dc3545'};">{dados['csat']:.1f}%</p>
            <p style="font-size: 12px; color: #666; margin: 0;">Meta: {dados['meta_csat']}%</p>
            <p style="font-size: 14px; color: {cor_delta}; margin: 0;">{delta_csat:+.1f}%</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        diff_avaliacoes = dados['perc_avaliacoes'] - dados['meta_geral']
        cor_diff = "green" if diff_avaliacoes >= 0 else "red"
        st.markdown(f"""
        <div style="background: #f8f9fa; padding: 15px; border-radius: 10px; text-align: center; border-left: 4px solid {'#28a745' if dados['perc_avaliacoes'] >= dados['meta_geral'] else '#dc3545'};">
            <p style="font-size: 12px; color: #666; margin: 0;">📊 % Avaliações</p>
            <p style="font-size: 28px; font-weight: bold; margin: 5px 0; color: {'#28a745' if dados['perc_avaliacoes'] >= dados['meta_geral'] else '#dc3545'};">{dados['perc_avaliacoes']:.1f}%</p>
            <p style="font-size: 12px; color: #666; margin: 0;">Meta: {dados['meta_geral']}%</p>
            <p style="font-size: 14px; color: {cor_diff}; margin: 0;">{diff_avaliacoes:+.1f}%</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div style="background: #f8f9fa; padding: 15px; border-radius: 10px; text-align: center; border-left: 4px solid #17a2b8;">
            <p style="font-size: 12px; color: #666; margin: 0;">📤 % Envio</p>
            <p style="font-size: 28px; font-weight: bold; margin: 5px 0; color: #17a2b8;">{dados['perc_envio']:.1f}%</p>
            <p style="font-size: 12px; color: #666; margin: 0;">Clientes não avaliaram</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        diff_atend = dados['total_atendimentos'] - media_operacao
        cor_atend = "green" if diff_atend >= 0 else "red"
        st.markdown(f"""
        <div style="background: #f8f9fa; padding: 15px; border-radius: 10px; text-align: center; border-left: 4px solid #6c757d;">
            <p style="font-size: 12px; color: #666; margin: 0;">💬 Atendimentos</p>
            <p style="font-size: 28px; font-weight: bold; margin: 5px 0; color: #6c757d;">{dados['total_atendimentos']}</p>
            <p style="font-size: 12px; color: #666; margin: 0;">Média: {media_operacao}</p>
            <p style="font-size: 14px; color: {cor_atend}; margin: 0;">{diff_atend:+.0f}</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col5:
        if posicao_podio:
            medalha = ["🥇", "🥈", "🥉"][posicao_podio-1]
            cor_podio = ['#FFD700', '#C0C0C0', '#CD7F32'][posicao_podio-1]
            st.markdown(f"""
            <div style="background: #f8f9fa; padding: 15px; border-radius: 10px; text-align: center; border-left: 4px solid {cor_podio};">
                <p style="font-size: 12px; color: #666; margin: 0;">🏆 Ranking</p>
                <p style="font-size: 40px; margin: 0;">{medalha}</p>
                <p style="font-size: 16px; font-weight: bold; margin: 5px 0; color: {cor_podio};">{posicao_podio}º Lugar</p>
                <p style="font-size: 12px; color: #666; margin: 0;">Pódio do Mês</p>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div style="background: #f8f9fa; padding: 15px; border-radius: 10px; text-align: center; border-left: 4px solid #6c757d;">
                <p style="font-size: 12px; color: #666; margin: 0;">🏆 Ranking</p>
                <p style="font-size: 28px; margin: 5px 0;">📋</p>
                <p style="font-size: 14px; font-weight: bold; margin: 5px 0; color: #6c757d;">Não está no pódio</p>
                <p style="font-size: 12px; color: #666; margin: 0;">Meta: CSAT ≥ 90%</p>
            </div>
            """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    col1, col2 = st.columns(2)
    
    with col1:
        fig_barras = go.Figure()
        
        fig_barras.add_trace(go.Bar(
            x=['CSAT'],
            y=[dados['csat']],
            name='Alcançado',
            marker_color='#2ecc71',
            text=[f"{dados['csat']:.1f}%"],
            textposition='outside'
        ))
        
        fig_barras.add_trace(go.Bar(
            x=['CSAT'],
            y=[dados['meta_csat']],
            name='Meta',
            marker_color='#e74c3c',
            text=[f"{dados['meta_csat']}%"],
            textposition='outside'
        ))
        
        fig_barras.update_layout(
            title='CSAT - Resultado vs Meta',
            yaxis_title='Percentual (%)',
            yaxis_range=[0, 100],
            height=350,
            showlegend=True,
            legend=dict(orientation='h', yanchor='bottom', y=1.02)
        )
        st.plotly_chart(fig_barras, use_container_width=True)
        
        fig_gauge = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=dados['csat'],
            delta={'reference': dados['meta_csat']},
            title={'text': "CSAT"},
            gauge={
                'axis': {'range': [None, 100]},
                'bar': {'color': "#2ecc71" if dados['csat'] >= dados['meta_csat'] else "#e74c3c"},
                'steps': [
                    {'range': [0, dados['meta_csat']], 'color': "rgba(231, 76, 60, 0.3)"},
                    {'range': [dados['meta_csat'], 100], 'color': "rgba(46, 204, 113, 0.3)"}
                ],
                'threshold': {
                    'line': {'color': "red", 'width': 4},
                    'thickness': 0.75,
                    'value': dados['meta_csat']
                }
            }
        ))
        fig_gauge.update_layout(height=250)
        st.plotly_chart(fig_gauge, use_container_width=True)
    
    with col2:
        categorias = ['CSAT', '% Avaliações', '% Envio', 'Atendimentos']
        
        csat_norm = dados['csat']
        avaliacoes_norm = dados['perc_avaliacoes']
        envio_norm = dados['perc_envio']
        
        if media_operacao > 0:
            atend_norm = min(100, (dados['total_atendimentos'] / media_operacao) * 100)
        else:
            atend_norm = 0
        
        valores_analista = [csat_norm, avaliacoes_norm, envio_norm, atend_norm]
        valores_meta = [dados['meta_csat'], dados['meta_geral'], 50, 100]
        
        fig_radar = go.Figure()
        
        fig_radar.add_trace(go.Scatterpolar(
            r=valores_analista,
            theta=categorias,
            fill='toself',
            name=analista,
            line_color='#2ecc71',
            fillcolor='rgba(46, 204, 113, 0.3)'
        ))
        
        fig_radar.add_trace(go.Scatterpolar(
            r=valores_meta,
            theta=categorias,
            fill='toself',
            name='Meta',
            line_color='#e74c3c',
            fillcolor='rgba(231, 76, 60, 0.1)',
            line_dash='dash'
        ))
        
        fig_radar.update_layout(
            title='Radar de Performance',
            polar=dict(
                radialaxis=dict(
                    visible=True,
                    range=[0, 100]
                )
            ),
            height=400,
            showlegend=True,
            legend=dict(orientation='h', yanchor='bottom', y=1.02)
        )
        st.plotly_chart(fig_radar, use_container_width=True)
    
    st.markdown("---")
    
    return posicao_podio

# ============================================
# GERAÇÃO DE FEEDBACK COM IA
# ============================================

def gerar_feedback_ia(analista, dados, media_operacao, posicao_podio=None, feedback_editado=None):
    """Gera feedback usando IA com padrão MIMO refinado"""
    
    genero = get_genero_neutro(analista)
    
    texto_podio = ""
    if posicao_podio:
        texto_podio = f"🏆 {posicao_podio}º lugar no pódio do mês!"
    
    prompt_base = f"""
Você é um especialista em gestão de performance e desenvolvimento de equipes de atendimento ao cliente.

## DADOS DO COLABORADOR:
- Nome: {analista}
- Gênero: {genero}
- CSAT: {dados['csat']:.2f}% (Meta: ≥ {dados['meta_csat']:.0f}%)
- Delta CSAT: {dados['delta_csat']:+.2f} pontos
- % Avaliações: {dados['perc_avaliacoes']:.2f}% (Meta: ≥ {dados['meta_geral']:.0f}%)
- % Envio: {dados['perc_envio']:.2f}%
- Atendimentos: {dados['total_atendimentos']}
- Média da operação: {media_operacao}
- Avaliações: {dados['avaliacoes']} ({dados['positivos']} positivas, {dados['negativos']} negativas)
- Status: {dados['status']}
- Posição no pódio: {texto_podio if texto_podio else "Não está no pódio"}

## ESTRUTURA OBRIGATÓRIA DO FEEDBACK:
Todo feedback deve seguir esta estrutura:

### 1. CONTEXTO E OBSERVAÇÃO
Descreva de forma objetiva o fato observado, utilizando dados, comportamentos ou situações concretas, sem julgamentos pessoais.

### 2. IMPACTO
Explique quais foram os impactos gerados para o cliente, para a equipe, para os resultados ou para a operação.

### 3. RECONHECIMENTO
Destaque pontos positivos identificados, reforçando comportamentos desejados e demonstrando equilíbrio na análise.

### 4. OPORTUNIDADE DE DESENVOLVIMENTO
Apresente claramente qual comportamento, habilidade ou atitude precisa ser ajustado.

### 5. DIRECIONAMENTO PRÁTICO
Informe ações concretas que o colaborador pode realizar para evoluir.

### 6. ENCERRAMENTO MOTIVADOR
Finalize reforçando confiança na capacidade de desenvolvimento do colaborador.

## DIRETRIZES OBRIGATÓRIAS:
- Use Comunicação Não Violenta (CNV)
- Evite linguagem acusatória, punitiva ou agressiva
- Fale sobre comportamentos observáveis, nunca sobre características pessoais
- Use linguagem respeitosa, empática e profissional
- Demonstre interesse genuíno pelo desenvolvimento do colaborador
- Mantenha equilíbrio entre reconhecimento e desenvolvimento
- Use linguagem orientada para soluções e evolução
- Direcione o foco para comportamentos futuros desejados

## PERGUNTAS QUE O FEEDBACK DEVE RESPONDER:
1. O que aconteceu? (Contexto)
2. Qual foi o impacto? (Impacto)
3. O que deve continuar sendo feito? (Reconhecimento)
4. O que precisa mudar? (Oportunidade)
5. Como essa mudança pode acontecer na prática? (Direcionamento)
6. Qual o benefício esperado? (Encerramento)
"""
    
    if feedback_editado:
        prompt = prompt_base + f"""

## FEEDBACK ATUAL (EDITADO PELO GESTOR):
{feedback_editado}

## INSTRUÇÃO:
Analise o feedback acima e melhore/ajuste seguindo:
1. Mantenha a essência do que o gestor escreveu
2. Reorganize para seguir a estrutura obrigatória (6 seções)
3. Aplique as diretrizes de comunicação (CNV, PNL)
4. Mantenha o tom profissional e motivador
5. Enriqueça com os dados do colaborador quando necessário
6. Responda às 6 perguntas obrigatórias

Gere o feedback revisado e aprimorado:
"""
    else:
        prompt = prompt_base + f"""

## INSTRUÇÃO:
Gere um feedback completo de performance seguindo a estrutura obrigatória de 6 seções.
Aplique todas as diretrizes de comunicação e responda às 6 perguntas obrigatórias.

O feedback deve ser profissional, construtivo e motivador, com base nos dados fornecidos.

Gere o feedback:
"""
    
    try:
        github_token = st.secrets.get("GITHUB_TOKEN", os.environ.get("GITHUB_TOKEN", ""))
        
        if github_token:
            headers = {
                "Authorization": f"Bearer {github_token}",
                "Content-Type": "application/json"
            }
            
            url = "https://models.inference.ai.azure.com/chat/completions"
            
            payload = {
                "model": "gpt-4o-mini",
                "messages": [
                    {
                        "role": "system",
                        "content": "Você é um especialista em gestão de performance e desenvolvimento de equipes de atendimento ao cliente."
                    },
                    {
                        "role": "user",
                        "content": prompt
                    }
                ],
                "temperature": 0.7,
                "max_tokens": 1500
            }
            
            response = requests.post(url, headers=headers, json=payload)
            
            if response.status_code == 200:
                result = response.json()
                return result['choices'][0]['message']['content'].strip()
        
        return gerar_feedback_manual(analista, dados, media_operacao, posicao_podio)
        
    except Exception as e:
        st.warning(f"Erro ao gerar feedback com IA: {str(e)}. Usando feedback manual.")
        return gerar_feedback_manual(analista, dados, media_operacao, posicao_podio)

def gerar_feedback_manual(analista, dados, media_operacao, posicao_podio=None):
    """Gera feedback manual (fallback sem IA)"""
    
    genero = get_genero_neutro(analista)
    
    texto_podio = ""
    if posicao_podio:
        texto_podio = f"🏆 {posicao_podio}º lugar no pódio do mês!"
    
    if dados['status'] == "🟢 Meta Superada":
        status_texto = "excelente, superando todas as metas estabelecidas"
        reconhecimento = "Parabéns pelo excelente desempenho! Você tem sido um exemplo para a equipe."
        direcionamento = "Continue mantendo este alto padrão de qualidade e compartilhe suas boas práticas com os colegas."
    elif dados['status'] == "🟡 Atenção":
        status_texto = "bom, com alguns pontos que merecem atenção"
        reconhecimento = "Seu desempenho tem sido consistente, mas há espaço para crescimento."
        direcionamento = "Foque em melhorar os indicadores que estão abaixo da meta. Conte com o suporte da gestão."
    else:
        status_texto = "precisa de atenção e melhoria"
        reconhecimento = "Reconhecemos seu esforço, mas é necessário um plano de ação para melhorar os resultados."
        direcionamento = "Vamos trabalhar juntos em um plano de ação focado nas áreas que precisam de desenvolvimento."
    
    feedback = f"""
### 1. CONTEXTO E OBSERVAÇÃO

No período analisado, o(a) colaborador(a) {analista} apresentou um CSAT de {dados['csat']:.2f}%, 
com uma taxa de avaliações de {dados['perc_avaliacoes']:.2f}% e {dados['total_atendimentos']} atendimentos realizados. 
O percentual de envio foi de {dados['perc_envio']:.2f}%. {texto_podio}

### 2. IMPACTO

O desempenho do(a) colaborador(a) impactou diretamente a experiência do cliente e os resultados da operação. 
Com {dados['positivos']} avaliações positivas e apenas {dados['negativos']} negativas, o índice de satisfação 
demonstra {'alta qualidade' if dados['csat'] >= 90 else 'oportunidades de melhoria'} no atendimento prestado.

### 3. RECONHECIMENTO

{reconhecimento}
- CSAT {'acima' if dados['delta_csat'] >= 0 else 'próximo'} da meta estabelecida
- {'Boa' if dados['perc_avaliacoes'] >= dados['meta_geral'] else 'Atenção à'} taxa de coleta de feedback
- {'Produtividade' if dados['total_atendimentos'] >= media_operacao else 'Potencial para'} {'acima' if dados['total_atendimentos'] >= media_operacao else 'aumentar'} da média da operação

### 4. OPORTUNIDADE DE DESENVOLVIMENTO

Para evoluir ainda mais, sugere-se focar em:
- {'Manter o alto padrão de CSAT e buscar excelência contínua' if dados['csat'] >= 92 else 'Elevar o CSAT através de atendimentos mais resolutivos e personalizados'}
- {'Continuar engajando os clientes na coleta de feedback' if dados['perc_avaliacoes'] >= dados['meta_geral'] else 'Aumentar a taxa de coleta de avaliações, oferecendo a pesquisa ao final de cada atendimento'}
- {'Manter a produtividade elevada' if dados['total_atendimentos'] >= media_operacao else 'Otimizar o tempo entre atendimentos para aumentar a produtividade'}

### 5. DIRECIONAMENTO PRÁTICO

Ações recomendadas:
1. {'Compartilhar boas práticas com a equipe' if dados['csat'] >= 92 else 'Revisar atendimentos com avaliações negativas para identificar padrões de melhoria'}
2. {'Manter a abordagem atual de coleta de feedback' if dados['perc_avaliacoes'] >= dados['meta_geral'] else 'Implementar uma rotina de oferecimento da pesquisa de satisfação em todos os atendimentos'}
3. Agendar uma conversa de alinhamento com a gestão para acompanhamento do desenvolvimento

### 6. ENCERRAMENTO MOTIVADOR

Acreditamos no seu potencial de crescimento e evolução contínua! 
O desenvolvimento individual fortalece todo o time, e sua dedicação é fundamental para alcançarmos 
resultados ainda melhores. Continue se dedicando e contando com o apoio da gestão para 
superar os desafios e celebrar as conquistas! 🚀

---
**Status:** {dados['status']}
**Data:** {datetime.now().strftime('%d/%m/%Y')}
"""
    return feedback.strip()

# ============================================
# GERADOR DE GRÁFICO MENSAL
# ============================================

def gerar_grafico_mensal(analista, dados_mensais, meta_csat, meta_avaliacoes):
    """Gera gráfico mensal de evolução do analista (apenas se tiver 2+ meses)"""
    
    if dados_mensais is None or dados_mensais.empty:
        return None
    
    df_analista = dados_mensais[dados_mensais['analista'] == analista]
    
    if df_analista.empty:
        return None
    
    meses = df_analista['mes_ano'].nunique()
    if meses < 2:
        return None
    
    df_analista = df_analista.sort_values('mes_ano')
    
    fig = px.line(
        df_analista,
        x='mes_ano',
        y='csat',
        title=f'📈 Evolução Mensal - {analista}',
        labels={'mes_ano': 'Período', 'csat': 'CSAT (%)'},
        markers=True
    )
    
    fig.update_traces(
        line=dict(color='#2ecc71', width=3),
        marker=dict(size=10, color='#2ecc71'),
        name='CSAT Alcançado'
    )
    
    fig.add_trace(go.Bar(
        x=df_analista['mes_ano'],
        y=df_analista['perc_avaliacoes'],
        name='% Avaliações',
        marker_color='#3498db',
        opacity=0.6,
        yaxis='y2'
    ))
    
    meta_csat_data = [meta_csat] * len(df_analista)
    fig.add_trace(go.Scatter(
        x=df_analista['mes_ano'],
        y=meta_csat_data,
        name=f'Meta CSAT: {meta_csat}%',
        line=dict(color='#e74c3c', width=2, dash='dash'),
        mode='lines'
    ))
    
    meta_avaliacoes_data = [meta_avaliacoes] * len(df_analista)
    fig.add_trace(go.Scatter(
        x=df_analista['mes_ano'],
        y=meta_avaliacoes_data,
        name=f'Meta Avaliações: {meta_avaliacoes}%',
        line=dict(color='#f39c12', width=2, dash='dot'),
        mode='lines',
        yaxis='y2'
    ))
    
    fig.update_layout(
        xaxis_title='Período',
        yaxis=dict(
            title='CSAT (%)',
            range=[0, 100],
            side='left'
        ),
        yaxis2=dict(
            title='% Avaliações',
            range=[0, 100],
            overlaying='y',
            side='right'
        ),
        hovermode='x unified',
        legend=dict(
            orientation='h',
            yanchor='bottom',
            y=1.02,
            xanchor='right',
            x=1
        ),
        height=450,
        template='plotly_white'
    )
    
    return fig

# ============================================
# GERAÇÃO DE ANÁLISE TÉCNICA
# ============================================

def gerar_analise_tecnica(analista, dados, media_operacao, podio):
    """Gera análise técnica detalhada"""
    genero = get_genero_neutro(analista)
    
    posicao = "não está no pódio"
    for i, (nome, csat, atendimentos, perc_avaliacoes) in enumerate(podio, 1):
        if nome == analista:
            posicao = f"está em {i}º lugar no pódio com CSAT de {csat:.2f}%, {atendimentos} atendimentos e {perc_avaliacoes:.2f}% de avaliações"
            break
    
    if dados['delta_csat'] > 0:
        analise_csat = f"superou a meta em {dados['delta_csat']:.2f} pontos percentuais"
    elif dados['delta_csat'] < 0:
        analise_csat = f"ficou abaixo da meta em {abs(dados['delta_csat']):.2f} pontos percentuais"
    else:
        analise_csat = "atingiu exatamente a meta"
    
    diff_avaliacoes = dados['perc_avaliacoes'] - dados['meta_geral']
    if diff_avaliacoes >= 0:
        analise_avaliacoes = f"superou a meta de {dados['meta_geral']:.0f}% em {diff_avaliacoes:.2f} pontos percentuais"
    else:
        analise_avaliacoes = f"ficou abaixo da meta de {dados['meta_geral']:.0f}% em {abs(diff_avaliacoes):.2f} pontos percentuais"
    
    if dados['total_atendimentos'] > media_operacao:
        produtividade = f"{(dados['total_atendimentos'] / media_operacao - 1) * 100:.2f}% superior à média da operação"
    elif dados['total_atendimentos'] < media_operacao:
        produtividade = f"{(1 - dados['total_atendimentos'] / media_operacao) * 100:.2f}% inferior à média da operação"
    else:
        produtividade = "igual à média da operação"
    
    if dados['csat'] >= 94:
        nivel_csat = "extremamente alto"
    elif dados['csat'] >= 90:
        nivel_csat = "alto"
    elif dados['csat'] >= 85:
        nivel_csat = "bom"
    else:
        nivel_csat = "precisa de atenção"
    
    texto = f"""
### 1. Qualidade e Satisfação do Cliente (CSAT)

O(A) {genero} registrou um índice de **Satisfação (CSAT) de {dados['csat']:.2f}%**.

- **Comparativo com a Meta:** O resultado {analise_csat} (meta: ≥ {dados['meta_csat']:.0f}%).
- **Análise Detalhada:** Do volume total de feedbacks recebidos ({dados['avaliacoes']}), **{dados['positivos']} foram positivos**, resultando em um índice de aprovação {nivel_csat}. Houve apenas {dados['negativos']} registros negativos, demonstrando consistência na entrega de um atendimento cordial, eficiente e resolutivo.

### 2. Engajamento e Coleta de Feedback

O(A) {genero} alcançou uma **Taxa de Avaliações de {dados['perc_avaliacoes']:.2f}%**.

- **Comparativo com a Meta:** A meta esperada é de no mínimo {dados['meta_geral']:.0f}% de conversão de atendimentos em avaliações. O resultado {analise_avaliacoes}, refletindo uma {'proatividade exemplar' if dados['perc_avaliacoes'] >= dados['meta_geral'] else 'oportunidade de melhoria'} na solicitação de feedback ao final de cada contato.

- **Cálculo:** A taxa foi calculada sobre o volume total de {dados['avaliacoes']} avaliações divididas pelos **{dados['validos']} atendimentos válidos**.

### 3. Produtividade e Volumetria

O volume total de atendimentos realizados pelo(a) {genero} foi de **{dados['total_atendimentos']} chamados**.

- **Comparativo com a Operação:** A média de atendimentos por agente da operação no período foi de {media_operacao}. {analista} absorveu uma demanda operacional **{produtividade}**.

- **Destaque:** {posicao} do mês em CSAT.
"""
    return texto.strip()

# ============================================
# GERAÇÃO DE RELATÓRIO WORD
# ============================================

def gerar_relatorio_word(analista, dados, analise_tecnica, feedback, media_operacao, podio, periodo, fig_grafico=None):
    """Gera relatório em formato Word (.docx) com gráfico"""
    
    doc = Document()
    
    titulo = doc.add_heading(f'Relatório de Performance - {analista}', 0)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph(f'Período: {periodo}')
    doc.add_paragraph('')
    
    doc.add_heading('Esperado:', level=1)
    doc.add_paragraph(f'≥ {dados["meta_geral"]:.0f}% de avaliações', style='List Bullet')
    doc.add_paragraph(f'≥ {dados["meta_csat"]:.0f}% de Satisfação', style='List Bullet')
    doc.add_paragraph('')
    
    doc.add_heading('Atingido:', level=1)
    
    posicao_texto = "Não está no pódio"
    for i, (nome, csat, atendimentos, perc_avaliacoes) in enumerate(podio, 1):
        if nome == analista:
            posicao_texto = f"{i}º Lugar - CSAT: {csat:.2f}% | {atendimentos} atendimentos | {perc_avaliacoes:.2f}% avaliações"
            break
    
    doc.add_paragraph(f'CSAT: {dados["csat"]:.2f}%', style='List Bullet')
    doc.add_paragraph(f'Avaliações: {dados["perc_avaliacoes"]:.2f}% ({dados["positivos"]} positivos + {dados["negativos"]} negativos = {dados["avaliacoes"]})', style='List Bullet')
    doc.add_paragraph(f'% Envio: {dados["perc_envio"]:.2f}%', style='List Bullet')
    doc.add_paragraph(f'Atendidos: {dados["total_atendimentos"]} - {dados["total_inativos"]} = {dados["validos"]}', style='List Bullet')
    doc.add_paragraph(f'Média por agente: {media_operacao}', style='List Bullet')
    doc.add_paragraph(f'Posição no Pódio: {posicao_texto}', style='List Bullet')
    doc.add_paragraph('')
    
    if fig_grafico:
        doc.add_heading('Gráfico de Evolução Mensal', level=1)
        doc.add_paragraph('(Gráfico disponível na versão digital do relatório)')
        doc.add_paragraph('')
    
    doc.add_heading('Análise Técnica de Desempenho', level=1)
    
    analise_limpa = analise_tecnica.replace('###', '')
    analise_limpa = analise_limpa.replace('**', '')
    
    for linha in analise_limpa.split('\n'):
        if linha.strip():
            if linha.strip().startswith('1.'):
                doc.add_heading('Qualidade e Satisfação do Cliente (CSAT)', level=2)
            elif linha.strip().startswith('2.'):
                doc.add_heading('Engajamento e Coleta de Feedback', level=2)
            elif linha.strip().startswith('3.'):
                doc.add_heading('Produtividade e Volumetria', level=2)
            else:
                doc.add_paragraph(linha.strip())
    
    doc.add_paragraph('')
    
    doc.add_heading('Feedback de Performance', level=1)
    doc.add_paragraph(f'Status Geral do Período: {dados["status"]}')
    doc.add_paragraph('')
    
    for linha in feedback.split('\n'):
        if linha.strip():
            if linha.strip().startswith('#'):
                continue
            elif linha.strip().startswith('###'):
                doc.add_heading(linha.strip().replace('###', '').strip(), level=2)
            elif linha.strip().startswith('**'):
                doc.add_paragraph(linha.strip())
            else:
                doc.add_paragraph(linha.strip())
    
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ============================================
# GERENCIAR ANALISTAS
# ============================================

def gerenciar_analistas_completo(analistas_config):
    """Função para gerenciar analistas com opção de mover entre times"""
    
    st.header("📝 Gerenciar Analistas")
    
    gestor_selecionado = st.selectbox(
        "Selecione o Gestor para editar",
        list(analistas_config.keys())
    )
    
    if not gestor_selecionado:
        return
    
    config = analistas_config[gestor_selecionado]
    
    nova_meta = st.number_input(
        "Meta Geral de Avaliações (%)",
        min_value=0,
        max_value=100,
        value=config['meta_geral_avaliacoes']
    )
    if nova_meta != config['meta_geral_avaliacoes']:
        config['meta_geral_avaliacoes'] = nova_meta
        salvar_analistas(analistas_config)
    
    st.markdown("---")
    
    st.subheader("🔄 Mover Analista entre Times")
    
    todos_analistas = []
    for gestor, cfg in analistas_config.items():
        for analista, dados in cfg['membros'].items():
            if dados['ativo']:
                todos_analistas.append((analista, gestor))
    
    if todos_analistas:
        col1, col2, col3 = st.columns([2, 2, 1])
        
        with col1:
            analista_para_mover = st.selectbox(
                "Selecione o analista para mover",
                [a[0] for a in todos_analistas]
            )
        
        with col2:
            gestor_atual = next((g for a, g in todos_analistas if a == analista_para_mover), None)
            todos_gestores = list(analistas_config.keys())
            indice_atual = todos_gestores.index(gestor_atual) if gestor_atual in todos_gestores else 0
            
            novo_gestor = st.selectbox(
                f"Time atual: {gestor_atual}",
                todos_gestores,
                index=indice_atual
            )
        
        with col3:
            if novo_gestor and novo_gestor != gestor_atual:
                if st.button("🔄 Mover", use_container_width=True):
                    meta_atual = analistas_config[gestor_atual]['membros'][analista_para_mover]['meta_csat']
                    del analistas_config[gestor_atual]['membros'][analista_para_mover]
                    analistas_config[novo_gestor]['membros'][analista_para_mover] = {
                        "meta_csat": meta_atual,
                        "ativo": True
                    }
                    salvar_analistas(analistas_config)
                    st.success(f"✅ {analista_para_mover} movido para {novo_gestor}!")
                    st.rerun()
            elif novo_gestor == gestor_atual:
                st.info("ℹ️ O analista já está neste time.")
    else:
        st.info("Nenhum analista ativo para mover.")
    
    st.markdown("---")
    
    st.subheader(f"👥 Membros do Time: {gestor_selecionado}")
    
    membros = config['membros']
    
    col1, col2, col3 = st.columns([3, 1, 1])
    with col1:
        novo_nome = st.text_input("Nome do novo analista")
    with col2:
        nova_meta_csat = st.number_input("Meta CSAT", min_value=0, max_value=100, value=86)
    with col3:
        if st.button("➕ Adicionar") and novo_nome:
            membros[novo_nome] = {"meta_csat": nova_meta_csat, "ativo": True}
            salvar_analistas(analistas_config)
            st.rerun()
    
    st.markdown("---")
    
    dados_tabela = []
    for nome, dados in membros.items():
        dados_tabela.append({
            "Analista": nome,
            "Meta CSAT": f"{dados['meta_csat']}%",
            "Status": "✅ Ativo" if dados['ativo'] else "❌ Inativo"
        })
    
    if dados_tabela:
        df_membros = pd.DataFrame(dados_tabela)
        st.dataframe(df_membros, use_container_width=True, hide_index=True)
    
    st.subheader("✏️ Editar/Remover Membros")
    membro_selecionado = st.selectbox(
        "Selecione um membro para editar",
        list(membros.keys())
    )
    
    if membro_selecionado:
        col1, col2, col3 = st.columns([2, 1, 1])
        with col1:
            nova_meta_membro = st.number_input(
                "Nova Meta CSAT",
                min_value=0,
                max_value=100,
                value=membros[membro_selecionado]['meta_csat']
            )
        with col2:
            novo_status = st.checkbox(
                "Ativo",
                value=membros[membro_selecionado]['ativo']
            )
        with col3:
            if st.button("🗑️ Remover", use_container_width=True):
                del membros[membro_selecionado]
                salvar_analistas(analistas_config)
                st.rerun()
        
        if nova_meta_membro != membros[membro_selecionado]['meta_csat'] or novo_status != membros[membro_selecionado]['ativo']:
            membros[membro_selecionado]['meta_csat'] = nova_meta_membro
            membros[membro_selecionado]['ativo'] = novo_status
            salvar_analistas(analistas_config)
            st.success("✅ Alterações salvas!")
            st.rerun()

# ============================================
# GRÁFICOS DO DASHBOARD
# ============================================

def criar_graficos_dashboard(df_dashboard):
    """Cria todos os gráficos do dashboard"""
    
    if df_dashboard.empty:
        st.warning("Nenhum dado disponível para gráficos.")
        return
    
    fig_csat = px.bar(
        df_dashboard,
        x='Analista',
        y='CSAT',
        color='CSAT',
        color_continuous_scale=['red', 'yellow', 'green'],
        range_color=[0, 100],
        title='📊 CSAT por Analista',
        text=df_dashboard['CSAT'].apply(lambda x: f'{x:.1f}%')
    )
    fig_csat.update_traces(textposition='outside')
    fig_csat.update_layout(height=400, showlegend=False)
    st.plotly_chart(fig_csat, use_container_width=True)
    
    fig_avaliacoes = px.bar(
        df_dashboard,
        x='Analista',
        y='% Avaliações',
        color='% Avaliações',
        color_continuous_scale=['red', 'yellow', 'green'],
        range_color=[0, 50],
        title='📈 % Avaliações por Analista',
        text=df_dashboard['% Avaliações'].apply(lambda x: f'{x:.1f}%')
    )
    fig_avaliacoes.update_traces(textposition='outside')
    fig_avaliacoes.update_layout(height=400, showlegend=False)
    st.plotly_chart(fig_avaliacoes, use_container_width=True)
    
    fig_envio = px.bar(
        df_dashboard,
        x='Analista',
        y='% Envio',
        color='% Envio',
        color_continuous_scale=['green', 'yellow', 'red'],
        range_color=[0, 100],
        title='📤 % Envio por Analista',
        text=df_dashboard['% Envio'].apply(lambda x: f'{x:.1f}%')
    )
    fig_envio.update_traces(textposition='outside')
    fig_envio.update_layout(height=400, showlegend=False)
    st.plotly_chart(fig_envio, use_container_width=True)
    
    fig_scatter = px.scatter(
        df_dashboard,
        x='💬 Atendimentos',
        y='CSAT',
        size='💬 Atendimentos',
        color='Status',
        hover_data=['Analista', '% Avaliações', '% Envio'],
        title='🎯 CSAT vs 💬 Atendimentos',
        labels={'💬 Atendimentos': '💬 Quantidade de Atendimentos', 'CSAT': 'CSAT (%)'}
    )
    fig_scatter.update_layout(height=400)
    st.plotly_chart(fig_scatter, use_container_width=True)
    
    st.subheader("📊 Distribuição de Status - Métricas Agregadas")
    
    status_agg = df_dashboard.groupby('Status').agg({
        'Analista': 'count',
        '% Avaliações': 'mean',
        '% Envio': 'mean',
        'CSAT': 'mean'
    }).reset_index()
    
    status_agg.columns = ['Status', 'Quantidade', '% Avaliações Médio', '% Envio Médio', 'CSAT Médio']
    status_agg['%'] = (status_agg['Quantidade'] / len(df_dashboard) * 100).round(1)
    
    fig_stacked = go.Figure()
    
    fig_stacked.add_trace(go.Bar(
        name='% Avaliações',
        x=status_agg['Status'],
        y=status_agg['% Avaliações Médio'],
        text=status_agg['% Avaliações Médio'].apply(lambda x: f'{x:.1f}%'),
        textposition='outside',
        marker_color='#2ecc71'
    ))
    
    fig_stacked.add_trace(go.Bar(
        name='% Envio',
        x=status_agg['Status'],
        y=status_agg['% Envio Médio'],
        text=status_agg['% Envio Médio'].apply(lambda x: f'{x:.1f}%'),
        textposition='outside',
        marker_color='#e74c3c'
    ))
    
    fig_stacked.add_trace(go.Bar(
        name='CSAT Médio',
        x=status_agg['Status'],
        y=status_agg['CSAT Médio'],
        text=status_agg['CSAT Médio'].apply(lambda x: f'{x:.1f}%'),
        textposition='outside',
        marker_color='#3498db'
    ))
    
    fig_stacked.update_layout(
        title='📊 Distribuição por Status - CSAT, % Avaliações e % Envio',
        barmode='group',
        height=450,
        yaxis_title='Percentual (%)',
        legend_title='Métrica'
    )
    
    st.plotly_chart(fig_stacked, use_container_width=True)
    
    st.dataframe(
        status_agg[['Status', 'Quantidade', '%', 'CSAT Médio', '% Avaliações Médio', '% Envio Médio']],
        use_container_width=True,
        hide_index=True
    )
    
    col1, col2 = st.columns(2)
    
    with col1:
        status_counts = df_dashboard['Status'].value_counts()
        fig_pie = px.pie(
            values=status_counts.values,
            names=status_counts.index,
            title='📊 Distribuição de Status',
            color_discrete_map={
                '🟢 Meta Superada': '#28a745',
                '🟡 Atenção': '#ffc107',
                '🔴 Crítico': '#dc3545'
            }
        )
        fig_pie.update_traces(textposition='inside', textinfo='percent+label')
        fig_pie.update_layout(height=400)
        st.plotly_chart(fig_pie, use_container_width=True)
    
    with col2:
        status_detalhe = df_dashboard.groupby('Status').agg({
            'Analista': 'count',
            'CSAT': 'mean',
            '% Avaliações': 'mean',
            '% Envio': 'mean'
        }).reset_index()
        status_detalhe.columns = ['Status', 'Qtd', 'CSAT Médio', '% Avaliações Médio', '% Envio Médio']
        status_detalhe['CSAT Médio'] = status_detalhe['CSAT Médio'].round(2)
        status_detalhe['% Avaliações Médio'] = status_detalhe['% Avaliações Médio'].round(2)
        status_detalhe['% Envio Médio'] = status_detalhe['% Envio Médio'].round(2)
        
        st.dataframe(status_detalhe, use_container_width=True, hide_index=True)

# ============================================
# INTERFACE PRINCIPAL
# ============================================

def main():
    st.title("📊 Sistema de Performance - Relatórios Automáticos")
    st.markdown("---")
    
    # ===== LOGIN =====
    if not fazer_login():
        st.info("👋 Faça login na barra lateral para acessar o sistema.")
        return
    
    # ===== CADASTRO DE USUÁRIO =====
    if st.session_state.get('cadastrar_usuario', False):
        cadastrar_usuario()
        return
    
    # ===== INICIALIZAR SUPABASE =====
    supabase = init_supabase()
    if supabase:
        st.sidebar.success("✅ Conectado ao Supabase")
    else:
        st.sidebar.warning("⚠️ Supabase não configurado")
    
    # ===== CARREGAR DADOS =====
    analistas_config = carregar_analistas()
    
    # ===== SIDEBAR =====
    with st.sidebar:
        st.header("📁 Upload de Arquivos")
        
        arquivo_satisfacao = st.file_uploader(
            "Arquivo de Satisfação (Good vs Bad)",
            type=['xlsx', 'csv']
        )
        
        arquivo_inativos = st.file_uploader(
            "Arquivo de Inatividade",
            type=['xlsx', 'csv']
        )
        
        st.markdown("---")
        
        # Período
        st.header("📅 Período")
        
        opcao_periodo = st.radio(
            "Como definir o período?",
            ["Selecionar manualmente", "Extrair do arquivo"],
            index=0
        )
        
        if opcao_periodo == "Selecionar manualmente":
            periodo_manual = st.text_input(
                "Informe o período",
                placeholder="Ex: Maio 2026",
                value=""
            )
            periodo_selecionado = periodo_manual if periodo_manual else None
        else:
            periodo_manual = None
            periodo_selecionado = None
            st.info("ℹ️ O período será extraído automaticamente do arquivo após o upload.")
        
        periodo_valido = periodo_selecionado is not None and periodo_selecionado.strip() != ""
        
        st.markdown("---")
        
        # Botão Processar
        if st.button("🚀 Processar Dados", use_container_width=True):
            # Validar período
            if not periodo_valido and opcao_periodo == "Selecionar manualmente":
                st.error("⚠️ Selecione o período antes de processar o relatório.")
            elif not arquivo_satisfacao or not arquivo_inativos:
                st.error("❌ Envie os dois arquivos.")
            else:
                with st.spinner("Processando dados..."):
                    try:
                        # Carregar arquivos de forma robusta
                        df_satisfacao = carregar_arquivo_satisfacao(arquivo_satisfacao)
                        if df_satisfacao is None:
                            st.stop()
                        
                        df_inativos = carregar_arquivo_inativos(arquivo_inativos)
                        if df_inativos is None:
                            st.stop()
                        
                        # Extrair período
                        if opcao_periodo == "Extrair do arquivo":
                            periodo = extrair_periodo(df_satisfacao)
                        else:
                            periodo = periodo_manual
                        
                        st.session_state.periodo = periodo
                        
                        # Verificar período existente
                        gestor_logado = st.session_state.gestor
                        if supabase:
                            existe = verificar_periodo_existente(supabase, periodo, gestor_logado)
                            if existe:
                                st.warning(f"⚠️ Já existe um relatório importado para o período {periodo}.")
                                
                                col1, col2 = st.columns(2)
                                with col1:
                                    if st.button("❌ Cancelar importação", use_container_width=True):
                                        st.stop()
                                with col2:
                                    if st.button("🔄 Substituir período existente", use_container_width=True):
                                        resultados = processar_dados(df_satisfacao, df_inativos, analistas_config)
                                        st.session_state.resultados = resultados
                                        st.session_state.processado = True
                                        
                                        sucesso, mensagem = substituir_historico(supabase, resultados, periodo, gestor_logado)
                                        if sucesso:
                                            st.success(f"✅ Período {periodo} substituído com sucesso!")
                                        else:
                                            st.error(f"❌ Erro ao substituir: {mensagem}")
                                        st.rerun()
                                st.stop()
                        
                        # Processar normalmente
                        resultados = processar_dados(df_satisfacao, df_inativos, analistas_config)
                        st.session_state.resultados = resultados
                        st.session_state.processado = True
                        
                        if supabase:
                            sucesso, mensagem = salvar_historico(supabase, resultados, periodo, gestor_logado)
                            if sucesso:
                                st.success(f"✅ Dados salvos no Supabase! Período: {periodo}")
                            else:
                                st.warning(f"⚠️ Dados processados, mas NÃO salvos: {mensagem}")
                        else:
                            st.success(f"✅ Dados processados! Período: {periodo}")
                    except Exception as e:
                        st.error(f"❌ Erro ao processar dados: {str(e)}")
                        st.info("Verifique se os arquivos estão no formato correto.")
        
        st.markdown("---")
        
        # Gerenciar Períodos
        st.header("📂 Gerenciar Períodos")
        
        if st.button("📊 Ver Períodos Importados", use_container_width=True):
            st.session_state.mostrar_periodos = True
        
        st.markdown("---")
        
        # Consultar Histórico
        st.header("📊 Consultar Histórico")
        
        if st.button("📊 Ver Histórico", use_container_width=True):
            st.session_state.mostrar_historico = True
        
        st.markdown("---")
        st.header("⚙️ Gerenciar")
        
        if st.button("📝 Gerenciar Analistas", use_container_width=True):
            st.session_state.gerenciar_analistas = True
    
    # ===== GERENCIAR ANALISTAS =====
    if st.session_state.get('gerenciar_analistas', False):
        gerenciar_analistas_completo(analistas_config)
        
        if st.button("🔙 Voltar", use_container_width=True):
            st.session_state.gerenciar_analistas = False
            st.rerun()
        
        st.markdown("---")
    
    # ===== GERENCIAR PERÍODOS =====
    if st.session_state.get('mostrar_periodos', False):
        st.header("📂 Gerenciar Períodos Importados")
        
        if supabase:
            if st.session_state.get('acesso_total', False):
                periodos = listar_periodos(supabase)
            else:
                periodos = listar_periodos(supabase, st.session_state.gestor)
            
            if periodos:
                dados_tabela = []
                for p in periodos:
                    df_periodo = carregar_historico(supabase, mes_ano=p['mes_ano'], gestor=p['gestor'])
                    qtd_analistas = len(df_periodo) if df_periodo is not None else 0
                    
                    dados_tabela.append({
                        'Período': p['mes_ano'],
                        'Gestor': p['gestor'],
                        'Analistas': qtd_analistas,
                        'Data Importação': p.get('data_criacao', 'N/A')[:10] if p.get('data_criacao') else 'N/A'
                    })
                
                df_periodos = pd.DataFrame(dados_tabela)
                st.dataframe(df_periodos, use_container_width=True, hide_index=True)
                
                st.markdown("---")
                
                st.subheader("🗑️ Excluir Período")
                
                col1, col2 = st.columns([2, 1])
                with col1:
                    periodo_para_excluir = st.selectbox(
                        "Selecione o período para excluir",
                        [p['mes_ano'] for p in periodos]
                    )
                
                with col2:
                    confirmar = st.checkbox("✅ Confirmar exclusão deste período")
                
                if st.button("🗑️ Excluir Período", use_container_width=True):
                    if not confirmar:
                        st.error("⚠️ Marque a caixa de confirmação antes de excluir.")
                    else:
                        gestor_periodo = next((p['gestor'] for p in periodos if p['mes_ano'] == periodo_para_excluir), None)
                        
                        if gestor_periodo:
                            if not st.session_state.get('acesso_total', False) and gestor_periodo != st.session_state.gestor:
                                st.error("❌ Você não tem permissão para excluir este período.")
                            else:
                                with st.spinner(f"Excluindo período {periodo_para_excluir}..."):
                                    sucesso, mensagem = excluir_periodo(supabase, periodo_para_excluir, gestor_periodo)
                                    if sucesso:
                                        st.success(f"✅ {mensagem}")
                                        st.rerun()
                                    else:
                                        st.error(f"❌ Erro: {mensagem}")
            else:
                st.info("Nenhum período importado encontrado.")
        else:
            st.error("❌ Supabase não configurado.")
        
        if st.button("🔙 Voltar"):
            st.session_state.mostrar_periodos = False
            st.rerun()
        
        st.markdown("---")
    
    # ===== CONSULTAR HISTÓRICO =====
    if st.session_state.get('mostrar_historico', False):
        st.header("📊 Consultar Histórico")
        
        if supabase:
            df_historico = carregar_historico(supabase)
            
            if df_historico is not None and not df_historico.empty:
                col1, col2, col3 = st.columns(3)
                with col1:
                    meses_disponiveis = ['Todos'] + sorted(df_historico['mes_ano'].unique().tolist())
                    mes_selecionado = st.selectbox("Selecione o Mês", meses_disponiveis)
                with col2:
                    gestores_disponiveis = ['Todos'] + sorted(df_historico['gestor'].unique().tolist())
                    gestor_filtro = st.selectbox("Selecione o Gestor", gestores_disponiveis)
                with col3:
                    analistas_historicos = ['Todos'] + sorted(df_historico['analista'].unique().tolist())
                    analista_filtro = st.selectbox("Selecione o Analista", analistas_historicos)
                
                df_filtrado = df_historico.copy()
                if mes_selecionado != 'Todos':
                    df_filtrado = df_filtrado[df_filtrado['mes_ano'] == mes_selecionado]
                if gestor_filtro != 'Todos':
                    df_filtrado = df_filtrado[df_filtrado['gestor'] == gestor_filtro]
                if analista_filtro != 'Todos':
                    df_filtrado = df_filtrado[df_filtrado['analista'] == analista_filtro]
                
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total Registros", len(df_filtrado))
                with col2:
                    st.metric("CSAT Médio", f"{df_filtrado['csat'].mean():.2f}%")
                with col3:
                    st.metric("% Avaliações Médio", f"{df_filtrado['perc_avaliacoes'].mean():.2f}%")
                with col4:
                    st.metric("% Envio Médio", f"{df_filtrado['perc_envio'].mean():.2f}%")
                
                st.markdown("---")
                
                if not df_filtrado.empty:
                    fig_historico = px.line(
                        df_filtrado.groupby('mes_ano').agg({
                            'csat': 'mean',
                            'perc_avaliacoes': 'mean',
                            'perc_envio': 'mean'
                        }).reset_index(),
                        x='mes_ano',
                        y=['csat', 'perc_avaliacoes', 'perc_envio'],
                        title='📈 Evolução Mensal',
                        labels={'value': 'Percentual (%)', 'mes_ano': 'Mês/Ano', 'variable': 'Métrica'}
                    )
                    fig_historico.update_layout(height=400)
                    st.plotly_chart(fig_historico, use_container_width=True)
                
                st.subheader("📋 Dados Históricos")
                st.dataframe(df_filtrado, use_container_width=True)
            else:
                st.warning("Nenhum dado histórico encontrado.")
        else:
            st.error("❌ Supabase não configurado.")
        
        if st.button("🔙 Voltar"):
            st.session_state.mostrar_historico = False
            st.rerun()
        
        st.markdown("---")
    
    # ===== VISUALIZAÇÃO DE RESULTADOS ATUAIS =====
    if st.session_state.get('processado', False) and not st.session_state.get('gerenciar_analistas', False) and not st.session_state.get('mostrar_historico', False) and not st.session_state.get('mostrar_periodos', False):
        resultados = st.session_state.resultados
        periodo = st.session_state.get('periodo', datetime.now().strftime('%B %Y'))
        
        # ===== FILTRO POR GESTOR =====
        if st.session_state.get('acesso_total', False):
            gestor_selecionado = None
            resultados_filtrados = resultados
            st.info("🔑 Acesso Total - Visualizando todos os times")
        else:
            gestor_selecionado = st.session_state.gestor
            resultados_filtrados = {k: v for k, v in resultados.items() if v['gestor'] == gestor_selecionado}
        
        if not resultados_filtrados:
            st.warning("Nenhum dado encontrado para seu time.")
            return
        
        # Métricas
        col1, col2, col3, col4, col5 = st.columns(5)
        
        total_analistas = len(resultados_filtrados)
        total_atendimentos = sum([d['total_atendimentos'] for d in resultados_filtrados.values()])
        media_atendimentos = total_atendimentos / total_analistas if total_analistas > 0 else 0
        
        with col1:
            st.metric("📊 Analistas", total_analistas)
        with col2:
            st.metric("💬 Atendimentos", f"{total_atendimentos:,}")
        with col3:
            st.metric("📈 Média Atend.", f"{media_atendimentos:.0f}")
        with col4:
            metas = len([d for d in resultados_filtrados.values() if d['status'] == '🟢 Meta Superada'])
            st.metric("🏆 Metas", f"{metas}/{total_analistas}")
        with col5:
            csat_medio = sum([d['csat'] for d in resultados_filtrados.values()]) / total_analistas if total_analistas > 0 else 0
            st.metric("⭐ CSAT Médio", f"{csat_medio:.2f}%")
        
        st.info(f"📅 Período: {periodo}")
        st.markdown("---")
        
        # ===== DASHBOARD =====
        st.header("📊 Dashboard de Performance")
        
        df_dashboard = pd.DataFrame([
            {
                'Analista': nome,
                'CSAT': dados['csat'],
                '% Avaliações': dados['perc_avaliacoes'],
                '% Envio': dados['perc_envio'],
                '💬 Atendimentos': dados['total_atendimentos'],
                'Status': dados['status']
            }
            for nome, dados in resultados_filtrados.items()
        ])
        
        if not df_dashboard.empty:
            criar_graficos_dashboard(df_dashboard)
        
        st.markdown("---")
        
        # ===== PÓDIO =====
        st.subheader("🏆 Pódio do Mês")
        
        if st.session_state.get('acesso_total', False):
            podio_gestor = None
        else:
            podio_gestor = st.session_state.gestor
        
        podio = calcular_podio(resultados, podio_gestor, media_atendimentos)
        
        if supabase and podio_gestor:
            try:
                podio_manual = carregar_podio_manual(supabase, periodo, podio_gestor)
                if podio_manual:
                    podio = podio_manual
            except:
                pass
        
        with st.expander("✏️ Editar Pódio Manualmente", expanded=False):
            st.info("Ajuste manualmente os resultados do pódio se necessário.")
            
            podio_manual = []
            for i in range(3):
                col1, col2, col3 = st.columns([2, 1, 1])
                with col1:
                    if i < len(podio):
                        nome = podio[i][0]
                    else:
                        nome = ""
                    nome_edit = st.text_input(f"{i+1}º - Nome", value=nome, key=f"podio_nome_{i}")
                with col2:
                    if i < len(podio):
                        csat = podio[i][1]
                    else:
                        csat = 0
                    csat_edit = st.number_input(f"CSAT (%)", value=float(csat), key=f"podio_csat_{i}")
                with col3:
                    if i < len(podio):
                        atend = podio[i][2]
                    else:
                        atend = 0
                    atend_edit = st.number_input(f"💬 Atendimentos", value=int(atend), key=f"podio_atend_{i}")
                
                if nome_edit:
                    podio_manual.append((nome_edit, csat_edit, atend_edit, 0))
            
            col1, col2 = st.columns(2)
            with col1:
                if podio_manual and st.button("💾 Salvar Pódio Manual"):
                    if supabase and podio_gestor:
                        try:
                            sucesso, mensagem = salvar_podio_manual(supabase, periodo, podio_gestor, podio_manual)
                            if sucesso:
                                podio = podio_manual
                                st.success("✅ Pódio manual salvo com sucesso!")
                                st.rerun()
                            else:
                                st.error(f"Erro ao salvar pódio: {mensagem}")
                        except Exception as e:
                            st.error(f"Erro ao salvar pódio: {str(e)}")
                    else:
                        st.error("❌ Supabase não configurado ou acesso total.")
            
            with col2:
                if st.button("🔄 Resetar Pódio"):
                    if supabase and podio_gestor:
                        try:
                            supabase.table('podio_manual').delete().eq('mes_ano', periodo).eq('gestor', podio_gestor).execute()
                            st.success("✅ Pódio resetado!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Erro ao resetar pódio: {str(e)}")
                    else:
                        st.error("❌ Supabase não configurado ou acesso total.")
        
        if podio:
            col1, col2, col3 = st.columns(3)
            for i, (col, (nome, csat, atendimentos, perc_avaliacoes)) in enumerate(zip([col1, col2, col3], podio), 1):
                medalha = ["🥇", "🥈", "🥉"][i-1]
                cores = ['#FFD700', '#C0C0C0', '#CD7F32']
                with col:
                    st.markdown(f"""
                    <div style="text-align: center; padding: 20px; border: 2px solid #ddd; border-radius: 10px; background-color: {cores[i-1]}20;">
                        <h1 style="font-size: 48px; margin: 0;">{medalha}</h1>
                        <h3 style="margin: 5px 0;">{i}º Lugar</h3>
                        <h2 style="margin: 5px 0;">{nome}</h2>
                        <p style="font-size: 24px; font-weight: bold; margin: 5px 0;">{csat:.2f}%</p>
                        <p style="margin: 5px 0;">CSAT</p>
                        <p style="margin: 5px 0; font-size: 16px; color: #444;">💬 {atendimentos} atendimentos</p>
                        <p style="margin: 5px 0; font-size: 14px; color: #28a745;">{perc_avaliacoes:.2f}% avaliações</p>
                    </div>
                    """, unsafe_allow_html=True)
        else:
            st.info("Nenhum analista atingiu os critérios do pódio (CSAT ≥ 90% e 💬 Atendimentos ≥ Média)")
        
        st.markdown("---")
        
        # ===== TABELA DE DESEMPENHO =====
        st.subheader("📋 Desempenho Individual")
        
        dados_tabela = []
        for analista, dados in sorted(resultados_filtrados.items(), key=lambda x: x[1]['csat'], reverse=True):
            no_podio = any(p[0] == analista for p in podio)
            podio_icon = "🏆" if no_podio else ""
            
            dados_tabela.append({
                'Analista': f"{podio_icon} {analista}",
                'CSAT': f"{dados['csat']:.2f}%",
                'Meta CSAT': f"{dados['meta_csat']:.0f}%",
                'Delta': f"{dados['delta_csat']:+.2f}%",
                '% Avaliações': f"{dados['perc_avaliacoes']:.2f}%",
                'Meta Avaliações': f"{dados['meta_geral']:.0f}%",
                '% Envio': f"{dados['perc_envio']:.2f}%",
                '💬 Atendimentos': dados['total_atendimentos'],
                'Média Operação': f"{media_atendimentos:.0f}",
                'Status': dados['status']
            })
        
        df_tabela = pd.DataFrame(dados_tabela)
        st.dataframe(df_tabela, use_container_width=True, hide_index=True)
        
        st.markdown("---")
        
        # ===== RELATÓRIO INDIVIDUAL =====
        st.subheader("📄 Gerar Relatório Individual")
        
        analista_selecionado = st.selectbox(
            "Selecione o Analista",
            list(resultados_filtrados.keys())
        )
        
        if analista_selecionado:
            dados = resultados_filtrados[analista_selecionado]
            
            # ===== 1. PAINEL VISUAL DO ANALISTA =====
            posicao_podio = criar_painel_analista(
                analista_selecionado,
                dados,
                media_atendimentos,
                podio
            )
            
            # ===== 2. GRÁFICO DE EVOLUÇÃO =====
            if supabase:
                df_historico_analista = carregar_historico(
                    supabase, 
                    analista=analista_selecionado
                )
                
                fig_mensal = gerar_grafico_mensal(
                    analista_selecionado,
                    df_historico_analista,
                    dados['meta_csat'],
                    dados['meta_geral']
                )
                if fig_mensal:
                    st.subheader("📈 Evolução Mensal")
                    st.plotly_chart(fig_mensal, use_container_width=True)
                    st.markdown("---")
            
            # ===== 3. ANÁLISE TÉCNICA =====
            with st.expander("📊 Análise Técnica", expanded=True):
                analise_tecnica = gerar_analise_tecnica(
                    analista_selecionado,
                    dados,
                    media_atendimentos,
                    podio
                )
                st.markdown(analise_tecnica)
            
            # ===== 4. FEEDBACK DE PERFORMANCE =====
            with st.expander("📝 Feedback de Performance", expanded=True):
                st.subheader(f"📝 Feedback - {analista_selecionado}")
                
                feedback_atual = gerar_feedback_manual(
                    analista_selecionado,
                    dados,
                    media_atendimentos,
                    posicao_podio
                )
                
                feedback_editado = st.text_area(
                    "✏️ Edite o feedback abaixo. Após editar, clique em 'Gerar com IA' para melhorar:",
                    value=feedback_atual,
                    height=400,
                    key="feedback_editor"
                )
                
                col1, col2 = st.columns([1, 3])
                with col1:
                    if st.button("🤖 Gerar com IA", use_container_width=True):
                        with st.spinner("Gerando feedback com IA..."):
                            feedback_gerado = gerar_feedback_ia(
                                analista_selecionado,
                                dados,
                                media_atendimentos,
                                posicao_podio,
                                feedback_editado
                            )
                            if feedback_gerado:
                                st.session_state.feedback_gerado = feedback_gerado
                                st.success("✅ Feedback gerado com IA!")
                                st.rerun()
                
                if st.session_state.get('feedback_gerado'):
                    st.markdown("---")
                    st.subheader("🤖 Feedback Gerado pela IA")
                    st.markdown(st.session_state.feedback_gerado)
                    
                    if st.button("📋 Usar este feedback"):
                        st.session_state.feedback_final = st.session_state.feedback_gerado
                        st.success("✅ Feedback selecionado!")
                
                if st.session_state.get('feedback_final'):
                    st.markdown("---")
                    st.subheader("📋 Feedback Final")
                    st.markdown(st.session_state.feedback_final)
                    feedback = st.session_state.feedback_final
                else:
                    feedback = feedback_editado
            
            # ===== 5. RELATÓRIO COMPLETO =====
            with st.expander("📄 Relatório Completo (Visualizar)", expanded=False):
                relatorio_markdown = f"""
# {analista_selecionado}

**Período:** {periodo}

## Esperado:

- ≥ {dados['meta_geral']:.0f}% de avaliações
- ≥ {dados['meta_csat']:.0f}% de Satisfação

## Atingido:

- **CSAT:** {dados['csat']:.2f}%
- **Avaliações:** {dados['perc_avaliacoes']:.2f}% ({dados['positivos']} positivos + {dados['negativos']} negativos = {dados['avaliacoes']})
- **% Envio:** {dados['perc_envio']:.2f}%
- **Atendidos:** {dados['total_atendimentos']} - {dados['total_inativos']} = {dados['validos']}
- **Média por agente:** {media_atendimentos}

---

## Análise Técnica de Desempenho

{analise_tecnica}

---

## Feedback de Performance

**Status Geral do Período:** {dados['status']}

{feedback}
"""
                st.markdown(relatorio_markdown)
            
            # ===== BOTÕES DE DOWNLOAD =====
            col1, col2 = st.columns(2)
            
            with col1:
                st.download_button(
                    label="📥 Baixar Análise (Word)",
                    data=gerar_relatorio_word(
                        analista_selecionado,
                        dados,
                        analise_tecnica,
                        "",
                        media_atendimentos,
                        podio,
                        periodo,
                        None
                    ),
                    file_name=f"Analise_{analista_selecionado.replace(' ', '_')}_{periodo.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            
            with col2:
                st.download_button(
                    label="📥 Baixar Relatório Completo (Word)",
                    data=gerar_relatorio_word(
                        analista_selecionado,
                        dados,
                        analise_tecnica,
                        feedback,
                        media_atendimentos,
                        podio,
                        periodo,
                        None
                    ),
                    file_name=f"Relatorio_Completo_{analista_selecionado.replace(' ', '_')}_{periodo.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
    
    st.markdown("---")
    st.markdown(
        """
        <div style="text-align: center; color: #666; font-size: 12px;">
            Sistema de Performance v9.0 | Controle de Períodos + IA + Histórico
        </div>
        """,
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
