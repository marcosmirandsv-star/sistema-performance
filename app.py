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
import calendar
import bcrypt

# ============================================
# CONSTANTES DE GESTORES (NOMES CORRETOS)
# ============================================

GESTOR_MARCOS = "Sua Gestão - Chat Notas"
GESTOR_POLYANA = "Gestão Polyana Ventura - Chat Outros"
GESTORES_VALIDOS = [GESTOR_MARCOS, GESTOR_POLYANA]

# ============================================
# CONFIGURACOES INICIAIS
# ============================================

st.set_page_config(
    page_title="Sistema de Performance",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================
# SUPABASE CONFIG
# ============================================

def init_supabase():
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
# FUNCAO DE PADRONIZACAO DE PERIODO
# ============================================

def padronizar_periodo(periodo):
    """Padroniza o formato do periodo para 'Mes Ano'"""
    if not periodo:
        return periodo
    return periodo.strip().title()

# ============================================
# FUNCOES DE FORCAGEM DE PERFIL E CACHE
# ============================================

def forcar_perfil_correto():
    """Forca o perfil correto para cada usuario"""
    if not st.session_state.get('logado', False):
        return
    
    usuario = st.session_state.get('usuario', '')
    
    if usuario == 'marcos':
        st.session_state.acesso_total = False
        st.session_state.perfil = "Gestor"
        st.session_state.gestor = GESTOR_MARCOS
    elif usuario == 'polyana':
        st.session_state.acesso_total = False
        st.session_state.perfil = "Gestor"
        st.session_state.gestor = GESTOR_POLYANA
    elif usuario == 'carine':
        st.session_state.acesso_total = True
        st.session_state.perfil = "Coordenador"
        st.session_state.gestor = GESTOR_MARCOS

def limpar_cache_completo():
    """Limpa todo o cache do sistema"""
    st.cache_data.clear()
    st.cache_resource.clear()
    
    keys_to_clear = ['resultados', 'processado', 'periodo', 'df_historico']
    for key in keys_to_clear:
        if key in st.session_state:
            del st.session_state[key]
    
    forcar_perfil_correto()
    st.success("✅ Cache limpo e perfil forçado!")
    st.rerun()

# ============================================
# FUNCOES DE RESET E DIAGNOSTICO
# ============================================

def resetar_usuario_carine():
    supabase = init_supabase()
    if not supabase:
        st.error("❌ Supabase não conectado")
        return False
    
    try:
        supabase.table('usuarios').delete().eq('usuario', 'carine').execute()
        supabase.table('usuarios').insert({
            'usuario': 'carine',
            'nome': 'Carine Melo',
            'senha': hash_senha('carine2026'),
            'gestor': GESTOR_MARCOS,
            'acesso_total': True
        }).execute()
        st.success("✅ Usuário Carine resetado com sucesso!")
        return True
    except Exception as e:
        st.error(f"❌ Erro ao resetar: {str(e)}")
        return False

def resetar_usuario_marcos():
    supabase = init_supabase()
    if not supabase:
        st.error("❌ Supabase não conectado")
        return False
    
    try:
        supabase.table('usuarios').delete().eq('usuario', 'marcos').execute()
        supabase.table('usuarios').insert({
            'usuario': 'marcos',
            'nome': 'Marcos Miranda',
            'senha': hash_senha('marcos2026'),
            'gestor': GESTOR_MARCOS,
            'acesso_total': False
        }).execute()
        st.success("✅ Usuário Marcos resetado com sucesso!")
        return True
    except Exception as e:
        st.error(f"❌ Erro ao resetar: {str(e)}")
        return False

def resetar_usuario_polyana():
    supabase = init_supabase()
    if not supabase:
        st.error("❌ Supabase não conectado")
        return False
    
    try:
        supabase.table('usuarios').delete().eq('usuario', 'polyana').execute()
        supabase.table('usuarios').insert({
            'usuario': 'polyana',
            'nome': 'Polyana Ventura',
            'senha': hash_senha('polyana2026'),
            'gestor': GESTOR_POLYANA,
            'acesso_total': False
        }).execute()
        st.success("✅ Usuário Polyana resetado com sucesso!")
        return True
    except Exception as e:
        st.error(f"❌ Erro ao resetar: {str(e)}")
        return False

def adicionar_dados_teste_polyana():
    supabase = init_supabase()
    if not supabase:
        st.error("❌ Supabase não conectado")
        return False
    
    try:
        existing = supabase.table('historico_performance').select('*').eq('gestor', GESTOR_POLYANA).eq('mes_ano', 'Maio 2026').execute()
        if existing.data:
            st.info("ℹ️ Dados para Polyana já existem.")
            return True
        
        analistas_polyana = [
            'Christian Matozinho', 'Diego Machado', 'Igor Siqueira',
            'Ismael Chagas Bessa', 'João Pedro Santana', 'Karolyne Moreira',
            'Luan Pereira', 'Mario Junior', 'Maycon Oliveira',
            'Miguel Augusto', 'Polliana Santana'
        ]
        
        dados_teste = []
        for i, analista in enumerate(analistas_polyana):
            csat_base = 85 + (i * 0.7) % 12
            avaliacoes_base = 22 + (i * 0.5) % 8
            atendimentos_base = 130 + (i * 3) % 30
            meta_csat = 90 if analista in ['João Pedro Santana', 'Mario Junior'] else 86
            
            if csat_base >= meta_csat and avaliacoes_base >= 25:
                status = "🟢 Meta Superada"
            elif csat_base >= meta_csat or avaliacoes_base >= 25:
                status = "🟡 Atenção"
            else:
                status = "🔴 Crítico"
            
            dados_teste.append({
                'mes_ano': 'Maio 2026',
                'analista': analista,
                'gestor': GESTOR_POLYANA,
                'csat': round(csat_base, 1),
                'perc_avaliacoes': round(avaliacoes_base, 1),
                'perc_envio': round(100 - avaliacoes_base, 1),
                'total_atendimentos': int(atendimentos_base),
                'total_inativos': 3 + (i % 5),
                'validos': int(atendimentos_base) - (3 + (i % 5)),
                'avaliacoes': int(round(avaliacoes_base / 100 * atendimentos_base)),
                'positivos': int(round(csat_base / 100 * (avaliacoes_base / 100 * atendimentos_base))),
                'negativos': int(round((100 - csat_base) / 100 * (avaliacoes_base / 100 * atendimentos_base))),
                'meta_csat': meta_csat,
                'delta_csat': round(csat_base - meta_csat, 1),
                'meta_geral': 25,
                'status': status
            })
        
        for dado in dados_teste:
            supabase.table('historico_performance').insert(dado).execute()
        
        st.success(f"✅ Dados para {len(dados_teste)} analistas da Polyana adicionados!")
        return True
    except Exception as e:
        st.error(f"❌ Erro: {str(e)}")
        return False

def diagnosticar_sistema():
    supabase = init_supabase()
    if not supabase:
        st.error("❌ Supabase não conectado")
        return
    
    st.subheader("🔍 Diagnóstico do Sistema")
    
    try:
        response = supabase.table('usuarios').select('*').execute()
        st.write("📋 Usuários no Supabase:")
        for user in response.data:
            st.write(f"- **{user['usuario']}** ({user['nome']}) - Acesso Total: {user.get('acesso_total', False)} - Gestor: {user['gestor']}")
    except Exception as e:
        st.warning(f"⚠️ Erro ao buscar usuários: {str(e)}")
    
    st.write("📋 Configuração dos Analistas:")
    analistas_config = carregar_analistas()
    for gestor, config in analistas_config.items():
        st.write(f"**{gestor}** - {len(config['membros'])} analistas")

# ============================================
# GERENCIAR USUARIOS - COMPLETO
# ============================================

def gerenciar_usuarios_supabase():
    st.header("👥 Gerenciar Usuários no Supabase")
    supabase = init_supabase()
    if not supabase:
        st.error("❌ Supabase não conectado")
        return
    
    try:
        # Carrega usuarios
        response = supabase.table('usuarios').select('*').execute()
        usuarios = response.data
        
        # ===== LISTA DE USUARIOS =====
        if usuarios:
            st.subheader("📋 Usuários Cadastrados")
            dados_tabela = []
            for u in usuarios:
                dados_tabela.append({
                    'Usuário': u['usuario'],
                    'Nome': u['nome'],
                    'Gestor': u['gestor'],
                    'Perfil': 'Coordenador' if u.get('acesso_total') else 'Gestor',
                    'Ativo': '✅ Sim' if u.get('ativo', True) else '❌ Não'
                })
            st.dataframe(pd.DataFrame(dados_tabela), use_container_width=True, hide_index=True)
        
        st.markdown("---")
        
        # ===== CRIAR USUARIO =====
        st.subheader("➕ Criar Novo Usuário")
        col1, col2, col3 = st.columns(3)
        with col1:
            novo_usuario = st.text_input("Usuário (login)", key="novo_usuario_login")
            novo_nome = st.text_input("Nome completo", key="novo_usuario_nome")
        with col2:
            nova_senha = st.text_input("Senha", type="password", key="novo_usuario_senha")
            confirma_senha = st.text_input("Confirmar senha", type="password", key="novo_usuario_confirma")
        with col3:
            novo_gestor = st.selectbox("Gestor", GESTORES_VALIDOS, key="novo_usuario_gestor")
            novo_perfil = st.selectbox("Perfil", ["Gestor", "Coordenador"], key="novo_usuario_perfil")
            novo_ativo = st.checkbox("Ativo", value=True, key="novo_usuario_ativo")
        
        if st.button("✅ Cadastrar Usuário", use_container_width=True, key="btn_criar_usuario"):
            if not novo_usuario or not novo_nome or not nova_senha:
                st.error("❌ Preencha todos os campos!")
            elif novo_usuario in [u['usuario'] for u in usuarios]:
                st.error("❌ Usuário já existe!")
            elif nova_senha != confirma_senha:
                st.error("❌ Senhas não conferem!")
            elif len(nova_senha) < 6:
                st.error("❌ Senha deve ter no mínimo 6 caracteres!")
            else:
                senha_hash = hash_senha(nova_senha)
                acesso_total = True if novo_perfil == "Coordenador" else False
                try:
                    supabase.table('usuarios').insert({
                        'usuario': novo_usuario,
                        'nome': novo_nome,
                        'senha': senha_hash,
                        'gestor': novo_gestor,
                        'acesso_total': acesso_total,
                        'ativo': novo_ativo
                    }).execute()
                    st.success(f"✅ Usuário {novo_usuario} criado com sucesso!")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Erro ao criar usuário: {str(e)}")
        
        st.markdown("---")
        
        # ===== EDITAR USUARIO =====
        st.subheader("✏️ Editar Usuário")
        if usuarios:
            usuario_selecionado = st.selectbox(
                "Selecione um usuário para editar",
                [u['usuario'] for u in usuarios],
                key="editar_usuario_select"
            )
            
            if usuario_selecionado:
                user_data = next((u for u in usuarios if u['usuario'] == usuario_selecionado), None)
                if user_data:
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        edit_nome = st.text_input("Nome", value=user_data['nome'], key="edit_usuario_nome")
                        edit_gestor = st.selectbox("Gestor", GESTORES_VALIDOS, 
                            index=GESTORES_VALIDOS.index(user_data['gestor']) if user_data['gestor'] in GESTORES_VALIDOS else 0,
                            key="edit_usuario_gestor")
                    with col2:
                        edit_senha = st.text_input("Nova senha (deixe em branco para manter)", type="password", key="edit_usuario_senha")
                        edit_confirma = st.text_input("Confirmar nova senha", type="password", key="edit_usuario_confirma")
                    with col3:
                        edit_perfil = st.selectbox("Perfil", ["Gestor", "Coordenador"],
                            index=1 if user_data.get('acesso_total') else 0,
                            key="edit_usuario_perfil")
                        edit_ativo = st.checkbox("Ativo", value=user_data.get('ativo', True), key="edit_usuario_ativo")
                    
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        if st.button("💾 Salvar Alterações", use_container_width=True, key="btn_salvar_usuario"):
                            updates = {
                                'nome': edit_nome,
                                'gestor': edit_gestor,
                                'acesso_total': True if edit_perfil == "Coordenador" else False,
                                'ativo': edit_ativo
                            }
                            if edit_senha and edit_senha == edit_confirma and len(edit_senha) >= 6:
                                updates['senha'] = hash_senha(edit_senha)
                            elif edit_senha and edit_senha != edit_confirma:
                                st.error("❌ Senhas não conferem!")
                            else:
                                try:
                                    supabase.table('usuarios').update(updates).eq('usuario', usuario_selecionado).execute()
                                    st.success(f"✅ Usuário {usuario_selecionado} atualizado!")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"❌ Erro ao atualizar: {str(e)}")
                    
                    with col2:
                        if st.button("🗑️ Excluir Usuário", use_container_width=True, key="btn_excluir_usuario"):
                            confirmar = st.checkbox("Confirmar exclusão permanente", key="confirmar_excluir_usuario")
                            if confirmar:
                                try:
                                    supabase.table('usuarios').delete().eq('usuario', usuario_selecionado).execute()
                                    st.success(f"✅ Usuário {usuario_selecionado} excluído!")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"❌ Erro ao excluir: {str(e)}")
                            else:
                                st.warning("⚠️ Marque a caixa de confirmação para excluir.")
        
        st.markdown("---")
        if st.button("🔙 Voltar", key="voltar_usuarios"):
            st.session_state.gerenciar_usuarios = False
            st.rerun()
    except Exception as e:
        st.error(f"❌ Erro: {str(e)}")

# ============================================
# FUNCOES DE LEITURA DE ARQUIVOS
# ============================================

def normalizar_nome_coluna(nome):
    import re
    nome = nome.lower().strip()
    nome = re.sub(r'[^a-záéíóúãõç0-9\s]', '', nome)
    nome = re.sub(r'\s+', ' ', nome).strip()
    return nome

def identificar_coluna_flexivel(df, padroes):
    colunas = df.columns.tolist()
    for col in colunas:
        col_normalizado = normalizar_nome_coluna(col)
        for padrao in padroes:
            padrao_normalizado = normalizar_nome_coluna(padrao)
            if padrao_normalizado in col_normalizado or col_normalizado in padrao_normalizado:
                return col
    return None

def carregar_arquivo_satisfacao(arquivo):
    try:
        df = pd.read_excel(arquivo)
    except:
        try:
            df = pd.read_csv(arquivo, encoding='utf-8')
        except:
            try:
                df = pd.read_csv(arquivo, encoding='latin-1')
            except Exception as e:
                st.error(f"❌ Não foi possível ler o arquivo de satisfação: {str(e)}")
                return None
    
    colunas = df.columns.tolist()
    st.info(f"🔍 Colunas encontradas: {', '.join(colunas)}")
    
    col_id = identificar_coluna_flexivel(df, ['ID do ticket', 'Ticket ID', 'ID', 'ticket_id'])
    col_satisfacao = identificar_coluna_flexivel(df, ['Índice de satisfação do ticket', 'Satisfaction', 'satisfacao', 'satisfação'])
    col_atribuido = identificar_coluna_flexivel(df, ['Nome do atribuído', 'Assignee', 'Atribuído', 'Responsável'])
    
    if not col_id or not col_satisfacao or not col_atribuido:
        st.error("❌ Colunas necessárias não encontradas")
        return None
    
    df_renomeado = df.rename(columns={
        col_id: 'ID do ticket',
        col_satisfacao: 'Índice de satisfação do ticket',
        col_atribuido: 'Nome do atribuído'
    })
    
    colunas_necessarias = ['ID do ticket', 'Índice de satisfação do ticket', 'Nome do atribuído']
    df_final = df_renomeado[colunas_necessarias].copy()
    
    mapa_conversao = {
        'Boa': 'Good', 'boa': 'Good', 'Ruim': 'Bad', 'ruim': 'Bad',
        'Oferecida': 'Offered', 'oferecida': 'Offered',
        'Não oferecida': 'Unoffered', 'não oferecida': 'Unoffered',
        'Nao oferecida': 'Unoffered', 'nao oferecida': 'Unoffered'
    }
    df_final['Índice de satisfação do ticket'] = df_final['Índice de satisfação do ticket'].replace(mapa_conversao)
    
    st.success(f"✅ Arquivo carregado! {len(df_final)} registros")
    return df_final

def carregar_arquivo_inativos(arquivo):
    try:
        df = pd.read_excel(arquivo)
    except:
        try:
            df = pd.read_csv(arquivo, encoding='utf-8')
        except:
            try:
                df = pd.read_csv(arquivo, encoding='latin-1')
            except Exception as e:
                st.error(f"❌ Não foi possível ler o arquivo de inativos: {str(e)}")
                return None
    
    colunas = df.columns.tolist()
    st.info(f"🔍 Colunas encontradas: {', '.join(colunas)}")
    
    col_id = identificar_coluna_flexivel(df, ['ID do ticket', 'Ticket ID', 'ID', 'ticket_id'])
    col_atribuido = identificar_coluna_flexivel(df, ['Nome do atribuído', 'Assignee', 'Atribuído', 'Responsável'])
    
    if not col_id or not col_atribuido:
        st.error("❌ Colunas necessárias não encontradas")
        return None
    
    df_renomeado = df.rename(columns={
        col_id: 'ID do ticket',
        col_atribuido: 'Nome do atribuído'
    })
    
    colunas_necessarias = ['ID do ticket', 'Nome do atribuído']
    df_final = df_renomeado[colunas_necessarias].copy()
    
    st.success(f"✅ Arquivo carregado! {len(df_final)} registros")
    return df_final

# ============================================
# FUNCOES DE BANCO DE DADOS
# ============================================

def __tratar_erro_unique(e):
    erro_str = str(e).lower()
    if 'duplicate key' in erro_str or 'unique constraint' in erro_str or '23505' in erro_str:
        return True, "Já existe um registro para este analista neste período e gestor. Nenhum dado foi duplicado."
    return False, None

def salvar_historico(supabase, dados, mes_ano, gestor):
    if not supabase:
        return False, "Supabase não configurado"
    
    # Padroniza o periodo
    mes_ano = padronizar_periodo(mes_ano)
    
    try:
        dados_por_gestor = {}
        for analista, d in dados.items():
            gestor_real = d.get('gestor', gestor)
            if gestor_real not in dados_por_gestor:
                dados_por_gestor[gestor_real] = []
            dados_por_gestor[gestor_real].append((analista, d))
        
        for gestor_real, registros_do_gestor in dados_por_gestor.items():
            existing = supabase.table('historico_performance').select('*').eq('mes_ano', mes_ano).eq('gestor', gestor_real).execute()
            if existing.data:
                return False, f"Já existe relatório para {mes_ano} do gestor {gestor_real}"
            
            for analista, d in registros_do_gestor:
                existing_registro = supabase.table('historico_performance').select('*').eq('mes_ano', mes_ano).eq('analista', analista).eq('gestor', gestor_real).execute()
                if existing_registro.data:
                    return False, f"Já existe registro para {analista} em {mes_ano} do gestor {gestor_real}"
                
                try:
                    supabase.table('historico_performance').insert({
                        'mes_ano': mes_ano,
                        'analista': analista,
                        'gestor': gestor_real,
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
                    }).execute()
                except Exception as insert_error:
                    is_unique, msg = __tratar_erro_unique(insert_error)
                    if is_unique:
                        return False, f"Já existe registro para {analista} em {mes_ano} do gestor {gestor_real}. {msg}"
                    raise insert_error
        
        return True, "Salvo com sucesso"
    except Exception as e:
        return False, str(e)

def substituir_historico(supabase, dados, mes_ano, gestor):
    if not supabase:
        return False, "Supabase não configurado"
    
    # Padroniza o periodo
    mes_ano = padronizar_periodo(mes_ano)
    
    try:
        dados_por_gestor = {}
        for analista, d in dados.items():
            gestor_real = d.get('gestor', gestor)
            if gestor_real not in dados_por_gestor:
                dados_por_gestor[gestor_real] = []
            dados_por_gestor[gestor_real].append((analista, d))
        
        for gestor_real, registros_do_gestor in dados_por_gestor.items():
            try:
                delete_result = supabase.table('historico_performance').delete().eq('mes_ano', mes_ano).eq('gestor', gestor_real).execute()
                if hasattr(delete_result, 'error') and delete_result.error:
                    return False, f"Erro ao excluir registros antigos do gestor {gestor_real}: {delete_result.error}"
            except Exception as delete_error:
                return False, f"Erro ao excluir registros antigos do gestor {gestor_real}: {str(delete_error)}"
            
            for analista, d in registros_do_gestor:
                try:
                    supabase.table('historico_performance').insert({
                        'mes_ano': mes_ano,
                        'analista': analista,
                        'gestor': gestor_real,
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
                    }).execute()
                except Exception as insert_error:
                    is_unique, msg = __tratar_erro_unique(insert_error)
                    if is_unique:
                        return False, f"Já existe registro para {analista} em {mes_ano} do gestor {gestor_real} (possível conflito durante substituição). {msg}"
                    raise insert_error
        
        return True, "Substituído com sucesso"
    except Exception as e:
        return False, str(e)

def carregar_historico(supabase, mes_ano=None, gestor=None, analista=None):
    if not supabase:
        return None
    
    # Padroniza o periodo se fornecido
    if mes_ano:
        mes_ano = padronizar_periodo(mes_ano)
    
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

def listar_periodos(supabase, gestor=None):
    if not supabase:
        return []
    
    try:
        query = supabase.table('historico_performance').select('mes_ano, gestor, data_criacao')
        if gestor:
            query = query.eq('gestor', gestor)
        response = query.execute()
        
        if gestor and len(response.data) == 0:
            all_response = supabase.table('historico_performance').select('gestor').execute()
            gestores_existentes = set([r['gestor'] for r in all_response.data])
            if gestores_existentes:
                st.info(f"🔍 Gestor '{gestor}' não encontrado. Gestores disponíveis: {', '.join(gestores_existentes)}")
        
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
        st.error(f"Erro ao listar períodos: {e}")
        return []

def excluir_periodo(supabase, mes_ano, gestor):
    if not supabase:
        return False, "Supabase não configurado"
    
    # Padroniza o periodo
    mes_ano = padronizar_periodo(mes_ano)
    
    try:
        supabase.table('historico_performance').delete().eq('mes_ano', mes_ano).eq('gestor', gestor).execute()
        supabase.table('podio_manual').delete().eq('mes_ano', mes_ano).eq('gestor', gestor).execute()
        supabase.table('podio_exclusoes').delete().eq('mes_ano', mes_ano).eq('gestor', gestor).execute()
        
        st.session_state.processado = False
        if 'resultados' in st.session_state:
            del st.session_state.resultados
        
        st.rerun()
        return True, "Período excluído com sucesso"
    except Exception as e:
        return False, str(e)

def verificar_periodo_existente(supabase, mes_ano, gestor):
    if not supabase:
        return False
    
    # Padroniza o periodo
    mes_ano = padronizar_periodo(mes_ano)
    
    try:
        response = supabase.table('historico_performance').select('*').eq('mes_ano', mes_ano).eq('gestor', gestor).execute()
        return len(response.data) > 0
    except:
        return False

def limpar_estado_sessao():
    if 'processado' in st.session_state and st.session_state.processado:
        if 'resultados' not in st.session_state or not st.session_state.resultados:
            keys_to_remove = ['resultados', 'processado', 'periodo', 'mostrar_periodos', 'mostrar_historico']
            for key in keys_to_remove:
                if key in st.session_state:
                    del st.session_state[key]
            return True
    return False

def carregar_podio_manual(supabase, mes_ano, gestor):
    if not supabase:
        return None
    
    # Padroniza o periodo
    mes_ano = padronizar_periodo(mes_ano)
    
    try:
        response = supabase.table('podio_manual').select('*').eq('mes_ano', mes_ano).eq('gestor', gestor).order('posicao').execute()
        if response.data:
            return [(d['analista'], d['csat'], d['atendimentos'], 0) for d in response.data]
        return None
    except:
        return None

def salvar_podio_manual(supabase, mes_ano, gestor, podio):
    if not supabase:
        return False, "Supabase não configurado"
    
    # Padroniza o periodo
    mes_ano = padronizar_periodo(mes_ano)
    
    try:
        supabase.table('podio_manual').delete().eq('mes_ano', mes_ano).eq('gestor', gestor).execute()
        for i, (nome, csat, atendimentos, _) in enumerate(podio, 1):
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

def carregar_exclusoes_podio(supabase, mes_ano, gestor):
    if not supabase:
        return []
    
    # Padroniza o periodo
    mes_ano = padronizar_periodo(mes_ano)
    
    try:
        response = supabase.table('podio_exclusoes').select('analista').eq('mes_ano', mes_ano).eq('gestor', gestor).execute()
        if response.data:
            return [item['analista'] for item in response.data]
        return []
    except Exception as e:
        return []

def salvar_exclusao_podio(supabase, mes_ano, gestor, analista):
    if not supabase:
        return False
    
    # Padroniza o periodo
    mes_ano = padronizar_periodo(mes_ano)
    
    try:
        supabase.table('podio_exclusoes').insert({
            'mes_ano': mes_ano,
            'gestor': gestor,
            'analista': analista
        }).execute()
        return True
    except Exception as e:
        if 'duplicate key' in str(e).lower() or 'unique constraint' in str(e).lower():
            return True
        return False

def remover_exclusao_podio(supabase, mes_ano, gestor, analista):
    if not supabase:
        return False
    
    # Padroniza o periodo
    mes_ano = padronizar_periodo(mes_ano)
    
    try:
        supabase.table('podio_exclusoes').delete().eq('mes_ano', mes_ano).eq('gestor', gestor).eq('analista', analista).execute()
        return True
    except Exception as e:
        return False

# ============================================
# FUNCOES DE BACKUP E RESTORE
# ============================================

def gerar_backup_excel(supabase):
    if not supabase:
        return None, "Supabase não configurado"
    
    try:
        buffer = io.BytesIO()
        
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            response = supabase.table('historico_performance').select('*').execute()
            df = pd.DataFrame(response.data)
            if not df.empty:
                df = df.drop(columns=['id'], errors='ignore')
            df.to_excel(writer, sheet_name='historico_performance', index=False)
            
            response = supabase.table('podio_manual').select('*').execute()
            df = pd.DataFrame(response.data)
            if not df.empty:
                df = df.drop(columns=['id'], errors='ignore')
            df.to_excel(writer, sheet_name='podio_manual', index=False)
            
            response = supabase.table('podio_exclusoes').select('*').execute()
            df = pd.DataFrame(response.data)
            if not df.empty:
                df = df.drop(columns=['id'], errors='ignore')
            df.to_excel(writer, sheet_name='podio_exclusoes', index=False)
        
        buffer.seek(0)
        return buffer, "Backup gerado com sucesso!"
    
    except Exception as e:
        return None, f"Erro ao gerar backup: {str(e)}"

def restaurar_backup_excel(supabase, arquivo_excel, modo):
    if not supabase:
        return False, "Supabase não configurado"
    
    try:
        excel_data = pd.read_excel(arquivo_excel, sheet_name=None)
        
        abas_esperadas = ['historico_performance', 'podio_manual', 'podio_exclusoes']
        for aba in abas_esperadas:
            if aba not in excel_data:
                return False, f"Arquivo de backup inválido: aba '{aba}' não encontrada"
        
        total_restaurados = 0
        total_pulados = 0
        
        chaves_unicas = {
            'historico_performance': ['mes_ano', 'analista', 'gestor'],
            'podio_manual': ['mes_ano', 'gestor', 'posicao'],
            'podio_exclusoes': ['mes_ano', 'gestor', 'analista']
        }
        
        for aba_nome, df in excel_data.items():
            if df.empty:
                continue
            
            df = df.drop(columns=['id'], errors='ignore')
            
            if modo == 'substituir':
                try:
                    supabase.table(aba_nome).delete().neq('id', 0).execute()
                except Exception as e:
                    response = supabase.table(aba_nome).select('id').execute()
                    for item in response.data:
                        supabase.table(aba_nome).delete().eq('id', item['id']).execute()
            
            registros = df.to_dict('records')
            chaves = chaves_unicas.get(aba_nome, [])
            
            for registro in registros:
                if modo == 'mesclar' and chaves:
                    query = supabase.table(aba_nome).select('*')
                    for chave in chaves:
                        if chave in registro:
                            query = query.eq(chave, registro[chave])
                    response = query.execute()
                    
                    if response.data:
                        total_pulados += 1
                        continue
                
                try:
                    supabase.table(aba_nome).insert(registro).execute()
                    total_restaurados += 1
                except Exception as e:
                    is_unique, _ = __tratar_erro_unique(e)
                    if is_unique:
                        total_pulados += 1
                    else:
                        raise e
        
        if modo == 'substituir':
            mensagem = f"Substituição concluída: {total_restaurados} registros restaurados, {total_pulados} falhas"
        else:
            mensagem = f"Mesclagem concluída: {total_restaurados} registros inseridos, {total_pulados} registros ignorados (já existentes)"
        
        return True, mensagem
    
    except Exception as e:
        return False, f"Erro ao restaurar backup: {str(e)}"

# ============================================
# GERENCIAMENTO DE USUARIOS (PRIORIZA SUPABASE)
# ============================================

def hash_senha(senha):
    return hashlib.sha256(senha.encode()).hexdigest()

def carregar_usuarios():
    supabase = init_supabase()
    if supabase:
        try:
            response = supabase.table('usuarios').select('*').execute()
            if response.data:
                usuarios = {}
                for u in response.data:
                    usuarios[u['usuario']] = {
                        'senha': u['senha'],
                        'nome': u['nome'],
                        'gestor': u['gestor'],
                        'acesso_total': u.get('acesso_total', False),
                        'ativo': u.get('ativo', True)
                    }
                return usuarios
        except Exception as e:
            st.warning(f"⚠️ Erro ao carregar do Supabase: {str(e)}")
    
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
            "gestor": GESTOR_MARCOS,
            "acesso_total": False,
            "ativo": True
        },
        "polyana": {
            "senha": hash_senha("polyana2026"),
            "nome": "Polyana Ventura",
            "gestor": GESTOR_POLYANA,
            "acesso_total": False,
            "ativo": True
        },
        "carine": {
            "senha": hash_senha("carine2026"),
            "nome": "Carine Melo",
            "gestor": GESTOR_MARCOS,
            "acesso_total": True,
            "ativo": True
        }
    }
    
    try:
        with open('usuarios.json', 'w', encoding='utf-8') as f:
            json.dump(usuarios, f, ensure_ascii=False, indent=2)
    except:
        pass
    
    return usuarios

def salvar_usuario_supabase(supabase, usuario, nome, senha_hash, gestor, acesso_total=False):
    if not supabase:
        return False
    
    try:
        existing = supabase.table('usuarios').select('*').eq('usuario', usuario).execute()
        if existing.data:
            supabase.table('usuarios').update({
                'nome': nome,
                'senha': senha_hash,
                'gestor': gestor,
                'acesso_total': acesso_total
            }).eq('usuario', usuario).execute()
        else:
            supabase.table('usuarios').insert({
                'usuario': usuario,
                'nome': nome,
                'senha': senha_hash,
                'gestor': gestor,
                'acesso_total': acesso_total
            }).execute()
        return True
    except Exception as e:
        print(f"Erro ao salvar usuário: {e}")
        return False

def salvar_usuarios_local(usuarios):
    try:
        with open('usuarios.json', 'w', encoding='utf-8') as f:
            json.dump(usuarios, f, ensure_ascii=False, indent=2)
        return True
    except:
        return False

# ============================================
# GERENCIAMENTO DE ANALISTAS
# ============================================

def carregar_analistas():
    try:
        if os.path.exists('analistas.json'):
            with open('analistas.json', 'r', encoding='utf-8') as f:
                config = json.load(f)
                if GESTOR_POLYANA in config and len(config[GESTOR_POLYANA]['membros']) >= 11:
                    return config
    except:
        pass
    
    config = {
        GESTOR_MARCOS: {
            "gestor": "Marcos Miranda",
            "meta_geral_avaliacoes": 25,
            "membros": {
                "Ana Claudia Corrêa": {"meta_csat": 90, "ativo": True},
                "João Vitor Almeida": {"meta_csat": 90, "ativo": True},
                "João Pedro Vianey": {"meta_csat": 90, "ativo": True},
                "Carlos Lemos": {"meta_csat": 90, "ativo": True},
                "Lorena Almeida": {"meta_csat": 86, "ativo": True},
                "Paulo Victor": {"meta_csat": 86, "ativo": True},
                "Rayane Nunes": {"meta_csat": 86, "ativo": True},
                "Thiago Reis": {"meta_csat": 90, "ativo": True},
                "Vanessa Silva": {"meta_csat": 86, "ativo": True}
            }
        },
        GESTOR_POLYANA: {
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
    try:
        with open('analistas.json', 'w', encoding='utf-8') as f:
            json.dump(analistas, f, ensure_ascii=False, indent=2)
        return True
    except:
        return False

# ============================================
# FUNÇÃO DE LOGIN CORRIGIDA
# ============================================

def fazer_login_corrigido():
    """Versão corrigida do login usando bcrypt"""
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔐 Login")
    
    supabase = init_supabase()
    if not supabase:
        st.sidebar.error("❌ Erro de conexão com o banco")
        return False
    
    usuario = st.sidebar.text_input("Usuário")
    senha = st.sidebar.text_input("Senha", type="password")
    
    if st.sidebar.button("Entrar", use_container_width=True):
        if not usuario or not senha:
            st.sidebar.error("❌ Preencha usuário e senha!")
            return False
        
        try:
            # Buscar usuário no Supabase
            response = supabase.table('usuarios').select('*').eq('usuario', usuario).execute()
            
            if not response.data:
                st.sidebar.error("❌ Usuário não encontrado!")
                return False
            
            user = response.data[0]
            
            # Verificar se está ativo
            if not user.get('ativo', True):
                st.sidebar.error("❌ Usuário desativado!")
                return False
            
            # Verificar senha com bcrypt
            try:
                import bcrypt
                senha_hash = user.get('senha', '')
                if bcrypt.checkpw(senha.encode('utf-8'), senha_hash.encode('utf-8')):
                    # Login bem-sucedido
                    st.session_state.logado = True
                    st.session_state.usuario = user['usuario']
                    st.session_state.nome_usuario = user['nome']
                    st.session_state.gestor = user.get('gestor', user.get('time_nome', ''))
                    st.session_state.acesso_total = user.get('acesso_total', False)
                    st.session_state.perfil = "Coordenador" if user.get('acesso_total', False) else "Gestor"
                    st.session_state.cargo = user.get('cargo', '')
                    st.session_state.time_nome = user.get('time_nome', '')
                    st.session_state.supervisor_nome = user.get('supervisor_nome', '')
                    
                    # Atualizar último login
                    try:
                        supabase.table('usuarios').update({
                            'ultimo_login': datetime.now().isoformat()
                        }).eq('id', user['id']).execute()
                    except:
                        pass
                    
                    # Forçar perfil correto
                    forcar_perfil_correto()
                    
                    st.sidebar.success(f"✅ Bem-vindo, {user['nome']}!")
                    st.rerun()
                    return True
                else:
                    st.sidebar.error("❌ Senha incorreta!")
                    return False
                    
            except ImportError:
                st.sidebar.error("❌ Erro ao verificar senha!")
                return False
                
        except Exception as e:
            st.sidebar.error(f"❌ Erro: {str(e)}")
            return False
    
    # Mostrar perfil se já estiver logado
    if st.session_state.get('logado', False):
        st.sidebar.markdown("---")
        st.sidebar.subheader("👤 Perfil")
        st.sidebar.write(f"**Usuário:** {st.session_state.get('nome_usuario')}")
        st.sidebar.write(f"**Perfil:** {st.session_state.get('perfil')}")
        if st.session_state.get('acesso_total', False):
            st.sidebar.write("**Acesso:** 🔑 Total")
        else:
            st.sidebar.write(f"**Time:** {st.session_state.get('gestor')}")
        
        st.sidebar.success("✅ Logado")
        
        if st.sidebar.button("Sair", use_container_width=True):
            keys_to_clear = ['logado', 'usuario', 'nome_usuario', 'gestor', 'acesso_total', 
                           'perfil', 'cargo', 'time_nome', 'supervisor_nome', 'resultados', 
                           'processado', 'periodo', 'df_historico']
            for key in keys_to_clear:
                if key in st.session_state:
                    del st.session_state[key]
            st.rerun()
        return True
    
    return False

# ============================================
# FUNCOES AUXILIARES
# ============================================

def get_genero_neutro(nome):
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

def calcular_media_operacao(resultados):
    if not resultados:
        return 0
    total_atendimentos = sum([d['total_atendimentos'] for d in resultados.values()])
    total_analistas = len(resultados)
    return round(total_atendimentos / total_analistas) if total_analistas > 0 else 0

def calcular_podio(resultados, media_atendimentos=None, limiar_csat=90):
    if not resultados:
        return []
    
    if media_atendimentos is None:
        media_atendimentos = calcular_media_operacao(resultados)
    
    dados_validos = {k: v for k, v in resultados.items() 
                     if v['csat'] >= limiar_csat 
                     and v['total_atendimentos'] >= media_atendimentos 
                     and v['perc_avaliacoes'] >= 25}
    
    if not dados_validos:
        return []
    
    sorted_analistas = sorted(dados_validos.items(), key=lambda x: x[1]['csat'], reverse=True)
    top_3 = sorted_analistas[:3]
    return [(nome, dados['csat'], dados['total_atendimentos'], dados['perc_avaliacoes']) for nome, dados in top_3]

# ============================================
# FUNCOES DE DASHBOARD VISUAL
# ============================================

def criar_cards_indicadores(dados, meta_csat=90, meta_avaliacoes=25, meta_envio=90):
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
        <div style="background: #ffffff; padding: 20px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); border-left: 4px solid #2ecc71;">
            <p style="font-size: 14px; color: #666; margin: 0;">📊 Total de Registros</p>
            <p style="font-size: 34px; font-weight: bold; margin: 5px 0;">{dados.get('total_registros', 0)}</p>
            <p style="font-size: 11px; color: #999; margin: 0;">Analistas no período</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        cor_csat = "#2ecc71" if dados.get('csat_medio', 0) >= meta_csat else "#e74c3c"
        st.markdown(f"""
        <div style="background: #ffffff; padding: 20px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); border-left: 4px solid {cor_csat};">
            <p style="font-size: 14px; color: #666; margin: 0;">⭐ CSAT Médio</p>
            <p style="font-size: 34px; font-weight: bold; margin: 5px 0; color: {cor_csat};">{dados.get('csat_medio', 0):.2f}%</p>
            <p style="font-size: 11px; color: #999; margin: 0;">Meta: {meta_csat}%</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        cor_avaliacoes = "#2ecc71" if dados.get('perc_avaliacoes_medio', 0) >= meta_avaliacoes else "#e74c3c"
        st.markdown(f"""
        <div style="background: #ffffff; padding: 20px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); border-left: 4px solid {cor_avaliacoes};">
            <p style="font-size: 14px; color: #666; margin: 0;">📝 % Avaliações Médio</p>
            <p style="font-size: 34px; font-weight: bold; margin: 5px 0; color: {cor_avaliacoes};">{dados.get('perc_avaliacoes_medio', 0):.2f}%</p>
            <p style="font-size: 11px; color: #999; margin: 0;">Meta: {meta_avaliacoes}%</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        cor_envio = "#2ecc71" if dados.get('perc_envio_medio', 0) >= meta_envio else "#e74c3c"
        st.markdown(f"""
        <div style="background: #ffffff; padding: 20px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); border-left: 4px solid {cor_envio};">
            <p style="font-size: 14px; color: #666; margin: 0;">📤 % Envio Médio</p>
            <p style="font-size: 34px; font-weight: bold; margin: 5px 0; color: {cor_envio};">{dados.get('perc_envio_medio', 0):.2f}%</p>
            <p style="font-size: 11px; color: #999; margin: 0;">Meta: {meta_envio}%</p>
        </div>
        """, unsafe_allow_html=True)

def criar_saude_equipe(csat_medio, perc_avaliacoes_medio, perc_envio_medio, meta_csat=90, meta_avaliacoes=25, meta_envio=90):
    st.subheader("🚦 Saúde da Equipe")
    
    if csat_medio >= meta_csat:
        cor_csat = "🟢"; status_csat = "Saudável"; bg_csat = "#d4edda"
    elif csat_medio >= meta_csat - 2:
        cor_csat = "🟡"; status_csat = "Atenção"; bg_csat = "#fff3cd"
    else:
        cor_csat = "🔴"; status_csat = "Crítico"; bg_csat = "#f8d7da"
    
    if perc_avaliacoes_medio >= meta_avaliacoes:
        cor_avaliacoes = "🟢"; status_avaliacoes = "Saudável"; bg_avaliacoes = "#d4edda"
    elif perc_avaliacoes_medio >= meta_avaliacoes - 5:
        cor_avaliacoes = "🟡"; status_avaliacoes = "Atenção"; bg_avaliacoes = "#fff3cd"
    else:
        cor_avaliacoes = "🔴"; status_avaliacoes = "Crítico"; bg_avaliacoes = "#f8d7da"
    
    if perc_envio_medio >= meta_envio:
        cor_envio = "🟢"; status_envio = "Saudável"; bg_envio = "#d4edda"
    elif perc_envio_medio >= 80:
        cor_envio = "🟡"; status_envio = "Atenção"; bg_envio = "#fff3cd"
    else:
        cor_envio = "🔴"; status_envio = "Crítico"; bg_envio = "#f8d7da"
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(f"""
        <div style="background: {bg_csat}; padding: 15px; border-radius: 10px; text-align: center;">
            <p style="font-size: 28px; margin: 0;">{cor_csat}</p>
            <p style="font-size: 20px; font-weight: bold; margin: 0;">CSAT</p>
            <p style="font-size: 18px; color: #333;">{csat_medio:.2f}%</p>
            <p style="font-size: 12px; color: #666;">Status: {status_csat}</p>
            <p style="font-size: 11px; color: #999;">Meta: {meta_csat}%</p>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div style="background: {bg_avaliacoes}; padding: 15px; border-radius: 10px; text-align: center;">
            <p style="font-size: 28px; margin: 0;">{cor_avaliacoes}</p>
            <p style="font-size: 20px; font-weight: bold; margin: 0;">Avaliações</p>
            <p style="font-size: 18px; color: #333;">{perc_avaliacoes_medio:.2f}%</p>
            <p style="font-size: 12px; color: #666;">Status: {status_avaliacoes}</p>
            <p style="font-size: 11px; color: #999;">Meta: {meta_avaliacoes}%</p>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown(f"""
        <div style="background: {bg_envio}; padding: 15px; border-radius: 10px; text-align: center;">
            <p style="font-size: 28px; margin: 0;">{cor_envio}</p>
            <p style="font-size: 20px; font-weight: bold; margin: 0;">Envio</p>
            <p style="font-size: 18px; color: #333;">{perc_envio_medio:.2f}%</p>
            <p style="font-size: 12px; color: #666;">Status: {status_envio}</p>
            <p style="font-size: 11px; color: #999;">Meta: {meta_envio}%</p>
        </div>
        """, unsafe_allow_html=True)
    st.markdown("---")

def criar_grafico_evolucao(df_historico, coluna, titulo, cor, meta=None, meta_label=None):
    if df_historico is None or df_historico.empty:
        return None
    
    meses_ordenados = ordenar_meses(df_historico['mes_ano'].unique().tolist())
    df_ordenado = df_historico.copy()
    df_ordenado['ordem'] = df_ordenado['mes_ano'].apply(lambda x: meses_ordenados.index(x) if x in meses_ordenados else 999)
    df_ordenado = df_ordenado.sort_values('ordem')
    
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=df_ordenado['mes_ano'],
        y=df_ordenado[coluna],
        mode='lines+markers',
        name=titulo,
        line=dict(color=cor, width=3),
        marker=dict(size=8, color=cor),
        hovertemplate='<b>%{x}</b><br>%{y:.2f}%<extra></extra>'
    ))
    
    if meta is not None:
        meta_data = [meta] * len(df_ordenado)
        fig.add_trace(go.Scatter(
            x=df_ordenado['mes_ano'],
            y=meta_data,
            mode='lines',
            name=meta_label if meta_label else 'Meta',
            line=dict(color='red', width=2, dash='dash'),
            hovertemplate='<b>%{x}</b><br>Meta: %{y:.0f}%<extra></extra>'
        ))
    
    fig.update_layout(
        height=400,
        margin=dict(l=20, r=20, t=50, b=20),
        yaxis=dict(
            title='Percentual (%)',
            range=[0, 100],
            gridcolor='rgba(255,255,255,0.1)',
            title_font=dict(color='#e0e0e0', size=13),
            tickfont=dict(color='#e0e0e0', size=12)
        ),
        xaxis=dict(
            title='Período',
            gridcolor='rgba(255,255,255,0.05)',
            title_font=dict(color='#e0e0e0', size=13),
            tickfont=dict(color='#e0e0e0', size=12)
        ),
        legend=dict(
            orientation='h',
            yanchor='bottom',
            y=-0.3,
            xanchor='center',
            x=0.5,
            font=dict(color='#e0e0e0', size=12)
        ),
        hovermode='x unified',
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='#e0e0e0', size=13)
    )
    return fig

def ordenar_meses(meses):
    ordem_meses = {
        'January': 1, 'February': 2, 'March': 3, 'April': 4,
        'May': 5, 'June': 6, 'July': 7, 'August': 8,
        'September': 9, 'October': 10, 'November': 11, 'December': 12,
        'Janeiro': 1, 'Fevereiro': 2, 'Março': 3, 'Abril': 4,
        'Maio': 5, 'Junho': 6, 'Julho': 7, 'Agosto': 8,
        'Setembro': 9, 'Outubro': 10, 'Novembro': 11, 'Dezembro': 12
    }
    
    def extrair_ano_mes(mes_str):
        if not mes_str:
            return (9999, 13)
        partes = mes_str.split()
        if len(partes) >= 2:
            nome_mes = partes[0]
            try:
                ano = int(partes[1])
            except ValueError:
                ano = 9999
            numero_mes = ordem_meses.get(nome_mes, 13)
            return (ano, numero_mes)
        return (9999, 13)
    
    return sorted(meses, key=extrair_ano_mes)

# ============================================
# DASHBOARD GESTOR
# ============================================

def dashboard_gestor(periodo, gestor_nome, supabase):
    # Padroniza o periodo
    periodo = padronizar_periodo(periodo)
    
    df_historico = carregar_historico(supabase, mes_ano=periodo, gestor=gestor_nome)
    
    if df_historico is None or df_historico.empty:
        st.warning(f"⚠️ Nenhum dado encontrado para {gestor_nome} no período {periodo}")
        return None
    
    resultados = {}
    for _, row in df_historico.iterrows():
        resultados[row['analista']] = {
            'total_atendimentos': row['total_atendimentos'],
            'total_inativos': row['total_inativos'],
            'validos': row['validos'],
            'avaliacoes': row['avaliacoes'],
            'positivos': row['positivos'],
            'negativos': row['negativos'],
            'perc_avaliacoes': row['perc_avaliacoes'],
            'perc_envio': row['perc_envio'],
            'csat': row['csat'],
            'meta_csat': row['meta_csat'],
            'delta_csat': row['delta_csat'],
            'meta_geral': row['meta_geral'],
            'status': row['status'],
            'gestor': row['gestor']
        }
    
    total_analistas = len(resultados)
    total_atendimentos = sum([d['total_atendimentos'] for d in resultados.values()])
    media_atendimentos = total_atendimentos / total_analistas if total_analistas > 0 else 0
    csat_medio = sum([d['csat'] for d in resultados.values()]) / total_analistas if total_analistas > 0 else 0
    perc_avaliacoes_medio = sum([d['perc_avaliacoes'] for d in resultados.values()]) / total_analistas if total_analistas > 0 else 0
    perc_envio_medio = sum([d['perc_envio'] for d in resultados.values()]) / total_analistas if total_analistas > 0 else 0
    metas_superadas = len([d for d in resultados.values() if d['status'] == '🟢 Meta Superada'])
    
    dados_cards = {
        'total_registros': total_analistas,
        'csat_medio': csat_medio,
        'perc_avaliacoes_medio': perc_avaliacoes_medio,
        'perc_envio_medio': perc_envio_medio
    }
    
    meta_csat = 90
    meta_avaliacoes = 25
    meta_envio = 90
    
    criar_cards_indicadores(dados_cards, meta_csat, meta_avaliacoes, meta_envio)
    st.info(f"📅 Período: {periodo} | 👤 Gestor: {gestor_nome} | 🏆 Metas Superadas: {metas_superadas}/{total_analistas}")
    st.markdown("---")
    
    criar_saude_equipe(csat_medio, perc_avaliacoes_medio, perc_envio_medio, meta_csat, meta_avaliacoes, meta_envio)
    
    st.subheader("📈 Evolução dos Indicadores")
    df_historico_completo = carregar_historico(supabase, gestor=gestor_nome)
    
    if df_historico_completo is not None and not df_historico_completo.empty:
        df_mensal = df_historico_completo.groupby('mes_ano').agg({
            'csat': 'mean',
            'perc_avaliacoes': 'mean',
            'perc_envio': 'mean'
        }).reset_index()
        
        fig_csat = criar_grafico_evolucao(df_mensal, 'csat', 'CSAT Médio', '#2ecc71', meta=meta_csat, meta_label=f'Meta: {meta_csat}%')
        if fig_csat:
            st.plotly_chart(fig_csat, use_container_width=True)
        
        fig_avaliacoes = criar_grafico_evolucao(df_mensal, 'perc_avaliacoes', '% Avaliações Médio', '#3498db', meta=meta_avaliacoes, meta_label=f'Meta: {meta_avaliacoes}%')
        if fig_avaliacoes:
            st.plotly_chart(fig_avaliacoes, use_container_width=True)
        
        fig_envio = criar_grafico_evolucao(df_mensal, 'perc_envio', '% Envio Médio', '#f39c12', meta=meta_envio, meta_label=f'Meta: {meta_envio}%')
        if fig_envio:
            st.plotly_chart(fig_envio, use_container_width=True)
    else:
        st.info("Dados insuficientes para gráficos de evolução. Importe mais meses.")
    
    st.markdown("---")
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("🏆 Top Performers da Equipe")
        top_analistas = sorted(resultados.items(), key=lambda x: x[1]['csat'], reverse=True)[:3]
        if top_analistas:
            for i, (nome, dados) in enumerate(top_analistas, 1):
                medalha = ["🥇", "🥈", "🥉"][i-1]
                st.markdown(f"""
                <div style="background: #f0f8ff; padding: 10px; border-radius: 8px; margin-bottom: 5px;">
                    <b>{medalha} {nome}</b> - CSAT: {dados['csat']:.2f}% | Avaliações: {dados['perc_avaliacoes']:.2f}% | Status: {dados['status']}
                </div>
                """, unsafe_allow_html=True)
    with col2:
        st.subheader("📊 Oportunidades de Melhoria")
        bottom_analistas = sorted(resultados.items(), key=lambda x: x[1]['csat'])[:3]
        if bottom_analistas:
            for i, (nome, dados) in enumerate(bottom_analistas, 1):
                st.markdown(f"""
                <div style="background: #fff5f5; padding: 10px; border-radius: 8px; margin-bottom: 5px;">
                    <b>{i}º {nome}</b> - CSAT: {dados['csat']:.2f}% | Avaliações: {dados['perc_avaliacoes']:.2f}% | Status: {dados['status']}
                </div>
                """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    st.subheader("📋 Desempenho da Equipe")
    dados_tabela = []
    for analista, dados in sorted(resultados.items(), key=lambda x: x[1]['csat'], reverse=True):
        dados_tabela.append({
            'Analista': analista,
            'CSAT': f"{dados['csat']:.2f}%",
            'Meta CSAT': f"{dados['meta_csat']:.0f}%",
            'Delta': f"{dados['delta_csat']:+.2f}%",
            '% Avaliações': f"{dados['perc_avaliacoes']:.2f}%",
            'Meta Avaliações': f"{dados['meta_geral']:.0f}%",
            '% Envio': f"{dados['perc_envio']:.2f}%",
            '💬 Atendimentos': dados['total_atendimentos'],
            'Status': dados['status']
        })
    df_tabela = pd.DataFrame(dados_tabela)
    st.dataframe(df_tabela, use_container_width=True, hide_index=True)
    
    return resultados

# ============================================
# DASHBOARD COORDENADOR
# ============================================

def dashboard_coordenador(periodo, nome_usuario, supabase):
    # Padroniza o periodo
    periodo = padronizar_periodo(periodo)
    
    df_historico = carregar_historico(supabase, mes_ano=periodo)
    
    if df_historico is None or df_historico.empty:
        st.warning(f"⚠️ Nenhum dado encontrado para o período {periodo}")
        return
    
    resultados = {}
    for _, row in df_historico.iterrows():
        resultados[row['analista']] = {
            'total_atendimentos': row['total_atendimentos'],
            'total_inativos': row['total_inativos'],
            'validos': row['validos'],
            'avaliacoes': row['avaliacoes'],
            'positivos': row['positivos'],
            'negativos': row['negativos'],
            'perc_avaliacoes': row['perc_avaliacoes'],
            'perc_envio': row['perc_envio'],
            'csat': row['csat'],
            'meta_csat': row['meta_csat'],
            'delta_csat': row['delta_csat'],
            'meta_geral': row['meta_geral'],
            'status': row['status'],
            'gestor': row['gestor']
        }
    
    total_analistas = len(resultados)
    total_atendimentos = sum([d['total_atendimentos'] for d in resultados.values()])
    media_atendimentos = total_atendimentos / total_analistas if total_analistas > 0 else 0
    csat_medio = sum([d['csat'] for d in resultados.values()]) / total_analistas if total_analistas > 0 else 0
    perc_avaliacoes_medio = sum([d['perc_avaliacoes'] for d in resultados.values()]) / total_analistas if total_analistas > 0 else 0
    perc_envio_medio = sum([d['perc_envio'] for d in resultados.values()]) / total_analistas if total_analistas > 0 else 0
    metas_superadas = len([d for d in resultados.values() if d['status'] == '🟢 Meta Superada'])
    
    dados_cards = {
        'total_registros': total_analistas,
        'csat_medio': csat_medio,
        'perc_avaliacoes_medio': perc_avaliacoes_medio,
        'perc_envio_medio': perc_envio_medio
    }
    
    meta_csat = 90
    meta_avaliacoes = 25
    meta_envio = 90
    
    criar_cards_indicadores(dados_cards, meta_csat, meta_avaliacoes, meta_envio)
    st.info(f"📅 Período: {periodo} | 👤 Coordenador: {nome_usuario} | 🏆 Metas Superadas: {metas_superadas}/{total_analistas}")
    st.markdown("---")
    
    criar_saude_equipe(csat_medio, perc_avaliacoes_medio, perc_envio_medio, meta_csat, meta_avaliacoes, meta_envio)
    
    st.subheader("📈 Evolução dos Indicadores - Operação")
    df_historico_completo = carregar_historico(supabase)
    
    if df_historico_completo is not None and not df_historico_completo.empty:
        df_mensal = df_historico_completo.groupby('mes_ano').agg({
            'csat': 'mean',
            'perc_avaliacoes': 'mean',
            'perc_envio': 'mean'
        }).reset_index()
        
        fig_csat = criar_grafico_evolucao(df_mensal, 'csat', 'CSAT Médio - Operação', '#2ecc71', meta=meta_csat, meta_label=f'Meta: {meta_csat}%')
        if fig_csat:
            st.plotly_chart(fig_csat, use_container_width=True)
            st.write("")
        
        fig_avaliacoes = criar_grafico_evolucao(df_mensal, 'perc_avaliacoes', '% Avaliações Médio - Operação', '#3498db', meta=meta_avaliacoes, meta_label=f'Meta: {meta_avaliacoes}%')
        if fig_avaliacoes:
            st.plotly_chart(fig_avaliacoes, use_container_width=True)
            st.write("")
        
        fig_envio = criar_grafico_evolucao(df_mensal, 'perc_envio', '% Envio Médio - Operação', '#f39c12', meta=meta_envio, meta_label=f'Meta: {meta_envio}%')
        if fig_envio:
            st.plotly_chart(fig_envio, use_container_width=True)
            st.write("")
    else:
        st.info("Dados insuficientes para gráficos de evolução. Importe mais meses.")
    
    st.markdown("---")
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("🥇 Top Performers - Operação")
        top_analistas = sorted(resultados.items(), key=lambda x: x[1]['csat'], reverse=True)[:5]
        for i, (nome, dados) in enumerate(top_analistas, 1):
            medalha = ["🥇", "🥈", "🥉", "4º", "5º"][i-1]
            st.markdown(f"""
            <div style="background: #f0f8ff; padding: 8px; border-radius: 8px; margin-bottom: 4px;">
                <b>{medalha} {nome}</b> - CSAT: {dados['csat']:.2f}% | {dados['gestor']} | Status: {dados['status']}
            </div>
            """, unsafe_allow_html=True)
    with col2:
        st.subheader("📊 Oportunidades - Operação")
        bottom_analistas = sorted(resultados.items(), key=lambda x: x[1]['csat'])[:5]
        for i, (nome, dados) in enumerate(bottom_analistas, 1):
            st.markdown(f"""
            <div style="background: #fff5f5; padding: 8px; border-radius: 8px; margin-bottom: 4px;">
                <b>{i}º {nome}</b> - CSAT: {dados['csat']:.2f}% | {dados['gestor']} | Status: {dados['status']}
            </div>
            """, unsafe_allow_html=True)
    
    st.markdown("---")
    st.subheader("📋 Desempenho da Operação")
    dados_tabela = []
    for analista, dados in sorted(resultados.items(), key=lambda x: x[1]['csat'], reverse=True):
        dados_tabela.append({
            'Analista': analista,
            'Gestor': dados['gestor'],
            'CSAT': f"{dados['csat']:.2f}%",
            'Meta CSAT': f"{dados['meta_csat']:.0f}%",
            'Delta': f"{dados['delta_csat']:+.2f}%",
            '% Avaliações': f"{dados['perc_avaliacoes']:.2f}%",
            'Meta Avaliações': f"{dados['meta_geral']:.0f}%",
            '% Envio': f"{dados['perc_envio']:.2f}%",
            '💬 Atendimentos': dados['total_atendimentos'],
            'Status': dados['status']
        })
    df_tabela = pd.DataFrame(dados_tabela)
    st.dataframe(df_tabela, use_container_width=True, hide_index=True)
    
    return resultados

# ============================================
# FUNCOES DE RELATORIOS
# ============================================

def gerar_analise_tecnica(analista, dados, media_operacao, podio):
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
    
    return f"""
### 1. Qualidade e Satisfação do Cliente (CSAT)

O(A) {genero} registrou um índice de **Satisfação (CSAT) de {dados['csat']:.2f}%**.

- **Comparativo com a Meta:** O resultado {analise_csat} (meta: ≥ {dados['meta_csat']:.0f}%).
- **Análise Detalhada:** Do volume total de feedbacks recebidos ({dados['avaliacoes']}), **{dados['positivos']} foram positivos**, resultando em um índice de aprovação {nivel_csat}. Houve apenas {dados['negativos']} registros negativos.

### 2. Engajamento e Coleta de Feedback

O(A) {genero} alcançou uma **Taxa de Avaliações de {dados['perc_avaliacoes']:.2f}%**.

- **Comparativo com a Meta:** A meta esperada é de no mínimo {dados['meta_geral']:.0f}% de conversão de atendimentos em avaliações. O resultado {analise_avaliacoes}.
- **Cálculo:** A taxa foi calculada sobre o volume total de {dados['avaliacoes']} avaliações divididas pelos **{dados['validos']} atendimentos válidos**.

### 3. Produtividade e Volumetria

O volume total de atendimentos realizados pelo(a) {genero} foi de **{dados['total_atendimentos']} chamados**.

- **Comparativo com a Operação:** A média de atendimentos por agente foi de {media_operacao}. {analista} absorveu uma demanda operacional **{produtividade}**.
- **Destaque:** {posicao} do mês em CSAT.
"""

def gerar_feedback_manual(analista, dados, media_operacao, posicao_podio=None):
    genero = get_genero_neutro(analista)
    texto_podio = f"🏆 {posicao_podio}º lugar no pódio do mês!" if posicao_podio else ""
    
    if dados['status'] == "🟢 Meta Superada":
        reconhecimento = "Parabéns pelo excelente desempenho! Você tem sido um exemplo para a equipe."
        direcionamento = "Continue mantendo este alto padrão de qualidade e compartilhe suas boas práticas com os colegas."
    elif dados['status'] == "🟡 Atenção":
        reconhecimento = "Seu desempenho tem sido consistente, mas há espaço para crescimento."
        direcionamento = "Foque em melhorar os indicadores que estão abaixo da meta. Conte com o suporte da gestão."
    else:
        reconhecimento = "Reconhecemos seu esforço, mas é necessário um plano de ação para melhorar os resultados."
        direcionamento = "Vamos trabalhar juntos em um plano de ação focado nas áreas que precisam de desenvolvimento."
    
    return f"""
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
- {'Continuar engajando os clientes na coleta de feedback' if dados['perc_avaliacoes'] >= dados['meta_geral'] else 'Aumentar a taxa de coleta de avaliações'}
- {'Manter a produtividade elevada' if dados['total_atendimentos'] >= media_operacao else 'Otimizar o tempo entre atendimentos'}

### 5. DIRECIONAMENTO PRÁTICO

Ações recomendadas:
1. {'Compartilhar boas práticas com a equipe' if dados['csat'] >= 92 else 'Revisar atendimentos com avaliações negativas'}
2. {'Manter a abordagem atual de coleta de feedback' if dados['perc_avaliacoes'] >= dados['meta_geral'] else 'Implementar rotina de oferecimento da pesquisa'}
3. Agendar conversa de alinhamento com a gestão

### 6. ENCERRAMENTO MOTIVADOR

Acreditamos no seu potencial de crescimento e evolução contínua! 
O desenvolvimento individual fortalece todo o time. Continue se dedicando e contando com o apoio da gestão 
para superar os desafios e celebrar as conquistas! 🚀

---
**Status:** {dados['status']}
**Data:** {datetime.now().strftime('%d/%m/%Y')}
"""

def gerar_relatorio_word(analista, dados, analise_tecnica, feedback, media_operacao, podio, periodo):
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
    doc.add_heading('Análise Técnica de Desempenho', level=1)
    
    for linha in analise_tecnica.split('\n'):
        if linha.strip():
            if linha.strip().startswith('### 1.'):
                doc.add_heading('Qualidade e Satisfação do Cliente (CSAT)', level=2)
            elif linha.strip().startswith('### 2.'):
                doc.add_heading('Engajamento e Coleta de Feedback', level=2)
            elif linha.strip().startswith('### 3.'):
                doc.add_heading('Produtividade e Volumetria', level=2)
            else:
                doc.add_paragraph(linha.strip())
    doc.add_paragraph('')
    doc.add_heading('Feedback de Performance', level=1)
    doc.add_paragraph(f'Status Geral do Período: {dados["status"]}')
    doc.add_paragraph('')
    
    for linha in feedback.split('\n'):
        if linha.strip() and not linha.strip().startswith('#'):
            if linha.strip().startswith('###'):
                doc.add_heading(linha.strip().replace('###', '').strip(), level=2)
            else:
                doc.add_paragraph(linha.strip())
    
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ============================================
# GERADOR DE FEEDBACK COM IA
# ============================================

def gerar_feedback_ia(analista, dados, media_operacao, posicao_podio=None, feedback_editado=None):
    genero = get_genero_neutro(analista)
    texto_podio = f"🏆 {posicao_podio}º lugar no pódio do mês!" if posicao_podio else "Não está no pódio"
    
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
- Posição no pódio: {texto_podio}

## ESTRUTURA OBRIGATÓRIA DO FEEDBACK:
### 1. CONTEXTO E OBSERVAÇÃO
### 2. IMPACTO
### 3. RECONHECIMENTO
### 4. OPORTUNIDADE DE DESENVOLVIMENTO
### 5. DIRECIONAMENTO PRÁTICO
### 6. ENCERRAMENTO MOTIVADOR
"""
    
    if feedback_editado:
        prompt = prompt_base + f"""
## FEEDBACK ATUAL (EDITADO PELO GESTOR):
{feedback_editado}

## INSTRUÇÃO:
Analise o feedback acima e melhore/ajuste seguindo a estrutura obrigatória.
Gere o feedback revisado e aprimorado:
"""
    else:
        prompt = prompt_base + f"""
## INSTRUÇÃO:
Gere um feedback completo de performance seguindo a estrutura obrigatória de 6 seções.
O feedback deve ser profissional, construtivo e motivador.

Gere o feedback:
"""
    
    try:
        github_token = st.secrets.get("GITHUB_TOKEN", os.environ.get("GITHUB_TOKEN", ""))
        if github_token:
            headers = {"Authorization": f"Bearer {github_token}", "Content-Type": "application/json"}
            url = "https://models.inference.ai.azure.com/chat/completions"
            payload = {
                "model": "gpt-4o-mini",
                "messages": [{"role": "system", "content": "Você é especialista em gestão de performance."}, {"role": "user", "content": prompt}],
                "temperature": 0.7,
                "max_tokens": 1500
            }
            response = requests.post(url, headers=headers, json=payload)
            if response.status_code == 200:
                result = response.json()
                return result['choices'][0]['message']['content'].strip()
        return gerar_feedback_manual(analista, dados, media_operacao, posicao_podio)
    except Exception as e:
        return gerar_feedback_manual(analista, dados, media_operacao, posicao_podio)

# ============================================
# FUNCOES DE PAINEL E GRAFICO MENSAL
# ============================================

def gerar_grafico_mensal(analista, dados_mensais, meta_csat, meta_avaliacoes):
    if dados_mensais is None or dados_mensais.empty:
        return None
    
    df_analista = dados_mensais[dados_mensais['analista'] == analista]
    if df_analista.empty or df_analista['mes_ano'].nunique() < 2:
        return None
    
    df_analista = df_analista.sort_values('mes_ano')
    meses_ordenados = ordenar_meses(df_analista['mes_ano'].unique().tolist())
    df_analista['mes_ano_ordem'] = df_analista['mes_ano'].apply(lambda x: meses_ordenados.index(x) if x in meses_ordenados else 999)
    df_analista = df_analista.sort_values('mes_ano_ordem')
    
    fig = px.line(df_analista, x='mes_ano', y='csat', title=f'📈 Evolução Mensal - {analista}', labels={'mes_ano': 'Período', 'csat': 'CSAT (%)'}, markers=True)
    fig.update_traces(line=dict(color='#2ecc71', width=3), marker=dict(size=10), name='CSAT Alcançado')
    fig.add_trace(go.Bar(x=df_analista['mes_ano'], y=df_analista['perc_avaliacoes'], name='% Avaliações', marker_color='#3498db', opacity=0.6, yaxis='y2'))
    fig.add_trace(go.Scatter(x=df_analista['mes_ano'], y=[meta_csat]*len(df_analista), name=f'Meta CSAT: {meta_csat}%', line=dict(color='#e74c3c', width=2, dash='dash'), mode='lines'))
    fig.add_trace(go.Scatter(x=df_analista['mes_ano'], y=[meta_avaliacoes]*len(df_analista), name=f'Meta Avaliações: {meta_avaliacoes}%', line=dict(color='#f39c12', width=2, dash='dot'), mode='lines', yaxis='y2'))
    
    fig.update_layout(
        xaxis_title='Período',
        xaxis=dict(
            title_font=dict(color='#e0e0e0', size=13),
            tickfont=dict(color='#e0e0e0', size=12),
            gridcolor='rgba(255,255,255,0.05)'
        ),
        yaxis=dict(
            title='CSAT (%)',
            range=[0, 100],
            side='left',
            title_font=dict(color='#e0e0e0', size=13),
            tickfont=dict(color='#e0e0e0', size=12),
            gridcolor='rgba(255,255,255,0.1)'
        ),
        yaxis2=dict(
            title='% Avaliações',
            range=[0, 100],
            overlaying='y',
            side='right',
            title_font=dict(color='#e0e0e0', size=13),
            tickfont=dict(color='#e0e0e0', size=12),
            gridcolor='rgba(255,255,255,0.05)'
        ),
        hovermode='x unified',
        legend=dict(
            orientation='h',
            yanchor='bottom',
            y=-0.3,
            xanchor='center',
            x=0.5,
            font=dict(color='#e0e0e0', size=12)
        ),
        height=450,
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='#e0e0e0', size=13)
    )
    return fig

def criar_painel_analista(analista, dados, media_operacao, podio):
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
            <p style="font-size: 34px; font-weight: bold; margin: 5px 0; color: {'#28a745' if dados['csat'] >= dados['meta_csat'] else '#dc3545'};">{dados['csat']:.1f}%</p>
            <p style="font-size: 11px; color: #666; margin: 0;">Meta: {dados['meta_csat']}%</p>
            <p style="font-size: 14px; color: {cor_delta}; margin: 0;">{delta_csat:+.1f}%</p>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        diff_avaliacoes = dados['perc_avaliacoes'] - dados['meta_geral']
        cor_diff = "green" if diff_avaliacoes >= 0 else "red"
        st.markdown(f"""
        <div style="background: #f8f9fa; padding: 15px; border-radius: 10px; text-align: center; border-left: 4px solid {'#28a745' if dados['perc_avaliacoes'] >= dados['meta_geral'] else '#dc3545'};">
            <p style="font-size: 12px; color: #666; margin: 0;">📊 % Avaliações</p>
            <p style="font-size: 34px; font-weight: bold; margin: 5px 0; color: {'#28a745' if dados['perc_avaliacoes'] >= dados['meta_geral'] else '#dc3545'};">{dados['perc_avaliacoes']:.1f}%</p>
            <p style="font-size: 11px; color: #666; margin: 0;">Meta: {dados['meta_geral']}%</p>
            <p style="font-size: 14px; color: {cor_diff}; margin: 0;">{diff_avaliacoes:+.1f}%</p>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown(f"""
        <div style="background: #f8f9fa; padding: 15px; border-radius: 10px; text-align: center; border-left: 4px solid #17a2b8;">
            <p style="font-size: 12px; color: #666; margin: 0;">📤 % Envio</p>
            <p style="font-size: 34px; font-weight: bold; margin: 5px 0; color: #17a2b8;">{dados['perc_envio']:.1f}%</p>
            <p style="font-size: 11px; color: #666; margin: 0;">Clientes não avaliaram</p>
        </div>
        """, unsafe_allow_html=True)
    with col4:
        diff_atend = dados['total_atendimentos'] - media_operacao
        cor_atend = "green" if diff_atend >= 0 else "red"
        st.markdown(f"""
        <div style="background: #f8f9fa; padding: 15px; border-radius: 10px; text-align: center; border-left: 4px solid #6c757d;">
            <p style="font-size: 12px; color: #666; margin: 0;">💬 Atendimentos</p>
            <p style="font-size: 34px; font-weight: bold; margin: 5px 0; color: #6c757d;">{dados['total_atendimentos']}</p>
            <p style="font-size: 11px; color: #666; margin: 0;">Média: {media_operacao}</p>
            <p style="font-size: 14px; color: {cor_atend}; margin: 0;">{diff_atend:+.0f}</p>
        </div>
        """, unsafe_allow_html=True)
    with col5:
        if posicao_podio:
            medalha = ["🥇", "🥈", "🥉"][posicao_podio-1]
            cores = ['#FFD700', '#C0C0C0', '#CD7F32']
            st.markdown(f"""
            <div style="background: #f8f9fa; padding: 15px; border-radius: 10px; text-align: center; border-left: 4px solid {cores[posicao_podio-1]};">
                <p style="font-size: 12px; color: #666; margin: 0;">🏆 Ranking</p>
                <p style="font-size: 40px; margin: 0;">{medalha}</p>
                <p style="font-size: 16px; font-weight: bold; margin: 5px 0; color: {cores[posicao_podio-1]};">{posicao_podio}º Lugar</p>
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
            textposition='outside',
            textfont=dict(size=14)
        ))
        fig_barras.add_trace(go.Bar(
            x=['CSAT'], 
            y=[dados['meta_csat']], 
            name='Meta', 
            marker_color='#e74c3c', 
            text=[f"{dados['meta_csat']}%"], 
            textposition='outside',
            textfont=dict(size=14)
        ))
        fig_barras.update_layout(
            title=dict(
                text='CSAT - Resultado vs Meta',
                font=dict(size=18, color='#e0e0e0')
            ),
            yaxis_title=dict(
                text='Percentual (%)',
                font=dict(size=14, color='#e0e0e0')
            ),
            yaxis=dict(
                range=[0, 100],
                tickfont=dict(size=12, color='#e0e0e0'),
                gridcolor='rgba(255,255,255,0.1)'
            ),
            xaxis=dict(
                tickfont=dict(size=12, color='#e0e0e0'),
                gridcolor='rgba(255,255,255,0.05)'
            ),
            height=480,
            showlegend=False,
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(color='#e0e0e0', size=13)
        )
        st.plotly_chart(fig_barras, use_container_width=True)
        
        fig_gauge = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=dados['csat'],
            delta={'reference': dados['meta_csat']},
            title={'text': "CSAT", 'font': {'size': 16, 'color': '#e0e0e0'}},
            number={'font': {'size': 24, 'color': '#e0e0e0'}},
            gauge={
                'axis': {'range': [None, 100], 'tickfont': {'size': 12, 'color': '#e0e0e0'}},
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
        fig_gauge.update_layout(
            height=320,
            font=dict(color='#e0e0e0', size=13),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)'
        )
        st.plotly_chart(fig_gauge, use_container_width=True)
    
    with col2:
        categorias = ['CSAT', '% Avaliações', '% Envio', 'Atendimentos']
        csat_norm = dados['csat']
        avaliacoes_norm = dados['perc_avaliacoes']
        envio_norm = dados['perc_envio']
        atend_norm = min(100, (dados['total_atendimentos'] / media_operacao) * 100) if media_operacao > 0 else 0
        
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
            title=dict(
                text='Radar de Performance',
                font=dict(size=18, color='#e0e0e0')
            ),
            polar=dict(
                radialaxis=dict(
                    visible=True,
                    range=[0, 100],
                    tickfont=dict(size=12, color='#e0e0e0'),
                    gridcolor='rgba(255,255,255,0.1)'
                ),
                angularaxis=dict(
                    tickfont=dict(size=12, color='#e0e0e0'),
                    gridcolor='rgba(255,255,255,0.05)'
                )
            ),
            height=480,
            showlegend=True,
            legend=dict(
                orientation='h',
                yanchor='bottom',
                y=-0.15,
                xanchor='center',
                x=0.5,
                font=dict(color='#e0e0e0', size=12)
            ),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(color='#e0e0e0', size=13)
        )
        st.plotly_chart(fig_radar, use_container_width=True)
    
    st.markdown("---")
    return posicao_podio

def gerar_relatorio_individual(analista, dados, media_operacao, podio, periodo, supabase, gestor_ativo, acesso_total):
    posicao_podio = criar_painel_analista(analista, dados, media_operacao, podio)
    
    if supabase:
        if acesso_total:
            df_historico_analista = carregar_historico(supabase, analista=analista)
        else:
            df_historico_analista = carregar_historico(supabase, analista=analista, gestor=gestor_ativo)
        fig_mensal = gerar_grafico_mensal(analista, df_historico_analista, dados['meta_csat'], dados['meta_geral'])
        if fig_mensal:
            st.subheader("📈 Evolução Mensal")
            st.plotly_chart(fig_mensal, use_container_width=True)
            st.markdown("---")
    
    with st.expander("📊 Análise Técnica", expanded=True):
        analise_tecnica = gerar_analise_tecnica(analista, dados, media_operacao, podio)
        st.markdown(analise_tecnica)
    
    with st.expander("📝 Feedback de Performance", expanded=True):
        st.subheader(f"📝 Feedback - {analista}")
        feedback_atual = gerar_feedback_manual(analista, dados, media_operacao, posicao_podio)
        feedback_editado = st.text_area("✏️ Edite o feedback abaixo. Após editar, clique em 'Gerar com IA' para melhorar:", value=feedback_atual, height=400, key="feedback_editor")
        
        col1, col2 = st.columns([1, 3])
        with col1:
            if st.button("🤖 Gerar com IA", use_container_width=True):
                with st.spinner("Gerando feedback com IA..."):
                    feedback_gerado = gerar_feedback_ia(analista, dados, media_operacao, posicao_podio, feedback_editado)
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
        
        feedback = st.session_state.get('feedback_final', feedback_editado)
    
    with st.expander("📄 Relatório Completo", expanded=False):
        relatorio_markdown = f"""
# {analista}

**Período:** {periodo}

## Esperado:
- ≥ {dados['meta_geral']:.0f}% de avaliações
- ≥ {dados['meta_csat']:.0f}% de Satisfação

## Atingido:
- **CSAT:** {dados['csat']:.2f}%
- **Avaliações:** {dados['perc_avaliacoes']:.2f}% ({dados['positivos']} positivos + {dados['negativos']} negativos = {dados['avaliacoes']})
- **% Envio:** {dados['perc_envio']:.2f}%
- **Atendidos:** {dados['total_atendimentos']} - {dados['total_inativos']} = {dados['validos']}
- **Média por agente:** {media_operacao}

---

## Análise Técnica de Desempenho

{analise_tecnica}

---

## Feedback de Performance

**Status Geral do Período:** {dados['status']}

{feedback}
"""
        st.markdown(relatorio_markdown)
    
    col1, col2 = st.columns(2)
    with col1:
        st.download_button(label="📥 Baixar Análise (Word)", data=gerar_relatorio_word(analista, dados, analise_tecnica, "", media_operacao, podio, periodo), file_name=f"Analise_{analista.replace(' ', '_')}_{periodo.replace(' ', '_')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    with col2:
        st.download_button(label="📥 Baixar Relatório Completo (Word)", data=gerar_relatorio_word(analista, dados, analise_tecnica, feedback, media_operacao, podio, periodo), file_name=f"Relatorio_Completo_{analista.replace(' ', '_')}_{periodo.replace(' ', '_')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

# ============================================
# NOVAS FUNCOES DE NAVEGACAO
# ============================================

def pagina_importar_periodo():
    """Pagina de importacao de periodo"""
    st.header("📁 Importar Período")
    
    supabase = init_supabase()
    if not supabase:
        st.error("❌ Supabase não configurado")
        return
    
    analistas_config = carregar_analistas()
    acesso_total = st.session_state.get('acesso_total', False)
    gestor_ativo = st.session_state.get('gestor', None)
    
    col1, col2 = st.columns(2)
    with col1:
        arquivo_satisfacao = st.file_uploader("Arquivo de Satisfação", type=['xlsx', 'csv'], key="import_satisfacao")
    with col2:
        arquivo_inativos = st.file_uploader("Arquivo de Inatividade", type=['xlsx', 'csv'], key="import_inativos")
    
    st.markdown("---")
    
    # ===== SELETORES DE MES E ANO (SUBSTITUI O TEXTO LIVRE) =====
    col1, col2 = st.columns(2)
    with col1:
        meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 
                 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
        mes_selecionado = st.selectbox("Mês", meses, key="import_mes")
    with col2:
        anos = list(range(2020, 2031))
        ano_selecionado = st.selectbox("Ano", anos, index=anos.index(2026), key="import_ano")
    
    # Monta o periodo padronizado
    periodo = padronizar_periodo(f"{mes_selecionado} {ano_selecionado}")
    st.info(f"📅 Período selecionado: **{periodo}**")
    
    if st.button("🚀 Processar Dados", use_container_width=True, key="import_processar"):
        if not arquivo_satisfacao or not arquivo_inativos:
            st.error("❌ Envie os dois arquivos.")
        else:
            with st.spinner("Processando dados..."):
                try:
                    df_satisfacao = carregar_arquivo_satisfacao(arquivo_satisfacao)
                    if df_satisfacao is None:
                        st.stop()
                    df_inativos = carregar_arquivo_inativos(arquivo_inativos)
                    if df_inativos is None:
                        st.stop()
                    
                    gestor_salvar = st.session_state.get('gestor', GESTOR_MARCOS)
                    
                    resultados_temp = processar_dados(df_satisfacao, df_inativos, analistas_config)
                    dados_por_gestor = {}
                    for analista, d in resultados_temp.items():
                        gestor_real = d.get('gestor', gestor_salvar)
                        if gestor_real not in dados_por_gestor:
                            dados_por_gestor[gestor_real] = []
                        dados_por_gestor[gestor_real].append((analista, d))
                    
                    existe_duplicidade = False
                    gestores_com_dados = []
                    for gestor_real, _ in dados_por_gestor.items():
                        if verificar_periodo_existente(supabase, periodo, gestor_real):
                            existe_duplicidade = True
                            gestores_com_dados.append(gestor_real)
                    
                    if existe_duplicidade:
                        st.warning(f"⚠️ Já existe relatório para {periodo} do(s) gestor(es): {', '.join(gestores_com_dados)}")
                        col1, col2 = st.columns(2)
                        with col1:
                            if st.button("❌ Cancelar", use_container_width=True):
                                st.stop()
                        with col2:
                            if st.button("🔄 Substituir", use_container_width=True):
                                resultados = processar_dados(df_satisfacao, df_inativos, analistas_config)
                                st.session_state.resultados = resultados
                                st.session_state.processado = True
                                st.session_state.periodo = periodo
                                sucesso, mensagem = substituir_historico(supabase, resultados, periodo, gestor_salvar)
                                if sucesso:
                                    st.success(f"✅ Período {periodo} substituído!")
                                    st.rerun()
                                else:
                                    st.error(f"❌ Erro: {mensagem}")
                        st.stop()
                    
                    resultados = processar_dados(df_satisfacao, df_inativos, analistas_config)
                    st.session_state.resultados = resultados
                    st.session_state.processado = True
                    st.session_state.periodo = periodo
                    
                    sucesso, mensagem = salvar_historico(supabase, resultados, periodo, gestor_salvar)
                    if sucesso:
                        st.success(f"✅ Dados salvos! Período: {periodo}")
                        st.rerun()
                    else:
                        st.warning(f"⚠️ Dados NÃO salvos: {mensagem}")
                except Exception as e:
                    st.error(f"❌ Erro: {str(e)}")

def pagina_gerenciar_periodos():
    """Pagina de gerenciamento de periodos"""
    st.header("📂 Gerenciar Períodos")
    
    supabase = init_supabase()
    if not supabase:
        st.error("❌ Supabase não configurado")
        return
    
    acesso_total = st.session_state.get('acesso_total', False)
    gestor_ativo = st.session_state.get('gestor', None)
    
    if acesso_total:
        periodos = listar_periodos(supabase)
    else:
        periodos = listar_periodos(supabase, gestor_ativo)
    
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
            periodo_para_excluir = st.selectbox("Selecione o período", [p['mes_ano'] for p in periodos], key="periodo_excluir")
        with col2:
            confirmar = st.checkbox("✅ Confirmar exclusão", key="confirmar_excluir")
        
        if st.button("🗑️ Excluir Período", use_container_width=True, key="btn_excluir_periodo"):
            if not confirmar:
                st.error("⚠️ Marque a caixa de confirmação.")
            else:
                gestor_periodo = next((p['gestor'] for p in periodos if p['mes_ano'] == periodo_para_excluir), None)
                if gestor_periodo:
                    if not acesso_total and gestor_periodo != gestor_ativo:
                        st.error("❌ Sem permissão para excluir.")
                    else:
                        sucesso, mensagem = excluir_periodo(supabase, periodo_para_excluir, gestor_periodo)
                        if sucesso:
                            st.success(f"✅ {mensagem}")
                            st.rerun()
                        else:
                            st.error(f"❌ Erro: {mensagem}")
    else:
        if acesso_total:
            st.info("Nenhum período importado na operação.")
        else:
            st.info(f"Nenhum período importado para {gestor_ativo}.")

def pagina_historico():
    """Pagina de consulta de historico"""
    st.header("📈 Histórico")
    
    supabase = init_supabase()
    if not supabase:
        st.error("❌ Supabase não configurado")
        return
    
    acesso_total = st.session_state.get('acesso_total', False)
    gestor_ativo = st.session_state.get('gestor', None)
    
    if acesso_total:
        df_historico = carregar_historico(supabase)
    else:
        df_historico = carregar_historico(supabase, gestor=gestor_ativo)
    
    if df_historico is not None and not df_historico.empty:
        col1, col2, col3 = st.columns(3)
        with col1:
            meses_disponiveis = ['Todos'] + sorted(df_historico['mes_ano'].unique().tolist())
            mes_selecionado = st.selectbox("Selecione o Mês", meses_disponiveis, key="hist_mes")
        with col2:
            if acesso_total:
                gestores_disponiveis = ['Todos'] + sorted(df_historico['gestor'].unique().tolist())
            else:
                gestores_disponiveis = [gestor_ativo]
            gestor_filtro = st.selectbox("Selecione o Gestor", gestores_disponiveis, key="hist_gestor")
        with col3:
            if acesso_total and gestor_filtro != 'Todos':
                df_analistas = df_historico[df_historico['gestor'] == gestor_filtro]
            elif not acesso_total:
                df_analistas = df_historico[df_historico['gestor'] == gestor_ativo]
            else:
                df_analistas = df_historico
            analistas_disponiveis = ['Todos'] + sorted(df_analistas['analista'].unique().tolist())
            analista_filtro = st.selectbox("Selecione o Analista", analistas_disponiveis, key="hist_analista")
        
        df_filtrado = df_historico.copy()
        if mes_selecionado != 'Todos':
            df_filtrado = df_filtrado[df_filtrado['mes_ano'] == mes_selecionado]
        if gestor_filtro != 'Todos' and acesso_total:
            df_filtrado = df_filtrado[df_filtrado['gestor'] == gestor_filtro]
        elif not acesso_total:
            df_filtrado = df_filtrado[df_filtrado['gestor'] == gestor_ativo]
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
            df_agrupado = df_filtrado.groupby('mes_ano').agg({
                'csat': 'mean',
                'perc_avaliacoes': 'mean',
                'perc_envio': 'mean'
            }).reset_index()
            
            meses_ordenados = ordenar_meses(df_agrupado['mes_ano'].unique().tolist())
            df_agrupado['ordem'] = df_agrupado['mes_ano'].apply(lambda x: meses_ordenados.index(x) if x in meses_ordenados else 999)
            df_agrupado = df_agrupado.sort_values('ordem')
            
            fig_historico = px.line(
                df_agrupado,
                x='mes_ano',
                y=['csat', 'perc_avaliacoes', 'perc_envio'],
                title='📈 Evolução Mensal',
                labels={'value': 'Percentual (%)', 'mes_ano': 'Mês/Ano', 'variable': 'Métrica'}
            )
            fig_historico.update_layout(height=400)
            st.plotly_chart(fig_historico, use_container_width=True)
            
            # ===== GRAFICO INDIVIDUAL SO QUANDO ANALISTA SELECIONADO =====
            if analista_filtro != 'Todos':
                st.subheader(f"📊 Evolução Mensal - {analista_filtro}")
                df_analista = df_filtrado[df_filtrado['analista'] == analista_filtro]
                if not df_analista.empty:
                    meta_csat = df_analista['meta_csat'].iloc[0] if 'meta_csat' in df_analista.columns else 90
                    meta_geral = df_analista['meta_geral'].iloc[0] if 'meta_geral' in df_analista.columns else 25
                    fig_mensal = gerar_grafico_mensal(analista_filtro, df_historico, meta_csat, meta_geral)
                    if fig_mensal:
                        st.plotly_chart(fig_mensal, use_container_width=True)
        
        st.subheader("📋 Dados Históricos")
        st.dataframe(df_filtrado, use_container_width=True)
    else:
        if acesso_total:
            st.warning("Nenhum dado histórico encontrado na operação.")
        else:
            st.warning(f"Nenhum dado histórico encontrado para {gestor_ativo}.")

def pagina_gerenciar_usuarios():
    """Pagina de gerenciamento de usuarios"""
    st.header("👥 Gerenciar Usuários")
    gerenciar_usuarios_supabase()

def pagina_gerenciar_analistas():
    """Pagina de gerenciamento de analistas (completa)"""
    st.header("📝 Gerenciar Analistas")
    
    analistas_config = carregar_analistas()
    supabase = init_supabase()
    
    gestor_selecionado = st.selectbox("Selecione o Gestor", list(analistas_config.keys()), key="gerenciar_gestor")
    if not gestor_selecionado:
        return
    
    config = analistas_config[gestor_selecionado]
    
    nova_meta = st.number_input(
        "Meta Geral de Avaliações (%)", 
        min_value=0, max_value=100, 
        value=config['meta_geral_avaliacoes'],
        key="meta_geral_input"
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
            analista_para_mover = st.selectbox("Selecione o analista", [a[0] for a in todos_analistas], key="mover_analista")
        with col2:
            gestor_atual = next((g for a, g in todos_analistas if a == analista_para_mover), None)
            todos_gestores = list(analistas_config.keys())
            indice_atual = todos_gestores.index(gestor_atual) if gestor_atual in todos_gestores else 0
            novo_gestor = st.selectbox(f"Time atual: {gestor_atual}", todos_gestores, index=indice_atual, key="novo_gestor")
        with col3:
            if novo_gestor and novo_gestor != gestor_atual:
                if st.button("🔄 Mover", use_container_width=True, key="btn_mover_analista"):
                    meta_atual = analistas_config[gestor_atual]['membros'][analista_para_mover]['meta_csat']
                    del analistas_config[gestor_atual]['membros'][analista_para_mover]
                    analistas_config[novo_gestor]['membros'][analista_para_mover] = {"meta_csat": meta_atual, "ativo": True}
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
        novo_nome = st.text_input("Nome do novo analista", key="novo_analista_nome")
    with col2:
        nova_meta_csat = st.number_input("Meta CSAT", min_value=0, max_value=100, value=86, key="novo_analista_meta")
    with col3:
        if st.button("➕ Adicionar", use_container_width=True, key="btn_adicionar_analista") and novo_nome:
            if novo_nome not in membros:
                membros[novo_nome] = {"meta_csat": nova_meta_csat, "ativo": True}
                salvar_analistas(analistas_config)
                st.rerun()
            else:
                st.error("❌ Analista já existe!")
    
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
    membro_selecionado = st.selectbox("Selecione um membro para editar", list(membros.keys()), key="editar_membro")
    if membro_selecionado:
        col1, col2, col3 = st.columns([2, 1, 1])
        with col1:
            nova_meta_membro = st.number_input(
                "Nova Meta CSAT", 
                min_value=0, max_value=100, 
                value=membros[membro_selecionado]['meta_csat'],
                key="editar_meta"
            )
        with col2:
            novo_status = st.checkbox("Ativo", value=membros[membro_selecionado]['ativo'], key="editar_status")
        with col3:
            if st.button("🗑️ Remover", use_container_width=True, key="btn_remover_analista"):
                del membros[membro_selecionado]
                salvar_analistas(analistas_config)
                st.rerun()
        
        if nova_meta_membro != membros[membro_selecionado]['meta_csat'] or novo_status != membros[membro_selecionado]['ativo']:
            membros[membro_selecionado]['meta_csat'] = nova_meta_membro
            membros[membro_selecionado]['ativo'] = novo_status
            salvar_analistas(analistas_config)
            st.success("✅ Alterações salvas!")
            st.rerun()

def pagina_configuracoes():
    """Pagina de configuracoes do sistema"""
    st.header("⚙️ Configurações do Sistema")
    
    supabase = init_supabase()
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("🔄 Resetar Carine", use_container_width=True):
            resetar_usuario_carine()
            st.rerun()
        if st.button("🔄 Resetar Marcos", use_container_width=True):
            resetar_usuario_marcos()
            st.rerun()
        if st.button("🔄 Resetar Polyana", use_container_width=True):
            resetar_usuario_polyana()
            st.rerun()
        if st.button("📥 Adicionar TODOS os Analistas da Polyana", use_container_width=True):
            adicionar_dados_teste_polyana()
            st.rerun()
    with col2:
        if st.button("🧹 LIMPAR CACHE E FORÇAR PERFIL", use_container_width=True):
            limpar_cache_completo()
            st.rerun()
        if st.button("🔍 Diagnóstico do Sistema", use_container_width=True):
            diagnosticar_sistema()
    
    st.markdown("---")
    st.subheader("💾 Backup e Restore")
    
    if supabase:
        buffer, mensagem = gerar_backup_excel(supabase)
        if buffer:
            data_atual = datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_arquivo = f"backup_sistema_{data_atual}.xlsx"
            
            st.download_button(
                label="📥 Baixar Backup (Excel)",
                data=buffer,
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.warning(mensagem)
    else:
        st.error("❌ Supabase não configurado.")
    
    st.markdown("---")
    
    st.subheader("📤 Restaurar Backup")
    st.caption("⚠️ Restaure apenas com dados do mesmo sistema para evitar conflitos.")
    
    arquivo_backup = st.file_uploader(
        "Selecione o arquivo de backup (.xlsx)",
        type=['xlsx'],
        key="upload_backup_config"
    )
    
    modo_restore = st.radio(
        "Modo de restauração:",
        ["Mesclar (recomendado, não apaga nada)", "Substituir tudo (APAGA os dados atuais antes de restaurar)"],
        key="modo_restore_config"
    )
    
    confirmacao_ok = True
    if "Substituir" in modo_restore:
        confirmacao = st.text_input(
            "Digite CONFIRMAR para prosseguir com a substituição:",
            type="password",
            key="confirmacao_substituir_config"
        )
        if confirmacao != "CONFIRMAR":
            confirmacao_ok = False
            st.warning("⚠️ Digite CONFIRMAR exatamente como está escrito para habilitar a restauração.")
    
    if arquivo_backup and confirmacao_ok and supabase:
        if st.button("🔄 Restaurar Dados", use_container_width=True, key="btn_restaurar_config"):
            with st.spinner("Restaurando dados..."):
                modo = "substituir" if "Substituir" in modo_restore else "mesclar"
                sucesso, mensagem = restaurar_backup_excel(supabase, arquivo_backup, modo)
                if sucesso:
                    st.success(f"✅ {mensagem}")
                    st.rerun()
                else:
                    st.error(f"❌ {mensagem}")
    elif not supabase:
        st.error("❌ Supabase não configurado.")

def dashboard_principal():
    """Dashboard principal (tela inicial)"""
    supabase = init_supabase()
    acesso_total = st.session_state.get('acesso_total', False)
    gestor_ativo = st.session_state.get('gestor', None)
    nome_usuario = st.session_state.get('nome_usuario', '')
    
    # ===== SELETOR DE PERIODO NO DASHBOARD =====
    if supabase:
        if acesso_total:
            periodos_disponiveis = listar_periodos(supabase)
        else:
            periodos_disponiveis = listar_periodos(supabase, gestor_ativo)
        
        if periodos_disponiveis:
            # Ordena cronologicamente
            periodos_lista = ordenar_meses([p['mes_ano'] for p in periodos_disponiveis])[::-1]
            
            # Define o periodo atual
            periodo_atual = st.session_state.get('periodo', periodos_lista[0] if periodos_lista else None)
            
            if periodo_atual not in periodos_lista and periodos_lista:
                periodo_atual = periodos_lista[0]
            
            periodo_selecionado = st.selectbox(
                "📅 Selecione o Período para visualizar",
                periodos_lista,
                index=periodos_lista.index(periodo_atual) if periodo_atual in periodos_lista else 0,
                key="seletor_periodo_dashboard_principal"
            )
            
            if periodo_selecionado != periodo_atual:
                st.session_state.periodo = periodo_selecionado
                st.rerun()
    
    # Se não há dados processados, tenta carregar o último período
    if not st.session_state.get('processado', False):
        if supabase and periodos_disponiveis:
            ultimo_periodo = ordenar_meses([p['mes_ano'] for p in periodos_disponiveis])[::-1][0]
            
            df_historico = carregar_historico(supabase, mes_ano=ultimo_periodo)
            if df_historico is not None and not df_historico.empty:
                resultados = {}
                for _, row in df_historico.iterrows():
                    resultados[row['analista']] = {
                        'total_atendimentos': row['total_atendimentos'],
                        'total_inativos': row['total_inativos'],
                        'validos': row['validos'],
                        'avaliacoes': row['avaliacoes'],
                        'positivos': row['positivos'],
                        'negativos': row['negativos'],
                        'perc_avaliacoes': row['perc_avaliacoes'],
                        'perc_envio': row['perc_envio'],
                        'csat': row['csat'],
                        'meta_csat': row['meta_csat'],
                        'delta_csat': row['delta_csat'],
                        'meta_geral': row['meta_geral'],
                        'status': row['status'],
                        'gestor': row['gestor']
                    }
                st.session_state.resultados = resultados
                st.session_state.processado = True
                st.session_state.periodo = ultimo_periodo
    
    if not st.session_state.get('processado', False) or not st.session_state.get('resultados'):
        st.info("📊 Nenhum dado carregado. Faça a importação de um período para visualizar o dashboard.")
        return
    
    periodo = st.session_state.get('periodo', datetime.now().strftime('%B %Y'))
    dashboard_resultados = st.session_state.resultados
    gestor_utilizado = gestor_ativo
    
    if acesso_total:
        st.info("🔑 Perfil: Coordenador - Visualizando toda a operação")
        
        if 'visao_coordenador' not in st.session_state:
            st.session_state.visao_coordenador = "🏢 Visão Geral (Toda a Operação)"
        
        opcoes_visao = ["🏢 Visão Geral (Toda a Operação)"]
        for gestor in GESTORES_VALIDOS:
            opcoes_visao.append(f"👥 {gestor}")
        
        visao_selecionada = st.radio(
            "📊 Visualizar:",
            opcoes_visao,
            index=opcoes_visao.index(st.session_state.visao_coordenador) if st.session_state.visao_coordenador in opcoes_visao else 0,
            horizontal=True,
            key="seletor_visao_coordenador_dashboard"
        )
        
        st.session_state.visao_coordenador = visao_selecionada
        
        if visao_selecionada == "🏢 Visão Geral (Toda a Operação)":
            dashboard_resultados = dashboard_coordenador(periodo, nome_usuario, supabase)
            gestor_utilizado = None
        else:
            gestor_escolhido = visao_selecionada.replace("👥 ", "")
            gestor_utilizado = gestor_escolhido
            st.info(f"👥 Perfil: Coordenador - Visualizando time: {gestor_escolhido}")
            dashboard_resultados = dashboard_gestor(periodo, gestor_escolhido, supabase)
    else:
        st.info(f"👥 Perfil: Gestor - Visualizando equipe: {gestor_ativo}")
        dashboard_resultados = dashboard_gestor(periodo, gestor_ativo, supabase)
    
    if dashboard_resultados is not None and len(dashboard_resultados) > 0:
        
        analistas_excluidos = []
        if gestor_utilizado is not None and supabase:
            try:
                analistas_excluidos = carregar_exclusoes_podio(supabase, periodo, gestor_utilizado)
            except:
                pass
        
        resultados_para_podio = {
            k: v for k, v in dashboard_resultados.items() 
            if k not in analistas_excluidos
        }
        
        media_atendimentos = calcular_media_operacao(resultados_para_podio)
        podio = calcular_podio(resultados_para_podio, media_atendimentos)
        
        st.markdown("---")
        st.subheader("🏆 Pódio do Mês")
        
        if analistas_excluidos:
            st.caption(f"ℹ️ {len(analistas_excluidos)} analista(s) excluído(s) do cálculo do pódio.")
        
        podio_manual_carregado = None
        if gestor_utilizado is not None and supabase:
            try:
                podio_manual_carregado = carregar_podio_manual(supabase, periodo, gestor_utilizado)
            except:
                pass
        
        podio_efetivo = podio_manual_carregado if podio_manual_carregado is not None else podio
        
        if podio_manual_carregado is not None:
            st.caption("✏️ Pódio ajustado manualmente pelo gestor")
        
        if podio_efetivo and len(podio_efetivo) > 0:
            col1, col2, col3 = st.columns(3)
            for i, (col, (nome, csat, atendimentos, perc_avaliacoes)) in enumerate(zip([col1, col2, col3], podio_efetivo), 1):
                medalha = ["🥇", "🥈", "🥉"][i-1]
                cores = ['#FFD700', '#C0C0C0', '#CD7F32']
                with col:
                    st.markdown(f"""
                    <div style="text-align: center; padding: 20px; border: 2px solid #444; border-radius: 10px; background-color: rgba(255,255,255,0.05);">
                        <h1 style="font-size: 48px; margin: 0;">{medalha}</h1>
                        <h3 style="margin: 5px 0; color: #e0e0e0;">{i}º Lugar</h3>
                        <h2 style="margin: 5px 0; color: #fff;">{nome}</h2>
                        <p style="font-size: 24px; font-weight: bold; margin: 5px 0; color: {cores[i-1]};">{csat:.2f}%</p>
                        <p style="margin: 5px 0; color: #ccc;">CSAT</p>
                        <p style="margin: 5px 0; font-size: 16px; color: #aaa;">💬 {atendimentos} atendimentos</p>
                        <p style="margin: 5px 0; font-size: 14px; color: #28a745;">{perc_avaliacoes:.2f}% avaliações</p>
                    </div>
                    """, unsafe_allow_html=True)
        else:
            st.info("🏆 Nenhum analista atingiu todos os critérios do pódio neste mês.")
        
        if gestor_utilizado is not None:
            with st.expander("✏️ Editar Pódio Manualmente", expanded=False):
                
                st.markdown("**Excluir analistas do cálculo do pódio:**")
                st.caption("Analistas com volume atípico (ex: ajuda ocasional) podem ser removidos da base de cálculo da média e elegibilidade do pódio.")
                
                exclusoes_atuais = carregar_exclusoes_podio(supabase, periodo, gestor_utilizado) if supabase else []
                
                analistas_selecionados = st.multiselect(
                    "Selecione os analistas para excluir do cálculo do pódio:",
                    options=list(dashboard_resultados.keys()),
                    default=exclusoes_atuais,
                    key="multiselect_exclusoes_podio_dashboard"
                )
                
                if set(analistas_selecionados) != set(exclusoes_atuais):
                    for analista in analistas_selecionados:
                        if analista not in exclusoes_atuais:
                            salvar_exclusao_podio(supabase, periodo, gestor_utilizado, analista)
                    for analista in exclusoes_atuais:
                        if analista not in analistas_selecionados:
                            remover_exclusao_podio(supabase, periodo, gestor_utilizado, analista)
                    st.success("✅ Exclusões atualizadas! Recarregando...")
                    st.rerun()
                
                st.markdown("---")
                
                st.info("Ajuste manualmente os resultados do pódio se necessário.")
                
                podio_atual = podio_manual_carregado if podio_manual_carregado is not None else podio
                podio_manual_edit = []
                
                for i in range(3):
                    col1, col2, col3 = st.columns([2, 1, 1])
                    with col1:
                        nome = podio_atual[i][0] if i < len(podio_atual) else ""
                        nome_edit = st.text_input(f"{i+1}º - Nome", value=nome, key=f"podio_nome_dash_{i}")
                    with col2:
                        csat = podio_atual[i][1] if i < len(podio_atual) else 0.0
                        csat_edit = st.number_input(f"CSAT (%)", value=float(csat), key=f"podio_csat_dash_{i}", step=0.1)
                    with col3:
                        atend = podio_atual[i][2] if i < len(podio_atual) else 0
                        atend_edit = st.number_input(f"💬 Atendimentos", value=int(atend), key=f"podio_atend_dash_{i}", step=1)
                    if nome_edit:
                        podio_manual_edit.append((nome_edit, csat_edit, atend_edit, 0))
                
                col1, col2 = st.columns(2)
                with col1:
                    if podio_manual_edit and st.button("💾 Salvar Pódio Manual", use_container_width=True, key="salvar_podio_dash"):
                        if supabase:
                            try:
                                sucesso, mensagem = salvar_podio_manual(supabase, periodo, gestor_utilizado, podio_manual_edit)
                                if sucesso:
                                    st.success("✅ Pódio salvo com sucesso!")
                                    st.rerun()
                                else:
                                    st.error(f"❌ Erro: {mensagem}")
                            except Exception as e:
                                st.error(f"❌ Erro: {str(e)}")
                        else:
                            st.error("❌ Supabase não configurado.")
                
                with col2:
                    if st.button("🔄 Resetar Pódio", use_container_width=True, key="resetar_podio_dash"):
                        if supabase:
                            try:
                                supabase.table('podio_manual').delete().eq('mes_ano', periodo).eq('gestor', gestor_utilizado).execute()
                                st.success("✅ Pódio resetado com sucesso!")
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ Erro: {str(e)}")
                        else:
                            st.error("❌ Supabase não configurado.")
        
        st.markdown("---")
        st.subheader("📄 Relatório Individual por Analista")
        
        if acesso_total and st.session_state.visao_coordenador == "🏢 Visão Geral (Toda a Operação)":
            analistas_disponiveis = list(dashboard_resultados.keys())
            analistas_disponiveis.sort()
        else:
            analistas_disponiveis = list(dashboard_resultados.keys())
            analistas_disponiveis.sort()
        
        if analistas_disponiveis:
            analista_selecionado = st.selectbox(
                "Selecione um analista para ver o relatório completo:",
                analistas_disponiveis,
                key="seletor_analista_relatorio_dashboard"
            )
            
            if analista_selecionado:
                dados_analista = dashboard_resultados[analista_selecionado]
                gestor_do_analista = dados_analista.get('gestor', gestor_utilizado)
                
                gerar_relatorio_individual(
                    analista_selecionado,
                    dados_analista,
                    media_atendimentos,
                    podio_efetivo,
                    periodo,
                    supabase,
                    gestor_do_analista,
                    acesso_total
                )

# ============================================
# INTERFACE PRINCIPAL
# ============================================

def main():
    st.title("📊 Sistema de Performance - Relatórios Automáticos")
    st.markdown("---")
    
    # USAR A FUNÇÃO DE LOGIN CORRIGIDA
    if not fazer_login_corrigido():
        st.info("👋 Faça login na barra lateral para acessar o sistema.")
        return
    
    forcar_perfil_correto()
    
    # ===== MENU PERMANENTE NA SIDEBAR =====
    st.sidebar.markdown("---")
    st.sidebar.header("📋 Navegação")
    
    menu_opcoes = [
        "📊 Dashboard",
        "📁 Importar Período",
        "📂 Gerenciar Períodos",
        "📈 Histórico",
        "👥 Gerenciar Usuários",
        "📝 Gerenciar Analistas",
        "⚙️ Configurações"
    ]
    
    if 'pagina_atual' not in st.session_state:
        st.session_state.pagina_atual = "📊 Dashboard"
    
    for opcao in menu_opcoes:
        if st.sidebar.button(opcao, use_container_width=True, key=f"menu_{opcao}"):
            st.session_state.pagina_atual = opcao
            st.rerun()
    
    # ===== RENDERIZA PAGINA SELECIONADA =====
    pagina = st.session_state.pagina_atual
    
    if pagina == "📊 Dashboard":
        dashboard_principal()
    elif pagina == "📁 Importar Período":
        pagina_importar_periodo()
    elif pagina == "📂 Gerenciar Períodos":
        pagina_gerenciar_periodos()
    elif pagina == "📈 Histórico":
        pagina_historico()
    elif pagina == "👥 Gerenciar Usuários":
        pagina_gerenciar_usuarios()
    elif pagina == "📝 Gerenciar Analistas":
        pagina_gerenciar_analistas()
    elif pagina == "⚙️ Configurações":
        pagina_configuracoes()
    
    st.markdown("---")
    st.markdown('<div style="text-align: center; color: #666; font-size: 12px;">Sistema de Performance v12.0 | Navegação Reorganizada</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
