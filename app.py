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

# ============================================
# CONSTANTES DE GESTORES (NOMES CORRETOS)
# ============================================

GESTOR_MARCOS = "Sua Gestão - Chat Notas"
GESTOR_POLYANA = "Gestão Polyana Ventura - Chat Outros"
GESTORES_VALIDOS = [GESTOR_MARCOS, GESTOR_POLYANA]

# ============================================
# CONFIGURAÇÕES INICIAIS
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
# FUNÇÕES DE RESET E DIAGNÓSTICO
# ============================================

def resetar_usuario_carine():
    """Resetar a Carine no Supabase"""
    supabase = init_supabase()
    if not supabase:
        st.error("❌ Supabase não conectado")
        return False
    
    try:
        # Remove usuário se existir
        supabase.table('usuarios').delete().eq('usuario', 'carine').execute()
        
        # Insere com senha padrão e acesso total
        supabase.table('usuarios').insert({
            'usuario': 'carine',
            'nome': 'Carine Melo',
            'senha': hash_senha('carine2026'),
            'gestor': GESTOR_MARCOS,
            'acesso_total': True
        }).execute()
        
        st.success("✅ Usuário Carine resetado com sucesso no Supabase!")
        return True
    except Exception as e:
        st.error(f"❌ Erro ao resetar: {str(e)}")
        return False

def resetar_usuario_marcos():
    """Resetar o Marcos no Supabase"""
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
        
        st.success("✅ Usuário Marcos resetado com sucesso no Supabase!")
        return True
    except Exception as e:
        st.error(f"❌ Erro ao resetar: {str(e)}")
        return False

def resetar_usuario_polyana():
    """Resetar a Polyana no Supabase"""
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
        
        st.success("✅ Usuário Polyana resetado com sucesso no Supabase!")
        return True
    except Exception as e:
        st.error(f"❌ Erro ao resetar: {str(e)}")
        return False

def adicionar_dados_teste_polyana():
    """Função para adicionar dados de teste para TODOS os analistas da Polyana"""
    supabase = init_supabase()
    if not supabase:
        st.error("❌ Supabase não conectado")
        return False
    
    try:
        # Verifica se já existem dados
        existing = supabase.table('historico_performance').select('*').eq('gestor', GESTOR_POLYANA).eq('mes_ano', 'Maio 2026').execute()
        if existing.data:
            st.info("ℹ️ Dados para Polyana já existem.")
            return True
        
        # LISTA COMPLETA DOS ANALISTAS DA POLYANA
        analistas_polyana = [
            'Christian Matozinho',
            'Diego Machado',
            'Igor Siqueira',
            'Ismael Chagas Bessa',
            'João Pedro Santana',
            'Karolyne Moreira',
            'Luan Pereira',
            'Mario Junior',
            'Maycon Oliveira',
            'Miguel Augusto',
            'Polliana Santana'
        ]
        
        # Dados de exemplo para TODOS os analistas da Polyana
        dados_teste = []
        
        # Dados variados para cada analista
        for i, analista in enumerate(analistas_polyana):
            # Gera dados variados para cada analista
            csat_base = 85 + (i * 0.7) % 12  # Varia entre 85 e 97
            avaliacoes_base = 22 + (i * 0.5) % 8  # Varia entre 22 e 30
            atendimentos_base = 130 + (i * 3) % 30  # Varia entre 130 e 160
            
            # Determina meta CSAT (90 para João Pedro Santana e Mario Junior, 86 para os demais)
            meta_csat = 90 if analista in ['João Pedro Santana', 'Mario Junior'] else 86
            
            # Calcula status
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
        
        st.success(f"✅ Dados de teste para {len(dados_teste)} analistas da Polyana adicionados com sucesso!")
        return True
    except Exception as e:
        st.error(f"❌ Erro ao adicionar dados: {str(e)}")
        return False

def diagnosticar_sistema():
    """Função de diagnóstico do sistema"""
    supabase = init_supabase()
    if not supabase:
        st.error("❌ Supabase não conectado")
        return
    
    st.subheader("🔍 Diagnóstico do Sistema")
    
    # Verifica usuários no Supabase
    try:
        response = supabase.table('usuarios').select('*').execute()
        st.write("📋 Usuários no Supabase:")
        for user in response.data:
            st.write(f"- **{user['usuario']}** ({user['nome']}) - Acesso Total: {user.get('acesso_total', False)} - Gestor: {user['gestor']}")
    except Exception as e:
        st.warning(f"⚠️ Erro ao buscar usuários: {str(e)}")
        st.info("💡 A tabela 'usuarios' pode não existir. Execute o script SQL para criá-la.")
    
    # Verifica configuração dos analistas
    st.write("📋 Configuração dos Analistas:")
    analistas_config = carregar_analistas()
    
    for gestor, config in analistas_config.items():
        st.write(f"**{gestor}** - {len(config['membros'])} analistas:")
        for analista, dados in config['membros'].items():
            st.write(f"  - {analista} (Meta: {dados['meta_csat']}%, Ativo: {dados['ativo']})")
    
    # Verifica períodos
    try:
        response = supabase.table('historico_performance').select('mes_ano, gestor, analista').execute()
        st.write("📊 Dados no Supabase por Gestor:")
        dados_por_gestor = {}
        for item in response.data:
            key = item['gestor']
            if key not in dados_por_gestor:
                dados_por_gestor[key] = []
            dados_por_gestor[key].append(item['analista'])
        
        for gestor, analistas in dados_por_gestor.items():
            st.write(f"**{gestor}** - {len(analistas)} analistas:")
            for analista in analistas[:5]:  # Mostra apenas os 5 primeiros para não poluir
                st.write(f"  - {analista}")
            if len(analistas) > 5:
                st.write(f"  - ... e mais {len(analistas) - 5} analistas")
    except Exception as e:
        st.error(f"Erro ao buscar dados: {e}")

def gerenciar_usuarios_supabase():
    """Interface para gerenciar usuários no Supabase"""
    st.header("👥 Gerenciar Usuários no Supabase")
    
    supabase = init_supabase()
    if not supabase:
        st.error("❌ Supabase não conectado")
        return
    
    try:
        # Lista usuários
        response = supabase.table('usuarios').select('*').execute()
        usuarios = response.data
        
        if usuarios:
            st.subheader("📋 Usuários Cadastrados")
            dados_tabela = []
            for u in usuarios:
                dados_tabela.append({
                    'Usuário': u['usuario'],
                    'Nome': u['nome'],
                    'Gestor': u['gestor'],
                    'Acesso Total': '✅ Sim' if u.get('acesso_total') else '❌ Não'
                })
            st.dataframe(pd.DataFrame(dados_tabela), use_container_width=True, hide_index=True)
        
        st.markdown("---")
        st.subheader("✏️ Atualizar/Resetar Usuário")
        
        if usuarios:
            col1, col2 = st.columns(2)
            with col1:
                usuario_selecionado = st.selectbox("Selecione um usuário", [u['usuario'] for u in usuarios])
            with col2:
                if st.button("🔄 Resetar Senha", use_container_width=True):
                    st.session_state.resetar_senha_usuario = usuario_selecionado
                    st.rerun()
            
            if st.session_state.get('resetar_senha_usuario'):
                usuario_reset = st.session_state.resetar_senha_usuario
                st.info(f"Resetando senha de: {usuario_reset}")
                nova_senha = st.text_input("Nova senha (mínimo 6 caracteres)", type="password", key="nova_senha_input")
                if st.button("Confirmar Reset", use_container_width=True):
                    if nova_senha and len(nova_senha) >= 6:
                        supabase.table('usuarios').update({
                            'senha': hash_senha(nova_senha)
                        }).eq('usuario', usuario_reset).execute()
                        st.success(f"✅ Senha de {usuario_reset} resetada!")
                        del st.session_state.resetar_senha_usuario
                        st.rerun()
                    else:
                        st.error("❌ Senha deve ter no mínimo 6 caracteres")
        
        st.markdown("---")
        st.subheader("🔄 Trocar Perfil")
        
        if usuarios:
            col1, col2 = st.columns(2)
            with col1:
                usuario_para_trocar = st.selectbox("Selecione o usuário", [u['usuario'] for u in usuarios], key="trocar_perfil")
            with col2:
                # Pega o perfil atual
                perfil_atual = next((u.get('acesso_total') for u in usuarios if u['usuario'] == usuario_para_trocar), False)
                perfil_atual_texto = "Coordenador" if perfil_atual else "Gestor"
                
                novo_perfil = st.selectbox(
                    "Novo Perfil", 
                    ['Gestor', 'Coordenador'],
                    index=1 if perfil_atual else 0
                )
                acesso = True if novo_perfil == 'Coordenador' else False
                
                if st.button("✅ Atualizar Perfil", use_container_width=True):
                    supabase.table('usuarios').update({
                        'acesso_total': acesso
                    }).eq('usuario', usuario_para_trocar).execute()
                    st.success(f"✅ Perfil de {usuario_para_trocar} atualizado para {novo_perfil}!")
                    st.rerun()
        
        st.markdown("---")
        if st.button("🔙 Voltar", key="voltar_usuarios"):
            st.session_state.gerenciar_usuarios = False
            st.rerun()
            
    except Exception as e:
        st.error(f"❌ Erro: {str(e)}")

# ============================================
# FUNÇÕES DE LEITURA DE ARQUIVOS
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
    except Exception as e:
        try:
            df = pd.read_csv(arquivo, encoding='utf-8')
        except:
            try:
                df = pd.read_csv(arquivo, encoding='latin-1')
            except:
                st.error(f"❌ Não foi possível ler o arquivo de satisfação: {str(e)}")
                return None
    
    colunas = df.columns.tolist()
    st.info(f"🔍 Colunas encontradas: {', '.join(colunas)}")
    
    col_id = identificar_coluna_flexivel(df, ['ID do ticket', 'Ticket ID', 'ID', 'ticket_id'])
    col_satisfacao = identificar_coluna_flexivel(df, [
        'Índice de satisfação do ticket', 'Satisfaction', 'satisfacao', 'satisfação'
    ])
    col_atribuido = identificar_coluna_flexivel(df, [
        'Nome do atribuído', 'Assignee', 'Atribuído', 'Responsável'
    ])
    
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
    except Exception as e:
        try:
            df = pd.read_csv(arquivo, encoding='utf-8')
        except:
            try:
                df = pd.read_csv(arquivo, encoding='latin-1')
            except:
                st.error(f"❌ Não foi possível ler o arquivo de inativos: {str(e)}")
                return None
    
    colunas = df.columns.tolist()
    st.info(f"🔍 Colunas encontradas: {', '.join(colunas)}")
    
    col_id = identificar_coluna_flexivel(df, ['ID do ticket', 'Ticket ID', 'ID', 'ticket_id'])
    col_atribuido = identificar_coluna_flexivel(df, [
        'Nome do atribuído', 'Assignee', 'Atribuído', 'Responsável'
    ])
    
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
# FUNÇÕES DE BANCO DE DADOS
# ============================================

def salvar_historico(supabase, dados, mes_ano, gestor):
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
        
        existing = supabase.table('historico_performance').select('*').eq('mes_ano', mes_ano).eq('gestor', gestor).execute()
        if existing.data:
            return False, f"Já existe relatório para {mes_ano}"
        
        for registro in registros:
            supabase.table('historico_performance').insert(registro).execute()
        return True, "Salvo com sucesso"
    except Exception as e:
        return False, str(e)

def substituir_historico(supabase, dados, mes_ano, gestor):
    if not supabase:
        return False, "Supabase não configurado"
    
    try:
        supabase.table('historico_performance').delete().eq('mes_ano', mes_ano).eq('gestor', gestor).execute()
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

def listar_periodos(supabase, gestor=None):
    if not supabase:
        return []
    
    try:
        query = supabase.table('historico_performance').select('mes_ano, gestor, data_criacao')
        if gestor:
            query = query.eq('gestor', gestor)
        response = query.execute()
        
        # ===== DIAGNÓSTICO =====
        if gestor and len(response.data) == 0:
            # Tenta buscar todos os gestores para diagnóstico
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
    
    try:
        supabase.table('historico_performance').delete().eq('mes_ano', mes_ano).eq('gestor', gestor).execute()
        supabase.table('podio_manual').delete().eq('mes_ano', mes_ano).eq('gestor', gestor).execute()
        
        keys_to_remove = ['resultados', 'processado', 'periodo', 'mostrar_periodos', 'mostrar_historico']
        for key in keys_to_remove:
            if key in st.session_state:
                del st.session_state[key]
        st.rerun()
        return True, "Período excluído com sucesso"
    except Exception as e:
        return False, str(e)

def verificar_periodo_existente(supabase, mes_ano, gestor):
    if not supabase:
        return False
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

# ============================================
# GERENCIAMENTO DE USUÁRIOS (PRIORIZA SUPABASE)
# ============================================

def hash_senha(senha):
    return hashlib.sha256(senha.encode()).hexdigest()

def carregar_usuarios():
    """Carrega usuários do Supabase (prioridade) ou do arquivo local"""
    
    # Tenta carregar do Supabase primeiro
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
                        'acesso_total': u.get('acesso_total', False)
                    }
                return usuarios
        except Exception as e:
            st.warning(f"⚠️ Erro ao carregar do Supabase: {str(e)}")
    
    # Fallback: arquivo local
    try:
        if os.path.exists('usuarios.json'):
            with open('usuarios.json', 'r', encoding='utf-8') as f:
                return json.load(f)
    except:
        pass
    
    # Usuários padrão (emergência)
    usuarios = {
        "marcos": {
            "senha": hash_senha("marcos2026"),
            "nome": "Marcos Miranda",
            "gestor": GESTOR_MARCOS,
            "acesso_total": False
        },
        "polyana": {
            "senha": hash_senha("polyana2026"),
            "nome": "Polyana Ventura",
            "gestor": GESTOR_POLYANA,
            "acesso_total": False
        },
        "carine": {
            "senha": hash_senha("carine2026"),
            "nome": "Carine Melo",
            "gestor": GESTOR_MARCOS,
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
    """Salva ou atualiza um usuário no Supabase"""
    if not supabase:
        return False
    
    try:
        # Verifica se o usuário já existe
        existing = supabase.table('usuarios').select('*').eq('usuario', usuario).execute()
        
        if existing.data:
            # Atualiza usuário existente
            supabase.table('usuarios').update({
                'nome': nome,
                'senha': senha_hash,
                'gestor': gestor,
                'acesso_total': acesso_total
            }).eq('usuario', usuario).execute()
        else:
            # Insere novo usuário
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
    """Salva usuários no arquivo local (fallback)"""
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
    """Carrega analistas com a configuração correta para ambos os gestores"""
    try:
        if os.path.exists('analistas.json'):
            with open('analistas.json', 'r', encoding='utf-8') as f:
                config = json.load(f)
                # Verifica se a configuração está completa
                if GESTOR_POLYANA in config and len(config[GESTOR_POLYANA]['membros']) >= 11:
                    return config
                else:
                    st.warning("⚠️ Configuração dos analistas incompleta. Recriando...")
    except:
        pass
    
    # Configuração COMPLETA dos analistas
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
        st.success("✅ Arquivo analistas.json recriado com todos os analistas!")
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
# LOGIN
# ============================================

def fazer_login():
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔐 Login")
    
    # ===== BOTÃO DE FERRAMENTAS =====
    with st.sidebar.expander("🔧 Ferramentas do Sistema", expanded=False):
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
        
        if st.button("👥 Gerenciar Usuários (Supabase)", use_container_width=True):
            st.session_state.gerenciar_usuarios = True
            st.rerun()
        
        if st.button("🔍 Diagnóstico do Sistema", use_container_width=True):
            diagnosticar_sistema()
    
    # ===== LOGIN =====
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
            st.session_state.perfil = "Coordenador" if st.session_state.acesso_total else "Gestor"
            st.rerun()
        else:
            st.sidebar.error("❌ Usuário ou senha inválidos!")
            if usuario in usuarios:
                st.sidebar.warning(f"⚠️ Usuário '{usuario}' existe, mas a senha está incorreta")
                st.sidebar.info("💡 Use 'Resetar' nas ferramentas do sistema.")
            else:
                st.sidebar.warning(f"⚠️ Usuário '{usuario}' não encontrado")
                st.sidebar.info("Usuários disponíveis: " + ", ".join(usuarios.keys()))
                st.sidebar.info("💡 Use 'Resetar' nas ferramentas do sistema.")
    
    if st.session_state.get('logado', False):
        st.sidebar.success(f"✅ Logado como {st.session_state.nome_usuario}")
        
        if st.session_state.get('acesso_total', False):
            st.sidebar.info("🔑 Perfil: Coordenador - Acesso Total")
        else:
            st.sidebar.info(f"👥 Perfil: Gestor - Time: {st.session_state.gestor}")
        
        if st.session_state.get('acesso_total', False):
            if st.sidebar.button("👤 Cadastrar Usuário", use_container_width=True):
                st.session_state.cadastrar_usuario = True
        
        if st.sidebar.button("Sair", use_container_width=True):
            st.session_state.logado = False
            st.session_state.usuario = None
            st.session_state.nome_usuario = None
            st.session_state.gestor = None
            st.session_state.acesso_total = False
            st.session_state.perfil = None
            st.rerun()
        return True
    return False

def cadastrar_usuario():
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
        gestor_usuario = st.selectbox("Gestor", GESTORES_VALIDOS)
        acesso_total = st.checkbox("🔑 Acesso Total ao Sistema (Coordenador)", value=False)
    
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
                    st.success(f"✅ Usuário {novo_usuario} cadastrado no Supabase!")
                    st.rerun()
                else:
                    st.error("❌ Erro ao salvar no Supabase!")
            else:
                # Fallback para arquivo local
                usuarios[novo_usuario] = {
                    "senha": senha_hash,
                    "nome": novo_nome,
                    "gestor": gestor_usuario,
                    "acesso_total": acesso_total
                }
                if salvar_usuarios_local(usuarios):
                    st.success(f"✅ Usuário {novo_usuario} cadastrado localmente!")
                    st.rerun()
                else:
                    st.error("❌ Erro ao salvar usuário!")
    
    if st.button("🔙 Voltar"):
        st.session_state.cadastrar_usuario = False
        st.rerun()

# ============================================
# FUNÇÕES AUXILIARES
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
# FUNÇÕES DE DASHBOARD VISUAL
# ============================================

def criar_cards_indicadores(dados, meta_csat=90, meta_avaliacoes=25, meta_envio=90):
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
        <div style="background: #ffffff; padding: 20px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); border-left: 4px solid #2ecc71;">
            <p style="font-size: 14px; color: #666; margin: 0;">📊 Total de Registros</p>
            <p style="font-size: 28px; font-weight: bold; margin: 5px 0;">{dados.get('total_registros', 0)}</p>
            <p style="font-size: 12px; color: #999; margin: 0;">Analistas no período</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        cor_csat = "#2ecc71" if dados.get('csat_medio', 0) >= meta_csat else "#e74c3c"
        st.markdown(f"""
        <div style="background: #ffffff; padding: 20px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); border-left: 4px solid {cor_csat};">
            <p style="font-size: 14px; color: #666; margin: 0;">⭐ CSAT Médio</p>
            <p style="font-size: 28px; font-weight: bold; margin: 5px 0; color: {cor_csat};">{dados.get('csat_medio', 0):.2f}%</p>
            <p style="font-size: 12px; color: #999; margin: 0;">Meta: {meta_csat}%</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        cor_avaliacoes = "#2ecc71" if dados.get('perc_avaliacoes_medio', 0) >= meta_avaliacoes else "#e74c3c"
        st.markdown(f"""
        <div style="background: #ffffff; padding: 20px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); border-left: 4px solid {cor_avaliacoes};">
            <p style="font-size: 14px; color: #666; margin: 0;">📝 % Avaliações Médio</p>
            <p style="font-size: 28px; font-weight: bold; margin: 5px 0; color: {cor_avaliacoes};">{dados.get('perc_avaliacoes_medio', 0):.2f}%</p>
            <p style="font-size: 12px; color: #999; margin: 0;">Meta: {meta_avaliacoes}%</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        cor_envio = "#2ecc71" if dados.get('perc_envio_medio', 0) >= meta_envio else "#e74c3c"
        st.markdown(f"""
        <div style="background: #ffffff; padding: 20px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); border-left: 4px solid {cor_envio};">
            <p style="font-size: 14px; color: #666; margin: 0;">📤 % Envio Médio</p>
            <p style="font-size: 28px; font-weight: bold; margin: 5px 0; color: {cor_envio};">{dados.get('perc_envio_medio', 0):.2f}%</p>
            <p style="font-size: 12px; color: #999; margin: 0;">Meta: {meta_envio}%</p>
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
            <p style="font-size: 24px; margin: 0;">{cor_csat}</p>
            <p style="font-size: 18px; font-weight: bold; margin: 0;">CSAT</p>
            <p style="font-size: 14px; color: #333;">{csat_medio:.2f}%</p>
            <p style="font-size: 12px; color: #666;">Status: {status_csat}</p>
            <p style="font-size: 11px; color: #999;">Meta: {meta_csat}%</p>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div style="background: {bg_avaliacoes}; padding: 15px; border-radius: 10px; text-align: center;">
            <p style="font-size: 24px; margin: 0;">{cor_avaliacoes}</p>
            <p style="font-size: 18px; font-weight: bold; margin: 0;">Avaliações</p>
            <p style="font-size: 14px; color: #333;">{perc_avaliacoes_medio:.2f}%</p>
            <p style="font-size: 12px; color: #666;">Status: {status_avaliacoes}</p>
            <p style="font-size: 11px; color: #999;">Meta: {meta_avaliacoes}%</p>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown(f"""
        <div style="background: {bg_envio}; padding: 15px; border-radius: 10px; text-align: center;">
            <p style="font-size: 24px; margin: 0;">{cor_envio}</p>
            <p style="font-size: 18px; font-weight: bold; margin: 0;">Envio</p>
            <p style="font-size: 14px; color: #333;">{perc_envio_medio:.2f}%</p>
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
        height=350,
        margin=dict(l=20, r=20, t=40, b=20),
        yaxis=dict(title='Percentual (%)', range=[0, 100], gridcolor='#f0f0f0'),
        xaxis=dict(title='Período', gridcolor='#f0f0f0'),
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5),
        hovermode='x unified',
        plot_bgcolor='white'
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
    def get_ordem(mes):
        nome_mes = mes.split()[0] if mes else mes
        return ordem_meses.get(nome_mes, 13)
    return sorted(meses, key=get_ordem)

# ============================================
# DASHBOARD GESTOR
# ============================================

def dashboard_gestor(periodo, gestor_nome, supabase):
    """Dashboard específica para Gestor - mostra apenas sua equipe"""
    
    # Carregar dados do período e gestor
    df_historico = carregar_historico(supabase, mes_ano=periodo, gestor=gestor_nome)
    
    if df_historico is None or df_historico.empty:
        st.warning(f"⚠️ Nenhum dado encontrado para {gestor_nome} no período {periodo}")
        return None
    
    # Converter para dicionário
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
    
    # Calcular métricas agregadas
    total_analistas = len(resultados)
    total_atendimentos = sum([d['total_atendimentos'] for d in resultados.values()])
    media_atendimentos = total_atendimentos / total_analistas if total_analistas > 0 else 0
    csat_medio = sum([d['csat'] for d in resultados.values()]) / total_analistas if total_analistas > 0 else 0
    perc_avaliacoes_medio = sum([d['perc_avaliacoes'] for d in resultados.values()]) / total_analistas if total_analistas > 0 else 0
    perc_envio_medio = sum([d['perc_envio'] for d in resultados.values()]) / total_analistas if total_analistas > 0 else 0
    metas_superadas = len([d for d in resultados.values() if d['status'] == '🟢 Meta Superada'])
    
    # CARDS DE INDICADORES
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
    
    # SAÚDE DA EQUIPE
    criar_saude_equipe(csat_medio, perc_avaliacoes_medio, perc_envio_medio, meta_csat, meta_avaliacoes, meta_envio)
    
    # GRÁFICOS DE EVOLUÇÃO
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
    
    # TOP E BOTTOM PERFORMERS
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
    
    # TABELA DE DESEMPENHO
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
        
        fig_avaliacoes = criar_grafico_evolucao(df_mensal, 'perc_avaliacoes', '% Avaliações Médio - Operação', '#3498db', meta=meta_avaliacoes, meta_label=f'Meta: {meta_avaliacoes}%')
        if fig_avaliacoes:
            st.plotly_chart(fig_avaliacoes, use_container_width=True)
        
        fig_envio = criar_grafico_evolucao(df_mensal, 'perc_envio', '% Envio Médio - Operação', '#f39c12', meta=meta_envio, meta_label=f'Meta: {meta_envio}%')
        if fig_envio:
            st.plotly_chart(fig_envio, use_container_width=True)
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

# ============================================
# FUNÇÕES DE RELATÓRIOS (RESUMIDAS)
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
# FUNÇÕES DE PAINEL E GRÁFICO MENSAL
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
    fig.update_layout(xaxis_title='Período', yaxis=dict(title='CSAT (%)', range=[0, 100], side='left'), yaxis2=dict(title='% Avaliações', range=[0, 100], overlaying='y', side='right'), hovermode='x unified', legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1), height=450, template='plotly_white')
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
        fig_barras.add_trace(go.Bar(x=['CSAT'], y=[dados['csat']], name='Alcançado', marker_color='#2ecc71', text=[f"{dados['csat']:.1f}%"], textposition='outside'))
        fig_barras.add_trace(go.Bar(x=['CSAT'], y=[dados['meta_csat']], name='Meta', marker_color='#e74c3c', text=[f"{dados['meta_csat']}%"], textposition='outside'))
        fig_barras.update_layout(title='CSAT - Resultado vs Meta', yaxis_title='Percentual (%)', yaxis_range=[0, 100], height=350, showlegend=True, legend=dict(orientation='h', yanchor='bottom', y=1.02))
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
                'threshold': {'line': {'color': "red", 'width': 4}, 'thickness': 0.75, 'value': dados['meta_csat']}
            }
        ))
        fig_gauge.update_layout(height=250)
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
        fig_radar.add_trace(go.Scatterpolar(r=valores_analista, theta=categorias, fill='toself', name=analista, line_color='#2ecc71', fillcolor='rgba(46, 204, 113, 0.3)'))
        fig_radar.add_trace(go.Scatterpolar(r=valores_meta, theta=categorias, fill='toself', name='Meta', line_color='#e74c3c', fillcolor='rgba(231, 76, 60, 0.1)', line_dash='dash'))
        fig_radar.update_layout(title='Radar de Performance', polar=dict(radialaxis=dict(visible=True, range=[0, 100])), height=400, showlegend=True, legend=dict(orientation='h', yanchor='bottom', y=1.02))
        st.plotly_chart(fig_radar, use_container_width=True)
    
    st.markdown("---")
    return posicao_podio

# ============================================
# FUNÇÃO PARA GERAR RELATÓRIO INDIVIDUAL
# ============================================

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
# INTERFACE PRINCIPAL
# ============================================

def main():
    st.title("📊 Sistema de Performance - Relatórios Automáticos")
    st.markdown("---")
    
    if not fazer_login():
        st.info("👋 Faça login na barra lateral para acessar o sistema.")
        return
    
    if st.session_state.get('cadastrar_usuario', False):
        cadastrar_usuario()
        return
    
    if st.session_state.get('gerenciar_usuarios', False):
        gerenciar_usuarios_supabase()
        st.markdown("---")
        return
    
    limpar_estado_sessao()
    supabase = init_supabase()
    
    if supabase:
        st.sidebar.success("✅ Conectado ao Supabase")
    else:
        st.sidebar.warning("⚠️ Supabase não configurado")
    
    analistas_config = carregar_analistas()
    acesso_total = st.session_state.get('acesso_total', False)
    gestor_ativo = st.session_state.get('gestor', None)
    nome_usuario = st.session_state.get('nome_usuario', '')
    
    # SIDEBAR
    with st.sidebar:
        st.header("📁 Upload de Arquivos")
        arquivo_satisfacao = st.file_uploader("Arquivo de Satisfação", type=['xlsx', 'csv'])
        arquivo_inativos = st.file_uploader("Arquivo de Inatividade", type=['xlsx', 'csv'])
        
        st.markdown("---")
        st.header("📅 Período")
        opcao_periodo = st.radio("Como definir o período?", ["Selecionar manualmente", "Extrair do arquivo"], index=0)
        
        if opcao_periodo == "Selecionar manualmente":
            periodo_manual = st.text_input("Informe o período", placeholder="Ex: Maio 2026", value="")
            periodo_selecionado = periodo_manual if periodo_manual else None
        else:
            periodo_selecionado = None
            st.info("ℹ️ O período será extraído automaticamente do arquivo.")
        
        periodo_valido = periodo_selecionado is not None and periodo_selecionado.strip() != ""
        
        st.markdown("---")
        if st.button("🚀 Processar Dados", use_container_width=True):
            if not periodo_valido and opcao_periodo == "Selecionar manualmente":
                st.error("⚠️ Selecione o período antes de processar o relatório.")
            elif not arquivo_satisfacao or not arquivo_inativos:
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
                        
                        if opcao_periodo == "Extrair do arquivo":
                            periodo = extrair_periodo(df_satisfacao)
                        else:
                            periodo = periodo_manual
                        
                        st.session_state.periodo = periodo
                        gestor_salvar = st.session_state.get('gestor', GESTOR_MARCOS)
                        
                        if supabase:
                            existe = verificar_periodo_existente(supabase, periodo, gestor_salvar)
                            if existe:
                                st.warning(f"⚠️ Já existe relatório para {periodo}.")
                                col1, col2 = st.columns(2)
                                with col1:
                                    if st.button("❌ Cancelar", use_container_width=True):
                                        st.stop()
                                with col2:
                                    if st.button("🔄 Substituir", use_container_width=True):
                                        resultados = processar_dados(df_satisfacao, df_inativos, analistas_config)
                                        st.session_state.resultados = resultados
                                        st.session_state.processado = True
                                        sucesso, mensagem = substituir_historico(supabase, resultados, periodo, gestor_salvar)
                                        if sucesso:
                                            st.success(f"✅ Período {periodo} substituído!")
                                        else:
                                            st.error(f"❌ Erro: {mensagem}")
                                        st.rerun()
                                st.stop()
                        
                        resultados = processar_dados(df_satisfacao, df_inativos, analistas_config)
                        st.session_state.resultados = resultados
                        st.session_state.processado = True
                        
                        if supabase:
                            sucesso, mensagem = salvar_historico(supabase, resultados, periodo, gestor_salvar)
                            if sucesso:
                                st.success(f"✅ Dados salvos! Período: {periodo}")
                            else:
                                st.warning(f"⚠️ Dados NÃO salvos: {mensagem}")
                        else:
                            st.success(f"✅ Dados processados! Período: {periodo}")
                    except Exception as e:
                        st.error(f"❌ Erro: {str(e)}")
        
        st.markdown("---")
        st.header("📂 Gerenciar Períodos")
        if st.button("📊 Ver Períodos Importados", use_container_width=True):
            st.session_state.mostrar_periodos = True
        
        st.markdown("---")
        st.header("📊 Consultar Histórico")
        if st.button("📊 Ver Histórico", use_container_width=True):
            st.session_state.mostrar_historico = True
        
        st.markdown("---")
        st.header("⚙️ Gerenciar")
        if st.button("📝 Gerenciar Analistas", use_container_width=True):
            st.session_state.gerenciar_analistas = True
    
    # ===== GERENCIAR ANALISTAS =====
    if st.session_state.get('gerenciar_analistas', False):
        st.header("📝 Gerenciar Analistas")
        gestor_selecionado = st.selectbox("Selecione o Gestor", list(analistas_config.keys()))
        if gestor_selecionado:
            config = analistas_config[gestor_selecionado]
            st.write(f"**{gestor_selecionado}** - {len(config['membros'])} analistas:")
            for analista, dados in config['membros'].items():
                st.write(f"- {analista} (Meta: {dados['meta_csat']}%, Ativo: {dados['ativo']})")
        
        if st.button("🔙 Voltar", key="voltar_gerenciar_analistas"):
            st.session_state.gerenciar_analistas = False
            st.rerun()
        st.markdown("---")
    
    # ===== GERENCIAR PERÍODOS =====
    if st.session_state.get('mostrar_periodos', False):
        st.header("📂 Gerenciar Períodos Importados")
        if supabase:
            if acesso_total:
                periodos = listar_periodos(supabase)
            else:
                periodos = listar_periodos(supabase, gestor_ativo)
            
            if periodos:
                dados_tabela = []
                for p in periodos:
                    df_periodo = carregar_historico(supabase, mes_ano=p['mes_ano'], gestor=p['gestor'])
                    qtd_analistas = len(df_periodo) if df_periodo is not None else 0
                    dados_tabela.append({'Período': p['mes_ano'], 'Gestor': p['gestor'], 'Analistas': qtd_analistas})
                st.dataframe(pd.DataFrame(dados_tabela), use_container_width=True, hide_index=True)
            else:
                st.info("Nenhum período encontrado.")
        else:
            st.error("❌ Supabase não configurado.")
        
        if st.button("🔙 Voltar", key="voltar_periodos"):
            st.session_state.mostrar_periodos = False
            st.rerun()
        st.markdown("---")
    
    # ===== CONSULTAR HISTÓRICO =====
    if st.session_state.get('mostrar_historico', False):
        st.header("📊 Consultar Histórico")
        if supabase:
            if acesso_total:
                df_historico = carregar_historico(supabase)
            else:
                df_historico = carregar_historico(supabase, gestor=gestor_ativo)
            
            if df_historico is not None and not df_historico.empty:
                st.dataframe(df_historico, use_container_width=True)
            else:
                st.warning("Nenhum dado histórico encontrado.")
        else:
            st.error("❌ Supabase não configurado.")
        
        if st.button("🔙 Voltar", key="voltar_historico"):
            st.session_state.mostrar_historico = False
            st.rerun()
        st.markdown("---")
    
    # ============================================
    # DASHBOARD PRINCIPAL
    # ============================================
    
    if st.session_state.get('processado', False) and not st.session_state.get('gerenciar_analistas', False) and not st.session_state.get('mostrar_historico', False) and not st.session_state.get('mostrar_periodos', False):
        
        periodo = st.session_state.get('periodo', datetime.now().strftime('%B %Y'))
        
        if supabase:
            if acesso_total:
                periodos_disponiveis = listar_periodos(supabase)
            else:
                periodos_disponiveis = listar_periodos(supabase, gestor_ativo)
            
            if periodos_disponiveis:
                periodos_lista = sorted([p['mes_ano'] for p in periodos_disponiveis], reverse=True)
                periodo_selecionado = st.selectbox("📅 Selecione o Período para visualizar", periodos_lista, index=0 if periodo in periodos_lista else 0, key="seletor_periodo_dashboard")
                if periodo_selecionado != periodo:
                    st.session_state.periodo = periodo_selecionado
                    st.rerun()
        
        if acesso_total:
            df_historico = carregar_historico(supabase, mes_ano=st.session_state.periodo)
            st.info("🔑 Perfil: Coordenador - Visualizando toda a operação")
            if df_historico is not None and not df_historico.empty:
                dashboard_coordenador(st.session_state.periodo, nome_usuario, supabase)
            else:
                st.warning(f"⚠️ Nenhum dado encontrado para o período {st.session_state.periodo}")
        else:
            st.info(f"👥 Perfil: Gestor - Visualizando equipe: {gestor_ativo}")
            df_historico = carregar_historico(supabase, mes_ano=st.session_state.periodo, gestor=gestor_ativo)
            
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
                dashboard_gestor(st.session_state.periodo, gestor_ativo, supabase)
            else:
                st.warning(f"⚠️ Nenhum dado encontrado para {gestor_ativo} no período {st.session_state.periodo}")
                periodos_gestor = listar_periodos(supabase, gestor_ativo)
                if periodos_gestor:
                    periodos_ordenados = sorted(periodos_gestor, key=lambda x: x['mes_ano'], reverse=True)
                    ultimo_periodo = periodos_ordenados[0]
                    st.info(f"📅 Carregando último período disponível: {ultimo_periodo['mes_ano']}")
                    st.session_state.periodo = ultimo_periodo['mes_ano']
                    st.rerun()
            
            # EXIBIR DASHBOARD POR PERFIL
            if acesso_total:
                st.info("🔑 Perfil: Coordenador - Visualizando toda a operação")
                dashboard_coordenador(st.session_state.periodo, nome_usuario, supabase)
            else:
                st.info(f"👥 Perfil: Gestor - Visualizando equipe: {gestor_ativo}")
                resultados = dashboard_gestor(st.session_state.periodo, gestor_ativo, supabase)
                
                if resultados:
                    st.markdown("---")
                    st.subheader("🏆 Pódio do Mês")
                    
                    media_atendimentos = calcular_media_operacao(resultados)
                    podio = calcular_podio(resultados, media_atendimentos)
                    
                    if supabase and not acesso_total:
                        try:
                            podio_manual = carregar_podio_manual(supabase, st.session_state.periodo, gestor_ativo)
                            if podio_manual:
                                podio = podio_manual
                        except:
                            pass
                    
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
                        st.info("🏆 Nenhum analista atingiu todos os critérios do pódio neste mês.")
                    
                    st.markdown("---")
                    st.subheader("📄 Gerar Relatório Individual")
                    analista_selecionado = st.selectbox("Selecione o Analista", list(resultados.keys()))
                    if analista_selecionado:
                        dados = resultados[analista_selecionado]
                        gerar_relatorio_individual(analista_selecionado, dados, media_atendimentos, podio, st.session_state.periodo, supabase, gestor_ativo, acesso_total)
    
    else:
        if not st.session_state.get('gerenciar_analistas', False) and not st.session_state.get('mostrar_historico', False) and not st.session_state.get('mostrar_periodos', False):
            if acesso_total:
                st.info("📊 Bem-vindo, Coordenador! Faça upload dos arquivos ou consulte o histórico.")
            else:
                st.info(f"📊 Bem-vindo, Gestor! Faça upload dos arquivos para sua equipe: {gestor_ativo}")
            
            if supabase:
                if acesso_total:
                    periodos = listar_periodos(supabase)
                else:
                    periodos = listar_periodos(supabase, gestor_ativo)
                
                if periodos:
                    periodos_ordenados = sorted(periodos, key=lambda x: x['mes_ano'], reverse=True)
                    ultimo_periodo = periodos_ordenados[0]
                    st.info(f"📅 Carregando último período disponível: {ultimo_periodo['mes_ano']}")
                    
                    df_historico = carregar_historico(supabase, mes_ano=ultimo_periodo['mes_ano'], gestor=ultimo_periodo['gestor'])
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
                        st.session_state.periodo = ultimo_periodo['mes_ano']
                        st.rerun()
    
    st.markdown("---")
    st.markdown('<div style="text-align: center; color: #666; font-size: 12px;">Sistema de Performance v11.0 | UX Melhorado + Dashboard Visual + Evolução de Indicadores</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
