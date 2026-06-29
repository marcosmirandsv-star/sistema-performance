import streamlit as st
import pandas as pd
import numpy as np
from supabase import create_client, Client
from datetime import datetime, timedelta
import bcrypt
import plotly.express as px
import plotly.graph_objects as go
from streamlit_option_menu import option_menu
import json
import io
import base64
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
import xlsxwriter
import time
import re
from typing import List, Dict, Any, Optional

# ==================== CONFIGURAÇÃO DA PÁGINA ====================
st.set_page_config(
    page_title="Sistema de Gestão de Performance",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== ESTILO CSS PERSONALIZADO ====================
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: 700;
        color: #1f3a8a;
        margin-bottom: 1rem;
        padding: 1rem;
        background: linear-gradient(135deg, #e0e7ff, #c7d2fe);
        border-radius: 0.5rem;
    }
    .sub-header {
        font-size: 1.5rem;
        font-weight: 600;
        color: #1e40af;
        margin-top: 1rem;
        margin-bottom: 0.5rem;
    }
    .metric-card {
        background: white;
        padding: 1.5rem;
        border-radius: 0.75rem;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);
        border-left: 4px solid #3b82f6;
    }
    .podium-gold {
        background: linear-gradient(135deg, #fbbf24, #f59e0b);
        padding: 1.5rem;
        border-radius: 0.75rem;
        color: white;
        text-align: center;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);
    }
    .podium-silver {
        background: linear-gradient(135deg, #9ca3af, #6b7280);
        padding: 1.5rem;
        border-radius: 0.75rem;
        color: white;
        text-align: center;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);
    }
    .podium-bronze {
        background: linear-gradient(135deg, #d97706, #92400e);
        padding: 1.5rem;
        border-radius: 0.75rem;
        color: white;
        text-align: center;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);
    }
    .status-badge {
        padding: 0.25rem 0.75rem;
        border-radius: 9999px;
        font-size: 0.875rem;
        font-weight: 500;
    }
    .status-active {
        background-color: #d1fae5;
        color: #065f46;
    }
    .status-inactive {
        background-color: #fee2e2;
        color: #991b1b;
    }
    .info-box {
        background-color: #eff6ff;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #3b82f6;
        margin: 0.5rem 0;
    }
    .warning-box {
        background-color: #fffbeb;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #f59e0b;
        margin: 0.5rem 0;
    }
    .success-box {
        background-color: #f0fdf4;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #22c55e;
        margin: 0.5rem 0;
    }
</style>
""", unsafe_allow_html=True)

# ==================== INICIALIZAÇÃO DO SUPABASE ====================
try:
    SUPABASE_URL = st.secrets["SUPABASE_URL"]
    SUPABASE_KEY = st.secrets["SUPABASE_KEY"]
    supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
    # Teste de conexão
    test_response = supabase.table("usuarios").select("count").limit(1).execute()
except Exception as e:
    st.error(f"""
    ❌ **Erro ao conectar ao Supabase**
    
    Verifique:
    - Se as credenciais estão corretas no arquivo `.streamlit/secrets.toml`
    - Se as tabelas foram criadas no Supabase
    - Se a URL e KEY estão corretas
    
    **Erro técnico:** {str(e)}
    """)
    st.stop()

# ==================== FUNÇÕES DE AUTENTICAÇÃO ====================

def hash_password(password: str) -> str:
    """Cria hash da senha usando bcrypt"""
    try:
        salt = bcrypt.gensalt(rounds=12)
        return bcrypt.hashpw(password.encode('utf-8'), salt).decode('utf-8')
    except Exception as e:
        st.error(f"Erro ao gerar hash da senha: {str(e)}")
        return None

def verify_password(password: str, hashed: str) -> bool:
    """Verifica se a senha corresponde ao hash"""
    try:
        if not password or not hashed:
            return False
        return bcrypt.checkpw(password.encode('utf-8'), hashed.encode('utf-8'))
    except Exception as e:
        return False

def login_user(username: str, password: str) -> Optional[Dict]:
    """Realiza login do usuário"""
    try:
        # Buscar usuário
        response = supabase.table("usuarios").select("*").eq("usuario", username).execute()
        
        if not response.data:
            return None
            
        user = response.data[0]
        
        # Verificar se usuário está ativo
        if not user.get('ativo', True):
            st.warning("⚠️ Usuário inativo. Entre em contato com o administrador.")
            return None
            
        # Verificar senha
        if verify_password(password, user.get('senha', '')):
            # Removido: atualização de ultimo_login porque a coluna não existe
            # supabase.table("usuarios").update({
            #     "ultimo_login": datetime.now().isoformat()
            # }).eq("id", user['id']).execute()
            
            # Registrar log (opcional)
            # registrar_log(user['id'], "LOGIN", f"Login realizado com sucesso")
            
            return user
        
        return None
        
    except Exception as e:
        st.error(f"Erro ao fazer login: {str(e)}")
        return None
        
    except Exception as e:
        st.error(f"Erro ao fazer login: {str(e)}")
        return None

def get_user_by_id(user_id: int) -> Optional[Dict]:
    """Busca usuário por ID"""
    try:
        response = supabase.table("usuarios").select("*").eq("id", user_id).execute()
        return response.data[0] if response.data else None
    except Exception as e:
        st.error(f"Erro ao buscar usuário: {str(e)}")
        return None

def get_all_users() -> List[Dict]:
    """Busca todos os usuários ativos e inativos"""
    try:
        response = supabase.table("usuarios").select("*").order("nome").execute()
        return response.data
    except Exception as e:
        st.error(f"Erro ao buscar usuários: {str(e)}")
        return []

def get_active_users() -> List[Dict]:
    """Busca apenas usuários ativos"""
    try:
        response = supabase.table("usuarios").select("*").eq("ativo", True).order("nome").execute()
        return response.data
    except Exception as e:
        st.error(f"Erro ao buscar usuários ativos: {str(e)}")
        return []

def create_user(
    usuario: str, 
    nome: str, 
    senha: str, 
    cargo: str, 
    time_nome: str, 
    supervisor_nome: str, 
    ativo: bool = True
) -> Optional[Dict]:
    """Cria um novo usuário"""
    try:
        # Validar dados
        if not usuario or not nome or not senha:
            st.error("❌ Todos os campos obrigatórios devem ser preenchidos!")
            return None
            
        # Verificar se usuário já existe
        existing = supabase.table("usuarios").select("*").eq("usuario", usuario).execute()
        if existing.data:
            st.error("❌ Usuário já existe!")
            return None
            
        # Validar e-mail se for fornecido como usuário
        if '@' in usuario and not re.match(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$', usuario):
            st.warning("⚠️ Formato de e-mail inválido!")
            
        # Gerar hash da senha
        hashed = hash_password(senha)
        if not hashed:
            st.error("❌ Erro ao gerar hash da senha!")
            return None
        
        # Preparar dados
        data = {
            "usuario": usuario,
            "nome": nome,
            "senha": hashed,
            "cargo": cargo,
            "time_nome": time_nome,
            "supervisor_nome": supervisor_nome,
            "ativo": ativo,
            "criado_em": datetime.now().isoformat(),
            "criado_por": st.session_state.get('user', {}).get('id', None)
        }
        
        # Inserir no banco
        response = supabase.table("usuarios").insert(data).execute()
        
        if response.data:
            # Registrar log
            registrar_log(response.data[0]['id'], "CRIACAO", f"Usuário {nome} criado com sucesso")
            st.success(f"✅ Usuário {nome} criado com sucesso!")
            return response.data[0]
        
        st.error("❌ Erro ao criar usuário!")
        return None
        
    except Exception as e:
        st.error(f"❌ Erro ao criar usuário: {str(e)}")
        return None

def update_user(user_id: int, data: Dict) -> Optional[Dict]:
    """Atualiza um usuário"""
    try:
        # Se tiver senha, fazer hash
        if 'senha' in data and data['senha']:
            hashed = hash_password(data['senha'])
            if not hashed:
                st.error("❌ Erro ao gerar hash da nova senha!")
                return None
            data['senha'] = hashed
        else:
            data.pop('senha', None)
        
        # Adicionar data de atualização
        data['atualizado_em'] = datetime.now().isoformat()
        data['atualizado_por'] = st.session_state.get('user', {}).get('id', None)
        
        # Atualizar no banco
        response = supabase.table("usuarios").update(data).eq("id", user_id).execute()
        
        if response.data:
            # Registrar log
            registrar_log(user_id, "ATUALIZACAO", "Dados do usuário atualizados")
            st.success("✅ Usuário atualizado com sucesso!")
            return response.data[0]
        
        st.error("❌ Erro ao atualizar usuário!")
        return None
        
    except Exception as e:
        st.error(f"❌ Erro ao atualizar usuário: {str(e)}")
        return None

def delete_user(user_id: int) -> bool:
    """Exclui um usuário"""
    try:
        # Buscar nome do usuário para log
        user = get_user_by_id(user_id)
        nome = user.get('nome', 'Unknown') if user else 'Unknown'
        
        # Registrar log antes de excluir
        registrar_log(user_id, "EXCLUSAO", f"Usuário {nome} excluído do sistema")
        
        # Excluir do banco
        response = supabase.table("usuarios").delete().eq("id", user_id).execute()
        
        st.success(f"✅ Usuário {nome} excluído com sucesso!")
        return True
        
    except Exception as e:
        st.error(f"❌ Erro ao excluir usuário: {str(e)}")
        return False

def registrar_log(usuario_id: int, acao: str, descricao: str):
    """Registra log de atividades"""
    try:
        data = {
            "usuario_id": usuario_id,
            "acao": acao,
            "descricao": descricao,
            "data_hora": datetime.now().isoformat(),
            "ip": st.session_state.get('ip', '0.0.0.0')
        }
        supabase.table("logs").insert(data).execute()
    except Exception as e:
        # Não interrompe a execução em caso de erro no log
        pass

def get_logs(usuario_id: Optional[int] = None, limit: int = 100) -> List[Dict]:
    """Busca logs de atividades"""
    try:
        query = supabase.table("logs").select("*").order("data_hora", desc=True).limit(limit)
        if usuario_id:
            query = query.eq("usuario_id", usuario_id)
        response = query.execute()
        return response.data
    except Exception as e:
        return []

# ==================== FUNÇÕES DE DADOS ====================

def get_indicadores(time_nome: Optional[str] = None, mes: Optional[str] = None, ano: Optional[int] = None) -> List[Dict]:
    """Busca indicadores de performance"""
    try:
        query = supabase.table("indicadores").select("*")
        
        if time_nome:
            query = query.eq("time_nome", time_nome)
        if mes:
            query = query.eq("mes", mes)
        if ano:
            query = query.eq("ano", ano)
            
        response = query.order("data_criacao", desc=True).execute()
        return response.data
    except Exception as e:
        st.error(f"Erro ao buscar indicadores: {str(e)}")
        return []

def get_historico(time_nome: Optional[str] = None, mes: Optional[str] = None, ano: Optional[int] = None) -> List[Dict]:
    """Busca histórico de performance"""
    try:
        query = supabase.table("historico").select("*")
        
        if time_nome:
            query = query.eq("time_nome", time_nome)
        if mes:
            query = query.eq("mes", mes)
        if ano:
            query = query.eq("ano", ano)
            
        response = query.order("data_criacao", desc=True).execute()
        return response.data
    except Exception as e:
        st.error(f"Erro ao buscar histórico: {str(e)}")
        return []

def get_podio(time_nome: Optional[str] = None, mes: Optional[str] = None, ano: Optional[int] = None) -> List[Dict]:
    """Busca dados do pódio"""
    try:
        query = supabase.table("podio").select("*")
        
        if time_nome:
            query = query.eq("time_nome", time_nome)
        if mes:
            query = query.eq("mes", mes)
        if ano:
            query = query.eq("ano", ano)
            
        response = query.order("pontuacao", desc=True).execute()
        return response.data
    except Exception as e:
        st.error(f"Erro ao buscar pódio: {str(e)}")
        return []

def get_podio_manual(time_nome: str, mes: str, ano: int) -> Optional[Dict]:
    """Busca pódio manual"""
    try:
        response = supabase.table("podio_manual").select("*")\
            .eq("time_nome", time_nome)\
            .eq("mes", mes)\
            .eq("ano", ano)\
            .execute()
        return response.data[0] if response.data else None
    except Exception as e:
        st.error(f"Erro ao buscar pódio manual: {str(e)}")
        return None

def save_podio_manual(time_nome: str, mes: str, ano: int, dados: str) -> bool:
    """Salva pódio manual"""
    try:
        existing = get_podio_manual(time_nome, mes, ano)
        
        data = {
            "time_nome": time_nome,
            "mes": mes,
            "ano": ano,
            "dados": dados,
            "atualizado_em": datetime.now().isoformat(),
            "atualizado_por": st.session_state.get('user', {}).get('id', None)
        }
        
        if existing:
            response = supabase.table("podio_manual").update(data)\
                .eq("time_nome", time_nome)\
                .eq("mes", mes)\
                .eq("ano", ano)\
                .execute()
        else:
            data["criado_em"] = datetime.now().isoformat()
            response = supabase.table("podio_manual").insert(data).execute()
            
        st.success("✅ Pódio manual salvo com sucesso!")
        return True
        
    except Exception as e:
        st.error(f"❌ Erro ao salvar pódio manual: {str(e)}")
        return False

def get_exclusoes(time_nome: str, mes: str, ano: int) -> List[Dict]:
    """Busca exclusões do pódio"""
    try:
        response = supabase.table("podio_exclusoes").select("*")\
            .eq("time_nome", time_nome)\
            .eq("mes", mes)\
            .eq("ano", ano)\
            .execute()
        return response.data
    except Exception as e:
        st.error(f"Erro ao buscar exclusões: {str(e)}")
        return []

def add_exclusao(time_nome: str, mes: str, ano: int, analista: str) -> bool:
    """Adiciona exclusão de analista do pódio"""
    try:
        # Verificar se já existe
        existing = supabase.table("podio_exclusoes").select("*")\
            .eq("time_nome", time_nome)\
            .eq("mes", mes)\
            .eq("ano", ano)\
            .eq("analista", analista)\
            .execute()
            
        if existing.data:
            st.warning("⚠️ Este analista já está excluído do pódio!")
            return False
            
        data = {
            "time_nome": time_nome,
            "mes": mes,
            "ano": ano,
            "analista": analista,
            "criado_em": datetime.now().isoformat(),
            "criado_por": st.session_state.get('user', {}).get('id', None)
        }
        
        response = supabase.table("podio_exclusoes").insert(data).execute()
        st.success(f"✅ Analista {analista} excluído do pódio!")
        return True
        
    except Exception as e:
        st.error(f"❌ Erro ao adicionar exclusão: {str(e)}")
        return False

def remove_exclusao(exclusao_id: int) -> bool:
    """Remove exclusão do pódio"""
    try:
        response = supabase.table("podio_exclusoes").delete().eq("id", exclusao_id).execute()
        st.success("✅ Exclusão removida com sucesso!")
        return True
    except Exception as e:
        st.error(f"❌ Erro ao remover exclusão: {str(e)}")
        return False

def importar_periodo(mes: str, ano: int, time_nome: str, arquivo) -> bool:
    """Importa dados de um período"""
    try:
        # Ler arquivo Excel
        df = pd.read_excel(arquivo)
        
        # Validar colunas obrigatórias
        colunas_obrigatorias = ['analista', 'performance', 'produtividade', 'qualidade']
        for col in colunas_obrigatorias:
            if col not in df.columns:
                st.error(f"❌ Coluna '{col}' não encontrada no arquivo!")
                return False
        
        # Processar cada linha
        total_importados = 0
        for _, row in df.iterrows():
            data = {
                "time_nome": time_nome,
                "mes": mes,
                "ano": ano,
                "analista": row.get('analista', ''),
                "performance": float(row.get('performance', 0)),
                "produtividade": float(row.get('produtividade', 0)),
                "qualidade": float(row.get('qualidade', 0)),
                "metricas": row.to_dict(),
                "data_importacao": datetime.now().isoformat(),
                "importado_por": st.session_state.get('user', {}).get('id', None)
            }
            
            # Inserir no histórico
            supabase.table("historico").insert(data).execute()
            total_importados += 1
        
        st.success(f"✅ {total_importados} registros importados com sucesso!")
        return True
        
    except Exception as e:
        st.error(f"❌ Erro ao importar período: {str(e)}")
        return False

def get_periodos_disponiveis() -> List[Dict]:
    """Busca períodos disponíveis no histórico"""
    try:
        response = supabase.table("historico").select("mes", "ano", "time_nome")\
            .distinct("mes", "ano", "time_nome")\
            .order("ano", desc=True)\
            .order("mes", desc=True)\
            .execute()
        return response.data
    except Exception as e:
        return []

# ==================== FUNÇÕES DE DASHBOARD ====================

def criar_dashboard_metricas(df: pd.DataFrame, titulo: str):
    """Cria dashboard com métricas e gráficos"""
    if df.empty:
        st.warning("📊 Sem dados para exibir")
        return
    
    # Métricas
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <p style="color: #6b7280; font-size: 0.875rem;">Total Analistas</p>
            <h2 style="color: #1f2937; margin: 0;">{len(df)}</h2>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        if 'performance' in df.columns:
            media = df['performance'].mean()
            st.markdown(f"""
            <div class="metric-card">
                <p style="color: #6b7280; font-size: 0.875rem;">Performance Média</p>
                <h2 style="color: #1f2937; margin: 0;">{media:.1f}%</h2>
            </div>
            """, unsafe_allow_html=True)
    
    with col3:
        if 'produtividade' in df.columns:
            media = df['produtividade'].mean()
            st.markdown(f"""
            <div class="metric-card">
                <p style="color: #6b7280; font-size: 0.875rem;">Produtividade Média</p>
                <h2 style="color: #1f2937; margin: 0;">{media:.1f}</h2>
            </div>
            """, unsafe_allow_html=True)
    
    with col4:
        if 'qualidade' in df.columns:
            media = df['qualidade'].mean()
            st.markdown(f"""
            <div class="metric-card">
                <p style="color: #6b7280; font-size: 0.875rem;">Qualidade Média</p>
                <h2 style="color: #1f2937; margin: 0;">{media:.1f}%</h2>
            </div>
            """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # Gráficos
    col1, col2 = st.columns(2)
    
    with col1:
        if 'performance' in df.columns and 'analista' in df.columns:
            fig = px.bar(
                df, 
                x='analista', 
                y='performance',
                title=f'📊 Performance por Analista',
                color='performance',
                color_continuous_scale='Viridis',
                text='performance'
            )
            fig.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
            fig.update_layout(
                xaxis_title="Analista",
                yaxis_title="Performance (%)",
                height=400,
                showlegend=False
            )
            st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        if 'produtividade' in df.columns and 'qualidade' in df.columns and 'analista' in df.columns:
            fig = go.Figure()
            fig.add_trace(go.Bar(
                name='Produtividade',
                x=df['analista'],
                y=df['produtividade'],
                marker_color='#3b82f6',
                text=df['produtividade'],
                textposition='outside'
            ))
            fig.add_trace(go.Bar(
                name='Qualidade',
                x=df['analista'],
                y=df['qualidade'],
                marker_color='#22c55e',
                text=df['qualidade'],
                textposition='outside'
            ))
            fig.update_layout(
                title='📈 Produtividade vs Qualidade',
                xaxis_title="Analista",
                yaxis_title="Valor",
                barmode='group',
                height=400,
                legend=dict(
                    yanchor="top",
                    y=0.99,
                    xanchor="left",
                    x=0.01
                )
            )
            st.plotly_chart(fig, use_container_width=True)
    
    # Gráfico de pizza - Distribuição por performance
    if 'performance' in df.columns:
        st.subheader("🎯 Distribuição de Performance")
        
        # Criar categorias
        df['categoria'] = pd.cut(
            df['performance'],
            bins=[0, 70, 85, 100],
            labels=['Baixa (<70%)', 'Média (70-85%)', 'Alta (>85%)']
        )
        
        fig = px.pie(
            df,
            names='categoria',
            values='performance',
            title='Distribuição por Categoria de Performance',
            color='categoria',
            color_discrete_map={
                'Baixa (<70%)': '#ef4444',
                'Média (70-85%)': '#f59e0b',
                'Alta (>85%)': '#22c55e'
            }
        )
        fig.update_traces(textposition='inside', textinfo='percent+label')
        st.plotly_chart(fig, use_container_width=True)

def mostrar_podio(time_nome: str, mes: str, ano: int, is_coordenador: bool = False, supervisor_selecionado: str = None):
    """Exibe o pódio com todas as funcionalidades"""
    st.markdown(f"### 🏆 Pódio - {mes}/{ano}")
    
    # Definir filtro de time
    if is_coordenador and supervisor_selecionado and supervisor_selecionado != "Todos":
        time_filtro = "Time Marcos" if "Marcos" in supervisor_selecionado else "Time Polyana"
    else:
        time_filtro = time_nome
    
    # Buscar dados
    podio_data = get_podio(time_filtro, mes, ano)
    podio_manual = get_podio_manual(time_filtro, mes, ano)
    exclusoes = get_exclusoes(time_filtro, mes, ano)
    
    # Aplicar exclusões
    analistas_excluidos = [e['analista'] for e in exclusoes]
    
    if podio_data:
        df_podio = pd.DataFrame(podio_data)
        
        # Filtrar excluídos
        if analistas_excluidos:
            df_podio = df_podio[~df_podio['analista'].isin(analistas_excluidos)]
        
        # Ordenar por pontuação
        df_podio = df_podio.sort_values('pontuacao', ascending=False)
        
        # Exibir pódio
        col1, col2, col3 = st.columns([1, 2, 1])
        
        # Medalhas
        medalhas = ['🥇', '🥈', '🥉']
        cores = ['podium-gold', 'podium-silver', 'podium-bronze']
        
        for i in range(min(3, len(df_podio))):
            if i == 0:
                with col2:
                    st.markdown(f"""
                    <div class="{cores[i]}">
                        <h1 style="font-size: 3rem;">{medalhas[i]}</h1>
                        <h2>{df_podio.iloc[i]['analista']}</h2>
                        <p style="font-size: 1.5rem;">{df_podio.iloc[i]['pontuacao']:.1f} pts</p>
                    </div>
                    """, unsafe_allow_html=True)
            elif i == 1:
                with col1:
                    st.markdown(f"""
                    <div class="{cores[i]}">
                        <h1 style="font-size: 2.5rem;">{medalhas[i]}</h1>
                        <h3>{df_podio.iloc[i]['analista']}</h3>
                        <p style="font-size: 1.2rem;">{df_podio.iloc[i]['pontuacao']:.1f} pts</p>
                    </div>
                    """, unsafe_allow_html=True)
            else:
                with col3:
                    st.markdown(f"""
                    <div class="{cores[i]}">
                        <h1 style="font-size: 2.5rem;">{medalhas[i]}</h1>
                        <h3>{df_podio.iloc[i]['analista']}</h3>
                        <p style="font-size: 1.2rem;">{df_podio.iloc[i]['pontuacao']:.1f} pts</p>
                    </div>
                    """, unsafe_allow_html=True)
        
        # Ranking completo
        with st.expander("📋 Ranking Completo"):
            # Adicionar coluna de posição
            df_podio['Posição'] = range(1, len(df_podio) + 1)
            
            # Formatar para exibição
            display_df = df_podio[['Posição', 'analista', 'pontuacao']].copy()
            display_df.columns = ['Posição', 'Analista', 'Pontuação']
            
            # Aplicar cores para os 3 primeiros
            def color_posicao(val):
                if val <= 3:
                    return ['background-color: #fef3c7'] * len(display_df.columns)
                return [''] * len(display_df.columns)
            
            st.dataframe(
                display_df.style.apply(color_posicao, axis=1),
                use_container_width=True,
                hide_index=True
            )
            
            # Exportar ranking
            csv = display_df.to_csv(index=False)
            st.download_button(
                "📥 Exportar Ranking (CSV)",
                csv,
                f"ranking_{time_filtro}_{mes}_{ano}.csv",
                "text/csv",
                key="export_ranking"
            )
    
    # Pódio Manual
    if podio_manual:
        with st.expander("📝 Pódio Manual"):
            st.json(podio_manual['dados'])
            
            col1, col2 = st.columns(2)
            with col1:
                if st.button("✏️ Editar Pódio Manual", key="edit_manual_btn"):
                    st.session_state.edit_manual = True
            
            if st.session_state.get('edit_manual', False):
                novo_dados = st.text_area(
                    "Dados do Pódio Manual (JSON)",
                    value=podio_manual['dados'],
                    height=200,
                    key="manual_data"
                )
                
                col1, col2 = st.columns(2)
                with col1:
                    if st.button("💾 Salvar Pódio Manual"):
                        # Validar JSON
                        try:
                            json.loads(novo_dados)
                            if save_podio_manual(time_filtro, mes, ano, novo_dados):
                                st.session_state.edit_manual = False
                                st.rerun()
                        except:
                            st.error("❌ JSON inválido! Verifique a formatação.")
                
                with col2:
                    if st.button("❌ Cancelar"):
                        st.session_state.edit_manual = False
                        st.rerun()
    
    # Gerenciar Exclusões
    with st.expander("🚫 Gerenciar Exclusões"):
        st.markdown("### Adicionar Exclusão")
        
        col1, col2 = st.columns([3, 1])
        with col1:
            analista = st.text_input("Nome do Analista", key="excluir_analista_input")
        with col2:
            st.write("")
            if st.button("➕ Adicionar", key="add_exclusao_btn"):
                if analista:
                    if add_exclusao(time_filtro, mes, ano, analista):
                        st.rerun()
                else:
                    st.warning("⚠️ Digite o nome do analista!")
        
        # Listar exclusões
        if exclusoes:
            st.markdown("### Analistas Excluídos")
            
            for exc in exclusoes:
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.write(f"🔴 {exc['analista']}")
                with col2:
                    if st.button("🗑️", key=f"remover_{exc['id']}"):
                        if remove_exclusao(exc['id']):
                            st.rerun()
        else:
            st.info("Nenhuma exclusão cadastrada para este período.")

# ==================== FUNÇÕES DE RELATÓRIOS ====================

def exportar_relatorio_pdf(df: pd.DataFrame, titulo: str):
    """Exporta relatório em PDF"""
    try:
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4)
        elements = []
        
        # Estilos
        styles = getSampleStyleSheet()
        title_style = styles['Title']
        heading_style = styles['Heading2']
        
        # Título
        elements.append(Paragraph(f"Relatório de Performance - {titulo}", title_style))
        elements.append(Spacer(1, 20))
        
        # Data
        elements.append(Paragraph(f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", styles['Normal']))
        elements.append(Spacer(1, 20))
        
        # Tabela de dados
        if not df.empty:
            # Preparar dados para tabela
            data = [df.columns.tolist()] + df.values.tolist()
            
            table = Table(data)
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 12),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            
            elements.append(table)
        
        # Construir PDF
        doc.build(elements)
        buffer.seek(0)
        
        return buffer
        
    except Exception as e:
        st.error(f"❌ Erro ao gerar PDF: {str(e)}")
        return None

def exportar_relatorio_excel(df: pd.DataFrame, titulo: str):
    """Exporta relatório em Excel"""
    try:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            # Escrever dados
            df.to_excel(writer, sheet_name='Dados', index=False)
            
            # Formatar
            workbook = writer.book
            worksheet = writer.sheets['Dados']
            
            # Adicionar título
            worksheet.write(0, 0, f"Relatório de Performance - {titulo}")
            
            # Formatar colunas
            for i, col in enumerate(df.columns):
                column_width = max(df[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(i, i, column_width)
        
        buffer.seek(0)
        return buffer
        
    except Exception as e:
        st.error(f"❌ Erro ao gerar Excel: {str(e)}")
        return None

# ==================== FUNÇÕES DE GERENCIAMENTO ====================

def gerenciar_usuarios():
    """Tela de gerenciamento de usuários"""
    st.markdown('<h1 class="main-header">👤 Gerenciamento de Usuários</h1>', unsafe_allow_html=True)
    
    # Estatísticas rápidas
    users = get_all_users()
    if users:
        total = len(users)
        ativos = sum(1 for u in users if u.get('ativo', False))
        inativos = total - ativos
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total de Usuários", total)
        with col2:
            st.metric("Usuários Ativos", ativos)
        with col3:
            st.metric("Usuários Inativos", inativos)
    
    st.markdown("---")
    
    # Tabs
    tab1, tab2, tab3 = st.tabs(["➕ Criar Usuário", "✏️ Editar Usuário", "📋 Listar Usuários"])
    
    with tab1:
        with st.form("criar_usuario_form"):
            st.markdown("### 📝 Dados do Usuário")
            
            col1, col2 = st.columns(2)
            with col1:
                usuario = st.text_input("Usuário (login) *", placeholder="exemplo@email.com")
                nome = st.text_input("Nome completo *", placeholder="Nome Sobrenome")
                senha = st.text_input("Senha *", type="password")
                confirmar_senha = st.text_input("Confirmar Senha *", type="password")
            
            with col2:
                cargo = st.selectbox("Cargo *", ["Supervisor", "Coordenador"])
                
                if cargo == "Coordenador":
                    time_nome = "Operacao"
                    supervisor_nome = "Carine Melo"
                    st.info("ℹ️ Coordenador terá acesso a toda operação")
                else:
                    time_nome = st.selectbox("Time *", ["Time Marcos", "Time Polyana"])
                    supervisor_nome = st.selectbox("Supervisor *", ["Marcos Miranda", "Polyana Ventura"])
                
                ativo = st.checkbox("Ativo", value=True)
            
            st.markdown("---")
            
            if st.form_submit_button("🚀 Criar Usuário"):
                # Validações
                if not usuario or not nome or not senha:
                    st.error("❌ Preencha todos os campos obrigatórios (*)")
                elif senha != confirmar_senha:
                    st.error("❌ As senhas não coincidem!")
                elif len(senha) < 6:
                    st.error("❌ A senha deve ter pelo menos 6 caracteres!")
                else:
                    user = create_user(usuario, nome, senha, cargo, time_nome, supervisor_nome, ativo)
                    if user:
                        st.rerun()
    
    with tab2:
        users = get_all_users()
        if users:
            user_options = {f"{u['nome']} ({u['usuario']})": u['id'] for u in users}
            selected_user = st.selectbox("Selecionar usuário para editar", list(user_options.keys()))
            
            if selected_user:
                user_id = user_options[selected_user]
                user = get_user_by_id(user_id)
                
                if user:
                    with st.form("editar_usuario_form"):
                        st.markdown(f"### ✏️ Editando: {user['nome']}")
                        
                        col1, col2 = st.columns(2)
                        with col1:
                            nome = st.text_input("Nome", value=user['nome'])
                            usuario = st.text_input("Usuário", value=user['usuario'], disabled=True)
                            nova_senha = st.text_input("Nova senha (deixe em branco para manter)", type="password")
                        
                        with col2:
                            cargo = st.selectbox("Cargo", ["Supervisor", "Coordenador"], 
                                               index=0 if user['cargo'] == "Supervisor" else 1)
                            
                            if cargo == "Coordenador":
                                time_nome = "Operacao"
                                supervisor_nome = "Carine Melo"
                            else:
                                time_nome = st.selectbox("Time", ["Time Marcos", "Time Polyana"],
                                                        index=0 if user['time_nome'] == "Time Marcos" else 1)
                                supervisor_nome = st.selectbox("Supervisor", 
                                    ["Marcos Miranda", "Polyana Ventura"],
                                    index=0 if user['supervisor_nome'] == "Marcos Miranda" else 1)
                            
                            ativo = st.checkbox("Ativo", value=user['ativo'])
                        
                        st.markdown("---")
                        
                        col1, col2 = st.columns(2)
                        with col1:
                            if st.form_submit_button("💾 Atualizar"):
                                update_data = {
                                    "nome": nome,
                                    "cargo": cargo,
                                    "time_nome": time_nome,
                                    "supervisor_nome": supervisor_nome,
                                    "ativo": ativo
                                }
                                
                                if nova_senha:
                                    if len(nova_senha) < 6:
                                        st.error("❌ A senha deve ter pelo menos 6 caracteres!")
                                        return
                                    update_data["senha"] = nova_senha
                                
                                if update_user(user_id, update_data):
                                    st.rerun()
                        
                        with col2:
                            if st.form_submit_button("🗑️ Excluir Usuário", type="primary"):
                                if st.checkbox("Confirmar exclusão permanente"):
                                    if delete_user(user_id):
                                        st.rerun()
        else:
            st.info("Nenhum usuário cadastrado")
    
    with tab3:
        users = get_all_users()
        if users:
            # Preparar dados para exibição
            df = pd.DataFrame(users)
            
            # Selecionar colunas para exibição
            display_cols = ['nome', 'usuario', 'cargo', 'time_nome', 'supervisor_nome', 'ativo', 'criado_em']
            df_display = df[display_cols].copy()
            
            # Renomear colunas
            df_display.columns = ['Nome', 'Usuário', 'Cargo', 'Time', 'Supervisor', 'Ativo', 'Criado em']
            
            # Formatar data
            df_display['Criado em'] = pd.to_datetime(df_display['Criado em']).dt.strftime('%d/%m/%Y')
            
            # Formatar status
            df_display['Ativo'] = df_display['Ativo'].apply(
                lambda x: '✅ Ativo' if x else '❌ Inativo'
            )
            
            st.dataframe(
                df_display,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Ativo": st.column_config.TextColumn("Status"),
                }
            )
            
            # Exportar
            csv = df_display.to_csv(index=False)
            st.download_button(
                "📥 Exportar Lista (CSV)",
                csv,
                "usuarios.csv",
                "text/csv"
            )
        else:
            st.info("Nenhum usuário cadastrado")

# ==================== FUNÇÕES DE LOGS ====================

def visualizar_logs():
    """Tela de visualização de logs"""
    st.markdown('<h1 class="main-header">📋 Logs do Sistema</h1>', unsafe_allow_html=True)
    
    # Filtros
    col1, col2 = st.columns(2)
    with col1:
        usuario_id = st.selectbox(
            "Filtrar por Usuário",
            [None] + [u['id'] for u in get_all_users()],
            format_func=lambda x: "Todos" if x is None else get_user_by_id(x)['nome'] if get_user_by_id(x) else str(x)
        )
    with col2:
        limite = st.slider("Quantidade de registros", 10, 500, 100)
    
    # Buscar logs
    logs = get_logs(usuario_id, limite)
    
    if logs:
        df_logs = pd.DataFrame(logs)
        
        # Formatar para exibição
        display_df = df_logs[['data_hora', 'descricao', 'acao']].copy()
        display_df['data_hora'] = pd.to_datetime(display_df['data_hora']).dt.strftime('%d/%m/%Y %H:%M')
        display_df.columns = ['Data/Hora', 'Descrição', 'Ação']
        
        st.dataframe(
            display_df,
            use_container_width=True,
            hide_index=True
        )
        
        # Exportar
        csv = display_df.to_csv(index=False)
        st.download_button(
            "📥 Exportar Logs (CSV)",
            csv,
            f"logs_{datetime.now().strftime('%Y%m%d')}.csv",
            "text/csv"
        )
    else:
        st.info("Nenhum log encontrado")

# ==================== NAVEGAÇÃO PRINCIPAL ====================

def mostrar_visao_coordenador(menu: str):
    """Mostra visão do coordenador"""
    if menu == "Dashboard":
        st.markdown('<h1 class="main-header">📊 Dashboard - Operação Completa</h1>', unsafe_allow_html=True)
        
        # Filtros
        col1, col2, col3 = st.columns(3)
        with col1:
            mes = st.selectbox("Mês", 
                ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
                 "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"],
                index=datetime.now().month - 1,
                key="coord_mes"
            )
        with col2:
            ano = st.selectbox("Ano", [2024, 2025, 2026], index=2, key="coord_ano")
        with col3:
            supervisor = st.selectbox("Supervisor", ["Todos", "Marcos Miranda", "Polyana Ventura"], key="coord_sup")
        
        mes_ano_display = f"{mes} {ano}"
        st.info(f"📅 **Período:** {mes_ano_display}")
        
        # Buscar dados
        with st.spinner("Carregando dados..."):
            if supervisor == "Todos":
                dados = get_indicadores(mes=mes, ano=ano)
            else:
                time_nome = "Time Marcos" if "Marcos" in supervisor else "Time Polyana"
                dados = get_indicadores(time_nome=time_nome, mes=mes, ano=ano)
        
        if dados:
            df = pd.DataFrame(dados)
            titulo = f"{mes_ano_display}" + (f" - {supervisor}" if supervisor != "Todos" else " - Todos os Times")
            criar_dashboard_metricas(df, titulo)
        else:
            st.warning("📊 Nenhum dado encontrado para o período selecionado")
    
    elif menu == "Histórico":
        st.markdown('<h1 class="main-header">📜 Histórico - Operação Completa</h1>', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        with col1:
            supervisor = st.selectbox("Filtrar por Supervisor", ["Todos", "Marcos Miranda", "Polyana Ventura"], key="hist_sup")
        with col2:
            periodo = st.selectbox("Período", ["Últimos 30 dias", "Últimos 90 dias", "Último ano", "Todos"], key="hist_per")
        
        with st.spinner("Carregando histórico..."):
            if supervisor == "Todos":
                historico = get_historico()
            else:
                time_nome = "Time Marcos" if "Marcos" in supervisor else "Time Polyana"
                historico = get_historico(time_nome=time_nome)
        
        if historico:
            df = pd.DataFrame(historico)
            
            # Filtrar por período
            if periodo != "Todos":
                df['data_criacao'] = pd.to_datetime(df['data_criacao'])
                hoje = datetime.now()
                if periodo == "Últimos 30 dias":
                    df = df[df['data_criacao'] > hoje - timedelta(days=30)]
                elif periodo == "Últimos 90 dias":
                    df = df[df['data_criacao'] > hoje - timedelta(days=90)]
                elif periodo == "Último ano":
                    df = df[df['data_criacao'] > hoje - timedelta(days=365)]
            
            # Exibir
            display_df = df[['time_nome', 'mes', 'ano', 'analista', 'performance', 'data_criacao']].copy()
            display_df['data_criacao'] = pd.to_datetime(display_df['data_criacao']).dt.strftime('%d/%m/%Y')
            display_df.columns = ['Time', 'Mês', 'Ano', 'Analista', 'Performance', 'Data']
            
            st.dataframe(
                display_df,
                use_container_width=True,
                hide_index=True
            )
            
            # Botões de exportação
            col1, col2 = st.columns(2)
            with col1:
                # Exportar CSV
                csv = display_df.to_csv(index=False)
                st.download_button(
                    "📥 Exportar CSV",
                    csv,
                    f"historico_{datetime.now().strftime('%Y%m%d')}.csv",
                    "text/csv"
                )
            
            with col2:
                # Exportar Excel
                excel_buffer = exportar_relatorio_excel(display_df, "Histórico")
                if excel_buffer:
                    st.download_button(
                        "📥 Exportar Excel",
                        excel_buffer,
                        f"historico_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        else:
            st.info("Nenhum histórico encontrado")
    
    elif menu == "Importação":
        st.markdown('<h1 class="main-header">📤 Importar Período</h1>', unsafe_allow_html=True)
        
        with st.form("importar_form"):
            col1, col2 = st.columns(2)
            with col1:
                mes = st.selectbox("Mês *", 
                    ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
                     "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"],
                    index=datetime.now().month - 1,
                    key="imp_mes"
                )
            with col2:
                ano = st.selectbox("Ano *", [2024, 2025, 2026], index=2, key="imp_ano")
            
            time_nome = st.selectbox("Time *", ["Time Marcos", "Time Polyana"], key="imp_time")
            arquivo = st.file_uploader("Selecione o arquivo Excel *", type=['xlsx', 'xls'], key="imp_file")
            
            # Template de exemplo
            with st.expander("📄 Modelo de arquivo"):
                st.markdown("""
                O arquivo deve conter as seguintes colunas:
                - **analista**: Nome do analista
                - **performance**: Percentual de performance (0-100)
                - **produtividade**: Score de produtividade
                - **qualidade**: Percentual de qualidade (0-100)
                
                Exemplo:
                | analista | performance | produtividade | qualidade |
                |----------|-------------|---------------|-----------|
                | João Silva | 85.5 | 12.3 | 92.0 |
                | Maria Santos | 78.0 | 11.5 | 88.5 |
                """)
            
            st.markdown("---")
            
            if st.form_submit_button("🚀 Importar"):
                if not arquivo:
                    st.error("❌ Selecione um arquivo!")
                else:
                    if importar_periodo(mes, ano, time_nome, arquivo):
                        st.success("✅ Dados importados com sucesso!")
                        st.rerun()
    
    elif menu == "Pódio":
        st.markdown('<h1 class="main-header">🏆 Pódio - Operação Completa</h1>', unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        with col1:
            supervisor = st.selectbox("Supervisor", ["Todos", "Marcos Miranda", "Polyana Ventura"], key="pod_sup")
        with col2:
            mes = st.selectbox("Mês", 
                ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
                 "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"],
                index=datetime.now().month - 1,
                key="pod_mes"
            )
        with col3:
            ano = st.selectbox("Ano", [2024, 2025, 2026], index=2, key="pod_ano")
        
        if supervisor == "Todos":
            # Pódio Geral
            st.markdown("### 🏆 Pódio Geral da Operação")
            mostrar_podio("Operacao", mes, ano, True, "Todos")
            
            st.markdown("---")
            
            # Pódio por Time
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("### 🏆 Time Marcos")
                mostrar_podio("Time Marcos", mes, ano, True, "Marcos")
            
            with col2:
                st.markdown("### 🏆 Time Polyana")
                mostrar_podio("Time Polyana", mes, ano, True, "Polyana")
        else:
            time_nome = "Time Marcos" if "Marcos" in supervisor else "Time Polyana"
            mostrar_podio(time_nome, mes, ano, True, supervisor)
    
    elif menu == "Usuários":
        gerenciar_usuarios()
    
    elif menu == "Logs":
        visualizar_logs()

def mostrar_visao_supervisor(menu: str):
    """Mostra visão do supervisor"""
    time_nome = st.session_state.time_nome
    supervisor_nome = st.session_state.supervisor_nome
    
    if menu == "Dashboard":
        st.markdown(f'<h1 class="main-header">📊 Dashboard - {time_nome}</h1>', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        with col1:
            mes = st.selectbox("Mês", 
                ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
                 "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"],
                index=datetime.now().month - 1,
                key="sup_mes"
            )
        with col2:
            ano = st.selectbox("Ano", [2024, 2025, 2026], index=2, key="sup_ano")
        
        mes_ano_display = f"{mes} {ano}"
        st.info(f"📅 **Período:** {mes_ano_display}")
        
        with st.spinner("Carregando dados..."):
            dados = get_indicadores(time_nome=time_nome, mes=mes, ano=ano)
        
        if dados:
            df = pd.DataFrame(dados)
            criar_dashboard_metricas(df, f"{time_nome} - {mes_ano_display}")
        else:
            st.warning("📊 Nenhum dado encontrado para o período selecionado")
    
    elif menu == "Histórico":
        st.markdown(f'<h1 class="main-header">📜 Histórico - {time_nome}</h1>', unsafe_allow_html=True)
        
        with st.spinner("Carregando histórico..."):
            historico = get_historico(time_nome=time_nome)
        
        if historico:
            df = pd.DataFrame(historico)
            
            # Exibir
            display_df = df[['mes', 'ano', 'analista', 'performance', 'data_criacao']].copy()
            display_df['data_criacao'] = pd.to_datetime(display_df['data_criacao']).dt.strftime('%d/%m/%Y')
            display_df.columns = ['Mês', 'Ano', 'Analista', 'Performance', 'Data']
            
            st.dataframe(
                display_df,
                use_container_width=True,
                hide_index=True
            )
            
            # Botões de exportação
            col1, col2 = st.columns(2)
            with col1:
                csv = display_df.to_csv(index=False)
                st.download_button(
                    "📥 Exportar CSV",
                    csv,
                    f"historico_{time_nome}_{datetime.now().strftime('%Y%m%d')}.csv",
                    "text/csv"
                )
            
            with col2:
                excel_buffer = exportar_relatorio_excel(display_df, f"Histórico - {time_nome}")
                if excel_buffer:
                    st.download_button(
                        "📥 Exportar Excel",
                        excel_buffer,
                        f"historico_{time_nome}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        else:
            st.info("Nenhum histórico encontrado")
    
    elif menu == "Importação":
        st.markdown(f'<h1 class="main-header">📤 Importar Período - {time_nome}</h1>', unsafe_allow_html=True)
        
        with st.form("importar_sup_form"):
            col1, col2 = st.columns(2)
            with col1:
                mes = st.selectbox("Mês *", 
                    ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
                     "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"],
                    index=datetime.now().month - 1,
                    key="sup_imp_mes"
                )
            with col2:
                ano = st.selectbox("Ano *", [2024, 2025, 2026], index=2, key="sup_imp_ano")
            
            st.info(f"📌 Importando para: **{time_nome}**")
            arquivo = st.file_uploader("Selecione o arquivo Excel *", type=['xlsx', 'xls'], key="sup_imp_file")
            
            st.markdown("---")
            
            if st.form_submit_button("🚀 Importar"):
                if not arquivo:
                    st.error("❌ Selecione um arquivo!")
                else:
                    if importar_periodo(mes, ano, time_nome, arquivo):
                        st.success("✅ Dados importados com sucesso!")
                        st.rerun()
    
    elif menu == "Pódio":
        st.markdown(f'<h1 class="main-header">🏆 Pódio - {time_nome}</h1>', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        with col1:
            mes = st.selectbox("Mês", 
                ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
                 "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"],
                index=datetime.now().month - 1,
                key="sup_pod_mes"
            )
        with col2:
            ano = st.selectbox("Ano", [2024, 2025, 2026], index=2, key="sup_pod_ano")
        
        mostrar_podio(time_nome, mes, ano)
    
    elif menu == "Usuários":
        st.warning("🔒 Apenas coordenadores podem gerenciar usuários")
    
    elif menu == "Logs":
        visualizar_logs()

# ==================== MAIN ====================

def main():
    """Função principal do sistema"""
    
    # Verificar se está logado
    if 'logged_in' not in st.session_state:
        st.session_state.logged_in = False
    
    # Tela de Login
    if not st.session_state.logged_in:
        st.markdown("""
        <div style="text-align: center; padding: 2rem;">
            <h1 style="font-size: 3rem; color: #1f3a8a;">📊 Sistema de Gestão de Performance</h1>
            <p style="font-size: 1.2rem; color: #6b7280;">Faça login para acessar o sistema</p>
        </div>
        """, unsafe_allow_html=True)
        
        with st.container():
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                with st.form("login_form"):
                    st.markdown("### 🔐 Login")
                    
                    usuario = st.text_input("Usuário", placeholder="Digite seu usuário")
                    senha = st.text_input("Senha", type="password", placeholder="Digite sua senha")
                    
                    st.markdown("---")
                    
                    if st.form_submit_button("🚀 Entrar", use_container_width=True):
                        if usuario and senha:
                            with st.spinner("Autenticando..."):
                                user = login_user(usuario, senha)
                                if user:
                                    st.session_state.logged_in = True
                                    st.session_state.user = user
                                    st.session_state.cargo = user['cargo']
                                    st.session_state.time_nome = user['time_nome']
                                    st.session_state.supervisor_nome = user['supervisor_nome']
                                    st.success(f"✅ Bem-vindo, {user['nome']}!")
                                    time.sleep(1)
                                    st.rerun()
                                else:
                                    st.error("❌ Usuário ou senha inválidos")
                        else:
                            st.warning("⚠️ Preencha todos os campos")
        
        # Informações de acesso
        with st.expander("ℹ️ Informações de Acesso"):
            st.markdown("""
            **Credenciais padrão:**
            
            **Coordenadora:**
            - Usuário: carine
            - Senha: carine123
            
            **Supervisores:**
            - Usuário: marcos
            - Senha: marcos123
            
            - Usuário: polyana
            - Senha: polyana123
            """)
        
        return
    
    # ===== MENU LATERAL =====
    with st.sidebar:
        # Cabeçalho do usuário
        st.markdown(f"""
        <div style="text-align: center; padding: 1rem; background: #f3f4f6; border-radius: 0.5rem;">
            <p style="font-size: 1.2rem; font-weight: 600;">👋 {st.session_state.user['nome']}</p>
            <p style="color: #6b7280;">📌 {st.session_state.cargo}</p>
            <p style="color: #6b7280;">🏢 {st.session_state.time_nome}</p>
        </div>
        """, unsafe_allow_html=True)
        
        st.divider()
        
        # Menu de navegação
        menu_options = ["Dashboard", "Histórico", "Importação", "Pódio"]
        menu_icons = ["📊", "📜", "📤", "🏆"]
        
        # Adicionar opções extras para coordenador
        if st.session_state.cargo == "Coordenador":
            menu_options.append("Usuários")
            menu_options.append("Logs")
            menu_icons.append("👤")
            menu_icons.append("📋")
        
        selected = option_menu(
            menu_title="MENU",
            options=menu_options,
            icons=menu_icons,
            default_index=0,
            styles={
                "container": {"padding": "0!important", "background-color": "#fafafa"},
                "icon": {"color": "orange", "font-size": "20px"},
                "nav-link": {
                    "font-size": "16px",
                    "text-align": "left",
                    "margin": "0px",
                    "--hover-color": "#eee"
                },
                "nav-link-selected": {"background-color": "#ff4b4b"},
            }
        )
        
        st.divider()
        
        # Botão de logout
        if st.button("🚪 Sair", use_container_width=True):
            # Registrar logout
            if 'user' in st.session_state:
                registrar_log(st.session_state.user['id'], "LOGOUT", "Usuário saiu do sistema")
            
            # Limpar sessão
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()
        
        # Versão
        st.caption("Versão 2.0.0 | © 2026")
    
    # ===== CONTEÚDO PRINCIPAL =====
    # Mostrar visão baseada no cargo
    if st.session_state.cargo == "Coordenador":
        mostrar_visao_coordenador(selected)
    else:
        mostrar_visao_supervisor(selected)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        st.error(f"❌ Erro inesperado: {str(e)}")
        st.info("🔄 Recarregue a página ou entre em contato com o administrador.")
