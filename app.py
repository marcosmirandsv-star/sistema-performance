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
        
        # Deletar dados existentes do mesmo mês/gestor para não duplicar
        supabase.table('historico_performance').delete().eq('mes_ano', mes_ano).eq('gestor', gestor).execute()
        
        for registro in registros:
            supabase.table('historico_performance').insert(registro).execute()
        
        return True, "Salvo com sucesso"
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

# ============================================
# GERENCIAMENTO DE USUÁRIOS
# ============================================

def hash_senha(senha):
    """Cria hash da senha"""
    return hashlib.sha256(senha.encode()).hexdigest()

def carregar_usuarios():
    """Carrega usuários do Supabase ou arquivo local"""
    # Tentar carregar do Supabase primeiro
    supabase = init_supabase()
    if supabase:
        try:
            response = supabase.table('usuarios').select('*').execute()
            usuarios = {}
            for u in response.data:
                usuarios[u['usuario']] = {
                    'senha': u['senha'],
                    'nome': u['nome'],
                    'gestor': u['gestor']
                }
            return usuarios
        except:
            pass
    
    # Fallback para arquivo local
    try:
        if os.path.exists('usuarios.json'):
            with open('usuarios.json', 'r', encoding='utf-8') as f:
                return json.load(f)
    except:
        pass
    
    # Usuários padrão
    usuarios = {
        "marcos": {
            "senha": hash_senha("marcos2026"),
            "nome": "Marcos Miranda",
            "gestor": "Sua Gestão - Chat Notas"
        },
        "polyana": {
            "senha": hash_senha("polyana2026"),
            "nome": "Polyana Ventura",
            "gestor": "Gestão Polyana Ventura - Chat Outros"
        }
    }
    
    # Salvar localmente
    try:
        with open('usuarios.json', 'w', encoding='utf-8') as f:
            json.dump(usuarios, f, ensure_ascii=False, indent=2)
    except:
        pass
    
    return usuarios

def salvar_usuario_supabase(supabase, usuario, nome, senha_hash, gestor):
    """Salva usuário no Supabase"""
    if not supabase:
        return False
    
    try:
        supabase.table('usuarios').insert({
            'usuario': usuario,
            'nome': nome,
            'senha': senha_hash,
            'gestor': gestor
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
    
    # Configuração inicial
    config = {
        "Sua Gestão - Chat Notas": {
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
        "Gestão Polyana Ventura - Chat Outros": {
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
            st.rerun()
        else:
            st.sidebar.error("❌ Usuário ou senha inválidos!")
    
    if st.session_state.get('logado', False):
        st.sidebar.success(f"✅ Logado como {st.session_state.nome_usuario}")
        
        # Botão para cadastrar novos usuários
        if st.sidebar.button("👤 Cadastrar Usuário", use_container_width=True):
            st.session_state.cadastrar_usuario = True
        
        if st.sidebar.button("Sair", use_container_width=True):
            st.session_state.logado = False
            st.session_state.usuario = None
            st.session_state.nome_usuario = None
            st.session_state.gestor = None
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
            ["Sua Gestão - Chat Notas", "Gestão Polyana Ventura - Chat Outros"]
        )
    
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
            
            # Salvar no Supabase
            if supabase:
                if salvar_usuario_supabase(supabase, novo_usuario, novo_nome, senha_hash, gestor_usuario):
                    st.success(f"✅ Usuário {novo_usuario} cadastrado com sucesso!")
                    st.rerun()
                else:
                    st.error("❌ Erro ao salvar no Supabase!")
            else:
                # Salvar localmente
                usuarios[novo_usuario] = {
                    "senha": senha_hash,
                    "nome": novo_nome,
                    "gestor": gestor_usuario
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
        'Claudia', 'Almeida'
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
# GERAÇÃO DE RELATÓRIO INDIVIDUAL COM GRÁFICO
# ============================================

def gerar_grafico_mensal(analista, dados_mensais, meta_csat, meta_avaliacoes):
    """Gera gráfico mensal de evolução do analista"""
    
    if dados_mensais is None or dados_mensais.empty:
        return None
    
    # Filtrar dados do analista
    df_analista = dados_mensais[dados_mensais['analista'] == analista]
    
    if df_analista.empty:
        return None
    
    # Criar gráfico
    fig = go.Figure()
    
    # Adicionar CSAT
    fig.add_trace(go.Scatter(
        x=df_analista['mes_ano'],
        y=df_analista['csat'],
        name='CSAT Alcançado',
        mode='lines+markers',
        line=dict(color='#2ecc71', width=3),
        marker=dict(size=10)
    ))
    
    # Adicionar meta CSAT (linha constante)
    fig.add_hline(
        y=meta_csat, 
        line_dash="dash", 
        line_color="#e74c3c",
        annotation_text=f"Meta CSAT: {meta_csat}%",
        annotation_position="bottom right"
    )
    
    # Adicionar % Avaliações
    fig.add_trace(go.Bar(
        x=df_analista['mes_ano'],
        y=df_analista['perc_avaliacoes'],
        name='% Avaliações',
        marker_color='#3498db',
        opacity=0.7,
        yaxis='y2'
    ))
    
    # Adicionar meta avaliações (linha constante)
    fig.add_hline(
        y=meta_avaliacoes, 
        line_dash="dot", 
        line_color="#f39c12",
        annotation_text=f"Meta Avaliações: {meta_avaliacoes}%",
        annotation_position="top right",
        yaxis='y2'
    )
    
    # Configurar layout com dois eixos Y
    fig.update_layout(
        title=f'📈 Evolução Mensal - {analista}',
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
        height=450
    )
    
    return fig

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

def gerar_feedback_mimo(analista, dados, media_operacao, posicao_podio=None):
    """Gera feedback no padrão MIMO"""
    
    genero = get_genero_neutro(analista)
    
    texto_podio = ""
    if posicao_podio:
        texto_podio = f" 🏆 {posicao_podio}º lugar no pódio do mês!"
    
    if dados['status'] == "🟢 Meta Superada":
        abertura = f"Parabéns, {analista}! 🌟 Seu desempenho no último mês foi excelente{texto_podio}"
    elif dados['status'] == "🟡 Atenção":
        abertura = f"Olá, {analista}! Seu desempenho no último mês foi bom, mas temos alguns pontos para ajustar{texto_podio}"
    else:
        abertura = f"Olá, {analista}! Precisamos conversar sobre seu desempenho no último mês"
    
    indicadores = f"""
**📊 Indicadores de Performance:**

- **CSAT:** {dados['csat']:.2f}% (Meta: ≥ {dados['meta_csat']:.0f}%)
- **Delta CSAT:** {dados['delta_csat']:+.2f} pontos percentuais
- **% Avaliações:** {dados['perc_avaliacoes']:.2f}% (Meta: ≥ {dados['meta_geral']:.0f}%)
- **% Envio:** {dados['perc_envio']:.2f}%
- **💬 Atendimentos:** {dados['total_atendimentos']} chamados
- **Média da Operação:** {media_operacao} chamados
- **Avaliações:** {dados['avaliacoes']} ({dados['positivos']} positivas, {dados['negativos']} negativas)
"""
    
    pontos_fortes = []
    oportunidades = []
    
    if dados['csat'] >= dados['meta_csat'] + 2:
        pontos_fortes.append("CSAT acima da meta, demonstrando excelente qualidade no atendimento")
    elif dados['csat'] < dados['meta_csat']:
        oportunidades.append("CSAT abaixo da meta, foque em entender as causas das avaliações negativas")
    
    if dados['perc_avaliacoes'] >= dados['meta_geral'] + 5:
        pontos_fortes.append("alta taxa de engajamento, excelente trabalho na coleta de feedback")
    elif dados['perc_avaliacoes'] < dados['meta_geral']:
        oportunidades.append("taxa de avaliações abaixo da meta, lembre-se de sempre oferecer a pesquisa ao final do atendimento")
    
    if dados['total_atendimentos'] > media_operacao:
        pontos_fortes.append(f"produtividade {((dados['total_atendimentos']/media_operacao - 1)*100):.0f}% acima da média")
    elif dados['total_atendimentos'] < media_operacao * 0.8:
        oportunidades.append("produtividade abaixo da média, busque otimizar seu tempo entre os atendimentos")
    
    if pontos_fortes:
        mensagem_desenvolvimento = f"**🌟 Pontos Fortes:**\n- " + "\n- ".join(pontos_fortes)
    else:
        mensagem_desenvolvimento = "**🌟 Pontos Fortes:** Continue se dedicando ao atendimento de qualidade."
    
    if oportunidades:
        mensagem_desenvolvimento += f"\n\n**🎯 Oportunidades de Melhoria:**\n- " + "\n- ".join(oportunidades)
    
    if dados['status'] == "🟢 Meta Superada":
        orientacao = f"""
**📌 Próximos Passos:**

1. Continue com o excelente trabalho! 🚀
2. Compartilhe suas boas práticas com a equipe
3. Mantenha o foco na qualidade e engajamento
4. Meta para o próximo mês: manter CSAT ≥ {dados['meta_csat']:.0f}% e avaliações ≥ {dados['meta_geral']:.0f}%
"""
    elif dados['status'] == "🟡 Atenção":
        orientacao = f"""
**📌 Próximos Passos:**

1. {oportunidades[0] if oportunidades else 'Foque em melhorar sua performance geral'}
2. Agende uma conversa com seu gestor para alinhamento
3. Revise seus atendimentos com avaliações negativas
4. Meta para o próximo mês: CSAT ≥ {dados['meta_csat']:.0f}% e avaliações ≥ {dados['meta_geral']:.0f}%
"""
    else:
        orientacao = f"""
**📌 Próximos Passos:**

1. 🚨 Plano de ação urgente necessário
2. {oportunidades[0] if oportunidades else 'Revise completamente sua abordagem de atendimento'}
3. Acompanhamento diário com seu gestor
4. Meta para o próximo mês: CSAT ≥ {dados['meta_csat']:.0f}% e avaliações ≥ {dados['meta_geral']:.0f}%
"""
    
    feedback = f"""
## 📝 Feedback de Performance

### M - Mensagem de Abertura
{abertura}

### I - Indicadores de Performance
{indicadores}

### M - Mensagem de Desenvolvimento
{mensagem_desenvolvimento}

### O - Orientação e Próximos Passos
{orientacao}

---
**Status Geral:** {dados['status']}
**Data:** {datetime.now().strftime('%d/%m/%Y')}
"""
    return feedback.strip()

# ============================================
# GERAÇÃO DE RELATÓRIO WORD
# ============================================

def gerar_relatorio_word(analista, dados, analise_tecnica, feedback, media_operacao, podio, periodo, fig_grafico=None):
    """Gera relatório em formato Word (.docx) com gráfico"""
    
    doc = Document()
    
    # Título
    titulo = doc.add_heading(f'Relatório de Performance - {analista}', 0)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Período
    doc.add_paragraph(f'Período: {periodo}')
    doc.add_paragraph('')
    
    # Esperado
    doc.add_heading('Esperado:', level=1)
    doc.add_paragraph(f'≥ {dados["meta_geral"]:.0f}% de avaliações', style='List Bullet')
    doc.add_paragraph(f'≥ {dados["meta_csat"]:.0f}% de Satisfação', style='List Bullet')
    doc.add_paragraph('')
    
    # Atingido
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
    
    # Gráfico (se disponível)
    if fig_grafico:
        doc.add_heading('Gráfico de Evolução Mensal', level=1)
        doc.add_paragraph('(Gráfico disponível na versão digital do relatório)')
        doc.add_paragraph('')
    
    # Análise Técnica
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
    
    # Feedback
    doc.add_heading('Conclusão e Feedback', level=1)
    doc.add_paragraph(f'Status Geral do Período: {dados["status"]}')
    doc.add_paragraph('')
    
    for linha in feedback.split('\n'):
        if linha.strip():
            if linha.strip().startswith('##'):
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
# GERENCIAR ANALISTAS COM MOVER ENTRE TIMES
# ============================================

def gerenciar_analistas_completo(analistas_config):
    """Função para gerenciar analistas com opção de mover entre times"""
    
    st.header("📝 Gerenciar Analistas")
    
    # Selecionar gestor para editar
    gestor_selecionado = st.selectbox(
        "Selecione o Gestor para editar",
        list(analistas_config.keys())
    )
    
    if not gestor_selecionado:
        return
    
    config = analistas_config[gestor_selecionado]
    
    # Editar meta geral
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
    
    # ===== MOVER ANALISTA =====
    st.subheader("🔄 Mover Analista entre Times")
    
    # Listar todos os analistas ativos com seus gestores
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
            # Encontrar gestor atual
            gestor_atual = next((g for a, g in todos_analistas if a == analista_para_mover), None)
            
            # Listar todos os gestores disponíveis (incluindo o atual)
            todos_gestores = list(analistas_config.keys())
            
            # Mostrar gestor atual como padrão
            indice_atual = todos_gestores.index(gestor_atual) if gestor_atual in todos_gestores else 0
            
            novo_gestor = st.selectbox(
                f"Time atual: {gestor_atual}",
                todos_gestores,
                index=indice_atual
            )
        
        with col3:
            if novo_gestor and novo_gestor != gestor_atual:
                if st.button("🔄 Mover", use_container_width=True):
                    # Salvar meta atual
                    meta_atual = analistas_config[gestor_atual]['membros'][analista_para_mover]['meta_csat']
                    
                    # Remover do time atual
                    del analistas_config[gestor_atual]['membros'][analista_para_mover]
                    
                    # Adicionar ao novo time
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
    
    # ===== MEMBROS DO TIME =====
    st.subheader(f"👥 Membros do Time: {gestor_selecionado}")
    
    membros = config['membros']
    
    # Adicionar novo membro
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
    
    # Tabela de membros
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
    
    # Editar membro
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
    
    # 1. Gráfico de Barras - CSAT
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
    
    # 2. Gráfico de Barras - % Avaliações
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
    
    # 3. Gráfico de Barras - % Envio
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
    
    # 4. Gráfico de Dispersão - CSAT vs Atendimentos
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
    
    # 5. Gráfico de Barras Empilhadas - Distribuição por Status
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
    
    # Tabela resumo
    st.dataframe(
        status_agg[['Status', 'Quantidade', '%', 'CSAT Médio', '% Avaliações Médio', '% Envio Médio']],
        use_container_width=True,
        hide_index=True
    )
    
    # 6. Gráfico de Pizza
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
            type=['xlsx']
        )
        
        arquivo_inativos = st.file_uploader(
            "Arquivo de Inatividade",
            type=['xlsx']
        )
        
        st.markdown("---")
        st.header("📅 Período")
        
        periodo_manual = st.text_input(
            "Informe o período (ex: Maio 2026)",
            value=datetime.now().strftime('%B %Y')
        )
        
        st.markdown("---")
        
        if st.button("🚀 Processar Dados", use_container_width=True):
            if arquivo_satisfacao and arquivo_inativos:
                with st.spinner("Processando dados..."):
                    try:
                        df_satisfacao = pd.read_excel(arquivo_satisfacao)
                        df_inativos = pd.read_excel(arquivo_inativos)
                        
                        if periodo_manual:
                            periodo = periodo_manual
                        else:
                            periodo = extrair_periodo(df_satisfacao)
                        
                        st.session_state.periodo = periodo
                        
                        resultados = processar_dados(df_satisfacao, df_inativos, analistas_config)
                        st.session_state.resultados = resultados
                        st.session_state.processado = True
                        
                        # Salvar no Supabase
                        if supabase:
                            gestor_logado = st.session_state.gestor
                            sucesso, mensagem = salvar_historico(supabase, resultados, periodo, gestor_logado)
                            if sucesso:
                                st.success(f"✅ Dados salvos no Supabase! Período: {periodo}")
                            else:
                                st.warning(f"⚠️ Dados processados, mas NÃO salvos: {mensagem}")
                        else:
                            st.success(f"✅ Dados processados! Período: {periodo}")
                    except Exception as e:
                        st.error(f"❌ Erro ao processar dados: {str(e)}")
            else:
                st.error("❌ Envie os dois arquivos.")
        
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
        gerenciar_analistas_completo(analistas_config)
        
        if st.button("🔙 Voltar", use_container_width=True):
            st.session_state.gerenciar_analistas = False
            st.rerun()
        
        st.markdown("---")
    
    # ===== CONSULTAR HISTÓRICO =====
    if st.session_state.get('mostrar_historico', False):
        st.header("📊 Consultar Histórico")
        
        if supabase:
            # Carregar histórico
            df_historico = carregar_historico(supabase)
            
            if df_historico is not None and not df_historico.empty:
                # Filtros
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
                
                # Aplicar filtros
                df_filtrado = df_historico.copy()
                if mes_selecionado != 'Todos':
                    df_filtrado = df_filtrado[df_filtrado['mes_ano'] == mes_selecionado]
                if gestor_filtro != 'Todos':
                    df_filtrado = df_filtrado[df_filtrado['gestor'] == gestor_filtro]
                if analista_filtro != 'Todos':
                    df_filtrado = df_filtrado[df_filtrado['analista'] == analista_filtro]
                
                # Métricas
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
                
                # Gráfico de evolução
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
                
                # Tabela de dados
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
    if st.session_state.get('processado', False) and not st.session_state.get('gerenciar_analistas', False) and not st.session_state.get('mostrar_historico', False):
        resultados = st.session_state.resultados
        periodo = st.session_state.get('periodo', datetime.now().strftime('%B %Y'))
        
        gestor_logado = st.session_state.gestor
        resultados_gestor = {k: v for k, v in resultados.items() if v['gestor'] == gestor_logado}
        
        if not resultados_gestor:
            st.warning("Nenhum dado encontrado para seu time.")
            return
        
        # Métricas
        col1, col2, col3, col4, col5 = st.columns(5)
        
        total_analistas = len(resultados_gestor)
        total_atendimentos = sum([d['total_atendimentos'] for d in resultados_gestor.values()])
        media_atendimentos = total_atendimentos / total_analistas if total_analistas > 0 else 0
        
        with col1:
            st.metric("📊 Analistas", total_analistas)
        with col2:
            st.metric("💬 Atendimentos", f"{total_atendimentos:,}")
        with col3:
            st.metric("📈 Média Atend.", f"{media_atendimentos:.0f}")
        with col4:
            metas = len([d for d in resultados_gestor.values() if d['status'] == '🟢 Meta Superada'])
            st.metric("🏆 Metas", f"{metas}/{total_analistas}")
        with col5:
            csat_medio = sum([d['csat'] for d in resultados_gestor.values()]) / total_analistas if total_analistas > 0 else 0
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
            for nome, dados in resultados_gestor.items()
        ])
        
        if not df_dashboard.empty:
            criar_graficos_dashboard(df_dashboard)
        
        st.markdown("---")
        
        # ===== PÓDIO =====
        st.subheader("🏆 Pódio do Mês")
        
        podio = calcular_podio(resultados, gestor_logado, media_atendimentos)
        
        # Carregar pódio manual do Supabase
        if supabase:
            try:
                podio_manual = carregar_podio_manual(supabase, periodo, gestor_logado)
                if podio_manual:
                    podio = podio_manual
            except:
                pass
        
        # Editor manual do pódio
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
                    if supabase:
                        try:
                            sucesso, mensagem = salvar_podio_manual(supabase, periodo, gestor_logado, podio_manual)
                            if sucesso:
                                podio = podio_manual
                                st.success("✅ Pódio manual salvo com sucesso!")
                                st.rerun()
                            else:
                                st.error(f"Erro ao salvar pódio: {mensagem}")
                        except Exception as e:
                            st.error(f"Erro ao salvar pódio: {str(e)}")
                    else:
                        st.error("❌ Supabase não configurado.")
            
            with col2:
                if st.button("🔄 Resetar Pódio"):
                    if supabase:
                        try:
                            supabase.table('podio_manual').delete().eq('mes_ano', periodo).eq('gestor', gestor_logado).execute()
                            st.success("✅ Pódio resetado!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Erro ao resetar pódio: {str(e)}")
                    else:
                        st.error("❌ Supabase não configurado.")
        
        # Exibir pódio
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
        for analista, dados in sorted(resultados_gestor.items(), key=lambda x: x[1]['csat'], reverse=True):
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
            list(resultados_gestor.keys())
        )
        
        if analista_selecionado:
            dados = resultados_gestor[analista_selecionado]
            
            # ===== RELATÓRIO DE ANÁLISE (separado do feedback) =====
            with st.expander("📊 Relatório de Análise", expanded=True):
                st.subheader(f"📊 Análise de Performance - {analista_selecionado}")
                
                # Gráfico mensal do analista
                if supabase:
                    df_historico_analista = carregar_historico(
                        supabase, 
                        analista=analista_selecionado,
                        gestor=gestor_logado
                    )
                    
                    if df_historico_analista is not None and not df_historico_analista.empty:
                        fig_mensal = gerar_grafico_mensal(
                            analista_selecionado,
                            df_historico_analista,
                            dados['meta_csat'],
                            dados['meta_geral']
                        )
                        if fig_mensal:
                            st.plotly_chart(fig_mensal, use_container_width=True)
                    else:
                        st.info("Dados históricos insuficientes para o gráfico mensal. Processe mais meses para ver a evolução.")
                else:
                    st.warning("Supabase não configurado para histórico.")
                
                # Análise técnica
                analise_tecnica = gerar_analise_tecnica(
                    analista_selecionado,
                    dados,
                    media_atendimentos,
                    podio
                )
                st.markdown(analise_tecnica)
            
            # ===== FEEDBACK (separado) =====
            with st.expander("📝 Feedback de Performance", expanded=True):
                st.subheader(f"📝 Feedback - {analista_selecionado}")
                
                posicao_podio = None
                for i, (nome, _, _, _) in enumerate(podio, 1):
                    if nome == analista_selecionado:
                        posicao_podio = i
                        break
                
                # Observações do gestor
                with st.expander("📝 Observações do Gestor", expanded=False):
                    observacoes = st.text_area(
                        "Adicione observações para o feedback:",
                        height=150,
                        placeholder="Ex: Colaborador teve dificuldade com atendimentos complexos no início do mês, mas evoluiu significativamente...",
                        key="observacoes_feedback"
                    )
                
                # Gerar feedback
                feedback = gerar_feedback_mimo(
                    analista_selecionado,
                    dados,
                    media_atendimentos,
                    posicao_podio
                )
                
                # Editor de feedback
                feedback_editado = st.text_area(
                    "Edite o feedback abaixo:",
                    value=feedback,
                    height=400,
                    key="feedback_editado"
                )
                if feedback_editado != feedback:
                    feedback = feedback_editado
            
            # ===== RELATÓRIO COMPLETO =====
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

## Conclusão e Feedback

**Status Geral do Período:** {dados['status']}

{feedback}
"""
                st.markdown(relatorio_markdown)
            
            # ===== BOTÕES DE DOWNLOAD =====
            col1, col2 = st.columns(2)
            
            with col1:
                # Relatório de Análise (sem feedback)
                relatorio_analise = f"""
# {analista_selecionado} - Análise de Performance

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

{analise_tecnica}
"""
                
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
            Sistema de Performance v7.0 | Supabase + Histórico + Relatórios
        </div>
        """,
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
