import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import io

# ============================================
# CONFIGURAÇÕES
# ============================================

st.set_page_config(
    page_title="Sistema de Performance",
    page_icon="📊",
    layout="wide"
)

# Definição dos times por gestor
GESTORES = {
    "Sua Gestão - Chat Notas": [
        "Ana Claudia Corrêa",
        "João Pedro Santana",
        "João Vitor Almeida",
        "Lorena Almeida",
        "Paulo Victor",
        "Rayane Nunes",
        "Thiago Reis",
        "Vanessa Silva"
    ],
    "Gestão Polyana Ventura - Chat Outros": [
        "Christian Matozinho",
        "Diego Machado",
        "Igor Siqueira",
        "Ismael Chagas Bessa",
        "Karolyne Moreira",
        "Luan Pereira",
        "Mario Junior",
        "Maycon Oliveira",
        "Miguel Augusto",
        "Polliana Santana"
    ]
}

# Metas individuais de CSAT
METAS_CSAT = {
    "Ana Claudia Corrêa": 90,
    "João Pedro Santana": 90,
    "João Vitor Almeida": 90,
    "Lorena Almeida": 86,
    "Paulo Victor": 86,
    "Rayane Nunes": 86,
    "Thiago Reis": 90,
    "Vanessa Silva": 86,
    "Christian Matozinho": 86,
    "Diego Machado": 86,
    "Igor Siqueira": 86,
    "Ismael Chagas Bessa": 86,
    "Karolyne Moreira": 86,
    "Luan Pereira": 86,
    "Mario Junior": 86,
    "Maycon Oliveira": 86,
    "Miguel Augusto": 86,
    "Polliana Santana": 86
}

META_GERAL_AVALIACOES = 25

# ============================================
# FUNÇÕES PRINCIPAIS
# ============================================

def identificar_gestor(analista):
    """Identifica qual gestor o analista pertence"""
    for gestor, analistas in GESTORES.items():
        if analista in analistas:
            return gestor
    return "Outros"

def processar_dados(df_satisfacao, df_inativos):
    """Processa os dados e calcula todas as métricas"""
    resultados = {}
    analistas = df_satisfacao['Nome do atribuído'].unique()
    
    for analista in analistas:
        tickets_analista = df_satisfacao[df_satisfacao['Nome do atribuído'] == analista]
        total_atendimentos = len(tickets_analista)
        
        inativos_analista = df_inativos[df_inativos['Nome do atribuído'] == analista]
        total_inativos = len(inativos_analista)
        validos = total_atendimentos - total_inativos
        
        avaliacoes = len(tickets_analista[tickets_analista['Índice de satisfação do ticket'].isin(['Good', 'Bad'])])
        positivos = len(tickets_analista[tickets_analista['Índice de satisfação do ticket'] == 'Good'])
        negativos = len(tickets_analista[tickets_analista['Índice de satisfação do ticket'] == 'Bad'])
        
        perc_avaliacoes = (avaliacoes / validos * 100) if validos > 0 else 0
        csat = (positivos / avaliacoes * 100) if avaliacoes > 0 else 0
        perc_envio = (avaliacoes / total_atendimentos * 100) if total_atendimentos > 0 else 0
        
        meta_csat = METAS_CSAT.get(analista, 86)
        delta_csat = csat - meta_csat
        gestor = identificar_gestor(analista)
        
        if csat >= meta_csat and perc_avaliacoes >= META_GERAL_AVALIACOES:
            status = "🟢 Meta Superada"
        elif csat >= meta_csat or perc_avaliacoes >= META_GERAL_AVALIACOES:
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
            'status': status,
            'gestor': gestor
        }
    
    return resultados

def calcular_estatisticas_por_gestor(resultados, gestor_nome):
    """Calcula estatísticas agregadas por gestor"""
    analistas_gestor = [a for a, d in resultados.items() if d['gestor'] == gestor_nome]
    
    if not analistas_gestor:
        return None
    
    dados_gestor = [resultados[a] for a in analistas_gestor]
    
    total_atendimentos = sum([d['total_atendimentos'] for d in dados_gestor])
    total_inativos = sum([d['total_inativos'] for d in dados_gestor])
    total_validos = sum([d['validos'] for d in dados_gestor])
    total_avaliacoes = sum([d['avaliacoes'] for d in dados_gestor])
    total_positivos = sum([d['positivos'] for d in dados_gestor])
    
    csat_medio = (total_positivos / total_avaliacoes * 100) if total_avaliacoes > 0 else 0
    perc_avaliacoes_medio = (total_avaliacoes / total_validos * 100) if total_validos > 0 else 0
    perc_envio_medio = (total_avaliacoes / total_atendimentos * 100) if total_atendimentos > 0 else 0
    media_atendimentos = total_atendimentos / len(analistas_gestor) if len(analistas_gestor) > 0 else 0
    
    return {
        'total_analistas': len(analistas_gestor),
        'total_atendimentos': total_atendimentos,
        'total_inativos': total_inativos,
        'total_validos': total_validos,
        'total_avaliacoes': total_avaliacoes,
        'total_positivos': total_positivos,
        'csat_medio': round(csat_medio, 2),
        'perc_avaliacoes_medio': round(perc_avaliacoes_medio, 2),
        'perc_envio_medio': round(perc_envio_medio, 2),
        'media_atendimentos': round(media_atendimentos, 2),
        'analistas': analistas_gestor
    }

def calcular_podio(resultados, gestor_nome=None):
    """Calcula o pódio baseado no CSAT"""
    if gestor_nome:
        analistas_filtrados = {k: v for k, v in resultados.items() 
                              if v['gestor'] == gestor_nome and v['avaliacoes'] > 0}
    else:
        analistas_filtrados = {k: v for k, v in resultados.items() if v['avaliacoes'] > 0}
    
    sorted_analistas = sorted(analistas_filtrados.items(), key=lambda x: x[1]['csat'], reverse=True)
    top_3 = sorted_analistas[:3]
    
    return [(nome, dados['csat']) for nome, dados in top_3]

# ============================================
# INTERFACE STREAMLIT
# ============================================

def main():
    st.title("📊 Sistema de Performance - Relatórios Automáticos")
    st.markdown("---")
    
    # Sidebar
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
        
        if st.button("🚀 Processar Dados", use_container_width=True):
            if arquivo_satisfacao and arquivo_inativos:
                with st.spinner("Processando dados..."):
                    df_satisfacao = pd.read_excel(arquivo_satisfacao)
                    df_inativos = pd.read_excel(arquivo_inativos)
                    resultados = processar_dados(df_satisfacao, df_inativos)
                    st.session_state.resultados = resultados
                    st.session_state.processado = True
                    st.success("✅ Dados processados com sucesso!")
            else:
                st.error("❌ Por favor, envie os dois arquivos.")
    
    # Conteúdo principal
    if st.session_state.get('processado', False):
        resultados = st.session_state.resultados
        
        # Métricas gerais
        col1, col2, col3, col4 = st.columns(4)
        total_analistas = len(resultados)
        total_atendimentos = sum([d['total_atendimentos'] for d in resultados.values()])
        media_atendimentos = total_atendimentos / total_analistas if total_analistas > 0 else 0
        metas_superadas = len([d for d in resultados.values() if d['status'] == '🟢 Meta Superada'])
        
        with col1:
            st.metric("Total de Analistas", total_analistas)
        with col2:
            st.metric("Total de Atendimentos", f"{total_atendimentos:,}")
        with col3:
            st.metric("Média de Atendimentos", f"{media_atendimentos:.0f}")
        with col4:
            st.metric("Metas Superadas", f"{metas_superadas}/{total_analistas}")
        
        st.markdown("---")
        
        # Seletor de gestor
        gestores_disponiveis = list(GESTORES.keys())
        gestor_selecionado = st.selectbox(
            "Selecione o Gestor",
            gestores_disponiveis
        )
        
        if gestor_selecionado:
            stats = calcular_estatisticas_por_gestor(resultados, gestor_selecionado)
            
            if stats:
                # Métricas do gestor
                col1, col2, col3, col4, col5 = st.columns(5)
                with col1:
                    st.metric("Analistas", stats['total_analistas'])
                with col2:
                    st.metric("CSAT Médio", f"{stats['csat_medio']:.2f}%")
                with col3:
                    st.metric("% Avaliações", f"{stats['perc_avaliacoes_medio']:.2f}%")
                with col4:
                    st.metric("% Envio", f"{stats['perc_envio_medio']:.2f}%")
                with col5:
                    st.metric("Média Atend.", f"{stats['media_atendimentos']:.0f}")
                
                st.markdown("---")
                
                # Tabela de desempenho
                st.subheader("📋 Desempenho Individual")
                
                dados_tabela = []
                for analista, dados in resultados.items():
                    if dados['gestor'] == gestor_selecionado:
                        dados_tabela.append({
                            'Analista': analista,
                            'CSAT': f"{dados['csat']:.2f}%",
                            'Meta': f"{dados['meta_csat']}%",
                            'Delta': f"{dados['delta_csat']:+.2f}%",
                            '% Avaliações': f"{dados['perc_avaliacoes']:.2f}%",
                            '% Envio': f"{dados['perc_envio']:.2f}%",
                            'Atendimentos': dados['total_atendimentos'],
                            'Status': dados['status']
                        })
                
                df_tabela = pd.DataFrame(dados_tabela)
                st.dataframe(df_tabela, use_container_width=True, hide_index=True)
                
                st.markdown("---")
                
                # Pódio
                st.subheader("🏆 Pódio do Mês")
                podio = calcular_podio(resultados, gestor_selecionado)
                
                if podio:
                    col1, col2, col3 = st.columns(3)
                    for i, (col, (nome, csat)) in enumerate(zip([col1, col2, col3], podio), 1):
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
                            </div>
                            """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
