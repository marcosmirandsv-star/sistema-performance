import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import io
import json
import os
import google.generativeai as genai

# ============================================
# CONFIGURAÇÕES INICIAIS
# ============================================

st.set_page_config(
    page_title="Sistema de Performance",
    page_icon="📊",
    layout="wide"
)

# ============================================
# CONFIGURAÇÃO DA GEMINI API
# ============================================

def configurar_gemini():
    """Configura a API do Gemini"""
    try:
        api_key = st.secrets["GEMINI_API_KEY"]
        genai.configure(api_key=api_key)
        return True
    except:
        api_key = os.environ.get("GEMINI_API_KEY")
        if api_key:
            genai.configure(api_key=api_key)
            return True
        return False

def gerar_analise_gemini(analista, dados, media_operacao, podio):
    """Gera análise técnica usando Gemini API"""
    
    if not configurar_gemini():
        return gerar_analise_tecnica_fallback(analista, dados, media_operacao, podio)
    
    try:
        podio_texto = ""
        for i, (nome, csat, atendimentos, perc_envio) in enumerate(podio, 1):
            podio_texto += f"{i}º: {nome} - CSAT: {csat:.2f}% | {atendimentos} atendimentos | % Envio: {perc_envio:.2f}%\n"
        
        prompt = f"""
        Você é um especialista em análise de desempenho de equipes de atendimento ao cliente.
        Analise os dados do colaborador {analista} e gere uma análise técnica detalhada.

        ## DADOS DO COLABORADOR:
        - CSAT (Satisfação do Cliente): {dados['csat']:.2f}%
        - Meta de CSAT: {dados['meta_csat']:.0f}%
        - Delta CSAT: {dados['delta_csat']:+.2f} pontos percentuais
        - % de Avaliações: {dados['perc_avaliacoes']:.2f}%
        - Meta de Avaliações: {dados['meta_geral']:.0f}%
        - Total de Atendimentos: {dados['total_atendimentos']}
        - Atendimentos Válidos: {dados['validos']}
        - Inativos: {dados['total_inativos']}
        - Total de Avaliações: {dados['avaliacoes']}
        - Avaliações Positivas: {dados['positivos']}
        - Avaliações Negativas: {dados['negativos']}
        - % de Envio de Avaliações: {dados['perc_envio']:.2f}%

        ## DADOS DA OPERAÇÃO:
        - Média de Atendimentos por Agente: {media_operacao}

        ## PÓDIO DO MÊS (por CSAT, com avaliações ≥ 25%):
        {podio_texto}

        ## INSTRUÇÕES:
        Gere uma análise técnica de desempenho seguindo EXATAMENTE esta estrutura:

        ### 1. Qualidade e Satisfação do Cliente (CSAT)
        - Apresente o índice de CSAT
        - Compare com a meta e calcule o delta
        - Analise o volume de avaliações positivas e negativas
        - Dê uma interpretação qualitativa do resultado

        ### 2. Engajamento e Coleta de Feedback
        - Apresente a taxa de avaliações
        - Compare com a meta
        - Analise a proatividade na coleta de feedback
        - Mostre o cálculo da taxa

        ### 3. Produtividade e Volumetria
        - Apresente o volume total de atendimentos
        - Compare com a média da operação
        - Destaque a posição no pódio (se aplicável)
        - Analise a produtividade geral

        ## TOM DE VOZ:
        - Profissional e objetivo
        - Use dados numéricos para embasar a análise
        - Seja construtivo e motivador
        - Destaque conquistas e oportunidades de melhoria
        - Use linguagem clara e acessível

        ## FORMATO:
        - Use markdown para formatação
        - Destaque números importantes em **negrito**
        - Mantenha a estrutura com ### para cada seção
        - Seja conciso, mas completo (máximo 400 palavras)

        Gere a análise agora:
        """
        
        model = genai.GenerativeModel('gemini-1.5-flash')
        response = model.generate_content(prompt)
        return response.text.strip()
        
    except Exception as e:
        st.warning(f"Erro ao gerar análise com Gemini: {str(e)}. Usando análise padrão.")
        return gerar_analise_tecnica_fallback(analista, dados, media_operacao, podio)

def gerar_analise_tecnica_fallback(analista, dados, media_operacao, podio):
    """Análise técnica padrão (fallback sem IA)"""
    
    posicao = "não está no pódio"
    for i, (nome, csat, atendimentos, perc_envio) in enumerate(podio, 1):
        if nome == analista:
            posicao = f"está em {i}º lugar no pódio com CSAT de {csat:.2f}%, {atendimentos} atendimentos e {perc_envio:.2f}% de não envio"
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

A colaboradora registrou um índice de **Satisfação (CSAT) de {dados['csat']:.2f}%**.

- **Comparativo com a Meta:** O resultado {analise_csat} (meta: ≥ {dados['meta_csat']:.0f}%).
- **Análise Detalhada:** Do volume total de feedbacks recebidos ({dados['avaliacoes']}), **{dados['positivos']} foram positivos**, resultando em um índice de aprovação {nivel_csat}. Houve apenas {dados['negativos']} registros negativos, demonstrando consistência na entrega de um atendimento cordial, eficiente e resolutivo.

### 2. Engajamento e Coleta de Feedback

A colaboradora alcançou uma **Taxa de Avaliações de {dados['perc_avaliacoes']:.2f}%**.

- **Comparativo com a Meta:** A meta esperada é de no mínimo {dados['meta_geral']:.0f}% de conversão de atendimentos em avaliações. O resultado {analise_avaliacoes}, refletindo uma {'proatividade exemplar' if dados['perc_avaliacoes'] >= dados['meta_geral'] else 'oportunidade de melhoria'} na solicitação de feedback ao final de cada contato.

- **Cálculo:** A taxa foi calculada sobre o volume total de {dados['avaliacoes']} avaliações divididas pelos **{dados['validos']} atendimentos válidos**.

### 3. Produtividade e Volumetria

O volume total de atendimentos realizados pela colaboradora foi de **{dados['total_atendimentos']} chamados**.

- **Comparativo com a Operação:** A média de atendimentos por agente da operação no período foi de {media_operacao}. {analista} absorveu uma demanda operacional **{produtividade}**.

- **Destaque:** {posicao} do mês em CSAT.
"""
    return texto.strip()

# ============================================
# GERENCIAMENTO DE ANALISTAS
# ============================================

def carregar_analistas():
    """Carrega a lista de analistas do session_state ou arquivo"""
    if 'analistas' not in st.session_state:
        st.session_state.analistas = {
            "Sua Gestão - Chat Notas": {
                "gestor": "Marcos Miranda",
                "meta_geral_avaliacoes": 25,
                "membros": {
                    "Ana Claudia Corrêa": {"meta_csat": 90, "ativo": True},
                    "João Pedro Santana": {"meta_csat": 90, "ativo": True},
                    "João Vitor Almeida": {"meta_csat": 90, "ativo": True},
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
                    "Karolyne Moreira": {"meta_csat": 86, "ativo": True},
                    "Luan Pereira": {"meta_csat": 86, "ativo": True},
                    "Mario Junior": {"meta_csat": 86, "ativo": True},
                    "Maycon Oliveira": {"meta_csat": 86, "ativo": True},
                    "Miguel Augusto": {"meta_csat": 86, "ativo": True},
                    "Polliana Santana": {"meta_csat": 86, "ativo": True}
                }
            }
        }
    return st.session_state.analistas

def salvar_analistas(analistas):
    """Salva a lista de analistas no session_state"""
    st.session_state.analistas = analistas

# ============================================
# FUNÇÕES DE CÁLCULO
# ============================================

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
        
        # ✅ CÁLCULO CORRETO DO % DE ENVIO
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

def calcular_podio(resultados, gestor=None):
    """Calcula o pódio baseado no CSAT (apenas quem tem >= 25% avaliações)"""
    if gestor:
        dados_filtrados = {k: v for k, v in resultados.items() 
                          if v['gestor'] == gestor and v['avaliacoes'] > 0}
    else:
        dados_filtrados = {k: v for k, v in resultados.items() if v['avaliacoes'] > 0}
    
    dados_validos = {k: v for k, v in dados_filtrados.items() 
                     if v['perc_avaliacoes'] >= 25}
    
    sorted_analistas = sorted(dados_validos.items(), key=lambda x: x[1]['csat'], reverse=True)
    top_3 = sorted_analistas[:3]
    
    return [(nome, dados['csat'], dados['total_atendimentos'], dados['perc_envio']) for nome, dados in top_3]

# ============================================
# GERAÇÃO DE RELATÓRIOS
# ============================================

def gerar_feedback(analista, dados):
    """Gera feedback no formato GUP (Mimo/Sari)"""
    
    pontos_fortes = []
    areas_melhoria = []
    
    if dados['csat'] >= dados['meta_csat']:
        pontos_fortes.append("excelente performance em CSAT")
    else:
        areas_melhoria.append("melhorar o CSAT")
    
    if dados['perc_avaliacoes'] >= dados['meta_geral']:
        pontos_fortes.append("alta taxa de engajamento com coleta de feedback")
    else:
        areas_melhoria.append("aumentar a taxa de coleta de avaliações")
    
    if dados['total_atendimentos'] > 0:
        pontos_fortes.append("boa produtividade")
    
    if dados['status'] == "🟢 Meta Superada":
        avaliacao = "excelente"
        recomendacao = "Continue mantendo este alto padrão de qualidade e produtividade. Seu desempenho é referência para a equipe."
    elif dados['status'] == "🟡 Atenção":
        avaliacao = "bom, com pontos de atenção"
        recomendacao = f"Foque em {' e '.join(areas_melhoria)} para alcançar todas as metas. Você tem potencial para melhorar ainda mais."
    else:
        avaliacao = "precisa de atenção"
        recomendacao = f"É necessário um plano de ação focado em {' e '.join(areas_melhoria)}. Conte com o suporte da gestão para isso."
    
    feedback = f"""
**Feedback {dados['status']}**

**Colaborador:** {analista}

**Avaliação Geral:** O desempenho do colaborador no período foi {avaliacao}.

**Pontos Fortes:** {'; '.join(pontos_fortes) if pontos_fortes else 'Nenhum ponto forte identificado'}.

**Áreas de Melhoria:** {'; '.join(areas_melhoria) if areas_melhoria else 'Nenhuma área de melhoria identificada'}.

**Recomendação:** {recomendacao}

**Próximos Passos:** 
- Revisar os resultados em 30 dias
- {f'Aplicar plano de ação focado em {areas_melhoria[0] if areas_melhoria else "manutenção do desempenho"}'}
"""
    return feedback.strip()

def gerar_relatorio_completo(analista, dados, analise_tecnica, feedback, media_operacao, podio):
    """Gera o relatório completo no formato do exemplo da Ana Cláudia"""
    
    posicao_texto = "Não está no pódio"
    for i, (nome, csat, atendimentos, perc_envio) in enumerate(podio, 1):
        if nome == analista:
            posicao_texto = f"{i}º Lugar - 🏆 (CSAT: {csat:.2f}% | {atendimentos} atendimentos | % Envio: {perc_envio:.2f}%)"
            break
    
    relatorio = f"""
# {analista}

**Período:** {datetime.now().strftime('%d/%m/%Y a %d/%m/%Y')}

## Esperado:

- ≥ {dados['meta_geral']:.0f}% de avaliações
- ≥ {dados['meta_csat']:.0f}% de Satisfação

## Atingido:

- **CSAT:** {dados['csat']:.2f}%
- **Avaliações:** {dados['perc_avaliacoes']:.2f}% ({dados['positivos']} positivos + {dados['negativos']} negativos = {dados['avaliacoes']})
- **% Envio:** {dados['perc_envio']:.2f}% ({dados['validos'] - dados['avaliacoes']} clientes não avaliaram)
- **Atendidos:** {dados['total_atendimentos']} - {dados['total_inativos']} = {dados['validos']}
- **Média por agente:** {media_operacao}
- **Posição no Pódio:** {posicao_texto}

---

## Análise Técnica de Desempenho

{analise_tecnica}

---

## Conclusão e Feedback

**Status Geral do Período:** {dados['status']}

{feedback}

---

## 📊 Gráfico de Evolução Histórica

*[Área reservada para gráfico de evolução histórica - AVD]*

"""
    return relatorio

# ============================================
# INTERFACE STREAMLIT
# ============================================

def main():
    st.title("📊 Sistema de Performance - Relatórios Automáticos")
    st.markdown("---")
    
    analistas_config = carregar_analistas()
    
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
        st.header("🤖 Configuração Gemini")
        
        api_key_input = st.text_input(
            "Chave da API Gemini",
            type="password",
            placeholder="Cole sua chave API aqui"
        )
        
        if api_key_input:
            os.environ["GEMINI_API_KEY"] = api_key_input
            st.session_state.gemini_configured = True
            st.success("✅ API Gemini configurada!")
        
        st.markdown("---")
        
        if st.button("🚀 Processar Dados", use_container_width=True):
            if arquivo_satisfacao and arquivo_inativos:
                with st.spinner("Processando dados..."):
                    df_satisfacao = pd.read_excel(arquivo_satisfacao)
                    df_inativos = pd.read_excel(arquivo_inativos)
                    resultados = processar_dados(df_satisfacao, df_inativos, analistas_config)
                    st.session_state.resultados = resultados
                    st.session_state.processado = True
                    st.success("✅ Dados processados com sucesso!")
            else:
                st.error("❌ Por favor, envie os dois arquivos.")
        
        st.markdown("---")
        st.header("⚙️ Gerenciar Analistas")
        
        if st.button("📝 Gerenciar Analistas", use_container_width=True):
            st.session_state.gerenciar_analistas = True
    
    if st.session_state.get('gerenciar_analistas', False):
        st.header("📝 Gerenciar Analistas")
        
        gestor_selecionado = st.selectbox(
            "Selecione o Gestor",
            list(analistas_config.keys())
        )
        
        if gestor_selecionado:
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
            st.subheader("👥 Membros do Time")
            
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
        
        if st.button("🔙 Voltar", use_container_width=True):
            st.session_state.gerenciar_analistas = False
            st.rerun()
        
        st.markdown("---")
    
    if st.session_state.get('processado', False) and not st.session_state.get('gerenciar_analistas', False):
        resultados = st.session_state.resultados
        
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
            st.metric("Média Geral de Atendimentos", f"{media_atendimentos:.0f}")
        with col4:
            st.metric("Metas Superadas", f"{metas_superadas}/{total_analistas}")
        
        st.markdown("---")
        
        gestores_disponiveis = list(analistas_config.keys())
        gestor_selecionado = st.selectbox(
            "Selecione o Gestor",
            gestores_disponiveis
        )
        
        if gestor_selecionado:
            resultados_gestor = {k: v for k, v in resultados.items() if v['gestor'] == gestor_selecionado}
            
            if resultados_gestor:
                media_gestor = calcular_media_operacao(resultados, gestor_selecionado)
                podio_gestor = calcular_podio(resultados, gestor_selecionado)
                
                col1, col2, col3, col4, col5 = st.columns(5)
                with col1:
                    st.metric("Analistas", len(resultados_gestor))
                with col2:
                    csat_medio = sum([d['csat'] for d in resultados_gestor.values()]) / len(resultados_gestor)
                    st.metric("CSAT Médio", f"{csat_medio:.2f}%")
                with col3:
                    perc_medio = sum([d['perc_avaliacoes'] for d in resultados_gestor.values()]) / len(resultados_gestor)
                    st.metric("% Avaliações Médio", f"{perc_medio:.2f}%")
                with col4:
                    st.metric("Média Atendimentos", f"{media_gestor:.0f}")
                with col5:
                    metas_time = len([d for d in resultados_gestor.values() if d['status'] == '🟢 Meta Superada'])
                    st.metric("Metas Superadas", f"{metas_time}/{len(resultados_gestor)}")
                
                st.markdown("---")
                
                st.subheader("📋 Desempenho Individual")
                
                dados_tabela = []
                for analista, dados in sorted(resultados_gestor.items(), key=lambda x: x[1]['csat'], reverse=True):
                    dados_tabela.append({
                        'Analista': analista,
                        'CSAT': f"{dados['csat']:.2f}%",
                        'Meta CSAT': f"{dados['meta_csat']:.0f}%",
                        'Delta': f"{dados['delta_csat']:+.2f}%",
                        '% Avaliações': f"{dados['perc_avaliacoes']:.2f}%",
                        'Meta Avaliações': f"{dados['meta_geral']:.0f}%",
                        '% Envio': f"{dados['perc_envio']:.2f}%",
                        'Atendimentos': dados['total_atendimentos'],
                        'Status': dados['status']
                    })
                
                df_tabela = pd.DataFrame(dados_tabela)
                st.dataframe(df_tabela, use_container_width=True, hide_index=True)
                
                st.markdown("---")
                
                st.subheader("🏆 Pódio do Mês")
                podio = calcular_podio(resultados, gestor_selecionado)
                
                if podio:
                    col1, col2, col3 = st.columns(3)
                    for i, (col, (nome, csat, atendimentos, perc_envio)) in enumerate(zip([col1, col2, col3], podio), 1):
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
                                <p style="margin: 5px 0; font-size: 16px; color: #444;">📞 {atendimentos} atendimentos</p>
                                <p style="margin: 5px 0; font-size: 16px; color: #444;">📤 {perc_envio:.2f}% não avaliaram</p>
                                <p style="margin: 5px 0; font-size: 14px; color: #666;">✅ Avaliações ≥ 25%</p>
                            </div>
                            """, unsafe_allow_html=True)
                else:
                    st.info("Nenhum analista atingiu os critérios para o pódio (CSAT e avaliações ≥ 25%)")
                
                st.markdown("---")
                
                st.subheader("📄 Gerar Relatório Individual")
                
                analista_selecionado = st.selectbox(
                    "Selecione o Analista",
                    list(resultados_gestor.keys())
                )
                
                if analista_selecionado:
                    dados = resultados_gestor[analista_selecionado]
                    
                    media_operacao_geral = calcular_media_operacao(resultados)
                    podio_geral = calcular_podio(resultados)
                    
                    usar_gemini = st.checkbox(
                        "🤖 Usar Gemini para análise técnica (requer API configurada)",
                        value=st.session_state.get('gemini_configured', False)
                    )
                    
                    with st.spinner("Gerando análise técnica..."):
                        if usar_gemini and st.session_state.get('gemini_configured', False):
                            analise_tecnica = gerar_analise_gemini(
                                analista_selecionado,
                                dados,
                                media_operacao_geral,
                                podio_geral
                            )
                        else:
                            analise_tecnica = gerar_analise_tecnica_fallback(
                                analista_selecionado,
                                dados,
                                media_operacao_geral,
                                podio_geral
                            )
                    
                    feedback = gerar_feedback(analista_selecionado, dados)
                    
                    with st.expander("✏️ Editar Feedback", expanded=False):
                        feedback_editado = st.text_area(
                            "Edite o feedback abaixo:",
                            value=feedback,
                            height=300
                        )
                        if feedback_editado != feedback:
                            feedback = feedback_editado
                    
                    relatorio = gerar_relatorio_completo(
                        analista_selecionado,
                        dados,
                        analise_tecnica,
                        feedback,
                        media_operacao_geral,
                        podio_geral
                    )
                    
                    with st.expander("👁️ Visualizar Relatório", expanded=True):
                        st.markdown(relatorio)
                    
                    st.download_button(
                        label="📥 Baixar Relatório (Markdown)",
                        data=relatorio,
                        file_name=f"Relatorio_{analista_selecionado.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.md",
                        mime="text/markdown"
                    )
                    
                    st.download_button(
                        label="📥 Baixar Relatório (Word)",
                        data=relatorio,
                        file_name=f"Relatorio_{analista_selecionado.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            else:
                st.warning("Nenhum dado encontrado para este gestor.")
    
    st.markdown("---")
    st.markdown(
        """
        <div style="text-align: center; color: #666; font-size: 12px;">
            Sistema de Performance v2.0 | Desenvolvido com ❤️
        </div>
        """,
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
