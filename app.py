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
        return False
    
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
        
        for registro in registros:
            supabase.table('historico_performance').insert(registro).execute()
        
        return True
    except Exception as e:
        st.error(f"Erro ao salvar histórico: {str(e)}")
        return False

def carregar_historico(supabase, mes_ano=None, gestor=None):
    """Carrega histórico do Supabase"""
    if not supabase:
        return None
    
    try:
        query = supabase.table('historico_performance').select('*')
        
        if mes_ano:
            query = query.eq('mes_ano', mes_ano)
        if gestor:
            query = query.eq('gestor', gestor)
        
        response = query.execute()
        return pd.DataFrame(response.data)
    except Exception as e:
        st.error(f"Erro ao carregar histórico: {str(e)}")
        return None

def salvar_podio_manual(supabase, mes_ano, gestor, podio):
    """Salva edição manual do pódio"""
    if not supabase:
        return False
    
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
        
        return True
    except Exception as e:
        st.error(f"Erro ao salvar pódio: {str(e)}")
        return False

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
        st.error(f"Erro ao carregar pódio: {str(e)}")
        return None

# ============================================
# MULTI-LOGIN
# ============================================

USUARIOS = {
    "marcos": {
        "senha": "marcos2026",
        "nome": "Marcos Miranda",
        "gestor": "Sua Gestão - Chat Notas"
    },
    "polyana": {
        "senha": "polyana2026",
        "nome": "Polyana Ventura",
        "gestor": "Gestão Polyana Ventura - Chat Outros"
    }
}

def fazer_login():
    """Tela de login"""
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔐 Login")
    
    usuario = st.sidebar.text_input("Usuário")
    senha = st.sidebar.text_input("Senha", type="password")
    
    if st.sidebar.button("Entrar", use_container_width=True):
        if usuario in USUARIOS and USUARIOS[usuario]["senha"] == senha:
            st.session_state.logado = True
            st.session_state.usuario = usuario
            st.session_state.nome_usuario = USUARIOS[usuario]["nome"]
            st.session_state.gestor = USUARIOS[usuario]["gestor"]
            st.rerun()
        else:
            st.sidebar.error("❌ Usuário ou senha inválidos!")
    
    if st.session_state.get('logado', False):
        st.sidebar.success(f"✅ Logado como {st.session_state.nome_usuario}")
        if st.sidebar.button("Sair", use_container_width=True):
            st.session_state.logado = False
            st.session_state.usuario = None
            st.session_state.nome_usuario = None
            st.session_state.gestor = None
            st.rerun()
        return True
    return False

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

def carregar_analistas():
    """Carrega a lista de analistas"""
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
# GERAÇÃO DE ANÁLISE E FEEDBACK
# ============================================

def gerar_analise_tecnica(analista, dados, media_operacao, podio):
    """Gera análise técnica detalhada"""
    genero = get_genero_neutro(analista)
    
    posicao = "não está no pódio"
    for i, (nome, csat, atendimentos, perc_envio) in enumerate(podio, 1):
        if nome == analista:
            posicao = f"está em {i}º lugar no pódio com CSAT de {csat:.2f}%, {atendimentos} atendimentos"
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

def gerar_relatorio_word(analista, dados, analise_tecnica, feedback, media_operacao, podio, periodo):
    """Gera relatório em formato Word (.docx)"""
    
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
            posicao_texto = f"{i}º Lugar - CSAT: {csat:.2f}% | {atendimentos} atendimentos"
            break
    
    doc.add_paragraph(f'CSAT: {dados["csat"]:.2f}%', style='List Bullet')
    doc.add_paragraph(f'Avaliações: {dados["perc_avaliacoes"]:.2f}% ({dados["positivos"]} positivos + {dados["negativos"]} negativos = {dados["avaliacoes"]})', style='List Bullet')
    doc.add_paragraph(f'Atendidos: {dados["total_atendimentos"]} - {dados["total_inativos"]} = {dados["validos"]}', style='List Bullet')
    doc.add_paragraph(f'Média por agente: {media_operacao}', style='List Bullet')
    doc.add_paragraph(f'Posição no Pódio: {posicao_texto}', style='List Bullet')
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
# DASHBOARD COM HISTÓRICO
# ============================================

def criar_dashboard_historico(df_historico, gestor_selecionado):
    """Cria dashboard com gráficos históricos"""
    
    if df_historico is None or df_historico.empty:
        st.info("Nenhum dado histórico disponível ainda.")
        return
    
    df_filtrado = df_historico[df_historico['gestor'] == gestor_selecionado]
    
    if df_filtrado.empty:
        st.info(f"Nenhum dado histórico para {gestor_selecionado}")
        return
    
    df_mensal = df_filtrado.groupby('mes_ano').agg({
        'csat': 'mean',
        'perc_avaliacoes': 'mean',
        'total_atendimentos': 'sum'
    }).reset_index()
    
    fig_csat = px.line(
        df_mensal,
        x='mes_ano',
        y='csat',
        title='📈 Evolução do CSAT Médio',
        labels={'mes_ano': 'Mês/Ano', 'csat': 'CSAT (%)'}
    )
    fig_csat.add_hline(y=90, line_dash="dash", line_color="green", annotation_text="Meta CSAT (90%)")
    fig_csat.update_layout(height=400)
    st.plotly_chart(fig_csat, use_container_width=True)
    
    fig_avaliacoes = px.line(
        df_mensal,
        x='mes_ano',
        y='perc_avaliacoes',
        title='📊 Evolução da Taxa de Avaliações',
        labels={'mes_ano': 'Mês/Ano', 'perc_avaliacoes': '% Avaliações'}
    )
    fig_avaliacoes.add_hline(y=25, line_dash="dash", line_color="orange", annotation_text="Meta (25%)")
    fig_avaliacoes.update_layout(height=400)
    st.plotly_chart(fig_avaliacoes, use_container_width=True)
    
    st.subheader("📋 Evolução Individual por Mês")
    
    df_analista_mes = df_filtrado.pivot_table(
        index='analista',
        columns='mes_ano',
        values='csat',
        aggfunc='mean'
    ).round(2)
    
    st.dataframe(df_analista_mes, use_container_width=True)

# ============================================
# INTERFACE PRINCIPAL
# ============================================

def main():
    st.title("📊 Sistema de Performance - Relatórios Automáticos")
    st.markdown("---")
    
    if not fazer_login():
        st.info("👋 Faça login na barra lateral para acessar o sistema.")
        return
    
    supabase = init_supabase()
    if supabase:
        st.sidebar.success("✅ Conectado ao Supabase")
    else:
        st.sidebar.warning("⚠️ Supabase não configurado. Configure as variáveis de ambiente.")
    
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
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("🚀 Processar", use_container_width=True):
                if arquivo_satisfacao and arquivo_inativos:
                    with st.spinner("Processando dados..."):
                        df_satisfacao = pd.read_excel(arquivo_satisfacao)
                        df_inativos = pd.read_excel(arquivo_inativos)
                        
                        periodo = extrair_periodo(df_satisfacao)
                        st.session_state.periodo = periodo
                        
                        resultados = processar_dados(df_satisfacao, df_inativos, analistas_config)
                        st.session_state.resultados = resultados
                        st.session_state.processado = True
                        
                        if supabase:
                            gestor_logado = st.session_state.gestor
                            if salvar_historico(supabase, resultados, periodo, gestor_logado):
                                st.success(f"✅ Dados salvos! Período: {periodo}")
                            else:
                                st.warning("⚠️ Dados processados, mas NÃO salvos no Supabase.")
                        else:
                            st.success(f"✅ Dados processados! Período: {periodo}")
                else:
                    st.error("❌ Envie os dois arquivos.")
        
        with col2:
            if st.button("📊 Ver Histórico", use_container_width=True):
                if supabase:
                    gestor_logado = st.session_state.gestor
                    df_historico = carregar_historico(supabase, gestor=gestor_logado)
                    if df_historico is not None and not df_historico.empty:
                        st.session_state.mostrar_historico = True
                        st.session_state.df_historico = df_historico
                        st.success("✅ Histórico carregado!")
                    else:
                        st.warning("Nenhum histórico encontrado.")
                else:
                    st.error("❌ Supabase não configurado.")
        
        st.markdown("---")
        st.header("⚙️ Gerenciar Analistas")
        
        if st.button("📝 Gerenciar Analistas", use_container_width=True):
            st.session_state.gerenciar_analistas = True
    
    # ===== GERENCIAR ANALISTAS =====
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
    
    # ===== VISUALIZAR HISTÓRICO =====
    if st.session_state.get('mostrar_historico', False):
        st.header("📊 Dashboard Histórico")
        
        df_historico = st.session_state.df_historico
        
        col1, col2 = st.columns(2)
        with col1:
            meses_disponiveis = ['Todos'] + sorted(df_historico['mes_ano'].unique().tolist())
            mes_selecionado = st.selectbox("Selecione o Mês", meses_disponiveis)
        with col2:
            gestores_disponiveis = ['Todos'] + sorted(df_historico['gestor'].unique().tolist())
            gestor_filtro = st.selectbox("Selecione o Gestor", gestores_disponiveis)
        
        df_filtrado = df_historico.copy()
        if mes_selecionado != 'Todos':
            df_filtrado = df_filtrado[df_filtrado['mes_ano'] == mes_selecionado]
        if gestor_filtro != 'Todos':
            df_filtrado = df_filtrado[df_filtrado['gestor'] == gestor_filtro]
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Registros", len(df_filtrado))
        with col2:
            st.metric("CSAT Médio", f"{df_filtrado['csat'].mean():.2f}%")
        with col3:
            st.metric("% Avaliações Médio", f"{df_filtrado['perc_avaliacoes'].mean():.2f}%")
        with col4:
            st.metric("Total Atendimentos", f"{df_filtrado['total_atendimentos'].sum():,}")
        
        st.markdown("---")
        
        criar_dashboard_historico(df_filtrado, st.session_state.gestor)
        
        st.subheader("📋 Dados Históricos")
        st.dataframe(df_filtrado, use_container_width=True)
        
        if st.button("🔙 Voltar para Dados Atuais"):
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
        
        col1, col2, col3, col4, col5 = st.columns(5)
        
        total_analistas = len(resultados_gestor)
        total_atendimentos = sum([d['total_atendimentos'] for d in resultados_gestor.values()])
        media_atendimentos = total_atendimentos / total_analistas if total_analistas > 0 else 0
        
        with col1:
            st.metric("📊 Analistas", total_analistas)
        with col2:
            st.metric("📞 Atendimentos", f"{total_atendimentos:,}")
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
        
        # Dashboard com gráficos
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
            col1, col2 = st.columns(2)
            
            with col1:
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
                fig_csat.update_layout(height=400)
                st.plotly_chart(fig_csat, use_container_width=True)
            
            with col2:
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
                fig_avaliacoes.update_layout(height=400)
                st.plotly_chart(fig_avaliacoes, use_container_width=True)
            
            col3, col4 = st.columns(2)
            
            with col3:
                fig_scatter = px.scatter(
                    df_dashboard,
                    x='💬 Atendimentos',
                    y='CSAT',
                    size='💬 Atendimentos',
                    color='Status',
                    hover_data=['Analista'],
                    title='🎯 CSAT vs 💬 Atendimentos',
                    labels={'💬 Atendimentos': '💬 Quantidade de Atendimentos', 'CSAT': 'CSAT (%)'}
                )
                fig_scatter.update_layout(height=400)
                st.plotly_chart(fig_scatter, use_container_width=True)
            
            with col4:
                status_counts = df_dashboard['Status'].value_counts()
                fig_pie = px.pie(
                    values=status_counts.values,
                    names=status_counts.index,
                    title='📊 Distribuição de Status'
                )
                fig_pie.update_layout(height=400)
                st.plotly_chart(fig_pie, use_container_width=True)
        
        st.markdown("---")
        
        # Pódio
        st.subheader("🏆 Pódio do Mês")
        
        podio = calcular_podio(resultados, gestor_logado, media_atendimentos)
        
        if supabase:
            podio_manual = carregar_podio_manual(supabase, periodo, gestor_logado)
            if podio_manual:
                podio = podio_manual
        
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
                        if salvar_podio_manual(supabase, periodo, gestor_logado, podio_manual):
                            podio = podio_manual
                            st.success("✅ Pódio manual salvo com sucesso!")
                            st.rerun()
                    else:
                        st.error("❌ Supabase não configurado.")
            
            with col2:
                if st.button("🔄 Resetar Pódio"):
                    if supabase:
                        supabase.table('podio_manual').delete().eq('mes_ano', periodo).eq('gestor', gestor_logado).execute()
                        st.success("✅ Pódio resetado!")
                        st.rerun()
        
        if podio:
            col1, col2, col3 = st.columns(3)
            for i, (col, (nome, csat, atendimentos, _)) in enumerate(zip([col1, col2, col3], podio), 1):
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
                        <p style="margin: 5px 0; font-size: 14px; color: #28a745;">✅ CSAT ≥ 90% | Atend. ≥ Média</p>
                    </div>
                    """, unsafe_allow_html=True)
        else:
            st.info("Nenhum analista atingiu os critérios do pódio (CSAT ≥ 90% e 💬 Atendimentos ≥ Média)")
        
        st.markdown("---")
        
        # Tabela de desempenho
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
                '💬 Atendimentos': dados['total_atendimentos'],
                'Média Operação': f"{media_atendimentos:.0f}",
                'Status': dados['status']
            })
        
        df_tabela = pd.DataFrame(dados_tabela)
        st.dataframe(df_tabela, use_container_width=True, hide_index=True)
        
        st.markdown("---")
        
        # Gerar relatório individual
        st.subheader("📄 Gerar Relatório Individual")
        
        analista_selecionado = st.selectbox(
            "Selecione o Analista",
            list(resultados_gestor.keys())
        )
        
        if analista_selecionado:
            dados = resultados_gestor[analista_selecionado]
            
            analise_tecnica = gerar_analise_tecnica(
                analista_selecionado,
                dados,
                media_atendimentos,
                podio
            )
            
            posicao_podio = None
            for i, (nome, _, _, _) in enumerate(podio, 1):
                if nome == analista_selecionado:
                    posicao_podio = i
                    break
            
            feedback = gerar_feedback_mimo(
                analista_selecionado,
                dados,
                media_atendimentos,
                posicao_podio
            )
            
            with st.expander("✏️ Editar Feedback (Padrão MIMO)", expanded=False):
                feedback_editado = st.text_area(
                    "Edite o feedback abaixo:",
                    value=feedback,
                    height=400
                )
                if feedback_editado != feedback:
                    feedback = feedback_editado
            
            relatorio_markdown = f"""
# {analista_selecionado}

**Período:** {periodo}

## Esperado:

- ≥ {dados['meta_geral']:.0f}% de avaliações
- ≥ {dados['meta_csat']:.0f}% de Satisfação

## Atingido:

- **CSAT:** {dados['csat']:.2f}%
- **Avaliações:** {dados['perc_avaliacoes']:.2f}% ({dados['positivos']} positivos + {dados['negativos']} negativos = {dados['avaliacoes']})
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
            
            with st.expander("👁️ Visualizar Relatório", expanded=True):
                st.markdown(relatorio_markdown)
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.download_button(
                    label="📥 Baixar Relatório (Word)",
                    data=gerar_relatorio_word(
                        analista_selecionado,
                        dados,
                        analise_tecnica,
                        feedback,
                        media_atendimentos,
                        podio,
                        periodo
                    ),
                    file_name=f"Relatorio_{analista_selecionado.replace(' ', '_')}_{periodo.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            
            with col2:
                st.download_button(
                    label="📥 Baixar Relatório (Markdown)",
                    data=relatorio_markdown,
                    file_name=f"Relatorio_{analista_selecionado.replace(' ', '_')}_{periodo.replace(' ', '_')}.md",
                    mime="text/markdown"
                )
    
    st.markdown("---")
    st.markdown(
        """
        <div style="text-align: center; color: #666; font-size: 12px;">
            Sistema de Performance v4.0 | Supabase + Histórico
        </div>
        """,
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
