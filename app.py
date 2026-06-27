import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle, BarChart, PieChart, Drawing, AreaChart, LineChart
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib import colors
from reportlab.chart.barcharts import VerticalBarChart
from reportlab.chart.piecharts import PieChart as RLPieChart
from reportlab.chart.lineplots import LinePlot
from reportlab.chart.axes import XValueAxis, YValueAxis
from datetime import datetime
from io import BytesIO
import random

st.set_page_config(
    page_title="Ebook Creator Pro Premium",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# CSS DESIGN PROFISSIONAL
st.markdown("""
    <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    
    [data-testid="stAppViewContainer"] {
        background: linear-gradient(135deg, #0a0e27 0%, #1a1f3a 50%, #0f172a 100%);
        color: #f1f5f9;
    }
    
    [data-testid="stHeader"] { background: transparent !important; }
    
    .stButton>button {
        background: linear-gradient(135deg, #6366f1 0%, #4f46e5 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 14px 28px !important;
        font-weight: 700 !important;
        font-size: 16px !important;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
        box-shadow: 0 8px 24px rgba(99, 102, 241, 0.35) !important;
        letter-spacing: 0.5px !important;
    }
    
    .stButton>button:hover {
        transform: translateY(-4px) !important;
        box-shadow: 0 12px 32px rgba(99, 102, 241, 0.5) !important;
        background: linear-gradient(135deg, #4f46e5 0%, #4338ca 100%) !important;
    }
    
    .stButton>button:active {
        transform: translateY(-2px) !important;
    }
    
    .header-premium {
        background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 50%, #a855f7 100%);
        border-radius: 20px;
        padding: 48px 32px;
        margin: 24px 0;
        text-align: center;
        box-shadow: 0 20px 60px rgba(99, 102, 241, 0.3);
        border: 1px solid rgba(255, 255, 255, 0.1);
    }
    
    .header-premium h1 {
        font-size: 52px !important;
        font-weight: 800 !important;
        margin: 0 !important;
        color: white !important;
        text-shadow: 0 2px 10px rgba(0, 0, 0, 0.3);
        letter-spacing: -1px;
    }
    
    .header-premium p {
        font-size: 20px !important;
        color: rgba(255, 255, 255, 0.95) !important;
        margin: 12px 0 0 0 !important;
        font-weight: 500;
    }
    
    .card-container {
        background: linear-gradient(135deg, rgba(30, 41, 59, 0.8) 0%, rgba(15, 23, 42, 0.9) 100%);
        border: 1px solid rgba(99, 102, 241, 0.2);
        border-radius: 16px;
        padding: 28px;
        margin: 20px 0;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3);
        backdrop-filter: blur(10px);
    }
    
    .section-header {
        display: flex;
        align-items: center;
        gap: 12px;
        margin: 32px 0 20px 0;
        padding-bottom: 16px;
        border-bottom: 2px solid rgba(99, 102, 241, 0.3);
    }
    
    .section-header h2 {
        color: #6366f1 !important;
        font-size: 22px !important;
        font-weight: 700 !important;
        margin: 0 !important;
    }
    
    .metric-box {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%);
        border-radius: 12px;
        padding: 20px;
        text-align: center;
        box-shadow: 0 8px 20px rgba(16, 185, 129, 0.3);
        border: 1px solid rgba(255, 255, 255, 0.1);
    }
    
    .metric-box h3 {
        color: white !important;
        font-size: 36px !important;
        margin: 0 !important;
        font-weight: 800;
    }
    
    .metric-box p {
        color: rgba(255, 255, 255, 0.85) !important;
        font-size: 13px !important;
        margin: 8px 0 0 0 !important;
        text-transform: uppercase;
        letter-spacing: 1px;
        font-weight: 600;
    }
    
    [data-testid="stTabs"] [role="tablist"] {
        gap: 8px;
        border: none !important;
    }
    
    [data-testid="stTabs"] [role="tab"] {
        background: rgba(99, 102, 241, 0.1) !important;
        border: 1px solid rgba(99, 102, 241, 0.2) !important;
        border-radius: 10px !important;
        padding: 12px 24px !important;
        color: #cbd5e1 !important;
        font-weight: 600 !important;
        transition: all 0.3s ease !important;
    }
    
    [data-testid="stTabs"] [role="tab"][aria-selected="true"] {
        background: linear-gradient(135deg, #6366f1 0%, #4f46e5 100%) !important;
        color: white !important;
        border: 1px solid #4f46e5 !important;
        box-shadow: 0 8px 16px rgba(99, 102, 241, 0.3) !important;
    }
    
    .stSlider [role="slider"] {
        background: linear-gradient(135deg, #6366f1 0%, #4f46e5 100%) !important;
    }
    
    .stSelectbox, .stTextInput, .stNumberInput {
        border-radius: 10px !important;
    }
    
    .divider-premium {
        height: 2px;
        background: linear-gradient(90deg, transparent 0%, #6366f1 50%, transparent 100%);
        margin: 40px 0;
    }
    
    .info-box {
        background: linear-gradient(135deg, rgba(99, 102, 241, 0.1) 0%, rgba(139, 92, 246, 0.05) 100%);
        border-left: 4px solid #6366f1;
        border-radius: 8px;
        padding: 16px 20px;
        margin: 12px 0;
    }
    
    .success-box {
        background: linear-gradient(135deg, rgba(16, 185, 129, 0.1) 0%, rgba(5, 150, 105, 0.05) 100%);
        border-left: 4px solid #10b981;
        border-radius: 8px;
        padding: 16px 20px;
        margin: 12px 0;
    }
    
    footer { display: none !important; }
    
    .footer-custom {
        text-align: center;
        padding: 32px 24px;
        margin-top: 48px;
        background: linear-gradient(135deg, rgba(99, 102, 241, 0.05) 0%, rgba(139, 92, 246, 0.05) 100%);
        border-top: 1px solid rgba(99, 102, 241, 0.2);
        border-radius: 16px;
    }
    
    .footer-custom p {
        color: #cbd5e1 !important;
        font-size: 14px !important;
        margin: 4px 0 !important;
    }
    
    .footer-custom strong {
        color: #f1f5f9 !important;
    }
    </style>
""", unsafe_allow_html=True)

# FUNCAO: Criar gráfico de barras
def criar_grafico_barras():
    drawing = Drawing(400, 250)
    bc = VerticalBarChart()
    bc.x = 50
    bc.y = 50
    bc.width = 300
    bc.height = 150
    bc.data = [[10, 20, 30, 15, 25]]
    bc.strokeColor = colors.black
    bc.fillColor = colors.HexColor('#6366f1')
    bc.categoryAxis.labels.boxAnchor = 'n'
    bc.categoryAxis.labels.angle = 0
    bc.categoryAxis.categoryNames = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai']
    drawing.add(bc)
    return drawing

# FUNCAO: Criar gráfico de pizza
def criar_grafico_pizza():
    drawing = Drawing(400, 250)
    pc = RLPieChart()
    pc.x = 100
    pc.y = 50
    pc.width = 200
    pc.height = 150
    pc.data = [25, 35, 20, 20]
    pc.labels = ['Categoria A', 'Categoria B', 'Categoria C', 'Categoria D']
    pc.slices.strokeWidth = 2
    drawing.add(pc)
    return drawing

# FUNCAO: Gerar PDF Profissional com Gráficos e Tabelas
def gerar_pdf_profissional(tema, audience, idioma, regiao, conteudo, sugestoes_imagens):
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        topMargin=0.8*inch,
        bottomMargin=0.8*inch,
        leftMargin=0.8*inch,
        rightMargin=0.8*inch
    )
    
    story = []
    styles = getSampleStyleSheet()
    
    # ESTILOS CUSTOMIZADOS
    titulo_capa = ParagraphStyle(
        'TituloCapa',
        parent=styles['Heading1'],
        fontSize=48,
        textColor=colors.HexColor('#6366f1'),
        spaceAfter=20,
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    )
    
    subtitulo_capa = ParagraphStyle(
        'SubtituloCapa',
        parent=styles['Normal'],
        fontSize=24,
        textColor=colors.HexColor('#8b5cf6'),
        spaceAfter=40,
        alignment=TA_CENTER,
        fontName='Helvetica'
    )
    
    info_capa = ParagraphStyle(
        'InfoCapa',
        parent=styles['Normal'],
        fontSize=12,
        textColor=colors.HexColor('#64748b'),
        spaceAfter=8,
        alignment=TA_CENTER
    )
    
    heading1 = ParagraphStyle(
        'CustomHeading1',
        parent=styles['Heading1'],
        fontSize=28,
        textColor=colors.HexColor('#0f172a'),
        spaceAfter=12,
        spaceBefore=12,
        fontName='Helvetica-Bold'
    )
    
    heading2 = ParagraphStyle(
        'CustomHeading2',
        parent=styles['Heading2'],
        fontSize=16,
        textColor=colors.HexColor('#1e293b'),
        spaceAfter=10,
        spaceBefore=8,
        fontName='Helvetica-Bold'
    )
    
    body = ParagraphStyle(
        'CustomBody',
        parent=styles['BodyText'],
        fontSize=10,
        alignment=TA_JUSTIFY,
        spaceAfter=10,
        leading=15,
        textColor=colors.HexColor('#334155')
    )
    
    # PÁGINA 1: CAPA
    story.append(Spacer(1, 1.5*inch))
    story.append(Paragraph("📚", titulo_capa))
    story.append(Spacer(1, 0.2*inch))
    story.append(Paragraph(tema, titulo_capa))
    story.append(Spacer(1, 0.3*inch))
    story.append(Paragraph(f"Um guia essencial para {audience}", subtitulo_capa))
    story.append(Spacer(1, 1.5*inch))
    story.append(Paragraph("Por Luciana Britto | L&B Marketing", info_capa))
    story.append(Spacer(1, 0.1*inch))
    story.append(Paragraph("Estratégias de Valor", info_capa))
    story.append(Spacer(1, 0.2*inch))
    story.append(Paragraph(f"Idioma: {idioma} | Região: {regiao}", info_capa))
    story.append(Spacer(1, 0.1*inch))
    story.append(Paragraph(f"© 2026 — Todos os direitos reservados", info_capa))
    story.append(Spacer(1, 0.1*inch))
    story.append(Paragraph(datetime.now().strftime("Gerado em %d de %B de %Y"), info_capa))
    
    # PÁGINA 2: ÍNDICE
    story.append(PageBreak())
    story.append(Paragraph("ÍNDICE", heading1))
    story.append(Spacer(1, 0.2*inch))
    
    indice_items = [
        "Introdução",
        "Capítulo 1",
        "Capítulo 2",
        "Capítulo 3",
        "Capítulo 4",
        "Capítulo 5",
        "Conclusão",
        "Sugestões de Imagens",
    ]
    
    for idx, item in enumerate(indice_items, 1):
        story.append(Paragraph(f"{idx}. {item}", body))
    
    # PÁGINA 3+: CONTEÚDO COM CAPÍTULOS EM NOVAS PÁGINAS
    story.append(PageBreak())
    
    linhas = conteudo.split('\n')
    current_chapter = 0
    
    for linha in linhas:
        linha = linha.strip()
        if not linha:
            story.append(Spacer(1, 0.08*inch))
        elif linha.startswith('###'):
            # Nova página para novo capítulo
            if current_chapter > 0:
                story.append(PageBreak())
            current_chapter += 1
            
            titulo = linha.replace('###', '').replace('**', '').strip()
            story.append(Paragraph(titulo, heading1))
            story.append(Spacer(1, 0.2*inch))
            
            # Adicionar gráfico aleatório
            if current_chapter % 2 == 0:
                try:
                    story.append(criar_grafico_barras())
                    story.append(Spacer(1, 0.15*inch))
                except:
                    pass
            elif current_chapter % 3 == 0:
                try:
                    story.append(criar_grafico_pizza())
                    story.append(Spacer(1, 0.15*inch))
                except:
                    pass
        elif linha.startswith('##'):
            titulo = linha.replace('##', '').replace('**', '').strip()
            story.append(Paragraph(titulo, heading2))
            story.append(Spacer(1, 0.1*inch))
        elif linha.startswith('[**'):
            pass
        elif linha.startswith('---'):
            story.append(Spacer(1, 0.2*inch))
        elif '|' in linha and linha.count('|') > 2:
            # Detectar tabelas (linhas com |)
            try:
                rows = [cell.strip() for cell in linha.split('|') if cell.strip()]
                if rows:
                    table_data = [rows]
                    table = Table(table_data, colWidths=[2*inch, 2*inch])
                    table.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#6366f1')),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0, 0), (-1, 0), 10),
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                        ('GRID', (0, 0), (-1, -1), 1, colors.black)
                    ]))
                    story.append(table)
                    story.append(Spacer(1, 0.15*inch))
            except:
                clean_text = linha.replace('**', '').replace('*', '').strip()
                if clean_text and len(clean_text) > 2:
                    story.append(Paragraph(clean_text, body))
        else:
            clean_text = linha.replace('**', '').replace('*', '').replace('`', '').strip()
            if clean_text and len(clean_text) > 2:
                try:
                    story.append(Paragraph(clean_text, body))
                except:
                    pass
    
    # PÁGINA FINAL: IMAGENS
    story.append(PageBreak())
    story.append(Paragraph("SUGESTÕES DE IMAGENS PARA CANVA", heading1))
    story.append(Spacer(1, 0.2*inch))
    
    linhas_imagens = sugestoes_imagens.split('\n')
    for linha in linhas_imagens:
        linha = linha.strip()
        if linha and len(linha) > 2:
            clean = linha.replace('**', '').replace('*', '')
            try:
                story.append(Paragraph(clean, body))
            except:
                pass
    
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# INICIALIZAR STATE
if "ebooks" not in st.session_state:
    st.session_state.ebooks = []

# TEMAS
TEMAS = {
    "🤰 Maternidade Real": "Maternidade real, mães, experiências honestas",
    "👶 Sono do Bebé": "Sono infantil, técnicas de dormir, rotina",
    "💪 Pós-parto": "Recuperação pós-parto, saúde, bem-estar",
    "📚 Parenting": "Educação infantil, desenvolvimento, parentalidade",
    "💰 Renda Variável": "Renda extra, freelance, negócios digitais",
    "🧘 Bem-estar": "Mindfulness, meditação, saúde mental",
    "💼 Autónomo": "Ser autónomo, gestão, financeiro",
    "⚖️ Peso Saudável": "Perda de peso, fitness, saúde, nutrição",
    "🎨 Criatividade": "Criatividade, inovação, arte, inspiração",
    "💳 Finanças": "Educação financeira, investimentos, economia",
    "❤️ Relacionamentos": "Relacionamentos, amor, comunicação",
    "🚀 Carreira": "Carreira, emprego, desenvolvimento profissional",
}

IDIOMAS = {
    "🇵🇹 Português": "Português",
    "🇬🇧 Inglês": "English",
    "🇪🇸 Espanhol": "Español",
    "🇫🇷 Francês": "Français",
    "🇮🇹 Italiano": "Italiano",
}

ESTILOS_ARTE = {
    "📷 Foto": "Fotografia realista, alta qualidade",
    "🎨 Ilustração": "Ilustração artística e colorida",
    "🔷 Abstrato": "Arte abstrata e moderna",
    "🌐 Digital": "Arte digital e futurista",
    "🌊 Aquarela": "Estilo aquarela e suave",
    "⬜ Minimalista": "Design minimalista e limpo",
}

# HEADER
st.markdown("""
    <div class="header-premium">
        <h1>📚 Ebook Creator Pro Premium</h1>
        <p>Com gráficos, tabelas, vetores e capítulos em páginas separadas</p>
    </div>
""", unsafe_allow_html=True)

# TABS
tab1, tab2, tab3 = st.tabs(["✍️ Criar Ebook", "📚 Meus Ebooks", "📊 Dashboard"])

# TAB 1: CRIAR EBOOK
with tab1:
    st.markdown('<div class="divider-premium"></div>', unsafe_allow_html=True)
    
    st.markdown('<div class="section-header"><h2>1️⃣ Configuração Básica</h2></div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns([1.5, 1])
    with col1:
        st.markdown('<div class="info-box">Selecione o tema principal</div>', unsafe_allow_html=True)
        tema_option = st.selectbox("Tema", list(TEMAS.keys()), label_visibility="collapsed")
        tema_nome = tema_option.split(" ", 1)[1] if " " in tema_option else tema_option
        tema_desc = TEMAS[tema_option]
    
    with col2:
        st.markdown('<div class="info-box">Seu público-alvo</div>', unsafe_allow_html=True)
        audience = st.text_input("Público-alvo", "Mães portuguesas", label_visibility="collapsed")
    
    st.markdown('<div class="divider-premium"></div>', unsafe_allow_html=True)
    
    st.markdown('<div class="section-header"><h2>2️⃣ Estrutura do Ebook</h2></div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("**📑 Capítulos**")
        num_chapters = st.slider("Quantidade", 2, 10, 3, label_visibility="collapsed")
        st.caption(f"Total: {num_chapters} capítulos")
    
    with col2:
        st.markdown("**📝 Palavras/Capítulo**")
        palavras_capitulo = st.select_slider(
            "Palavras",
            options=[200, 400, 600, 800, 1000, 1200],
            value=600,
            label_visibility="collapsed"
        )
        st.caption(f"~{palavras_capitulo} palavras")
    
    with col3:
        st.markdown("**✍️ Estilo de Escrita**")
        estilo = st.selectbox(
            "Estilo",
            ["Profissional", "Descontraído", "Científico", "Inspirador"],
            label_visibility="collapsed"
        )
        st.caption(f"{estilo}")
    
    st.markdown('<div class="divider-premium"></div>', unsafe_allow_html=True)
    
    st.markdown('<div class="section-header"><h2>3️⃣ Idioma & Localização</h2></div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        idioma_option = st.selectbox("Idioma", list(IDIOMAS.keys()), label_visibility="collapsed")
        idioma = IDIOMAS[idioma_option]
    
    with col2:
        regiao = st.text_input("País/Região", "Portugal", label_visibility="collapsed")
    
    st.markdown('<div class="divider-premium"></div>', unsafe_allow_html=True)
    
    st.markdown('<div class="section-header"><h2>4️⃣ Estilo Visual</h2></div>', unsafe_allow_html=True)
    
    estilo_arte_option = st.selectbox("Estilo de Arte", list(ESTILOS_ARTE.keys()), label_visibility="collapsed")
    estilo_arte = estilo_arte_option.split(" ", 1)[1] if " " in estilo_arte_option else estilo_arte_option
    
    st.markdown(f'<div class="info-box">Estilo: <strong>{ESTILOS_ARTE[estilo_arte_option]}</strong></div>', unsafe_allow_html=True)
    
    st.markdown('<div class="divider-premium"></div>', unsafe_allow_html=True)
    
    st.markdown('<div class="section-header"><h2>5️⃣ Conteúdo Adicional</h2></div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        incluir_graficos = st.checkbox("📊 Incluir Gráficos", value=True)
    with col2:
        incluir_tabelas = st.checkbox("📋 Incluir Tabelas", value=True)
    with col3:
        incluir_vetores = st.checkbox("🎨 Incluir Vetores", value=True)
    
    st.markdown('<div class="divider-premium"></div>', unsafe_allow_html=True)
    
    st.markdown('<div class="section-header"><h2>6️⃣ Formatos de Download</h2></div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        formato_txt = st.checkbox("📄 TXT", value=True)
    with col2:
        formato_word = st.checkbox("📋 Word", value=True)
    with col3:
        formato_pdf = st.checkbox("🎨 PDF Premium", value=True)
    
    st.markdown('<div class="divider-premium"></div>', unsafe_allow_html=True)
    
    col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
    with col_btn2:
        if st.button("🚀 GERAR EBOOK PREMIUM", use_container_width=True, key="generate"):
            try:
                api_key = st.secrets["GEMINI_API_KEY"]
                genai.configure(api_key=api_key)
                
                models = [m.name for m in genai.list_models()]
                model_name = [m for m in models if "gemini" in m.lower()][0]
                model = genai.GenerativeModel(model_name)
                
                progress_bar = st.progress(0)
                status = st.empty()
                
                with st.spinner("⏳ Gerando conteúdo com IA..."):
                    status.text("🔄 Conectando à IA...")
                    progress_bar.progress(30)
                    
                    # Adicionar requisitos de gráficos/tabelas
                    extras = ""
                    if incluir_graficos:
                        extras += "- Incluir 1-2 dados/estatísticas por capítulo para gráficos\n"
                    if incluir_tabelas:
                        extras += "- Incluir 1 tabela/comparação por capítulo\n"
                    if incluir_vetores:
                        extras += "- Incluir descrições de vetores/diagramas para ilustrar conceitos\n"
                    
                    prompt = f"""Crie um ebook PROFISSIONAL em {idioma} sobre '{tema_desc}' para '{audience}' em {regiao}.

PARAMETROS:
- Capitulos: {num_chapters}
- Palavras/cap: {palavras_capitulo}
- Estilo: {estilo}
{extras}

ESTRUTURA:
- INTRODUCAO (2-3 paragrafos)
- {num_chapters} CAPITULOS (CADA UM COM DADOS/TABELA/VETOR)
- CONCLUSAO
- BONUS: 5 dicas

Escreva COMPLETAMENTE em {idioma}. IMPORTANTE: Cada capítulo deve iniciar com '### CAPÍTULO X:' e ter dados/números para gráficos."""
                    
                    response = model.generate_content(prompt)
                    content = response.text
                    
                    progress_bar.progress(70)
                    status.text("📸 Gerando sugestões...")
                    
                    prompt_imagens = f"""Crie 10 prompts de imagem para '{tema_desc}' em estilo {ESTILOS_ARTE[estilo_arte_option]}.
                    
Formato:
[N]. [Titulo]: [Descricao em {idioma}]"""
                    
                    response_imagens = model.generate_content(prompt_imagens)
                    sugestoes_imagens = response_imagens.text
                    
                    ebook_data = {
                        "id": len(st.session_state.ebooks) + 1,
                        "tema": tema_nome,
                        "idioma": idioma,
                        "data": datetime.now().strftime("%d/%m/%Y %H:%M"),
                        "publico": audience,
                        "capitulos": num_chapters,
                        "palavras": palavras_capitulo,
                        "estilo_arte": estilo_arte,
                        "conteudo": content,
                        "sugestoes_imagens": sugestoes_imagens
                    }
                    st.session_state.ebooks.append(ebook_data)
                    
                    progress_bar.progress(100)
                    status.empty()
                
                st.markdown('<div class="success-box"><strong>✅ Ebook gerado com sucesso! Com gráficos, tabelas e capítulos em páginas separadas!</strong></div>', unsafe_allow_html=True)
                
                res_col1, res_col2 = st.tabs(["📖 Conteúdo", "🎨 Prompts de Imagens"])
                
                with res_col1:
                    with st.expander("Ver Conteúdo Completo", expanded=True):
                        st.write(content)
                
                with res_col2:
                    with st.expander("Ver Sugestões de Imagens", expanded=True):
                        st.write(sugestoes_imagens)
                
                st.markdown('<div class="divider-premium"></div>', unsafe_allow_html=True)
                st.markdown("### 📥 Baixar seu Ebook")
                
                col_d1, col_d2, col_d3 = st.columns(3)
                
                if formato_txt:
                    with col_d1:
                        txt = f"{tema_nome}\n\nPor Luciana Britto | L&B Marketing\n{idioma} | {regiao}\n{datetime.now().strftime('%d/%m/%Y')}\n\n{content}\n\n--- IMAGENS ---\n{sugestoes_imagens}\n\nCopyright 2026 Luciana Britto"
                        st.download_button("📄 TXT", txt, f"{tema_nome}.txt", "text/plain", use_container_width=True)
                
                if formato_word:
                    with col_d2:
                        doc = Document()
                        title = doc.add_heading(tema_nome, 0)
                        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        info = doc.add_paragraph()
                        info.add_run(f"Por Luciana Britto | L&B Marketing\n{idioma} | {regiao}").bold = True
                        info.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        doc.add_paragraph(content)
                        doc.add_page_break()
                        doc.add_heading("Imagens Sugeridas", level=1)
                        doc.add_paragraph(sugestoes_imagens)
                        
                        doc_bytes = BytesIO()
                        doc.save(doc_bytes)
                        doc_bytes.seek(0)
                        st.download_button("📋 Word", doc_bytes.getvalue(), f"{tema_nome}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
                
                if formato_pdf:
                    with col_d3:
                        pdf_data = gerar_pdf_profissional(tema_nome, audience, idioma, regiao, content, sugestoes_imagens)
                        st.download_button("🎨 PDF Premium", pdf_data, f"{tema_nome}.pdf", "application/pdf", use_container_width=True)
                
            except Exception as e:
                st.error(f"❌ Erro: {str(e)}")

# TAB 2: MEUS EBOOKS
with tab2:
    st.markdown('<div class="divider-premium"></div>', unsafe_allow_html=True)
    
    if not st.session_state.ebooks:
        st.info("📭 Nenhum ebook criado. Comece agora!")
    else:
        st.markdown(f"### 📚 Total: {len(st.session_state.ebooks)} Ebooks")
        
        for ebook in reversed(st.session_state.ebooks):
            with st.expander(f"📖 {ebook['tema']} ({ebook['idioma']}) • {ebook['capitulos']} cap • {ebook['data']}", expanded=False):
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.markdown(f"**Público:** {ebook['publico']}")
                    st.markdown(f"**Estilo:** {ebook['estilo_arte']}")
                    st.write(ebook['conteudo'][:400] + "...")

# TAB 3: DASHBOARD
with tab3:
    st.markdown('<div class="divider-premium"></div>', unsafe_allow_html=True)
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f'<div class="metric-box"><h3>{len(st.session_state.ebooks)}</h3><p>Ebooks Criados</p></div>', unsafe_allow_html=True)
    
    with col2:
        total = sum([e['capitulos'] for e in st.session_state.ebooks]) if st.session_state.ebooks else 0
        st.markdown(f'<div class="metric-box"><h3>{total}</h3><p>Capítulos</p></div>', unsafe_allow_html=True)
    
    with col3:
        st.markdown('<div class="metric-box"><h3>12</h3><p>Temas</p></div>', unsafe_allow_html=True)
    
    with col4:
        st.markdown('<div class="metric-box"><h3>∞</h3><p>Uso GRÁTIS</p></div>', unsafe_allow_html=True)

# FOOTER
st.markdown("""
    <div class="footer-custom">
        <p><strong>Luciana Britto | L&B Marketing — Estratégias de Valor</strong></p>
        <p>© 2026 • Ferramenta Premium de IA para Empreendoras</p>
        <p style="font-size: 12px; opacity: 0.7;">Powered by Google Gemini AI + ReportLab</p>
    </div>
""", unsafe_allow_html=True)
