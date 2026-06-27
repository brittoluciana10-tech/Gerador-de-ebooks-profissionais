import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
from reportlab.lib import colors
from datetime import datetime
from io import BytesIO

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
    
    /* ONBOARDING */
    .onboarding-container {
        background: linear-gradient(135deg, #e0f2fe 0%, #dbeafe 50%, #f0f9ff 100%);
        min-height: 100vh;
        padding: 60px 20px;
    }
    
    .onboarding-header {
        text-align: center;
        margin-bottom: 60px;
    }
    
    .onboarding-header h1 {
        font-size: 56px;
        font-weight: 800;
        color: #0c2340;
        margin: 0 0 12px 0;
        letter-spacing: -1px;
    }
    
    .onboarding-header p {
        font-size: 24px;
        color: #1e40af;
        margin: 0;
        font-weight: 500;
    }
    
    .card-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
        gap: 24px;
        max-width: 1200px;
        margin: 0 auto 80px auto;
        padding: 0 20px;
    }
    
    .option-card {
        background: white;
        border-radius: 16px;
        overflow: hidden;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.1);
        transition: all 0.3s ease;
        cursor: pointer;
    }
    
    .option-card:hover {
        transform: translateY(-8px);
        box-shadow: 0 20px 50px rgba(0, 0, 0, 0.15);
    }
    
    .card-image {
        width: 100%;
        height: 180px;
        background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 50%, #a855f7 100%);
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 72px;
    }
    
    .card-content {
        padding: 24px;
    }
    
    .card-content h3 {
        font-size: 20px;
        font-weight: 700;
        color: #0c2340;
        margin: 0 0 8px 0;
    }
    
    .card-content p {
        font-size: 14px;
        color: #64748b;
        margin: 0;
        line-height: 1.5;
    }
    
    .recent-prompts {
        max-width: 1000px;
        margin: 0 auto;
        padding: 0 20px;
    }
    
    .recent-prompts h2 {
        font-size: 28px;
        font-weight: 700;
        color: #0c2340;
        margin: 0 0 24px 0;
        text-align: center;
    }
    
    .prompt-item {
        background: white;
        border-radius: 12px;
        padding: 20px 24px;
        box-shadow: 0 5px 15px rgba(0, 0, 0, 0.08);
        margin-bottom: 12px;
        cursor: pointer;
        transition: all 0.3s ease;
    }
    
    .prompt-item:hover {
        box-shadow: 0 8px 20px rgba(0, 0, 0, 0.12);
        transform: translateX(4px);
    }
    
    .prompt-title {
        font-size: 16px;
        font-weight: 600;
        color: #0c2340;
        margin: 0 0 4px 0;
    }
    
    .prompt-meta {
        font-size: 12px;
        color: #94a3b8;
        margin: 0;
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
    </style>
""", unsafe_allow_html=True)

# FUNCAO: Gerar PDF Simples
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
    
    # CONTEÚDO
    story.append(PageBreak())
    
    linhas = conteudo.split('\n')
    current_chapter = 0
    
    for linha in linhas:
        linha = linha.strip()
        if not linha:
            story.append(Spacer(1, 0.08*inch))
        elif linha.startswith('###'):
            if current_chapter > 0:
                story.append(PageBreak())
            current_chapter += 1
            
            titulo = linha.replace('###', '').replace('**', '').strip()
            story.append(Paragraph(titulo, heading1))
            story.append(Spacer(1, 0.2*inch))
        elif linha.startswith('##'):
            titulo = linha.replace('##', '').replace('**', '').strip()
            story.append(Paragraph(titulo, heading2))
            story.append(Spacer(1, 0.1*inch))
        elif linha.startswith('[**'):
            pass
        elif linha.startswith('---'):
            story.append(Spacer(1, 0.2*inch))
        else:
            clean_text = linha.replace('**', '').replace('*', '').replace('`', '').strip()
            if clean_text and len(clean_text) > 2:
                try:
                    story.append(Paragraph(clean_text, body))
                except:
                    pass
    
    # IMAGENS
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

if "page" not in st.session_state:
    st.session_state.page = "onboarding"

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

# PAGE: ONBOARDING
if st.session_state.page == "onboarding":
    st.markdown("""
        <div class="onboarding-container">
            <div class="onboarding-header">
                <h1>Criar com IA</h1>
                <p>Como você quer começar?</p>
            </div>
        </div>
    """, unsafe_allow_html=True)
    
    # CARDS
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown("""
            <div class="option-card">
                <div class="card-image">📚</div>
                <div class="card-content">
                    <h3>Gerar</h3>
                    <p>Criar a partir de um prompt de uma linha em poucos segundos</p>
                </div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("Começar", key="btn_gerar", use_container_width=True):
            st.session_state.page = "criar"
            st.rerun()
    
    with col2:
        st.markdown("""
            <div class="option-card">
                <div class="card-image">✏️</div>
                <div class="card-content">
                    <h3>Colar texto</h3>
                    <p>Criar a partir de anotações, um esboço ou conteúdo existente</p>
                </div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("Colar", key="btn_colar", use_container_width=True):
            st.info("⏳ Funcionalidade em desenvolvimento!")
    
    with col3:
        st.markdown("""
            <div class="option-card">
                <div class="card-image">📋</div>
                <div class="card-content">
                    <h3>Usar modelo</h3>
                    <p>Crie usando a estrutura ou os layouts de um modelo</p>
                </div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("Modelo", key="btn_modelo", use_container_width=True):
            st.info("⏳ Funcionalidade em desenvolvimento!")
    
    with col4:
        st.markdown("""
            <div class="option-card">
                <div class="card-image">📂</div>
                <div class="card-content">
                    <h3>Importar</h3>
                    <p>Aprimorar documentos, apresentações ou páginas da Web existentes</p>
                </div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("Importar", key="btn_importar", use_container_width=True):
            st.info("⏳ Funcionalidade em desenvolvimento!")
    
    # PROMPTS RECENTES
    st.markdown("""
        <div style="margin-top: 80px;">
        <div class="recent-prompts">
            <h2>Seus prompts recentes</h2>
        </div>
        </div>
    """, unsafe_allow_html=True)
    
    recent_items = [
        {
            "title": "Quero que você escreva o conteúdo completo de um Infoproduto no form...",
            "meta": "Colar texto • há 7 horas"
        },
        {
            "title": "Gerar um ebook sobre maternidade real para mães portuguesas",
            "meta": "Gerar • há 2 dias"
        },
        {
            "title": "Criar tabelas de comparação e gráficos de dados para ebook",
            "meta": "Gerar • há 3 dias"
        }
    ]
    
    for item in recent_items:
        st.markdown(f"""
            <div class="prompt-item">
                <p class="prompt-title">{item['title']}</p>
                <p class="prompt-meta">{item['meta']}</p>
            </div>
        """, unsafe_allow_html=True)

# PAGE: CRIAR EBOOK
elif st.session_state.page == "criar":
    if st.button("← Voltar ao Início", key="back_criar"):
        st.session_state.page = "onboarding"
        st.rerun()
    
    st.markdown("""
        <div class="header-premium">
            <h1>📚 Ebook Creator Pro Premium</h1>
            <p>Com gráficos, tabelas, vetores e capítulos em páginas separadas</p>
        </div>
    """, unsafe_allow_html=True)
    
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
                    
                    extras = ""
                    if incluir_graficos:
                        extras += "- Incluir 1-2 dados/estatísticas por capítulo\n"
                    if incluir_tabelas:
                        extras += "- Incluir 1 tabela/comparação por capítulo\n"
                    if incluir_vetores:
                        extras += "- Incluir descrições de vetores\n"
                    
                    prompt = f"""Crie um ebook PROFISSIONAL em {idioma} sobre '{tema_desc}' para '{audience}' em {regiao}.

PARAMETROS:
- Capitulos: {num_chapters}
- Palavras/cap: {palavras_capitulo}
- Estilo: {estilo}
{extras}

ESTRUTURA:
- INTRODUCAO (2-3 paragrafos)
- {num_chapters} CAPITULOS (comece cada com ### CAPITULO X:)
- CONCLUSAO
- BONUS: 5 dicas

Escreva em {idioma}."""
                    
                    response = model.generate_content(prompt)
                    content = response.text
                    
                    progress_bar.progress(70)
                    status.text("📸 Gerando sugestões...")
                    
                    prompt_imagens = f"""Crie 10 prompts de imagem para '{tema_desc}' em estilo {ESTILOS_ARTE[estilo_arte_option]}.
                    
Formato: [N]. [Titulo]: [Descricao]"""
                    
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
                
                st.markdown('<div class="success-box"><strong>✅ Ebook gerado com sucesso!</strong></div>', unsafe_allow_html=True)
                
                res_col1, res_col2 = st.tabs(["📖 Conteúdo", "🎨 Imagens"])
                
                with res_col1:
                    with st.expander("Ver Conteúdo", expanded=True):
                        st.write(content)
                
                with res_col2:
                    with st.expander("Ver Sugestões", expanded=True):
                        st.write(sugestoes_imagens)
                
                st.markdown('<div class="divider-premium"></div>', unsafe_allow_html=True)
                st.markdown("### 📥 Baixar seu Ebook")
                
                col_d1, col_d2, col_d3 = st.columns(3)
                
                if formato_txt:
                    with col_d1:
                        txt = f"{tema_nome}\n\nPor Luciana Britto\n{idioma} | {regiao}\n\n{content}\n\n{sugestoes_imagens}"
                        st.download_button("📄 TXT", txt, f"{tema_nome}.txt", "text/plain", use_container_width=True)
                
                if formato_word:
                    with col_d2:
                        doc = Document()
                        title = doc.add_heading(tema_nome, 0)
                        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
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
                        st.download_button("🎨 PDF", pdf_data, f"{tema_nome}.pdf", "application/pdf", use_container_width=True)
                
            except Exception as e:
                st.error(f"❌ Erro: {str(e)}")

# FOOTER
st.markdown("""
    <div class="footer-custom">
        <p><strong>Luciana Britto | L&B Marketing — Estratégias de Valor</strong></p>
        <p>© 2026 • Ferramenta Premium de IA para Empreendoras</p>
        <p style="font-size: 12px; opacity: 0.7;">Powered by Google Gemini AI</p>
    </div>
""", unsafe_allow_html=True)
