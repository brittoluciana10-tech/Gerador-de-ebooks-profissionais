import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from reportlab.lib.pagesizes import letter
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
        <p>Crie ebooks profissionais com CONTROLE TOTAL + Imagens IA</p>
    </div>
""", unsafe_allow_html=True)

# TABS
tab1, tab2, tab3 = st.tabs(["✍️ Criar Ebook", "📚 Meus Ebooks", "📊 Dashboard"])

# TAB 1: CRIAR EBOOK
with tab1:
    st.markdown('<div class="divider-premium"></div>', unsafe_allow_html=True)
    
    # SECAO 1: BASICO
    st.markdown('<div class="section-header"><h2>1️⃣ Configuração Básica</h2></div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns([1.5, 1])
    with col1:
        st.markdown('<div class="info-box">Selecione o tema principal do seu ebook</div>', unsafe_allow_html=True)
        tema_option = st.selectbox(
            "Tema",
            list(TEMAS.keys()),
            label_visibility="collapsed"
        )
        tema_nome = tema_option.split(" ", 1)[1] if " " in tema_option else tema_option
        tema_desc = TEMAS[tema_option]
    
    with col2:
        st.markdown('<div class="info-box">Público que irá ler</div>', unsafe_allow_html=True)
        audience = st.text_input("👥 Público-alvo", "Mães portuguesas", label_visibility="collapsed")
    
    st.markdown('<div class="divider-premium"></div>', unsafe_allow_html=True)
    
    # SECAO 2: ESTRUTURA
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
    
    # SECAO 3: IDIOMA E LOCALIZACAO
    st.markdown('<div class="section-header"><h2>3️⃣ Idioma & Localização</h2></div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        idioma_option = st.selectbox("Idioma", list(IDIOMAS.keys()), label_visibility="collapsed")
        idioma = IDIOMAS[idioma_option]
    
    with col2:
        regiao = st.text_input("País/Região", "Portugal", label_visibility="collapsed")
    
    st.markdown('<div class="divider-premium"></div>', unsafe_allow_html=True)
    
    # SECAO 4: ESTILO VISUAL
    st.markdown('<div class="section-header"><h2>4️⃣ Estilo Visual (para Imagens)</h2></div>', unsafe_allow_html=True)
    
    estilo_arte_option = st.selectbox("Estilo de Arte", list(ESTILOS_ARTE.keys()), label_visibility="collapsed")
    estilo_arte = estilo_arte_option.split(" ", 1)[1] if " " in estilo_arte_option else estilo_arte_option
    
    st.markdown(f'<div class="info-box">Estilo selecionado: <strong>{ESTILOS_ARTE[estilo_arte_option]}</strong></div>', unsafe_allow_html=True)
    
    st.markdown('<div class="divider-premium"></div>', unsafe_allow_html=True)
    
    # SECAO 5: DOWNLOADS
    st.markdown('<div class="section-header"><h2>5️⃣ Formatos de Download</h2></div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        formato_txt = st.checkbox("📄 TXT", value=True)
    with col2:
        formato_word = st.checkbox("📋 Word", value=True)
    with col3:
        formato_pdf = st.checkbox("🎨 PDF", value=True)
    
    st.markdown('<div class="divider-premium"></div>', unsafe_allow_html=True)
    
    # BOTAO GERAR
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
                
                with st.spinner("⏳ Gerando conteúdo com IA Google Gemini..."):
                    status.text("🔄 Conectando à IA...")
                    progress_bar.progress(30)
                    
                    prompt = f"""Crie um ebook PROFISSIONAL em {idioma} sobre '{tema_desc}' para '{audience}' em {regiao}.

PARAMETROS:
- Capitulos: {num_chapters}
- Palavras/cap: {palavras_capitulo}
- Estilo: {estilo}
- Visual: {ESTILOS_ARTE[estilo_arte_option]}

ESTRUTURA:
- CAPA: [Espaco para imagem]
- TITULO
- INTRODUCAO (2-3 paragrafos)
- {num_chapters} CAPITULOS com conteudo detalhado
- CONCLUSAO
- BONUS: 5 dicas

Escreva COMPLETAMENTE em {idioma}."""
                    
                    response = model.generate_content(prompt)
                    content = response.text
                    
                    progress_bar.progress(70)
                    status.text("📸 Gerando sugestões de imagens...")
                    
                    prompt_imagens = f"""Crie 10 prompts de imagem para usar em Canva/DALL-E baseado neste ebook sobre '{tema_desc}'.
                    
Formato:
[Numero]. [Titulo]: [Descricao detalhada em {idioma} com estilo {ESTILOS_ARTE[estilo_arte_option]}]"""
                    
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
                
                # TABS RESULTADO
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
                        txt = f"{tema_nome}\n\nPor Luciana Britto | L&B Marketing\n{idioma} | {regiao}\n{datetime.now().strftime('%d/%m/%Y')}\n\n{content}\n\n--- IMAGENS ---\n{sugestoes_imagens}"
                        st.download_button("📄 Download TXT", txt, f"{tema_nome}.txt", "text/plain", use_container_width=True)
                
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
                        st.download_button("📋 Download Word", doc_bytes.getvalue(), f"{tema_nome}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
                
                if formato_pdf:
                    with col_d3:
                        buffer = BytesIO()
                        doc = SimpleDocTemplate(buffer, pagesize=letter)
                        story = []
                        styles = getSampleStyleSheet()
                        
                        title_style = ParagraphStyle('T', parent=styles['Heading1'], fontSize=28, textColor=colors.HexColor('#6366f1'), alignment=TA_CENTER)
                        body_style = ParagraphStyle('B', parent=styles['BodyText'], fontSize=11, alignment=TA_JUSTIFY)
                        
                        story.append(Paragraph(tema_nome, title_style))
                        story.append(Spacer(1, 0.2*inch))
                        story.append(Paragraph(f"{idioma} | {regiao}", styles['Normal']))
                        story.append(PageBreak())
                        story.append(Paragraph(content, body_style))
                        
                        doc.build(story)
                        buffer.seek(0)
                        st.download_button("🎨 Download PDF", buffer.getvalue(), f"{tema_nome}.pdf", "application/pdf", use_container_width=True)
                
            except Exception as e:
                st.error(f"❌ Erro: {str(e)}")

# TAB 2: MEUS EBOOKS
with tab2:
    st.markdown('<div class="divider-premium"></div>', unsafe_allow_html=True)
    
    if not st.session_state.ebooks:
        st.info("📭 Nenhum ebook criado ainda. Comece agora na aba 'Criar Ebook'!")
    else:
        st.markdown(f"### 📚 Total: {len(st.session_state.ebooks)} Ebooks Criados")
        
        for ebook in reversed(st.session_state.ebooks):
            with st.expander(f"📖 {ebook['tema']} ({ebook['idioma']}) • {ebook['capitulos']} cap • {ebook['data']}", expanded=False):
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.markdown(f"**👥 Público:** {ebook['publico']}")
                    st.markdown(f"**🎨 Estilo:** {ebook['estilo_arte']}")
                    st.write(ebook['conteudo'][:400] + "...")

# TAB 3: DASHBOARD
with tab3:
    st.markdown('<div class="divider-premium"></div>', unsafe_allow_html=True)
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f'<div class="metric-box"><h3>{len(st.session_state.ebooks)}</h3><p>Ebooks Criados</p></div>', unsafe_allow_html=True)
    
    with col2:
        total = sum([e['capitulos'] for e in st.session_state.ebooks]) if st.session_state.ebooks else 0
        st.markdown(f'<div class="metric-box"><h3>{total}</h3><p>Capítulos Gerados</p></div>', unsafe_allow_html=True)
    
    with col3:
        st.markdown('<div class="metric-box"><h3>12</h3><p>Temas Disponíveis</p></div>', unsafe_allow_html=True)
    
    with col4:
        st.markdown('<div class="metric-box"><h3>∞</h3><p>Uso GRÁTIS</p></div>', unsafe_allow_html=True)

# FOOTER
st.markdown("""
    <div class="footer-custom">
        <p><strong>Luciana Britto | L&B Marketing — Estratégias de Valor</strong></p>
        <p>© 2026 • Ferramenta Premium de IA para Empreendoras Portuguesas</p>
        <p style="font-size: 12px; opacity: 0.7;">Powered by Google Gemini AI</p>
    </div>
""", unsafe_allow_html=True)
