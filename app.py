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
    page_title="Ebook Creator Pro",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# CSS PROFISSIONAL
st.markdown("""
    <style>
    * { margin: 0; padding: 0; }
    html, body, [class*="css"] { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif !important; }
    
    .main { 
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        color: #f1f5f9;
    }
    
    .stButton>button {
        background: linear-gradient(135deg, #3b82f6 0%, #1d4ed8 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 12px 24px !important;
        font-weight: 600 !important;
        font-size: 16px !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 4px 15px rgba(59, 130, 246, 0.4) !important;
    }
    
    .stButton>button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 20px rgba(59, 130, 246, 0.6) !important;
    }
    
    .stSelectbox, .stTextInput, .stSlider { 
        border-radius: 8px !important;
    }
    
    .card {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        border: 1px solid rgba(59, 130, 246, 0.2);
        border-radius: 12px;
        padding: 24px;
        margin: 12px 0;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
    }
    
    .header-card {
        background: linear-gradient(135deg, #3b82f6 0%, #1d4ed8 100%);
        border-radius: 12px;
        padding: 32px;
        margin: 24px 0;
        color: white;
        text-align: center;
    }
    
    .metric-card {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%);
        border-radius: 8px;
        padding: 16px;
        color: white;
        text-align: center;
    }
    
    h1, h2, h3 { color: #ffffff !important; }
    
    .divider {
        background: linear-gradient(90deg, transparent 0%, #3b82f6 50%, transparent 100%);
        height: 2px;
        margin: 32px 0;
    }
    </style>
""", unsafe_allow_html=True)

# INICIALIZAR STATE
if "ebooks" not in st.session_state:
    st.session_state.ebooks = []
if "current_content" not in st.session_state:
    st.session_state.current_content = None

# TEMAS
TEMAS = {
    "Maternidade Real": "Maternidade real, mães, experiências honestas",
    "Sono do Bebé": "Sono infantil, técnicas de dormir, rotina",
    "Pós-parto": "Recuperação pós-parto, saúde, bem-estar",
    "Parenting": "Educação infantil, desenvolvimento, parentalidade",
    "Renda Variável": "Renda extra, freelance, negócios digitais",
    "Bem-estar": "Mindfulness, meditação, saúde mental",
    "Autónomo": "Ser autónomo, gestão, financeiro",
    "Peso Saudável": "Perda de peso, fitness, saúde, nutrição",
    "Criatividade": "Criatividade, inovação, arte, inspiração",
    "Finanças": "Educação financeira, investimentos, economia",
    "Relacionamentos": "Relacionamentos, amor, comunicação",
    "Carreira": "Carreira, emprego, desenvolvimento profissional",
}

# HEADER
st.markdown("""
    <div class="header-card">
        <h1 style="margin: 0; font-size: 48px;">📚 Ebook Creator Pro</h1>
        <p style="margin: 12px 0 0 0; font-size: 18px; opacity: 0.9;">Crie ebooks profissionais em segundos</p>
    </div>
""", unsafe_allow_html=True)

# TABS
tab1, tab2, tab3 = st.tabs(["Criar Ebook", "Meus Ebooks", "Dashboard"])

# TAB 1: CRIAR EBOOK
with tab1:
    st.markdown("<div class='divider'></div>", unsafe_allow_html=True)
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("### Configurar seu Ebook")
        tema_nome = st.selectbox(
            "Escolha o tema",
            list(TEMAS.keys()),
            label_visibility="collapsed"
        )
        tema_desc = TEMAS[tema_nome]
    
    with col2:
        st.markdown("### Público")
        audience = st.text_input(
            "Público-alvo",
            "Mães portuguesas",
            label_visibility="collapsed"
        )
    
    st.markdown("<div class='divider'></div>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        num_chapters = st.slider("Capitulos", 2, 5, 3, label_visibility="collapsed")
    
    with col2:
        st.markdown("**Download**")
        formato_txt = st.checkbox("TXT", value=True)
        formato_word = st.checkbox("Word", value=True)
    
    with col3:
        st.markdown("**Opcoes**")
        formato_pdf = st.checkbox("PDF", value=True)
        estilo = st.selectbox("Estilo", ["Profissional", "Descontraido", "Cientifico"], label_visibility="collapsed")
    
    st.markdown("<div class='divider'></div>", unsafe_allow_html=True)
    
    # BOTAO GERAR
    col_btn = st.columns([1, 2, 1])
    with col_btn[1]:
        if st.button("Gerar Ebook Premium", use_container_width=True, key="generate_btn"):
            try:
                api_key = st.secrets["GEMINI_API_KEY"]
                genai.configure(api_key=api_key)
                
                models = [m.name for m in genai.list_models()]
                model_name = [m for m in models if "gemini" in m.lower()][0]
                model = genai.GenerativeModel(model_name)
                
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                with st.spinner("Gerando conteudo com IA..."):
                    status_text.text("Conectando a IA...")
                    progress_bar.progress(25)
                    
                    prompt = f"""Crie um ebook PROFISSIONAL sobre '{tema_desc}' para '{audience}'.

ESTRUTURA OBRIGATORIA:
- TITULO: (titulo atrativo e profissional)
- INTRODUCAO: (2-3 paragrafos inspiradores)
- {num_chapters} CAPITULOS:
  * Titulo do capitulo
  * Conteudo detalhado (400+ palavras)
  * 3 dicas praticas
- CONCLUSAO: (mensagem final inspiradora)
- BONUS: (5 dicas extras)

ESTILO: {estilo}
QUALIDADE: Profissional, detalhado e altamente pratico
FOCO: Resolver problemas reais de '{audience}'"""
                    
                    response = model.generate_content(prompt)
                    content = response.text
                    st.session_state.current_content = content
                    
                    progress_bar.progress(70)
                    status_text.text("Processando conteudo...")
                    
                    # Adicionar ao historico
                    ebook_data = {
                        "id": len(st.session_state.ebooks) + 1,
                        "tema": tema_nome,
                        "data": datetime.now().strftime("%d/%m/%Y %H:%M"),
                        "publico": audience,
                        "capitulos": num_chapters,
                        "conteudo": content
                    }
                    st.session_state.ebooks.append(ebook_data)
                    
                    progress_bar.progress(100)
                    status_text.empty()
                
                # RESULTADO
                st.markdown("<div class='divider'></div>", unsafe_allow_html=True)
                st.success("Ebook gerado com sucesso!")
                
                # PREVIEW
                with st.expander("Ver Preview", expanded=True):
                    st.write(content[:1000] + "...")
                
                # DOWNLOADS
                st.markdown("### Download seu Ebook")
                
                col_d1, col_d2, col_d3 = st.columns(3)
                
                # TXT
                if formato_txt:
                    with col_d1:
                        txt = f"{tema_nome}\n\nPor Luciana Britto | L&B Marketing\n{datetime.now().strftime('%d/%m/%Y')}\n\n{content}\n\nCopyright 2026 Luciana Britto | L&B Marketing"
                        st.download_button(
                            "Download TXT",
                            txt,
                            f"{tema_nome}.txt",
                            "text/plain",
                            use_container_width=True
                        )
                
                # WORD
                if formato_word:
                    with col_d2:
                        doc = Document()
                        title = doc.add_heading(tema_nome, 0)
                        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        subtitle = doc.add_paragraph("Por Luciana Britto | L&B Marketing")
                        subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        doc.add_paragraph(f"Gerado em {datetime.now().strftime('%d/%m/%Y')}")
                        doc.add_paragraph()
                        doc.add_paragraph(content)
                        
                        doc_bytes = BytesIO()
                        doc.save(doc_bytes)
                        doc_bytes.seek(0)
                        
                        st.download_button(
                            "Download Word",
                            doc_bytes.getvalue(),
                            f"{tema_nome}.docx",
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                
                # PDF
                if formato_pdf:
                    with col_d3:
                        buffer = BytesIO()
                        doc = SimpleDocTemplate(buffer, pagesize=letter)
                        story = []
                        styles = getSampleStyleSheet()
                        
                        title_style = ParagraphStyle(
                            'CustomTitle',
                            parent=styles['Heading1'],
                            fontSize=28,
                            textColor=colors.HexColor('#3b82f6'),
                            spaceAfter=12,
                            alignment=TA_CENTER,
                            fontName='Helvetica-Bold'
                        )
                        
                        body_style = ParagraphStyle(
                            'CustomBody',
                            parent=styles['BodyText'],
                            fontSize=11,
                            alignment=TA_JUSTIFY,
                            spaceAfter=12,
                            leading=16
                        )
                        
                        story.append(Paragraph(tema_nome, title_style))
                        story.append(Spacer(1, 0.2*inch))
                        story.append(Paragraph("Por Luciana Britto | L&B Marketing", styles['Normal']))
                        story.append(Paragraph(f"{datetime.now().strftime('%d de %B de %Y')}", styles['Normal']))
                        story.append(PageBreak())
                        story.append(Paragraph(content, body_style))
                        
                        doc.build(story)
                        buffer.seek(0)
                        
                        st.download_button(
                            "Download PDF",
                            buffer.getvalue(),
                            f"{tema_nome}.pdf",
                            "application/pdf",
                            use_container_width=True
                        )
                
            except Exception as e:
                st.error(f"Erro: {str(e)}")

# TAB 2: MEUS EBOOKS
with tab2:
    st.markdown("<div class='divider'></div>", unsafe_allow_html=True)
    
    if len(st.session_state.ebooks) == 0:
        st.info("Nenhum ebook criado ainda. Comece agora!")
    else:
        st.markdown(f"### Total: {len(st.session_state.ebooks)} Ebooks")
        
        for ebook in reversed(st.session_state.ebooks):
            titulo_expander = f"EBOOK: {ebook['tema']} - {ebook['publico']} - {ebook['data']}"
            with st.expander(titulo_expander, expanded=False):
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.write(ebook['conteudo'][:500] + "...")
                with col2:
                    if st.button(f"Ver Completo", key=f"view_{ebook['id']}"):
                        st.write(ebook['conteudo'])

# TAB 3: DASHBOARD
with tab3:
    st.markdown("<div class='divider'></div>", unsafe_allow_html=True)
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
            <div class="metric-card">
                <h3 style="margin: 0;">{len(st.session_state.ebooks)}</h3>
                <p style="margin: 0; font-size: 12px;">Ebooks Criados</p>
            </div>
        """, unsafe_allow_html=True)
    
    with col2:
        total_capitulos = sum([e['capitulos'] for e in st.session_state.ebooks]) if st.session_state.ebooks else 0
        st.markdown(f"""
            <div class="metric-card">
                <h3 style="margin: 0;">{total_capitulos}</h3>
                <p style="margin: 0; font-size: 12px;">Capitulos Gerados</p>
            </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown("""
            <div class="metric-card">
                <h3 style="margin: 0;">12</h3>
                <p style="margin: 0; font-size: 12px;">Temas Disponiveis</p>
            </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown("""
            <div class="metric-card">
                <h3 style="margin: 0;">Infinito</h3>
                <p style="margin: 0; font-size: 12px;">Uso GRATIS</p>
            </div>
        """, unsafe_allow_html=True)
    
    st.markdown("<div class='divider'></div>", unsafe_allow_html=True)
    
    st.markdown("### Dicas para Monetizar")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        **1. Crie em Canva**
        - Faca design bonito
        - Capa profissional
        
        **2. Publique em:**
        - Gumroad (20% taxa)
        - Hotmart (50% taxa)
        - Seu proprio site
        """)
    
    with col2:
        st.markdown("""
        **3. Precos Recomendados**
        - Basico: R$ 27
        - Standard: R$ 67
        - Premium: R$ 127
        
        **4. Anuncie em:**
        - Google Ads
        - Facebook Ads
        - Email marketing
        """)

# FOOTER
st.markdown("<div class='divider'></div>", unsafe_allow_html=True)
st.markdown("""
    <div style="text-align: center; padding: 24px; color: #94a3b8;">
        <p style="margin: 0; font-size: 14px;">
            <strong>Luciana Britto | L&B Marketing - Estrategias de Valor</strong>
        </p>
        <p style="margin: 8px 0 0 0; font-size: 12px; opacity: 0.7;">
            Copyright 2026 - Ferramenta de IA para Empreendoras - Usando Google Gemini
        </p>
    </div>
""", unsafe_allow_html=True)
