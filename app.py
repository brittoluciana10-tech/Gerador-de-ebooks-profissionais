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
import json

st.set_page_config(page_title="Gerador de Ebooks Premium 2.0", page_icon="📚", layout="wide")

# CSS
st.markdown("""
    <style>
    .main { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); }
    .stButton>button { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; }
    </style>
""", unsafe_allow_html=True)

# INICIALIZAR SESSION STATE
if "user" not in st.session_state:
    st.session_state.user = None
if "ebooks" not in st.session_state:
    st.session_state.ebooks = []
if "page" not in st.session_state:
    st.session_state.page = "home"

# TEMAS
TEMAS = {
    "Maternidade Real": "Maternidade real, maes, experiencias honestas",
    "Sono do Bebe": "Sono infantil, tecnicas de dormir, rotina",
    "Pos-parto": "Recuperacao pos-parto, saude, bem-estar",
    "Parenting": "Educacao infantil, desenvolvimento, parentalidade",
    "Renda Variavel": "Renda extra, freelance, negocios digitais",
    "Bem-estar": "Mindfulness, meditacao, saude mental",
    "Autonomo": "Ser autonomo, gestao, financeiro",
    "Peso Saudavel": "Perda de peso, fitness, saude, nutricao",
    "Criatividade": "Criatividade, inovacao, arte, inspiracao",
    "Financas": "Educacao financeira, investimentos, economia",
    "Relacionamentos": "Relacionamentos, amor, comunicacao",
    "Carreira": "Carreira, emprego, desenvolvimento profissional",
}

# FUNCAO: Gerar PDF Profissional
def gerar_pdf_profissional(tema_nome, content):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    story = []
    styles = getSampleStyleSheet()
    
    # Estilos customizados
    titulo_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=28,
        textColor=colors.HexColor('#667eea'),
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
    
    # Titulo
    story.append(Paragraph(tema_nome, titulo_style))
    story.append(Spacer(1, 0.3*inch))
    
    # Autor
    author_style = ParagraphStyle('Author', parent=styles['Normal'], alignment=TA_CENTER)
    story.append(Paragraph("<i>Por Luciana Britto | L&B Marketing</i>", author_style))
    story.append(Spacer(1, 0.2*inch))
    
    # Data
    story.append(Paragraph(f"<i>{datetime.now().strftime('%d de %B de %Y')}</i>", author_style))
    story.append(PageBreak())
    
    # Conteudo
    story.append(Paragraph(content, body_style))
    
    # Footer
    story.append(Spacer(1, 0.5*inch))
    footer_style = ParagraphStyle('Footer', parent=styles['Normal'], fontSize=8, alignment=TA_CENTER)
    story.append(Paragraph("© 2026 Luciana Britto | L&B Marketing - Estrategias de Valor", footer_style))
    
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# PAGINA: HOME
if st.session_state.page == "home":
    st.title("📚 Gerador de Ebooks Premium 2.0")
    st.markdown("Crie ebooks profissionais + PDF + Word + Historico")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if st.button("🆕 Novo Ebook", use_container_width=True):
            st.session_state.page = "create"
            st.rerun()
    
    with col2:
        if st.button("📚 Meus Ebooks", use_container_width=True):
            st.session_state.page = "history"
            st.rerun()
    
    with col3:
        if st.button("⚙️ Configuracoes", use_container_width=True):
            st.session_state.page = "settings"
            st.rerun()
    
    st.markdown("---")
    st.info(f"Total de ebooks gerados: **{len(st.session_state.ebooks)}**")

# PAGINA: CREATE
elif st.session_state.page == "create":
    st.title("🆕 Novo Ebook")
    
    if st.button("← Voltar"):
        st.session_state.page = "home"
        st.rerun()
    
    col1, col2 = st.columns(2)
    
    with col1:
        tema_nome = st.selectbox("Tema", list(TEMAS.keys()))
        tema_desc = TEMAS[tema_nome]
    
    with col2:
        audience = st.text_input("Publico-alvo", "maes")
    
    col1, col2 = st.columns(2)
    with col1:
        num_chapters = st.slider("Capitulos", 2, 5, 3)
    with col2:
        formato = st.multiselect("Download", ["TXT", "Word", "PDF"], default=["TXT", "Word", "PDF"])
    
    if st.button("🚀 Gerar Ebook", use_container_width=True):
        try:
            api_key = st.secrets["GEMINI_API_KEY"]
            genai.configure(api_key=api_key)
            
            models = [m.name for m in genai.list_models()]
            model_name = [m for m in models if "gemini" in m.lower()][0]
            model = genai.GenerativeModel(model_name)
            
            with st.spinner("Gerando conteudo..."):
                prompt = f"""Crie um ebook sobre '{tema_desc}' para '{audience}'.

Estrutura:
- TITULO
- INTRODUCAO (2-3 paragrafos)
- {num_chapters} CAPITULOS com conteudo detalhado
- CONCLUSAO
- BONUS (3 dicas)

Profissional e pratico!"""
                
                response = model.generate_content(prompt)
                content = response.text
                
                st.success("Ebook gerado!")
                st.write(content)
                
                # Adicionar ao historico
                ebook_data = {
                    "id": len(st.session_state.ebooks) + 1,
                    "tema": tema_nome,
                    "data": datetime.now().strftime("%d/%m/%Y %H:%M"),
                    "conteudo": content
                }
                st.session_state.ebooks.append(ebook_data)
                
                # DOWNLOADS
                col1, col2, col3 = st.columns(3)
                
                # TXT
                if "TXT" in formato:
                    with col1:
                        txt = f"{tema_nome}\n\nPor Luciana Britto\n{datetime.now().strftime('%d/%m/%Y')}\n\n{content}\n\n© 2026 Luciana Britto | L&B Marketing"
                        st.download_button("📥 TXT", txt, f"{tema_nome}.txt", "text/plain")
                
                # WORD
                if "Word" in formato:
                    with col2:
                        doc = Document()
                        title = doc.add_heading(tema_nome, 0)
                        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        doc.add_paragraph(f"Por Luciana Britto | L&B Marketing\n{datetime.now().strftime('%d/%m/%Y')}", style='Normal').alignment = WD_ALIGN_PARAGRAPH.CENTER
                        doc.add_paragraph(content)
                        
                        doc_bytes = BytesIO()
                        doc.save(doc_bytes)
                        doc_bytes.seek(0)
                        
                        st.download_button("📄 Word", doc_bytes.getvalue(), f"{tema_nome}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                
                # PDF
                if "PDF" in formato:
                    with col3:
                        pdf_data = gerar_pdf_profissional(tema_nome, content)
                        st.download_button("🎨 PDF", pdf_data, f"{tema_nome}.pdf", "application/pdf")
                
        except Exception as e:
            st.error(f"Erro: {str(e)}")

# PAGINA: HISTORY
elif st.session_state.page == "history":
    st.title("📚 Meus Ebooks")
    
    if st.button("← Voltar"):
        st.session_state.page = "home"
        st.rerun()
    
    if len(st.session_state.ebooks) == 0:
        st.info("Nenhum ebook gerado ainda. Comece agora!")
    else:
        for ebook in reversed(st.session_state.ebooks):
            with st.expander(f"📖 {ebook['tema']} - {ebook['data']}"):
                st.write(ebook['conteudo'][:500] + "...")
                if st.button(f"Ver completo - {ebook['id']}"):
                    st.write(ebook['conteudo'])

# PAGINA: SETTINGS
elif st.session_state.page == "settings":
    st.title("⚙️ Configuracoes")
    
    if st.button("← Voltar"):
        st.session_state.page = "home"
        st.rerun()
    
    st.info("""
**Sobre sua conta:**
- Total de ebooks: {}
- Ultimo gerado: {}
    """.format(len(st.session_state.ebooks), st.session_state.ebooks[-1]['data'] if st.session_state.ebooks else "Nenhum"))

st.markdown("---")
st.markdown("**Luciana Britto | L&B Marketing — Estrategias de Valor**")
