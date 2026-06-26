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
    
    .section-title {
        color: #3b82f6;
        font-size: 18px;
        font-weight: 700;
        margin: 24px 0 12px 0;
        border-left: 4px solid #3b82f6;
        padding-left: 12px;
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

# IDIOMAS
IDIOMAS = {
    "Português": "Português",
    "Inglês": "English",
    "Espanhol": "Español",
    "Francês": "Français",
    "Italiano": "Italiano",
}

# ESTILOS DE ARTE
ESTILOS_ARTE = {
    "Foto": "Fotografia realista, alta qualidade, profissional",
    "Ilustração": "Ilustração, desenho artístico, colorido",
    "Abstrato": "Arte abstrata, formas geométricas, moderna",
    "Digital": "Arte digital, cyberpunk, futurista",
    "Aquarela": "Estilo aquarela, suave, artístico",
    "Minimalista": "Design minimalista, limpo, moderno",
}

# HEADER
st.markdown("""
    <div class="header-card">
        <h1 style="margin: 0; font-size: 48px;">📚 Ebook Creator Pro Premium</h1>
        <p style="margin: 12px 0 0 0; font-size: 18px; opacity: 0.9;">Crie ebooks profissionais com CONTROLE TOTAL</p>
    </div>
""", unsafe_allow_html=True)

# TABS
tab1, tab2, tab3 = st.tabs(["Criar Ebook", "Meus Ebooks", "Dashboard"])

# TAB 1: CRIAR EBOOK
with tab1:
    st.markdown("<div class='divider'></div>", unsafe_allow_html=True)
    
    # SECAO 1: BASICO
    st.markdown("<div class='section-title'>1. Configuracao Basica</div>", unsafe_allow_html=True)
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        tema_nome = st.selectbox(
            "Escolha o tema",
            list(TEMAS.keys()),
            label_visibility="collapsed"
        )
        tema_desc = TEMAS[tema_nome]
    
    with col2:
        audience = st.text_input(
            "Publico-alvo",
            "Maes portuguesas",
            label_visibility="collapsed"
        )
    
    st.markdown("<div class='divider'></div>", unsafe_allow_html=True)
    
    # SECAO 2: ESTRUTURA
    st.markdown("<div class='section-title'>2. Estrutura do Ebook</div>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        num_chapters = st.slider(
            "Numero de Capitulos",
            min_value=2,
            max_value=10,
            value=3,
            step=1
        )
    
    with col2:
        palavras_capitulo = st.select_slider(
            "Palavras por Capitulo",
            options=[200, 400, 600, 800, 1000, 1200],
            value=600
        )
    
    with col3:
        estilo = st.selectbox(
            "Estilo de Escrita",
            ["Profissional", "Descontraido", "Cientifico", "Inspirador"],
            label_visibility="collapsed"
        )
    
    st.markdown("<div class='divider'></div>", unsafe_allow_html=True)
    
    # SECAO 3: IDIOMA E LOCALIZACAO
    st.markdown("<div class='section-title'>3. Idioma e Localizacao</div>", unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        idioma = st.selectbox(
            "Idioma do Ebook",
            list(IDIOMAS.keys()),
            label_visibility="collapsed"
        )
    
    with col2:
        regiao = st.text_input(
            "Regiao/Pais (ex: Portugal, Brasil)",
            "Portugal",
            label_visibility="collapsed"
        )
    
    st.markdown("<div class='divider'></div>", unsafe_allow_html=True)
    
    # SECAO 4: ESTILO VISUAL
    st.markdown("<div class='section-title'>4. Estilo Visual (para suas imagens)</div>", unsafe_allow_html=True)
    
    estilo_arte = st.selectbox(
        "Escolha o estilo de arte para as imagens",
        list(ESTILOS_ARTE.keys()),
        label_visibility="collapsed"
    )
    
    st.info(f"Estilo selecionado: {ESTILOS_ARTE[estilo_arte]}")
    
    st.markdown("<div class='divider'></div>", unsafe_allow_html=True)
    
    # SECAO 5: DOWNLOADS
    st.markdown("<div class='section-title'>5. Formatos de Download</div>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        formato_txt = st.checkbox("TXT", value=True)
    with col2:
        formato_word = st.checkbox("Word (com espacos para imagens)", value=True)
    with col3:
        formato_pdf = st.checkbox("PDF", value=True)
    
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
                    
                    prompt = f"""Crie um ebook PROFISSIONAL em {idioma} sobre '{tema_desc}' para '{audience}' em {regiao}.

PARAMETROS EXATOS:
- Numero de capitulos: {num_chapters}
- Palavras por capitulo: {palavras_capitulo} aproximadamente
- Estilo: {estilo}
- Estilo visual das imagens: {ESTILOS_ARTE[estilo_arte]}

ESTRUTURA OBRIGATORIA:
- CAPA: [Deixar espaco para imagem de capa]
- TITULO: (titulo atrativo)
- INTRODUCAO: (2-3 paragrafos em {idioma})
- {num_chapters} CAPITULOS:
  * CAPITULO N: [Deixar espaco para imagem relacionada]
  * Titulo do capitulo
  * Conteudo com ~{palavras_capitulo} palavras
  * Adaptado para {regiao}
  * 3 dicas praticas
- CONCLUSAO: (mensagem final)
- BONUS: (5 dicas extras)
- CREDITOS DE IMAGENS: [Espaco para creditos das imagens]

Escreva COMPLETAMENTE em {idioma}.
Use linguagem apropriada para {regiao}.
Deixe claro onde as imagens devem ser inseridas."""
                    
                    response = model.generate_content(prompt)
                    content = response.text
                    st.session_state.current_content = content
                    
                    progress_bar.progress(70)
                    status_text.text("Processando conteudo...")
                    
                    # Gerar sugestoes de imagens
                    prompt_imagens = f"""Baseado neste ebook sobre '{tema_desc}', gere 10 prompts de imagem para usar em Canva, DALL-E ou Midjourney.

Formato:
1. [Titulo da imagem]: [Descricao detalhada em {idioma} usando estilo {ESTILOS_ARTE[estilo_arte]}]

PROMPTS:
1. Capa do Ebook: Uma imagem que represente {tema_desc} em estilo {ESTILOS_ARTE[estilo_arte]}"""
                    
                    response_imagens = model.generate_content(prompt_imagens)
                    sugestoes_imagens = response_imagens.text
                    
                    # Adicionar ao historico
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
                    status_text.empty()
                
                # RESULTADO
                st.markdown("<div class='divider'></div>", unsafe_allow_html=True)
                st.success("Ebook gerado com sucesso!")
                
                # TABS DE RESULTADO
                res_col1, res_col2 = st.tabs(["Conteudo do Ebook", "Prompts de Imagens"])
                
                with res_col1:
                    with st.expander("Ver Conteudo Completo", expanded=True):
                        st.write(content)
                
                with res_col2:
                    with st.expander("Ver Sugestoes de Imagens para Canva/DALL-E", expanded=True):
                        st.write(sugestoes_imagens)
                
                # DOWNLOADS
                st.markdown("<div class='divider'></div>", unsafe_allow_html=True)
                st.markdown("### Download seu Ebook")
                
                col_d1, col_d2, col_d3 = st.columns(3)
                
                # TXT
                if formato_txt:
                    with col_d1:
                        txt = f"{tema_nome}\n\nPor Luciana Britto | L&B Marketing\nIdioma: {idioma}\nRegiao: {regiao}\nCapitulos: {num_chapters}\nEstilo de Imagens: {estilo_arte}\n{datetime.now().strftime('%d/%m/%Y')}\n\n{content}\n\n--- SUGESTOES DE IMAGENS ---\n{sugestoes_imagens}\n\nCopyright 2026 Luciana Britto | L&B Marketing"
                        st.download_button(
                            "Download TXT",
                            txt,
                            f"{tema_nome}_{idioma}.txt",
                            "text/plain",
                            use_container_width=True
                        )
                
                # WORD COM ESPACOS PARA IMAGENS
                if formato_word:
                    with col_d2:
                        doc = Document()
                        
                        # Titulo
                        title = doc.add_heading(tema_nome, 0)
                        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        
                        # Info
                        info = doc.add_paragraph()
                        info.add_run("Por Luciana Britto | L&B Marketing\n").bold = True
                        info.add_run(f"Idioma: {idioma} | Regiao: {regiao}\n")
                        info.add_run(f"Capitulos: {num_chapters} | Estilo: {estilo_arte}\n")
                        info.add_run(f"Gerado em {datetime.now().strftime('%d/%m/%Y')}")
                        info.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        
                        doc.add_paragraph()
                        doc.add_paragraph("[ESPACO RESERVADO PARA CAPA - Insira imagem aqui em Canva ou Word]").italic = True
                        doc.add_paragraph()
                        
                        # Conteudo
                        doc.add_paragraph(content)
                        
                        doc.add_page_break()
                        doc.add_heading("Sugestoes de Imagens para Adicionar", level=1)
                        doc.add_paragraph(sugestoes_imagens)
                        
                        doc_bytes = BytesIO()
                        doc.save(doc_bytes)
                        doc_bytes.seek(0)
                        
                        st.download_button(
                            "Download Word",
                            doc_bytes.getvalue(),
                            f"{tema_nome}_{idioma}_COM_ESPACOS.docx",
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
                        story.append(Paragraph(f"Por Luciana Britto | L&B Marketing", styles['Normal']))
                        story.append(Paragraph(f"Idioma: {idioma} | Regiao: {regiao}", styles['Normal']))
                        story.append(Paragraph(f"{datetime.now().strftime('%d de %B de %Y')}", styles['Normal']))
                        story.append(PageBreak())
                        story.append(Paragraph(content, body_style))
                        story.append(PageBreak())
                        story.append(Paragraph("Sugestoes de Imagens", title_style))
                        story.append(Paragraph(sugestoes_imagens, body_style))
                        
                        doc.build(story)
                        buffer.seek(0)
                        
                        st.download_button(
                            "Download PDF",
                            buffer.getvalue(),
                            f"{tema_nome}_{idioma}.pdf",
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
            titulo_expander = f"EBOOK: {ebook['tema']} ({ebook['idioma']}) - {ebook['capitulos']} cap - {ebook['data']}"
            with st.expander(titulo_expander, expanded=False):
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.write(f"**Publico:** {ebook['publico']}")
                    st.write(f"**Estilo de Imagens:** {ebook['estilo_arte']}")
                    st.write(f"**Palavras por capitulo:** {ebook['palavras']}")
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
                <h3 style="margin: 0;">5 Idiomas</h3>
                <p style="margin: 0; font-size: 12px;">Disponiveis</p>
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
    
    st.markdown("### Como Usar Este Gerador")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        **1. Gere o Ebook**
        - Escolha tema, idioma, capitulos
        - IA gera conteudo profissional
        
        **2. Baixe em Word**
        - Arquivo pronto para editar
        - Espacos marcados para imagens
        - Estrutura profissional
        
        **3. Obtenha Sugestoes**
        - Prompts prontos para Canva
        - Descrições detalhadas
        - Estilos personalizados
        """)
    
    with col2:
        st.markdown("""
        **4. Adicione Imagens em Canva**
        - Use os prompts sugeridos
        - Crie design profissional
        - Mantenha consistencia visual
        
        **5. Exporte e Venda**
        - PDF pronto para vender
        - Gumroad, Hotmart, etc
        - Comece a ganhar!
        
        **6. Repita**
        - Gere novos ebooks
        - Crie sua biblioteca
        - Escale o negocio
        """)

# FOOTER
st.markdown("<div class='divider'></div>", unsafe_allow_html=True)
st.markdown("""
    <div style="text-align: center; padding: 24px; color: #94a3b8;">
        <p style="margin: 0; font-size: 14px;">
            <strong>Luciana Britto | L&B Marketing - Estrategias de Valor</strong>
        </p>
        <p style="margin: 8px 0 0 0; font-size: 12px; opacity: 0.7;">
            Copyright 2026 - Ferramenta PRO Premium de IA para Empreendoras - Com Espacos para Imagens
        </p>
    </div>
""", unsafe_allow_html=True)
