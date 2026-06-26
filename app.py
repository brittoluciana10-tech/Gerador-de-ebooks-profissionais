import streamlit as st
from anthropic import Anthropic
from openai import OpenAI
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, PageBreak, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from datetime import datetime
import io
import json
import requests
from PIL import Image as PILImage

st.set_page_config(page_title="Premium Ebook Generator", page_icon="📚", layout="wide")

st.markdown("""
    <style>
    .main { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); }
    .stButton > button { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border: none; }
    h1 { color: #667eea; }
    </style>
""", unsafe_allow_html=True)

@st.cache_resource
def init_clients():
    claude = Anthropic(api_key=st.secrets.get("ANTHROPIC_API_KEY", ""))
    openai = OpenAI(api_key=st.secrets.get("OPENAI_API_KEY", ""))
    return claude, openai

claude_client, openai_client = init_clients()

st.title("📚 Premium Ebook Generator")
st.subtitle("Crie ebooks profissionais tipo Gamma com IA")

col1, col2 = st.columns(2)

with col1:
    topic = st.text_input("📖 Tema do Ebook", "Maternidade real")
    audience = st.text_input("👥 Público-alvo", "mães e pais")

with col2:
    num_chapters = st.slider("📑 Número de capítulos", 2, 5, 3)
    style = st.selectbox("🎨 Estilo", ["Moderno", "Clássico", "Minimalista"])

if st.button("🚀 Gerar Ebook Premium", key="generate", use_container_width=True):
    
    with st.spinner("⏳ Gerando conteúdo com Claude..."):
        try:
            prompt = f"""Crie um ebook profissional sobre "{topic}" para {audience}.

Responda APENAS em JSON válido (sem markdown):
{{
  "title": "Título atrativo",
  "subtitle": "Subtítulo inspirador",
  "author": "Luciana Britto",
  "description": "Descrição breve",
  "chapters": [
    {{
      "number": 1,
      "title": "Título do capítulo",
      "introduction": "Introdução (2-3 frases)",
      "sections": [
        {{
          "heading": "Seção 1",
          "content": "Conteúdo detalhado (300+ palavras)"
        }},
        {{
          "heading": "Seção 2",
          "content": "Conteúdo detalhado (300+ palavras)"
        }}
      ],
      "conclusion": "Conclusão (2-3 frases)",
      "image_prompt": "Descrição para imagem de capa"
    }}
  ],
  "conclusion": "Conclusão geral"
}}"""

            response = claude_client.messages.create(
                model="claude-opus-4-6",
                max_tokens=4000,
                messages=[{"role": "user", "content": prompt}]
            )

            content_text = response.content[0].text
            json_start = content_text.find('{')
            json_end = content_text.rfind('}') + 1
            json_str = content_text[json_start:json_end]
            ebook_content = json.loads(json_str)
            
            st.success("✅ Conteúdo gerado!")

    with st.spinner("🖼️ Gerando imagens com DALL-E..."):
        images = {}
        
        try:
            cover_response = openai_client.images.generate(
                model="dall-e-3",
                prompt=f'Professional premium book cover for "{ebook_content["title"]}" - modern design, gradient background, elegant typography, high quality, 1024x1024',
                n=1,
                size="1024x1024"
            )
            images["cover"] = cover_response.data[0].url
            st.success("✅ Cover gerada!")
        except Exception as e:
            st.warning(f"⚠️ Erro ao gerar cover: {str(e)}")
            images["cover"] = None

        for chapter in ebook_content["chapters"][:2]:
            try:
                img_response = openai_client.images.generate(
                    model="dall-e-3",
                    prompt=f'{chapter.get("image_prompt", chapter["title"])} - professional, high quality, modern design',
                    n=1,
                    size="1024x1024"
                )
                images[f'chapter_{chapter["number"]}'] = img_response.data[0].url
                st.success(f"✅ Imagem capítulo {chapter['number']} gerada!")
            except Exception as e:
                st.warning(f"⚠️ Erro ao gerar imagem capítulo {chapter['number']}")

    with st.spinner("📄 Gerando PDF profissional..."):
        try:
            pdf_buffer = io.BytesIO()
            doc = SimpleDocTemplate(
                pdf_buffer,
                pagesize=A4,
                rightMargin=0.75*inch,
                leftMargin=0.75*inch,
                topMargin=0.75*inch,
                bottomMargin=0.75*inch
            )

            story = []
            styles = getSampleStyleSheet()

            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=36,
                textColor=colors.HexColor('#667eea'),
                spaceAfter=6,
                alignment=TA_CENTER,
                fontName='Helvetica-Bold'
            )

            subtitle_style = ParagraphStyle(
                'CustomSubtitle',
                parent=styles['Normal'],
                fontSize=14,
                textColor=colors.HexColor('#764ba2'),
                spaceAfter=30,
                alignment=TA_CENTER
            )

            heading_style = ParagraphStyle(
                'CustomHeading',
                parent=styles['Heading2'],
                fontSize=18,
                textColor=colors.HexColor('#333333'),
                spaceAfter=12,
                spaceBefore=12,
                borderColor=colors.HexColor('#667eea'),
                borderWidth=2,
                borderPadding=5
            )

            body_style = ParagraphStyle(
                'CustomBody',
                parent=styles['BodyText'],
                fontSize=11,
                alignment=TA_JUSTIFY,
                spaceAfter=12,
                leading=16
            )

            story.append(Spacer(1, 1*inch))
            story.append(Paragraph(ebook_content["title"], title_style))
            story.append(Paragraph(ebook_content["subtitle"], subtitle_style))
            
            if images.get("cover"):
                try:
                    img_data = requests.get(images["cover"]).content
                    img = PILImage.open(io.BytesIO(img_data))
                    img_width = 4*inch
                    img_height = (img.height / img.width) * img_width
                    img_path = io.BytesIO(img_data)
                    story.append(Image(img_path, width=img_width, height=img_height, hAlign='CENTER'))
                    story.append(Spacer(1, 0.5*inch))
                except:
                    pass

            story.append(Spacer(1, 1*inch))
            story.append(Paragraph(f"por {ebook_content['author']}", styles['Normal']))
            story.append(Spacer(1, 0.5*inch))
            story.append(Paragraph(str(datetime.now().year), styles['Normal']))
            story.append(PageBreak())

            story.append(Paragraph("Índice", styles['Heading1']))
            story.append(Spacer(1, 0.2*inch))
            for chapter in ebook_content["chapters"]:
                story.append(Paragraph(f"Capítulo {chapter['number']}: {chapter['title']}", styles['Normal']))
            story.append(PageBreak())

            for chapter in ebook_content["chapters"]:
                story.append(Paragraph(f"CAPÍTULO {chapter['number']}", styles['Heading2']))
                story.append(Paragraph(chapter["title"], heading_style))
                story.append(Spacer(1, 0.2*inch))

                if images.get(f'chapter_{chapter["number"]}'):
                    try:
                        img_data = requests.get(images[f'chapter_{chapter["number"]}']).content
                        img = PILImage.open(io.BytesIO(img_data))
                        img_width = 3.5*inch
                        img_height = (img.height / img.width) * img_width
                        img_path = io.BytesIO(img_data)
                        story.append(Image(img_path, width=img_width, height=img_height, hAlign='CENTER'))
                        story.append(Spacer(1, 0.2*inch))
                    except:
                        pass

                story.append(Paragraph(chapter["introduction"], body_style))

                for section in chapter.get("sections", []):
                    story.append(Paragraph(section["heading"], heading_style))
                    story.append(Paragraph(section["content"], body_style))

                story.append(Spacer(1, 0.1*inch))
                story.append(Paragraph(f"<i>{chapter['conclusion']}</i>", body_style))
                story.append(PageBreak())

            story.append(Spacer(1, 1*inch))
            story.append(Paragraph("Conclusão", styles['Heading1']))
            story.append(Paragraph(ebook_content["conclusion"], body_style))
            story.append(Spacer(1, 1*inch))
            story.append(Paragraph("© 2026 Luciana Britto | L&B Marketing — Estratégias de Valor", styles['Normal']))

            doc.build(story)
            pdf_buffer.seek(0)

            st.success("✅ PDF gerado com sucesso!")

            st.download_button(
                label="📥 Download PDF",
                data=pdf_buffer,
                file_name=f"{ebook_content['title']}.pdf",
                mime="application/pdf",
                use_container_width=True
            )

            st.markdown("---")
            st.subheader("📖 Conteúdo Gerado")
            st.json(ebook_content)

        except Exception as e:
            st.error(f"❌ Erro ao gerar PDF: {str(e)}")

st.markdown("---")
st.markdown("**Desenvolvido por Luciana Britto | L&B Marketing — Estratégias de Valor**")
