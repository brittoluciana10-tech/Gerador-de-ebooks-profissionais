import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime
from io import BytesIO
import base64

st.set_page_config(page_title="Gerador de Ebooks Premium", page_icon="📚", layout="wide")

st.title("📚 Gerador de Ebooks Premium")
st.markdown("Crie ebooks profissionais + PDF + Word + Pagina de Venda")

# TEMAS EXPANDIDOS
TEMAS = {
    "Maternidade Real": "Maternidade real, maes, experiencias honestas, emocoes",
    "Sono do Bebe": "Sono infantil, tecnicas de dormir, rotina, sono profundo",
    "Pos-parto": "Recuperacao pos-parto, saude, bem-estar, corpo",
    "Parenting": "Educacao infantil, desenvolvimento, parentalidade, criancas",
    "Renda Variavel": "Renda extra, freelance, negocios digitais, ganho",
    "Bem-estar": "Mindfulness, meditacao, saude mental, calma",
    "Autonomo": "Ser autonomo, gestao, financeiro, independencia",
    "Peso Saudavel": "Perda de peso, fitness, saude, nutricao",
    "Criatividade": "Criatividade, inovacao, arte, inspiracao",
    "Financas": "Educacao financeira, investimentos, economia, dinheiro",
    "Relacionamentos": "Relacionamentos, amor, comunicacao, casal",
    "Carreira": "Carreira, emprego, desenvolvimento, profissao",
}

col1, col2 = st.columns(2)

with col1:
    tema_nome = st.selectbox("Escolha o Tema", list(TEMAS.keys()))
    tema_desc = TEMAS[tema_nome]

with col2:
    audience = st.text_input("Publico-alvo", "maes")

# Opcoes
col1, col2 = st.columns(2)
with col1:
    num_chapters = st.slider("Capitulos", 2, 5, 3)
with col2:
    formato = st.multiselect("Download", ["TXT", "Word", "PDF"], default=["TXT", "Word"])

if st.button("Gerar Ebook Premium", use_container_width=True):
    try:
        api_key = st.secrets["GEMINI_API_KEY"]
        genai.configure(api_key=api_key)
        
        models = [m.name for m in genai.list_models()]
        model_name = [m for m in models if "gemini" in m.lower()][0]
        model = genai.GenerativeModel(model_name)
        
        with st.spinner("Gerando conteudo com IA..."):
            prompt = f"""Crie um ebook profissional sobre '{tema_desc}' para '{audience}'.

Estrutura obrigatoria:
- TITULO: (titulo atrativo e profissional)
- INTRODUCAO: (2-3 paragrafos inspiradores)
- {num_chapters} CAPITULOS com:
  * CAPITULO N: Titulo
  * Conteudo detalhado (300+ palavras)
  * Dicas praticas
- CONCLUSAO: (mensagem final inspiradora)
- BONUS: (3 dicas extras)

Seja profissional, detalhado e pratico!"""

            response = model.generate_content(prompt)
            content = response.text
            
            st.success("Ebook gerado com sucesso!")
            st.markdown("---")
            st.write(content)
            
            # DOWNLOAD TXT
            if "TXT" in formato:
                txt_content = f"""{tema_nome}

Por Luciana Britto | L&B Marketing
Gerado em {datetime.now().strftime('%d/%m/%Y')}

---

{content}

---

© 2026 Luciana Britto | L&B Marketing - Estrategias de Valor
"""
                st.download_button(
                    "📥 Download TXT",
                    txt_content,
                    f"{tema_nome}.txt",
                    "text/plain"
                )
            
            # DOWNLOAD WORD
            if "Word" in formato:
                doc = Document()
                
                # Titulo
                title = doc.add_heading(tema_nome, 0)
                title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Subtitulo
                subtitle = doc.add_paragraph()
                subtitle_run = subtitle.add_run(f"Por Luciana Britto | L&B Marketing")
                subtitle_run.bold = True
                subtitle_run.font.size = Pt(12)
                subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Data
                date_para = doc.add_paragraph(f"Gerado em {datetime.now().strftime('%d/%m/%Y')}")
                date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                doc.add_paragraph()
                
                # Conteudo
                doc.add_paragraph(content)
                
                # Footer
                footer_para = doc.add_paragraph()
                footer_run = footer_para.add_run("© 2026 Luciana Britto | L&B Marketing - Estrategias de Valor")
                footer_run.font.size = Pt(8)
                footer_run.italic = True
                footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Salvar em bytes
                doc_bytes = BytesIO()
                doc.save(doc_bytes)
                doc_bytes.seek(0)
                
                st.download_button(
                    "📄 Download Word",
                    doc_bytes.getvalue(),
                    f"{tema_nome}.docx",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            
            # PAGINA DE VENDA
            st.markdown("---")
            st.subheader("🛍️ Crie Pagina de Venda")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.info(f"""
**Sua pagina deve incluir:**

1. **Titulo Atrativo**
   {tema_nome}

2. **Problema do Cliente**
   Que dor sua audiencia ({audience}) tem?

3. **Solucao (Seu Ebook)**
   Por que seu ebook resolve?

4. **Beneficios**
   - Benefit 1
   - Benefit 2
   - Benefit 3

5. **Prova Social**
   Depoimentos ou dados

6. **CTA (Call to Action)**
   "Compre agora por R$29"
                """)
            
            with col2:
                st.warning(f"""
**Ferramentas Recomendadas:**

1. **Canva** (design gratis)
   - Templates de landing page
   - Cores profissionais
   - Fonts bonitos

2. **Plataformas de Venda:**
   - Gumroad (mais facil)
   - Hotmart (mais profissional)
   - Sua propria loja

3. **Email Marketing:**
   - Brevo (gratis)
   - Mailchimp
   - Converkit

4. **Anuncios:**
   - Google Ads
   - Facebook Ads
   - TikTok Ads
                """)
            
            # Sugestoes de preco
            st.markdown("---")
            st.subheader("💰 Sugestoes de Preco")
            
            pricing_col1, pricing_col2, pricing_col3 = st.columns(3)
            
            with pricing_col1:
                st.success("""
**BASICO**
R$ 17-27

- 1 Ebook
- Email de suporte
- Atualizacoes
                """)
            
            with pricing_col2:
                st.info("""
**STANDARD**
R$ 47-77

- Ebook + PDF premium
- Email ilimitado
- Grupo no Discord
- Atualizacoes vitalicia
                """)
            
            with pricing_col3:
                st.warning("""
**PREMIUM**
R$ 97-197

- Ebook + Templates
- 3 chamadas 1:1
- Grupo privado
- Suporte prioritario
                """)
            
    except Exception as e:
        st.error(f"Erro: {str(e)}")

st.markdown("---")
st.markdown("""
**Luciana Britto | L&B Marketing — Estrategias de Valor**

🚀 Seus proximos passos:
1. Gere o ebook
2. Baixe em Word/TXT
3. Crie pagina de venda em Canva
4. Publique em Gumroad/Hotmart
5. Anuncie no Google Ads
6. Ganhe passivamente! 💰
""")
