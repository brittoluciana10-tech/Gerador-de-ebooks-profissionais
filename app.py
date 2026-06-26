import streamlit as st
import google.generativeai as genai
from fpdf import FPDF
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Gerador de Ebooks Premium", page_icon="📚", layout="wide")

# CSS customizado
st.markdown("""
    <style>
    .main { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); }
    </style>
""", unsafe_allow_html=True)

st.title("📚 Gerador de Ebooks Premium")
st.markdown("Crie ebooks profissionais com IA Google (GRATIS!) + PDF + Pagina de Venda")

# Temas disponíveis
TEMAS = {
    "Maternidade Real": "Maternidade real, maes, experiencias honestas",
    "Sono do Bebe": "Sono infantil, tecnicas de dormir, rotina",
    "Pos-parto": "Recuperacao pos-parto, saude, bem-estar",
    "Parenting": "Educacao infantil, desenvolvimento, parentalidade",
    "Renda Variavel": "Renda extra, freelance, negocios digitais",
    "Bem-estar": "Mindfulness, meditacao, saude mental",
    "Autonom": "Ser autonomo, gestao, financeiro",
}

col1, col2 = st.columns(2)

with col1:
    tema_nome = st.selectbox("Escolha o Tema", list(TEMAS.keys()))
    tema_desc = TEMAS[tema_nome]

with col2:
    audience = st.text_input("Publico-alvo", "maes")

# Opcoes avancadas
with st.expander("Opcoes Avancadas"):
    num_chapters = st.slider("Numero de capitulos", 2, 5, 3)
    style = st.selectbox("Estilo", ["Profissional", "Descontraido", "Cientifico"])

if st.button("Gerar Ebook Premium", use_container_width=True):
    try:
        api_key = st.secrets["GEMINI_API_KEY"]
        genai.configure(api_key=api_key)
        
        # Descobrir modelo disponivel
        models = [m.name for m in genai.list_models()]
        model_name = [m for m in models if "gemini" in m.lower()][0]
        model = genai.GenerativeModel(model_name)
        
        with st.spinner("Gerando conteudo com Google Gemini..."):
            prompt = f"""Crie um ebook profissional sobre '{tema_desc}' para '{audience}'.

Estrutura EXATA (importante):
- Titulo atrativo
- Introducao (2-3 paragrafos)
- {num_chapters} capitulos, cada um com:
  * Titulo do capitulo
  * 2-3 secoes tematicas
  * Conteudo detalhado (300+ palavras total por capitulo)
- Conclusao
- Dicas praticas

Estilo: {style}
Leve, profissional e pratico."""

            response = model.generate_content(prompt)
            content = response.text
            
            st.success("Ebook gerado com sucesso!")
            
            # Exibir conteúdo
            st.markdown("---")
            st.write(content)
            
            # Gerar PDF
            with st.spinner("Gerando PDF..."):
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", "B", 18)
                pdf.cell(0, 10, tema_nome, ln=True, align="C")
                
                pdf.set_font("Arial", "I", 10)
                pdf.cell(0, 10, "Por Luciana Britto | L&B Marketing", ln=True, align="C")
                pdf.cell(0, 10, f"Gerado em {datetime.now().strftime('%d/%m/%Y')}", ln=True, align="C")
                
                pdf.ln(10)
                pdf.set_font("Arial", "", 10)
                
                # Quebrar texto em linhas (sem emojis)
                for line in content.split('\n'):
                    if line.strip():
                        try:
                            # Remove caracteres especiais problemáticos
                            clean_line = line.encode('ascii', 'ignore').decode('ascii')
                            if clean_line.strip():
                                pdf.multi_cell(0, 5, clean_line[:500])
                        except:
                            pass
                
                pdf.ln(5)
                pdf.set_font("Arial", "I", 8)
                pdf.cell(0, 5, "© 2026 Luciana Britto | L&B Marketing - Estrategias de Valor", ln=True)
                
                # Salvar em bytes
                pdf_bytes = pdf.output()
                
                st.download_button(
                    label="Download PDF",
                    data=pdf_bytes,
                    file_name=f"{tema_nome}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            
            # Seção de página de venda
            st.markdown("---")
            st.subheader("Proximo Passo: Pagina de Venda")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.info("""
Para vender este ebook:
1. Crie pagina de venda em Canva
2. Integre Stripe para pagamentos
3. Configure email automatico
4. Anuncie no Google Ads/Facebook
                """)
            
            with col2:
                st.warning("""
Ou use seu skill ebook-architect para:
- Validar nicho automaticamente
- Gerar ebook completo
- Criar landing page
- Integrar com Stripe
                """)
            
    except Exception as e:
        st.error(f"Erro: {str(e)}")

# Footer
st.markdown("---")
st.markdown("""
**Luciana Britto | L&B Marketing — Estrategias de Valor**

Dica: Use este gerador para:
- Criar produtos digitais rapido
- Testar nichos antes de investir
- Gerar conteudo para email marketing
- Criar lead magnets gratis
""")
