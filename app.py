import streamlit as st
import google.generativeai as genai
from fpdf import FPDF
from datetime import datetime
import io

st.set_page_config(page_title="Gerador de Ebooks Premium", page_icon="📚", layout="wide")

st.title("📚 Gerador de Ebooks Premium")
st.markdown("Crie ebooks profissionais com IA Google (GRATIS!)")

# Temas disponíveis
TEMAS = {
    "Maternidade Real": "Maternidade real, maes, experiencias honestas",
    "Sono do Bebe": "Sono infantil, tecnicas de dormir, rotina",
    "Pos-parto": "Recuperacao pos-parto, saude, bem-estar",
    "Parenting": "Educacao infantil, desenvolvimento, parentalidade",
    "Renda Variavel": "Renda extra, freelance, negocios digitais",
    "Bem-estar": "Mindfulness, meditacao, saude mental",
    "Autonomo": "Ser autonomo, gestao, financeiro",
}

col1, col2 = st.columns(2)

with col1:
    tema_nome = st.selectbox("Escolha o Tema", list(TEMAS.keys()))
    tema_desc = TEMAS[tema_nome]

with col2:
    audience = st.text_input("Publico-alvo", "maes")

if st.button("Gerar Ebook Premium", use_container_width=True):
    try:
        api_key = st.secrets["GEMINI_API_KEY"]
        genai.configure(api_key=api_key)
        
        models = [m.name for m in genai.list_models()]
        model_name = [m for m in models if "gemini" in m.lower()][0]
        model = genai.GenerativeModel(model_name)
        
        with st.spinner("Gerando conteudo..."):
            prompt = f"""Crie um ebook sobre '{tema_desc}' para '{audience}'.
Estrutura: titulo, introducao, 3 capitulos detalhados, conclusao."""

            response = model.generate_content(prompt)
            content = response.text
            
            st.success("Ebook gerado!")
            st.write(content)
            
            # Gerar PDF
            with st.spinner("Gerando PDF..."):
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", "B", 16)
                pdf.cell(0, 10, tema_nome, ln=True, align="C")
                pdf.set_font("Arial", "I", 9)
                pdf.cell(0, 10, "Por Luciana Britto | L&B Marketing", ln=True, align="C")
                pdf.ln(8)
                pdf.set_font("Arial", "", 10)
                
                for line in content.split('\n'):
                    if line.strip():
                        clean = line.encode('latin-1', 'ignore').decode('latin-1')
                        pdf.multi_cell(0, 5, clean)
                
                pdf_output = pdf.output(dest='S').encode('latin-1')
                
                st.download_button(
                    label="Download PDF",
                    data=pdf_output,
                    file_name=f"{tema_nome}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            
            st.markdown("---")
            st.info("Proximo passo: Crie uma pagina de venda em Canva e venda no Gumroad ou Hotmart!")
            
    except Exception as e:
        st.error(f"Erro: {str(e)}")

st.markdown("---")
st.markdown("**Luciana Britto | L&B Marketing**")
