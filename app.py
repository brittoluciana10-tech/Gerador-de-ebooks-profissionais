import streamlit as st
import google.generativeai as genai
from datetime import datetime

st.set_page_config(page_title="Gerador de Ebooks", page_icon="📚")

st.title("📚 Gerador de Ebooks Premium")

TEMAS = {
    "Maternidade Real": "Maternidade real, maes, experiencias honestas",
    "Sono do Bebe": "Sono infantil, tecnicas de dormir, rotina",
    "Pos-parto": "Recuperacao pos-parto, saude, bem-estar",
    "Parenting": "Educacao infantil, desenvolvimento",
    "Renda Variavel": "Renda extra, freelance, negocios",
    "Bem-estar": "Mindfulness, meditacao, saude mental",
    "Autonomo": "Ser autonomo, gestao, financeiro",
}

tema_nome = st.selectbox("Tema", list(TEMAS.keys()))
audience = st.text_input("Publico-alvo", "maes")

if st.button("Gerar Ebook"):
    try:
        api_key = st.secrets["GEMINI_API_KEY"]
        genai.configure(api_key=api_key)
        
        models = [m.name for m in genai.list_models()]
        model_name = [m for m in models if "gemini" in m.lower()][0]
        model = genai.GenerativeModel(model_name)
        
        with st.spinner("Gerando..."):
            response = model.generate_content(
                f"Ebook sobre {tema_nome} para {audience}. "
                f"3 capitulos com titulo, introducao e conteudo."
            )
            content = response.text
            
        st.success("Pronto!")
        st.write(content)
        
        # DOWNLOAD FUNCIONA 100%
        download_text = f"{tema_nome}\n\nPor Luciana Britto\n{datetime.now().strftime('%d/%m/%Y')}\n\n{content}"
        
        st.download_button(
            "Download TXT",
            download_text,
            f"{tema_nome}.txt",
            "text/plain"
        )
        
    except Exception as e:
        st.error(str(e))

st.markdown("**Luciana Britto | L&B Marketing**")
